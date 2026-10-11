# テスト用の PDF を組み立てる（外から PDF を持ってこない。作成者などのメタデータは入れない）。
#   . "$PSScriptRoot\..\helpers\pdf.ps1"
# 文字列はすべて Latin-1 の 1 バイト = 1 文字として扱う。

$script:pdfTestLatin1 = [System.Text.Encoding]::GetEncoding(28591)

function toPdfBytes([string]$text) {
    return , ($script:pdfTestLatin1.GetBytes($text))
}

function joinPdfBytes([object[]]$parts) {
    $ms = New-Object System.IO.MemoryStream
    foreach ($p in $parts) {
        if ($null -eq $p) { continue }
        if ($p -is [byte[]]) { $ms.Write($p, 0, $p.Length) }
        else { $b = $script:pdfTestLatin1.GetBytes([string]$p); $ms.Write($b, 0, $b.Length) }
    }
    return , ($ms.ToArray())
}

function compressPdfBytes([byte[]]$bytes) {
    # zlib 形式（2 バイトの先頭 + deflate + Adler-32）
    $ms = New-Object System.IO.MemoryStream
    $ms.WriteByte(0x78); $ms.WriteByte(0x9C)
    $ds = New-Object System.IO.Compression.DeflateStream($ms, [System.IO.Compression.CompressionMode]::Compress, $true)
    $ds.Write($bytes, 0, $bytes.Length)
    $ds.Dispose()
    [uint32]$a = 1; [uint32]$b = 0
    foreach ($x in $bytes) {
        $a = ($a + $x) % 65521
        $b = ($b + $a) % 65521
    }
    $adler = ([uint64]$b -shl 16) -bor $a
    foreach ($shift in 24, 16, 8, 0) { $ms.WriteByte([byte](($adler -shr $shift) -band 0xFF)) }
    return , ($ms.ToArray())
}

function newPdfStreamObject([string]$dictBody, [byte[]]$data) {
    # /Length を足した辞書と本体
    return @{ Dict = "<< $dictBody /Length $($data.Length) >>"; Stream = $data }
}

function newPdfFromObjects {
    # objects: 番号 -> 文字列（辞書など）か、newPdfStreamObject の戻り値。xref は表形式
    param (
        [System.Collections.IDictionary]$objects,
        [string]$trailerExtra = '/Root 1 0 R',
        [string]$header = "%PDF-1.4`n"
    )

    $ms = New-Object System.IO.MemoryStream
    $h = toPdfBytes $header
    $ms.Write($h, 0, $h.Length)
    $offsets = @{}
    $maxId = ($objects.Keys | Measure-Object -Maximum).Maximum
    foreach ($id in ($objects.Keys | Sort-Object)) {
        $offsets[[int]$id] = [int]$ms.Length
        $o = $objects[$id]
        if ($o -is [string]) {
            $b = joinPdfBytes @("$id 0 obj`n", $o, "`nendobj`n")
        }
        else {
            $b = joinPdfBytes @("$id 0 obj`n", $o.Dict, "`nstream`n", $o.Stream, "`nendstream`nendobj`n")
        }
        $ms.Write($b, 0, $b.Length)
    }
    $xrefPos = [int]$ms.Length
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("xref`n0 $($maxId + 1)`n")
    [void]$sb.Append("0000000000 65535 f `n")
    for ($i = 1; $i -le $maxId; $i++) {
        if ($offsets.ContainsKey($i)) { [void]$sb.Append(('{0:D10} 00000 n ' -f $offsets[$i]) + "`n") }
        else { [void]$sb.Append("0000000000 00000 f `n") }
    }
    [void]$sb.Append("trailer`n<< /Size $($maxId + 1) $trailerExtra >>`nstartxref`n$xrefPos`n%%EOF`n")
    $t = toPdfBytes $sb.ToString()
    $ms.Write($t, 0, $t.Length)
    return , ($ms.ToArray())
}

function newSimplePdf {
    # pages: 各ページの内容ストリーム（文字列）。フォント F1 は Helvetica（WinAnsi）。compress なら Flate にする
    param (
        [string[]]$pages,
        [switch]$compress,
        [string]$fontObject = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>',
        [System.Collections.IDictionary]$extraObjects = $null
    )

    $objs = @{}
    $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
    $kids = (0..($pages.Count - 1) | ForEach-Object { "$(10 + $_ * 2) 0 R" }) -join ' '
    $objs[2] = "<< /Type /Pages /Kids [$kids] /Count $($pages.Count) >>"
    $objs[3] = $fontObject
    for ($i = 0; $i -lt $pages.Count; $i++) {
        $pageId = 10 + $i * 2
        $contentId = $pageId + 1
        $objs[$pageId] = "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R >> >> /Contents $contentId 0 R >>"
        $data = toPdfBytes $pages[$i]
        if ($compress) {
            $objs[$contentId] = newPdfStreamObject '/Filter /FlateDecode' (compressPdfBytes $data)
        }
        else {
            $objs[$contentId] = newPdfStreamObject '' $data
        }
    }
    if ($null -ne $extraObjects) { foreach ($k in $extraObjects.Keys) { $objs[[int]$k] = $extraObjects[$k] } }
    return , (newPdfFromObjects $objs)
}

function newPdfWithXrefStream {
    # normal: 番号 -> 文字列か newPdfStreamObject（そのまま書く）。packed: 番号 -> 辞書の文字列（オブジェクトのストリームにまとめる）
    # 相互参照は xref ストリーム（Flate。predictor なら PNG の Up を使う）
    param (
        [System.Collections.IDictionary]$normal,
        [System.Collections.IDictionary]$packed = @{},
        [switch]$predictor,
        [string]$trailerExtra = '/Root 1 0 R',
        [int[]]$widths = @(1, 4, 2),
        [int]$pngType = 2,
        [switch]$useIndex,
        [switch]$badHeader,
        [switch]$reverseIndex
    )

    $objs = @{}
    foreach ($k in $normal.Keys) { $objs[[int]$k] = $normal[$k] }
    $maxId = ($objs.Keys + $packed.Keys | Measure-Object -Maximum).Maximum
    $objStmId = $maxId + 1
    $xrefId = $maxId + 2
    $packedIds = @($packed.Keys | ForEach-Object { [int]$_ } | Sort-Object)
    if ($packedIds.Count -gt 0) {
        $head = New-Object System.Text.StringBuilder
        $body = New-Object System.Text.StringBuilder
        foreach ($id in $packedIds) {
            [void]$head.Append("$id $($body.Length) ")
            [void]$body.Append($packed[$id] + "`n")
        }
        $headText = $head.ToString()
        if ($badHeader) { $headText = "x" + $headText.Substring(1) }
        $data = toPdfBytes ($headText + $body.ToString())
        $objs[$objStmId] = newPdfStreamObject "/Type /ObjStm /N $($packedIds.Count) /First $($head.Length) /Filter /FlateDecode" (compressPdfBytes $data)
    }

    $ms = New-Object System.IO.MemoryStream
    $h = toPdfBytes "%PDF-1.5`n"
    $ms.Write($h, 0, $h.Length)
    $offsets = @{}
    foreach ($id in ($objs.Keys | Sort-Object)) {
        $offsets[[int]$id] = [int]$ms.Length
        $o = $objs[$id]
        if ($o -is [string]) { $b = joinPdfBytes @("$id 0 obj`n", $o, "`nendobj`n") }
        else { $b = joinPdfBytes @("$id 0 obj`n", $o.Dict, "`nstream`n", $o.Stream, "`nendstream`nendobj`n") }
        $ms.Write($b, 0, $b.Length)
    }
    $xrefPos = [int]$ms.Length
    $offsets[$xrefId] = $xrefPos

    # 行の項目（種類・2 番目・3 番目）を、widths の幅のビッグエンディアンで並べる。useIndex なら、ある番号の行だけを /Index の組で書く
    $rowLength = $widths[0] + $widths[1] + $widths[2]
    $ids = @(0..$xrefId | Where-Object { -not $useIndex -or $_ -eq 0 -or $offsets.ContainsKey($_) -or $packedIds -contains $_ })
    $rows = New-Object System.Collections.ArrayList
    foreach ($i in $ids) {
        $idx = [array]::IndexOf($packedIds, $i)
        if ($idx -ge 0) {
            # reverseIndex: 相互参照の示す位置を逆順にして、番号と合わなくする
            if ($reverseIndex) { $idx = $packedIds.Count - 1 - $idx }
            $fields = @(2, $objStmId, $idx)
        }
        elseif ($offsets.ContainsKey($i)) { $fields = @(1, $offsets[$i], 0) }
        else { $fields = @(0, 0, 0) }
        $row = New-Object byte[] $rowLength
        $pos = 0
        for ($f = 0; $f -lt 3; $f++) {
            for ($j = $widths[$f] - 1; $j -ge 0; $j--) { $row[$pos + $j] = [byte]($fields[$f] -band 255); $fields[$f] = [int]($fields[$f] -shr 8) }
            $pos += $widths[$f]
        }
        [void]$rows.Add($row)
    }
    $raw = New-Object System.IO.MemoryStream
    $prev = New-Object byte[] $rowLength
    foreach ($row in $rows) {
        if ($predictor) {
            # PNG の行の種類（0 なし・1 Sub・2 Up・3 Average・4 Paeth）で符号化する。bpp は 1
            $raw.WriteByte([byte]$pngType)
            $d = New-Object byte[] $rowLength
            for ($j = 0; $j -lt $rowLength; $j++) {
                $left = 0; $upLeft = 0
                if ($j -ge 1) { $left = [int]$row[$j - 1]; $upLeft = [int]$prev[$j - 1] }
                $up = [int]$prev[$j]
                switch ($pngType) {
                    0 { $pred = 0 }
                    1 { $pred = $left }
                    2 { $pred = $up }
                    3 { $pred = [int][Math]::Floor(($left + $up) / 2) }
                    4 {
                        $p = $left + $up - $upLeft
                        $pa = [Math]::Abs($p - $left); $pb = [Math]::Abs($p - $up); $pc = [Math]::Abs($p - $upLeft)
                        if ($pa -le $pb -and $pa -le $pc) { $pred = $left } elseif ($pb -le $pc) { $pred = $up } else { $pred = $upLeft }
                    }
                }
                $d[$j] = [byte](([int]$row[$j] - $pred + 256) % 256)
            }
            $raw.Write($d, 0, $rowLength)
            $prev = $row
        }
        else { $raw.Write($row, 0, $rowLength) }
    }
    $parms = if ($predictor) { " /DecodeParms << /Predictor 12 /Columns $rowLength >>" } else { '' }
    $indexText = ''
    if ($useIndex) {
        $runs = New-Object System.Collections.ArrayList
        $startId = $ids[0]; $count = 1
        for ($k = 1; $k -lt $ids.Count; $k++) {
            if ($ids[$k] -eq $ids[$k - 1] + 1) { $count++ }
            else { [void]$runs.Add("$startId $count"); $startId = $ids[$k]; $count = 1 }
        }
        [void]$runs.Add("$startId $count")
        $indexText = " /Index [$($runs -join ' ')]"
    }
    $xs = newPdfStreamObject "/Type /XRef /W [$($widths -join ' ')] /Size $($xrefId + 1)$indexText /Filter /FlateDecode$parms $trailerExtra" (compressPdfBytes $raw.ToArray())
    $b = joinPdfBytes @("$xrefId 0 obj`n", $xs.Dict, "`nstream`n", $xs.Stream, "`nendstream`nendobj`n", "startxref`n$xrefPos`n%%EOF`n")
    $ms.Write($b, 0, $b.Length)
    return , ($ms.ToArray())
}

# ---- 暗号化された PDF（標準のセキュリティハンドラー）を組み立てる ----
# 読む側（scripts\shared\core\pdf_crypt.ps1）の MD5・RC4 を使わず、.NET の MD5・AES・BigInteger で別に書く（読む側と同じ間違いを共有しないため）。

function newTestRc4([byte[]]$key, [byte[]]$data) {
    $s = 0..255
    $j = 0
    for ($i = 0; $i -lt 256; $i++) {
        $j = ($j + $s[$i] + $key[$i % $key.Length]) % 256
        $t = $s[$i]; $s[$i] = $s[$j]; $s[$j] = $t
    }
    $out = New-Object byte[] $data.Length
    $i = 0; $j = 0
    for ($n = 0; $n -lt $data.Length; $n++) {
        $i = ($i + 1) % 256
        $j = ($j + $s[$i]) % 256
        $t = $s[$i]; $s[$i] = $s[$j]; $s[$j] = $t
        $out[$n] = $data[$n] -bxor $s[($s[$i] + $s[$j]) % 256]
    }
    return , $out
}

function getTestDigest([string]$name, [byte[]]$data) {
    $algorithm = [System.Security.Cryptography.HashAlgorithm]::Create($name)
    try { return , $algorithm.ComputeHash($data) } finally { $algorithm.Dispose() }
}

function invokeTestAes([byte[]]$key, [byte[]]$iv, [byte[]]$data, [bool]$pkcs7) {
    # 暗号化。pkcs7 なら詰め物を付ける（なしなら data は 16 の倍数）
    $aes = [System.Security.Cryptography.Aes]::Create()
    try {
        $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
        $aes.Padding = if ($pkcs7) { [System.Security.Cryptography.PaddingMode]::PKCS7 } else { [System.Security.Cryptography.PaddingMode]::None }
        $aes.Key = $key
        $aes.IV = $iv
        $enc = $aes.CreateEncryptor()
        try { return , $enc.TransformFinalBlock($data, 0, $data.Length) } finally { $enc.Dispose() }
    } finally { $aes.Dispose() }
}

function getTestHashR56([int]$revision, [byte[]]$password, [byte[]]$salt) {
    # R5 は SHA-256 だけ。R6 は ISO 32000-2 の 2.B（先頭 16 バイトを大きな数として 3 で割った余りで SHA の種類を選ぶ）
    $k = getTestDigest 'SHA256' (joinPdfBytes @($password, $salt))
    if ($revision -lt 6) { return , $k }
    $round = 0
    do {
        $unit = joinPdfBytes @($password, $k)
        $k1 = joinPdfBytes @(1..64 | ForEach-Object { , $unit })
        $e = invokeTestAes $k[0..15] $k[16..31] $k1 $false
        $head = [byte[]]($e[0..15])
        [array]::Reverse($head)
        $number = New-Object System.Numerics.BigInteger (, ([byte[]]($head + [byte]0)))
        $mod = [int]($number % 3)
        $k = getTestDigest @('SHA256', 'SHA384', 'SHA512')[$mod] $e
        $round++
    } while ($round -lt 64 -or $e[$e.Length - 1] -gt $round - 32)
    return , ([byte[]]($k[0..31]))
}

function newEncryptedPdf {
    # method: rc40（V1 R2）・rc128（V2 R3）・aes128（V4 R4）・aes256r5・aes256r6。userPassword が空でなければ、空のパスワードでは開けない。
    # encryptDict を渡すと、Encrypt の辞書をそれに替える（鍵と暗号は method のまま）。xrefStream なら相互参照をストリームにする
    param (
        [string[]]$pages,
        [string]$method = 'aes256r6',
        [string]$userPassword = '',
        [switch]$compress,
        [string]$encryptDict = '',
        [switch]$xrefStream
    )

    $pad = [byte[]]@(0x28, 0xBF, 0x4E, 0x5E, 0x4E, 0x75, 0x8A, 0x41, 0x64, 0x00, 0x4E, 0x56, 0xFF, 0xFA, 0x01, 0x08, 0x2E, 0x2E, 0x00, 0xB6, 0xD0, 0x68, 0x3E, 0x80, 0x2F, 0x0C, 0xA9, 0xFE, 0x64, 0x53, 0x69, 0x7A)
    $hex = { param([byte[]]$b) '<' + (($b | ForEach-Object { $_.ToString('x2') }) -join '') + '>' }
    $id = [byte[]](1..16 | ForEach-Object { $_ * 7 })
    $iv = [byte[]](1..16 | ForEach-Object { 200 - $_ })
    $pw = $script:pdfTestLatin1.GetBytes($userPassword)

    $mode = 'rc4'
    if ($method -eq 'rc40') { $v = 1; $r = 2; $n = 5; $dict = '/Filter /Standard /V 1 /R 2' }
    elseif ($method -eq 'rc128') { $v = 2; $r = 3; $n = 16; $dict = '/Filter /Standard /V 2 /R 3 /Length 128' }
    elseif ($method -eq 'aes128') {
        $v = 4; $r = 4; $n = 16; $mode = 'aesv2'
        $dict = '/Filter /Standard /V 4 /R 4 /Length 128 /CF << /StdCF << /CFM /AESV2 /AuthEvent /DocOpen /Length 16 >> >> /StmF /StdCF /StrF /StdCF'
    }
    else {
        $r = if ($method -eq 'aes256r5') { 5 } else { 6 }
        $v = 5; $n = 32; $mode = 'aesv3'
        $dict = "/Filter /Standard /V 5 /R $r /Length 256 /CF << /StdCF << /CFM /AESV3 /AuthEvent /DocOpen /Length 32 >> >> /StmF /StdCF /StrF /StdCF"
    }

    if ($v -lt 5) {
        $o = [byte[]](1..32 | ForEach-Object { $_ + 100 })
        $p = -4
        $pBytes = [byte[]](0xFC, 0xFF, 0xFF, 0xFF)
        $padded = [byte[]](($pw + $pad)[0..31])
        $h = getTestDigest 'MD5' (joinPdfBytes @($padded, $o, $pBytes, $id))
        if ($r -ge 3) { for ($i = 0; $i -lt 50; $i++) { $h = getTestDigest 'MD5' ([byte[]]$h[0..($n - 1)]) } }
        $key = [byte[]]$h[0..($n - 1)]
        if ($r -eq 2) {
            $u = newTestRc4 $key $pad
        }
        else {
            $x = newTestRc4 $key (getTestDigest 'MD5' (joinPdfBytes @($pad, $id)))
            for ($i = 1; $i -le 19; $i++) { $x = newTestRc4 ([byte[]]($key | ForEach-Object { $_ -bxor $i })) $x }
            $u = joinPdfBytes @($x, [byte[]](1..16))
        }
        $dict += " /O $(& $hex $o) /U $(& $hex $u) /P $p"
        $trailerId = "/ID [$(& $hex $id) $(& $hex $id)]"
    }
    else {
        $key = [byte[]](1..32 | ForEach-Object { ($_ * 11) % 256 })
        $vsalt = [byte[]](21..28)
        $ksalt = [byte[]](31..38)
        $u = joinPdfBytes @((getTestHashR56 $r $pw $vsalt), $vsalt, $ksalt)
        $ue = invokeTestAes (getTestHashR56 $r $pw $ksalt) (New-Object byte[] 16) $key $false
        $dict += " /O $(& $hex ([byte[]](1..48))) /U $(& $hex $u) /OE $(& $hex ([byte[]](1..32))) /UE $(& $hex $ue) /Perms $(& $hex ([byte[]](1..16))) /P -4"
        $trailerId = "/ID [$(& $hex $id) $(& $hex $id)]"
    }

    $protect = {
        param([int]$num, [byte[]]$data)
        if ($mode -eq 'aesv3') { $objectKey = $key }
        else {
            $suffix = [byte[]]@(($num -band 255), (($num -shr 8) -band 255), (($num -shr 16) -band 255), 0, 0)
            if ($mode -eq 'aesv2') { $suffix = [byte[]]($suffix + [byte[]](0x73, 0x41, 0x6C, 0x54)) }
            $objectKey = [byte[]](getTestDigest 'MD5' (joinPdfBytes @($key, $suffix)))[0..([Math]::Min($key.Length + 5, 16) - 1)]
        }
        if ($mode -eq 'rc4') { return , (newTestRc4 $objectKey $data) }
        return , (joinPdfBytes @($iv, (invokeTestAes $objectKey $iv $data $true)))
    }

    $objs = @{}
    $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
    $kids = (0..($pages.Count - 1) | ForEach-Object { "$(10 + $_ * 2) 0 R" }) -join ' '
    $objs[2] = "<< /Type /Pages /Kids [$kids] /Count $($pages.Count) >>"
    $objs[3] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>'
    if ($encryptDict -ne '') { $objs[5] = "<< $encryptDict >>" } else { $objs[5] = "<< $dict >>" }
    for ($i = 0; $i -lt $pages.Count; $i++) {
        $pageId = 10 + $i * 2
        $contentId = $pageId + 1
        $objs[$pageId] = "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R >> >> /Contents $contentId 0 R >>"
        $data = toPdfBytes $pages[$i]
        $flate = ''
        if ($compress) { $data = compressPdfBytes $data; $flate = '/Filter /FlateDecode' }
        $objs[$contentId] = newPdfStreamObject $flate (& $protect $contentId $data)
    }
    $extra = "/Root 1 0 R /Encrypt 5 0 R $trailerId"
    if ($xrefStream) { return , (newPdfWithXrefStream -normal $objs -trailerExtra $extra) }
    return , (newPdfFromObjects $objs $extra)
}
