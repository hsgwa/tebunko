# PDF の暗号化（標準のセキュリティハンドラー）を、空の利用者パスワードで復号する部品。PowerShell だけで書く（外の部品・Office を使わない）。
# 画面にもファイルにも触らない。pdf_file.ps1 の readPdfPages が、ストリームを取り出すときに使う。
#   読む    : RC4 40 ビット（V1・R2）、RC4 128 ビット（V2・R3）、AES-128（V4・R4）、AES-256（V5・R5・R6）。
#             開くのに利用者パスワードが要らない（空のパスワードで鍵が作れて、/U と合う）PDF だけ。
#   読まない: パスワードが要る PDF、上の方式以外（独自のフィルター・未知の方式）。例外にする（そのファイルだけ失敗）。
# 復号するのはストリームだけ（内容・フォントの対応表・オブジェクトのストリーム）。文字列は、読む内容がストリームの中にあるため復号しない。
# PDF の仕様が求める旧式の 128 ビットのハッシュ（RFC 1321）と RC4 は自前で書く。.NET の同じ働きの型は、FIPS 準拠を強制した機械では使えず、4 方式とも読めなくなるため。
# SHA-2 と AES は、FIPS でも使える *CryptoServiceProvider の型を使う（SHA256Managed・AesManaged は FIPS で例外になる）。
# 鍵・途中の値は、例外の文にも記録にも出さない。壊れた辞書は、決まった長さの確かめと、R6 の繰り返しの上限（288 回）で止める。
# 例外の文言は「PDF」で始める（readPdfPages は、それ以外の例外を「壊れている」に言い換える）。

# 利用者パスワードを鍵の元にするときの詰め物（PDF 32000-1 の表の Algorithm 2）
${pdfPasswordPad} = [byte[]]@(
    0x28, 0xBF, 0x4E, 0x5E, 0x4E, 0x75, 0x8A, 0x41, 0x64, 0x00, 0x4E, 0x56, 0xFF, 0xFA, 0x01, 0x08,
    0x2E, 0x2E, 0x00, 0xB6, 0xD0, 0x68, 0x3E, 0x80, 0x2F, 0x0C, 0xA9, 0xFE, 0x64, 0x53, 0x69, 0x7A
)

# RFC 1321 のハッシュの定数。K[i] = floor(2^32 * abs(sin(i + 1)))、s は各ラウンドの回転量
${pdfDigestK} = [long[]]@(0..63 | ForEach-Object { [long][Math]::Floor([Math]::Abs([Math]::Sin($_ + 1)) * 4294967296.0) })
${pdfDigestS} = [int[]]@(
    7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22,
    5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20,
    4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23,
    6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21
)

# R6 の鍵の計算の繰り返しの上限（仕様では最後のバイトで終わる。最大でも 64 + 224 = 288 回で必ず終わる）
${pdfCryptR6MaxRounds} = 300

function joinPdfCryptBytes {
    # バイト列を順につなげる
    param (
        [object[]]$parts
    )

    $ms = [System.IO.MemoryStream]::new()
    foreach ($part in $parts) {
        if ($null -eq $part) { continue }
        $bytes = [byte[]]$part
        $ms.Write($bytes, 0, $bytes.Length)
    }
    return , $ms.ToArray()
}

function getPdfRfc1321Hash {
    # RFC 1321 のハッシュ。32 ビットの値は long に入れ、そのつど 32 ビットに切る
    param (
        [byte[]]$data
    )

    $mask = [long]4294967295
    $length = $data.Length
    $total = ([int][Math]::Floor(($length + 8) / 64) + 1) * 64
    $buffer = New-Object byte[] $total
    [Array]::Copy($data, 0, $buffer, 0, $length)
    $buffer[$length] = 0x80
    $bits = [long]$length * 8
    for ($i = 0; $i -lt 8; $i++) {
        $buffer[$total - 8 + $i] = [byte](($bits -shr (8 * $i)) -band 255)
    }
    $a0 = 0x67452301L; $b0 = 0xEFCDAB89L; $c0 = 0x98BADCFEL; $d0 = 0x10325476L
    $m = New-Object long[] 16
    for ($block = 0; $block -lt $total; $block += 64) {
        for ($j = 0; $j -lt 16; $j++) {
            $o = $block + $j * 4
            $m[$j] = [long]$buffer[$o] -bor ([long]$buffer[$o + 1] -shl 8) -bor ([long]$buffer[$o + 2] -shl 16) -bor ([long]$buffer[$o + 3] -shl 24)
        }
        $a = $a0; $b = $b0; $c = $c0; $d = $d0
        for ($i = 0; $i -lt 64; $i++) {
            if ($i -lt 16) {
                $f = ($b -band $c) -bor ((-bnot $b) -band $d)
                $g = $i
            } elseif ($i -lt 32) {
                $f = ($d -band $b) -bor ((-bnot $d) -band $c)
                $g = (5 * $i + 1) % 16
            } elseif ($i -lt 48) {
                $f = $b -bxor $c -bxor $d
                $g = (3 * $i + 5) % 16
            } else {
                $f = ($c -bxor ($b -bor (-bnot $d))) -band $mask
                $g = (7 * $i) % 16
            }
            $f = ($f + $a + ${pdfDigestK}[$i] + $m[$g]) -band $mask
            $a = $d
            $d = $c
            $c = $b
            $s = ${pdfDigestS}[$i]
            $b = ($b + ((($f -shl $s) -bor ($f -shr (32 - $s))) -band $mask)) -band $mask
        }
        $a0 = ($a0 + $a) -band $mask
        $b0 = ($b0 + $b) -band $mask
        $c0 = ($c0 + $c) -band $mask
        $d0 = ($d0 + $d) -band $mask
    }
    $result = New-Object byte[] 16
    $words = @($a0, $b0, $c0, $d0)
    for ($w = 0; $w -lt 4; $w++) {
        for ($k = 0; $k -lt 4; $k++) {
            $result[$w * 4 + $k] = [byte](($words[$w] -shr (8 * $k)) -band 255)
        }
    }
    return , $result
}

function invokePdfRc4 {
    # RC4（暗号化と復号は同じ）
    param (
        [byte[]]$key,
        [byte[]]$data
    )

    $s = New-Object int[] 256
    for ($i = 0; $i -lt 256; $i++) { $s[$i] = $i }
    $keyLength = $key.Length
    $j = 0
    for ($i = 0; $i -lt 256; $i++) {
        $j = ($j + $s[$i] + $key[$i % $keyLength]) -band 255
        $t = $s[$i]; $s[$i] = $s[$j]; $s[$j] = $t
    }
    $out = New-Object byte[] $data.Length
    $i = 0
    $j = 0
    for ($n = 0; $n -lt $data.Length; $n++) {
        $i = ($i + 1) -band 255
        $j = ($j + $s[$i]) -band 255
        $t = $s[$i]; $s[$i] = $s[$j]; $s[$j] = $t
        $out[$n] = $data[$n] -bxor $s[($s[$i] + $s[$j]) -band 255]
    }
    return , $out
}

function getPdfSha {
    # SHA-256・384・512
    param (
        [int]$bits,
        [byte[]]$data
    )

    if ($bits -eq 256) { $hash = New-Object System.Security.Cryptography.SHA256CryptoServiceProvider }
    elseif ($bits -eq 384) { $hash = New-Object System.Security.Cryptography.SHA384CryptoServiceProvider }
    else { $hash = New-Object System.Security.Cryptography.SHA512CryptoServiceProvider }
    try {
        return , $hash.ComputeHash($data)
    } finally {
        $hash.Dispose()
    }
}

function invokePdfAesCbc {
    # AES-CBC（詰め物なし。data は 16 の倍数）
    param (
        [byte[]]$key,
        [byte[]]$iv,
        [byte[]]$data,
        [bool]$encrypt
    )

    $aes = New-Object System.Security.Cryptography.AesCryptoServiceProvider
    try {
        $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
        $aes.Padding = [System.Security.Cryptography.PaddingMode]::None
        $aes.Key = $key
        $aes.IV = $iv
        if ($encrypt) { $transform = $aes.CreateEncryptor() } else { $transform = $aes.CreateDecryptor() }
        try {
            return , $transform.TransformFinalBlock($data, 0, $data.Length)
        } finally {
            $transform.Dispose()
        }
    } finally {
        $aes.Dispose()
    }
}

function getPdfCryptHashR56 {
    # R5・R6 の利用者パスワードのハッシュ（PDF 2.0 の Algorithm 2.B。R5 は SHA-256 の 1 回だけ）。32 バイトを返す
    param (
        [int]$revision,
        [byte[]]$passBytes,
        [byte[]]$salt,
        [byte[]]$userData
    )

    $k = getPdfSha 256 (joinPdfCryptBytes @($passBytes, $salt, $userData))
    if ($revision -lt 6) { return , $k }
    $round = 0
    do {
        $unit = joinPdfCryptBytes @($passBytes, $k, $userData)
        $repeated = New-Object byte[] ($unit.Length * 64)
        for ($j = 0; $j -lt 64; $j++) {
            [Array]::Copy($unit, 0, $repeated, $j * $unit.Length, $unit.Length)
        }
        $aesKey = New-Object byte[] 16
        $aesIv = New-Object byte[] 16
        [Array]::Copy($k, 0, $aesKey, 0, 16)
        [Array]::Copy($k, 16, $aesIv, 0, 16)
        $e = invokePdfAesCbc $aesKey $aesIv $repeated $true
        # 先頭 16 バイトを 1 つの大きな数とみなして 3 で割った余り（256 を 3 で割ると 1 余るので、バイトの和の余りと同じ）
        $sum = 0
        for ($j = 0; $j -lt 16; $j++) { $sum += $e[$j] }
        $k = getPdfSha (256 + 128 * ($sum % 3)) $e
        $round++
    } while (($round -lt 64 -or $e[$e.Length - 1] -gt $round - 32) -and $round -lt ${pdfCryptR6MaxRounds})
    $result = New-Object byte[] 32
    [Array]::Copy($k, 0, $result, 0, 32)
    return , $result
}

function getPdfStringBytes {
    # PDF の文字列（@{PdfType="str"}）をバイト列にする。文字列でなければ $null
    param (
        $value
    )

    if ($value -is [hashtable] -and $value.ContainsKey("PdfType") -and $value["PdfType"] -eq "str") {
        return , (${pdfLatin1}.GetBytes([string]$value["V"]))
    }
    return $null
}

function getPdfCryptMethod {
    # 暗号化フィルター（CF の項目）の方式を返す。V4 は RC4（/V2）・AES-128（/AESV2）、V5 は AES-256（/AESV3）。
    # /None は復号しない（None）。それ以外は、対応しない方式として例外にする
    param (
        [hashtable]$doc,
        $cf,
        [string]$name,
        [int]$version
    )

    if ($name -ceq "/Identity") { return "None" }
    if (-not (testPdfDict $cf) -or $name.Length -lt 2) {
        throw "PDFの暗号化の方式に対応していないため読めません。"
    }
    $filter = resolvePdf $doc $cf[$name.Substring(1)]
    if (-not (testPdfDict $filter)) {
        throw "PDFの暗号化の方式に対応していないため読めません。"
    }
    $cfm = resolvePdf $doc $filter["CFM"]
    if ($cfm -ceq "/None") { return "None" }
    if ($version -eq 4 -and $cfm -ceq "/V2") { return "RC4" }
    if ($version -eq 4 -and $cfm -ceq "/AESV2") { return "AESV2" }
    if ($version -eq 5 -and $cfm -ceq "/AESV3") { return "AESV3" }
    throw "PDFの暗号化の方式に対応していないため読めません。"
}

function newPdfCrypt {
    # Encrypt の辞書（解決済み）から、空の利用者パスワードで鍵を作り、/U と照らして合えば復号の設定を返す。
    # 返す値: @{ Key; Version; StmMethod; Cf; EncryptMetadata }。
    # パスワードが要る（空のパスワードが /U と合わない）・対応しない方式・辞書が壊れているときは、そのファイルだけ例外にする
    param (
        [hashtable]$doc,
        [hashtable]$enc,
        [string]$id
    )

    $unsupported = "PDFの暗号化の方式に対応していないため読めません。"
    $broken = "PDFの暗号化の情報が壊れているため読めません。"
    if ((resolvePdf $doc $enc["Filter"]) -cne "/Standard") { throw $unsupported }
    $v = resolvePdf $doc $enc["V"]
    $r = resolvePdf $doc $enc["R"]
    if ($v -isnot [double] -or $r -isnot [double]) { throw $unsupported }
    $v = [int]$v
    $r = [int]$r
    $supported = (($v -eq 1 -or $v -eq 2) -and ($r -eq 2 -or $r -eq 3)) -or ($v -eq 4 -and $r -eq 4) -or ($v -eq 5 -and ($r -eq 5 -or $r -eq 6))
    if (-not $supported) { throw $unsupported }

    # 鍵の長さ（バイト）
    if ($v -eq 1) {
        $n = 5
    } elseif ($v -eq 2) {
        $bitLength = resolvePdf $doc $enc["Length"]
        if ($null -eq $bitLength) { $bitLength = 40.0 }
        if ($bitLength -isnot [double] -or $bitLength -lt 40 -or $bitLength -gt 128 -or $bitLength % 8 -ne 0) { throw $unsupported }
        $n = [int]($bitLength / 8)
    } elseif ($v -eq 4) {
        $n = 16
    } else {
        $n = 32
    }

    $ownerBytes = getPdfStringBytes (resolvePdf $doc $enc["O"])
    $userBytes = getPdfStringBytes (resolvePdf $doc $enc["U"])
    $encryptMetadata = $true
    if ((resolvePdf $doc $enc["EncryptMetadata"]) -eq $false) { $encryptMetadata = $false }

    # 復号の方式（ストリーム）
    $cf = $null
    $method = "RC4"
    if ($v -ge 4) {
        $cf = resolvePdf $doc $enc["CF"]
        $stmName = resolvePdf $doc $enc["StmF"]
        if ($null -eq $stmName) { $stmName = "/Identity" }
        if ($stmName -isnot [string]) { throw $unsupported }
        $method = getPdfCryptMethod $doc $cf $stmName $v
    }

    $key = $null
    if ($r -ge 5) {
        $userExtra = getPdfStringBytes (resolvePdf $doc $enc["UE"])
        if ($null -eq $userBytes -or $userBytes.Length -lt 48 -or $null -eq $userExtra -or $userExtra.Length -lt 32) { throw $broken }
        $empty = New-Object byte[] 0
        $validation = New-Object byte[] 8
        $keySalt = New-Object byte[] 8
        [Array]::Copy($userBytes, 32, $validation, 0, 8)
        [Array]::Copy($userBytes, 40, $keySalt, 0, 8)
        $check = getPdfCryptHashR56 $r $empty $validation $empty
        $same = $true
        for ($i = 0; $i -lt 32; $i++) { if ($check[$i] -ne $userBytes[$i]) { $same = $false } }
        if (-not $same) { throw "PDFにパスワードが必要なため読めません。" }
        $intermediate = getPdfCryptHashR56 $r $empty $keySalt $empty
        $ue = New-Object byte[] 32
        [Array]::Copy($userExtra, 0, $ue, 0, 32)
        $key = invokePdfAesCbc $intermediate (New-Object byte[] 16) $ue $false
    } else {
        if ($null -eq $ownerBytes -or $ownerBytes.Length -lt 32 -or $null -eq $userBytes -or $userBytes.Length -lt 32) { throw $broken }
        $p = resolvePdf $doc $enc["P"]
        if ($p -isnot [double]) { throw $broken }
        $pBytes = New-Object byte[] 4
        $pLong = [long]$p
        for ($i = 0; $i -lt 4; $i++) { $pBytes[$i] = [byte](($pLong -shr (8 * $i)) -band 255) }
        $owner32 = New-Object byte[] 32
        [Array]::Copy($ownerBytes, 0, $owner32, 0, 32)
        $idBytes = ${pdfLatin1}.GetBytes($id)
        $parts = @(${pdfPasswordPad}, $owner32, $pBytes, $idBytes)
        if ($r -ge 4 -and -not $encryptMetadata) { $parts += , ([byte[]]@(255, 255, 255, 255)) }
        $hash = getPdfRfc1321Hash (joinPdfCryptBytes $parts)
        if ($r -ge 3) {
            for ($i = 0; $i -lt 50; $i++) {
                $head = New-Object byte[] $n
                [Array]::Copy($hash, 0, $head, 0, $n)
                $hash = getPdfRfc1321Hash $head
            }
        }
        $key = New-Object byte[] $n
        [Array]::Copy($hash, 0, $key, 0, $n)

        # /U と照らす（Algorithm 4・5）
        if ($r -eq 2) {
            $expected = invokePdfRc4 $key ${pdfPasswordPad}
            $compare = 32
        } else {
            $x = invokePdfRc4 $key (getPdfRfc1321Hash (joinPdfCryptBytes @(${pdfPasswordPad}, $idBytes)))
            for ($i = 1; $i -le 19; $i++) {
                $step = New-Object byte[] $n
                for ($j = 0; $j -lt $n; $j++) { $step[$j] = $key[$j] -bxor $i }
                $x = invokePdfRc4 $step $x
            }
            $expected = $x
            $compare = 16
        }
        $same = $true
        for ($i = 0; $i -lt $compare; $i++) { if ($expected[$i] -ne $userBytes[$i]) { $same = $false } }
        if (-not $same) { throw "PDFにパスワードが必要なため読めません。" }
    }
    return @{ Key = $key; Version = $v; StmMethod = $method; Cf = $cf; EncryptMetadata = $encryptMetadata }
}

function getPdfStreamCryptMethod {
    # ストリーム 1 つの復号の方式（RC4・AESV2・AESV3・None）を返す。
    # 相互参照のストリームは暗号化されない。EncryptMetadata が false のメタデータも暗号化されない。
    # ストリームごとの /Crypt フィルターがあれば、その名前の方式（/Identity は復号しない）にする
    param (
        [hashtable]$doc,
        [hashtable]$stream
    )

    $crypt = $doc.Crypt
    $dict = $stream.Dict
    $type = resolvePdf $doc $dict["Type"]
    if ($type -ceq "/XRef") { return "None" }
    if ($type -ceq "/Metadata" -and -not $crypt.EncryptMetadata) { return "None" }
    $filters = resolvePdf $doc $dict["Filter"]
    if ($null -ne $filters) {
        $filterList = @($filters)
        $paramsList = @(resolvePdf $doc $dict["DecodeParms"])
        for ($i = 0; $i -lt $filterList.Count; $i++) {
            if ((resolvePdf $doc $filterList[$i]) -ceq "/Crypt") {
                $name = "/Identity"
                if ($i -lt $paramsList.Count) {
                    $params = resolvePdf $doc $paramsList[$i]
                    if ((testPdfDict $params) -and $null -ne (resolvePdf $doc $params["Name"])) { $name = resolvePdf $doc $params["Name"] }
                }
                if ($name -isnot [string]) { throw "PDFの暗号化の方式に対応していないため読めません。" }
                return (getPdfCryptMethod $doc $crypt.Cf $name $crypt.Version)
            }
        }
    }
    return $crypt.StmMethod
}

function unprotectPdfData {
    # ストリームの中身（圧縮を戻す前）を復号する。num・gen は、ストリームのオブジェクトの番号と世代
    param (
        [hashtable]$crypt,
        [string]$method,
        [int]$num,
        [int]$gen,
        [byte[]]$data
    )

    if ($method -ceq "None") { return , $data }
    $key = $crypt.Key
    if ($method -cne "AESV3") {
        # オブジェクトごとの鍵（Algorithm 1）。ハッシュ(鍵 + 番号 3 バイト + 世代 2 バイト [+ AES は "sAlT"]) の先頭 min(n + 5, 16) バイト
        $suffix = New-Object byte[] 5
        $suffix[0] = [byte]($num -band 255); $suffix[1] = [byte](($num -shr 8) -band 255); $suffix[2] = [byte](($num -shr 16) -band 255)
        $suffix[3] = [byte]($gen -band 255); $suffix[4] = [byte](($gen -shr 8) -band 255)
        $parts = @($key, $suffix)
        if ($method -ceq "AESV2") { $parts += , ([byte[]]@(0x73, 0x41, 0x6C, 0x54)) }
        $hash = getPdfRfc1321Hash (joinPdfCryptBytes $parts)
        $size = [Math]::Min($key.Length + 5, 16)
        $key = New-Object byte[] $size
        [Array]::Copy($hash, 0, $key, 0, $size)
    }
    if ($method -ceq "RC4") { return , (invokePdfRc4 $key $data) }
    # AES-CBC。先頭 16 バイトが初期ベクトル、最後に PKCS#7 の詰め物
    if ($data.Length -eq 0) { return , $data }
    if ($data.Length -lt 32 -or $data.Length % 16 -ne 0) {
        throw "PDFの暗号化されたデータが壊れているため読めません。"
    }
    $iv = New-Object byte[] 16
    [Array]::Copy($data, 0, $iv, 0, 16)
    $body = New-Object byte[] ($data.Length - 16)
    [Array]::Copy($data, 16, $body, 0, $body.Length)
    $plain = invokePdfAesCbc $key $iv $body $false
    $pad = $plain[$plain.Length - 1]
    if ($pad -lt 1 -or $pad -gt 16) {
        throw "PDFの暗号化されたデータが壊れているため読めません。"
    }
    $result = New-Object byte[] ($plain.Length - $pad)
    [Array]::Copy($plain, 0, $result, 0, $result.Length)
    return , $result
}
