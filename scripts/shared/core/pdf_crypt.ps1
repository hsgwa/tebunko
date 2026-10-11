# PDF の暗号化（標準のセキュリティハンドラー）を、空の利用者パスワードで復号する部品。PowerShell だけで書く（外の部品・Office を使わない）。
# 画面にもファイルにも触らない。pdf_file.ps1 の readPdfPages が、ストリームを取り出すときに使う。
#   読む    : AES-256（V5・R5・R6）。開くのに利用者パスワードが要らない（空のパスワードで鍵が作れて、/U の検証用のハッシュと合う）PDF だけ。
#   読まない: パスワードが要る PDF、AES-256 以外（RC4 40・128 ビット、AES-128 など。鍵の計算に旧式のハッシュが要り、FIPS 準拠を強制した機械で使えないため、
#             自前でも書かない。パスワードが要るかどうかも判定せず、必ず「方式に対応していない」にする）。例外にする（そのファイルだけ失敗）。
# 復号するのはストリームだけ（内容・フォントの対応表・オブジェクトのストリーム）。文字列は、読む内容がストリームの中にあるため復号しない。
# 使うのは SHA-2 と AES だけで、FIPS でも使える *CryptoServiceProvider の型を使う（SHA256Managed・AesManaged は FIPS で例外になる）。
# 鍵・途中の値は、例外の文にも記録にも出さない。壊れた辞書は、決まった長さの確かめと、R6 の繰り返しの上限（300 回。仕様では最大でも 288 回で終わる）で止める。
# 例外の文言は「PDF」で始める（readPdfPages は、それ以外の例外を「壊れている」に言い換える）。

# R6 の鍵の計算の繰り返しの上限（仕様では最後のバイトで終わる。最大でも 64 + 224 = 288 回で必ず終わる）
${pdfCryptR6MaxRounds} = 300

# 対応していない方式の例外の文
${pdfCryptUnsupported} = "PDFの暗号化の方式に対応していないため読めません。"

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

function getPdfSlice {
    # バイト列の一部を新しい配列に写す
    param (
        [byte[]]$bytes,
        [int]$offset,
        [int]$length
    )

    $out = New-Object byte[] $length
    [Array]::Copy($bytes, $offset, $out, 0, $length)
    return , $out
}

function getPdfKeyR56 {
    # R5・R6 のファイルの鍵を、空の利用者パスワードで作る。/U の検証用のハッシュと合わなければ「パスワードが必要」の例外。
    # userBytes は /U（48 バイト以上）、userExtra は /UE（32 バイト以上）
    param (
        [int]$revision,
        [byte[]]$userBytes,
        [byte[]]$userExtra
    )

    $empty = New-Object byte[] 0
    $check = getPdfCryptHashR56 $revision $empty (getPdfSlice $userBytes 32 8) $empty
    $same = $true
    for ($i = 0; $i -lt 32; $i++) { if ($check[$i] -ne $userBytes[$i]) { $same = $false } }
    if (-not $same) { throw "PDFにパスワードが必要なため読めません。" }
    $intermediate = getPdfCryptHashR56 $revision $empty (getPdfSlice $userBytes 40 8) $empty
    return , (invokePdfAesCbc $intermediate (New-Object byte[] 16) (getPdfSlice $userExtra 0 32) $false)
}

function getPdfCryptMethod {
    # 暗号化フィルター（CF の項目）の方式を返す。AES-256（/AESV3）は AESV3、/None と /Identity は復号しない（None）。
    # それ以外は、対応しない方式として例外にする
    param (
        [hashtable]$doc,
        $cf,
        [string]$name
    )

    if ($name -ceq "/Identity") { return "None" }
    if (-not (testPdfDict $cf) -or $name.Length -lt 2) {
        throw ${pdfCryptUnsupported}
    }
    $filter = resolvePdf $doc $cf[$name.Substring(1)]
    if (-not (testPdfDict $filter)) {
        throw ${pdfCryptUnsupported}
    }
    $cfm = resolvePdf $doc $filter["CFM"]
    if ($cfm -ceq "/None") { return "None" }
    if ($cfm -ceq "/AESV3") { return "AESV3" }
    throw ${pdfCryptUnsupported}
}

function newPdfCrypt {
    # Encrypt の辞書（解決済み）から、空の利用者パスワードで鍵を作り、/U と照らして合えば復号の設定を返す。
    # 返す値: @{ Key; StmMethod; Cf }。AES-256（V5・R5・R6）だけを読む。
    # パスワードが要る（空のパスワードが /U と合わない）・対応しない方式・辞書が壊れているときは、そのファイルだけ例外にする
    param (
        [hashtable]$doc,
        [hashtable]$enc
    )

    if ((resolvePdf $doc $enc["Filter"]) -cne "/Standard") { throw ${pdfCryptUnsupported} }
    $v = resolvePdf $doc $enc["V"]
    $r = resolvePdf $doc $enc["R"]
    if ($v -isnot [double] -or $r -isnot [double]) { throw ${pdfCryptUnsupported} }
    if ([int]$v -ne 5 -or ([int]$r -ne 5 -and [int]$r -ne 6)) { throw ${pdfCryptUnsupported} }
    $r = [int]$r

    # 復号の方式（ストリーム）
    $cf = resolvePdf $doc $enc["CF"]
    $stmName = resolvePdf $doc $enc["StmF"]
    if ($null -eq $stmName) { $stmName = "/Identity" }
    if ($stmName -isnot [string]) { throw ${pdfCryptUnsupported} }
    $method = getPdfCryptMethod $doc $cf $stmName

    $userBytes = getPdfStringBytes (resolvePdf $doc $enc["U"])
    $userExtra = getPdfStringBytes (resolvePdf $doc $enc["UE"])
    if ($null -eq $userBytes -or $userBytes.Length -lt 48 -or $null -eq $userExtra -or $userExtra.Length -lt 32) {
        throw "PDFの暗号化の情報が壊れているため読めません。"
    }
    $key = getPdfKeyR56 $r $userBytes $userExtra
    return @{ Key = $key; StmMethod = $method; Cf = $cf }
}

function getPdfStreamCryptMethod {
    # ストリーム 1 つの復号の方式（AESV3・None）を返す。相互参照のストリームは暗号化されない。
    # ストリームごとの /Crypt フィルターがあれば、その名前の方式（/Identity は復号しない）にする
    param (
        [hashtable]$doc,
        [hashtable]$stream
    )

    $crypt = $doc.Crypt
    $dict = $stream.Dict
    if ((resolvePdf $doc $dict["Type"]) -ceq "/XRef") { return "None" }
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
                if ($name -isnot [string]) { throw ${pdfCryptUnsupported} }
                return (getPdfCryptMethod $doc $crypt.Cf $name)
            }
        }
    }
    return $crypt.StmMethod
}

function unprotectPdfData {
    # ストリームの中身（圧縮を戻す前）を復号する。AES-256 はファイルの鍵をそのまま使う
    param (
        [hashtable]$crypt,
        [string]$method,
        [byte[]]$data
    )

    if ($method -ceq "None") { return , $data }
    # AES-CBC。先頭 16 バイトが初期ベクトル、最後に PKCS#7 の詰め物
    if ($data.Length -eq 0) { return , $data }
    if ($data.Length -lt 32 -or $data.Length % 16 -ne 0) {
        throw "PDFの暗号化されたデータが壊れているため読めません。"
    }
    $plain = invokePdfAesCbc $crypt.Key (getPdfSlice $data 0 16) (getPdfSlice $data 16 ($data.Length - 16)) $false
    $pad = $plain[$plain.Length - 1]
    if ($pad -lt 1 -or $pad -gt 16) {
        throw "PDFの暗号化されたデータが壊れているため読めません。"
    }
    return , (getPdfSlice $plain 0 ($plain.Length - $pad))
}
