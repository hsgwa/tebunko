# 暗号化されたファイルの種類を、開かずに見分ける（判断層）。ファイル・COM・Officeには触らない。
# 入力は「先頭バイト列」と「複合ドキュメント形式（CFB）のディレクトリのエントリ名の配列（CFBでなければ $null）」だけ。
# 実際にファイルを読むのは状態層（office_protection.ps1）。
#
# 種類（getOfficeProtectionKind の戻り値）:
#   Zip      : 新形式（.docx 等）。ふつうに読む
#   Password : パスワード付き（新形式）。Word・PowerPoint は Office を使わずに失敗にする（Excel は今のまま Office に任せる）
#   Rights   : IRM・秘密度ラベルの暗号化。Office で開かない（ライセンス取得・サインイン画面を防ぐ）
#   Legacy   : 旧形式（.doc / .xls / .ppt）など、権限保護と分からないCFB。今と同じ（Officeで開く。ダミーのパスワードで失敗）
#   Text     : 空・RTF・HTML・UTF-8・UTF-16（BOM）など、NULを含まないテキストらしい内容。今と同じ（Officeで開く）
#   Unknown  : 上のどれでもない（先頭4KBにNULを含む）バイナリ。透過暗号化の製品の暗号文の見込みがある。
#              守りを付けてOfficeで開く予備を使う（extract_office.ps1 がアプリごとに切り替える）

${zipSignature}      = [byte[]](0x50, 0x4B, 0x03, 0x04)
${compoundSignature} = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1)

# 権限保護（IRM・秘密度ラベル）を示すCFBのエントリ名（[MS-OFFCRYPTO] IRMDS）。
#   新形式: DataSpaceInfo配下の DRMEncryptedDataSpace・TransformInfo配下の DRMEncryptedTransform
#   旧形式: 保護された中身の \x09DRMContent（必須）と、\x09DRMDataSpace・\x09DRMTransform
#           （任意で \x09DRMViewerContent・\x09LZXDRMDataSpace・\x09LZXTransform）
${rightsProtectedEntryNames} = @(
    "DRMEncryptedDataSpace", "DRMEncryptedTransform",
    ([char]9 + "DRMContent"), ([char]9 + "DRMDataSpace"), ([char]9 + "DRMTransform"),
    ([char]9 + "DRMViewerContent"), ([char]9 + "LZXDRMDataSpace"), ([char]9 + "LZXTransform")
)
# パスワード保護（ECMA-376 標準・アジャイル暗号化）を示すCFBのエントリ名
${passwordProtectedEntryNames} = @("StrongEncryptionDataSpace", "EncryptionInfo")
# データスペース（新形式のパスワード・権限保護が共通で使うストレージ）の名前
${dataSpacesStorageName} = ([char]6 + "DataSpaces")

function testsByteSignature {
    # $bytes の先頭が $signature と一致するか。
    # 判断層はファイルを開けないため、office_reader.ps1 の isZipFile・isCompoundFile が読むのと同じ印を、
    # 既に読んである先頭バイト列（$bytes）の側で比べる（ファイルを開き直さない）
    param (
        [byte[]]$bytes,
        [byte[]]$signature
    )

    if ($null -eq $bytes -or $bytes.Length -lt $signature.Length) {
        return $false
    }
    for ($i = 0; $i -lt $signature.Length; $i++) {
        if ($bytes[$i] -ne $signature[$i]) {
            return $false
        }
    }
    return $true
}

function testsUnicodeTextBom {
    # UTF-16（Excelのテキスト保存・多くのUnicodeテキストエディタ）の先頭のBOM（FF FE）かどうか
    param (
        [byte[]]$bytes
    )

    return ($null -ne $bytes -and $bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
}

function getOfficeProtectionKind {
    # 先頭バイト列とCFBのエントリ名（CFBでなければ $null）から、暗号化の種類を返す
    param (
        [byte[]]$head,
        [string[]]$cfbNames
    )

    if ($null -eq $head -or $head.Length -eq 0) {
        return "Text"
    }
    if (testsByteSignature $head ${zipSignature}) {
        return "Zip"
    }
    if (testsByteSignature $head ${compoundSignature}) {
        $names = @($cfbNames)
        if (@($names | Where-Object { ${rightsProtectedEntryNames} -contains $_ }).Count -gt 0) {
            return "Rights"
        }
        $hasDataSpaces = ($names -contains ${dataSpacesStorageName})
        $hasPasswordNames = (@($names | Where-Object { ${passwordProtectedEntryNames} -contains $_ }).Count -gt 0)
        if ($hasDataSpaces -and !$hasPasswordNames) {
            # データスペースはあるが、名前を1つも取りこぼした場合。安全側（Officeで開かない）に倒す
            return "Rights"
        }
        if ($hasPasswordNames) {
            return "Password"
        }
        return "Legacy"
    }

    # ZIPでもCFBでもない。先頭4KBにNUL（0x00）を含む（UTF-16のBOMで始まるものを除く）＝形式の分からないバイナリ。
    # 空・RTF（`{\rtf`）・HTML・UTF-8・UTF-16（BOM）は、いずれもNULを含まないテキストのため「Text」になる
    if (testsUnicodeTextBom $head) {
        return "Text"
    }
    if (@($head) -contains [byte]0) {
        return "Unknown"
    }
    return "Text"
}

function getProtectionFailureText {
    # 種類ごとの、取り込み一覧のエラー列に出す短い文言。文言が無い種類（Zip・Legacy・Text）は $null
    param (
        [string]$kind
    )

    switch ($kind) {
        "Password" { return "読み取りパスワードが設定されているため開けません（パスワード付きのファイルは取り込めません）" }
        "Rights"   { return "IRM・秘密度ラベルで暗号化されているため取り込めません。" }
        "Unknown"  { return "暗号化されているか壊れているため取り込めません。" }
        default    { return $null }
    }
}

# Officeが保存した一時ファイル（変換した.docx・.pptx、Excelのテキスト保存）まで、透過暗号化の製品が
# 暗号化した場合の文言。ふつうのファイルでも起こりうるため、すべてのOfficeの出力でこの確認を行う
${officeOutputEncryptedMessage} = "ファイルを暗号化する製品が一時ファイルを暗号化したため取り込めません。"

function testOfficeOutput {
    # Officeが保存した出力の先頭バイト列が、期待する形（Zip・UnicodeText）かどうか
    param (
        [byte[]]$head,
        [string]$expectedKind
    )

    if ($expectedKind -eq "Zip") {
        return (testsByteSignature $head ${zipSignature})
    }
    if ($expectedKind -eq "UnicodeText") {
        return (testsUnicodeTextBom $head)
    }
    return $false
}

# Workbook.FileFormat が、テキスト・HTML・CSVとして取り込まれた印の値（xlCSV=6・xlTextMac=19・xlTextWindows=20・
# xlTextMSDOS=21・xlCSVMac=22・xlCSVWindows=23・xlCSVMSDOS=24・xlUnicodeText=42・xlHtml=44・xlCSVUTF8=62・
# xlCurrentPlatformText=-4158）
${excelTextLikeFileFormats} = @(6, 19, 20, 21, 22, 23, 24, 42, 44, 62, -4158)

function testWorkbookFormat {
    # Workbook.FileFormat が、テキスト・HTML・CSVとして取り込まれた印でないか（「形式の分からないバイナリ」のときだけ確かめる。
    # 中身がHTML・CSV等の「Text」の種類のブックには当てない。当てると本来読めているものまで失敗になる）
    param (
        $fileFormat
    )

    return (${excelTextLikeFileFormats} -notcontains $fileFormat)
}

function getWordOpenFormat {
    # 拡張子に合う WdOpenFormat の値。「形式の分からないバイナリ」をWordで開くときだけ渡し、自動判定に任せない
    # （wdOpenFormatDocument=1・wdOpenFormatTemplate=2・wdOpenFormatXMLDocument=9・
    #   wdOpenFormatXMLDocumentMacroEnabled=10・wdOpenFormatXMLTemplate=11・wdOpenFormatAuto=0）
    param (
        [string]$extension
    )

    switch ($extension.ToLowerInvariant()) {
        ".doc"  { return 1 }
        ".dot"  { return 2 }
        ".docx" { return 9 }
        ".docm" { return 10 }
        ".dotx" { return 11 }
        ".dotm" { return 11 }
        default { return 0 }
    }
}
