# インデックス名と TSV のファイル名の決め方（判断層。ファイルに触らない）。

function newIndexName {
    # 変換対象フォルダのインデックス名（work\index 直下のフォルダ名）を作る。
    # フォルダ名（ドライブ直下はドライブ名）を使い、usedNames と重複すれば「名前(2)」「名前(3)」…とする
    param (
        [string]$folderPath,
        $usedNames = $null  # HashSet[string]・配列・1 個の文字列・$null のいずれでもよい
    )

    $base = toSafeFileName (getFolderLeafName $folderPath)
    if ($base -eq "") {
        $base = "フォルダ"
    }

    # 呼び出し側から $null や文字列の配列で渡されても落ちないよう、ここで集合に直す（大文字・小文字は区別しない）
    $used = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($usedName in @($usedNames)) {
        if ($usedName) {
            [void]$used.Add([string]$usedName)
        }
    }

    $name = $base
    for ($i = 2; $used.Contains($name); $i++) {
        $name = "${base}(${i})"
    }
    return $name
}

function assignIndexNames {
    # 変換対象フォルダ（getTargetFolders）にインデックス名を割り当て、@{ Path; Enabled; Name } の配列を返す。
    # インデックス名は設定に持つ（getTargetFolders の Name）。フォルダの置き場所（Path）を書き換えても名前は変わらないため、
    # フォルダを移しても同じインデックスとして扱える（インデックスを作り直さない）。
    # 名前が無い場合（新しく追加したフォルダ・以前の版の設定）は、前回の変換一覧の同じパスの名前を使い、
    # それも無ければフォルダ名から重複しない名前を作る
    param (
        [object[]]$targetFolders,
        [object[]]$previousFolders  # readStatusFile の Folders（@{ Path; Name }）
    )

    $previousNames = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($folder in @($previousFolders | Where-Object { $_ -and $_.Name })) {
        $previousNames[$folder.Path] = $folder.Name
    }

    # 設定にある名前・前回の名前は使わない（別のフォルダに同じ名前を付けてインデックスを取り違えないため）
    $used = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($name in $previousNames.Values) {
        [void]$used.Add($name)
    }
    foreach ($folder in @($targetFolders | Where-Object { $_ -and $_.Name })) {
        [void]$used.Add($folder.Name)
    }

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($folder in @($targetFolders | Where-Object { $_ })) {
        $name = [string]$folder.Name
        if ($name -eq "") {
            if ($previousNames.ContainsKey($folder.Path)) {
                $name = $previousNames[$folder.Path]
            } else {
                $name = newIndexName $folder.Path $used
            }
            [void]$used.Add($name)
        }
        $result.Add([pscustomobject]@{ Path = $folder.Path; Enabled = $folder.Enabled; Name = $name })
    }
    return $result.ToArray()
}

function splitIndexRelPath {
    # work\index からの相対パスを、先頭のインデックス名と残り（変換対象フォルダからの相対パス）に分ける: @{ Name; Rest }
    # （String.Split([char], 2) は .NET Framework では Split(params char[]) になり、2 も区切り文字とみなされるため使わない）
    param (
        [string]$relPath
    )

    $i = $relPath.IndexOf("\")
    if ($i -lt 0) {
        return @{ Name = $relPath; Rest = "" }
    }
    return @{ Name = $relPath.Substring(0, $i); Rest = $relPath.Substring($i + 1) }
}


# Windows で使えないファイル名（インデックス名に使えるかの判定に使う）
${reservedFileNames} = @("CON", "PRN", "AUX", "NUL") +
    @(1..9 | ForEach-Object { "COM$_" }) + @(1..9 | ForEach-Object { "LPT$_" })


function testIndexName {
    # インデックス名（work\index 直下のフォルダ名）として使えるかを調べ、使えない理由を返す（使えれば空文字列）。
    #   usedNames: ほかのインデックスが使っている名前（大文字・小文字を区別しない）
    param (
        [string]$name,
        [object[]]$usedNames = @()
    )

    $name = [string]$name
    if ($name -eq "") {
        return "インデックス名を入力してください。"
    }
    if ($name -ne $name.Trim()) {
        return "インデックス名の前後に空白は使えません。"
    }
    if ($name.Length -gt ${maxFileNameLength}) {
        return "インデックス名が長すぎます（${maxFileNameLength} 文字まで）。"
    }
    if (@([System.IO.Path]::GetInvalidFileNameChars() | Where-Object { $name.IndexOf($_) -ge 0 }).Count -gt 0) {
        return "インデックス名に使えない文字が含まれています（\ / : * ? " + [char]34 + " < > | と制御文字）。"
    }
    if ($name.EndsWith(".")) {
        return "インデックス名の最後に . は使えません。"
    }
    if (${reservedFileNames} -contains $name.Split(".")[0].ToUpperInvariant()) {
        return "「${name}」は Windows で使えない名前です。"
    }
    foreach ($used in @($usedNames)) {
        if ([string]::Equals([string]$used, $name, [System.StringComparison]::OrdinalIgnoreCase)) {
            return "「${name}」は、ほかのインデックスが使っています。別の名前を付けてください。"
        }
    }
    return ""
}

function encodeIndexPlace {
    # インデックスのTSVのファイル名に入れる場所（シート名・ページ・スライド）を符号化する。
    # ファイル名に使えない文字・制御文字と、区切りの _・符号化に使う % を "%XX"（16進数）にする（decodeIndexPlace で元に戻す）。
    #   ・全角に置き換えると `a"b` と `a”b` が同じファイル名になり、後のシートで上書きされるため、元に戻せる形にする
    #   ・場所に _ が残らないため、ファイル名に .xlsx_ 等を含むブックでも、最後の _ でファイル名と場所に分けられる
    param (
        [string]$place
    )

    return [regex]::Replace($place, '[\x00-\x1F"%*/:<>?\\_|]', { param($m) '%{0:X2}' -f [int][char]$m.Value })
}

function decodeIndexPlace {
    # encodeIndexPlace で符号化した場所を元に戻す（encodeIndexPlace が作る "%XX" だけを戻す）
    param (
        [string]$place
    )

    return [regex]::Replace($place, '%(?:[01][0-9A-F]|2[25AF]|3[ACEF]|5[CF]|7C)', { param($m) [string][char][Convert]::ToInt32($m.Value.Substring(1), 16) })
}

function toIndexFileName {
    # インデックスのTSVのファイル名 "<場所>.tsv" を返す（場所は encodeIndexPlace で符号化する）。
    # 元のファイル名はフォルダ名（= 元のファイル名そのもの）にするため、ファイル名には入れない。
    # ファイル名の上限（255文字）は長いパスの対応（toLongPath）でも超えられないため、超える場合は分かるメッセージで例外にする
    param (
        [string]$place
    )

    $name = "{0}.tsv" -f (encodeIndexPlace $place)
    if ($name.Length -gt ${maxFileNameLength}) {
        throw "変換結果のファイル名が長すぎるため保存できません（$($name.Length) 文字。上限 ${maxFileNameLength} 文字）: ${name}"
    }
    return $name
}


# インデックスのTSVのファイル名を、元のファイル名（book）と場所（sheet。encodeIndexPlace で符号化したもの）に分ける正規表現
# （大文字・小文字を区別しない。検索処理の C# でも使う）。
#   ・場所は _ を符号化してある（toIndexFileName）ため、最後の _ で分ける（ファイル名に .xlsx_ 等を含むブックでも正しく分かれる）
#   ・以前の版のTSV（場所を全角に置き換え、_ はそのまま）で最後の _ の前が拡張子にならないものは、最初の「拡張子_」で分ける
${indexFileNamePattern} = "^(?:(?<book>.*\.(?:xls|doc|ppt)[a-z]?)_(?<sheet>[^_]*)|(?<book>.*?\.(?:xls|doc|ppt)[a-z]?)_(?<sheet>.*))\.tsv$"


function splitIndexFileName {
    # 以前の形式（フラット）の "ファイル名.拡張子_場所.tsv" を、ファイル名と場所に分解する（場所は decodeIndexPlace で元に戻す）。
    # 今の形式は「ファイル名のフォルダ＋場所.tsv」のため splitIndexTsvPath を使う
    #   例: "ブック名.xlsx_シート名.tsv" / "文書.docx_ページ001.tsv" / "資料.pptx_スライド003%5Fノート.tsv"
    param (
        [string]$fileName
    )

    if ($fileName -match ${indexFileNamePattern}) {
        return @{ book = $Matches.book; sheet = (decodeIndexPlace $Matches.sheet) }
    }

    return @{ book = $fileName; sheet = "" }
}


# インデックスの「元のファイル名のフォルダ」と分かる名前（Officeファイルの拡張子で終わる）
${indexBookDirPattern} = "\.(?:xls|doc|ppt)[a-z]?$"


function splitIndexTsvPath {
    # インデックスフォルダからのTSVの相対パスを @{ Book（元のファイル名）; Place（場所）; RelDir（元のファイルのあるフォルダ） } に分解する。
    # 今の形式（<相対フォルダ>\<ファイル名.xlsx>\<場所>.tsv）と、以前の形式（<相対フォルダ>\<ファイル名.xlsx>_<場所>.tsv）の両方を扱う。
    # 場所には _ を符号化して入れる（encodeIndexPlace）ため、ファイル名に _ があれば以前の形式と分かる
    param (
        [string]$relPath
    )

    $fileName = [System.IO.Path]::GetFileName($relPath)
    $dir = [System.IO.Path]::GetDirectoryName($relPath)
    if ($fileName.IndexOf("_") -lt 0 -and $dir -and ([System.IO.Path]::GetFileName($dir) -match ${indexBookDirPattern})) {
        return @{
            Book   = [System.IO.Path]::GetFileName($dir)
            Place  = (decodeIndexPlace ([System.IO.Path]::GetFileNameWithoutExtension($fileName)))
            RelDir = [System.IO.Path]::GetDirectoryName($dir)
        }
    }

    $name = splitIndexFileName $fileName
    return @{ Book = $name.book; Place = $name.sheet; RelDir = $dir }
}
