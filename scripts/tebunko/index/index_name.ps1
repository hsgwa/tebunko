# インデックス名と TSV のファイル名の決め方（判断層。ファイルに触らない）。

function newIndexName {
    # クロール対象フォルダのインデックス名（work\index 直下のフォルダ名）を作る。
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
        # 長いフォルダ名に (2) を付けるとファイル名の上限を超えてフォルダを作れないため、上限に収まるよう名前を切り詰める
        $suffix = "(${i})"
        $name = $base.Substring(0, [Math]::Min($base.Length, ${maxFileNameLength} - $suffix.Length)) + $suffix
    }
    return $name
}

function assignIndexNames {
    # クロール対象フォルダ（getTargetFolders）にインデックス名を割り当て、@{ Path; Enabled; Name } の配列を返す。
    # インデックス名は設定に持つ（getTargetFolders の Name）。フォルダの置き場所（Path）を書き換えても名前は変わらないため、
    # フォルダを移しても同じインデックスとして扱える（インデックスを作り直さない）。
    # 名前が無い場合（新しく追加したフォルダ）は、前回の取り込み一覧の同じパスの名前を使い、
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
    # work\index からの相対パスを、先頭のインデックス名と残り（クロール対象フォルダからの相対パス）に分ける: @{ Name; Rest }
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
    # ファイル名に使えない文字・制御文字と、_・符号化に使う % を "%XX"（16進数）にする（decodeIndexPlace で元に戻す）。
    #   ・全角に置き換えると `a"b` と `a”b` が同じファイル名になり、後のシートで上書きされるため、元に戻せる形にする
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

    # 符号化した文字が無ければそのまま返す（検索のたびに全TSVの名前を分解するため、置換の呼び出しを省く）
    if ($place.IndexOf("%") -lt 0) {
        return $place
    }
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
        throw "インデックスのファイル名が長すぎるため保存できません（$($name.Length) 文字。上限 ${maxFileNameLength} 文字）: ${name}"
    }
    return $name
}


# 図形・コメントなど、本文（セル・段落）以外の文字の場所は "<元の場所>[<種類>]" とする。
# 元の場所は、Excel ではシート名（例: "売上[図形]"。office_reader.ps1 の readXlsxObjectUnits）。
# Word はページ、PowerPoint はスライドを元の場所にする（例: "ページ003[コメント]" "スライド002[図形]"）。
# Excel のシート名には [ ] を使えず、Word・PowerPoint の場所（ページNNN・スライドNNN 等）にも付かないため、ふつうの場所と重ならない。
# 種類ごとに、検索に含めるかを画面で選べる（search_query.ps1 の newPlaceExclude）。
# 種類を足すときは、ここ・書き出す側（office_reader.ps1）・画面（types.ps1 の HitRow.ObjectPlaceRegex）をそろえる
${placeKindShape}   = "図形"      # 図形・テキストボックス・WordArt・SmartArt・グラフ（PowerPoint のテキストボックス・図形はスライドの本文）
${placeKindComment} = "コメント"  # コメント（メモ・スレッド形式のコメント）
${objectPlacePattern} = "^(?<base>.*)\[(?<kind>${placeKindShape}|${placeKindComment})\]$"


function splitObjectPlace {
    # 場所を @{ Base（元の場所）; Kind（種類。ふつうの場所は空） } に分ける。
    #   例: "売上[図形]" → @{ Base = "売上"; Kind = "図形" } / "ページ001" → @{ Base = "ページ001"; Kind = "" }
    param (
        [string]$place
    )

    if ($place -match ${objectPlacePattern}) {
        return @{ Base = $Matches.base; Kind = $Matches.kind }
    }
    return @{ Base = $place; Kind = "" }
}


function describePlace {
    # 画面の「場所」「種別」と検索結果ファイルに出す文字を @{ Place; Kind } で返す（TSV の名前は変えず、表示だけを変える）。
    #   Excel      : "売上" → [シート] 売上・セル / "売上[図形]" → [シート] 売上・図形 / "売上[コメント]" → [シート] 売上・コメント
    #   Word       : "ページ003" → [ページ] 3（目安）・本文（ページは保存時の区切りから数えた目安のため）/ "脚注" → [脚注]・本文
    #   PowerPoint : "スライド002（非表示）" → [スライド] 2（非表示）・本文 / "スライド002_ノート" → [スライド] 2・ノート
    # 種別は、図形・コメントなら場所の種類の名前そのまま（検索条件の［図形も検索］［コメントも検索］と同じ言葉）
    param (
        [string]$book,
        [string]$place
    )

    $split = splitObjectPlace $place
    $base = $split.Base
    $kind = $split.Kind

    if ($book -match '\.xls[a-z]?$') {
        return @{ Place = "[シート] $base"; Kind = $(if ($kind) { $kind } else { "セル" }) }
    }
    if ($base -eq "") {
        return @{ Place = ""; Kind = $(if ($kind) { $kind } else { "本文" }) }
    }
    if ($base -match '^ページ(\d+)$') {
        return @{ Place = "[ページ] $([int]$Matches[1])（目安）"; Kind = $(if ($kind) { $kind } else { "本文" }) }
    }
    if ($base -match '^スライド(\d+)_ノート$') {
        return @{ Place = "[スライド] $([int]$Matches[1])"; Kind = $(if ($kind) { $kind } else { "ノート" }) }
    }
    if ($base -match '^スライド(\d+)(（非表示）)?$') {
        return @{ Place = "[スライド] $([int]$Matches[1])$($Matches[2])"; Kind = $(if ($kind) { $kind } else { "本文" }) }
    }
    # ヘッダー・フッター・脚注など、番号の無い場所
    return @{ Place = "[$base]"; Kind = $(if ($kind) { $kind } else { "本文" }) }
}


# インデックスの「元のファイル名のフォルダ」と分かる名前（Officeファイルの拡張子で終わる）
${indexBookDirPattern} = "\.(?:xls|doc|ppt)[a-z]?$"
