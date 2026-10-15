# インデックス名と TSV のファイル名の決め方（判断層。ファイルに触らない）。

function newIndexName {
    # クロール対象フォルダのインデックス名（work\index 直下のフォルダ名）を作る。
    # フォルダ名（ドライブ直下はドライブ名）を使い、usedNames と重複すれば「名前(2)」「名前(3)」…とする
    param (
        [string]$folderPath,
        $usedNames = $null  # HashSet[string]・配列・1 個の文字列・$null のいずれでもよい
    )

    # 作る名前は、インデックス名の決まり（testIndexName）に必ず合わせる: 前後の空白（全角空白を含む）と最後の . を削り、
    # 長すぎれば切り、削った結果が空・予約名（CON など）になれば「フォルダ」にする
    $base = (toSafeFileName (getFolderLeafName $folderPath)).Trim()
    if ($base.Length -gt ${maxFileNameLength}) {
        $base = $base.Substring(0, ${maxFileNameLength})
    }
    do {
        $before = $base
        $base = $base.Trim().TrimEnd(".")
    } while ($base.Length -ne $before.Length)
    if ((getFolderNameProblemText $base) -ne "") {
        $base = "フォルダ"
    }

    # 呼び出し側から $null や文字列の配列で渡されても落ちないよう、ここで集合に直す（大文字・小文字は区別しない）
    $used = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    # 集合が 1 要素の配列に入って渡されても（@(関数) で受けた場合など）、中身を 1 つずつ取り出す
    foreach ($entry in @($usedNames)) {
        foreach ($usedName in @($entry)) {
            if ($usedName) {
                [void]$used.Add([string]$usedName)
            }
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
    # 設定は書き換えられうる。パスにつなぐ名前として使えない（.. ・末尾の . など）ものは名前が無いものとして割り当て直す。
    # 既にある名前は、予約名（aux.backup など）でも、そのまま使い続ける（testStoredIndexName）
    foreach ($folder in @($targetFolders | Where-Object { $_ -and $_.Name -and !(testStoredIndexName ([string]$_.Name)) })) {
        [void]$used.Add($folder.Name)
    }

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($folder in @($targetFolders | Where-Object { $_ })) {
        $name = [string]$folder.Name
        if ($name -ne "" -and (testStoredIndexName $name)) {
            $name = ""
        }
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


# Windows で使えないファイル名の一覧（${reservedFileNames}）は shared/core/text.ps1 にある


function testIndexName {
    # インデックス名（work\index 直下のフォルダ名）として使えるかを調べ、使えない理由を返す（使えれば空文字列）。
    #   usedNames: ほかのインデックスが使っている名前（大文字・小文字を区別しない）
    param (
        [string]$name,
        [object[]]$usedNames = @()
    )

    $problem = getFolderNameProblemText ([string]$name)
    if ($problem -ne "") {
        return $problem
    }
    $name = [string]$name
    foreach ($used in @($usedNames)) {
        if ([string]::Equals([string]$used, $name, [System.StringComparison]::OrdinalIgnoreCase)) {
            return "「${name}」は、ほかのインデックスが使っています。別の名前を付けてください。"
        }
    }
    return ""
}

function testStoredIndexName {
    # 保存してあるインデックス名（設定・取り込み一覧・状態のフォルダから読んだもの）が、パスにつなぐ名前として使えるかを調べ、
    # 使えない理由を返す（使えれば空文字列）。testIndexName との違いは、Windows の予約名（aux.backup など）を拒まないこと。
    # 前の版で付いた、既にある実在のフォルダの名前を空にせず使い続けるため。新しく付ける名前（newIndexName・画面の入力・取り込み）は
    # 厳しいまま testIndexName で調べる。置き場の外へ出る名前・最後の . や前後の空白・使えない文字は、ここでも拒む。
    # 消す・動かす所は、この確かめに頼らず、名前で組んだパスが置き場の下にあること（testPathUnder）を別に確かめる
    param (
        [string]$name
    )

    return (getFolderNameProblemText ([string]$name) $true)
}

# ファイル名に使えない文字の配列 ${invalidFileNameChars} は shared/core/text.ps1 にある（getFolderNameProblemText と共通）

function testPathSegments {
    # 区切り文字で分けたパスの各区切り（segments）が、ファイル名として安全に使えるかを調べ、
    # 使えない理由を返す（使えれば空文字列）。zip のエントリー名（/ 区切り）・取り込み一覧の相対パス（\ 区切り）の
    # どちらからも呼べる、区切り文字に依らない共通の確かめ。
    #   lenient: $true なら、最後の . や空白・Windows の予約語は拒まない（\\?\ で作られた実在のファイル名を、
    #            取り込み一覧から捨てないため。外へ出る区切り（空・.・..）・使えない文字・長すぎる区切りは拒む）
    param (
        [string[]]$segments,
        [bool]$lenient = $false
    )

    foreach ($segment in $segments) {
        if ($segment -eq "" -or $segment -eq "." -or $segment -eq "..") {
            return "使えない区切り（空・.・..）が含まれています。"
        }
        if ($segment.Length -gt ${maxFileNameLength}) {
            return "1 つの区切りが長すぎます（${maxFileNameLength} 文字まで）。"
        }
        if ($segment.IndexOfAny(${invalidFileNameChars}) -ge 0) {
            return "使えない文字が含まれています。"
        }
        if ($lenient) {
            continue
        }
        if ($segment.EndsWith(".") -or $segment.EndsWith(" ")) {
            return "区切りの最後に . や空白は使えません。"
        }
        if (${reservedFileNames} -contains $segment.Split(".")[0].ToUpperInvariant()) {
            return "Windows で使えない名前（$segment）が含まれています。"
        }
    }
    return ""
}

function testIndexRelPath {
    # 取り込み一覧・状態ファイルから読んだ相対パス（"インデックス名\フォルダからの相対パス"）が、
    # work\content_index の下のパスとして使えるかを調べ、使えない理由を返す（使えれば空文字列）。
    # 先頭は保存してあるインデックス名（testStoredIndexName。予約名は通す）、残りは区切りごとに空・.・..・ファイル名に使えない文字（\ / : など）を拒む
    # （残りには予約名・最後の . や空白の決まりを当てない。実在のファイル名を捨てないため）。
    #   allowNameOnly: 残りが無く、名前だけのもの（"営業"）も通す（システムインデックスの状態のフォルダ）
    #   checkedNames : testStoredIndexName を通った名前の集合（HashSet[string]）。数万行を読むとき、同じ名前を何度も調べないために
    #                  呼び出し側が渡す（通った名前をここに足す）
    param (
        [string]$relPath,
        [bool]$allowNameOnly = $false,
        $checkedNames = $null
    )

    $relPath = [string]$relPath
    $i = $relPath.IndexOf("\")
    $name = if ($i -lt 0) { $relPath } else { $relPath.Substring(0, $i) }
    if ($null -eq $checkedNames -or -not $checkedNames.Contains($name)) {
        $nameReason = testStoredIndexName $name
        if ($nameReason) {
            return $nameReason
        }
        if ($null -ne $checkedNames) {
            [void]$checkedNames.Add($name)
        }
    }
    if ($i -lt 0) {
        if ($allowNameOnly) {
            return ""
        }
        return "インデックス名の後に、フォルダからの相対パスがありません。"
    }
    return (testPathSegments ($relPath.Substring($i + 1).Split("\")) $true)
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

# 決まった場所（ページ・スライド・ヘッダー/フッター・脚注・文書・本文）は、ファイル名を ASCII の固定名にする。
# 符号化した場所（encodeIndexPlace）には "_" が出ないため（"_" は "%5F" にする）、固定名に "_" を使っても
# シート名由来のファイル名と重ならない（convertIndexFileNameToPlace で戻すときの手掛かりにする）。
function convertPlaceBaseToFixedFileName {
    # 場所の元の部分（種類を除いた部分）を、決まったASCIIのファイル名にする（当てはまらなければ空文字列）。
    #   数字はそのまま桁を写す（int にしない。"001" は "001" のまま）
    param (
        [string]$base
    )

    if ($base -match '^ページ(\d+)$') { return "page_$($Matches[1])" }
    if ($base -match '^スライド(\d+)（非表示）$') { return "slide_$($Matches[1])_hidden" }
    if ($base -match '^スライド(\d+)_ノート$') { return "slide_$($Matches[1])_notes" }
    if ($base -match '^スライド(\d+)$') { return "slide_$($Matches[1])" }
    if ($base -eq "ヘッダー・フッター") { return "header_footer" }
    if ($base -eq "脚注") { return "doc_footnotes" }
    if ($base -eq "文書") { return "doc_whole" }
    if ($base -eq "本文") { return "doc_body" }
    return ""
}

function convertFixedFileNameToPlaceBase {
    # convertPlaceBaseToFixedFileName の逆変換（当てはまらなければ空文字列）
    param (
        [string]$fileBase
    )

    if ($fileBase -match '^page_(\d+)$') { return "ページ$($Matches[1])" }
    if ($fileBase -match '^slide_(\d+)_hidden$') { return "スライド$($Matches[1])（非表示）" }
    if ($fileBase -match '^slide_(\d+)_notes$') { return "スライド$($Matches[1])_ノート" }
    if ($fileBase -match '^slide_(\d+)$') { return "スライド$($Matches[1])" }
    if ($fileBase -eq "header_footer") { return "ヘッダー・フッター" }
    if ($fileBase -eq "doc_footnotes") { return "脚注" }
    if ($fileBase -eq "doc_whole") { return "文書" }
    if ($fileBase -eq "doc_body") { return "本文" }
    return ""
}

function toIndexFileName {
    # インデックスのTSVのファイル名 "<場所>.tsv" を返す。
    #   ・決まった場所（ページ・スライド・ヘッダー/フッター・脚注・文書・本文）は ASCII の固定名にする
    #   ・それ以外（Excelのシート名）は encodeIndexPlace で符号化する
    #   ・図形・コメント・ヘッダー・フッター・埋め込みは、末尾に英語の種類 "[shape]" "[comment]" "[header_footer]" "[embed<N>]" を付ける（placeKindFileNames）
    # 元のファイル名はフォルダ名（= 元のファイル名そのもの）にするため、ファイル名には入れない。
    # ファイル名の上限（255文字）は長いパスの対応（toLongPath）でも超えられないため、超える場合は分かるメッセージで例外にする
    param (
        [string]$place
    )

    $split = splitObjectPlace $place
    $fixedName = convertPlaceBaseToFixedFileName $split.Base
    $base = if ($fixedName) { $fixedName } else { encodeIndexPlace $split.Base }
    $suffix = if ($split.Kind) { "[$(${placeKindFileNames}[$split.Kind])$(if ($split.Number -gt 0) { $split.Number })]" } else { "" }
    $name = "{0}{1}.tsv" -f $base, $suffix
    if ($name.Length -gt ${maxFileNameLength}) {
        throw "インデックスのファイル名が長すぎるため保存できません（$($name.Length) 文字。上限 ${maxFileNameLength} 文字）: ${name}"
    }
    return $name
}

function convertIndexFileNameToPlace {
    # toIndexFileName の逆変換。ファイル名（拡張子 .tsv を除いたもの）から場所を復元する。
    #   戻す順番: (1) 末尾の "[shape]"・"[comment]"・"[header_footer]"・"[embed<N>]"（placeKindFileNames の値）を外す →
    #             (2) 残りに "_" があれば固定の場所の表で戻し、無ければ decodeIndexPlace で戻す →
    #             (3) (1) で外した種類を日本語（[図形]・[コメント]・[ヘッダー・フッター]・[埋め込み<N>]）で付け直す
    param (
        [string]$fileNameWithoutExtension
    )

    $kind = ""
    $body = $fileNameWithoutExtension
    $fileKindAlternation = (${placeKindFileNames}.Values | ForEach-Object { [regex]::Escape($_) }) -join "|"
    if ($fileNameWithoutExtension -match "^(?<base>.+)\[(?<kind>$fileKindAlternation)(?<number>${objectPlaceNumberPattern})?\]`$") {
        $kindName = $Matches.kind
        $kind = (${placeKindFileNames}.GetEnumerator() | Where-Object { $_.Value -eq $kindName } | Select-Object -First 1).Key
        # 番号は埋め込みにだけ付く（ほかの種類に付いた名前は、この関数が作らない形のため戻さない）
        if (($kind -eq ${placeKindEmbed}) -eq [bool]$Matches.number) {
            $body = $Matches.base
            if ($Matches.number) { $kind = "${kind}$($Matches.number)" }
        } else {
            $kind = ""
        }
    }

    if ($body.IndexOf("_") -ge 0) {
        $base = convertFixedFileNameToPlaceBase $body
    } else {
        $base = decodeIndexPlace $body
    }

    if ($kind) { return "$base[$kind]" }
    return $base
}


# 図形・コメントなど、本文（セル・段落）以外の文字の場所は "<元の場所>[<種類>]" とする。
# 元の場所は、Excel ではシート名（例: "売上[図形]"。office_reader.ps1 の readXlsxObjectUnits）。
# Word はページ、PowerPoint はスライドを元の場所にする（例: "ページ003[コメント]" "スライド002[図形]"）。
# Excel のシート名には [ ] を使えず、Word・PowerPoint の場所（ページNNN・スライドNNN 等）にも付かないため、ふつうの場所と重ならない。
# 種類ごとに、検索に含めるかを画面で選べる（search_query.ps1 の newPlaceExclude）。
# 種類を足すときは、ここ（placeKindFileNames にも足す）と書き出す側（office_reader.ps1）をそろえる。検索の除外・画面の判定の正規表現は newObjectPlaceKindPattern が表から組み立てる
${placeKindShape}   = "図形"      # 図形・テキストボックス・WordArt・SmartArt・グラフ（PowerPoint のテキストボックス・図形はスライドの本文）
${placeKindComment} = "コメント"  # コメント（メモ・スレッド形式のコメント）
${placeKindHeaderFooter} = "ヘッダー・フッター"  # Excel のヘッダー・フッター（Word・PowerPoint の "ヘッダー・フッター" は場所の名前で、種類ではない）。いつも検索する（除外の選択肢は無い）
${placeKindEmbed}   = "埋め込み"  # Word・PowerPoint に埋め込んだ Office のファイル（.xlsx・.docx・.pptx など）の中の文字。場所は "<元の場所>[埋め込み<N>]"（N は 1 ファイルの中の通し番号。ほかの種類と違い、番号が付く）。図形と一緒に検索の有無を切り替える

# 種類ごとの、TSVのファイル名に付ける英語の名前（toIndexFileName・convertIndexFileNameToPlace で使う）。
# 番号の付く種類（埋め込み）のファイル名は "[embed<N>]"
${placeKindFileNames} = @{ ${placeKindShape} = "shape"; ${placeKindComment} = "comment"; ${placeKindHeaderFooter} = "header_footer"; ${placeKindEmbed} = "embed" }

# ${objectPlacePattern} の種類の選択肢は ${placeKindFileNames} のキーから組み立てる（足し忘れを防ぐ）。
# 番号の付く種類（埋め込み）は選択肢に入れず、番号（先頭が 0 でない数字）を付けた形だけを別に当てる
# （"[埋め込み]"・"[埋め込み01]" はふつうの場所とする）
${objectPlaceKindAlternation} = (${placeKindFileNames}.Keys | Where-Object { $_ -ne ${placeKindEmbed} } | ForEach-Object { [regex]::Escape($_) }) -join "|"
${objectPlaceNumberPattern} = "[1-9][0-9]{0,8}"
${objectPlacePattern} = "^(?<base>.*)\[(?:(?<kind>${objectPlaceKindAlternation})|(?<numberedKind>$([regex]::Escape(${placeKindEmbed})))(?<number>${objectPlaceNumberPattern}))\]$"


function newObjectPlaceKindPattern {
    # 場所の末尾の種類だけを当てる正規表現の文字列を、${placeKindFileNames} の表から組み立てる。
    # 検索から外す場所（search_query.ps1 の newPlaceExclude）と画面の判定（ui/search/result_types.ps1 の HitRow.ObjectPlaceRegex）が使う。
    #   例（全部）: \[(?:図形|コメント|ヘッダー・フッター|埋め込み[1-9][0-9]{0,8})\]$（選択肢の並びは表の順による）
    # kinds: 当てる種類の名前の並び（省略または空は全部）。番号の付く種類（埋め込み）は番号の形を付けて当てる
    param (
        [string[]]$kinds = @()
    )

    if ($kinds.Count -eq 0) {
        $kinds = @(${placeKindFileNames}.Keys)
    }
    $alternatives = foreach ($kind in $kinds) {
        if ($kind -eq ${placeKindEmbed}) { "$([regex]::Escape($kind))${objectPlaceNumberPattern}" } else { [regex]::Escape($kind) }
    }
    return "\[(?:$(@($alternatives) -join '|'))\]`$"
}


function splitObjectPlace {
    # 場所を @{ Base（元の場所）; Kind（種類。ふつうの場所は空）; Number（埋め込みの番号。ほかは 0） } に分ける。
    #   例: "売上[図形]" → @{ Base = "売上"; Kind = "図形"; Number = 0 } / "ページ001" → @{ Base = "ページ001"; Kind = ""; Number = 0 }
    #       "ページ003[埋め込み2]" → @{ Base = "ページ003"; Kind = "埋め込み"; Number = 2 }
    param (
        [string]$place
    )

    if ($place -match ${objectPlacePattern}) {
        if ($Matches.numberedKind) {
            return @{ Base = $Matches.base; Kind = $Matches.numberedKind; Number = [int]$Matches.number }
        }
        return @{ Base = $Matches.base; Kind = $Matches.kind; Number = 0 }
    }
    return @{ Base = $place; Kind = ""; Number = 0 }
}


function describePlace {
    # 場所ごとの表記と「種別」を @{ Place; Kind } で返す（TSV の名前は変えず、表示だけを変える）。
    # 画面の見出しの要約・検索結果ファイル・コピーに使う。表の「場所」の列は、これにセル番地を足した describeHitPlace を使う
    #   Excel      : "売上" → [シート]売上・セル / "売上[図形]" → [シート]売上・図形 / "売上[コメント]" → [シート]売上・コメント
    #                / "売上[ヘッダー・フッター]" → [シート]売上・ヘッダー・フッター
    #   Word       : "ページ003" → 3 ページ（目安）・本文（ページは保存時の区切りから数えた目安のため）/ "脚注" → 脚注・本文
    #   PowerPoint : "スライド002（非表示）" → スライド 2（非表示）・本文 / "スライド002_ノート" → スライド 2・ノート
    #   テキスト   : "本文" → 空・本文（場所は 1 つだけのため、表の「場所」の列は describeHitPlace の "N 行目" だけで足りる）
    # 種別は、図形・コメントなら場所の種類の名前そのまま（検索条件の［図形も検索］［コメントも検索］と同じ言葉）。
    # 埋め込みは番号を付けた "埋め込み 1"（場所は元の場所のページ・スライドの表記のまま）
    param (
        [string]$book,
        [string]$place
    )

    if (testTextExtension $book) {
        return @{ Place = ""; Kind = "本文" }
    }

    $split = splitObjectPlace $place
    $base = $split.Base
    $kind = $split.Kind
    if ($split.Number -gt 0) { $kind = "$kind $($split.Number)" }

    if ((getContentIndexFileKind $book) -eq "Excel") {
        return @{ Place = "[シート]$base"; Kind = $(if ($kind) { $kind } else { "セル" }) }
    }
    if ($base -eq "") {
        return @{ Place = ""; Kind = $(if ($kind) { $kind } else { "本文" }) }
    }
    if ($base -match '^ページ(\d+)$') {
        if (testPdfExtension $book) {
            return @{ Place = "$([int]$Matches[1]) ページ"; Kind = $(if ($kind) { $kind } else { "本文" }) }
        }
        return @{ Place = "$([int]$Matches[1]) ページ（目安）"; Kind = $(if ($kind) { $kind } else { "本文" }) }
    }
    if ($base -match '^スライド(\d+)_ノート$') {
        return @{ Place = "スライド $([int]$Matches[1])"; Kind = $(if ($kind) { $kind } else { "ノート" }) }
    }
    if ($base -match '^スライド(\d+)(（非表示）)?$') {
        return @{ Place = "スライド $([int]$Matches[1])$($Matches[2])"; Kind = $(if ($kind) { $kind } else { "本文" }) }
    }
    # ヘッダー・フッター・脚注など、番号の無い場所
    return @{ Place = $base; Kind = $(if ($kind) { $kind } else { "本文" }) }
}


function describeHitPlace {
    # 結果の表の「場所」の列とプレビューの題に出す、行ごとの場所の表記を返す（場所ごとの表記にセル番地を足したもの）。
    # Excel だけ足す（Word・PowerPoint は場所ごとと同じ）。テキストは行番号だけ（"N 行目"）を返す（場所は 1 つしかないため）。
    #   セル             : [シート]売上!B12 / 1 行に複数のセルが一致 [シート]売上!B12 ほか 2 / セル番地が求まらないとき [シート]売上 12 行目
    #   図形・コメント   : [シート]売上!D5 / セル番地が求まらないとき [シート]売上（行番号は通し番号のため出さない）
    #   ヘッダー・フッター: [シート]売上（行にセル番地が無く、matchCell が空のため）
    #   テキスト         : 12 行目
    param (
        [string]$place,
        [bool]$isExcel,
        [bool]$isObjectPlace,
        [string]$matchCell,
        [int]$matchCount,
        [int]$lineNumber,
        [bool]$isText = $false
    )

    if ($isText) { return "$lineNumber 行目" }
    if (-not $isExcel) { return $place }
    if ($isObjectPlace) {
        if ($matchCell) { return "$place!$matchCell" }
        return $place
    }
    if (-not $matchCell) { return "$place $lineNumber 行目" }
    if ($matchCount -gt 1) { return "$place!$matchCell ほか $($matchCount - 1)" }
    return "$place!$matchCell"
}


# インデックスの「元のファイル名のフォルダ」と分かる名前（Office・テキスト・PDF の拡張子で終わる）
# （testIndexBookName もこの正規表現で判定する。Office の拡張子は officeExtensions の表から作る）
${indexBookDirPattern} = "(?:$(((${officeExtensions} + ${textExtensions} + ${pdfExtensions}) | ForEach-Object { [regex]::Escape($_) }) -join '|'))$"
