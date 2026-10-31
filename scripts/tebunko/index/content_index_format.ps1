# 検索用の本文インデックスのファイル（フォルダ 1 つ・元のファイルの拡張子 1 つにつき、大きさで分けて 1 つ以上。content_index.xlsx.001.tsv など）の形式（判断層）。
# ファイル（元のファイル）ごと・場所（シート・ページ・スライドなど）ごとに、メタ情報の行と今の TSV の中身を並べる。
#
#   <RS> 版=2
#   <RS> ファイル名=(株)山田商事_見積書.xlsx
#   <RS> 種類=Excel
#   <RS> シート=見積
#   <RS> 対象=本文
#   品名→数量→単価        ← 今の TSV の中身そのまま（→ はタブ）。この行が 1 行目
#
# ・メタ情報の行は「RS（U+001E）、半角スペース、キー=値」。1 行に 1 項目。
#   ファイル名= で新しいファイル、シート= ページ= スライド= 部分= で新しい場所が始まる。知らないキーは読み飛ばす
# ・1 つで contentIndexSearchMaxBytes を超える元のファイルは、約 contentIndexFileMaxBytes の塊に分けて、塊ごとに専用の本文インデックスのファイルに入れる
#   （1 行目は「版=3」。分けないファイルは「版=2」のまま）。どの塊も、頭の「ファイル名=」「種類=」と場所のメタ情報の行を繰り返す。
#   場所の途中から始まる塊には「開始行=<N>」（その塊の最初の行が、場所の何行目か。無ければ 1）を付ける
# ・中身には U+001C～U+001F を入れない（書くときに取り除く）ため、行の先頭が RS ならメタ情報の行
# ・改行は LF にそろえる。行の分け方は StreamReader.ReadLine と同じ（CRLF・LF・CR）にし、行番号を変えない
# ・文字コードは UTF-16LE（BOM 付き。content_index_store.ps1 が読み書きする）
# ・Excel の図形・コメントの行は「セル番地 + タブ + 文字」の形。検索は区切りのタブ（最初のタブ）より後だけを照合する（CellPrefixed。content_index_search.ps1）

# 本文インデックスのファイルの名前は「content_index.<元のファイルの拡張子（小文字）>.<番号（3 桁以上）>.tsv」（content_index.xlsx.001.tsv など）。
# 1 つの本文インデックスのファイルが contentIndexFileMaxBytes 以上になったら、それ以上ブックを足さず、次の番号の本文インデックスのファイルに足す
# 先頭の語（content_index）は 1 か所の定数にする（index_store.ps1 の getContentIndexKeyOf も使う）
${contentIndexFileNamePrefix} = "content_index"
${contentIndexFilePattern} = "${contentIndexFileNamePrefix}.*.tsv"
${contentIndexFileNamePattern} = "^${contentIndexFileNamePrefix}\.(?<ext>[^.]+)\.(?<part>[0-9]{3,})\.tsv$"
${contentIndexFileMaxBytes} = 4MB
# 読む本文インデックスのファイル 1 つの大きさの上限（バイト。検索も、インデックス作成で前のファイルとまぜるときも使う）。超えるファイルは読まずに飛ばす。
# 1 つのファイルは丸ごと読み、並列の検索は最大 4 スレッドなので、同時に持つのは最大でこの 4 倍（正しく作ったファイルは contentIndexFileMaxBytes が目安）
${contentIndexSearchMaxBytes} = 64MB
# 元のファイル 1 つで上限を超えるときの、取り込み一覧のエラーの列の文言。上限は文字の量（UTF-16 のバイト数）のため、「ファイルサイズ」とは書かない。
# 読み込んだときの上限から作る（テストが上限を小さくしても、文言は変わらない）
${contentIndexTooLargeMessage} = "ファイルの文字が多すぎるため更新できません（上限 $([int](${contentIndexSearchMaxBytes} / 1MB)) MB）。"
# インデックスの格納形式の版（本文インデックスの「版=」の行・目録の版・zip の目録の版はすべてこの 1 つの番号にそろえる。
# 例外は、分けた塊を入れたファイルの「版=3」だけ（下の contentIndexSplitVersion）。）
# 読み飛ばしても古い版が誤らない列・キーを足すときは上げない。意味を変える・無いと古い版が誤るものを足すときは上げる（format.md「目録の版の規則」）
${indexFormatVersion} = 2
${contentIndexVersion} = ${indexFormatVersion}
# 元のファイルを分けた塊を入れた本文インデックスのファイルだけの版（目録・zip の版は上げない。古い版はこのファイルだけ読み飛ばす。format.md「目録の版の規則」）
${contentIndexSplitVersion} = 3
${contentIndexMark} = [char]0x1E

# 古い形式（前の版のインデックス）を利用者に知らせる文言。画面の一覧・インデクサのログ・取り込みの断りで同じ文言を使う
${indexRebuildMessage} = "古い形式です。作り直してください"
${indexRebuildingMessage} = "作り直しの途中です。［更新］で続きを作ってください"
${indexNewerMessage} = "新しい版の tebunko で作られています"

# インデックスの状態（getIndexFormatState）
${indexStateCurrent}    = "Current"
${indexStateBuilding}   = "Building"
${indexStateRebuilding} = "Rebuilding"
${indexStateOld}        = "Old"
${indexStateEmpty}      = "Empty"
${indexStateBroken}     = "Broken"
${indexStateNewer}      = "Newer"

# 目録（store\format.tsv）のインデックスごとの行の状態
${catalogRowBuilding} = "作成中"
${catalogRowFinished} = "作り終えた"

# 場所のまとまりを始めるキー
${contentIndexPlaceKeys} = @("シート", "ページ", "スライド", "部分")


function getContentIndexFileKind {
    # 元のファイル名から種類（Excel・Word・PowerPoint・テキスト・PDF）を返す。分からなければ空
    param (
        [string]$book
    )

    $officeKind = getOfficeFileKind $book
    if ($officeKind) { return $officeKind }
    if (testTextExtension $book) { return "テキスト" }
    if (testPdfExtension $book) { return "PDF" }
    return ""
}

function getContentIndexExtension {
    # 元のファイル名から、本文インデックスのファイルを分ける拡張子（小文字・先頭の . なし）を返す
    param (
        [string]$book
    )

    return [System.IO.Path]::GetExtension($book).TrimStart(".").ToLowerInvariant()
}


function getContentIndexFileName {
    # 拡張子・番号の本文インデックスのファイルの名前（content_index.xlsx.001.tsv など）を返す
    param (
        [string]$extension,
        [int]$part = 1
    )

    return "${contentIndexFileNamePrefix}.{0}.{1:D3}.tsv" -f $extension, $part
}


function getContentIndexBookHeadPrefix {
    # 本文インデックスで、元のファイル 1 つの始まりを表す行の、値の前までの部分（「<RS> ファイル名=」）
    return "$([string]${contentIndexMark}) ファイル名="
}


function readContentIndexFileName {
    # 本文インデックスのファイルの名前から @{ Extension; Part } を返す。本文インデックスのファイルの名前でなければ $null
    # （番号が大きすぎて整数にできない名前も、本文インデックスのファイルの名前とはみなさない）
    param (
        [string]$name
    )

    if ($name -match ${contentIndexFileNamePattern}) {
        $part = 0
        if ([int]::TryParse($Matches.part, [ref]$part)) {
            return @{ Extension = $Matches.ext.ToLowerInvariant(); Part = $part }
        }
    }
    return $null
}


function splitContentIndexBooksByExtension {
    # 元のファイルの並びを、拡張子ごとの並び（[ordered] 拡張子 → 並び。拡張子は現れた順）に分ける。各並びの中の順は変えない
    param (
        $books
    )

    $groups = [ordered]@{}
    foreach ($book in $books) {
        $extension = getContentIndexExtension ([string]$book.Name)
        if (!$groups.Contains($extension)) {
            $groups[$extension] = New-Object System.Collections.Generic.List[object]
        }
        $groups[$extension].Add($book)
    }
    return $groups
}


function encodeContentIndexValue {
    # メタ情報の値を 1 行に収まる形にする。制御文字（タブを除く）と % を "%XX" にする（decodeContentIndexValue で戻す）
    param (
        [string]$value
    )

    if ($value -notmatch '[\x00-\x08\x0A-\x1F%]') {
        return $value
    }
    return [regex]::Replace($value, '[\x00-\x08\x0A-\x1F%]', { param($m) '%{0:X2}' -f [int][char]$m.Value })
}


function decodeContentIndexValue {
    # encodeContentIndexValue で符号化した値を戻す
    param (
        [string]$value
    )

    if ($value.IndexOf("%") -lt 0) {
        return $value
    }
    return [regex]::Replace($value, '%(?:[01][0-9A-F]|25)', { param($m) [string][char][Convert]::ToInt32($m.Value.Substring(1), 16) })
}


function convertPlaceToContentIndexMeta {
    # 今の場所の名前（TSV のファイル名。"見積[図形]"・"ページ001"・"スライド002（非表示）"・"スライド001_ノート" など）を、
    # 場所のメタ情報（[ordered] キー → 値）にする。convertContentIndexMetaToPlace で必ず元の名前に戻る形にし、
    # 分けられない名前は「部分=名前そのまま・対象=本文」にする
    param (
        [string]$book,
        [string]$place
    )

    $split = splitObjectPlace $place
    $target = if ($split.Kind) { $split.Kind } else { "本文" }
    $meta = [ordered]@{}
    switch (getContentIndexFileKind $book) {
        "Excel" {
            $meta["シート"] = $split.Base
            $meta["対象"] = $target
        }
        "Word" {
            if ($split.Base -match '^ページ(\d+)$') {
                $meta["ページ"] = [string][int]$Matches[1]
            } else {
                $meta["部分"] = $split.Base
            }
            $meta["対象"] = $target
        }
        "PowerPoint" {
            if (!$split.Kind -and $split.Base -match '^スライド(\d+)_ノート$') {
                $meta["スライド"] = [string][int]$Matches[1]
                $meta["対象"] = "ノート"
            } elseif ($split.Base -match '^スライド(\d+)(（非表示）)?$') {
                $meta["スライド"] = [string][int]$Matches[1]
                $meta["対象"] = $target
                if ($Matches[2]) { $meta["非表示"] = "はい" }
            } else {
                $meta["部分"] = $split.Base
                $meta["対象"] = $target
            }
        }
        default {
            $meta["部分"] = $split.Base
            $meta["対象"] = $target
        }
    }
    if ($split.Number -gt 0) {
        # 埋め込みの番号（"ページ003[埋め込み2]" → 対象=埋め込み・埋め込み=2。場所を始めるキーではない）
        $meta["埋め込み"] = [string]$split.Number
    }
    if ((convertContentIndexMetaToPlace $meta) -cne $place) {
        # 番号の桁が違う（"ページ1"）など、組み立て直すと同じ名前にならないものは、名前をそのまま持つ
        $meta = [ordered]@{ "部分" = $place; "対象" = "本文" }
    }
    return $meta
}


function convertContentIndexMetaToPlace {
    # 場所のメタ情報から、今の場所の名前（画面の表示・図形とコメントの除外・元のファイルを開く処理が使う形）を組み立てる
    param (
        $meta
    )

    $target = [string]$meta["対象"]
    $suffix = if ($target -eq ${placeKindShape} -or $target -eq ${placeKindComment} -or $target -eq ${placeKindHeaderFooter}) { "[$target]" } else { "" }
    if ($target -eq ${placeKindEmbed} -and $meta.Contains("埋め込み")) {
        $suffix = "[$target$([string]$meta["埋め込み"])]"
    }
    if ($meta.Contains("シート")) {
        return [string]$meta["シート"] + $suffix
    }
    # 番号が数でないときは、その値の文字列のまま場所に使う（壊れた本文インデックスのファイルでも例外にしない）
    $number = 0
    if ($meta.Contains("ページ")) {
        $pageText = [string]$meta["ページ"]
        $pageName = if ([int]::TryParse($pageText, [ref]$number)) { "ページ{0:D3}" -f $number } else { $pageText }
        return $pageName + $suffix
    }
    if ($meta.Contains("スライド")) {
        $slideText = [string]$meta["スライド"]
        $base = if ([int]::TryParse($slideText, [ref]$number)) { "スライド{0:D3}" -f $number } else { $slideText }
        if ($target -eq "ノート") {
            return "${base}_ノート"
        }
        if ($meta["非表示"] -eq "はい") {
            $base += "（非表示）"
        }
        return $base + $suffix
    }
    if ($target -eq "本文") {
        return [string]$meta["部分"]
    }
    return [string]$meta["部分"] + $suffix
}


function convertToContentIndexBody {
    # TSV の中身を、本文インデックスのファイルに入れる形（LF 区切り・末尾に LF・U+001C～U+001F を除く）にする。
    # 行の分け方は StreamReader.ReadLine と同じ（CRLF・LF・CR で分け、末尾の改行の後ろは行にしない）。中身が空なら空を返す
    param (
        [string]$text
    )

    if ($text.Length -eq 0) {
        return ""
    }
    $body = $text.Replace("`r`n", "`n").Replace("`r", "`n")
    if ($body.IndexOfAny([char[]]@([char]0x1C, [char]0x1D, [char]0x1E, [char]0x1F)) -ge 0) {
        $body = [regex]::Replace($body, '[\x1C-\x1F]', "")
    }
    if ($body.EndsWith("`n")) {
        return $body
    }
    return $body + "`n"
}


function convertBookToContentIndexBlock {
    # 元のファイル 1 つの中身から、本文インデックスのファイルに入れるまとまり（「ファイル名=」の行から最後の行まで）の文字列を作る。
    #   book: @{ Name（元のファイル名）; Places（@{ Place（今の場所の名前）; Text（TSV の中身） } の並び） }
    #         または @{ Name; Block }（前の本文インデックスのファイルから取り出したまとまり。そのまま返す）
    param (
        $book
    )

    if ($null -ne $book.Block) {
        return [string]$book.Block
    }
    $mark = [string]${contentIndexMark}
    $name = [string]$book.Name
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append((getContentIndexBookHeadPrefix)).Append((encodeContentIndexValue $name)).Append("`n")
    $kind = getContentIndexFileKind $name
    if ($kind) {
        [void]$sb.Append("$mark 種類=$kind`n")
    }
    foreach ($place in $book.Places) {
        [void]$sb.Append((getContentIndexPlaceMetaText $name ([string]$place.Place)))
        [void]$sb.Append((convertToContentIndexBody ([string]$place.Text)))
    }
    return $sb.ToString()
}


function getContentIndexPlaceMetaText {
    # 場所 1 つ分のメタ情報の行（「<RS> キー=値」を LF で並べたもの）を返す
    param (
        [string]$book,
        [string]$place
    )

    $mark = [string]${contentIndexMark}
    $meta = convertPlaceToContentIndexMeta $book $place
    $sb = New-Object System.Text.StringBuilder
    foreach ($key in $meta.Keys) {
        [void]$sb.Append("$mark ").Append($key).Append("=").Append((encodeContentIndexValue ([string]$meta[$key]))).Append("`n")
    }
    return $sb.ToString()
}


function convertToContentIndexText {
    # 元のファイルの並びから、本文インデックスのファイルの文字列を作る（この順に書く）。
    #   books  : convertBookToContentIndexBlock に渡せるもの（@{ Name; Places } か @{ Name; Block }）の並び
    #   version: 1 行目に書く版（分けた塊を入れるファイルは contentIndexSplitVersion。既定は contentIndexVersion）
    param (
        $books,
        [int]$version = ${contentIndexVersion}
    )

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("$([string]${contentIndexMark}) 版=${version}`n")
    foreach ($book in $books) {
        [void]$sb.Append((convertBookToContentIndexBlock $book))
    }
    return $sb.ToString()
}


function planContentIndexParts {
    # フォルダ 1 つの、拡張子ごと・番号ごとの本文インデックスのファイルに、どの元のファイルを入れるかを決める。
    #   parts  : 前の本文インデックスのファイルの並び。@{ Extension; Part; Books（@{ Name; Block } の並び。ファイルの中の順）; Version（省略は 2。分けた塊のファイルは 3） }
    #   books  : 足す・入れ替える元のファイル（@{ Name; Block }。分けて入れるものは @{ Name; Chunks（塊の文字列の並び。splitContentIndexBookBlock の結果） }）。元のファイル名の順に足す
    #   remove : 外す元のファイル名
    #   maxBytes: この大きさ（UTF-16 のバイト数）以上の本文インデックスのファイルには、もう足さない
    # 決め方:
    #   ・入れ替える元のファイルは、今入っている本文インデックスのファイルの同じ位置で入れ替える（ほかの本文インデックスのファイルへ移さない）
    #   ・外す元のファイルは、入っている本文インデックスのファイルから外す
    #   ・新しい元のファイルは、その拡張子の最後の番号の本文インデックスのファイルが maxBytes 未満ならそこに、以上なら次の番号の新しい本文インデックスのファイルに足す
    #   ・分けた塊（Chunks）は、1 つずつ専用の本文インデックスのファイル（版 3）に入れる。その拡張子の最後の番号の次から続けて振り、ほかの元のファイルとはまぜない。
    #     分けた塊を入れたファイルには、あとから足す元のファイルも足さない（版 2 の最後のファイル、無ければ次の番号の新しい版 2 のファイルに足す）
    #   ・版 3 のファイルの塊は、同じ名前の元のファイルを入れ替える・外すときに、どの番号のファイルからも外す（入れ替えは、分けないなら版 2 のファイルに、分けるなら新しい版 3 のファイルに入れ直す）
    #   ・分けない元のファイルは 1 つの本文インデックスのファイルにおさめる（1 つで maxBytes を超えても、その 1 つで 1 つの本文インデックスのファイルにする）
    # @{ Extension; Part; Books; Changed（書き直しが要る）; Version（ファイルの 1 行目に書く版） } の並び（拡張子・番号の順）を返す。
    # 元のファイルが無くなった本文インデックスのファイルは、Books が空で Changed が $true（消す）
    param (
        $parts,
        $books,
        [string[]]$remove = @(),
        [long]$maxBytes = ${contentIndexFileMaxBytes}
    )

    $removeSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($name in $remove) { [void]$removeSet.Add($name) }
    $replace = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($book in $books) { $replace[[string]$book.Name] = $book }

    $plan = New-Object System.Collections.Generic.List[hashtable]
    foreach ($part in @($parts | Sort-Object { $_.Extension }, { $_.Part })) {
        $kept = New-Object System.Collections.Generic.List[object]
        $changed = $false
        $isSplit = ([int]$part.Version -eq ${contentIndexSplitVersion})
        foreach ($book in $part.Books) {
            $name = [string]$book.Name
            $next = if ($replace.ContainsKey($name)) { $replace[$name] } else { $null }
            if ($removeSet.Contains($name) -or ($isSplit -and $null -ne $next)) {
                $changed = $true
            } elseif ($null -ne $next -and $null -eq $next.Chunks) {
                $kept.Add($next)
                [void]$replace.Remove($name)
                $changed = $true
            } elseif ($null -ne $next) {
                $changed = $true
            } else {
                $kept.Add($book)
            }
        }
        $plan.Add(@{ Extension = [string]$part.Extension; Part = [int]$part.Part; Books = $kept; Changed = $changed; Version = $(if ($isSplit) { ${contentIndexSplitVersion} } else { ${contentIndexVersion} }) })
    }

    # 新しい元のファイル（どの本文インデックスのファイルにも無かったもの）を、元のファイル名の順に足す
    foreach ($book in @($replace.Values | Sort-Object { [string]$_.Name })) {
        addContentIndexBookToPlan $plan $book $maxBytes
    }
    return , @($plan | Sort-Object { $_.Extension }, { $_.Part })
}


function addContentIndexBookToPlan {
    # planContentIndexParts の計画（plan）に、どの本文インデックスのファイルにも無い元のファイル 1 つを足す。
    # 分ける元のファイル（Chunks）は塊ごとに新しい版 3 のファイルを、分けない元のファイルは版 2 の最後のファイル（maxBytes 以上なら次の番号の新しいファイル）に足す
    param (
        $plan,
        $book,
        [long]$maxBytes
    )

    $extension = getContentIndexExtension ([string]$book.Name)
    $last = $null
    $lastNumber = 0
    foreach ($p in $plan) {
        if ($p.Extension -ne $extension) { continue }
        if ($p.Part -gt $lastNumber) { $lastNumber = $p.Part }
        if ($p.Version -ne ${contentIndexSplitVersion} -and ($null -eq $last -or $p.Part -gt $last.Part)) { $last = $p }
    }
    if ($null -ne $book.Chunks) {
        foreach ($chunk in $book.Chunks) {
            $lastNumber++
            $single = New-Object System.Collections.Generic.List[object]
            $single.Add(@{ Name = [string]$book.Name; Block = [string]$chunk })
            $plan.Add(@{ Extension = $extension; Part = $lastNumber; Books = $single; Changed = $true; Version = ${contentIndexSplitVersion} })
        }
        return
    }
    if ($null -eq $last -or (measureContentIndexPartBytes $last.Books) -ge $maxBytes) {
        $last = @{ Extension = $extension; Part = $lastNumber + 1; Books = (New-Object System.Collections.Generic.List[object]); Changed = $true; Version = ${contentIndexVersion} }
        $plan.Add($last)
    }
    $last.Books.Add($book)
    $last.Changed = $true
}


function measureContentIndexPartBytes {
    # 本文インデックスのファイルの大きさ（UTF-16 のバイト数。BOM と版の行を含む）を、元のファイルのまとまり（@{ Name; Block }）の並びから求める
    param (
        $books
    )

    $chars = 1 + ("$([string]${contentIndexMark}) 版=${contentIndexVersion}`n").Length
    foreach ($book in $books) {
        $chars += ([string]$book.Block).Length
    }
    return [long]$chars * 2
}

function measureContentIndexPlaceChars {
    # 本文インデックスのファイルに入る、場所 1 つ分（メタ情報の行＋中身）の文字数を返す（convertBookToContentIndexBlock と同じ作り方）。
    # 元のファイル 1 つの大きさは、「ファイル名=」の行から最後の行までの先頭部分（Places を空にしたまとまり）にこの値を足して求める
    param (
        [string]$book,
        [string]$place,
        [string]$text
    )

    return [long]((getContentIndexPlaceMetaText $book $place).Length + (convertToContentIndexBody $text).Length)
}


function splitContentIndexBookBlock {
    # 元のファイル 1 つ（book: @{ Name; Places（@{ Place; Text } の並び） }）の、本文インデックスのファイルに入れるまとまりの文字列の並びを返す。
    # 1 つで maxBytes（UTF-16 のバイト数・BOM と版の行を含む）以下なら、分けずに 1 つ（convertBookToContentIndexBlock と同じ）。
    # 超えるなら、約 chunkBytes の塊に分ける。場所（シート・ページなど）の切れ目で切ることを優先し、1 つの場所が 1 つの塊に入らないときだけ
    # その場所を行の切れ目で切る（場所の途中から始まる塊には「開始行=<N>」を付ける）。どの塊も、頭の「ファイル名=」「種類=」の行を繰り返す。
    # 1 行がそれだけで maxBytes を超えるなど、maxBytes におさまらない塊ができるときは例外にする。
    # 分けるときは、塊にした場所の Text（book の Places の中身）を空にする（メモリを減らすため。分けたあとの book は使わない）
    param (
        $book,
        [long]$maxBytes = ${contentIndexSearchMaxBytes},
        [long]$chunkBytes = ${contentIndexFileMaxBytes}
    )

    $name = [string]$book.Name
    $head = convertBookToContentIndexBlock @{ Name = $name; Places = @() }
    # 本文インデックスのファイルの、BOM と版の行の文字数
    $base = 1 + ("$([string]${contentIndexMark}) 版=${contentIndexSplitVersion}`n").Length
    $total = $base + $head.Length
    foreach ($place in $book.Places) {
        $total += measureContentIndexPlaceChars $name ([string]$place.Place) ([string]$place.Text)
    }
    if ($total * 2 -le $maxBytes) {
        return , @((convertBookToContentIndexBlock $book))
    }
    $limit = [long][Math]::Floor($chunkBytes / 2) - $base
    $chunks = New-Object System.Collections.Generic.List[string]
    $current = @{ Sb = $null; Used = 0L }
    foreach ($place in $book.Places) {
        $body = convertToContentIndexBody ([string]$place.Text)
        # 塊にした場所の読んだ TSV は、もう使わない（大きな元のファイルで、メモリに同じ中身を何通りも持たないため。呼び出し側の book の Text を空にする）
        $place.Text = $null
        addContentIndexPlaceToChunks $chunks $current $head (getContentIndexPlaceMetaText $name ([string]$place.Place)) $body $limit
    }
    if ($null -ne $current.Sb) { $chunks.Add($current.Sb.ToString()) }
    foreach ($chunk in $chunks) {
        if (($base + $chunk.Length) * 2 -gt $maxBytes) { throw (getContentIndexTooLargeMessage) }
    }
    return , $chunks.ToArray()
}


function addContentIndexPlaceToChunks {
    # splitContentIndexBookBlock の、場所 1 つ分。開いている塊（current.Sb）に入るなら足す。入らなければ塊を閉じて（chunks に足す）新しい塊から始め、
    # 場所が 1 つの塊（limit 文字）に入らなければ、行の切れ目で切って続きの塊に「開始行」を付けて入れる。
    #   head: 塊の頭（「ファイル名=」「種類=」の行）  meta: 場所のメタ情報の行  body: 場所の中身（convertToContentIndexBody の結果）
    param (
        $chunks,
        $current,
        [string]$head,
        [string]$meta,
        [string]$body,
        [long]$limit
    )

    if ($null -ne $current.Sb -and $current.Used + $meta.Length + $body.Length -le $limit) {
        [void]$current.Sb.Append($meta).Append($body)
        $current.Used += $meta.Length + $body.Length
        return
    }
    if ($null -ne $current.Sb) { $chunks.Add($current.Sb.ToString()) }
    $lineNumber = 1
    $pos = 0
    do {
        $metaHere = if ($lineNumber -gt 1) { $meta + "$([string]${contentIndexMark}) 開始行=$lineNumber`n" } else { $meta }
        $current.Sb = New-Object System.Text.StringBuilder
        [void]$current.Sb.Append($head).Append($metaHere)
        $room = $limit - $head.Length - $metaHere.Length
        $end = $body.Length
        if ($body.Length - $pos -gt $room) {
            $cut = $(if ($room -gt 0) { $body.LastIndexOf([char]10, $pos + $room - 1, $room) } else { -1 })
            if ($cut -lt $pos) { $cut = $body.IndexOf([char]10, $pos) }
            $end = $cut + 1
        }
        $piece = $body.Substring($pos, $end - $pos)
        [void]$current.Sb.Append($piece)
        $current.Used = $head.Length + $metaHere.Length + $piece.Length
        $lineNumber += $piece.Length - $piece.Replace("`n", "").Length
        $pos = $end
        if ($pos -lt $body.Length) { $chunks.Add($current.Sb.ToString()) }
    } while ($pos -lt $body.Length)
}


function getContentIndexTooLargeMessage {
    # 元のファイル 1 つで本文インデックスのファイルが検索の上限（contentIndexSearchMaxBytes）を超えるときの、取り込み一覧のエラーの列の文言
    return ${contentIndexTooLargeMessage}
}


function testContentIndexCellPrefixed {
    # 場所の行が「セル番地 + 区切りのタブ + 文字」の形か（Excel の図形・コメント）を返す。検索は、この形の行では区切りのタブより後だけを照合する
    param (
        [string]$book,
        $meta
    )

    return testCellPrefixedPlace $book ([string]$meta["対象"])
}

function testCellPrefixedPlace {
    # 場所の行が「セル番地 + 区切りのタブ + 文字」の形か（元のファイルが Excel で、場所の種類が図形・コメント）を返す。
    # 種類はふつうの場所なら空。この判定はここだけに置き、検索結果の行（HitRow.IsCellPrefixed）もこれで決める
    param (
        [string]$book,
        [string]$kind
    )

    return ((getContentIndexFileKind $book) -eq "Excel") -and ($kind -eq ${placeKindShape} -or $kind -eq ${placeKindComment})
}


function readContentIndexPlaces {
    # 本文インデックスのファイルの文字列を読み、場所ごとの @{ Book; Location（今の場所の名前）; CellPrefixed（行の先頭がセル番地と区切りのタブか）;
    # FirstLine（この場所の最初の行が、場所の何行目か。分けた塊の途中からのときは「開始行」。ふつうは 1）; Start（中身の先頭の位置）; End（中身の終わりの次の位置） }
    # を先頭から順に返す。版が分からないときは例外にする（版 2 と、分けた塊のファイルの版 3 を読む）
    param (
        [string]$text
    )

    $places = New-Object System.Collections.Generic.List[hashtable]
    $mark = ${contentIndexMark}
    $placeKeys = ${contentIndexPlaceKeys}
    $book = ""
    $current = $null
    $meta = $null
    $versionSeen = $false
    $pos = $text.IndexOf($mark)
    while ($pos -ge 0) {
        $lineEnd = $text.IndexOf([char]10, $pos)
        if ($lineEnd -lt 0) { $lineEnd = $text.Length }
        $line = $text.Substring($pos + 2, [Math]::Max(0, $lineEnd - $pos - 2))
        $eq = $line.IndexOf("=")
        if ($eq -gt 0) {
            $key = $line.Substring(0, $eq)
            $value = decodeContentIndexValue $line.Substring($eq + 1)
            if ($key -eq "版") {
                if ($value -ne [string]${contentIndexVersion} -and $value -ne [string]${contentIndexSplitVersion}) {
                    throw (getIndexRebuildMessage)
                }
                $versionSeen = $true
            } elseif ($key -eq "ファイル名" -or $placeKeys -contains $key) {
                if ($current) {
                    addContentIndexPlace $places $current $meta $pos
                    $current = $null
                }
                if ($key -eq "ファイル名") {
                    $book = $value
                } else {
                    $meta = [ordered]@{ $key = $value }
                    $current = @{ Book = $book; Location = ""; CellPrefixed = $false; FirstLine = 1; Start = $lineEnd + 1; End = $lineEnd + 1 }
                }
            } elseif ($current) {
                $meta[$key] = $value
                $current.Start = $lineEnd + 1
            }
        }
        $pos = if ($lineEnd -lt $text.Length) { $text.IndexOf($mark, $lineEnd + 1) } else { -1 }
    }
    if ($current) {
        addContentIndexPlace $places $current $meta $text.Length
    }
    if (!$versionSeen -and $text.Length -gt 0) {
        throw (getIndexRebuildMessage)
    }
    return , $places
}


function addContentIndexPlace {
    # readContentIndexPlaces の、場所 1 つを閉じる処理。場所の終わりの位置を入れ、名前・セル番地の判定・開始行を決めて places に足す
    param (
        $places,
        $current,
        $meta,
        [int]$end
    )

    $current.End = $end
    # 開いてよい名前でない場所（書き換えたインデックスの .hta など）は、検索にもプレビューにも出さない
    if (!(testSourceBookName $current.Book)) {
        return
    }
    $current.Location = convertContentIndexMetaToPlace $meta
    $current.CellPrefixed = testContentIndexCellPrefixed $current.Book $meta
    # 開始行は正の整数だけを受け付ける（それ以外は 1。形が崩れていても行番号がずれるだけ）
    $firstLine = 0
    if ([int]::TryParse([string]$meta["開始行"], [ref]$firstLine) -and $firstLine -gt 0) {
        $current.FirstLine = $firstLine
    }
    $places.Add($current)
}


function splitContentIndexTextByBook {
    # 本文インデックスのファイルの文字列を、元のファイルごとのまとまり（@{ Name; Block }。Block は「ファイル名=」の行から次の「ファイル名=」の行の前まで）
    # に分けて先頭から順に返す（元のファイルを入れ替えるとき、変わらないファイルをそのまま写すため）。版が違えば例外にする
    param (
        [string]$text
    )

    $books = New-Object System.Collections.Generic.List[hashtable]
    if ($text.Length -eq 0) {
        return , $books
    }
    $head = getContentIndexBookHeadPrefix
    $firstEnd = $text.IndexOf([char]10)
    if ($firstEnd -lt 0 -or !(testContentIndexVersionLine $text.Substring(0, $firstEnd))) {
        throw (getIndexRebuildMessage)
    }
    $pos = $text.IndexOf("`n$head", [System.StringComparison]::Ordinal)
    while ($pos -ge 0) {
        $start = $pos + 1
        $lineEnd = $text.IndexOf([char]10, $start)
        if ($lineEnd -lt 0) {
            throw "インデックスのファイルの形が崩れています（ファイル名の行に改行がありません）"
        }
        $name = decodeContentIndexValue $text.Substring($start + $head.Length, $lineEnd - $start - $head.Length)
        $next = $text.IndexOf("`n$head", $lineEnd, [System.StringComparison]::Ordinal)
        $end = if ($next -ge 0) { $next + 1 } else { $text.Length }
        $books.Add(@{ Name = $name; Block = $text.Substring($start, $end - $start) })
        $pos = $next
    }
    return , $books
}


function getContentIndexBodyText {
    # 本文インデックスのファイルの文字列から、メタ情報の行（先頭が RS の行）を除いた中身だけを返す（検索用の索引の gram を作るため。
    # メタ情報の「ファイル名」「シート」などの語が入ると、その語で探したときにどのフォルダも候補になってしまう）
    param (
        [string]$text
    )

    if ($text.IndexOf([char]0x1E) -lt 0) {
        return $text
    }
    return [regex]::Replace($text, "(?m)^\x1E[^\n]*\n?", "")
}



function getIndexRebuildMessage {
    # 古い形式のインデックスを知らせる文言（1 か所）。本文インデックスの版の例外・一覧のステータス・インデクサのログ・zip の取り込みの断りで使う
    return ${indexRebuildMessage}
}


function getContentIndexDroppedText {
    # 検索の上限を超えていた前の本文インデックスのファイルを外したことを、ログ（黄色）に出す 1 行
    param (
        [string]$rel,
        [int]$count
    )

    return "    大きすぎるインデックスのファイルを外しました（${rel}）。中の ${count} 件は次の更新で取り込み直します"
}


function getContentIndexPublishSkippedText {
    # インデックス作成で、本文インデックスにまとめられず飛ばしたフォルダをログに出す 1 行（インデックスからの相対パス・段・理由）。
    # 段は「本文インデックス」（理由の頭に付けて、どこで止まったかが分かるようにする）
    param (
        [string]$rel,
        [string]$stage,
        [string]$reason
    )

    $shown = if ($rel) { $rel } else { "." }
    $oneLine = ($reason -replace "[\r\n]+", " ").Trim()
    return "    インデックスにまとめられなかったフォルダがあります（${shown}）: ${stage}: ${oneLine}。インデックスを作り直すと直ります"
}


function testContentIndexVersionLine {
    # 本文インデックスの 1 行目（RS + " 版=<値>"）が、読める版（版 2 か、分けた塊のファイルの版 3）かどうか
    param (
        [string]$line
    )

    $text = $line.TrimEnd("`r")
    $mark = [string]${contentIndexMark}
    return ($text -ceq "$mark 版=${indexFormatVersion}") -or ($text -ceq "$mark 版=${contentIndexSplitVersion}")
}


function getContentIndexTextVersion {
    # 本文インデックスのファイルの文字列の版（分けた塊のファイルなら contentIndexSplitVersion、それ以外は contentIndexVersion。1 行目だけを見る）
    param (
        [string]$text
    )

    $isSplit = $text.StartsWith("$([string]${contentIndexMark}) 版=${contentIndexSplitVersion}`n", [System.StringComparison]::Ordinal)
    return $(if ($isSplit) { ${contentIndexSplitVersion} } else { ${contentIndexVersion} })
}


function getIndexFormatState {
    # 目録とフォルダの有無から、ワークスペースと各インデックスの状態を決める（判断層。ファイルには触らない）。
    #   catalog        : readIndexFormat の結果 @{ Status（Ok・Missing・Broken・Newer・Older）; Indexes（名前 → 作成中・作り終えた） }
    #   oldNames       : 古い置き場（content_index\）に本文インデックスがあるインデックスの名前
    #   v2Names        : 今の置き場（store\content\）に本文インデックスがあるインデックスの名前
    #   hasLegacyIndex : 前の前の版（index\）のしるしがあるか
    # 戻り値: @{ State; Message; Indexes = 名前 → @{ State; Searchable; Message } }
    #   State: Current・Building・Rebuilding・Old・Empty・Broken（目録だけ壊れた。各インデックスは本文インデックスから見分ける）・Newer
    param (
        $catalog,
        [string[]]$oldNames = @(),
        [string[]]$v2Names = @(),
        [bool]$hasLegacyIndex = $false
    )

    $status = if ($catalog) { [string]$catalog.Status } else { "Missing" }
    $rows = @{}
    if ($status -eq "Ok" -and $catalog.Indexes) {
        foreach ($key in $catalog.Indexes.Keys) { $rows[[string]$key] = [string]$catalog.Indexes[$key] }
    }
    $oldSet = @{}
    foreach ($name in $oldNames) { $oldSet[[string]$name] = $true }
    $v2Set = @{}
    foreach ($name in $v2Names) { $v2Set[[string]$name] = $true }

    $names = New-Object System.Collections.Generic.List[string]
    foreach ($table in @($rows, $oldSet, $v2Set)) {
        foreach ($key in $table.Keys) {
            if (!$names.Contains([string]$key)) { $names.Add([string]$key) }
        }
    }

    $indexes = [ordered]@{}
    foreach ($name in ($names | Sort-Object)) {
        $inOld = $oldSet.ContainsKey($name)
        $inStore = $v2Set.ContainsKey($name)
        $finished = ($rows.ContainsKey($name) -and $rows[$name] -eq ${catalogRowFinished})
        $indexState = ${indexStateBuilding}
        if ($status -eq "Newer") {
            $indexState = ${indexStateNewer}
        } elseif ($finished -and !$inOld) {
            $indexState = ${indexStateCurrent}
        } elseif ($inStore -and $inOld) {
            $indexState = ${indexStateRebuilding}
        } elseif ($inStore) {
            $indexState = $(if ($finished) { ${indexStateCurrent} } else { ${indexStateBuilding} })
        } elseif ($inOld) {
            $indexState = ${indexStateOld}
        }
        $message = ""
        if ($indexState -eq ${indexStateOld}) { $message = ${indexRebuildMessage} }
        elseif ($indexState -eq ${indexStateRebuilding}) { $message = ${indexRebuildingMessage} }
        elseif ($indexState -eq ${indexStateNewer}) { $message = ${indexNewerMessage} }
        $searchable = ($indexState -eq ${indexStateCurrent} -or $indexState -eq ${indexStateBuilding} -or $indexState -eq ${indexStateRebuilding})
        $indexes[$name] = @{ State = $indexState; Searchable = $searchable; Message = $message }
    }

    $state = ${indexStateEmpty}
    if ($status -eq "Newer") {
        $state = ${indexStateNewer}
    } elseif ($status -eq "Broken") {
        $state = ${indexStateBroken}
    } elseif ($indexes.Count -gt 0) {
        $states = @($indexes.Values | ForEach-Object { $_.State })
        foreach ($candidate in @(${indexStateRebuilding}, ${indexStateOld}, ${indexStateBuilding})) {
            if ($states -contains $candidate) { $state = $candidate; break }
        }
        if ($state -eq ${indexStateEmpty}) { $state = ${indexStateCurrent} }
    } elseif ($hasLegacyIndex) {
        $state = ${indexStateOld}
    }

    $workspaceMessage = ""
    if ($state -eq ${indexStateOld}) { $workspaceMessage = ${indexRebuildMessage} }
    elseif ($state -eq ${indexStateRebuilding}) { $workspaceMessage = ${indexRebuildingMessage} }
    elseif ($state -eq ${indexStateNewer}) { $workspaceMessage = ${indexNewerMessage} }
    return @{ State = $state; Message = $workspaceMessage; Indexes = $indexes }
}
