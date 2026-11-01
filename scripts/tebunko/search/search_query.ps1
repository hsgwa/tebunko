# 検索の条件（正規表現・対象ファイル）の組み立て（判断層。ファイルに触らない）。

function isValidRegex {
    # 正規表現として正しいか
    param (
        [string]$pattern
    )

    try {
        [void][regex]::new($pattern)
        return $true
    } catch {
        return $false
    }
}


# 1行の照合にかけてよい時間。正規表現によっては終わらなくなるため、超えたら検索を止める
${regexTimeout} = [timespan]::FromSeconds(5)

function getRegexScanMode {
    # 正規表現を TSV の全文（改行は LF にそろえ、^ $ は行の先頭・末尾に一致させる）にかけてよいかを返す。
    # 1 行ずつ PowerShell で照合すると遅いため、全文への照合（.NET の中で走る）で済ませられる場合を見分ける。
    #   "lines" : 一致は必ず 1 行の中に収まり、行の外の文字を見ない。全文での一致の位置から、ヒットした行が分かる
    #   "filter": 行をまたいで一致することはあるが、1 行で一致するなら全文でも必ず一致する。
    #             全文で一致しない TSV は読み飛ばせる（一致した TSV は 1 行ずつ照合し直す）
    #   "scan"  : 全文では 1 行と結果が変わりうる（\A \z、否定の先読み・後読み、(?i) 等）。1 行ずつ照合する
    # 判定は安全側に倒す（分からない書き方は "filter" か "scan"）。検索結果は照合のしかたによらず同じになる
    param (
        [string]$pattern
    )

    $lines = $true
    $inClass = $false
    $classStart = 0
    $length = $pattern.Length
    for ($i = 0; $i -lt $length; $i++) {
        $c = $pattern[$i]
        if ($c -eq '\') {
            $i++
            if ($i -ge $length) {
                return "scan"
            }
            $e = $pattern[$i]
            if ("AzZG".IndexOf($e) -ge 0) {
                return "scan"   # 文字列の先頭・末尾は、1 行と全文とで位置が違う
            }
            if ("sWDpPnrxuc0123456789".IndexOf($e) -ge 0) {
                $lines = $false  # 改行にも一致しうる（\s \W \D \p{..} \n \x0A \u000A \cJ 8 進数・後方参照）
            }
            if ($inClass -and $i + 1 -lt $length -and $pattern[$i + 1] -eq '-') {
                $lines = $false  # エスケープから始まる範囲（[\t-z] 等）は改行を含みうる
            }
            continue
        }
        if ([int]$c -lt 0x20) {
            $lines = $false      # 制御文字をそのまま含む（改行・タブから始まる範囲など）
        }
        if ($inClass) {
            if ($c -eq ']' -and $i -gt $classStart) {
                $inClass = $false
            }
            continue
        }
        if ($c -eq '[') {
            $inClass = $true
            if ($i + 1 -lt $length -and $pattern[$i + 1] -eq '^') {
                $lines = $false  # 否定の文字クラスは改行に一致する
                $i++
            }
            $classStart = $i + 1  # 先頭の ] は文字として扱われる
            continue
        }
        if ($c -eq '(' -and $i + 1 -lt $length -and $pattern[$i + 1] -eq '?') {
            $rest = $pattern.Substring($i + 2)
            if ($rest.StartsWith(":") -or $rest.StartsWith(">")) {
                continue
            }
            if ($rest.StartsWith("=") -or $rest.StartsWith("<=")) {
                $lines = $false  # 肯定の先読み・後読みは行の外（改行）を見る。全文で一致が増えるだけなので "filter" にはできる
                continue
            }
            if ($rest -match "^(<[A-Za-z_]|'[A-Za-z_])") {
                continue         # 名前付きグループ
            }
            return "scan"        # 否定の先読み・後読み、(?i) 等のオプション、条件分岐、コメント
        }
    }
    if ($lines) {
        return "lines"
    }
    return "filter"
}


function newSearchRegex {
    # 検索条件から、検索と一致箇所の強調で使う正規表現を作る（画面の検索と結果の表示で共通）。
    #   simpleMatch  : $true ならワードを文字どおりに探す。$false なら正規表現として探し、正規表現として不正なら文字どおりに探す
    #   caseSensitive: 英字の大文字・小文字を区別する（既定は区別しない）
    # @{ Regex; SimpleMatch（実際に文字どおり探すか）; TextRegex; ScanMode } を返す。
    #   TextRegex: TSV の全文にかける正規表現（Regex に Multiline を足したもの。^ $ が行の先頭・末尾に一致する）
    #   ScanMode : TextRegex を全文にかけてよいか（getRegexScanMode）
    param (
        [string]$word,
        [bool]$simpleMatch = $true,
        [bool]$caseSensitive = $false
    )

    if (!$simpleMatch -and !(isValidRegex $word)) {
        $simpleMatch = $true
    }
    $pattern = if ($simpleMatch) { [regex]::Escape($word) } else { $word }
    # CultureInvariant: 大文字・小文字を区別しないときの照合が、区別する場合と同程度に速くなる（日本語の照合結果は変わらない）
    $options = [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
    if (!$caseSensitive) {
        $options = $options -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    }
    $textOptions = $options -bor [System.Text.RegularExpressions.RegexOptions]::Multiline
    return @{
        Regex       = New-Object System.Text.RegularExpressions.Regex($pattern, $options, ${regexTimeout})
        SimpleMatch = $simpleMatch
        TextRegex   = New-Object System.Text.RegularExpressions.Regex($pattern, $textOptions, ${regexTimeout})
        ScanMode    = getRegexScanMode $pattern
    }
}

function newPlaceExclude {
    # 検索から外す場所（図形・コメント）の正規表現を返す。どちらも検索するなら $null。
    # 場所の名前は "<元の場所>[<種類>]"（index_name.ps1 の objectPlacePattern）。Excel・Word・PowerPoint で共通
    param (
        [bool]$includeShapes = $true,
        [bool]$includeComments = $true
    )

    $kinds = @()
    # 埋め込み（"[埋め込み<N>]"）は図形と一緒に切り替える（検索の選択肢を増やさない）
    if (!$includeShapes) { $kinds += ${placeKindShape}; $kinds += ${placeKindEmbed} }
    if (!$includeComments) { $kinds += ${placeKindComment} }
    if ($kinds.Count -eq 0) {
        return $null
    }
    return [regex]::new((newObjectPlaceKindPattern -kinds $kinds))
}

# テキストのヒットの行・プレビューの前後の行を切る長さ（文字）と、一致の位置から遡る長さ（文字）。
# 1 行が数 MB になりうる .json・.xml・.log の描画で画面が固まらないよう、取り込みでは切らず、表示・出力の直前だけ切る
${hitLineMaxChars} = 1000
${hitLineBeforeMatchChars} = 200

function truncateHitLine {
    # 長い行（テキストのヒットの行・前後の行）を、一致の位置を中心に hitLineMaxChars 文字に切る。
    # 一致の位置（matchIndex。行の中の位置。分からなければ -1）の hitLineBeforeMatchChars 文字前から hitLineMaxChars 文字を取り、
    # 省いた側に "…" を付ける。hitLineMaxChars 以下ならそのまま返す
    param (
        [string]$line,
        [int]$matchIndex = -1
    )

    if ($line.Length -le ${hitLineMaxChars}) {
        return $line
    }
    $start = if ($matchIndex -lt 0) { 0 } else { [Math]::Max(0, $matchIndex - ${hitLineBeforeMatchChars}) }
    $end = [Math]::Min($line.Length, $start + ${hitLineMaxChars})
    $result = $line.Substring($start, $end - $start)
    if ($start -gt 0) { $result = "…" + $result }
    if ($end -lt $line.Length) { $result = $result + "…" }
    return $result
}

function getFileKindExtensions {
    # 検索バーのチップの種類（excel・word・powerpoint・text・pdf）ごとの拡張子（小文字）を返す。知らない種類は空
    param (
        [string]$kind
    )

    switch ($kind) {
        "excel" { return @(getOfficeExtensionsOfKind "Excel") }
        "word" { return @(getOfficeExtensionsOfKind "Word") }
        "powerpoint" { return @(getOfficeExtensionsOfKind "PowerPoint") }
        "text" { return @(${textExtensions}) }
        "pdf" { return @(${pdfExtensions}) }
    }
    return @()
}

function newFileKindFilter {
    # 選んだ種類（readFileKinds の値）から、対象ファイルの条件（newFileFilter に渡す文字列。"*.xlsx;*.xls" の形）を作る。
    # すべての種類を選んでいる・空・知らない種類だけのときは空（絞り込まない）
    param (
        [object[]]$kinds
    )

    $chosen = @(${fileKindNames} | Where-Object { @($kinds) -contains $_ })
    if ($chosen.Count -eq 0 -or $chosen.Count -eq ${fileKindNames}.Count) {
        return ""
    }
    return (@($chosen | ForEach-Object { getFileKindExtensions $_ } | ForEach-Object { "*$_" }) -join ";")
}

function newFileFilter {
    # 対象ファイルの指定（例: "*.xlsx;見積*;!*old*"）を、元のファイル名に対する正規表現 @{ Include; Exclude } にする（無い側は $null）。
    #   ; で区切る（全角の ； も可）。! で始まるものは除外。* は任意の文字列、? は任意の1文字。大文字・小文字を区別しない
    #   * も ? も無いものは部分一致（「見積」は「*見積*」）
    param (
        [string]$filter
    )

    $include = New-Object System.Collections.Generic.List[string]
    $exclude = New-Object System.Collections.Generic.List[string]
    foreach ($item in ([string]$filter).Split([char[]]";；")) {
        $item = $item.Trim()
        $list = $include
        if ($item.StartsWith("!") -or $item.StartsWith("！")) {
            $list = $exclude
            $item = $item.Substring(1).Trim()
        }
        if ($item -eq "") {
            continue
        }
        if ($item.IndexOfAny([char[]]"*?") -lt 0) {
            $item = "*${item}*"
        }
        $list.Add("^" + [regex]::Escape($item).Replace("\*", ".*").Replace("\?", ".") + "$")
    }

    $options = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
    $result = @{ Include = $null; Exclude = $null }
    if ($include.Count -gt 0) {
        $result.Include = New-Object System.Text.RegularExpressions.Regex(($include -join "|"), $options)
    }
    if ($exclude.Count -gt 0) {
        $result.Exclude = New-Object System.Text.RegularExpressions.Regex(($exclude -join "|"), $options)
    }
    return $result
}

function testHitFilter {
    # 結果の絞り込み（ファイル名の一部）に、元のファイル名が合うか。空の絞り込みは常に合う。大文字小文字は区別しない。
    # 表（HitRow.MatchesFileName）と、結果の出力（溢れの行。search_export.ps1）で同じ判定を使う
    param (
        [string]$book,
        [string]$filterText
    )

    if ($filterText -eq "") {
        return $true
    }
    return (("" + $book).IndexOf($filterText, [System.StringComparison]::CurrentCultureIgnoreCase) -ge 0)
}


function testLiteralSearchWord {
    # 検索語が、照合の式として文字どおりの 1 つの文字列か。文字どおりに探す指定か、正規表現を使う指定でも語に正規表現の記号が無いとき（Escape しても変わらない）は $true
    param (
        [string]$word,
        [bool]$simpleMatch
    )

    return (newSearchRegex $word $simpleMatch $false).SimpleMatch -or ([regex]::Escape($word) -eq $word)
}

function toSearchFolderKeys {
    # 検索対象（文字列、または @{ Root; RelPath; Recurse } の並び）を、比べるための "Root|RelPath|Recurse" の並び（大文字・小文字は区別しない）にする。
    # 文字列は getContentIndexFiles と同じく @{ Root = <文字列>; RelPath = ""; Recurse = $true } とみなす
    param (
        $folders
    )

    return @(@($folders) | ForEach-Object {
        $target = $_
        if ($target -is [string]) {
            $target = @{ Root = $target; RelPath = ""; Recurse = $true }
        }
        ("{0}|{1}|{2}" -f [string]$target.Root, ([string]$target.RelPath).Trim("\"), [bool]$target.Recurse).ToLowerInvariant()
    })
}

function getNarrowSearchReason {
    # 検索語が前の検索の語より長くなったとき、前の検索のヒットの行だけを照合し直してよいかを返す（前の結果の使い回し）。
    # "Narrow"（絞ってよい）か、絞らない理由の符号を返す。インデックスが変わっていないこと（今の列挙が前と同じ）は、状態層で確かめる。
    #   previous: 前の検索の記録（Word; SimpleMatch; CaseSensitive; FileFilter; IncludeShapes; IncludeComments; Folders; WorkDir; Limit）。無ければ $null
    #   request : 新しい要求（同じ項目を持つ）
    # 符号: NoRecord（記録なし）/ NoLimit（上限なし）/ Regex（どちらかの語が正規表現）/ NotLonger（新しい語が前の語より長くない）/
    #       NotContained（新しい語の中に前の語が無い）/ Condition（語以外の条件が違う）
    param (
        $previous,
        $request
    )

    if ($null -eq $previous) { return "NoRecord" }
    if ([int]$previous.Limit -le 0 -or [int]$request.Limit -le 0) { return "NoLimit" }
    $previousWord = [string]$previous.Word
    $word = [string]$request.Word
    if (!(testLiteralSearchWord $previousWord ([bool]$previous.SimpleMatch)) -or !(testLiteralSearchWord $word ([bool]$request.SimpleMatch))) { return "Regex" }
    if ($word.Length -le $previousWord.Length) { return "NotLonger" }
    $same = ([bool]$previous.CaseSensitive -eq [bool]$request.CaseSensitive) -and
        ([string]$previous.FileFilter -ceq [string]$request.FileFilter) -and
        ([bool]$previous.IncludeShapes -eq [bool]$request.IncludeShapes) -and
        ([bool]$previous.IncludeComments -eq [bool]$request.IncludeComments) -and
        ([string]$previous.WorkDir).Equals([string]$request.WorkDir, [System.StringComparison]::OrdinalIgnoreCase) -and
        ([int]$previous.Limit -eq [int]$request.Limit)
    if (!$same) { return "Condition" }
    $previousKeys = @(toSearchFolderKeys $previous.Folders)
    $keys = @(toSearchFolderKeys $request.Folders)
    if ($previousKeys.Count -ne $keys.Count) { return "Condition" }
    for ($i = 0; $i -lt $keys.Count; $i++) {
        if ($previousKeys[$i] -ne $keys[$i]) { return "Condition" }
    }
    $options = [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
    if (![bool]$request.CaseSensitive) { $options = $options -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase }
    if (![regex]::new([regex]::Escape($previousWord), $options).IsMatch($word)) { return "NotContained" }
    return "Narrow"
}
