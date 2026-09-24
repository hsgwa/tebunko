# 高速検索（Windows Search で検索語を含みうるフォルダを先に絞る）の決まり（判断層。ファイルにも画面にも触らない）。
# Windows インデックス（system_index の txt）は、文字の 2-gram を英数字の語にしたものを空白区切りで書く。
# Windows Search は語単位でしか一致を取らないため、本文をそのまま索引させると語の途中からの一致を落とすが、
# 2-gram の語なら、ワードを含む本文の txt には、ワードのすべての 2-gram が必ず入っている。

# Windows インデックスの名前（置き場所は paths_grep.ps1 の ${systemIndexDir}）
${windowsIndexFileName}      = "Windowsインデックス.txt"
${windowsIndexFileLike}      = "Windowsインデックス%"      # 問い合わせの LIKE（分けたものも含む）
${windowsIndexSplitLike}     = "Windowsインデックス[_]%"   # 分けたもの（Windowsインデックス_1.txt …）だけ

# 1 つの txt の大きさの上限（これを超えたら語の範囲で分ける。実測では 16MB までは末尾まで索引された）
${windowsIndexPartBytes} = 8MB
# txt のパスの長さの上限（これ以上は作らず、そのフォルダは .NET で照合する。Windows Search が長いパスを索引しないおそれがあるため）
${windowsIndexPathMax} = 240
# 1 回の問い合わせに使う語の数の上限（多いときは均等に間引く。間引いても候補が増えるだけで漏れない）
${searchGramMax} = 16

# 状態ファイルの行の種別
${windowsIndexCovered}  = "対応済み"   # インデックスのすべてのフォルダに txt がある（または対象外として記録した）
${windowsIndexPending}  = "反映待ち"   # txt を書いた（値は txt の更新日時。UTC の Ticks）
${windowsIndexExcluded} = "対象外"     # txt を作らなかったフォルダ（常に .NET で照合する）

function convertToGramToken {
    # 2 文字（小文字にしたもの）を語にする: x ＋ UTF-16LE の 4 バイトの 16 進（例: "ニタ" → xcb30bf30）
    param (
        [string]$pair
    )

    return "x" + [System.BitConverter]::ToString([System.Text.Encoding]::Unicode.GetBytes($pair)).Replace("-", "").ToLowerInvariant()
}

function getSearchGrams {
    # 検索語から、問い合わせに使う語を返す（重複なし・最大 searchGramMax 個）。
    # 空白で区切り、2 文字以上の部分の隣り合う 2 文字を語にする。語が無ければ空（高速検索は使えない）
    param (
        [string]$word
    )

    $grams = New-Object System.Collections.Generic.List[string]
    $seen = New-Object System.Collections.Generic.HashSet[string]
    foreach ($part in ([string]$word).ToLowerInvariant() -split "[\s\u2028]+") {
        for ($i = 0; $i -lt $part.Length - 1; $i++) {
            $token = convertToGramToken $part.Substring($i, 2)
            if ($seen.Add($token)) {
                $grams.Add($token)
            }
        }
    }
    if ($grams.Count -le ${searchGramMax}) {
        return , $grams.ToArray()
    }
    # 均等に間引く（先頭と末尾は残す）
    $picked = New-Object System.Collections.Generic.List[string]
    for ($k = 0; $k -lt ${searchGramMax}; $k++) {
        $picked.Add($grams[[int][Math]::Round($k * ($grams.Count - 1) / (${searchGramMax} - 1))])
    }
    return , $picked.ToArray()
}

function testFastSearchUsable {
    # 高速検索を使えるか（画面の「使用可 / 使用不可」）。
    #   available: Windows Search を開けて、system_index が索引の対象（windows_search.ps1 の testWindowsSearch）
    param (
        [bool]$available,
        [bool]$useRegex,
        [string]$word
    )

    return ($available -and !$useRegex -and (getSearchGrams $word).Count -gt 0)
}

function addTextGrams {
    # 本文の 2-gram を、語の値（uint32。UTF-16LE の 2 文字をそのまま 4 バイトの整数にしたもの）として set に足す。
    # 1 文字ずつ PowerShell で回すと遅いため、char[] を uint32[] に写して .NET の HashSet で重複を除く。
    # 空白を含む 2-gram も入るが、問い合わせには使わないため害は無い（除く方が遅い）
    param (
        [System.Collections.Generic.HashSet[uint32]]$set,
        [string]$text
    )

    $chars = $text.ToLowerInvariant().ToCharArray()
    $even = $chars.Length -shr 1
    $odd = ($chars.Length - 1) -shr 1
    if ($even -gt 0) {
        $u0 = New-Object 'uint32[]' $even
        [System.Buffer]::BlockCopy($chars, 0, $u0, 0, $even * 4)
        $set.UnionWith($u0)
    }
    if ($odd -gt 0) {
        # 1 文字ずらした並びは、2 バイトずらして写せば作れる
        $u1 = New-Object 'uint32[]' $odd
        [System.Buffer]::BlockCopy($chars, 2, $u1, 0, $odd * 4)
        $set.UnionWith($u1)
    }
}

function convertToGramText {
    # 語の値（昇順に並べたもの）の start から count 個を、txt に書く文字列にする（x ＋ 16 進 8 桁を空白区切り）
    param (
        [uint32[]]$values,
        [int]$start = 0,
        [int]$count = -1
    )

    if ($count -lt 0) {
        $count = $values.Length - $start
    }
    if ($count -le 0) {
        return ""
    }
    $bytes = New-Object 'byte[]' ($count * 4)
    [System.Buffer]::BlockCopy($values, $start * 4, $bytes, 0, $count * 4)
    $hex = [System.BitConverter]::ToString($bytes).Replace("-", "").ToLowerInvariant()
    return [regex]::Replace($hex, "(.{8})", 'x$1 ')
}

function getGramPartCount {
    # 語の数から、txt をいくつに分けるかを返す（1 語は "x" ＋ 8 桁 ＋ 空白の 10 バイト）
    param (
        [int]$count,
        [long]$partBytes = ${windowsIndexPartBytes}
    )

    $perPart = [Math]::Max(1, [int][Math]::Floor($partBytes / 10))
    return [Math]::Max(1, [int][Math]::Ceiling($count / $perPart))
}

function getWindowsIndexFileNames {
    # txt の名前（分けないときは 1 つ。分けるときは Windowsインデックス_1.txt …）
    param (
        [int]$parts
    )

    if ($parts -le 1) {
        return , @(${windowsIndexFileName})
    }
    $base = [System.IO.Path]::GetFileNameWithoutExtension(${windowsIndexFileName})
    return , @(1..$parts | ForEach-Object { "${base}_$_.txt" })
}

function testWindowsIndexPath {
    # txt のパスが長すぎないか（分けたときは最も長い名前で判定する）
    param (
        [string]$folder,
        [int]$parts = 1
    )

    $longest = @(getWindowsIndexFileNames $parts | Sort-Object Length -Descending)[0]
    return ("$folder\$longest").Length -lt ${windowsIndexPathMax}
}

function convertFolderRoot {
    # フォルダのパスの先頭（fromRoot）を toRoot に置き換える（system_index と index は同じ相対パスの構成）。
    # fromRoot の外のパスは $null
    param (
        [string]$path,
        [string]$fromRoot,
        [string]$toRoot
    )

    $from = $fromRoot.TrimEnd("\")
    if ($path.Equals($from, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $toRoot.TrimEnd("\")
    }
    if ($path.StartsWith("$from\", [System.StringComparison]::OrdinalIgnoreCase)) {
        return $toRoot.TrimEnd("\") + $path.Substring($from.Length)
    }
    return $null
}

function convertItemUrl {
    # Windows Search の System.ItemUrl（file:C:/…）をパスにする。% は符号化されていないため URL として戻さない
    param (
        [string]$url
    )

    return $url.Substring($url.IndexOf(":") + 1).Replace("/", "\")
}

function convertToScopeUrl {
    # SCOPE に書く URL（file:C:/…）。SQL の文字列に入れるため ' は重ねる
    param (
        [string]$path
    )

    return ("file:" + $path.TrimEnd("\").Replace("\", "/")).Replace("'", "''")
}

function newWindowsIndexQuery {
    # 候補の txt（分けていないもの）を探す問い合わせ
    param (
        [string]$scopeDir,
        [string[]]$grams
    )

    $condition = ($grams | ForEach-Object { "`"$_`"" }) -join " AND "
    return "SELECT System.ItemUrl FROM SystemIndex WHERE SCOPE='$(convertToScopeUrl $scopeDir)' AND System.FileName = '${windowsIndexFileName}' AND CONTAINS(System.Search.Contents, '$condition')"
}

function newWindowsIndexSplitQuery {
    # 分けた txt から、語を 1 つ含むものを探す問い合わせ（分けた txt は AND で 1 回に問えないため語ごとに問う）
    param (
        [string]$scopeDir,
        [string]$gram
    )

    return "SELECT System.ItemUrl FROM SystemIndex WHERE SCOPE='$(convertToScopeUrl $scopeDir)' AND System.FileName LIKE '${windowsIndexSplitLike}' AND CONTAINS(System.Search.Contents, '`"$gram`"')"
}

function newWindowsIndexStateQuery {
    # 反映の判定に使う値（GatherTime・DateModified）を取る問い合わせ
    param (
        [string]$scopeDir
    )

    return "SELECT System.ItemUrl, System.Search.GatherTime, System.DateModified FROM SystemIndex WHERE SCOPE='$(convertToScopeUrl $scopeDir)' AND System.FileName LIKE '${windowsIndexFileLike}'"
}

function testWindowsIndexReflected {
    # txt が Windows Search に反映済みか: 本文を読み終えた（GatherTime がある）かつ、今の版を索引した
    # （DateModified が txt の更新日時を秒で切り捨てた値と同じ。Windows Search は秒未満を切り捨てて持つ）
    param (
        $gatherTime,
        $dateModified,
        [long]$fileTicksUtc
    )

    if ($null -eq $gatherTime -or $gatherTime -is [System.DBNull] -or $null -eq $dateModified -or $dateModified -is [System.DBNull]) {
        return $false
    }
    return ([datetime]$dateModified).Ticks -eq ($fileTicksUtc - ($fileTicksUtc % [timespan]::TicksPerSecond))
}

function testFolderInTarget {
    # フォルダ（TSV のフォルダ）が検索対象に入るか。recurse が $false なら対象フォルダそのものだけ
    param (
        [string]$folder,
        [string]$targetDir,
        [bool]$recurse
    )

    $target = $targetDir.TrimEnd("\")
    if ($folder.Equals($target, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }
    return ($recurse -and $folder.StartsWith("$target\", [System.StringComparison]::OrdinalIgnoreCase))
}

function convertFromWindowsIndexState {
    # 状態ファイルの行を読む（パスは system_index からの相対パス）。形の違う行は無視する。
    #   Covered : 対応済みのインデックス名 / Pending: txt の相対パス → 書いた日時（Ticks）/ Excluded: 対象外のフォルダの相対パス
    param (
        [string[]]$lines
    )

    $state = newWindowsIndexState
    foreach ($line in $lines) {
        $cells = ([string]$line).Split("`t")
        if ($cells.Count -lt 2 -or $cells[1] -eq "") {
            continue
        }
        switch ($cells[0]) {
            ${windowsIndexCovered} { [void]$state.Covered.Add($cells[1]) }
            ${windowsIndexExcluded} { [void]$state.Excluded.Add($cells[1]) }
            ${windowsIndexPending} {
                $ticks = 0L
                if ($cells.Count -ge 3 -and [long]::TryParse($cells[2], [ref]$ticks)) {
                    $state.Pending[$cells[1]] = $ticks
                }
            }
        }
    }
    return $state
}

function newWindowsIndexState {
    # 空の状態
    return @{
        Covered  = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
        Pending  = New-Object 'System.Collections.Generic.Dictionary[string,long]' ([System.StringComparer]::OrdinalIgnoreCase)
        Excluded = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    }
}

function convertToWindowsIndexState {
    # 状態を状態ファイルの行にする（対応済み・対象外・反映待ちの順。それぞれ名前順）
    param (
        $state
    )

    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($name in ($state.Covered | Sort-Object)) {
        $lines.Add("${windowsIndexCovered}`t$name`t")
    }
    foreach ($rel in ($state.Excluded | Sort-Object)) {
        $lines.Add("${windowsIndexExcluded}`t$rel`t")
    }
    foreach ($rel in ($state.Pending.Keys | Sort-Object)) {
        $lines.Add("${windowsIndexPending}`t$rel`t$($state.Pending[$rel])")
    }
    return , $lines.ToArray()
}

function getIndexNameOfRelPath {
    # system_index（index）からの相対パスの先頭（インデックス名）
    param (
        [string]$relPath
    )

    return ([string]$relPath).Split("\")[0]
}
