# 制限モードの検索結果をブック（xlsx）にする（状態層）。制限言語モードで動く書き方だけで書く。
#
# 行の組み立て（newResultRows）→ XML（result_book_view.ps1）→ 部品を作業フォルダに書いて tar.exe で ZIP にまとめる（writeResultBook）。
# 結果のブックは検索ごとに時刻入りの名前で TEMP の tebunko_grep\検索結果 に作り、起動のたびに古いものを消す（removeOldResultBooks）。

${resultBookDir} = Join-Path ${tempRoot} "tebunko_grep\検索結果"

# ヒットの前後に付ける行の数（畳んで隠しておき、+ を押すと見える）
${resultContextLines} = 2

# tar.exe（Windows 10 1803 以降に入っている。PATH の先に Git の GNU tar があることがあるため、フルパスで呼ぶ）
${tarExe} = Join-Path $env:windir "System32\tar.exe"

function newResultRows {
    # 検索結果のシートの行（result_book_view.ps1 の形）を作る: @{ Rows; ColumnCount }。
    #   1 行目: 見出し（ファイル・場所・種別・行・A・B…）
    #   ファイルごとのまとまりの行: ファイル名（元のファイルへのリンク）・フォルダ（フォルダへのリンク）・件数
    #   その下にヒットの行（一致した部分に色を付ける。Excel はリンクで該当のシートとセルを開く）、
    #   さらにその下に前後の行（畳んで隠す）
    # 行をセルに分けるのは、Excel の TSV は " で囲んだセルを 1 つにし（splitTsvCells と同じ）、セル内改行（U+2028）を改行に戻す。
    # Word・PowerPoint の行（段落、表の 1 行）はタブで分ける（検索結果ファイルの toResultLine と同じ扱い）。
    # 行が多い（前後の行を含めて 5 万行）ため、行ごとに関数を呼ばずにここで分ける。
    #   hits : searchRestricted の Hits
    #   regex: 一致した部分に色を付ける正規表現（newSearchRegex の Regex）
    #   maps : resolveSourcePath のキャッシュ
    param (
        [object[]]$hits,
        [regex]$regex,
        [hashtable]$maps = @{},
        [int]$contextLines = ${resultContextLines}
    )

    $cellPattern = [regex]::new('\t(?:"(?:[^"]|"")*"[^\t]*|[^\t]*)')
    $quoted = [regex]::new('^"((?:[^"]|"")*)"(.*)$', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    $newLine = [string][char]10

    # 行は番号 → 行のハッシュテーブルに入れる（配列に += で足すと、件数が多いと遅いため）
    $rows = @{}
    $rows[0] = $null  # 見出しは最後に作る（列の数が決まってから）
    $maxCells = 0
    $lineCache = @{}
    $placeCache = @{}
    $sheetCache = @{}
    $folderCache = @{}
    $columnNames = @{}
    $groupRow = $null
    $groupKey = $null
    $groupCount = 0
    $source = $null
    foreach ($hit in $hits) {
        $key = "$($hit.Root)|$($hit.RelDir)|$($hit.Book)"
        $isExcel = $hit.Book -match "\.xls[a-z]?$"
        if ($key -ne $groupKey) {
            if ($groupRow) { $groupRow.Texts[2] = "$groupCount 件" }
            $groupKey = $key
            $groupCount = 0
            # 元のファイルの場所は、同じフォルダのファイルなら同じ元のフォルダの下にあるため、フォルダごとに 1 回だけ求める
            $folderKey = $hit.Root + "|" + $hit.RelDir
            if (!$folderCache.ContainsKey($folderKey)) {
                $resolved = resolveSourcePath $hit $maps
                $folderCache[$folderKey] = if ($resolved) { getPathParent $resolved } else { "" }
            }
            $source = if ($folderCache[$folderKey]) { $folderCache[$folderKey] + "\" + $hit.Book } else { $null }
            $title = if ($hit.RelDir) { "$($hit.RelDir)\$($hit.Book)" } else { $hit.Book }
            $groupRow = @{ Style = ${resultStyleGroup}; Level = 0; Texts = @($title, "元のフォルダが分かりません", "") }
            if ($source) {
                $groupRow.Texts[1] = "フォルダを開く"
                $groupRow.Styles = @{ 0 = ${resultStyleGroupLink}; 1 = ${resultStyleGroupLink} }
                $groupRow.Links = @{ 0 = @{ Target = $source }; 1 = @{ Target = (getPathParent $source) } }
            }
            $rows[$rows.Count] = $groupRow
        }
        $groupCount++

        # ヒットの行と、前後の行（同じ TSV の中）
        $lines = $null
        $firstLine = $hit.LineNumber
        $lastLine = $hit.LineNumber
        if ($contextLines -gt 0) {
            $path = "\\?\" + $hit.Root + "\" + $hit.RelPath
            if ($hit.Root.StartsWith("\\")) { $path = "\\?\UNC\" + $hit.Root.Substring(2) + "\" + $hit.RelPath }
            if (!$lineCache.ContainsKey($path)) {
                # 行ごとの Get-Content は遅い（1 行ずつオブジェクトを作る）ため、丸ごと読んでから分ける
                $raw = Get-Content -LiteralPath $path -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
                $lineCache[$path] = if ($null -eq $raw) { [string[]]@() } else { readRestrictedTsvLines $raw.Replace("`r`n", "`n").Replace("`r", "`n") }
            }
            $lines = $lineCache[$path]
            $firstLine = $hit.LineNumber - $contextLines
            if ($firstLine -lt 1) { $firstLine = 1 }
            $lastLine = $hit.LineNumber + $contextLines
            if ($lastLine -gt $lines.Count) { $lastLine = $lines.Count }
            if ($lastLine -lt $hit.LineNumber) { $lastLine = $hit.LineNumber }
        }
        $contextRows = @()
        for ($n = $firstLine; $n -le $lastLine; $n++) {
            # TSV に無い行（ヒットの後に TSV が短くなった等）は前後の行にしない
            if ($n -ne $hit.LineNumber -and $n -gt $lines.Count) { continue }
            $line = if ($n -eq $hit.LineNumber) { [string]$hit.Line } else { [string]$lines[$n - 1] }
            if (!$isExcel -or $line.IndexOf('"') -lt 0) {
                # " を含まない行（ほとんど）は、タブで分けるだけで splitTsvCells と同じになる
                $cells = $line.Split("`t")
                if ($isExcel -and $line.IndexOf(${cellNewLine}) -ge 0) {
                    $cells = [string[]]($cells -replace ${cellNewLine}, $newLine)
                }
            } else {
                $cells = @(foreach ($m in $cellPattern.Matches("`t" + $line)) {
                        $cell = $m.Value.Substring(1)
                        if ($cell.StartsWith('"')) {
                            $q = $quoted.Match($cell)
                            if ($q.Success) { $cell = $q.Groups[1].Value.Replace('""', '"') + $q.Groups[2].Value }
                        }
                        $cell.Replace(${cellNewLine}, $newLine)
                    })
            }
            if ($cells.Count -gt $maxCells) { $maxCells = $cells.Count }
            if ($n -eq $hit.LineNumber) {
                $placeKey = $hit.Book + "|" + $hit.Location
                if (!$placeCache.ContainsKey($placeKey)) {
                    $described = describePlace $hit.Book $hit.Location
                    $placeCache[$placeKey] = @(($described.Place -replace '[\x00-\x1F]', " "), $described.Kind)
                }
                $place = $placeCache[$placeKey]
                $hitRow = @{
                    Style = ${resultStyleNormal}; Level = 1; HighlightFrom = 4
                    Texts = @($hit.Book, $place[0], $place[1], [string]$hit.LineNumber) + $cells
                }
                if ($source) {
                    # Excel のリンクの行き先: 該当のシートと、一致した最初のセル（図形・コメントの TSV は 1 列目のセル番地。readXlsxObjectUnits）。
                    # ヒットごとに関数を呼ばないよう、シートの書き方はシートごとに覚える
                    $location = ""
                    if ($isExcel) {
                        if (!$sheetCache.ContainsKey($placeKey)) {
                            $split = splitObjectPlace $hit.Location
                            $sheetCache[$placeKey] = @($(if ($split.Base -ne "") { toXlsxLocation $split.Base "" } else { "" }), $split.Kind)
                        }
                        $sheetRef = $sheetCache[$placeKey]
                        if ($sheetRef[0] -ne "") {
                            if ($sheetRef[1]) {
                                $location = $sheetRef[0] + $(if ($cells.Count -gt 0 -and $cells[0] -match '^\$?[A-Z]{1,3}\$?[0-9]+$') { $cells[0].Replace('$', '') } else { "A1" })
                            } else {
                                $column = 1
                                for ($i = 0; $i -lt $cells.Count; $i++) {
                                    if ($regex -and $regex.IsMatch($cells[$i])) { $column = $i + 1; break }
                                }
                                if (!$columnNames.ContainsKey($column)) { $columnNames[$column] = toColumnName $column }
                                $location = $sheetRef[0] + $columnNames[$column] + $hit.LineNumber
                            }
                        }
                    }
                    $hitRow.Styles = @{ 0 = ${resultStyleLink} }
                    $hitRow.Links = @{ 0 = @{ Target = $source; Location = $location } }
                }
            } else {
                $label = if ($n -lt $hit.LineNumber) { "前の行" } else { "後の行" }
                $contextRows += @{ Style = ${resultStyleContext}; Level = 2; Hidden = $true; Texts = @("", "", $label, [string]$n) + $cells }
            }
        }
        $hitRow.Collapsed = ($contextRows.Count -gt 0)
        $rows[$rows.Count] = $hitRow
        foreach ($row in $contextRows) { $rows[$rows.Count] = $row }
    }
    if ($groupRow) { $groupRow.Texts[2] = "$groupCount 件" }

    $rows[0] = @{ Style = ${resultStyleHeader}; Level = 0; Texts = @("ファイル", "場所", "種別", "行") + @(for ($c = 1; $c -le $maxCells; $c++) { toColumnName $c }) }
    $list = @(for ($i = 0; $i -lt $rows.Count; $i++) { $rows[$i] })
    return @{ Rows = $list; ColumnCount = 4 + $maxCells }
}

function writeResultBook {
    # xlsx の部品（getResultBookParts）を作業フォルダに書き、tar.exe で ZIP にまとめて path に置く。
    # tar.exe は引数を ANSI のコードページで受け取るため、日本語のパスは引数に渡さない。
    # 作業フォルダに cmd で移ってから、ASCII の名前で作り、Move-Item で目的の名前にする
    param (
        [System.Collections.IDictionary]$parts,
        [string]$path
    )

    $work = Join-Path (getPathParent $path) ("作成中_" + $PID)
    if (Test-Path -LiteralPath $work) {
        Remove-Item -LiteralPath $work -Recurse -Force
    }
    try {
        foreach ($name in $parts.Keys) {
            $file = Join-Path $work $name.Replace("/", "\")
            New-Item -ItemType Directory -Path (getPathParent $file) -Force | Out-Null
            # XML は UTF-8（BOM 付き。XML の決まりで BOM があってもよい）
            Set-Content -LiteralPath $file -Value $parts[$name] -Encoding UTF8 -NoNewline
        }
        # 呼び出し側の $ErrorActionPreference が Stop でも、標準エラーに書かれただけで止めず、終了コードで判断する
        $ErrorActionPreference = "Continue"
        $output = cmd.exe /d /c "cd /d `"$work`" && `"${tarExe}`" --format zip -cf result.zip `"[Content_Types].xml`" _rels xl" 2>&1
        $ErrorActionPreference = "Stop"
        if ($LASTEXITCODE -ne 0) {
            throw "検索結果のブックを作れませんでした（tar.exe の終了コード $LASTEXITCODE）: $(@($output) -join ' ')"
        }
        Move-Item -LiteralPath (Join-Path $work "result.zip") -Destination $path -Force
    } finally {
        if (Test-Path -LiteralPath $work) {
            Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function newResultBookPath {
    # 検索結果のブックのパス（時刻入り。同じ秒に作ると上書きしないよう番号を付ける）
    param (
        [string]$dir = ${resultBookDir},
        [datetime]$now = (Get-Date)
    )

    $base = "検索結果_" + $now.ToString("yyyyMMdd_HHmmss")
    $path = Join-Path $dir "$base.xlsx"
    for ($i = 2; Test-Path -LiteralPath $path; $i++) {
        $path = Join-Path $dir "${base}_$i.xlsx"
    }
    return $path
}

function removeOldResultBooks {
    # 以前の検索結果のブックを消す（制限モードの起動時）。Excel で開いたままのものは消せないため飛ばす
    param (
        [string]$dir = ${resultBookDir}
    )

    if (!(Test-Path -LiteralPath $dir -PathType Container)) {
        return
    }
    foreach ($item in @(Get-ChildItem -LiteralPath $dir -ErrorAction SilentlyContinue)) {
        try {
            Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
        } catch {
            # 開いたままのブック
        }
    }
}

function saveResultBook {
    # 検索の結果をブックにして保存し、そのパスを返す。
    #   result: searchRestricted の結果 / regex: 一致した部分に色を付ける正規表現 / info: 条件のシートの項目（@( @("項目", "値"), ... )）
    param (
        $result,
        [regex]$regex,
        [object[]]$info,
        [string]$dir = ${resultBookDir},
        [hashtable]$maps = @{}
    )

    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    $built = newResultRows $result.Hits $regex $maps
    $widths = @(40, 24, 10, 7) + @(for ($c = 5; $c -le $built.ColumnCount; $c++) { 20 })
    $sheet = getResultSheetXml $built.Rows $built.ColumnCount $widths $regex
    $parts = getResultBookParts $sheet (getResultInfoSheetXml $info) $built.Rows.Count $built.ColumnCount
    $path = newResultBookPath $dir
    writeResultBook $parts $path
    return $path
}
