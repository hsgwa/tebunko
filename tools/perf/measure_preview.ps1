# 選択行のプレビューの時間とメモリを測る。tools\measure_perf.ps1 が別のプロセスで起動する。
# 検索してヒットを得たあと、ヒットの行を先頭から順に -Count 件、結果の表を↓キーで移るのと同じ順にプレビューする（前後 50 行ずつ。画面の上限と同じ 101 行）。
#   ・キャッシュ無し … 読みにキャッシュを渡さない（検索に入らなかったファイル・キャッシュを捨てたあとの選択に当たる）。先に測る
#   ・キャッシュあり … 検索のキャッシュ（newTsvTextCache）を渡す（検索の直後の選択に当たる）
# 1 回ごとに ReadMs（画面の最初の読みと同じ呼び方）・BuildMs（HitRow.BuildPreview）・CellsMs（画面が最初に出す範囲のセルを作る時間）と、
# 表の全部のセルの数（Cells。行 × 全部の列）・表の行がそのとき持っているセルの数（CreatedCells。見える範囲の分）、プロセスのメモリを記録する。
# 前の版（全部のセルを BuildPreview が作る版）を測るときは、CreatedCells は Cells と同じ。
# 検索の直後・プレビューのあと・GC のあと・検索のキャッシュを捨てたあとのワーキングセットも記録する（画面のメモリが戻るか・キャッシュの常駐の大きさを見る）。
# -ReadMore: 上の測りのあとに、ヒットが多い行（セルの数が最大のもの）の表へ、下に 200 行ずつ 5 回読み足す時間（読む・行を作る・セルを作る）を測る。
# -Render: 画面を出さずに、プレビューの表を preview.xaml の PreviewRows に入れて Measure・Arrange する時間も測る（STA のプロセスで動かす）。
#   あわせて、できた部品の数（Elements）、下と右へスクロールして見えている範囲を作り直す時間（ScrollMs）、描画のあとに残るマネージドヒープの増え（RenderRetainedMB）を記録する。
# 検索する語は -WordFile の JSON（Name・Word・Regex）で受け取る。結果は -Out の JSON に書く。
param (
    [Parameter(Mandatory = $true)][string]$Tool,
    [Parameter(Mandatory = $true)][string]$Index,
    [Parameter(Mandatory = $true)][string]$WordFile,
    [int]$Count = 20,
    [Parameter(Mandatory = $true)][string]$Out,
    [int]$SampleMs = 200,
    [switch]$Render,
    [switch]$ReadMore,
    [switch]$FullRead
)

$ErrorActionPreference = "Stop"
. ([WildcardPattern]::Escape("$PSScriptRoot\perf_common.ps1".Replace('`', '``')))
$lib = resolveTebunkoLib $Tool
. ([WildcardPattern]::Escape($lib.Replace('`', '``')))
foreach ($typesFile in @("scripts\shared\ui\types.ps1", "scripts\tebunko\ui\types.ps1")) {
    $typesPath = Join-Path $Tool $typesFile
    if (!(Test-Path -LiteralPath $typesPath)) { throw "測る tebunko に $typesFile がありません。" }
    . ([WildcardPattern]::Escape($typesPath.Replace('`', '``')))
}
assertTebunkoFunctions @("newTsvTextCache", "readContentIndexText")
# 前後を読む関数: 今の版は readContentIndexLines、前の版は readContentIndexContext
[void](resolveTebunkoFunction @("readContentIndexLines", "readContentIndexContext"))
# 見えている列の範囲を決める判断（列の仮想化がある版だけ）。無い版（前の版）は全部のセルを BuildPreview が作る
$previewViewPath = Join-Path $Tool "scripts\tebunko\ui\preview_view.ps1"
if (Test-Path -LiteralPath $previewViewPath) { . ([WildcardPattern]::Escape($previewViewPath.Replace('`', '``'))) }
$canWindow = $null -ne (Get-Command getPreviewColumnWindow -CommandType Function -ErrorAction SilentlyContinue)
if ($ReadMore -and $null -eq (Get-Command readContentIndexLines -CommandType Function -ErrorAction SilentlyContinue)) { throw "-ReadMore は、読み足し（readContentIndexLines）がある版でだけ測れます。" }
# readContentIndexText が何回呼ばれたかを数える（-FullRead で全部読む経路を本当に通ったかを確かめるため）。製品の関数は呼び出しをそのまま渡す
$script:readTextCalls = 0
$originalReadText = (Get-Command readContentIndexText -CommandType Function).ScriptBlock
function readContentIndexText { $script:readTextCalls++; & $originalReadText @args }
$getFiles = resolveTebunkoFunction @("getContentIndexFiles", "getIndexPackFiles")
$searchIndex = resolveTebunkoFunction @("searchContentIndex", "searchPackIndex")
$spec = [System.IO.File]::ReadAllText($WordFile, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
$Name = [string]$spec.Name
$Word = [string]$spec.Word
$Regex = [bool]$spec.Regex
$pattern = $null
if ($Regex) { $pattern = [regex]::new($Word, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase) }
$beforeLines = 50
$afterLines = 50

function roundMs($value) { return [Math]::Round([double]$value, 2) }

# 画面の最初の読み（showDetail）と同じ呼び方で、ヒットの前後を読む。画面と同じく、最後に読んだファイルの控えを次の読みに渡す。
# 読み足しの関数（readContentIndexLines）が無い前の版は、前後を読む関数（readContentIndexContext）で読む
$script:readFirstMarks = $null
$canReadLines = $null -ne (Get-Command readContentIndexLines -CommandType Function -ErrorAction SilentlyContinue)
function readPreviewFirst($hit, [string]$path, $cacheArg) {
    $lineNumber = [int]$hit.LineNumber
    if (-not $canReadLines) {
        return @(readContentIndexContext $path $hit.Book $hit.Location $lineNumber $beforeLines $afterLines $cacheArg)
    }
    $first = [Math]::Max(1, $lineNumber - $beforeLines)
    $result = readContentIndexLines $path $hit.Book $hit.Location $first ($lineNumber - $first + 1 + $afterLines) $cacheArg $script:readFirstMarks
    $script:readFirstMarks = $result.Marks
    return @($result.Rows)
}

# 画面を想定した大きさ（描画の計測の枠と同じ）。見える行の数は 1 行 22 px ほどとして 30 行
$viewWidth = 1400.0
$viewRows = 30
$columnMargin = 3
$cellRowMargin = 10

function hasMember($target, [string]$name) { return $null -ne ($target | Get-Member -Name $name) }

function countMadeCells($table) {
    # 表の行が、いま持っているセルの数（見える範囲の分）。セルを後から作る版でない前の版は、全部のセルを作るので 行 × 列
    if (-not (hasMember $table "EnsureCells")) { return $table.Rows.Count * $table.Columns.Count }
    $sum = 0
    foreach ($row in $table.Rows) { $sum += $row.Cells.Count }
    return $sum
}

function setWindowFor($table, [double]$offset, [double]$width) {
    # 横の位置 offset・見える幅 width の、見えている列の範囲に決め直す。範囲が変わったら $true（前の版・範囲を決める判断が無い版は $false）
    if (-not $canWindow -or -not (hasMember $table "EnsureCells")) { return $false }
    # 画面と同じく、列の幅の累積を渡す。累積を持たない前の版は、列の幅の並びを渡す版の関数なので Widths()
    $columnArg = $(if (hasMember $table "Prefix") { $table.Prefix() } else { $table.Widths() })
    $w = getPreviewColumnWindow $columnArg $offset $width ([HitRow]::NumberWidth) $columnMargin
    return $table.SetWindow($w.First, $w.Last, $w.Left, $w.Right)
}

function showCellsFrom($table, [int]$topIndex) {
    # 行の集まりの topIndex 行目から $viewRows 行（と上下の余白の行）に、今の範囲のセルが無ければ作る
    if (-not $canWindow -or -not (hasMember $table "EnsureCells")) { return }
    $table.EnsureCells($topIndex - $cellRowMargin, $topIndex + $viewRows + $cellRowMargin)
}

function showInitialCells($table) {
    # 画面が最初に出す範囲（一致したセルが見える横の位置・選択行の少し上から）のセルを作る
    if (-not $canWindow -or -not (hasMember $table "EnsureCells")) { return }
    [void](setWindowFor $table ([Math]::Max(0, $table.HitOffset - 120)) $viewWidth)
    $hitIndex = 0
    for ($i = 0; $i -lt $table.Rows.Count; $i++) { if ($table.Rows[$i].IsHitRow) { $hitIndex = $i; break } }
    showCellsFrom $table ([Math]::Max(0, $hitIndex - 2))
}

function getMemory {
    # ワーキングセット・プライベート・マネージドヒープ・GC の回数・ハンドル・スレッドの今の値
    $snap = getProcessSnapshot
    $snap["Gc0"] = [GC]::CollectionCount(0); $snap["Gc1"] = [GC]::CollectionCount(1); $snap["Gc2"] = [GC]::CollectionCount(2)
    return $snap
}

function countVisuals($element) {
    # 部品の木にある部品の数（仮想化で、見えている分だけ部品ができているかを見る）
    $count = 1
    $children = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($element)
    for ($i = 0; $i -lt $children; $i++) { $count += countVisuals ([System.Windows.Media.VisualTreeHelper]::GetChild($element, $i)) }
    return $count
}

function collectAll {
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
}

$monitor = startResourceMonitor $SampleMs
$previews = New-Object System.Collections.Generic.List[object]
$renderResult = $null
$readMoreRuns = New-Object System.Collections.Generic.List[object]
try {
    setMonitorPhase $monitor "検索"
    $found = & $getFiles @($Index)
    $filesName = @("ContentIndexFiles", "Packs") | Where-Object { $found.ContainsKey($_) } | Select-Object -First 1
    if (!$filesName) { throw "取得した本文インデックスの一覧に項目がありません" }
    $files = $found[$filesName]
    $cache = newTsvTextCache
    $result = & $searchIndex $Word $files (!$Regex) 10000 -cache $cache
    $hits = @($result.Hits | Select-Object -First $Count)
    if ($hits.Count -eq 0) { throw "ヒットが無いため、プレビューを測れません（語: $Name）。" }
    $afterSearch = getMemory
    $cacheChars = [long]$cache.Chars[0]

    # プレビューの前に GC して、あとで残った分（プレビューが持ち続けたもの）を比べる基準にする
    collectAll
    $beforePreview = getMemory
    setMonitorPhase $monitor "プレビュー"
    $wide = $null
    $n = 0
    foreach ($mode in @("無し", "あり")) {
        $cacheArg = if ($mode -eq "あり") { $cache } else { $null }
        foreach ($hit in $hits) {
            $n++
            $path = [System.IO.Path]::Combine($hit.Root, $hit.RelPath)
            # -FullRead: 控えを捨て、毎回ファイルの全体を読む経路に入れる（初めて開くファイルの最初の読みに当たる）
            # （contentIndexPlacesMemo は、前の版が場所の一覧を覚えていた変数。今の版には無いが、前の版を測るときに捨てるため残す）
            if ($FullRead) { $script:readFirstMarks = $null; $script:contentIndexPlacesMemo = $null }
            $callsBefore = $script:readTextCalls
            $watch = [System.Diagnostics.Stopwatch]::StartNew()
            $context = @(readPreviewFirst $hit $path $cacheArg)
            $readMs = $watch.Elapsed.TotalMilliseconds
            $textCalls = $script:readTextCalls - $callsBefore
            $relDir = [string]$hit.RelDir
            $cut = $relDir.IndexOf("\")
            $indexName = if ($cut -lt 0) { $relDir } else { $relDir.Substring(0, $cut) }
            $row = [HitRow]::Create($indexName, $hit.Root, $hit.RelPath, $relDir, $hit.FileName, $hit.Book, $hit.Location, [int]$hit.LineNumber, $hit.Line, $Word, $pattern)
            $numbers = [int[]]@($context | ForEach-Object { $_.LineNumber })
            $lines = [string[]]@($context | ForEach-Object { $_.Line })
            $watch.Restart()
            $table = $row.BuildPreview($numbers, $lines)
            $buildMs = $watch.Elapsed.TotalMilliseconds
            $watch.Restart()
            showInitialCells $table
            $cellsMs = $watch.Elapsed.TotalMilliseconds
            $columns = $table.Columns.Count   # 前の版は Columns が表示した列だけ。今の版は全部の列
            $cells = $table.Rows.Count * $columns
            $created = countMadeCells $table
            # 表を持ち続けるのは -Render のときだけ（GC のあとに残る増えに、このスクリプトが持つ表が入らないようにする）
            if ($mode -eq "無し" -and ($null -eq $wide -or $cells -gt $wide.Cells)) { $wide = @{ Cells = $cells; Table = $(if ($Render) { $table } else { $null }); Hit = $hit } }
            $item = [ordered]@{
                N = $n; Cache = $mode; LineNumber = [int]$hit.LineNumber; FileMB = [Math]::Round((Get-Item -LiteralPath (toLongPath $path)).Length / 1MB, 2)
                ReadMs = (roundMs $readMs); ReadTextCalls = $textCalls; BuildMs = (roundMs $buildMs); ContextRows = $context.Count; Rows = $table.Rows.Count; Columns = $columns; Cells = $cells; CreatedCells = $created; CellsMs = (roundMs $cellsMs)
            }
            $snap = getMemory
            foreach ($key in $snap.Keys) { $item[$key] = $snap[$key] }
            $previews.Add($item)
        }
    }
    $afterPreview = getMemory
    if ($FullRead) {
        # 切り替えが効かず、範囲読みのまま測ったときは、前後が同じ値で黙って並ぶので、失敗にする（時間の大小では判定しない）
        $fullReadProblems = @(getFullReadPathProblems $previews.ToArray())
        if ($fullReadProblems.Count -gt 0) { throw ("-FullRead で全部読む経路を通っていません: " + ($fullReadProblems -join " / ")) }
    }

    # このスクリプトが持つものを捨ててから、GC のあとのワーキングセットを見る（残った増えは、画面が持ち続けたものとみなす。戻らなければリークの疑い）
    $table = $null; $row = $null; $context = $null; $numbers = $null; $lines = $null; $item = $null; $snap = $null
    collectAll
    $afterGc = getMemory
    # 検索のキャッシュを捨てた分（キャッシュの常駐の大きさ）
    $cache.Texts.Clear(); $cache.Chars[0] = 0
    collectAll
    $afterCacheDrop = getMemory

    # -ReadMore も、残った増えを取ったあとに行う。セルの数が最大だった行の表を作り直し、下に 200 行ずつ 5 回読み足す
    if ($ReadMore) {
        setMonitorPhase $monitor "読み足し"
        $hit = $wide.Hit
        $path = [System.IO.Path]::Combine($hit.Root, $hit.RelPath)
        $relDir = [string]$hit.RelDir
        $cut = $relDir.IndexOf("\")
        $indexName = if ($cut -lt 0) { $relDir } else { $relDir.Substring(0, $cut) }
        $moreRow = [HitRow]::Create($indexName, $hit.Root, $hit.RelPath, $relDir, $hit.FileName, $hit.Book, $hit.Location, [int]$hit.LineNumber, $hit.Line, $Word, $pattern)
        # 初めて開くファイルとして最初の読みをし、その控えで読み足す（画面と同じ）
        $script:readFirstMarks = $null
        $context = @(readPreviewFirst $hit $path $null)
        $moreTable = $moreRow.BuildPreview([int[]]@($context | ForEach-Object { $_.LineNumber }), [string[]]@($context | ForEach-Object { $_.Line }))
        showInitialCells $moreTable
        $marks = $script:readFirstMarks
        $nextLine = $moreTable.LastLine() + 1
        for ($i = 0; $i -lt 5; $i++) {
            $watch = [System.Diagnostics.Stopwatch]::StartNew()
            $more = readContentIndexLines $path $hit.Book $hit.Location $nextLine 200 $null $marks
            $moreReadMs = $watch.Elapsed.TotalMilliseconds
            $marks = $more.Marks
            $watch.Restart()
            $numbers = [int[]]@($more.Rows | ForEach-Object { $_.LineNumber })
            $moreRows = $moreRow.BuildPreviewRows($moreTable, $numbers, [string[]]@($more.Rows | ForEach-Object { $_.Line }))
            $moreTable.InsertRows("down", $moreRows)
            $moreBuildMs = $watch.Elapsed.TotalMilliseconds
            $watch.Restart()
            # 列が増えていれば範囲を決め直し、読み足した行を含む見える範囲（下端）のセルを作る
            [void](setWindowFor $moreTable ([Math]::Max(0, $moreTable.HitOffset - 120)) $viewWidth)
            showCellsFrom $moreTable ([Math]::Max(0, $moreTable.Rows.Count - $viewRows))
            $moreCellsMs = $watch.Elapsed.TotalMilliseconds
            $readMoreRuns.Add([ordered]@{
                N = $i + 1; ReadMs = (roundMs $moreReadMs); BuildMs = (roundMs $moreBuildMs); CellsMs = (roundMs $moreCellsMs)
                Rows = $numbers.Count; TableRows = $moreTable.Rows.Count; Columns = $moreTable.Columns.Count
                CreatedCells = (countMadeCells $moreTable)
            })
            if ($numbers.Count -eq 0 -or $more.End) { break }
            $nextLine = $numbers[-1] + 1
        }
        $moreTable = $null; $moreRow = $null; $moreRows = $null; $more = $null; $context = $null
    }

    # -Render は、残った増え（GC のあとの値）を取ったあとに行う（描画で増えた分を、プレビューが持ち続けるものに入れない）。描画に使う表（$wide）だけは、それまで持っておく
    if ($Render) {
        setMonitorPhase $monitor "描画"
        if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) { throw "-Render は STA のプロセスで動かしてください。" }
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
        . ([WildcardPattern]::Escape((Join-Path $Tool "scripts\shared\ui\app_host.ps1").Replace('`', '``')))
        $content = loadXaml (Join-Path $Tool "scripts\tebunko\xaml\search\preview.xaml")
        $headerControl = $content.FindName("PreviewHeader")
        $headerPanel = $content.FindName("PreviewHeaderPanel")   # 前の版には無い
        $rowsControl = $content.FindName("PreviewRows")
        $size = [System.Windows.Size]::new($viewWidth, 700)
        $rect = [System.Windows.Rect]::new(0, 0, $viewWidth, 700)
        $table = $wide.Table
        $windowed = $canWindow -and (hasMember $table "EnsureCells")
        $times = New-Object System.Collections.Generic.List[double]
        collectAll
        $before = getMemory
        # 1 回目は部品の初期化（テンプレートの読み込みなど）が入るので、2 回目以降と分けて見られるよう 5 回測る
        for ($i = 0; $i -lt 5; $i++) {
            $headerControl.ItemsSource = $null; $rowsControl.ItemsSource = $null
            if ($null -ne $headerPanel) { $headerPanel.DataContext = $null }
            $content.Measure($size); $content.Arrange($rect); $content.UpdateLayout()
            $watch = [System.Diagnostics.Stopwatch]::StartNew()
            if ($windowed) {
                # 画面（preview.ps1）と同じ順: 行と空きを入れ、見える範囲の列とセルを決めてから配置する
                [void]$table.SetWindow(0, -1, 0, 0)   # 毎回、セルを作るところから測る
                $headerPanel.DataContext = $table.Window
                $rowsControl.ItemsSource = $table.Rows
                $content.Measure($size); $content.Arrange($rect); $content.UpdateLayout()
                [void](setWindowFor $table ([Math]::Max(0, $table.HitOffset - 120)) $viewWidth)
                $headerControl.ItemsSource = $table.WindowColumns()
                showCellsFrom $table 0
            } else {
                $headerControl.ItemsSource = $table.Columns
                $rowsControl.ItemsSource = $table.Rows
            }
            $content.Measure($size); $content.Arrange($rect); $content.UpdateLayout()
            $times.Add($watch.Elapsed.TotalMilliseconds)
        }
        $elements = countVisuals $content
        # 下と右へスクロールし、見えている範囲を作り直す時間（最後の行・最後の列まで 4 段階で動かして、合計を測る）
        $scrollMs = $null
        $scroll = $rowsControl.Template.FindName("PreviewScroll", $rowsControl)
        if ($null -ne $scroll) {
            $watch = [System.Diagnostics.Stopwatch]::StartNew()
            foreach ($step in 1..4) {
                $scroll.ScrollToVerticalOffset([Math]::Min($scroll.ScrollableHeight, $step * $scroll.ScrollableHeight / 4))
                $scroll.ScrollToHorizontalOffset($step * $scroll.ScrollableWidth / 4)
                $content.UpdateLayout()
                if ($windowed) {
                    if (setWindowFor $table $scroll.HorizontalOffset $scroll.ViewportWidth) { $headerControl.ItemsSource = $table.WindowColumns() }
                    showCellsFrom $table ([int]$scroll.VerticalOffset)
                    $content.UpdateLayout()
                }
            }
            $scrollMs = $watch.Elapsed.TotalMilliseconds
        }
        collectAll
        $afterRender = getMemory
        $renderResult = [ordered]@{
            Cells = $wide.Cells; Rows = $table.Rows.Count; Columns = $table.Columns.Count
            RenderMs = @($times | ForEach-Object { roundMs $_ }); WorkingSetBeforeMB = $before.WorkingSetMB; WorkingSetAfterMB = $afterRender.WorkingSetMB
            Elements = $elements; ScrollMs = $(if ($null -ne $scrollMs) { roundMs $scrollMs } else { $null })
            RenderRetainedMB = [Math]::Round([double]$afterRender.ManagedMB - [double]$before.ManagedMB, 1)
        }
        $headerControl.ItemsSource = $null; $rowsControl.ItemsSource = $null
        $content = $null; $headerControl = $null; $headerPanel = $null; $rowsControl = $null; $scroll = $null; $table = $null
    }
} finally {
    $samples = stopResourceMonitor $monitor
}
$resultData = [ordered]@{
    Name = $Name; Word = $Word; Regex = $Regex; FullRead = [bool]$FullRead; Count = $hits.Count; Previews = $previews.ToArray(); ReadMore = $(if ($ReadMore) { $readMoreRuns.ToArray() } else { $null })
    Memory = [ordered]@{
        AfterSearchWorkingSetMB = $afterSearch.WorkingSetMB; BeforePreviewManagedMB = $beforePreview.ManagedMB; AfterPreviewWorkingSetMB = $afterPreview.WorkingSetMB
        AfterGcWorkingSetMB = $afterGc.WorkingSetMB; AfterCacheDropWorkingSetMB = $afterCacheDrop.WorkingSetMB
        AfterGcManagedMB = $afterGc.ManagedMB; AfterCacheDropManagedMB = $afterCacheDrop.ManagedMB; CacheChars = $cacheChars
    }
    Render = $renderResult; Samples = $samples
    PeakWorkingSetMB = [Math]::Round([System.Diagnostics.Process]::GetCurrentProcess().PeakWorkingSet64 / 1MB, 1)
}
[System.IO.File]::WriteAllText($Out, ($resultData | ConvertTo-Json -Depth 6), (New-Object System.Text.UTF8Encoding($false)))
