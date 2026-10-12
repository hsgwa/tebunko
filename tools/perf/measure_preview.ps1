# 選択行のプレビューの時間とメモリを測る。tools\measure_perf.ps1 が別のプロセスで起動する。
# 検索してヒットを得たあと、ヒットの行を先頭から順に -Count 件、結果の表を↓キーで移るのと同じ順にプレビューする（前後 50 行ずつ。画面の上限と同じ 101 行）。
#   ・キャッシュ無し … readContentIndexContext にキャッシュを渡さない（検索に入らなかったファイル・キャッシュを捨てたあとの選択に当たる）。先に測る
#   ・キャッシュあり … 検索のキャッシュ（newTsvTextCache）を渡す（検索の直後の選択に当たる）
# 1 回ごとに ReadMs（readContentIndexContext）・BuildMs（HitRow.BuildPreview）と、作ったセルの数、プロセスのメモリを記録する。
# 検索の直後・プレビューのあと・GC のあと・検索のキャッシュを捨てたあとのワーキングセットも記録する（画面のメモリが戻るか・キャッシュの常駐の大きさを見る）。
# -Render: 画面を出さずに、プレビューの表を preview.xaml の PreviewRows に入れて Measure・Arrange する時間も測る（STA のプロセスで動かす）。
# 検索する語は -WordFile の JSON（Name・Word・Regex）で受け取る。結果は -Out の JSON に書く。
param (
    [Parameter(Mandatory = $true)][string]$Tool,
    [Parameter(Mandatory = $true)][string]$Index,
    [Parameter(Mandatory = $true)][string]$WordFile,
    [int]$Count = 20,
    [Parameter(Mandatory = $true)][string]$Out,
    [int]$SampleMs = 200,
    [switch]$Render,
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
assertTebunkoFunctions @("newTsvTextCache", "readContentIndexContext", "readContentIndexText")
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

function getMemory {
    # ワーキングセット・プライベート・マネージドヒープ・GC の回数・ハンドル・スレッドの今の値
    $snap = getProcessSnapshot
    $snap["Gc0"] = [GC]::CollectionCount(0); $snap["Gc1"] = [GC]::CollectionCount(1); $snap["Gc2"] = [GC]::CollectionCount(2)
    return $snap
}

function collectAll {
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
}

$monitor = startResourceMonitor $SampleMs
$previews = New-Object System.Collections.Generic.List[object]
$renderResult = $null
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
            # -FullRead: 場所の一覧の控えを捨て、毎回ファイルの全体を読む経路に入れる（範囲読みを入れる前の読み方に当たる）
            if ($FullRead) { $script:contentIndexPlacesMemo = $null }
            $callsBefore = $script:readTextCalls
            $watch = [System.Diagnostics.Stopwatch]::StartNew()
            $context = @(readContentIndexContext $path $hit.Book $hit.Location ([int]$hit.LineNumber) $beforeLines $afterLines $cacheArg)
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
            $columns = $table.ShownColumns
            $cells = $table.Rows.Count * $columns
            # 表を持ち続けるのは -Render のときだけ（GC のあとに残る増えに、このスクリプトが持つ表が入らないようにする）
            if ($mode -eq "無し" -and ($null -eq $wide -or $cells -gt $wide.Cells)) { $wide = @{ Cells = $cells; Table = $(if ($Render) { $table } else { $null }) } }
            $item = [ordered]@{
                N = $n; Cache = $mode; LineNumber = [int]$hit.LineNumber; FileMB = [Math]::Round((Get-Item -LiteralPath (toLongPath $path)).Length / 1MB, 2)
                ReadMs = (roundMs $readMs); ReadTextCalls = $textCalls; BuildMs = (roundMs $buildMs); ContextRows = $context.Count; Rows = $table.Rows.Count; Columns = $columns; Cells = $cells
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

    # -Render は、GC のあとの値を取る前に済ませる（表は描画に使うので、それまで捨てない。描画した分も GC のあとの値に入る）
    if ($Render) {
        setMonitorPhase $monitor "描画"
        if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) { throw "-Render は STA のプロセスで動かしてください。" }
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
        . ([WildcardPattern]::Escape((Join-Path $Tool "scripts\shared\ui\app_host.ps1").Replace('`', '``')))
        $content = loadXaml (Join-Path $Tool "scripts\tebunko\xaml\search\preview.xaml")
        $headerControl = $content.FindName("PreviewHeader")
        $rowsControl = $content.FindName("PreviewRows")
        $size = [System.Windows.Size]::new(1400, 700)
        $times = New-Object System.Collections.Generic.List[double]
        $before = getMemory
        # 1 回目は部品の初期化（テンプレートの読み込みなど）が入るので、2 回目以降と分けて見られるよう 5 回測る
        for ($i = 0; $i -lt 5; $i++) {
            $headerControl.ItemsSource = $null; $rowsControl.ItemsSource = $null
            $content.Measure($size); $content.Arrange([System.Windows.Rect]::new(0, 0, 1400, 700)); $content.UpdateLayout()
            $watch = [System.Diagnostics.Stopwatch]::StartNew()
            $headerControl.ItemsSource = $wide.Table.Columns
            $rowsControl.ItemsSource = $wide.Table.Rows
            $content.Measure($size); $content.Arrange([System.Windows.Rect]::new(0, 0, 1400, 700)); $content.UpdateLayout()
            $times.Add($watch.Elapsed.TotalMilliseconds)
        }
        $afterRender = getMemory
        $renderResult = [ordered]@{
            Cells = $wide.Cells; Rows = $wide.Table.Rows.Count; Columns = $wide.Table.ShownColumns
            RenderMs = @($times | ForEach-Object { roundMs $_ }); WorkingSetBeforeMB = $before.WorkingSetMB; WorkingSetAfterMB = $afterRender.WorkingSetMB
        }
        $headerControl.ItemsSource = $null; $rowsControl.ItemsSource = $null
        $content = $null; $headerControl = $null; $rowsControl = $null
    }

    # このスクリプトが持つものを捨ててから、GC のあとのワーキングセットを見る（残った増えは、画面が持ち続けたものとみなす。戻らなければリークの疑い）
    $table = $null; $row = $null; $context = $null; $numbers = $null; $lines = $null; $item = $null; $snap = $null; $wide = $null
    collectAll
    $afterGc = getMemory
    # 検索のキャッシュを捨てた分（キャッシュの常駐の大きさ）
    $cache.Texts.Clear(); $cache.Chars[0] = 0
    collectAll
    $afterCacheDrop = getMemory
} finally {
    $samples = stopResourceMonitor $monitor
}
$resultData = [ordered]@{
    Name = $Name; Word = $Word; Regex = $Regex; FullRead = [bool]$FullRead; Count = $hits.Count; Previews = $previews.ToArray()
    Memory = [ordered]@{
        AfterSearchWorkingSetMB = $afterSearch.WorkingSetMB; BeforePreviewManagedMB = $beforePreview.ManagedMB; AfterPreviewWorkingSetMB = $afterPreview.WorkingSetMB
        AfterGcWorkingSetMB = $afterGc.WorkingSetMB; AfterCacheDropWorkingSetMB = $afterCacheDrop.WorkingSetMB
        AfterGcManagedMB = $afterGc.ManagedMB; AfterCacheDropManagedMB = $afterCacheDrop.ManagedMB; CacheChars = $cacheChars
    }
    Render = $renderResult; Samples = $samples
    PeakWorkingSetMB = [Math]::Round([System.Diagnostics.Process]::GetCurrentProcess().PeakWorkingSet64 / 1MB, 1)
}
[System.IO.File]::WriteAllText($Out, ($resultData | ConvertTo-Json -Depth 6), (New-Object System.Text.UTF8Encoding($false)))
