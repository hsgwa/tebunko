# ［2 検索］タブ（検索条件・検索の実行・結果の一覧と絞り込み）。

# ============================================================================
# ［2 検索］
# ============================================================================

$script:hitRows = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$ui.ResultGrid.ItemsSource = $script:hitRows
$script:hitView = [System.Windows.Data.CollectionViewSource]::GetDefaultView($script:hitRows)
# 表示用（強調セグメント・DisplayLine・セル列）は、行が画面に出るときだけ作る（件数が多くても軽い）。
# HitRow.Prepare は1回だけ実行し、作った値は PropertyChanged で反映する
$ui.ResultGrid.Add_LoadingRow({
    param ($s, $e)
    if ($e.Row.Item -is [HitRow]) { $e.Row.Item.Prepare() }
})
# ファイルごとにまとめるときの見出し（元のファイルのフルパス → FileGroup）。検索のたびに作り直す
$script:fileGroups = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
# 表が初めて表示されたときに、まとめるかどうかを反映する（applyGrouping。見出しは表が表示された後でないと付けられない）
$ui.ResultGrid.Add_Loaded({ safe { applyGrouping } })
$script:search = $null
$script:lastSearch = $null
$script:sourceFolderMaps = @{}  # インデックスのフォルダ → インデックス名と変換対象フォルダの対応（getSourceLocation のキャッシュ）
$script:filterText = ""
# 検索で読んだ TSV の内容（画面を閉じるまで残し、次の検索では更新の無い TSV をファイルから読まない）
$script:tsvCache = newTsvTextCache

# 別スレッドで実行する検索（結果は $shared.Queue に少しずつ入れる）
${searchScript} = {
    param ($libPath, $word, $simpleMatch, $folders, $limit, $shared, $cache)
    try {
        . $libPath
        # TSV が多いと数え上げだけで数秒かかるため、途中の件数を画面に伝える（止まって見えないように）
        $index = getIndexTsvFiles $folders { param ($count) $shared.Scanned = $count }
        $shared.Folders = $index.Folders
        $shared.Total = $index.Files.Count
        $shared.IndexTotal = $index.Files.Count
        # 検索条件（大文字・小文字の区別・対象ファイル）は startSearch が $shared に入れる
        $result = searchIndex $word $index.Files $simpleMatch $limit 50 -caseSensitive $shared.CaseSensitive -fileFilter $shared.FileFilter -cache $cache -onProgress {
            param ($done, $total, $newHits)
            foreach ($hit in $newHits) {
                $shared.Queue.Enqueue($hit)
            }
            $shared.Total = $total
            $shared.Done = $done
        } -shouldStop { $shared.Stop }
        $shared.Total = $result.Total
        $shared.Truncated = $result.Truncated
        $shared.Cancelled = $result.Cancelled
    } catch {
        $shared.Error = $_.Exception.Message
    } finally {
        $shared.Finished = $true
    }
}

# 検索ワード（前後の空白を除く）
function getWordText {
    return $ui.WordBox.Text.Trim()
}

function getSearchOptionFromUi {
    # 画面の検索条件を readSearchOption と同じ形で返す
    return @{
        UseRegex      = [bool]$ui.RegexCheck.IsChecked
        CaseSensitive = [bool]$ui.CaseCheck.IsChecked
        FileFilter    = $ui.FileFilterBox.Text.Trim()
        GroupByFile   = [bool]$ui.GroupByFileCheck.IsChecked
    }
}

function setSearchOptionToUi {
    param (
        [hashtable]$option
    )

    $ui.RegexCheck.IsChecked = [bool]$option.UseRegex
    $ui.CaseCheck.IsChecked = [bool]$option.CaseSensitive
    $ui.FileFilterBox.Text = [string]$option.FileFilter
    $ui.GroupByFileCheck.IsChecked = [bool]$option.GroupByFile
    applyGrouping
}

function applyGrouping {
    # ［ファイルごとにまとめる］の状態に合わせて、結果の表をファイルごとのグループに分ける・戻す。
    # 行（HitRow）はそのままなので、選択・プレビュー・元のファイルを開く・絞り込み・出力はどちらでも同じに動く。
    # 見出し（GroupStyle。tab_search.xaml の FileGroupStyle）はまとめている間だけ付ける。GroupStyle のある表は
    # ・空のまま初めて表示されると、あとで行を入れても列の幅が決まらず、列見出しとセルが空になる
    # ・縦のスクロールバーが出ても「該当行」（幅 *）の幅を計算し直さず、横にはみ出す
    # という WPF の不具合があるため。前者を避けるため、表が表示される前（起動時）は付けず、Loaded で付ける
    $grid = $ui.ResultGrid
    $groups = $script:hitView.GroupDescriptions
    if ([bool]$ui.GroupByFileCheck.IsChecked) {
        if (!$grid.IsLoaded) {
            return
        }
        if ($grid.GroupStyle.Count -eq 0) {
            $grid.GroupStyle.Add($grid.Resources["FileGroupStyle"])
        }
        if ($groups.Count -eq 0) {
            # 行ごとの表示の間に入れた行には見出しが無いため、先に付ける
            assignFileGroups $script:hitRows
            $groups.Add((New-Object System.Windows.Data.PropertyGroupDescription "FileGroup"))
        }
    } else {
        $groups.Clear()
        $grid.GroupStyle.Clear()
    }
    refreshStarColumns
}

function refreshStarColumns {
    # 幅 * の列（該当行）の幅を、今の表の幅で計算し直す（applyGrouping の後者の不具合への対処）。
    # まとめ表示に切り替えたとき・戻したとき、まとめ表示で行の数が変わったとき（検索の終わり・絞り込み）に呼ぶ。
    # 行の並びが決まった後に計算させるため、画面の処理が済んでから行う
    [void]$ui.ResultGrid.Dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, [Action]{
        safe {
            foreach ($column in $ui.ResultGrid.Columns) {
                if ($column.Width.IsStar) {
                    $width = $column.Width
                    $column.Width = [System.Windows.Controls.DataGridLength]::Auto
                    $column.Width = $width
                }
            }
        }
    })
}

function isGroupedByFile {
    # 結果の表をファイルごとにまとめて表示しているか
    return $ui.ResultGrid.GroupStyle.Count -gt 0
}

function assignFileGroups {
    # 見出しの無い行に、元のファイルの見出し（FileGroup）を付ける（pumpSearch の中の処理と同じ）
    param (
        $rows
    )

    foreach ($row in $rows) {
        if ($null -ne $row.FileGroup) {
            continue
        }
        $fileKey = "$($row.Root)\$($row.RelDir)\$($row.Book)"
        $group = $null
        if (!$script:fileGroups.TryGetValue($fileKey, [ref]$group)) {
            $group = newFileGroup $fileKey $row.RelDir $row.Book
        }
        if ($group.AddLocation($row.Location)) {
            addFileGroupLocation $group $row.Book $row.Location
        }
        $row.FileGroup = $group
    }
}

function newFileGroup {
    # 元のファイルの見出し（FileGroup）を作って覚える（key は元のファイルのフルパス）
    param (
        [string]$key,
        [string]$relDir,
        [string]$book
    )

    $group = [FileGroup]::new()
    $group.Book = $book
    $group.RelDir = $relDir
    $group.FullPath = $key
    $group.AppKind = getAppKind $book
    $script:fileGroups[$key] = $group
    return $group
}

function addFileGroupLocation {
    # 見出しに、そのファイルで初めてヒットした場所を足し、右端の表記を作り直す
    param (
        [FileGroup]$group,
        [string]$book,
        [string]$location
    )

    $group.AddLabel((formatLocationLabel $book $location))
    $group.SetLocationText((describeFileLocations $group.GetLocations()))
}

function updateWordNotice {
    $notice = getWordNotice (getWordText) ([bool]$ui.RegexCheck.IsChecked)
    if ($notice -ne "") {
        $ui.WordNotice.Text = $notice
        $ui.WordNotice.Visibility = "Visible"
    } else {
        $ui.WordNotice.Visibility = "Collapsed"
    }
    updateSearchButton
}

function updateSearchButton {
    $noIndex = $script:indexSummary -and $script:indexSummary["Count"] -eq 0
    $state = newSearchButtonState ([bool]$script:search) ([bool]($script:search -and $script:search.Shared.Stop)) `
        (getWordText) (!$noIndex) @(getSearchTargets).Count
    $ui.SearchButton.Content = $state.Content
    $ui.SearchButton.IsEnabled = $state.Enabled
}

function updateSearchTarget {
    $targets = @(getSearchTargets)
    $summary = $script:indexSummary
    if ($summary -and $summary["Count"] -eq 0) {
        $ui.SearchTargetText.Text = "検索対象：なし（インデックスがありません。先に［1 インデックス管理］で作成してください）"
    } elseif ($targets.Count -eq 0) {
        $ui.SearchTargetText.Text = "検索対象：なし（左の一覧で、検索するインデックス・フォルダにチェックを付けてください）"
    } elseif (!(isAllIndexChecked)) {
        $ui.SearchTargetText.Text = "検索対象：$(describeSearchTargets $targets)"
    } elseif ($null -eq $summary) {
        $ui.SearchTargetText.Text = "検索対象：すべて（確認中…）"
    } else {
        $ui.SearchTargetText.Text = "検索対象：すべて（TSV $($summary['Count'].ToString('N0')) 件 ・ 最終変換 $(formatTime $summary['LastWrite'])）"
    }
    $ui.SearchTargetText.ToolTip = $ui.SearchTargetText.Text
    $ui.GoIndexTabButton.Visibility = if ($summary -and $summary["Count"] -eq 0) { "Visible" } else { "Collapsed" }
    updateSearchButton
}

function startSearch {
    if ($script:search) {
        cancelSearch
        return
    }
    $word = getWordText
    if ($word -eq "") {
        setStatus "検索ワードを入力してください。"
        return
    }

    $option = getSearchOptionFromUi
    writeSearchOption $option
    $useRegex = $option.UseRegex
    # 一致箇所の強調にも、検索と同じ正規表現を使う
    $searchRegex = newSearchRegex $word (!$useRegex) $option.CaseSensitive
    $simpleMatch = $searchRegex.SimpleMatch
    $pattern = $searchRegex.Regex

    $ui.FilterBox.Text = ""
    $script:filterText = ""
    $script:hitView.Filter = $null
    $script:hitRows.Clear()
    $script:fileGroups.Clear()
    clearDetail
    # 検索対象ツリーでチェックしたフォルダだけを検索する（結果の相対パスは、インデックスのフォルダからのまま）
    $folders = @(getSearchTargets)
    if ($folders.Count -eq 0) {
        setStatus "検索するフォルダに、左の「検索対象」でチェックを付けてください。"
        return
    }
    $ui.IndexColumn.Visibility = if (@($folders | ForEach-Object { (splitIndexRelPath ([string]$_.RelPath)).Name } | Sort-Object -Unique).Count -gt 1) { "Visible" } else { "Collapsed" }

    $shared = [hashtable]::Synchronized(@{
        Queue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
        Stop = $false; Finished = $false; Done = 0; Total = -1; IndexTotal = -1; Folders = $null; Scanned = 0
        Truncated = $false; Cancelled = $false; Error = $null
        CaseSensitive = $option.CaseSensitive; FileFilter = $option.FileFilter
    })
    $ps = [powershell]::Create()
    [void]$ps.AddScript(${searchScript}.ToString())
    foreach ($argument in @(${libPath}, $word, $simpleMatch, $folders, ${searchLimit}, $shared, $script:tsvCache)) {
        [void]$ps.AddArgument($argument)
    }
    $script:search = @{
        PS = $ps; Handle = $ps.BeginInvoke(); Shared = $shared
        Word = $word; Pattern = $pattern; SimpleMatch = $simpleMatch; UseRegex = $useRegex; Option = $option; Start = Get-Date
    }

    $ui.SearchProgress.Visibility = "Visible"
    $ui.SearchProgress.IsIndeterminate = $true
    $ui.SummaryText.Text = "検索中…"
    $taskbar.ProgressState = "Indeterminate"
    setStatus "検索しています：${word}"
    updateSearchButton
    $script:searchTimer.Start()
}

function cancelSearch {
    if ($script:search) {
        $script:search.Shared.Stop = $true
        $ui.SummaryText.Text = "中止しています…"
        updateSearchButton
    }
}

function pumpSearch {
    # 検索スレッドの結果を表に移し、進み具合を表示する
    $s = $script:search
    if ($null -eq $s) {
        $script:searchTimer.Stop()
        return
    }
    $shared = $s.Shared
    $hit = $null
    # 1 回に移す量は件数ではなく時間で区切る（件数で区切ると、ヒットが多いときに画面が 0.5 秒以上止まる）
    $elapsed = [System.Diagnostics.Stopwatch]::StartNew()
    $grouped = isGroupedByFile
    while ($elapsed.ElapsedMilliseconds -lt ${searchPumpMilliseconds} -and $shared.Queue.TryDequeue([ref]$hit)) {
        # インデックスのフォルダ（work\index）からの相対パスの先頭がインデックス名（splitIndexRelPath と同じ。1 件ごとの関数呼び出しを省く）
        $relDir = [string]$hit.RelDir
        $cut = $relDir.IndexOf("\")
        $indexName = if ($cut -lt 0) { $relDir } else { $relDir.Substring(0, $cut) }
        $row = [HitRow]::Create($indexName, $hit.Root, $hit.RelPath, $relDir, $hit.FileName,
                $hit.Book, $hit.Location, [int]$hit.LineNumber, $hit.Line, $s.Word, $s.Pattern)
        if ($grouped) {
            # まとめているときは、表に入れる前に見出しを決める（入れた時点の値でグループに振り分けるため）。
            # 行ごとの表示のときは作らない（切り替えたときに assignFileGroups で付ける）。
            # ヒットごとに通る処理なので assignFileGroups を呼ばずに同じことを書く（関数を呼ぶと、そのぶん検索が遅くなる）
            $fileKey = "$($hit.Root)\$relDir\$($hit.Book)"
            $group = $null
            if (!$script:fileGroups.TryGetValue($fileKey, [ref]$group)) {
                $group = newFileGroup $fileKey $relDir $hit.Book
            }
            if ($group.AddLocation($hit.Location)) {
                addFileGroupLocation $group $hit.Book $hit.Location
            }
            $row.FileGroup = $group
        }
        $script:hitRows.Add($row)
    }

    if ($shared.Total -gt 0) {
        $ratio = $shared.Done / $shared.Total
        $ui.SearchProgress.IsIndeterminate = $false
        $ui.SearchProgress.Value = $ratio
        $taskbar.ProgressState = "Normal"
        $taskbar.ProgressValue = $ratio
        if (!$shared.Stop) {
            $ui.SummaryText.Text = "検索中… $($shared.Done.ToString('N0')) / $($shared.Total.ToString('N0')) ファイル（$($script:hitRows.Count.ToString('N0')) 件）"
        }
    } elseif ($shared.Total -lt 0 -and !$shared.Stop) {
        # 数え上げの途中。件数が増えていくのが見えれば、止まっていないことが分かる
        $scanned = [int]$shared.Scanned
        $ui.SummaryText.Text = if ($scanned -gt 0) {
            "検索対象のファイルを確認しています…（$($scanned.ToString('N0')) 件）"
        } else {
            "検索対象のファイルを確認しています…"
        }
    }

    if ($shared.Finished -and $shared.Queue.IsEmpty) {
        finishSearch
    }
}

function finishSearch {
    $s = $script:search
    $script:search = $null
    $script:searchTimer.Stop()
    try {
        [void]$s.PS.EndInvoke($s.Handle)
    } catch {
    }
    $s.PS.Dispose()

    $shared = $s.Shared
    $ui.SearchProgress.Visibility = "Collapsed"
    $taskbar.ProgressState = if (isConverting) { $taskbar.ProgressState } else { "None" }
    $script:lastSearch = $s
    $seconds = ((Get-Date) - $s.Start).TotalSeconds
    updateSearchButton
    if (isGroupedByFile) {
        refreshStarColumns
    }

    if ($shared.Error) {
        $ui.SummaryText.Text = "検索できませんでした"
        setStatus "検索できませんでした：$($shared.Error)"
        return
    }

    $count = $script:hitRows.Count
    $files = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $script:hitRows) {
        [void]$files.Add("$($row.Root)\$($row.RelDir)\$($row.Book)")
    }

    if ($shared.IndexTotal -gt 0 -and $shared.Total -eq 0) {
        $ui.SummaryText.Text = "対象ファイル（$($s.Option.FileFilter)）に一致するファイルがありません。"
    } elseif ($shared.Total -eq 0) {
        $ui.SummaryText.Text = "検索対象の TSV がありません。先にインデックスを作成してください。"
    } elseif ($count -eq 0) {
        $text = "見つかりませんでした。"
        if (!$s.UseRegex -and $s.Word -match '[\\()\[\]{}.*+?^$|]') {
            $text += "（正規表現として探す場合は［正規表現を使う］をオンにしてください）"
        } elseif (describeSearchOption $s.Option) {
            $text += "（検索条件：$(describeSearchOption $s.Option)）"
        }
        $ui.SummaryText.Text = $text
    } else {
        $ui.SummaryText.Text = "$($count.ToString('N0')) 件（$($files.Count.ToString('N0')) ファイル） ・ $($seconds.ToString('0.0')) 秒"
    }

    $status = "検索しました（$($s.Word)：$($count.ToString('N0')) 件）"
    if (describeSearchOption $s.Option) {
        $status = "検索しました（$($s.Word)：$($count.ToString('N0')) 件　条件：$(describeSearchOption $s.Option)）"
    }
    if ($shared.Truncated) {
        # パスの順に検索して打ち切るため、この先のファイルのヒットは結果に出ない。そのことが分かる文面にする
        $status = "$(${searchLimit}.ToString('N0')) 件を超えたため、ここで打ち切りました。この先のファイルは検索していないため、ワード・対象ファイル・検索対象で絞り込んでください。"
    } elseif ($shared.Cancelled) {
        $status = "中止しました（$($count.ToString('N0')) 件まで表示）"
    }
    $missing = @($shared.Folders | Where-Object { !$_.Exists } | ForEach-Object { $_.Path })
    if ($missing.Count -gt 0) {
        $status += "　見つからない検索対象フォルダ：$($missing -join '、')"
    }
    if (isConverting) {
        $status += "　変換中のため、作成途中のインデックスを検索しています。"
    }
    setStatus $status
}

# 検索結果を表に移す 1 回あたりの時間（ミリ秒）。タイマーの間隔（100 ミリ秒）より短くし、その間も画面が操作できるようにする
${searchPumpMilliseconds} = 60
$script:searchTimer = newTimer 100 { safe { pumpSearch } }

# ---- 絞り込み・選択行の詳細 ----

function applyFilter {
    $script:filterText = $ui.FilterBox.Text.Trim()
    if ($script:filterText -eq "") {
        $script:hitView.Filter = $null
    } else {
        $script:hitView.Filter = [Predicate[object]] { param ($row) $row.Contains($script:filterText) }
    }
    if (isGroupedByFile) {
        refreshStarColumns
    }
    if ($script:lastSearch -and !$script:search -and $script:hitRows.Count -gt 0) {
        $shown = 0
        foreach ($row in $script:hitView) {
            $shown++
        }
        if ($script:filterText -eq "") {
            finishSummaryText
        } else {
            $ui.SummaryText.Text = "$($script:hitRows.Count.ToString('N0')) 件中 $($shown.ToString('N0')) 件を表示"
        }
    }
}

function finishSummaryText {
    $files = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $script:hitRows) {
        [void]$files.Add("$($row.Root)\$($row.RelDir)\$($row.Book)")
    }
    $ui.SummaryText.Text = "$($script:hitRows.Count.ToString('N0')) 件（$($files.Count.ToString('N0')) ファイル）"
}

$script:filterTimer = newTimer 300 {
    $script:filterTimer.Stop()
    safe { applyFilter }
}

# ---- イベント ----

$ui.WordBox.Add_TextChanged({ safe { updateWordNotice } })
$ui.WordBox.Add_PreviewKeyDown({
    param ($sender, $e)
    if ($e.Key -eq "Return") {
        safe {
            if (!$script:search) {
                startSearch
            }
        }
        $e.Handled = $true
    }
})
$ui.SearchButton.Add_Click({ safe { startSearch } })
$ui.RegexCheck.Add_Click({
    safe {
        writeSearchOption @{ UseRegex = [bool]$ui.RegexCheck.IsChecked }
        updateWordNotice
    }
})
$ui.CaseCheck.Add_Click({ safe { writeSearchOption @{ CaseSensitive = [bool]$ui.CaseCheck.IsChecked } } })
$ui.GroupByFileCheck.Add_Click({
    safe {
        writeSearchOption @{ GroupByFile = [bool]$ui.GroupByFileCheck.IsChecked }
        applyGrouping
    }
})
$ui.FileFilterBox.Add_TextChanged({
    $ui.FileFilterPlaceholder.Visibility = if ($ui.FileFilterBox.Text -eq "") { "Visible" } else { "Collapsed" }
})
# 対象ファイルは入力を終えたとき（フォーカスが外れたとき・検索したとき）に保存する
$ui.FileFilterBox.Add_LostFocus({ safe { writeSearchOption @{ FileFilter = $ui.FileFilterBox.Text.Trim() } } })
$ui.FileFilterBox.Add_KeyDown({
    param ($sender, $e)
    if ($e.Key -eq "Return") {
        safe {
            if (!$script:search) {
                startSearch
            }
        }
        $e.Handled = $true
    }
})

$ui.GoIndexTabButton.Add_Click({ $ui.Tabs.SelectedItem = $ui.IndexTab })
$ui.FilterBox.Add_TextChanged({
    $ui.FilterPlaceholder.Visibility = if ($ui.FilterBox.Text -eq "") { "Visible" } else { "Collapsed" }
    $script:filterTimer.Stop()
    $script:filterTimer.Start()
})
