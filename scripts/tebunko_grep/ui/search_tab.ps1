# ［2 検索］タブ（検索条件・検索の実行・結果の一覧と絞り込み）。

# ============================================================================
# ［2 検索］
# ============================================================================

# 結果（ヒットした行・ファイルごとの見出し）と表の中身は result_list.ps1
$script:search = $null
$script:lastSearch = $null
$script:sourceFolderMaps = @{}  # インデックスのフォルダ → インデックス名とクロール対象フォルダの対応（getSourceLocation のキャッシュ）
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
        $result = searchIndex $word $index.Files $simpleMatch $limit 50 -caseSensitive $shared.CaseSensitive -fileFilter $shared.FileFilter -cache $cache `
            -includeShapes $shared.IncludeShapes -includeComments $shared.IncludeComments -onProgress {
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
        IncludeShapes   = [bool]$ui.ShapeCheck.IsChecked
        IncludeComments = [bool]$ui.CommentCheck.IsChecked
    }
}

function setSearchOptionToUi {
    param (
        [hashtable]$option
    )

    $ui.RegexCheck.IsChecked = [bool]$option.UseRegex
    $ui.CaseCheck.IsChecked = [bool]$option.CaseSensitive
    $ui.FileFilterBox.Text = [string]$option.FileFilter
    $ui.ShapeCheck.IsChecked = [bool]$option.IncludeShapes
    $ui.CommentCheck.IsChecked = [bool]$option.IncludeComments
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
        $ui.SearchTargetText.Text = "検索対象：すべて（TSV $($summary['Count'].ToString('N0')) 件 ・ 最終取り込み $(formatTime $summary['LastWrite'])）"
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
    clearResults $word $pattern
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
        IncludeShapes = $option.IncludeShapes; IncludeComments = $option.IncludeComments
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
    while ($elapsed.ElapsedMilliseconds -lt ${searchPumpMilliseconds} -and $shared.Queue.TryDequeue([ref]$hit)) {
        # ヒットは元のファイルの見出しに生のまま貯めるだけにする。表の行（HitRow）は開いたときなどに作り（ensureRows）、
        # 表への反映は、この 1 回の最後に flushResults でまとめて行う。ヒットごとに通る処理なので、
        # 関数は初めてのファイル・場所のときだけ呼ぶ（呼ぶと、そのぶん結果を表に移すのが遅くなる）
        $relDir = [string]$hit.RelDir
        $fileKey = "$($hit.Root)\$relDir\$($hit.Book)"
        $group = $null
        if (!$script:fileGroups.TryGetValue($fileKey, [ref]$group)) {
            $group = newFileGroup $fileKey $relDir $hit.Book
        }
        $group.Hits.Add($hit)
        if ($group.AddLocation($hit.Location)) {
            addFileGroupLocation $group $hit.Book $hit.Location
        }
        [void]$script:dirtyGroups.Add($group)
        $script:hitCount++
    }    flushResults

    if ($shared.Total -gt 0) {
        $ratio = $shared.Done / $shared.Total
        $ui.SearchProgress.IsIndeterminate = $false
        $ui.SearchProgress.Value = $ratio
        $taskbar.ProgressState = "Normal"
        $taskbar.ProgressValue = $ratio
        if (!$shared.Stop) {
            $ui.SummaryText.Text = "検索中… $($shared.Done.ToString('N0')) / $($shared.Total.ToString('N0')) ファイル（$($script:hitCount.ToString('N0')) 件）"
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
    $taskbar.ProgressState = if (isIndexing) { $taskbar.ProgressState } else { "None" }
    $script:lastSearch = $s
    $seconds = ((Get-Date) - $s.Start).TotalSeconds
    updateSearchButton
    finishResults

    if ($shared.Error) {
        $ui.SummaryText.Text = "検索できませんでした"
        setStatus "検索できませんでした：$($shared.Error)"
        return
    }

    $count = $script:hitCount
    $files = $script:fileGroups

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
    if (isIndexing) {
        $status += "　インデックス作成中のため、作成途中のインデックスを検索しています。"
    }
    setStatus $status
}

# 検索結果を表に移す 1 回あたりの時間（ミリ秒）。タイマーの間隔（100 ミリ秒）より短くし、その間も画面が操作できるようにする
${searchPumpMilliseconds} = 60
$script:searchTimer = newTimer 100 { safe { pumpSearch } }

# ---- 絞り込み・選択行の詳細 ----

function applyFilter {
    $script:filterText = $ui.FilterBox.Text.Trim()
    applyResultFilter $script:filterText
    if ($script:lastSearch -and !$script:search -and $script:hitCount -gt 0) {
        $shown = getShownHitCount
        if ($script:filterText -eq "") {
            finishSummaryText
        } else {
            $ui.SummaryText.Text = "$($script:hitCount.ToString('N0')) 件中 $($shown.ToString('N0')) 件を表示"
        }
    }
}

function finishSummaryText {
    $ui.SummaryText.Text = "$($script:hitCount.ToString('N0')) 件（$($script:fileGroups.Count.ToString('N0')) ファイル）"
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
$ui.ShapeCheck.Add_Click({ safe { writeSearchOption @{ IncludeShapes = [bool]$ui.ShapeCheck.IsChecked } } })
$ui.CommentCheck.Add_Click({ safe { writeSearchOption @{ IncludeComments = [bool]$ui.CommentCheck.IsChecked } } })
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
