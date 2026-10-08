# 検索の実行（開始・中止・結果の取り込み・終わりの文言）の画面層。検索条件の画面は search_bar.ps1、結果の絞り込みは result_filter.ps1。

# 結果（ヒットした行・ファイルごとの見出し）と表の中身は result_list.ps1
$script:search = $null
$script:lastSearch = $null
$script:sourceFolderMaps = @{}  # インデックスのフォルダ → インデックス名とクロール対象フォルダの対応（getSourceLocation のキャッシュ）
$script:filterText = ""
# 検索で読んだ集約ファイルの内容（画面を閉じるまで残し、次の検索では更新の無い集約ファイルをファイルから読まない）
$script:tsvCache = newTsvTextCache
# 検索の司令のスレッド（画面を開いている間 1 つ。閉じるときに gui.ps1 が Close する）
$script:searchService = newSearchService $script:tsvCache
# Windows Search が使えるか（高速検索の使用可否に使う。$null はまだ確かめていない）
$script:fastAvailable = $null

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
    $kindError = getSearchKindError (getFileKindsFromUi)
    if ($kindError -ne "") {
        setStatus $kindError
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

    # 新しく検索を始めるため、待っている元のファイルの確認（別の行）は打ち切る
    cancelPendingSourceLookup

    $useFast = (getFastSearchView $script:fastAvailable $useRegex $word).Usable
    $shared = $script:searchService.Request((newSearchRequest $word $simpleMatch $folders ${searchLimit} $option $useFast))
    $script:search = @{
        Shared = $shared
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

function clearSearchView {
    # ワークスペースを変えたとき、前のワークスペースの検索を止め、結果・プレビュー・インデックスの対応を捨てる
    # 待っている元のファイルの確認も打ち切る（前のワークスペースの行は、ここから先は開けない）
    cancelPendingSourceLookup
    $s = $script:search
    if ($s) {
        # 止めた検索のヒットは画面に移さない（司令のスレッドは止まり次第、次の要求に移る）
        $s.Shared.Stop = $true
        $script:search = $null
        $script:searchTimer.Stop()
        $ui.SearchProgress.Visibility = "Collapsed"
        $taskbar.ProgressState = "None"
    }
    $script:lastSearch = $null
    $script:sourceFolderMaps = @{}
    $script:fastAvailable = $null
    $ui.FilterBox.Text = ""
    $script:filterText = ""
    clearResults "" $null
    clearDetail
    $ui.SummaryText.Text = ""
    updateSearchButton
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
    }
    flushResults

    if ($shared.Total -gt 0) {
        $ratio = $shared.Done / $shared.Total
        $ui.SearchProgress.IsIndeterminate = $false
        $ui.SearchProgress.Value = $ratio
        $taskbar.ProgressState = "Normal"
        $taskbar.ProgressValue = $ratio
        if (!$shared.Stop) {
            $ui.SummaryText.Text = getSearchProgressText $script:hitCount
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

    if (!$shared.Finished -and !$script:searchService.IsRunning()) {
        # 司令のスレッドが止まった（lib.ps1 を読み込めなかった等）。次の検索で作り直す
        $shared.Error = $script:searchService.GetFailure()
        $shared.Finished = $true
    }
    if ($shared.Finished -and $shared.Queue.IsEmpty) {
        finishSearch
    }
}

function finishSearch {
    $s = $script:search
    $script:search = $null
    $script:searchTimer.Stop()

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
    if ($null -ne $shared.FastAvailable) {
        $script:fastAvailable = [bool]$shared.FastAvailable
        updateFastSearchView
    }

    # 高速検索では、候補の無いフォルダの集約ファイルを集めないため、集めた数が 0 でも「インデックスが無い」とは限らない
    if (!$shared.FastUsed -and $shared.IndexTotal -gt 0 -and $shared.Total -eq 0) {
        $ui.SummaryText.Text = getNoKindMatchText $s.Option.FileKinds
    } elseif (!$shared.FastUsed -and $shared.Total -eq 0) {
        $ui.SummaryText.Text = "インデックスが未作成です"
    } elseif ($count -eq 0) {
        # 0 件は結果欄に文を出さず、ステータスバーの「検索しました：… 0 件」で伝える
        $ui.SummaryText.Text = ""
    } else {
        $ui.SummaryText.Text = getSearchSummaryText $count $files.Count $seconds
    }

    $status = getSearchStatusText $s.Word $count (describeSearchOption $s.Option)
    if ($shared.Truncated) {
        # パスの順に検索して打ち切るため、この先のファイルのヒットは結果に出ない。そのことが分かる文面にする
        $status = "$(${searchLimit}.ToString('N0')) 件を超えたため、ここで打ち切りました。この先のファイルは検索していないため、ワード・種類・検索対象で絞り込んでください。"
    } elseif ($shared.Cancelled) {
        $status = "中止しました（$($count.ToString('N0')) 件まで表示）"
    }
    $missing = @($shared.Folders | Where-Object { !$_.Exists } | ForEach-Object { $_.Path })
    if ($missing.Count -gt 0) {
        $status += "　見つからない検索対象フォルダ：$($missing -join '、')"
    }
    if (isIndexing) {
        $status += "　更新中のため、更新の途中のインデックスを検索しています。"
    }
    setStatus $status
}

# 検索結果を表に移す 1 回あたりの時間（ミリ秒）。タイマーの間隔（100 ミリ秒）より短くし、その間も画面が操作できるようにする
${searchPumpMilliseconds} = 60
$script:searchTimer = newTimer 100 { safe { pumpSearch } }
