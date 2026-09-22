# ［2 比較］タブ（トグル・入力・要約・進み具合と、比べる流れ）。
#
# 比べる流れ（画面のスレッドでは重い処理をしない）:
#   ファイル: 抽出プロセス（differ.ps1）で 2 つを抽出 → 比較の runspace で差分を取る → 表示
#   フォルダ: 比較の runspace でファイルを列挙し、ツリーを出す → サイズが同じ組はハッシュを比べる →
#             違うかもしれない組を抽出プロセスへ（選んだファイルは先に）→ そろった組から差分を取り、ツリーの状態を更新
# 比較の runspace は 1 つだけ作り、比較の部品（..\lib.ps1）を 1 回だけ読み込んで使い回す（呼ぶたびに読み込むと遅いため）

${diffCacheSize} = 30   # 差分を覚えておくファイルの数（選び直したとき、計算し直さずにすぐ出す）

$script:dt = @{
    Mode         = "file"   # 比べている（比べた）もの
    LeftPath     = ""
    RightPath    = ""
    Options      = $null
    Subfolders   = $true
    Job          = ""       # 作業フォルダ
    Process      = $null    # 抽出プロセス
    Stopping     = $false
    Entries      = @()      # フォルダのとき: ファイルの一覧（getFolderEntries）
    ById         = @{}      # 番号 → ファイル
    ByRel        = @{}      # 相対パス → ファイル（一覧を先頭からたどると、ファイルが多いとき遅いため）
    ByPath       = @{}      # 元のファイルのパス → ファイル（抽出中のファイルを「比較中」にするとき）
    Results      = @{}      # 抽出の結果（"番号|側" → 行）。足された行だけを読んで足していく
    ResultOffset = 0        # 抽出結果.tsv の、次に読む位置（バイト）
    LastCurrent  = ""       # 前に見た、抽出中のファイル
    TreeRefreshed = [datetime]::MinValue  # ツリーを最後に並べ直した時刻（比べている間は間を空けて並べ直す）
    FileEntry    = $null    # ファイルのとき: 比べる 1 組
    Diffs        = @{}      # キー（相対パス。ファイルのときは "file"）→ FileDiff
    DiffOrder    = New-Object System.Collections.Generic.List[string]
    Queued       = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    CompareTotal = 0        # 中身を比べるファイルの数（進み具合の分母）
    Scanning     = $false
    Generation   = 0        # 比べ直すたびに増やす（前の比較の結果が後から届いても使わない）
}
# 入力欄の内容は、トグルの側ごとに覚える
$script:diffInputs = @{ file = @{ Left = ""; Right = "" }; folder = @{ Left = ""; Right = "" } }
$script:diffInputMode = "file"
$script:diffLoading = $false

# ============================================================================
# 比較の runspace（裏で差分を計算する。仕事は 1 つずつ順に）
# ============================================================================

$script:worker = @{ Runspace = $null; Current = $null; Queue = (New-Object System.Collections.Generic.List[object]) }

function startDiffWorker {
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()
    $script:worker.Runspace = $runspace
    # 最初の仕事として、比較の部品を読み込む（グローバルに読み込むので、後の仕事から使える）
    $lib = "${diffDir}\lib.ps1"
    queueDiffTask "load" ". `"$lib`"" @() $null
}

function queueDiffTask {
    # 比較の runspace に仕事を頼む。onDone { param($result, $errorText, $context) } は画面のスレッドで呼ぶ。priority なら先に回す。
    # onDone に値を渡すときは context を使う（GetNewClosure で閉じた scriptblock からは、画面の関数が見えなくなるため）
    param (
        [string]$name,
        [string]$script,
        [object[]]$arguments,
        [scriptblock]$onDone,
        [bool]$priority = $false,
        $context = $null
    )

    $task = @{ Name = $name; Script = $script; Arguments = $arguments; OnDone = $onDone; Generation = $script:dt.Generation; Context = $context }
    if ($priority -and $script:worker.Queue.Count -gt 0) {
        $script:worker.Queue.Insert(0, $task)
    } else {
        $script:worker.Queue.Add($task)
    }
    pumpDiffWorker
    $script:diffTimer.Start()
}

function pumpDiffWorker {
    # 終わった仕事の結果を渡し、次の仕事を始める
    $current = $script:worker.Current
    if ($current -and $current.Handle.IsCompleted) {
        $script:worker.Current = $null
        $result = $null
        $errorText = $null
        try {
            $output = $current.PS.EndInvoke($current.Handle)
            if ($output.Count -gt 0) { $result = $output[$output.Count - 1] }
            if ($current.PS.Streams.Error.Count -gt 0) { $errorText = $current.PS.Streams.Error[0].ToString() }
        } catch {
            $errorText = $_.Exception.Message
        } finally {
            $current.PS.Dispose()
        }
        # 比べ直した後に届いた前の比較の結果は使わない
        if ($current.Task.OnDone -and $current.Task.Generation -eq $script:dt.Generation) {
            & $current.Task.OnDone $result $errorText $current.Task.Context
        }
    }
    if ($null -eq $script:worker.Current -and $script:worker.Queue.Count -gt 0 -and $script:worker.Runspace) {
        $task = $script:worker.Queue[0]
        $script:worker.Queue.RemoveAt(0)
        $ps = [powershell]::Create()
        $ps.Runspace = $script:worker.Runspace
        [void]$ps.AddScript($task.Script)
        foreach ($argument in $task.Arguments) { [void]$ps.AddArgument($argument) }
        $script:worker.Current = @{ PS = $ps; Handle = $ps.BeginInvoke(); Task = $task }
    }
}

function clearDiffTasks {
    # 待っている仕事を捨てる（走っている仕事は終わるのを待つ。結果は Generation で捨てる）
    $script:worker.Queue.Clear()
}

# ============================================================================
# 入力
# ============================================================================

function getDiffOptionsFromUi {
    return [ordered]@{
        includeShapes    = [bool]$ui.ShapeCheck.IsChecked
        includeComments  = [bool]$ui.CommentCheck.IsChecked
        ignoreWhitespace = [bool]$ui.WhitespaceCheck.IsChecked
        caseSensitive    = [bool]$ui.CaseCheck.IsChecked
    }
}

function updateDiffSettingsSafe {
    # 比較の設定を保存する（保存できなくても画面は続ける）
    param (
        [System.Collections.IDictionary]$values
    )

    if ($script:diffLoading) {
        return
    }
    try {
        updateDiffSettings $values
    } catch {
        setStatus "設定を保存できませんでした：$($_.Exception.Message)"
    }
}

function getInputItem {
    # 入力欄のパスが、あるか・フォルダか（getCompareBlockReason に渡す形）
    param (
        [string]$path
    )

    $path = $path.Trim().Trim('"')
    $item = @{ Path = $path; Exists = $false; IsFolder = $false }
    if ($path) {
        try {
            if (Test-Path -LiteralPath (toLongPath $path) -PathType Container) {
                $item.Exists = $true
                $item.IsFolder = $true
            } elseif (Test-Path -LiteralPath (toLongPath $path) -PathType Leaf) {
                $item.Exists = $true
            }
        } catch { }
    }
    return $item
}

function updateCompareButton {
    # ［比較］を押せるか。押せなければ理由を出す
    $ui.LeftPlaceholder.Visibility = if ($ui.LeftBox.Text -eq "") { "Visible" } else { "Collapsed" }
    $ui.RightPlaceholder.Visibility = if ($ui.RightBox.Text -eq "") { "Visible" } else { "Collapsed" }
    $reason = getCompareBlockReason $script:diffInputMode (getInputItem $ui.LeftBox.Text) (getInputItem $ui.RightBox.Text) ([bool]$ui.SubfolderCheck.IsChecked)
    $ui.BlockReasonText.Text = $reason
    $ui.CompareButton.IsEnabled = ($reason -eq "")
    return ($reason -eq "")
}

function setDiffMode {
    # トグル（file / folder）を切り替える。入力欄は、切り替えた側で最後に入れたパスに入れ替える
    param (
        [string]$mode
    )

    if ($mode -ne $script:diffInputMode) {
        $script:diffInputs[$script:diffInputMode].Left = $ui.LeftBox.Text
        $script:diffInputs[$script:diffInputMode].Right = $ui.RightBox.Text
        $script:diffInputMode = $mode
        $ui.LeftBox.Text = $script:diffInputs[$mode].Left
        $ui.RightBox.Text = $script:diffInputs[$mode].Right
    }
    $view = getModeView $mode
    $script:diffLoading = $true
    try {
        $ui.ModeFileButton.IsChecked = ($mode -eq "file")
        $ui.ModeFolderButton.IsChecked = ($mode -eq "folder")
        $ui.CompactFileButton.IsChecked = ($mode -eq "file")
        $ui.CompactFolderButton.IsChecked = ($mode -eq "folder")
    } finally {
        $script:diffLoading = $false
    }
    $ui.ModeNoteText.Text = $view.Note
    $ui.LeftPickButton.Content = $view.PickLabel
    $ui.RightPickButton.Content = $view.PickLabel
    $ui.LeftPlaceholder.Text = $view.Placeholder
    $ui.RightPlaceholder.Text = $view.Placeholder
    # 押せないオプションは隠さずに灰色にする（切り替えても配置が動かないように）
    foreach ($check in @($ui.SubfolderCheck, $ui.HideSameCheck)) {
        $check.IsEnabled = $view.FolderOptions
        $check.ToolTip = if ($view.FolderOptionsTip) { $view.FolderOptionsTip } else { $null }
        [System.Windows.Controls.ToolTipService]::SetShowOnDisabled($check, $true)
    }
    $ui.StartHintTitle.Text = if ($mode -eq "folder") { "比べる 2 つのフォルダを選ぶか、ここへドロップしてください" } else { "比べる 2 つのファイルを選ぶか、ここへドロップしてください" }
    [void](updateCompareButton)
    updateDiffSettingsSafe ([ordered]@{ diffMode = $mode })
}

function showDiffInput {
    # 入力欄を開く（$true）・比べた後の 1 行にたたむ（$false）
    param (
        [bool]$open
    )

    $ui.InputPanel.Visibility = if ($open) { "Visible" } else { "Collapsed" }
    $ui.CompactBar.Visibility = if ($open) { "Collapsed" } else { "Visible" }
}

function pickDiffPath {
    # ［ファイル…］［フォルダ…］で選ぶ
    param (
        [string]$side
    )

    $box = if ($side -eq "left") { $ui.LeftBox } else { $ui.RightBox }
    $name = if ($side -eq "left") { "比較元" } else { "比較先" }
    if ($script:diffInputMode -eq "folder") {
        $path = selectFolder "${name}のフォルダを選んでください" $box.Text
    } else {
        $dialog = New-Object Microsoft.Win32.OpenFileDialog
        $dialog.Title = "${name}のファイルを選んでください"
        $dialog.Filter = "Office ファイル|*.xlsx;*.xlsm;*.xls;*.xlsb;*.docx;*.docm;*.doc;*.pptx;*.pptm;*.ppt|すべてのファイル|*.*"
        if ($box.Text -and (Test-Path -LiteralPath $box.Text)) {
            $dialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($box.Text)
        }
        $path = if ($dialog.ShowDialog($window)) { $dialog.FileName } else { $null }
    }
    if ($path) {
        $box.Text = $path
        [void](updateCompareButton)
    }
}

function applyDropAction {
    # ドロップされたものを入力欄に入れる（2 つまとめてなら、そのまま比べる）
    param (
        [string[]]$paths,
        [string]$target
    )

    $items = @($paths | ForEach-Object { @{ Path = $_; IsFolder = (Test-Path -LiteralPath $_ -PathType Container) } })
    $action = getDropAction $items $target $script:diffInputMode
    if ($null -eq $action) {
        setStatus "ファイルとフォルダを混ぜて比べることはできません。どちらかにそろえてドロップしてください"
        return
    }
    if ($action.Mode -ne $script:diffInputMode) {
        setDiffMode $action.Mode
    }
    if ($null -ne $action.Left) { $ui.LeftBox.Text = $action.Left }
    if ($null -ne $action.Right) { $ui.RightBox.Text = $action.Right }
    showDiffInput $true
    if ((updateCompareButton) -and $null -ne $action.Left -and $null -ne $action.Right) {
        startDiff
    }
}

# ============================================================================
# 比べる
# ============================================================================

function startDiff {
    # 入力欄の内容で比べる（比べ直しも同じ）
    if (!(updateCompareButton)) {
        showDiffInput $true
        return
    }
    stopDiffWork
    $script:dt.Generation++
    $script:dt.Mode = $script:diffInputMode
    $script:dt.LeftPath = (getInputItem $ui.LeftBox.Text).Path
    $script:dt.RightPath = (getInputItem $ui.RightBox.Text).Path
    $script:dt.Options = getDiffOptionsFromUi
    $script:dt.Subfolders = [bool]$ui.SubfolderCheck.IsChecked
    $script:dt.Entries = @()
    $script:dt.ById = @{}
    $script:dt.ByRel = @{}
    $script:dt.ByPath = @{}
    $script:dt.Results = @{}
    $script:dt.ResultOffset = 0
    $script:dt.LastCurrent = ""
    $script:dt.FileEntry = $null
    $script:dt.Diffs = @{}
    $script:dt.DiffOrder.Clear()
    $script:dt.Queued.Clear()
    $script:dt.CompareTotal = 0
    $script:dt.Stopping = $false

    # 比べ方は保存する。比較元・比較先のパスは保存しない（画面を閉じたら忘れてよい）
    $mode = $script:dt.Mode
    $values = [ordered]@{ diffMode = $mode; diffOptions = $script:dt.Options }
    if ($mode -eq "folder") {
        $values.diffSubfolders = $script:dt.Subfolders
    }
    updateDiffSettingsSafe $values

    # 前の比較の作業フォルダは消す（今表示している比較の分だけを残す）
    try { removeDiffJobDir $script:dt.Job } catch { }
    $script:dt.Job = newDiffJobDir

    showDiffInput $false
    $ui.LeftNameText.Text = $script:dt.LeftPath
    $ui.RightNameText.Text = $script:dt.RightPath
    $ui.LeftNameText.ToolTip = "比較元: $($script:dt.LeftPath)（クリックで入力欄を開き直す）"
    $ui.RightNameText.ToolTip = "比較先: $($script:dt.RightPath)（クリックで入力欄を開き直す）"
    $ui.StartHint.Visibility = "Collapsed"
    $ui.SummaryBar.Visibility = "Visible"
    $ui.DiffPanel.Visibility = "Visible"
    $ui.ExportButton.IsEnabled = $false
    $folder = ($mode -eq "folder")
    $ui.TreePanel.Visibility = if ($folder) { "Visible" } else { "Collapsed" }
    $ui.TreeSplitter.Visibility = if ($folder) { "Visible" } else { "Collapsed" }
    $ui.TreeRowDefinition.Height = if ($folder) { New-Object System.Windows.GridLength ((readDiffSettings).diffTreeHeight) } else { New-Object System.Windows.GridLength 0 }
    $ui.SplitterRowDefinition.Height = if ($folder) { New-Object System.Windows.GridLength 8 } else { New-Object System.Windows.GridLength 0 }
    $ui.ExpandTreeButton.Visibility = if ($folder) { "Visible" } else { "Collapsed" }
    $ui.CollapseTreeButton.Visibility = if ($folder) { "Visible" } else { "Collapsed" }
    $ui.CountsText.Text = ""
    $ui.FileTitleText.Text = ""
    resetTree

    if ($folder) {
        $ui.TreeLeftTitle.Text = "比較元  $($script:dt.LeftPath)"
        $ui.TreeRightTitle.Text = "比較先  $($script:dt.RightPath)"
        $ui.SummaryText.Text = "ファイルを探しています…"
        clearFileDiff "ファイルを選ぶと、ここに差分を出します。"
        $script:dt.Scanning = $true
        queueDiffTask "scan" @'
param($left, $right, $subfolders)
$l = getFolderFiles $left $subfolders
$r = getFolderFiles $right $subfolders
$count = @($l.Files).Count + @($r.Files).Count
if ($count -gt ${diffMaxFiles}) { return @{ TooMany = $count } }
return @{ Entries = (getFolderEntries @($l.Files) @($r.Files)); HasError = ($l.HasError -or $r.HasError) }
'@ @($script:dt.LeftPath, $script:dt.RightPath, $script:dt.Subfolders) { param($result, $errorText) onFolderScanned $result $errorText }
    } else {
        $entry = @{ RelPath = "file"; Left = @{ Path = $script:dt.LeftPath }; Right = @{ Path = $script:dt.RightPath }; Status = "pending"; Id = 1; Error = "" }
        $script:dt.FileEntry = $entry
        $script:dt.ById[1] = $entry
        $script:dt.ByPath[$script:dt.LeftPath] = $entry
        $script:dt.ByPath[$script:dt.RightPath] = $entry
        $ui.SummaryText.Text = "比べています…"
        clearFileDiff "比較元と比較先を読み取っています…"
        writeExtractRequest $script:dt.Job @(@{ Id = 1; Side = "left"; Path = $script:dt.LeftPath }, @{ Id = 1; Side = "right"; Path = $script:dt.RightPath })
        launchDiffer
    }
    updateDiffRunning
    $script:diffTimer.Start()
}

function onFolderScanned {
    # フォルダの列挙が終わった: ツリーを出し、サイズが同じ組のハッシュを比べる
    param (
        $result,
        $errorText
    )

    $script:dt.Scanning = $false
    if ($errorText -or $null -eq $result) {
        $ui.SummaryText.Text = "フォルダを読めませんでした。（$errorText）"
        updateDiffRunning
        return
    }
    if ($result.TooMany) {
        $ui.SummaryText.Text = "ファイルが多すぎます（$($result.TooMany) 件。上限 ${diffMaxFiles} 件）。フォルダを絞ってください。"
        updateDiffRunning
        return
    }
    $script:dt.Entries = @($result.Entries)
    foreach ($entry in $script:dt.Entries) {
        $script:dt.ByRel[$entry.RelPath] = $entry
        if ($entry.Left) { $script:dt.ByPath[$entry.Left.Path] = $entry }
        if ($entry.Right) { $script:dt.ByPath[$entry.Right.Path] = $entry }
    }
    if ($result.HasError) {
        setStatus "アクセスできないフォルダがありました。その中のファイルは比べていません"
    }
    refreshTree
    updateDiffSummary

    $pairs = @($script:dt.Entries | Where-Object { $_.HashNeeded } | ForEach-Object { @{ RelPath = $_.RelPath; Left = $_.Left.Path; Right = $_.Right.Path } })
    if ($pairs.Count -eq 0) {
        startFolderExtraction
        return
    }
    $script:dt.Scanning = $true
    $ui.ProgressText.Text = "同じファイルを探しています…"
    queueDiffTask "hash" @'
param($pairs)
$same = New-Object System.Collections.Generic.List[string]
foreach ($pair in $pairs) {
    try {
        if ((getFileHash $pair.Left) -eq (getFileHash $pair.Right)) { $same.Add($pair.RelPath) }
    } catch { }
}
return @{ Same = $same.ToArray() }
'@ @(, $pairs) { param($result, $errorText) onFolderHashed $result $errorText }
    updateDiffRunning
}

function onFolderHashed {
    param (
        $result,
        $errorText
    )

    $script:dt.Scanning = $false
    if ($result -and $result.Same) {
        $same = New-Object 'System.Collections.Generic.HashSet[string]' ([string[]]$result.Same), ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($entry in $script:dt.Entries) {
            if ($same.Contains($entry.RelPath)) { $entry.Status = "same" }
        }
    }
    startFolderExtraction
}

function startFolderExtraction {
    # 違うかもしれない組を、ツリーの順に抽出プロセスへ渡す（片方にしか無いファイルは、選んだときだけ抽出する）
    $items = New-Object System.Collections.Generic.List[object]
    $id = 0
    foreach ($entry in (getEntriesInTreeOrder $script:dt.Entries)) {
        $id++
        $entry.Id = $id
        $script:dt.ById[$id] = $entry
        if ($entry.Status -eq "pending") {
            $items.Add(@{ Id = $id; Side = "left"; Path = $entry.Left.Path })
            $items.Add(@{ Id = $id; Side = "right"; Path = $entry.Right.Path })
        }
    }
    $script:dt.CompareTotal = $items.Count / 2
    refreshTree
    updateDiffSummary
    if ($items.Count -gt 0) {
        writeExtractRequest $script:dt.Job $items.ToArray()
        launchDiffer
    }
    # 列挙の間に選ばれていたファイルがあれば、先に回す
    if ($script:tree.Selected) {
        onDiffFileSelected $script:tree.Selected
    }
    updateDiffRunning
}

function launchDiffer {
    # 抽出プロセスを起動する（動いていれば何もしない。抽出プロセスは抽出要求を読み直すので、足した分も抽出する）
    if (isDiffRunning) {
        return
    }
    $differ = "${diffDir}\differ.ps1"
    $arguments = "-NoProfile -ExecutionPolicy RemoteSigned -WindowStyle Hidden -File `"$differ`" -JobDir `"$($script:dt.Job)`""
    $script:dt.Process = Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -WorkingDirectory ${rootDir} -WindowStyle Hidden -PassThru
    # PowerShell 5.1 では、起動直後にハンドルを取っておかないと終了コードを取得できないことがある
    $null = $script:dt.Process.Handle
    updateDiffRunning
}

function isDiffRunning {
    # 抽出プロセスが動いているか（［9 プロセス停止］の注意にも使う）
    return ($null -ne $script:dt.Process -and !$script:dt.Process.HasExited)
}

function ensureExtracted {
    # 片方にしか無いファイルを、選んだときに抽出する（抽出要求に足し、抽出プロセスが止まっていれば起動する）
    param (
        $entry
    )

    $side = if ($entry.Left) { "left" } else { "right" }
    $path = if ($entry.Left) { $entry.Left.Path } else { $entry.Right.Path }
    $key = "$($entry.Id)|$side"
    if ($script:dt.Queued.Contains("extract:$key")) {
        return
    }
    [void]$script:dt.Queued.Add("extract:$key")
    $requests = @()
    $requestFile = Join-Path $script:dt.Job ${diffRequestFileName}
    if (Test-Path -LiteralPath $requestFile) {
        $requests = @(readExtractRequest $script:dt.Job)
    }
    writeExtractRequest $script:dt.Job (@($requests) + @(@{ Id = $entry.Id; Side = $side; Path = $path }))
    writeDiffPriority $script:dt.Job @(@{ Id = $entry.Id; Side = $side })
    launchDiffer
}

function stopDiffWork {
    # 抽出プロセスを止め、待っている仕事を捨てる
    clearDiffTasks
    if (isDiffRunning) {
        try { requestDiffStop $script:dt.Job } catch { }
        # 抽出中の 1 ファイルが終わるのを待たずに、プロセスごと終わらせる（子の Office は抽出プロセスの watchdog が見ている）
        try {
            if (!$script:dt.Process.WaitForExit(1500)) {
                [void]$script:dt.Process.CloseMainWindow()
                $script:dt.Process.Kill()
            }
        } catch { }
    }
    $script:dt.Process = $null
}

function stopDiff {
    # ［中止］・Esc: 抽出プロセスにファイルの切れ目で止まってもらう（それまでの結果は残す）
    if (!(isDiffRunning) -and !$script:dt.Scanning) {
        return
    }
    $script:dt.Stopping = $true
    clearDiffTasks
    if (isDiffRunning) {
        requestDiffStop $script:dt.Job
    }
    $script:dt.Scanning = $false
    setStatus "比較を中止しています…（読み取り中のファイルが終わると止まります）"
    updateDiffRunning
}

function updateDiffRunning {
    # 比べている間は［中止］、終われば［比べ直す］を出す
    $running = (isDiffRunning) -or $script:dt.Scanning -or ($null -ne $script:worker.Current -and $script:worker.Current.Task.Name -ne "load")
    $ui.StopButton.Visibility = if ($running) { "Visible" } else { "Collapsed" }
    $ui.RecompareButton.Visibility = if ($running) { "Collapsed" } else { "Visible" }
    $ui.DiffProgress.Visibility = if ($running) { "Visible" } else { "Collapsed" }
    $ui.ExportButton.IsEnabled = !$running -and ($script:dt.Diffs.Count -gt 0 -or @($script:dt.Entries).Count -gt 0)
    if (!$running) {
        $ui.ProgressText.Text = ""
    }
}

function updateDiffSummary {
    # 要約と進み具合
    if ($script:dt.Mode -eq "folder") {
        $counts = getFolderCounts @($script:dt.Entries)
        $ui.SummaryText.Text = getFolderSummaryText $counts
        $progress = getFolderProgressText $counts $script:dt.CompareTotal
        if ($progress) {
            $ui.ProgressText.Text = $progress
            $waiting = $counts.pending + $counts.running
            $ui.DiffProgress.Value = if ($script:dt.CompareTotal -gt 0) { ($script:dt.CompareTotal - $waiting) / $script:dt.CompareTotal } else { 0 }
        } elseif (!$script:dt.Scanning) {
            $ui.ProgressText.Text = ""
        }
        $hidden = $counts.same
        $ui.HideSameCheck.Content = if ($hidden -gt 0) { "同じファイルを隠す（$hidden）" } else { "同じファイルを隠す" }
        return
    }
    $diff = $script:dt.Diffs["file"]
    if ($diff) {
        $text = getFileSummaryText $diff
        $ui.SummaryText.Text = $text.Title
        $ui.CountsText.Text = $text.Counts
    }
}

# ============================================================================
# 抽出の結果を見る（タイマー）
# ============================================================================

$script:diffTimer = newTimer 200 { safe { onDiffTick } }

function onDiffTick {
    pumpDiffWorker
    $job = $script:dt.Job
    if ($job -and (Test-Path -LiteralPath $job)) {
        # 抽出の結果は、足された行だけを読む
        $read = readNewExtractResults $job $script:dt.ResultOffset
        $script:dt.ResultOffset = $read.Offset
        foreach ($row in $read.Rows) {
            $script:dt.Results["$($row.Id)|$($row.Side)"] = $row
            $entry = $script:dt.ById[$row.Id]
            if ($entry -and (checkEntryExtracted $entry)) { $script:tree.Dirty = $true }
        }
        # 抽出中のファイルは「比較中」にする
        $progress = readDiffProgress $job
        if ($progress -and $progress.Current -and $progress.Current -ne $script:dt.LastCurrent) {
            $script:dt.LastCurrent = $progress.Current
            $entry = $script:dt.ByPath[$progress.Current]
            if ($entry -and $entry.Status -eq "pending") {
                $entry.Status = "running"
                $script:tree.Dirty = $true
            }
        }
        if ($script:dt.Mode -eq "file") {
            updateFileProgress $progress
        } elseif ($script:tree.Dirty -and ((Get-Date) - $script:dt.TreeRefreshed).TotalMilliseconds -ge 1500) {
            # ファイルが多いと並べ直しに時間がかかるため、比べている間は 1.5 秒に 1 回まで
            refreshTree
            updateDiffSummary
            $script:dt.TreeRefreshed = Get-Date
        }
        # 抽出プロセスが終わった
        if ($null -ne $script:dt.Process -and $script:dt.Process.HasExited) {
            $code = $script:dt.Process.ExitCode
            $script:dt.Process = $null
            if ($code -eq 1) {
                $errorFile = Join-Path $job ${diffErrorFileName}
                $message = if (Test-Path -LiteralPath $errorFile) { readSharedText $errorFile } else { "原因は分かりません" }
                setStatus "比較を続けられませんでした：$message"
            } elseif ($code -eq 2 -or $script:dt.Stopping) {
                setStatus "比較を中止しました。読み取り済みのファイルの結果は残しています"
            }
            if ($script:dt.Mode -eq "folder") {
                refreshTree
                updateDiffSummary
            }
        }
    }
    updateDiffRunning
    if (!(isDiffRunning) -and $null -eq $script:worker.Current -and $script:worker.Queue.Count -eq 0 -and !$script:dt.Scanning) {
        if ($script:tree.Dirty -and $script:dt.Mode -eq "folder") {
            refreshTree
            updateDiffSummary
        }
        $script:diffTimer.Stop()
    }
}

function checkEntryExtracted {
    # 1 組の抽出の結果を見て、そろっていれば差分を頼む。状態が変わったら $true
    param (
        $entry
    )

    $results = $script:dt.Results
    $sides = @()
    if ($entry.Left) { $sides += "left" }
    if ($entry.Right) { $sides += "right" }
    $done = 0
    foreach ($side in $sides) {
        $result = $results["$($entry.Id)|$side"]
        if ($null -eq $result) { continue }
        if ($result.State -eq ${diffStateFailed}) {
            if ($entry.Status -in @("insert", "delete")) {
                # 片方にしか無いファイルは、状態（追加・削除）を変えずに理由だけを持つ（件数の「追加」「削除」から外さない）
                if (!$entry.Error) {
                    $entry.Error = getExtractFailureText $side $result.Message
                    if ($script:tree.Selected -eq $entry.RelPath) { clearFileDiff $entry.Error }
                }
                return $false
            }
            if ($entry.Status -ne "failed") {
                $entry.Status = "failed"
                $entry.Error = getExtractFailureText $side $result.Message
                if ($entry.RelPath -eq "file") {
                    clearFileDiff $entry.Error
                    $ui.SummaryText.Text = "比べられませんでした。"
                } elseif ($script:tree.Selected -eq $entry.RelPath) {
                    clearFileDiff $entry.Error
                }
                return $true
            }
            return $false
        }
        $done++
    }
    if ($done -eq $sides.Count) {
        if ($entry.Status -in @("pending", "running") -or ($entry.Status -in @("insert", "delete") -and $script:dt.Queued.Contains("extract:$($entry.Id)|$($sides[0])"))) {
            queueEntryCompare $entry ($script:tree.Selected -eq $entry.RelPath -or $entry.RelPath -eq "file")
        }
        return $false
    }
    return $false
}

function updateFileProgress {
    # ファイル同士の進み具合
    param (
        $progress
    )

    $entry = $script:dt.FileEntry
    if (!$entry -or $entry.Status -notin @("pending", "running")) { return }
    if ($progress -and $progress.Current) {
        $side = if ($progress.Current -eq $entry.Left.Path -and $progress.Done -eq 0) { "比較元" } else { "比較先" }
        $ui.ProgressText.Text = "${side}を読み取り中…"
        $ui.DiffProgress.Value = [Math]::Min(1, $progress.Done / 2.0)
        $ui.DiffMessage.Text = "${side}を読み取っています…"
    } elseif ($script:dt.Queued.Contains("compare:file")) {
        $ui.ProgressText.Text = "違いを調べています…"
        $ui.DiffProgress.Value = 1
        $ui.DiffMessage.Text = "違いを調べています…"
    }
}

function queueEntryCompare {
    # 抽出がそろった 1 組の差分を、比較の runspace に頼む（priority: 選んでいるファイルなら先に）
    param (
        $entry,
        [bool]$priority
    )

    $key = "compare:$($entry.RelPath)"
    if ($script:dt.Queued.Contains($key)) {
        return
    }
    [void]$script:dt.Queued.Add($key)
    $path = if ($entry.Right) { $entry.Right.Path } else { $entry.Left.Path }
    $leftDir = if ($entry.Left) { getExtractDir $script:dt.Job "left" $entry.Id } else { "" }
    $rightDir = if ($entry.Right) { getExtractDir $script:dt.Job "right" $entry.Id } else { "" }
    queueDiffTask "compare" @'
param($kind, $leftDir, $rightDir, $options)
$left = if ($leftDir) { readExtractedUnits $leftDir } else { [ordered]@{} }
$right = if ($rightDir) { readExtractedUnits $rightDir } else { [ordered]@{} }
return @{ Diff = (compareOfficeUnits $kind $left $right $options) }
'@ @((getDiffKind $path), $leftDir, $rightDir, $script:dt.Options) { param($result, $errorText, $context) onEntryCompared $context $result $errorText } $priority $entry.RelPath
}

function onEntryCompared {
    # 1 組の差分が届いた
    param (
        [string]$relPath,
        $result,
        $errorText
    )

    $entry = if ($relPath -eq "file") { $script:dt.FileEntry } else { $script:dt.ByRel[$relPath] }
    if ($null -eq $entry) { return }
    if ($errorText -or $null -eq $result -or $null -eq $result.Diff) {
        $entry.Status = "failed"
        $entry.Error = "違いを調べられませんでした。（$errorText）"
        if ($relPath -eq "file" -or $script:tree.Selected -eq $relPath) { clearFileDiff $entry.Error }
        $script:tree.Dirty = $true
        return
    }
    $diff = $result.Diff
    rememberDiff $relPath $diff
    if ($relPath -eq "file") {
        $entry.Status = "change"
        updateDiffSummary
        $ui.ProgressText.Text = ""
        showFileDiff $diff "$([System.IO.Path]::GetFileName($script:dt.LeftPath)) ⇄ $([System.IO.Path]::GetFileName($script:dt.RightPath))" $script:dt.LeftPath $script:dt.RightPath
    } else {
        if ($entry.Left -and $entry.Right) {
            setEntryResult $entry $diff
        }
        $script:tree.Dirty = $true
        if ($script:tree.Selected -eq $relPath) {
            showEntryDiff $entry
        }
    }
    updateDiffRunning
}

function rememberDiff {
    # 差分を覚える（直近 ${diffCacheSize} ファイル。古いものから捨てる。選んでいるファイルは捨てない）
    param (
        [string]$key,
        $diff
    )

    $script:dt.Diffs[$key] = $diff
    [void]$script:dt.DiffOrder.Remove($key)
    $script:dt.DiffOrder.Add($key)
    while ($script:dt.DiffOrder.Count -gt ${diffCacheSize}) {
        $old = $script:dt.DiffOrder[0]
        $script:dt.DiffOrder.RemoveAt(0)
        if ($old -ne $script:tree.Selected -and $old -ne "file") {
            $script:dt.Diffs.Remove($old)
            [void]$script:dt.Queued.Remove("compare:$old")
        }
    }
}

function showEntryDiff {
    # フォルダの 1 組の差分を出す
    param (
        $entry
    )

    $diff = $script:dt.Diffs[$entry.RelPath]
    $leftPath = if ($entry.Left) { $entry.Left.Path } else { "" }
    $rightPath = if ($entry.Right) { $entry.Right.Path } else { "" }
    showFileDiff $diff $entry.RelPath $leftPath $rightPath
}

function onDiffFileSelected {
    # フォルダのツリーでファイルを選んだ（止まってから 150 ms 後）
    param (
        [string]$relPath
    )

    $entry = $script:dt.ByRel[$relPath]
    if ($null -eq $entry) { return }
    $ui.FileTitleText.Text = $relPath
    if ($script:dt.Diffs.ContainsKey($relPath)) {
        showEntryDiff $entry
        return
    }
    switch ($entry.Status) {
        "same" {
            clearFileDiff "同じファイルです。（バイトまで同じ）"
            return
        }
        "failed" {
            clearFileDiff $entry.Error
            return
        }
        { $_ -in @("insert", "delete") } {
            if ($entry.Error) { clearFileDiff $entry.Error; return }
            if (!$script:dt.Job -or $entry.Id -eq $null) { clearFileDiff "読み取りを待っています…"; return }
            clearFileDiff "読み取っています…"
            ensureExtracted $entry
            $script:diffTimer.Start()
            return
        }
    }
    # 比べる組: 抽出が終わっていれば先に差分を頼み、まだなら先に抽出してもらう
    clearFileDiff "読み取っています…"
    if ($script:dt.Job -and $entry.Id) {
        $results = $script:dt.Results
        if ($results.ContainsKey("$($entry.Id)|left") -and $results.ContainsKey("$($entry.Id)|right")) {
            clearFileDiff "違いを調べています…"
            queueEntryCompare $entry $true
        } else {
            writeDiffPriority $script:dt.Job @(@{ Id = $entry.Id; Side = "left" }, @{ Id = $entry.Id; Side = "right" })
        }
    }
    $script:diffTimer.Start()
}

function onDiffFolderSelected {
    # フォルダの行を選んだ: 中の件数の内訳を出す
    param (
        $row
    )

    $prefix = "$($row.RelPath)\"
    $inside = @($script:dt.Entries | Where-Object { $_.RelPath.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase) })
    $counts = getFolderCounts $inside
    clearFileDiff "$($row.RelPath)`n$(getFolderSummaryText $counts)"
    $ui.FileTitleText.Text = $row.RelPath
}

function exportDiffReport {
    # ［結果をファイルに出力］: 比較結果ファイルを書く（フォルダのときは、変更のあるファイルの中身を読み直して書く）
    $options = $script:dt.Options
    if ($script:dt.Mode -eq "file") {
        $diff = $script:dt.Diffs["file"]
        if (!$diff) { return }
        $path = writeDiffReport (getFileReport $script:dt.LeftPath $script:dt.RightPath $diff $options (Get-Date))
        setStatus "比較結果を書き出しました：$path"
        return
    }
    $ui.ExportButton.IsEnabled = $false
    setStatus "比較結果を書き出しています…"
    $changed = @($script:dt.Entries | Where-Object { $_.Status -eq "change" } | ForEach-Object {
        @{ RelPath = $_.RelPath; Kind = (getDiffKind $_.Right.Path); Left = (getExtractDir $script:dt.Job "left" $_.Id); Right = (getExtractDir $script:dt.Job "right" $_.Id) }
    })
    queueDiffTask "report" @'
param($leftPath, $rightPath, $entries, $changed, $options, $subfolders, $resultFile)
$diffs = @{}
foreach ($item in $changed) {
    $diffs[$item.RelPath] = compareOfficeUnits $item.Kind (readExtractedUnits $item.Left) (readExtractedUnits $item.Right) $options
}
$lines = getFolderReport $leftPath $rightPath $entries $diffs $options $subfolders (Get-Date)
return @{ Path = (writeDiffReport $lines $resultFile) }
'@ @($script:dt.LeftPath, $script:dt.RightPath, @($script:dt.Entries), $changed, $options, $script:dt.Subfolders, ${diffResultFile}) {
        param($result, $errorText)
        if ($errorText -or !$result) {
            setStatus "比較結果を書き出せませんでした：$errorText"
        } else {
            setStatus "比較結果を書き出しました：$($result.Path)"
        }
        updateDiffRunning
    }
}

# ============================================================================
# ウィンドウから呼ばれる関数（scripts\tebunko\gui.ps1）
# ============================================================================

function onDiffKeyDown {
    # ［2 比較］のキー操作。扱ったら $true を返す
    param (
        $e,
        $modifiers
    )

    if ($e.Key -eq "F5") {
        safe { startDiff }
        return $true
    }
    if ($e.Key -eq "F8") {
        safe { moveChange $(if ($modifiers -eq "Shift") { -1 } else { 1 }) }
        return $true
    }
    if ($e.Key -eq "Escape" -and ((isDiffRunning) -or $script:dt.Scanning)) {
        safe { stopDiff }
        return $true
    }
    if ($modifiers -eq "Control" -and ($e.Key -eq "Down" -or $e.Key -eq "Up")) {
        safe { [void](moveDiffFile $(if ($e.Key -eq "Down") { 1 } else { -1 })) }
        return $true
    }
    return $false
}

function confirmDiffClosing {
    # ウィンドウを閉じる前。比較は閉じれば要らなくなるため、聞かずに止める
    return $true
}

function closeDiffPage {
    # ウィンドウを閉じた後: 抽出プロセスを止め、比較の runspace を閉じ、この画面の作業フォルダを消す
    try { stopDiffWork } catch { }
    try {
        if ($script:worker.Current) { $script:worker.Current.PS.Stop() }
        if ($script:worker.Runspace) { $script:worker.Runspace.Dispose() }
    } catch { }
    try { removeDiffJobDir (Join-Path ${diffTempRoot} ([string]$PID)) } catch { }
}

function startDiffPage {
    # 起動時: 前回の設定（比べ方・表示）を入れ、強制終了で残った作業フォルダを消し、比較の runspace を用意する。
    # 比較元・比較先の欄は空で始める（パスは保存しない）
    $settings = readDiffSettings
    $script:diffLoading = $true
    try {
        $script:diffInputMode = $settings.diffMode
        $ui.LeftBox.Text = $script:diffInputs[$settings.diffMode].Left
        $ui.RightBox.Text = $script:diffInputs[$settings.diffMode].Right
        $ui.SubfolderCheck.IsChecked = $settings.diffSubfolders
        $ui.HideSameCheck.IsChecked = $settings.diffHideSame
        $ui.ShapeCheck.IsChecked = $settings.diffOptions.includeShapes
        $ui.CommentCheck.IsChecked = $settings.diffOptions.includeComments
        $ui.WhitespaceCheck.IsChecked = $settings.diffOptions.ignoreWhitespace
        $ui.CaseCheck.IsChecked = $settings.diffOptions.caseSensitive
        $ui.FoldCheck.IsChecked = $settings.diffFoldSame
        $ui.ViewSideButton.IsChecked = ($settings.diffView -ne "list")
        $ui.ViewListButton.IsChecked = ($settings.diffView -eq "list")
    } finally {
        $script:diffLoading = $false
    }
    setDiffView $settings.diffView
    setDiffMode $settings.diffMode
    showDiffInput $true
    try { [void](removeStaleDiffDirs) } catch { }
    startDiffWorker
}

# ============================================================================
# イベント
# ============================================================================

$ui.ModeFileButton.Add_Checked({ safe { if (!$script:diffLoading) { setDiffMode "file" } } })
$ui.ModeFolderButton.Add_Checked({ safe { if (!$script:diffLoading) { setDiffMode "folder" } } })
# 比べた後の 1 行でトグルを切り替えると、入力欄を開き直す
$ui.CompactFileButton.Add_Checked({ safe { if (!$script:diffLoading) { if ((isDiffRunning)) { stopDiff }; setDiffMode "file"; showDiffInput $true } } })
$ui.CompactFolderButton.Add_Checked({ safe { if (!$script:diffLoading) { if ((isDiffRunning)) { stopDiff }; setDiffMode "folder"; showDiffInput $true } } })

$ui.LeftBox.Add_TextChanged({ safe { [void](updateCompareButton) } })
$ui.RightBox.Add_TextChanged({ safe { [void](updateCompareButton) } })
$ui.SubfolderCheck.Add_Click({ safe { [void](updateCompareButton); updateDiffSettingsSafe ([ordered]@{ diffSubfolders = [bool]$ui.SubfolderCheck.IsChecked }) } })
$ui.HideSameCheck.Add_Click({
    safe {
        updateDiffSettingsSafe ([ordered]@{ diffHideSame = [bool]$ui.HideSameCheck.IsChecked })
        if ($script:dt.Mode -eq "folder" -and @($script:dt.Entries).Count -gt 0) { refreshTree }
    }
})
foreach ($check in @($ui.ShapeCheck, $ui.CommentCheck, $ui.WhitespaceCheck, $ui.CaseCheck)) {
    $check.Add_Click({ safe { updateDiffSettingsSafe ([ordered]@{ diffOptions = (getDiffOptionsFromUi) }) } })
}
$ui.LeftPickButton.Add_Click({ safe { pickDiffPath "left" } })
$ui.RightPickButton.Add_Click({ safe { pickDiffPath "right" } })
$ui.SwapButton.Add_Click({
    safe {
        $left = $ui.LeftBox.Text
        $ui.LeftBox.Text = $ui.RightBox.Text
        $ui.RightBox.Text = $left
    }
})
$ui.SwapCompactButton.Add_Click({
    safe {
        $left = $ui.LeftBox.Text
        $ui.LeftBox.Text = $ui.RightBox.Text
        $ui.RightBox.Text = $left
        startDiff
    }
})
$ui.CompareButton.Add_Click({ safe { startDiff } })
$ui.RecompareButton.Add_Click({ safe { startDiff } })
$ui.StopButton.Add_Click({ safe { stopDiff } })
$ui.EditInputButton.Add_Click({ safe { showDiffInput $true } })
$ui.LeftNameText.Add_MouseLeftButtonUp({ safe { showDiffInput $true } })
$ui.RightNameText.Add_MouseLeftButtonUp({ safe { showDiffInput $true } })
$ui.ExportButton.Add_Click({ safe { exportDiffReport } })
$ui.LeftBox.Add_KeyDown({ param ($sender, $e) if ($e.Key -eq "Enter") { safe { startDiff } } })
$ui.RightBox.Add_KeyDown({ param ($sender, $e) if ($e.Key -eq "Enter") { safe { startDiff } } })

# ツリーの高さを覚える
$ui.TreeSplitter.Add_DragCompleted({
    safe {
        $height = [int]$ui.TreeRowDefinition.ActualHeight
        if ($height -gt 0) { updateDiffSettingsSafe ([ordered]@{ diffTreeHeight = $height }) }
    }
})

# ドラッグ&ドロップ: 入力欄の上ならその側、それ以外は画面の左半分・右半分で決める（2 つまとめてなら名前の順）
$ui.DiffRoot.Add_PreviewDragOver({
    param ($sender, $e)
    $e.Effects = if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) { "Copy" } else { "None" }
    $e.Handled = $true
})
$ui.DiffRoot.Add_PreviewDrop({
    param ($sender, $e)
    safe {
        if (!$e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) { return }
        $paths = [string[]]$e.Data.GetData([System.Windows.DataFormats]::FileDrop)
        $x = $e.GetPosition($ui.DiffRoot).X
        $target = if ($x -lt $ui.DiffRoot.ActualWidth / 2) { "left" } else { "right" }
        applyDropAction $paths $target
        $e.Handled = $true
    }
})
