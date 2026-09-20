# エクスプローラー風のフォルダ選択ダイアログ。

function selectFolder {
    # エクスプローラー風のフォルダ選択ダイアログ（shared\xaml\dialog_folder_select.xaml）を開き、選んだフォルダを返す（キャンセルなら $null）。
    #   ・左：よく使う場所（デスクトップ・ドキュメント・ダウンロード）と PC のドライブのツリー
    #   ・右：今のフォルダの中身（フォルダと、そのフォルダにあるファイル。Office ファイルは色を変える）
    #   ・上：アドレスバー（パスの入力・貼り付けで移動）と［←］［→］［↑］
    #   ・下：選ぶフォルダのパス（一覧でフォルダを選ぶ・ドラッグ＆ドロップでも入る）
    # ※Windows 標準のフォルダ選択（WinForms の FolderBrowserDialog）はツリーだけでファイルが見えず、
    #   目的のフォルダにたどり着きにくいため、画面として作る。
    #   エクスプローラー形式の COM ダイアログ（IFileOpenDialog）は実行時コンパイル（csc.exe）が要るため使わない（12.2）
    param (
        [string]$description,
        [string]$initialPath,
        [System.Windows.Window]$owner = $window
    )

    $dialog = loadWindow "${sharedXamlDir}\dialog_folder_select.xaml"
    if ($owner) {
        $dialog.Owner = $owner
    }
    $ctrl = @{}
    foreach ($name in @(
            "DescriptionText", "BackButton", "ForwardButton", "UpButton", "AddressBox",
            "FolderTree", "EntryList", "EntryPlaceholder", "StatusText", "FolderBox", "OkButton", "ErrorText")) {
        $ctrl[$name] = $dialog.FindName($name)
    }
    $script:folderSelect = @{
        Window  = $dialog
        Ctrl    = $ctrl
        Current = ""                                                     # 今開いているフォルダ
        History = (New-Object System.Collections.Generic.List[string])   # ［←］［→］でたどる履歴
        Index   = -1                                                     # 履歴の今の位置
        All     = @()                                                    # 今のフォルダの中身（絞り込み前）
        Roots   = (New-Object 'System.Collections.ObjectModel.ObservableCollection[object]')
        Syncing = $false                                                 # ツリーの選択を合わせている間（移動を起こさない）
    }
    $ctrl.DescriptionText.Text = $description
    $ctrl.FolderTree.ItemsSource = $script:folderSelect.Roots
    loadFolderTreeRoots

    # ---- 操作 ----
    $ctrl.BackButton.Add_Click({ safe { moveFolderHistory -1 } })
    $ctrl.ForwardButton.Add_Click({ safe { moveFolderHistory 1 } })
    $ctrl.UpButton.Add_Click({ safe { goParentFolder } })
    $ctrl.AddressBox.Add_PreviewKeyDown({
        param ($sender, $e)
        if ($e.Key -eq "Return") {
            # Enter は［選択］（既定のボタン）ではなく、入力したパスへの移動にする
            safe { goFolder $script:folderSelect.Ctrl.AddressBox.Text | Out-Null }
            $e.Handled = $true
        }
    })
    $ctrl.EntryList.Add_SelectionChanged({ safe { onFolderEntrySelected } })
    $ctrl.EntryList.Add_MouseDoubleClick({ safe { openSelectedFolderEntry } })
    $ctrl.EntryList.Add_PreviewKeyDown({
        param ($sender, $e)
        if ($e.Key -eq "Return") {
            safe { openSelectedFolderEntry }
            $e.Handled = $true
        }
    })
    $ctrl.FolderTree.Add_SelectedItemChanged({
        safe {
            $d = $script:folderSelect
            $node = $d.Ctrl.FolderTree.SelectedItem
            if ($d.Syncing -or $null -eq $node -or $node.Path -eq "") {
                return
            }
            if (-not (testSamePath $node.Path $d.Current)) {
                goFolder $node.Path | Out-Null
            }
        }
    })
    # 展開したときに、そのフォルダのサブフォルダを読み込む（読み込みはイベントで駆動する。12 章）
    $ctrl.FolderTree.AddHandler([System.Windows.Controls.TreeViewItem]::ExpandedEvent, [System.Windows.RoutedEventHandler] {
        param ($s, $e)
        safe {
            $node = $e.OriginalSource.DataContext
            if ($node -is [FolderNode]) {
                loadFolderNode $node
            }
        }
    })
    # 選んだ項目が画面の外にあるとき（アドレスバーからの移動など）に見えるようにする
    $ctrl.FolderTree.AddHandler([System.Windows.Controls.TreeViewItem]::SelectedEvent, [System.Windows.RoutedEventHandler] {
        param ($s, $e)
        if ($e.OriginalSource -is [System.Windows.Controls.TreeViewItem]) {
            $e.OriginalSource.BringIntoView()
        }
    })
    $ctrl.FolderBox.Add_PreviewDragOver({ onFolderDragOver @args })
    $ctrl.FolderBox.Add_PreviewDrop({
        param ($sender, $e)
        safe {
            $folders = @(getDroppedFolders $e)
            if ($folders.Count -gt 0) {
                goFolder $folders[0] | Out-Null
            }
        }
        $e.Handled = $true
    })
    $ctrl.OkButton.Add_Click({
        safe {
            $d = $script:folderSelect
            $path = normalizeFolderPath $d.Ctrl.FolderBox.Text
            if ($path -eq "") {
                setFolderSelectError "フォルダを選んでください。"
                return
            }
            if (-not (Test-Path -LiteralPath (toLongPath $path) -PathType Container)) {
                setFolderSelectError "「${path}」は見つかりません。一覧から選ぶか、パスを確かめてください。"
                return
            }
            $d.Window.DialogResult = $true
        }
    })
    # キーボード操作はエクスプローラーに合わせる
    #   Alt+← / Alt+→：戻る・進む、Alt+↑ / BackSpace：1 つ上へ、F5 / Ctrl+R：読み直す、
    #   F4 / Alt+D / Ctrl+L：アドレスバーへ、Esc：キャンセル（IsCancel）
    $dialog.Add_PreviewKeyDown({
        param ($sender, $e)
        $key = if ($e.Key -eq "System") { $e.SystemKey } else { $e.Key }
        $modifiers = $e.KeyboardDevice.Modifiers
        $alt = (($modifiers -band [System.Windows.Input.ModifierKeys]::Alt) -ne 0)
        $ctrl = (($modifiers -band [System.Windows.Input.ModifierKeys]::Control) -ne 0)
        $inTextBox = ($e.OriginalSource -is [System.Windows.Controls.TextBox])
        if ($key -eq "F5" -or ($ctrl -and $key -eq "R")) {
            safe { reloadFolder }
            $e.Handled = $true
        } elseif ($alt -and $key -eq "Left") {
            safe { moveFolderHistory -1 }
            $e.Handled = $true
        } elseif ($alt -and $key -eq "Right") {
            safe { moveFolderHistory 1 }
            $e.Handled = $true
        } elseif (($alt -and $key -eq "Up") -or ($key -eq "Back" -and -not $inTextBox)) {
            # BackSpace は入力欄では文字を消すため、入力欄以外のときだけ 1 つ上へ
            safe { goParentFolder }
            $e.Handled = $true
        } elseif ($key -eq "F4" -or ($alt -and $key -eq "D") -or ($ctrl -and $key -eq "L")) {
            safe { focusAddressBox }
            $e.Handled = $true
        }
    })
    # マウスの戻る・進むボタン（サイドボタン）でも履歴をたどる
    $dialog.Add_PreviewMouseDown({
        param ($sender, $e)
        if ($e.ChangedButton -eq [System.Windows.Input.MouseButton]::XButton1) {
            safe { moveFolderHistory -1 }
            $e.Handled = $true
        } elseif ($e.ChangedButton -eq [System.Windows.Input.MouseButton]::XButton2) {
            safe { moveFolderHistory 1 }
            $e.Handled = $true
        }
    })

    # ---- 最初に開くフォルダ ----
    $start = normalizeFolderPath $initialPath
    if ($start -ne "" -and -not (Test-Path -LiteralPath (toLongPath $start) -PathType Container)) {
        # 指定のフォルダが無ければ、その上の、今もあるフォルダを開く
        $start = getExistingFolder $start
    }
    $opened = $false
    foreach ($candidate in @($start) + @(getQuickFolders | ForEach-Object { $_.Path }) + @(getComputerFolders | ForEach-Object { $_.Path })) {
        if ($candidate -eq "") {
            continue
        }
        if (goFolder $candidate) {
            $opened = $true
            break
        }
    }
    if (-not $opened) {
        setFolderSelectError "開けるフォルダが見つかりません。上の欄にフォルダのパスを入力してください。"
        updateFolderSelectButtons
    }
    $ctrl.EntryList.Focus() | Out-Null

    $result = $null
    if ($dialog.ShowDialog()) {
        $result = normalizeFolderPath $ctrl.FolderBox.Text
    }
    $script:folderSelect = $null
    return $result
}

function setFolderSelectError {
    # フォルダ選択ダイアログの下に出す、直してほしい内容（空なら消す）
    param (
        [string]$message
    )

    $ctrl = $script:folderSelect.Ctrl
    $ctrl.ErrorText.Text = $message
    $ctrl.ErrorText.Visibility = if ($message -eq "") { "Collapsed" } else { "Visible" }
}

function loadFolderTreeRoots {
    # ツリーの一番上（「よく使う場所」「PC」）を作る。どちらも最初から開いておく
    $d = $script:folderSelect
    $d.Roots.Clear()

    $quick = [FolderNode]::new($null, "よく使う場所", "")
    $quick.IsHeader = $true
    $quick.Loaded = $true
    foreach ($place in @(getQuickFolders)) {
        addFolderTreeChild $quick $place.Name $place.Path | Out-Null
    }
    if ($quick.Children.Count -gt 0) {
        $quick.SetExpanded($true)
        $d.Roots.Add($quick)
    }

    $computer = [FolderNode]::new($null, "PC", "")
    $computer.IsHeader = $true
    $computer.Loaded = $true
    foreach ($drive in @(getComputerFolders)) {
        addFolderTreeChild $computer $drive.Name $drive.Path | Out-Null
    }
    $computer.SetExpanded($true)
    $d.Roots.Add($computer)
}

function addFolderTreeChild {
    # ツリーに子（フォルダ）を1つ足す。サブフォルダがあれば ▷ を出すための仮の子を入れておく
    param (
        [FolderNode]$parent,
        [string]$name,
        [string]$path
    )

    $node = [FolderNode]::new($parent, $name, $path)
    if (testHasSubFolders $path) {
        $node.AddPlaceholder()
    }
    $parent.Children.Add($node)
    return $node
}

function loadFolderNode {
    # ツリーのノードのサブフォルダを読み込む（1回だけ）
    param (
        [FolderNode]$node
    )

    if ($null -eq $node -or $node.Loaded -or $node.IsPlaceholder -or $node.Path -eq "") {
        return
    }
    $node.Loaded = $true
    $node.Children.Clear()
    foreach ($entry in @((getFolderEntries $node.Path -foldersOnly).Entries)) {
        addFolderTreeChild $node $entry.Name $entry.Path | Out-Null
    }
}

function findFolderNode {
    # ツリーから path のノードを探す。**まだ開いていないフォルダは開かない**（エクスプローラーと同じく、
    # 移動しただけで左のツリーが勝手に展開されないようにする）。見つからなければ $null
    param (
        [FolderNode]$node,
        [string]$path
    )

    if ($node.IsPlaceholder -or $node.Path -eq "") {
        return $null
    }
    if (testSamePath $node.Path $path) {
        return $node
    }
    if (-not $node.Loaded) {
        return $null
    }
    if (-not ([string]$path).StartsWith($node.Path.TrimEnd("\") + "\", [System.StringComparison]::OrdinalIgnoreCase)) {
        return $null
    }
    foreach ($child in $node.Children) {
        $found = findFolderNode $child $path
        if ($null -ne $found) {
            return $found
        }
    }
    return $null
}

function findFolderNodeInTree {
    # ツリー全体（読み込んである範囲）から path のノードを探す
    param (
        [string]$path
    )

    foreach ($root in $script:folderSelect.Roots) {
        foreach ($child in $root.Children) {
            $found = findFolderNode $child $path
            if ($null -ne $found) {
                return $found
            }
        }
    }
    return $null
}

function revealFolderTree {
    # 開いているフォルダをツリーでも選んだ状態にする（ツリーの選択が移動を起こさないよう Syncing を立てる）。
    # ツリーは、すでに開いてあるところと、その 1 つ下までを追随させる（右で下りると 1 段ずつ開く）。
    # 深いパスを貼り付けたときに途中のフォルダをすべて開くことはしない
    param (
        [string]$path
    )

    $d = $script:folderSelect
    $d.Syncing = $true
    try {
        $selected = $d.Ctrl.FolderTree.SelectedItem
        if ($null -ne $selected -and (testSamePath $selected.Path $path)) {
            return
        }
        $found = findFolderNodeInTree $path
        if ($null -eq $found) {
            # 1 つ上のフォルダがツリーにあれば、そこだけ開いて中を出す
            $parentPath = getParentFolderPath $path
            if ($parentPath -ne "") {
                $parent = findFolderNodeInTree $parentPath
                if ($null -ne $parent) {
                    loadFolderNode $parent
                    $parent.SetExpanded($true)
                    $found = findFolderNodeInTree $path
                }
            }
        }
        if ($null -ne $selected) {
            $selected.SetSelected($false)
        }
        if ($null -ne $found) {
            $found.SetSelected($true)
        }
        # ツリーに無いフォルダ（開いていない深いパス・ネットワークのパスなど）は、ツリーの選択を外すだけにする
    } finally {
        $d.Syncing = $false
    }
}

function goFolder {
    # フォルダを開く（一覧・アドレスバー・ツリーをそのフォルダに合わせる）。開けたかを返す
    param (
        [string]$path,
        [bool]$addHistory = $true
    )

    $d = $script:folderSelect
    $path = normalizeFolderPath $path
    if ($path -eq "") {
        setFolderSelectError "フォルダのパスを入力してください。"
        return $false
    }
    $entries = getFolderEntries $path
    if ($entries.Error -ne "") {
        setFolderSelectError "「${path}」を開けません。$($entries.Error)"
        return $false
    }

    setFolderSelectError ""
    $d.Current = $path
    if ($addHistory) {
        pushFolderHistory $path
    }
    $d.Ctrl.AddressBox.Text = $path
    $d.Ctrl.FolderBox.Text = $path
    $d.All = @($entries.Entries | ForEach-Object { newFolderEntry $_ })
    updateFolderEntryList
    $d.Ctrl.StatusText.Text = describeFolderEntries $entries
    revealFolderTree $path
    updateFolderSelectButtons
    return $true
}

function describeFolderEntries {
    # 一覧の下に出す件数（目的のフォルダかどうかの目安にする）
    param (
        $entries
    )

    $text = "フォルダー {0:#,0} 個 ・ Office ファイル {1:#,0} 個" -f $entries.FolderCount, $entries.OfficeCount
    if ($entries.Truncated) {
        $text += "（中身が多いため、先頭だけを表示しています）"
    }
    return $text
}

function newFolderEntry {
    # 一覧の1行を作る（shared\core\folder.ps1 の getFolderEntries が返した中身から）
    param (
        $entry
    )

    $row = [FolderEntry]::new()
    $row.Name = $entry.Name
    $row.Path = $entry.Path
    $row.IsFolder = $entry.IsFolder
    $row.IsOffice = $entry.IsOffice
    $row.Kind = getFolderEntryKind $entry
    $row.UpdatedText = if ($null -ne $entry.Updated) { $entry.Updated.ToString("yyyy/MM/dd H:mm") } else { "" }
    return $row
}

function getFolderEntryKind {
    # 一覧の「種類」列に出す名前
    param (
        $entry
    )

    if ($entry.IsFolder) {
        return "フォルダー"
    }
    if (-not $entry.IsOffice) {
        return "ファイル"
    }
    $ext = [System.IO.Path]::GetExtension($entry.Name).ToLowerInvariant()
    if ($ext.StartsWith(".xls")) {
        return "Excel ブック"
    }
    if ($ext.StartsWith(".doc")) {
        return "Word 文書"
    }
    return "PowerPoint プレゼンテーション"
}

function updateFolderEntryList {
    # 今のフォルダの中身を一覧に出す。
    # ItemsSource には必ず配列を渡す（1 件のときに配列が展開されると渡せないため @() で包む）
    $d = $script:folderSelect
    $rows = @($d.All)
    $d.Ctrl.EntryList.ItemsSource = $rows
    $d.Ctrl.EntryPlaceholder.Text = "このフォルダの中にはフォルダもファイルもありません。このフォルダでよければ［選択］を押してください。"
    $d.Ctrl.EntryPlaceholder.Visibility = if ($rows.Count -eq 0) { "Visible" } else { "Collapsed" }
}

function onFolderEntrySelected {
    # 一覧でフォルダを選んだら、下の欄をそのフォルダにする（選んでいなければ今のフォルダ）
    $d = $script:folderSelect
    $row = $d.Ctrl.EntryList.SelectedItem
    $path = if ($null -ne $row -and $row.IsFolder) { $row.Path } else { $d.Current }
    if ($path -ne "" -and -not (testSamePath $d.Ctrl.FolderBox.Text $path)) {
        $d.Ctrl.FolderBox.Text = $path
        setFolderSelectError ""
    }
}

function openSelectedFolderEntry {
    # 一覧で選んでいるフォルダを開く（ダブルクリック・Enter）
    $row = $script:folderSelect.Ctrl.EntryList.SelectedItem
    if ($null -ne $row -and $row.IsFolder) {
        goFolder $row.Path | Out-Null
    }
}

function focusAddressBox {
    # アドレスバーへ移り、今のパスを選んだ状態にする（F4 / Alt+D / Ctrl+L）
    $address = $script:folderSelect.Ctrl.AddressBox
    $address.Focus() | Out-Null
    $address.SelectAll()
}

function goParentFolder {
    # 1つ上のフォルダへ
    $parent = getParentFolderPath $script:folderSelect.Current
    if ($parent -ne "") {
        goFolder $parent | Out-Null
    }
}

function reloadFolder {
    # 今のフォルダを読み直す（ツリーの下も読み込み直す）
    $d = $script:folderSelect
    if ($d.Current -eq "") {
        return
    }
    $node = $d.Ctrl.FolderTree.SelectedItem
    if ($node -is [FolderNode] -and $node.Loaded) {
        $node.Loaded = $false
        $node.Children.Clear()
        loadFolderNode $node
    }
    goFolder $d.Current $false | Out-Null
}

function pushFolderHistory {
    # ［←］［→］でたどる履歴に足す（今の位置より先は捨てる）
    param (
        [string]$path
    )

    $d = $script:folderSelect
    if ($d.Index -ge 0 -and (testSamePath $d.History[$d.Index] $path)) {
        return
    }
    while ($d.History.Count -gt ($d.Index + 1)) {
        $d.History.RemoveAt($d.History.Count - 1)
    }
    $d.History.Add($path)
    $d.Index = $d.History.Count - 1
}

function moveFolderHistory {
    # 履歴を1つ戻る・進む
    param (
        [int]$step
    )

    $d = $script:folderSelect
    $next = $d.Index + $step
    if ($next -lt 0 -or $next -ge $d.History.Count) {
        return
    }
    $before = $d.Index
    $d.Index = $next
    if (-not (goFolder $d.History[$next] $false)) {
        $d.Index = $before   # 消えたフォルダなどで開けなければ、位置は戻す
    }
    updateFolderSelectButtons
}

function updateFolderSelectButtons {
    # ［←］［→］［↑］の使える・使えないを合わせる
    $d = $script:folderSelect
    $d.Ctrl.BackButton.IsEnabled = ($d.Index -gt 0)
    $d.Ctrl.ForwardButton.IsEnabled = ($d.Index -ge 0 -and $d.Index -lt ($d.History.Count - 1))
    $d.Ctrl.UpButton.IsEnabled = ((getParentFolderPath $d.Current) -ne "")
}

function getDroppedFolders {
    param (
        [System.Windows.DragEventArgs]$e
    )

    if (!$e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
        return @()
    }
    return @($e.Data.GetData([System.Windows.DataFormats]::FileDrop) | Where-Object { Test-Path -LiteralPath $_ -PathType Container })
}

function onFolderDragOver {
    param ($sender, [System.Windows.DragEventArgs]$e)

    $e.Effects = if ((getDroppedFolders $e).Count -gt 0) { "Copy" } else { "None" }
    $e.Handled = $true
}
