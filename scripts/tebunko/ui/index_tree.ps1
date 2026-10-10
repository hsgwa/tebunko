# ［検索］タブの、検索対象インデックスのツリー。


# ---- 検索対象インデックスのツリー ----

$script:indexRoots = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$ui.IndexTree.ItemsSource = $script:indexRoots

$script:indexTreeRequest = @{ Value = 0 }   # ツリーの読み込みの依頼番号（新しい依頼が出たら、前の依頼の結果は捨てる）
$script:indexTreeLoading = $false           # 検索対象のツリーを読み込んでいる最中か（ネットワークのとき、裏で読む間）
$script:indexTreeConnectError = ""          # ワークスペースに接続できなかったときの、そのワークスペース（できたなら空）

function loadIndexTree {
    # インデックスの一覧（getIndexTreeData）をツリーに読み込む。一番上の項目がインデックス 1 件で、
    # ［インデックス管理］で作ったインデックスがすべて並ぶ。
    # 保存したチェックなしのフォルダと、読み込み前の展開の状態は戻す。
    # ワークスペースがネットワークの場所のときは、画面のスレッドで読まず、裏の列（network）で読んでから反映する
    $expanded = New-Object 'System.Collections.Generic.List[string]'
    foreach ($node in $script:indexRoots) {
        $node.AddExpanded($expanded)
    }
    $paths = @($expanded.ToArray()) + @(readSearchExcludes | ForEach-Object { [string]$_.Path })

    $requestBox = $script:indexTreeRequest
    $requestBox.Value++
    $requestId = $requestBox.Value
    $dir = $workspace.IndexDir
    $statusPath = $workspace.StatusFile
    $settingsPath = ${settingsFile}
    $expandedPaths = $expanded.ToArray()

    if (!(testNetworkPath $dir)) {
        applyIndexTreeData (getIndexTreeData $dir $statusPath $settingsPath $paths) $expandedPaths
        return
    }

    $script:indexRoots.Clear()
    $script:indexTreeLoading = $true
    $script:indexTreeConnectError = ""
    updateSearchTarget
    $otherState = ${pathStateOther}
    $applyData = ${function:applyIndexTreeData}   # 終わったときの処理は、関数を変数に取って呼ぶ（クロージャからは関数の名前を引けないため）
    startJob {
        param ($dir, $statusPath, $settingsPath, $paths)
        getIndexTreeData $dir $statusPath $settingsPath $paths
    } @($dir, $statusPath, $settingsPath, [string[]]$paths) {
        param ($output, $errorText)
        if ($requestId -ne $requestBox.Value) {
            # 待っている間に、読み込みをやり直した。前の依頼は捨てる
            return
        }
        if ($errorText -or !$output -or $output.Count -eq 0) {
            & $applyData @{ State = $otherState; Message = [string]$errorText } $expandedPaths
        } else {
            & $applyData $output[0] $expandedPaths
        }
    }.GetNewClosure() (getWorkspaceJobQueue $dir)
}

function applyIndexTreeChildren {
    # getIndexTreeData が集めた子（フォルダ → getIndexFolderChildren の結果）を、ノードの下に再帰して入れる
    param (
        [IndexNode]$node,
        [hashtable]$children
    )

    $key = $node.FullPath()
    if (!$children.ContainsKey($key)) {
        return
    }
    if ($node.NeedsLoad()) {
        if ($children[$key].Error) { return }
        $node.ApplyChildren($children[$key])
    }
    foreach ($child in @($node.Children)) {
        if (!$child.IsPlaceholder -and !$child.IsFiles) {
            applyIndexTreeChildren $child $children
        }
    }
}

function applyIndexTreeData {
    # getIndexTreeData の結果をツリーに反映する（読み込み口。画面のスレッドで呼ぶ）。
    # 接続できない・その他のときは、ツリーを空にして、ステータスに知らせる
    param (
        $data,
        [string[]]$expanded = @()
    )

    $script:indexTreeLoading = $false
    $script:indexTreeConnectError = ""
    $script:indexRoots.Clear()
    if ($data.State -eq ${pathStateUnreachable} -or $data.State -eq ${pathStateOther}) {
        $script:indexTreeConnectError = [string]$workspace.Dir
        setStatus (getWorkspaceUnreachableText $workspace.Dir)
        updateSearchTarget
        return
    }

    $root = [string]$data.Root
    $script:sourceFolderMaps[$root.TrimEnd("\")] = $data.Sources
    foreach ($index in @($data.Indexes)) {
        $sourcePath = if ($index.SourcePath) { $index.SourcePath } else { $null }
        $node = [IndexNode]::CreateRoot($root, $index.Name, $index.Name, $sourcePath, [bool]$index.Exists, [bool]$index.HasSubfolders)
        applyIndexTreeChildren $node $data.Children
        $script:indexRoots.Add($node)
    }

    foreach ($exclude in @(readSearchExcludes)) {
        foreach ($node in $script:indexRoots) {
            $node.ApplyExclude($exclude.Path, $exclude.Subfolders)
        }
    }
    foreach ($node in $script:indexRoots) {
        foreach ($path in $expanded) {
            $found = $node.Find($path)
            if ($found) {
                $found.SetExpanded($true)
            }
        }
    }
    updateSearchTarget
}

function expandIndexNode {
    # フォルダを展開したときに、子を読み込む。ネットワークの場所なら、裏の列（network）で読む間は「読み込み中…」を出しておく。
    # 読めなかったときは、フォルダを閉じて読み込み前の状態に戻し（「読み込み中…」を見せ続けない。もう一度開けば読み直す）、ステータスに知らせる
    param (
        [IndexNode]$node
    )

    if (!$node.NeedsLoad() -or $node.IsLoading) {
        return
    }
    $dir = $node.FullPath()
    if (!(testNetworkPath $dir)) {
        applyIndexFolderChildren $node (getIndexFolderChildren $dir)
        return
    }

    $node.IsLoading = $true
    $requestBox = $script:indexTreeRequest
    $generation = $requestBox.Value
    $applyChildren = ${function:applyIndexFolderChildren}
    startJob {
        param ($dir)
        getIndexFolderChildren $dir
    } @($dir) {
        param ($output, $errorText)
        $node.IsLoading = $false
        if ($generation -ne $requestBox.Value) {
            # ツリーを読み込み直した。このノードはもう画面に無い
            return
        }
        if ($errorText -or !$output -or $output.Count -eq 0) {
            & $applyChildren $node @{ Error = $(if ($errorText) { [string]$errorText } else { "結果がありません" }) }
            return
        }
        & $applyChildren $node $output[0]
    }.GetNewClosure() (getWorkspaceJobQueue $dir)
}

function applyIndexFolderChildren {
    # expandIndexNode の続き。子を入れる（読めなかったときは、フォルダを閉じて読み込み前に戻し、ステータスに知らせる）
    param (
        [IndexNode]$node,
        $data
    )

    if ($data.Error) {
        $node.SetExpanded($false)
        setStatus (getTreeFolderFailedText $node.Name ([string]$data.Error))
        return
    }
    $node.ApplyChildren($data)
    updateSearchTarget
}

function getSearchTargets {
    # 検索対象ツリーでチェックしたフォルダ（SearchTarget の配列。getIndexPackFiles に渡す）
    $targets = New-Object 'System.Collections.Generic.List[SearchTarget]'
    foreach ($node in $script:indexRoots) {
        $node.AddTargets($targets)
    }
    return $targets.ToArray()
}

function isAllIndexChecked {
    return @($script:indexRoots | Where-Object { $_.IsChecked -ne $true }).Count -eq 0
}

function describeSearchTargets {
    # 検索対象の表示（先頭の 3 件まで）。インデックスのフォルダ（work\index）からの相対パスは
    # 「インデックス名\その下のフォルダ」のため、そのまま表示に使う
    param (
        [object[]]$targets
    )

    $names = @($targets | ForEach-Object {
        $target = $_
        $name = ([string]$target.RelPath).Trim("\")
        if (!$target.Recurse) {
            $name += "（直下のファイル）"
        }
        $name
    })
    if ($names.Count -gt 3) {
        return "$($names[0..2] -join '、') ほか $($names.Count - 3) か所"
    }
    return $names -join "、"
}

function saveSearchExcludes {
    # チェックなしのフォルダを設定に保存する。見つからないインデックスのフォルダ（ネットワークのドライブが切れているなど）の記録は残す
    $excludes = New-Object 'System.Collections.Generic.List[SearchExclude]'
    foreach ($node in $script:indexRoots) {
        $node.AddExcludes($excludes)
    }
    $roots = @($script:indexRoots | Where-Object { $_.Exists } | ForEach-Object { $_.FullPath().TrimEnd("\") })
    $kept = @(readSearchExcludes | Where-Object {
        $path = $_.Path
        @($roots | Where-Object { $path -eq $_ -or $path.StartsWith("$_\", [System.StringComparison]::OrdinalIgnoreCase) }).Count -eq 0
    })
    writeSearchExcludes (@($kept) + @($excludes.ToArray()))
}

function onIndexTreeChecked {
    saveSearchExcludes
    updateSearchTarget
}

function setAllIndexChecked {
    param (
        [bool]$checked
    )

    foreach ($node in $script:indexRoots) {
        $node.SetChecked($checked)
    }
    onIndexTreeChecked
}

# ツリーのチェックボックスのクリック。チェックは OneWay バインドのため、クリックされたノードの Toggle() で
# 3状態（子・親への伝播）を反映してから保存する（PS class はセッターにロジックを書けないため、ここで行う）
$ui.IndexTree.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler] {
    param ($s, $e)
    safe {
        $cb = $e.OriginalSource
        if ($cb -is [System.Windows.Controls.CheckBox] -and $cb.DataContext -is [IndexNode]) {
            $cb.DataContext.Toggle()
            onIndexTreeChecked
        }
    }
})
# フォルダを展開したときに子を読み込む（IsExpanded は OneWay/プレーンなので、ここで expandIndexNode する）
$ui.IndexTree.AddHandler([System.Windows.Controls.TreeViewItem]::ExpandedEvent, [System.Windows.RoutedEventHandler] {
    param ($s, $e)
    safe {
        $node = $e.OriginalSource.DataContext
        if ($node -is [IndexNode]) { expandIndexNode $node }
    }
})
$ui.IndexTree.Add_PreviewKeyDown({
    param ($sender, $e)
    # スペースで選択中のフォルダのチェックを切り替える
    $node = $ui.IndexTree.SelectedItem
    if ($e.Key -eq "Space" -and $node -and !$node.IsPlaceholder) {
        safe {
            $node.Toggle()
            onIndexTreeChecked
        }
        $e.Handled = $true
    }
})
$ui.CheckAllIndexButton.Add_Click({ safe { setAllIndexChecked $true } })
$ui.UncheckAllIndexButton.Add_Click({ safe { setAllIndexChecked $false } })
