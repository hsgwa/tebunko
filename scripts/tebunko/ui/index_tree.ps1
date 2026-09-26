# ［2 検索］タブの、検索対象インデックスのツリー。


# ---- 検索対象インデックスのツリー ----

$script:indexRoots = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$ui.IndexTree.ItemsSource = $script:indexRoots

function loadIndexTree {
    # インデックスの一覧（getSearchIndexes）をツリーに読み込む。一番上の項目がインデックス 1 件で、
    # ［1 インデックス管理］で作ったインデックスがすべて並ぶ。
    # 保存したチェックなしのフォルダと、読み込み前の展開の状態は戻す
    $expanded = New-Object 'System.Collections.Generic.List[string]'
    foreach ($node in $script:indexRoots) {
        $node.AddExpanded($expanded)
    }

    $script:indexRoots.Clear()
    $root = ${indexDir}
    if (Test-Path -LiteralPath $root -PathType Container) {
        $root = (Resolve-Path -LiteralPath $root).ProviderPath.TrimEnd("\")
    }
    foreach ($index in @(getSearchIndexes)) {
        $sourcePath = if ($index.SourcePath) { $index.SourcePath } else { $null }
        $script:indexRoots.Add([IndexNode]::CreateRoot($root, $index.Name, $index.Name, $sourcePath))
    }
    $ui.IndexTreePlaceholder.Visibility = if ($script:indexRoots.Count -eq 0) { "Visible" } else { "Collapsed" }

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
# フォルダを展開したときに子を読み込む（IsExpanded は OneWay/プレーンなので、ここで LoadChildren する）
$ui.IndexTree.AddHandler([System.Windows.Controls.TreeViewItem]::ExpandedEvent, [System.Windows.RoutedEventHandler] {
    param ($s, $e)
    safe {
        $node = $e.OriginalSource.DataContext
        if ($node -is [IndexNode]) { $node.LoadChildren() }
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
