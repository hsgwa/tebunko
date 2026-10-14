# tebunko の画面で使う型: 検索対象とそのツリー。
# shared\ui\types.ps1 を先に読み込んでおくこと（NotifyBase を継承する）。

# 検索対象の1件（tebunko\search\content_index_search.ps1 の getContentIndexFiles に渡す）
class SearchTarget {
    [string]$Root
    [string]$RelPath
    [bool]$Recurse
}

# 検索対象から外したフォルダ（tebunko\core\settings.ps1 の readSearchExcludes / writeSearchExcludes と同じ項目）
class SearchExclude {
    [string]$Path
    [bool]$Subfolders
}

# 検索対象インデックスのツリーの1項目。3状態チェック（true/false/null）は子・親の状態から決まる。
# チェックはツリーのチェックボックスの Click で Toggle() を呼んで変える（TwoWay バインドはしない）
class IndexNode : NotifyBase {
    [string]$Name
    [string]$Root
    [string]$RelPath
    [bool]$IsFiles
    [bool]$IsPlaceholder
    [bool]$Exists
    [string]$SourcePath
    [string]$ToolTip
    [IndexNode]$Parent
    [System.Collections.ObjectModel.ObservableCollection[IndexNode]]$Children

    [Nullable[bool]]$IsChecked = $true   # チェックボックスは OneWay バインド。変更は Toggle/SetChecked で行う
    [bool]$IsExpanded                    # TreeViewItem.IsExpanded は OneWay。展開は Expanded イベント/SetExpanded で読み込む
    [bool]$IsLoading                     # 裏で子を読み込んでいる間（ネットワークのとき）
    hidden [bool]$loaded

    IndexNode([IndexNode]$parent, [string]$name, [string]$root, [string]$relPath, [bool]$isFiles) {
        $this.Parent = $parent
        $this.Name = $name
        $this.Root = $root
        $this.RelPath = $relPath
        $this.IsFiles = $isFiles
        $this.Exists = $true
        $this.Children = New-Object System.Collections.ObjectModel.ObservableCollection[IndexNode]
        if ($null -ne $parent) { $this.IsChecked = ($parent.IsChecked -ne $false) }
    }

    # 根（インデックス 1 件）を作る。フォルダの有無・サブフォルダの有無は呼ぶ側が集めて渡す（ここではファイルに触らない）
    static [IndexNode] CreateRoot([string]$root, [string]$name, [string]$relPath, [string]$sourcePath, [bool]$exists, [bool]$hasSubfolders) {
        $node = [IndexNode]::new($null, $name, $root, $relPath, $false)
        $node.SourcePath = $sourcePath
        $dir = $node.FullPath()
        $node.Exists = $exists
        # [string] の引数・プロパティは $null を "" にするため、元のフォルダが分からないことは空で判定する
        $node.ToolTip = $(if (-not [string]::IsNullOrEmpty($sourcePath)) { "元のフォルダ：" + $sourcePath + "`nインデックス：" + $dir } else { $dir })
        if (-not $node.Exists) { $node.ToolTip += "`n（フォルダが見つかりません。検索時はスキップします）" }
        if ($node.Exists -and $hasSubfolders) { $node.Children.Add([IndexNode]::NewPlaceholder($node)) } else { $node.loaded = $true }
        return $node
    }

    [string] FullPath() {
        return $(if ($this.RelPath -eq "") { $this.Root } else { $this.Root.TrimEnd('\') + "\" + $this.RelPath })
    }

    # 展開の状態を変える（プログラムから）。子の読み込みは呼ぶ側が行う（NeedsLoad・ApplyChildren）
    [void] SetExpanded([bool]$value) {
        if ($this.IsExpanded -eq $value) { return }
        $this.IsExpanded = $value
        $this.Raise("IsExpanded")
    }

    [void] SetChecked([bool]$value) {
        $this.SetTree($value)
        if ($null -ne $this.Parent) { $this.Parent.UpdateFromChildren() }
    }

    [void] Toggle() { $this.SetChecked($this.IsChecked -ne $true) }

    hidden [void] SetTree([bool]$value) {
        $this.SetState($value)
        foreach ($child in $this.Children) { if (-not $child.IsPlaceholder) { $child.SetTree($value) } }
    }

    hidden [void] SetState([Nullable[bool]]$value) {
        if ($this.IsChecked -eq $value) { return }
        $this.IsChecked = $value
        $this.Raise("IsChecked")
    }

    hidden [void] UpdateFromChildren() {
        $any = $false
        [Nullable[bool]]$state = $null
        foreach ($child in $this.Children) {
            if ($child.IsPlaceholder) { continue }
            if (-not $any) { $state = $child.IsChecked; $any = $true }
            elseif ($state -ne $child.IsChecked) { $state = $null; break }
        }
        if (-not $any) { return }
        $this.SetState($state)
        if ($null -ne $this.Parent) { $this.Parent.UpdateFromChildren() }
    }

    # 子をまだ読み込んでいないか（読み込み前の子はプレースホルダーだけ）。読み込み中（IsLoading）の間も読み込み前
    [bool] NeedsLoad() {
        return (-not $this.loaded -and -not $this.IsFiles -and -not $this.IsPlaceholder -and $this.Exists)
    }

    # 読み込んだ子（getIndexFolderChildren の結果）を入れる。1 回だけ入れる（ここではファイルに触らない）
    [void] ApplyChildren($data) {
        if ($this.loaded -or $this.IsFiles -or $this.IsPlaceholder) { return }
        $this.loaded = $true
        $this.IsLoading = $false
        $this.Children.Clear()
        if (-not $this.Exists) { return }
        $dir = $this.FullPath()
        $names = @($data.Folders)
        if ($names.Count -gt 0 -and $data.HasFiles) {
            $files = [IndexNode]::new($this, "（このフォルダ直下のファイル）", $this.Root, $this.RelPath, $true)
            $files.ToolTip = "サブフォルダを除く、" + $(if (-not [string]::IsNullOrEmpty($this.SourcePath)) { $this.SourcePath } else { $dir }) + " の直下のファイル"
            $this.Children.Add($files)
        }
        foreach ($folder in $names) {
            $n = [string]$folder.Name
            $childRel = $(if ($this.RelPath -eq "") { $n } else { $this.RelPath + "\" + $n })
            $child = [IndexNode]::new($this, $n, $this.Root, $childRel, $false)
            if (-not [string]::IsNullOrEmpty($this.SourcePath)) { $child.SourcePath = $this.SourcePath.TrimEnd('\') + "\" + $n }
            $child.ToolTip = $(if (-not [string]::IsNullOrEmpty($child.SourcePath)) { "元のフォルダ：" + $child.SourcePath } else { $child.FullPath() })
            if ($folder.HasSubfolders) { $child.Children.Add([IndexNode]::NewPlaceholder($child)) } else { $child.loaded = $true }
            $this.Children.Add($child)
        }
    }

    [IndexNode] Find([string]$path) {
        if ($this.IsFiles -or $this.IsPlaceholder) { return $null }
        $full = $this.FullPath().TrimEnd('\')
        $path = $path.TrimEnd('\')
        if ([string]::Equals($path, $full, [System.StringComparison]::OrdinalIgnoreCase)) { return $this }
        if (-not $path.StartsWith($full + "\", [System.StringComparison]::OrdinalIgnoreCase)) { return $null }
        # 読み込み前の子はたどらない（ファイルに触らない。読み込みは呼ぶ側が先に済ませる）
        foreach ($child in $this.Children) {
            $found = $child.Find($path)
            if ($null -ne $found) { return $found }
        }
        return $null
    }

    [void] ApplyExclude([string]$path, [bool]$subfolders) {
        $node = $this.Find($path)
        if ($null -eq $node) { return }
        if (-not $subfolders) {
            foreach ($child in $node.Children) {
                if ($child.IsFiles) { $child.SetChecked($false); return }
            }
            if ($node.Children.Count -gt 0) { return }
        }
        $node.SetChecked($false)
    }

    [void] AddTargets([System.Collections.Generic.List[SearchTarget]]$targets) {
        if ($this.IsPlaceholder -or $this.IsChecked -eq $false) { return }
        if ($this.IsChecked -eq $true) {
            $targets.Add([SearchTarget]@{ Root = $this.Root; RelPath = $this.RelPath; Recurse = (-not $this.IsFiles) })
            return
        }
        foreach ($child in $this.Children) { $child.AddTargets($targets) }
    }

    [void] AddExcludes([System.Collections.Generic.List[SearchExclude]]$excludes) {
        if ($this.IsPlaceholder -or $this.IsChecked -eq $true) { return }
        if ($this.IsChecked -eq $false) {
            $excludes.Add([SearchExclude]@{ Path = $this.FullPath(); Subfolders = (-not $this.IsFiles) })
            return
        }
        foreach ($child in $this.Children) { $child.AddExcludes($excludes) }
    }

    [void] AddExpanded([System.Collections.Generic.List[string]]$paths) {
        if ($this.IsFiles -or $this.IsPlaceholder) { return }
        if ($this.IsExpanded) { $paths.Add($this.FullPath()) }
        foreach ($child in $this.Children) { $child.AddExpanded($paths) }
    }

    static [IndexNode] NewPlaceholder([IndexNode]$parent) {
        $node = [IndexNode]::new($parent, "読み込み中…", $parent.Root, $parent.RelPath, $false)
        $node.IsPlaceholder = $true
        return $node
    }
}
