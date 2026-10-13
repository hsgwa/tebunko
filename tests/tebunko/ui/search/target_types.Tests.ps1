# 検索対象と、そのツリーの型（tebunko\ui\search\target_types.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\shared\ui\types.ps1"
    . "${scriptsDir}\tebunko\ui\search\target_types.ps1"

    # PropertyChanged で通知されたプロパティ名を集める
    function watchChanges($target) {
        $names = New-Object System.Collections.Generic.List[string]
        $target.add_PropertyChanged([System.ComponentModel.PropertyChangedEventHandler]{ param($sender, $e) $names.Add($e.PropertyName) }.GetNewClosure())
        , $names
    }
}

Describe "IndexNode（チェック）" -Tag Unit {
    BeforeAll {
        # 根 ─ 営業部 ─ 東京・大阪、総務部（フォルダは読まない。子を手で足す）
        function newTree {
            $root = [IndexNode]::new($null, "営業", "C:\index\営業", "", $false)
            $sales = [IndexNode]::new($root, "営業部", $root.Root, "営業部", $false)
            $tokyo = [IndexNode]::new($sales, "東京", $root.Root, "営業部\東京", $false)
            $osaka = [IndexNode]::new($sales, "大阪", $root.Root, "営業部\大阪", $false)
            $general = [IndexNode]::new($root, "総務部", $root.Root, "総務部", $false)
            $sales.Children.Add($tokyo); $sales.Children.Add($osaka); $sales.Children.Add([IndexNode]::NewPlaceholder($sales))
            $root.Children.Add($sales); $root.Children.Add($general)
            @{ Root = $root; Sales = $sales; Tokyo = $tokyo; Osaka = $osaka; General = $general }
        }
    }

    It "作ったときは親のチェックを受け継ぐ" {
        $parent = [IndexNode]::new($null, "営業", "C:\index\営業", "", $false)
        [IndexNode]::new($parent, "a", "C:\index\営業", "a", $false).IsChecked | Should -Be $true
        $parent.IsChecked = $false
        [IndexNode]::new($parent, "b", "C:\index\営業", "b", $false).IsChecked | Should -Be $false
        $parent.IsChecked = $null
        [IndexNode]::new($parent, "c", "C:\index\営業", "c", $false).IsChecked | Should -Be $true
    }

    It "外すと子も外れ、親は一部だけの状態（null）になる" {
        $t = newTree
        $t.Sales.SetChecked($false)
        $t.Tokyo.IsChecked | Should -Be $false
        $t.Osaka.IsChecked | Should -Be $false
        $t.Root.IsChecked | Should -Be $null
        $t.General.IsChecked | Should -Be $true
    }

    It "子の 1 つを外すと、親と根は一部だけになる" {
        $t = newTree
        $t.Tokyo.Toggle()
        $t.Tokyo.IsChecked | Should -Be $false
        $t.Sales.IsChecked | Should -Be $null
        $t.Root.IsChecked | Should -Be $null
    }

    It "子をすべて外すと親も外れる" {
        $t = newTree
        $t.Tokyo.SetChecked($false)
        $t.Osaka.SetChecked($false)
        $t.Sales.IsChecked | Should -Be $false
        $t.General.SetChecked($false)
        $t.Root.IsChecked | Should -Be $false
    }

    It "一部だけの状態で切り替えるとすべて付く" {
        $t = newTree
        $t.Tokyo.SetChecked($false)
        $t.Sales.Toggle()
        $t.Sales.IsChecked | Should -Be $true
        $t.Tokyo.IsChecked | Should -Be $true
        $t.Root.IsChecked | Should -Be $true
    }

    It "チェックの変化を通知し、変わらなければ通知しない" {
        $t = newTree
        $names = watchChanges $t.Tokyo
        $t.Tokyo.SetChecked($true)
        $names.Count | Should -Be 0
        $t.Tokyo.SetChecked($false)
        @($names) | Should -Be @("IsChecked")
    }

    It "検索対象は、すべて付いた枝だけを 1 件にまとめる" {
        $t = newTree
        $t.Tokyo.SetChecked($false)
        $targets = [System.Collections.Generic.List[SearchTarget]]::new()
        $t.Root.AddTargets($targets)
        @($targets | ForEach-Object { "$($_.RelPath)|$($_.Recurse)" }) | Should -Be @("営業部\大阪|True", "総務部|True")
    }

    It "すべて付いていれば根だけを対象にする" {
        $t = newTree
        $targets = [System.Collections.Generic.List[SearchTarget]]::new()
        $t.Root.AddTargets($targets)
        $targets.Count | Should -Be 1
        $targets[0].Root | Should -Be "C:\index\営業"
        $targets[0].RelPath | Should -Be ""
    }

    It "直下のファイルの項目は、サブフォルダを含めない対象にする" {
        $root = [IndexNode]::new($null, "営業", "C:\index\営業", "", $false)
        $files = [IndexNode]::new($root, "（このフォルダ直下のファイル）", $root.Root, "", $true)
        $sub = [IndexNode]::new($root, "営業部", $root.Root, "営業部", $false)
        $root.Children.Add($files); $root.Children.Add($sub)
        $sub.SetChecked($false)
        $targets = [System.Collections.Generic.List[SearchTarget]]::new()
        $root.AddTargets($targets)
        $targets.Count | Should -Be 1
        $targets[0].Recurse | Should -Be $false
    }

    It "外したものは、外した枝ごとに記録する" {
        $t = newTree
        $t.Tokyo.SetChecked($false)
        $t.General.SetChecked($false)
        $excludes = [System.Collections.Generic.List[SearchExclude]]::new()
        $t.Root.AddExcludes($excludes)
        @($excludes | ForEach-Object { "$($_.Path)|$($_.Subfolders)" }) | Should -Be @("C:\index\営業\営業部\東京|True", "C:\index\営業\総務部|True")
    }

    It "すべて付いていれば外したものは無い" {
        $t = newTree
        $excludes = [System.Collections.Generic.List[SearchExclude]]::new()
        $t.Root.AddExcludes($excludes)
        $excludes.Count | Should -Be 0
    }

    It "展開しているフォルダのパスを集める" {
        $t = newTree
        $t.Root.IsExpanded = $true
        $t.Sales.IsExpanded = $true
        $paths = [System.Collections.Generic.List[string]]::new()
        $t.Root.AddExpanded($paths)
        @($paths) | Should -Be @("C:\index\営業", "C:\index\営業\営業部")
    }

    It "FullPath は根ならそのまま、下なら相対パスをつなぐ" {
        $t = newTree
        $t.Root.FullPath() | Should -Be "C:\index\営業"
        $t.Tokyo.FullPath() | Should -Be "C:\index\営業\営業部\東京"
    }
}

Describe "IndexNode（子の受け取り）" -Tag Unit {
    # ファイルには触らない。子の材料（getIndexFolderChildren の結果の形）は、ここで手で作る。
    #   営業\直下のファイル
    #   営業\営業部\東京\（営業部の下にサブフォルダがある）
    #   営業\総務部\
    BeforeAll {
        function newChildrenData([bool]$hasFiles, [object[]]$folders) {
            return @{ HasFiles = $hasFiles; Folders = @($folders); Error = "" }
        }
        function newFolderData([string]$name, [bool]$hasSubfolders) {
            return @{ Name = $name; HasSubfolders = $hasSubfolders }
        }
        function newSalesRoot([string]$sourcePath = "C:\共有\営業") {
            return [IndexNode]::CreateRoot("C:\index\営業", "営業", "", $sourcePath, $true, $true)
        }
        function newLoadedRoot([string]$sourcePath = $null) {
            # 直下のファイルあり・営業部（東京あり）・総務部を読み込み済みの根
            $root = newSalesRoot $sourcePath
            $root.ApplyChildren((newChildrenData $true @((newFolderData "営業部" $true), (newFolderData "総務部" $false))))
            return $root
        }
    }

    It "根を作ると、サブフォルダがあれば読み込み中の子を置く" {
        $root = newSalesRoot
        $root.Exists | Should -Be $true
        $root.SourcePath | Should -Be "C:\共有\営業"
        $root.ToolTip | Should -Be "元のフォルダ：C:\共有\営業`nインデックス：C:\index\営業"
        $root.Children.Count | Should -Be 1
        $root.Children[0].IsPlaceholder | Should -Be $true
        $root.NeedsLoad() | Should -Be $true
    }

    It "サブフォルダが無ければ、子を置かず読み込み済みにする" {
        $root = [IndexNode]::CreateRoot("C:\index\営業", "営業", "", $null, $true, $false)
        $root.Children.Count | Should -Be 0
        $root.NeedsLoad() | Should -Be $false
    }

    # 引数 [string]$sourcePath は $null を "" にするため、$null だけを見ると「元のフォルダ：」が空のまま出ていた
    It "元のフォルダが分からなければ、ツールチップはインデックスの場所" {
        $root = [IndexNode]::CreateRoot("C:\index\営業", "営業", "", $null, $true, $true)
        $root.ToolTip | Should -Be "C:\index\営業"
    }

    It "フォルダが無ければ、その旨をツールチップに出し、子を読まない" {
        $root = [IndexNode]::CreateRoot("C:\index\無い", "無い", "", $null, $false, $true)
        $root.Exists | Should -Be $false
        $root.ToolTip | Should -Match "フォルダが見つかりません"
        $root.Children.Count | Should -Be 0
        $root.NeedsLoad() | Should -Be $false
        $root.ApplyChildren((newChildrenData $true @((newFolderData "営業部" $false))))
        $root.Children.Count | Should -Be 0
    }

    It "展開の状態を変えても、子は読み込まない（呼ぶ側が ApplyChildren で入れる）" {
        $root = newSalesRoot
        $names = watchChanges $root
        $root.SetExpanded($true)
        $root.IsExpanded | Should -Be $true
        @($names) | Should -Be @("IsExpanded")
        $root.Children.Count | Should -Be 1
        $root.Children[0].IsPlaceholder | Should -Be $true
    }

    It "子を受け取ると、直下のファイルの項目とサブフォルダを入れる" {
        $root = newLoadedRoot "C:\共有\営業"
        $root.NeedsLoad() | Should -Be $false
        @($root.Children | ForEach-Object { $_.Name }) | Should -Be @("（このフォルダ直下のファイル）", "営業部", "総務部")
        $root.Children[0].IsFiles | Should -Be $true
        $root.Children[0].ToolTip | Should -Be "サブフォルダを除く、C:\共有\営業 の直下のファイル"
        $sales = $root.Children[1]
        $sales.RelPath | Should -Be "営業部"
        $sales.SourcePath | Should -Be "C:\共有\営業\営業部"
        $sales.ToolTip | Should -Be "元のフォルダ：C:\共有\営業\営業部"
        # 営業部にはサブフォルダがあるので読み込み中の子を置く。総務部には置かない
        $sales.Children.Count | Should -Be 1
        $sales.Children[0].IsPlaceholder | Should -Be $true
        $sales.NeedsLoad() | Should -Be $true
        $root.Children[2].Children.Count | Should -Be 0
        $root.Children[2].NeedsLoad() | Should -Be $false
    }

    It "2 回目に受け取った子は入れない" {
        $root = newLoadedRoot
        $root.ApplyChildren((newChildrenData $false @((newFolderData "経理部" $false))))
        @($root.Children | ForEach-Object { $_.Name }) -contains "経理部" | Should -Be $false
    }

    It "元のフォルダが分からなければ、子のツールチップはインデックスの場所" {
        $root = newLoadedRoot
        $root.Children[1].SourcePath | Should -BeNullOrEmpty
        $root.Children[0].ToolTip | Should -Be "サブフォルダを除く、C:\index\営業 の直下のファイル"
        $root.Children[1].ToolTip | Should -Be "C:\index\営業\営業部"
    }

    It "サブフォルダがあっても直下にファイルが無ければ、直下のファイルの項目を作らない" {
        $root = newSalesRoot
        $root.ApplyChildren((newChildrenData $false @((newFolderData "営業部" $true), (newFolderData "総務部" $false))))
        @($root.Children | ForEach-Object { $_.Name }) | Should -Be @("営業部", "総務部")
    }

    It "読み込み中の印が付いている間も、まだ読み込んでいない" {
        $root = newSalesRoot
        $root.IsLoading = $true
        $root.NeedsLoad() | Should -Be $true
        $root.ApplyChildren((newChildrenData $false @((newFolderData "総務部" $false))))
        $root.IsLoading | Should -Be $false
        $root.NeedsLoad() | Should -Be $false
    }

    It "直下のファイルの項目・読み込み中の子は、子を受け取らない" {
        $files = [IndexNode]::new($null, "直下", "C:\index\営業", "", $true)
        $files.ApplyChildren((newChildrenData $true @((newFolderData "営業部" $false))))
        $files.Children.Count | Should -Be 0
        $files.NeedsLoad() | Should -Be $false
        $placeholder = [IndexNode]::NewPlaceholder((newSalesRoot))
        $placeholder.NeedsLoad() | Should -Be $false
    }

    It "Find はパスでたどる（読み込み前の子はたどらない）" {
        $root = newLoadedRoot
        $root.Find("C:\index\営業\営業部").RelPath | Should -Be "営業部"
        $root.Find("C:\index\営業") | Should -Be $root
        # 営業部の下はまだ読み込んでいないので、東京は見つからない
        $root.Find("C:\index\営業\営業部\東京") | Should -Be $null
        $root.Find("C:\index\営業\無い") | Should -Be $null
        $root.Find("C:\ほか") | Should -Be $null
        $root.Find("C:\index\営業\営業部").ApplyChildren((newChildrenData $false @((newFolderData "東京" $false))))
        $root.Find("C:\index\営業\営業部\東京\").RelPath | Should -Be "営業部\東京"
    }

    It "ApplyExclude でサブフォルダごと外す" {
        $root = newLoadedRoot
        $root.ApplyExclude("C:\index\営業\営業部", $true)
        $root.Find("C:\index\営業\営業部").IsChecked | Should -Be $false
        $root.IsChecked | Should -Be $null
    }

    It "ApplyExclude でサブフォルダを除く（直下のファイルだけ外す）" {
        $root = newLoadedRoot
        $root.ApplyExclude("C:\index\営業", $false)
        $root.Children[0].IsFiles | Should -Be $true
        $root.Children[0].IsChecked | Should -Be $false
        $root.Children[1].IsChecked | Should -Be $true
        $root.IsChecked | Should -Be $null
    }

    It "ApplyExclude でサブフォルダを除くとき、子が無いフォルダはそのものを外す" {
        $root = newLoadedRoot
        $root.ApplyExclude("C:\index\営業\総務部", $false)
        $root.Find("C:\index\営業\総務部").IsChecked | Should -Be $false
    }

    It "ApplyExclude でサブフォルダを除くとき、直下のファイルの項目が無くサブフォルダがあれば何もしない" {
        $root = newLoadedRoot
        $root.Children[2].Children.Add([IndexNode]::new($root.Children[2], "人事", "C:\index\営業", "総務部\人事", $false))
        $root.ApplyExclude("C:\index\営業\総務部", $false)
        $root.Find("C:\index\営業\総務部").IsChecked | Should -Be $true
        $root.IsChecked | Should -Be $true
    }

    It "ApplyExclude でサブフォルダを除くとき、直下のファイルの項目があればそれだけ外す" {
        $root = newLoadedRoot
        $sales = $root.Find("C:\index\営業\営業部")
        $sales.ApplyChildren((newChildrenData $true @((newFolderData "東京" $false))))
        $root.ApplyExclude("C:\index\営業\営業部", $false)
        $sales.Children[0].IsFiles | Should -Be $true
        $sales.Children[0].IsChecked | Should -Be $false
        $sales.IsChecked | Should -Be $null
    }

    It "ApplyExclude で見つからないパスは無視する" {
        $root = newLoadedRoot
        $root.ApplyExclude("C:\ほか", $true)
        $root.IsChecked | Should -Be $true
    }
}

Describe "IndexNode（境界値）" -Tag Unit {
    It "読み込み中の子を変えても親は変わらない" {
        $root = [IndexNode]::new($null, "営業", "C:\index\営業", "", $false)
        $placeholder = [IndexNode]::NewPlaceholder($root)
        $root.Children.Add($placeholder)
        $placeholder.SetChecked($false)
        $root.IsChecked | Should -Be $true
    }

    It "読み込み中の子は検索対象・除外・展開に入れない" {
        $root = [IndexNode]::new($null, "営業", "C:\index\営業", "", $false)
        $placeholder = [IndexNode]::NewPlaceholder($root)
        $placeholder.IsChecked = $null
        $placeholder.IsExpanded = $true
        $targets = [System.Collections.Generic.List[SearchTarget]]::new()
        $excludes = [System.Collections.Generic.List[SearchExclude]]::new()
        $paths = [System.Collections.Generic.List[string]]::new()
        $placeholder.AddTargets($targets)
        $placeholder.AddExcludes($excludes)
        $placeholder.AddExpanded($paths)
        $targets.Count | Should -Be 0
        $excludes.Count | Should -Be 0
        $paths.Count | Should -Be 0
        $placeholder.Find("C:\index\営業") | Should -Be $null
    }

    It "直下のファイルの項目は、外すとサブフォルダを含めない除外にし、展開には入れない" {
        $root = [IndexNode]::new($null, "営業", "C:\index\営業", "", $false)
        $files = [IndexNode]::new($root, "（このフォルダ直下のファイル）", $root.Root, "", $true)
        $root.Children.Add($files)
        $files.IsExpanded = $true
        $files.SetChecked($false)
        $excludes = [System.Collections.Generic.List[SearchExclude]]::new()
        $files.AddExcludes($excludes)
        $excludes[0].Path | Should -Be "C:\index\営業"
        $excludes[0].Subfolders | Should -Be $false
        $paths = [System.Collections.Generic.List[string]]::new()
        $files.AddExpanded($paths)
        $paths.Count | Should -Be 0
        $files.Find("C:\index\営業") | Should -Be $null
    }

    It "Find はパスの大文字小文字・末尾の \ を区別しない" {
        $root = [IndexNode]::new($null, "営業", "C:\Index\Sales", "", $false)
        $root.Find("c:\index\sales\") | Should -Be $root
    }

    It "Find は名前の先頭が同じだけの別のフォルダを選ばない" {
        $root = [IndexNode]::new($null, "営業", "C:\index\営業", "", $false)
        $root.Find("C:\index\営業部") | Should -Be $null
    }

    It "根の末尾に \ があっても、子のパスは \ を重ねない" {
        [IndexNode]::new($null, "a", "C:\index\", "営業部", $false).FullPath() | Should -Be "C:\index\営業部"
    }

    It "閉じるときは通知し、閉じたままなら通知しない" {
        $node = [IndexNode]::new($null, "営業", "C:\index\無い", "", $true)
        $names = watchChanges $node
        $node.SetExpanded($false)
        $names.Count | Should -Be 0
        $node.SetExpanded($true)
        $node.SetExpanded($false)
        @($names) | Should -Be @("IsExpanded", "IsExpanded")
    }
}
