# 画面で使う共通の型（shared\ui\types.ps1）のテスト。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\shared\ui\types.ps1"

# PropertyChanged で通知されたプロパティ名を集める
function watchChanges($target) {
    $names = New-Object System.Collections.Generic.List[string]
    $target.add_PropertyChanged([System.ComponentModel.PropertyChangedEventHandler]{ param($sender, $e) $names.Add($e.PropertyName) }.GetNewClosure())
    , $names
}

Describe "NotifyBase" -Tag Unit {
    It "登録したハンドラーに変わったプロパティ名を渡す" {
        $node = [FolderNode]::new($null, "営業部", "C:\共有\営業部")
        $names = watchChanges $node
        $node.Raise("Name")
        @($names) | Should Be @("Name")
    }

    It "ハンドラーを外すと通知しない" {
        $node = [FolderNode]::new($null, "営業部", "C:\共有\営業部")
        $names = New-Object System.Collections.Generic.List[string]
        $handler = [System.ComponentModel.PropertyChangedEventHandler]{ param($sender, $e) $names.Add($e.PropertyName) }.GetNewClosure()
        $node.add_PropertyChanged($handler)
        $node.remove_PropertyChanged($handler)
        $node.Raise("Name")
        $names.Count | Should Be 0
    }

    It "ハンドラーが無くても失敗しない" {
        $node = [FolderNode]::new($null, "営業部", "C:\共有\営業部")
        { $node.Raise("Name") } | Should Not Throw
    }
}

Describe "FolderNode" -Tag Unit {
    It "親・名前・パスを持ち、ツールチップはパス" {
        $parent = [FolderNode]::new($null, "共有", "C:\共有")
        $node = [FolderNode]::new($parent, "営業部", "C:\共有\営業部")
        $node.Parent | Should Be $parent
        $node.Name | Should Be "営業部"
        $node.ToolTip | Should Be "C:\共有\営業部"
        $node.Children.Count | Should Be 0
    }

    It "展開・選択は変わったときだけ通知する" {
        $node = [FolderNode]::new($null, "営業部", "C:\共有\営業部")
        $names = watchChanges $node
        $node.SetExpanded($true)
        $node.SetExpanded($true)
        $node.SetSelected($true)
        $node.SetSelected($true)
        $node.SetExpanded($false)
        $node.IsExpanded | Should Be $false
        $node.IsSelected | Should Be $true
        @($names) | Should Be @("IsExpanded", "IsSelected", "IsExpanded")
    }

    It "読み込み中の子を足す" {
        $node = [FolderNode]::new($null, "営業部", "C:\共有\営業部")
        $node.AddPlaceholder()
        $node.Children.Count | Should Be 1
        $node.Children[0].IsPlaceholder | Should Be $true
        $node.Children[0].Name | Should Be "読み込み中…"
        $node.Children[0].Path | Should Be ""
        $node.Children[0].Parent | Should Be $node
    }

    It "文字列にするとフォルダ名" {
        "$([FolderNode]::new($null, "営業部", "C:\共有\営業部"))" | Should Be "営業部"
    }
}

Describe "FolderEntry" -Tag Unit {
    It "文字列にすると名前" {
        $entry = [FolderEntry]@{ Name = "見積.xlsx"; Path = "C:\共有\見積.xlsx"; IsOffice = $true }
        $entry.ToString() | Should Be "見積.xlsx"
    }
}

Describe "ConfirmFact" -Tag Unit {
    It "詳細の既定は空" {
        ([ConfirmFact]@{ Mark = "✓"; Title = "インデックスは残ります" }).Detail | Should Be ""
    }
}
