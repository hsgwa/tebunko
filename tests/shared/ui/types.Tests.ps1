# 画面で使う共通の型（shared\ui\types.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\shared\ui\types.ps1"

    # PropertyChanged で通知されたプロパティ名を集める
    function watchChanges($target) {
        $names = New-Object System.Collections.Generic.List[string]
        $target.add_PropertyChanged([System.ComponentModel.PropertyChangedEventHandler]{ param($sender, $e) $names.Add($e.PropertyName) }.GetNewClosure())
        , $names
    }
}

Describe "NotifyBase" -Tag Unit {
    It "登録したハンドラーに変わったプロパティ名を渡す" {
        $node = [NotifyBase]::new()
        $names = watchChanges $node
        $node.Raise("Name")
        @($names) | Should -Be @("Name")
    }

    It "ハンドラーを外すと通知しない" {
        $node = [NotifyBase]::new()
        $names = New-Object System.Collections.Generic.List[string]
        $handler = [System.ComponentModel.PropertyChangedEventHandler]{ param($sender, $e) $names.Add($e.PropertyName) }.GetNewClosure()
        $node.add_PropertyChanged($handler)
        $node.remove_PropertyChanged($handler)
        $node.Raise("Name")
        $names.Count | Should -Be 0
    }
}
