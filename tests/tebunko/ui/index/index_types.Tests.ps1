# ［インデックス管理］の一覧・詳細・失敗した行の型（tebunko\ui\index\index_types.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\shared\ui\types.ps1"
    . "${scriptsDir}\tebunko\ui\index\index_types.ps1"

    # PropertyChanged で通知されたプロパティ名を集める
    function watchChanges($target) {
        $names = New-Object System.Collections.Generic.List[string]
        $target.add_PropertyChanged([System.ComponentModel.PropertyChangedEventHandler]{ param($sender, $e) $names.Add($e.PropertyName) }.GetNewClosure())
        , $names
    }
}

Describe "FolderItem" -Tag Unit {
    It "有効・名前・パスは変わったときだけ通知する" {
        $item = [FolderItem]::new()
        $names = watchChanges $item
        $item.SetEnabled($true); $item.SetEnabled($true)
        $item.SetName("営業"); $item.SetName("営業")
        $item.SetPath("C:\共有\営業部"); $item.SetPath("C:\共有\営業部")
        $item.Enabled | Should -Be $true
        $item.Name | Should -Be "営業"
        $item.Path | Should -Be "C:\共有\営業部"
        @($names) | Should -Be @("Enabled", "Name", "Path")
    }

    It "状態・件数をまとめて変えて通知する" {
        $item = [FolderItem]::new()
        $names = watchChanges $item
        $item.SetStatus("作成済み", "Green")
        $item.SetStats("12 件", "Excel 10 件", "2026/09/01")
        $item.StatusText | Should -Be "作成済み"
        $item.StatusBrush | Should -Be "Green"
        $item.FileCountText | Should -Be "12 件"
        $item.FileCountToolTip | Should -Be "Excel 10 件"
        $item.LastIngestedText | Should -Be "2026/09/01"
        @($names) | Should -Be @("StatusText", "StatusBrush", "FileCountText", "FileCountToolTip", "LastIngestedText")
    }

    It "ステータス・高速検索の列をまとめて変えて通知する" {
        $item = [FolderItem]::new()
        $names = watchChanges $item
        $item.SetIndexState("最新", "更新済み 10 件", "Ok", "", 0)
        $item.SetFast("可", "すぐ検索できます", "Ok")
        $item.IndexText | Should -Be "最新"
        $item.IndexToolTip | Should -Be "更新済み 10 件"
        $item.IndexLevel | Should -Be "Ok"
        $item.FastText | Should -Be "可"
        $item.FastToolTip | Should -Be "すぐ検索できます"
        $item.FastLevel | Should -Be "Ok"
        @($names) | Should -Be @("IndexText", "IndexToolTip", "IndexLevel", "IndexSub", "IndexPercent", "FastText", "FastToolTip", "FastLevel", "FastSteps")
    }

    It "高速検索の対処の手順を渡すと FastSteps に入る" {
        $item = [FolderItem]::new()
        $item.SetFast("不可", "対象外", "Warn", [string[]]@("手順 1", "手順 2"))
        $item.FastLevel | Should -Be "Warn"
        @($item.FastSteps) | Should -Be @("手順 1", "手順 2")
        $item.SetFast("可", "", "Ok")
        @($item.FastSteps).Count | Should -Be 0
    }
}
