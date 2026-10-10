# 設定の画面の判断（tebunko\ui\settings\settings_view.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\settings\settings_view.ps1"
}

Describe "getWorkspaceView" -Tag Unit {
    It "既定の場所なら既定だと伝え、［既定に戻す］は出さない" {
        $view = getWorkspaceView "C:\tool\work" "C:\Tool\work\"
        $view.Path | Should -Be "C:\tool\work"
        $view.Note | Should -Be "既定の場所（ドキュメントの tebunko）です。"
        $view.CanReset | Should -Be $false
    }

    It "既定でない場所なら、既定の場所を伝え、［既定に戻す］を出す" {
        $view = getWorkspaceView "D:\データ" "C:\tool\work"
        $view.Path | Should -Be "D:\データ"
        $view.Note | Should -Be "既定の場所は「C:\tool\work」です。"
        $view.CanReset | Should -Be $true
    }
}

Describe "getSettingsFileView" -Tag Unit {
    It "ツールのフォルダにあれば、そう伝える" {
        $view = getSettingsFileView "C:\tool\setting.config" "C:\Tool\"
        $view.Path | Should -Be "C:\tool\setting.config"
        $view.Note | Should -Be "ツールのフォルダに置いています。"
    }

    It "既定のワークスペースにあれば、ツールのフォルダに書き込めないためだと伝える" {
        $view = getSettingsFileView "C:\Users\test\Documents\tebunko_ws\setting.config" "C:\Program Files\tebunko"
        $view.Note | Should -Be "ツールのフォルダ（C:\Program Files\tebunko）に書き込めないため、既定のワークスペースに置いています。"
    }
}

Describe "getCountedEntryPaths（既定のワークスペースの設定ファイルを数えない）" -Tag Unit {
    It "<Case>" -TestCases @(
        @{ Case = "既定のワークスペースなら、設定ファイルと付いてできるファイルを数えない"; IsDefault = $true
           Paths = @("C:\ws\setting.config", "C:\ws\setting.config.tmp", "C:\ws\setting.config.broken-20261003-120000", "C:\ws\setting.config.broken-20261003-120000-2")
           Expected = @() }
        @{ Case = "既定のワークスペースでも、似た名前とほかのファイルは数える"; IsDefault = $true
           Paths = @("C:\ws\setting.config", "C:\ws\setting.config.bak", "C:\ws\README.md")
           Expected = @("C:\ws\setting.config.bak", "C:\ws\README.md") }
        @{ Case = "既定のワークスペースでないフォルダでは、設定ファイルも数える"; IsDefault = $false
           Paths = @("C:\other\setting.config", "C:\other\a.txt")
           Expected = @("C:\other\setting.config", "C:\other\a.txt") }
    ) {
        @(getCountedEntryPaths $Paths $IsDefault) | Should -Be @($Expected)
    }

    It "既定のワークスペースに設定ファイルだけなら、数えた結果は 0 個で「空でない」の警告を出さない" {
        $counted = @(getCountedEntryPaths @("C:\ws\setting.config", "C:\ws\setting.config.tmp") $true)
        $confirm = newWorkspaceConfirm "C:\ws" "C:\old" $counted.Count @() $false @() $true $false
        @($confirm.Facts).Count | Should -Be 0
        $confirm.Choices[0].Value | Should -Be "change"
    }

    It "既定のワークスペースでないフォルダでは、設定ファイルが 1 つあれば「空でない」の警告を出す" {
        $counted = @(getCountedEntryPaths @("C:\other\setting.config") $false)
        $confirm = newWorkspaceConfirm "C:\other" "C:\old" $counted.Count @("setting.config") $false @() $true $false
        @($confirm.Facts | Where-Object { $_.Kind -eq "warn" }).Count | Should -Be 1
    }
}

Describe "testWorkspaceChoice" -Tag Unit {
    It "今と同じ場所（書き方の違いも）なら same" {
        (testWorkspaceChoice "C:\TOOL\work\" "C:\tool\work" $true @{}).Kind | Should -Be "same"
    }

    It "空なら選ぶよう伝える" {
        $result = testWorkspaceChoice "" "C:\tool\work" $true @{}
        $result.Kind | Should -Be "error"
        $result.Message | Should -Be "フォルダを選んでください。"
    }

    It "今のインデックスのフォルダの中は選べない（<sub>）" -TestCases @(
        @{ sub = "content_index\営業" }
        @{ sub = "index\営業" }  # 前の版の index\ の中も同じ
    ) {
        param ($sub)
        $result = testWorkspaceChoice "C:\tool\work\$sub" "C:\tool\work" $true @{}
        $result.Kind | Should -Be "error"
        $result.Message | Should -Match "今のインデックスのフォルダの中です"
    }

    It "書き込めないフォルダは選べない" {
        $result = testWorkspaceChoice "C:\Program Files\共有" "C:\tool\work" $false @{}
        $result.Kind | Should -Be "error"
        $result.Message | Should -Be "「C:\Program Files\共有」にはファイルを作れません。書き込めるフォルダを選んでください。"
    }

    It "書き込めるほかのフォルダなら ok（今の work の中の、インデックスの外も選べる）" {
        (testWorkspaceChoice "D:\データ" "C:\tool\work" $true @{}).Kind | Should -Be "ok"
        (testWorkspaceChoice "C:\tool\work\新しい場所" "C:\tool\work" $true @{}).Kind | Should -Be "ok"
    }
}

Describe "newWorkspaceConfirm" -Tag Unit {
    It "空のフォルダなら、移す先を見出しに、移動中は使えないことを補足に出し、移動するボタンを 1 つ出す" {
        $confirm = newWorkspaceConfirm "D:\データ" "C:\tool\work" 0
        $confirm.Title | Should -Be "ワークスペースの変更"
        $confirm.Heading | Should -Be "インデックスとログを「D:\データ」へ移動します。"
        @($confirm.Facts).Count | Should -Be 0
        $confirm.Hint | Should -Be "移動中は、検索とインデックスの更新はできません。"
        @($confirm.Choices | ForEach-Object { $_.Value }) -join "," | Should -Be "change"
        $confirm.Choices[0].Text | Should -Be "移動する"
    }

    It "既定に戻すときは、題と文言を戻す向けにする" {
        $confirm = newWorkspaceConfirm "C:\Tools\work" "D:\x" 0 @() $false @() $true $true
        $confirm.Title | Should -Be "既定の場所に戻す"
        $confirm.Heading | Should -Be "インデックスとログを既定の場所「C:\Tools\work」へ移動します。"
        $confirm.Choices[0].Text | Should -Be "戻す"
    }

    It "空でなければ警告し、中身の数と例を出し、中に workspace を作るか・そのまま使うかを選ばせる" {
        $confirm = newWorkspaceConfirm "D:\データ" "C:\tool\work" 5 @("見積.xlsx", "報告書", "メモ.txt")
        $confirm.Heading | Should -Be "選んだフォルダは空ではありません。ワークスペースには空のフォルダを選んでください。"
        @($confirm.Facts | ForEach-Object { $_.Kind }) -join "," | Should -Be "warn,next,next"
        $confirm.Facts[0].Title | Should -Be "このフォルダは空ではありません（ファイル・フォルダが 5 個）"
        $confirm.Facts[0].Detail | Should -Be "見積.xlsx、報告書、メモ.txt など"
        @($confirm.Choices | ForEach-Object { $_.Value }) -join "," | Should -Be "asis,sub"
        $confirm.Facts[2].Detail | Should -Be "D:\データ\workspace"
        $confirm.Choices[0].Careful | Should -Be $true
        $confirm.Hint | Should -Be "空のフォルダを選び直すときは［キャンセル］を押してください。"
    }

    It "中身がすべて例に出ていれば「など」を付けず、数え切れなければ「以上」と出す" {
        (newWorkspaceConfirm "D:\データ" "C:\tool\work" 1 @("a.txt")).Facts[0].Detail | Should -Be "a.txt"
        (newWorkspaceConfirm "D:\データ" "C:\tool\work" 1000 @("a", "b", "c") $true).Facts[0].Title | Should -Be "このフォルダは空ではありません（ファイル・フォルダが 1,000 個以上）"
    }

    It "中に workspace を作れなければ、その選択肢を出さない" {
        $confirm = newWorkspaceConfirm "D:\データ" "C:\tool\work" 3 @("a", "b", "c") $false @() $false
        @($confirm.Choices | ForEach-Object { $_.Value }) -join "," | Should -Be "asis"
    }

    It "インデックスなどがあれば、今のを移すか、そのフォルダのを使うかを選ばせる（移すほうは、そのフォルダのインデックスが消えると補足に書く）" {
        $confirm = newWorkspaceConfirm "D:\共有\tebunko_ws" "C:\tool\work" 3 @("index", "ingest_status.tsv", "memo.txt") $false @("index", "ingest_status.tsv")
        $confirm.Heading | Should -Be "「D:\共有\tebunko_ws」には、すでにインデックスがあります。"
        $confirm.Hint | Should -Match "そのフォルダにあるインデックスは削除されます"
        @($confirm.Choices | ForEach-Object { $_.Value }) -join "," | Should -Be "reset,use"
        @($confirm.Choices | ForEach-Object { $_.Text }) -join "," | Should -Be "今のインデックスを移動する,そのフォルダのインデックスを使う"
        $confirm.Choices[0].Danger | Should -Be $true
        $confirm.Choices[0].Careful | Should -Be $true
    }
}
