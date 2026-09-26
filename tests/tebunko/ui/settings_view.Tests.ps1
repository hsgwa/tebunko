# ［8 設定］の判断（tebunko\ui\settings_view.ps1）のテスト。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\tebunko\ui\settings_view.ps1"

Describe "getWorkspaceView" -Tag Unit {
    It "既定の場所なら既定だと伝え、［既定に戻す］は出さない" {
        $view = getWorkspaceView "C:\tool\work" "C:\Tool\work\"
        $view.Path | Should Be "C:\tool\work"
        $view.Note | Should Be "既定の場所（ドキュメントの tebunko）です。"
        $view.CanReset | Should Be $false
    }

    It "既定でない場所なら、既定の場所を伝え、［既定に戻す］を出す" {
        $view = getWorkspaceView "D:\データ" "C:\tool\work"
        $view.Path | Should Be "D:\データ"
        $view.Note | Should Be "既定の場所は「C:\tool\work」です。"
        $view.CanReset | Should Be $true
    }
}

Describe "getSettingsFileView" -Tag Unit {
    It "ツールのフォルダにあれば、そう伝える" {
        $view = getSettingsFileView "C:\tool\setting.config" "C:\Tool\"
        $view.Path | Should Be "C:\tool\setting.config"
        $view.Note | Should Be "ツールのフォルダに置いています。"
    }

    It "利用者ごとの場所にあれば、ツールのフォルダに書き込めないためだと伝える" {
        $view = getSettingsFileView "C:\Users\test\AppData\Local\tebunko\0123456789ABCDEF\setting.config" "C:\Program Files\tebunko"
        $view.Note | Should Be "ツールのフォルダ（C:\Program Files\tebunko）に書き込めないため、利用者ごとの場所に置いています。"
    }
}

Describe "testWorkspaceChoice" -Tag Unit {
    It "今と同じ場所（書き方の違いも）なら same" {
        (testWorkspaceChoice "C:\TOOL\work\" "C:\tool\work" $true @{}).Kind | Should Be "same"
    }

    It "空なら選ぶよう伝える" {
        $result = testWorkspaceChoice "" "C:\tool\work" $true @{}
        $result.Kind | Should Be "error"
        $result.Message | Should Be "フォルダを選んでください。"
    }

    It "今のインデックスのフォルダの中は選べない" {
        $result = testWorkspaceChoice "C:\tool\work\index\営業" "C:\tool\work" $true @{}
        $result.Kind | Should Be "error"
        $result.Message | Should Match "今のインデックスのフォルダの中です"
    }

    It "書き込めないフォルダは選べない" {
        $result = testWorkspaceChoice "C:\Program Files\共有" "C:\tool\work" $false @{}
        $result.Kind | Should Be "error"
        $result.Message | Should Be "「C:\Program Files\共有」にはファイルを作れません。書き込めるフォルダを選んでください。"
    }

    It "書き込めるほかのフォルダなら ok（今の work の中の、インデックスの外も選べる）" {
        (testWorkspaceChoice "D:\データ" "C:\tool\work" $true @{}).Kind | Should Be "ok"
        (testWorkspaceChoice "C:\tool\work\新しい場所" "C:\tool\work" $true @{}).Kind | Should Be "ok"
    }
}

Describe "newWorkspaceConfirm" -Tag Unit {
    It "空のフォルダなら、そこに置くことと、今のワークスペースの中身を移すことを伝え、変えるボタンを 1 つ出す" {
        $confirm = newWorkspaceConfirm "D:\データ" "C:\tool\work" 0
        $confirm.Heading | Should Be "ワークスペースを変えますか？"
        @($confirm.Facts | ForEach-Object { $_.Kind }) -join "," | Should Be "next,next"
        $confirm.Facts[0].Detail | Should Be "D:\データ"
        $confirm.Facts[1].Title | Should Be "今のワークスペースの中身（インデックス・取り込み一覧・ログ）は、新しいワークスペースへ移します"
        $confirm.Facts[1].Detail | Should Be "移す前の場所：C:\tool\work"
        $confirm.Hint | Should Be ""
        @($confirm.Choices | ForEach-Object { $_.Value }) -join "," | Should Be "change"
        $confirm.Choices[0].Text | Should Be "ワークスペースを変える"
    }

    It "空でなければ警告し、中身の数と例を出し、中に workspace を作るか・そのまま使うかを選ばせる" {
        $confirm = newWorkspaceConfirm "D:\データ" "C:\tool\work" 5 @("見積.xlsx", "報告書", "メモ.txt")
        $confirm.Heading | Should Be "選んだフォルダは空ではありません。ワークスペースには空のフォルダを選んでください。"
        @($confirm.Facts | ForEach-Object { $_.Kind }) -join "," | Should Be "warn,next"
        $confirm.Facts[0].Title | Should Be "このフォルダは空ではありません（ファイル・フォルダが 5 個）"
        $confirm.Facts[0].Detail | Should Be "見積.xlsx、報告書、メモ.txt など"
        @($confirm.Choices | ForEach-Object { $_.Value }) -join "," | Should Be "sub,asis"
        $confirm.Choices[0].Detail | Should Be "D:\データ\workspace"
        $confirm.Choices[1].Careful | Should Be $true
        $confirm.Hint | Should Be "空のフォルダを選び直すときは［キャンセル］を押してください。"
    }

    It "中身がすべて例に出ていれば「など」を付けず、数え切れなければ「以上」と出す" {
        (newWorkspaceConfirm "D:\データ" "C:\tool\work" 1 @("a.txt")).Facts[0].Detail | Should Be "a.txt"
        (newWorkspaceConfirm "D:\データ" "C:\tool\work" 1000 @("a", "b", "c") $true).Facts[0].Title | Should Be "このフォルダは空ではありません（ファイル・フォルダが 1,000 個以上）"
    }

    It "中に workspace を作れなければ、その選択肢を出さない" {
        $confirm = newWorkspaceConfirm "D:\データ" "C:\tool\work" 3 @("a", "b", "c") $false @() $false
        @($confirm.Choices | ForEach-Object { $_.Value }) -join "," | Should Be "asis"
    }

    It "インデックスなどがあれば、使うか、消して最初からやり直すかを選ばせる（消すほうは赤いボタンで、キャンセルを既定にする）" {
        $confirm = newWorkspaceConfirm "D:\共有\tebunko_ws" "C:\tool\work" 3 @("index", "取り込み一覧.tsv", "memo.txt") $false @("index", "取り込み一覧.tsv")
        $confirm.Heading | Should Be "選んだフォルダには、すでにインデックスがあります。どうしますか？"
        @($confirm.Facts | ForEach-Object { $_.Kind }) -join "," | Should Be "kept,next"
        $confirm.Facts[0].Detail | Should Be "index、取り込み一覧.tsv"
        $confirm.Facts[1].Detail | Should Be "今のワークスペースの中身は移さず、元の場所に残します：C:\tool\work"
        @($confirm.Choices | ForEach-Object { $_.Value }) -join "," | Should Be "use,reset"
        @($confirm.Choices | ForEach-Object { $_.Text }) -join "," | Should Be "あるインデックスを使う,消して、最初からやり直す"
        $confirm.Choices[1].Danger | Should Be $true
        $confirm.Choices[1].Careful | Should Be $true
    }
}