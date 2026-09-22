# ［8 設定］の判断（tebunko_grep\ui\settings_view.ps1）のテスト。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\tebunko_grep\ui\settings_view.ps1"

Describe "getWorkDirView" -Tag Unit {
    It "既定の場所なら既定だと伝え、［既定に戻す］は出さない" {
        $view = getWorkDirView "C:\tool\work" "C:\Tool\work\"
        $view.Path | Should Be "C:\tool\work"
        $view.Note | Should Be "既定の場所（設定ファイルと同じフォルダの work）です。"
        $view.CanReset | Should Be $false
    }

    It "既定でない場所なら、既定の場所を伝え、［既定に戻す］を出す" {
        $view = getWorkDirView "D:\データ" "C:\tool\work"
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

Describe "testWorkFolderChoice" -Tag Unit {
    It "今と同じ場所（書き方の違いも）なら same" {
        (testWorkFolderChoice "C:\TOOL\work\" "C:\tool\work" $true @{}).Kind | Should Be "same"
    }

    It "空なら選ぶよう伝える" {
        $result = testWorkFolderChoice "" "C:\tool\work" $true @{}
        $result.Kind | Should Be "error"
        $result.Message | Should Be "フォルダを選んでください。"
    }

    It "今のインデックスのフォルダの中は選べない" {
        $result = testWorkFolderChoice "C:\tool\work\index\営業" "C:\tool\work" $true @{}
        $result.Kind | Should Be "error"
        $result.Message | Should Match "今のインデックスのフォルダの中です"
    }

    It "書き込めないフォルダは選べない" {
        $result = testWorkFolderChoice "C:\Program Files\共有" "C:\tool\work" $false @{}
        $result.Kind | Should Be "error"
        $result.Message | Should Be "「C:\Program Files\共有」にはファイルを作れません。書き込めるフォルダを選んでください。"
    }

    It "書き込めるほかのフォルダなら ok（今の work の中の、インデックスの外も選べる）" {
        (testWorkFolderChoice "D:\データ" "C:\tool\work" $true @{}).Kind | Should Be "ok"
        (testWorkFolderChoice "C:\tool\work\新しい場所" "C:\tool\work" $true @{}).Kind | Should Be "ok"
    }
}

Describe "newWorkFolderConfirm" -Tag Unit {
    It "新しい場所にインデックスが無ければ、取り込み直すことと、今のインデックスは残ることを伝える" {
        $confirm = newWorkFolderConfirm "D:\データ" "C:\tool\work" $false
        @($confirm.Facts | ForEach-Object { $_.Kind }) -join "," | Should Be "next,next,kept"
        $confirm.Facts[0].Detail | Should Be "D:\データ"
        $confirm.Facts[1].Title | Should Be "このフォルダにはまだインデックスがありません"
        $confirm.Facts[2].Detail | Should Be "C:\tool\work"
        $confirm.Hint | Should Match "「C:\\tool\\work」の中身を新しいフォルダへ移して"
    }

    It "新しい場所にインデックスがあれば、そのまま使うことを伝える" {
        $confirm = newWorkFolderConfirm "D:\データ" "C:\tool\work" $true
        @($confirm.Facts | ForEach-Object { $_.Kind }) -join "," | Should Be "next,kept,kept"
        $confirm.Facts[1].Title | Should Be "このフォルダにあるインデックスを、そのまま使います"
    }
}