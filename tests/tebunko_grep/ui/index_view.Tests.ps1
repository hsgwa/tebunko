# ［1 インデックス管理］の判断（tebunko_grep\ui\index_view.ps1）のテスト。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\tebunko_grep\ui\index_view.ps1"

function newItem {
    param ([string]$name, [string]$path)
    return [pscustomobject]@{ Name = $name; Path = $path }
}

Describe "getUsedIndexNames" -Tag Unit {
    $items = @((newItem "売上" "C:\data\売上"), (newItem "見積" "C:\data\見積"), (newItem "" "C:\data\新規"))

    It "名前のある行の名前を集める" {
        $used = getUsedIndexNames $items
        @($used).Count | Should Be 2
        $used.Contains("売上") | Should Be $true
    }

    It "大文字と小文字を区別しない" {
        $used = getUsedIndexNames @((newItem "Sales" "C:\data\a"))
        $used.Contains("sales") | Should Be $true
    }

    It "except に渡した行は含めない" {
        $used = getUsedIndexNames $items $items[0]
        $used.Contains("売上") | Should Be $false
        $used.Contains("見積") | Should Be $true
    }
}

Describe "testIndexEditInput" -Tag Unit {
    $items = @((newItem "売上" "C:\data\売上"), (newItem "見積" "C:\data\見積"))

    It "フォルダが空なら、指定するよう伝える" {
        testIndexEditInput "" "新しい名前" $items | Should Be "元のフォルダを指定してください。"
    }

    It "同じフォルダのインデックスがあれば断る" {
        testIndexEditInput "C:\data\売上" "別名" $items | Should Match "インデックス \[売上\] が既にあります"
    }

    It "既にあるインデックスの中のフォルダは断る" {
        testIndexEditInput "C:\data\売上\2024" "別名" $items | Should Match "の中のフォルダです"
    }

    It "既にあるインデックスを含むフォルダは断る" {
        testIndexEditInput "C:\data" "別名" $items | Should Match "があります"
    }

    It "編集中の行は重複の判定から外す" {
        testIndexEditInput "C:\data\売上" "売上" $items $items[0] | Should Be ""
    }

    It "名前が重複していれば断る" {
        testIndexEditInput "C:\data\新規" "見積" $items | Should Not Be ""
    }

    It "問題が無ければ空文字列" {
        testIndexEditInput "C:\data\新規" "新規" $items | Should Be ""
    }
}

Describe "getWorkDirView" -Tag Unit {
    It "既定の場所なら（既定）と付け、［既定に戻す］は出さない" {
        $view = getWorkDirView "C:\tool\work" "C:\Tool\work\"
        $view.Text | Should Be "インデックスの置き場所：C:\tool\work（既定）"
        $view.CanReset | Should Be $false
    }

    It "既定でない場所なら、その場所を出し、［既定に戻す］を出す" {
        $view = getWorkDirView "D:\データ" "C:\tool\work"
        $view.Text | Should Be "インデックスの置き場所：D:\データ"
        $view.CanReset | Should Be $true
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