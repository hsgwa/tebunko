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
        testIndexEditInput "C:\data\新規" "見積" $items | Should Be "「見積」は、ほかのインデックスが使っています。別の名前を付けてください。"
    }

    It "問題が無ければ空文字列" {
        testIndexEditInput "C:\data\新規" "新規" $items | Should Be ""
    }
}
