# gram 索引で絞った、照合する本文インデックスのファイルの集め方（tebunko\search\gram_search.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "getGramContentIndexFiles" -Tag Io {
    It "gram 索引の置き場が無いときは絞らず、null を返す（呼び出し側が全部を集める）" {
        $root = Join-Path $TestDrive "gram_none_idx"
        [void][System.IO.Directory]::CreateDirectory($root)
        getGramContentIndexFiles "見積書" $true @($root) -indexRoot $root -gramRoot (Join-Path $TestDrive "gram_none_gram") | Should -BeNullOrEmpty
    }
}
