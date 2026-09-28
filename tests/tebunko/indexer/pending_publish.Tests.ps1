# 本文インデックスへの書き出し待ち（tebunko\indexer\pending_publish.ps1 の PendingPublish）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\indexer\indexer_plan.ps1"  # getBookDir
    . "${scriptsDir}\tebunko\indexer\pending_publish.ps1"
}

Describe "PendingPublish" -Tag Unit {
    It "取り込んだファイルのフォルダは、無くなったファイル名を記録せずに書き出し待ちにする" {
        $pending = [PendingPublish]::new()
        $pending.Add("営業\見積.xlsx", $false)
        $flush = $pending.TakeFlushable(@())
        $flush.Count | Should -Be 1
        $folder = [System.IO.Path]::GetDirectoryName((getBookDir "営業\見積.xlsx"))
        $flush[$folder].Count | Should -Be 0
    }

    It "無くなったファイルは、そのファイル名も書き出し待ちに記録する" {
        $pending = [PendingPublish]::new()
        $pending.Add("営業\無い.xlsx", $true)
        $folder = [System.IO.Path]::GetDirectoryName((getBookDir "営業\無い.xlsx"))
        $flush = $pending.TakeFlushable(@())
        $flush[$folder].Contains("無い.xlsx") | Should -Be $true
    }

    It "まだ取り込みが続くフォルダ（keepFolders）は取り出さず、書き出し待ちに残す" {
        $pending = [PendingPublish]::new()
        $pending.Add("営業\a.xlsx", $false)
        $folder = [System.IO.Path]::GetDirectoryName((getBookDir "営業\a.xlsx"))

        $flush = $pending.TakeFlushable(@($folder))
        $flush.Count | Should -Be 0

        # keepFolders から外れれば、次に取り出せる
        $flush = $pending.TakeFlushable(@())
        $flush.ContainsKey($folder) | Should -Be $true
    }

    It "一度取り出したフォルダは、書き出し待ちから消える（二重に書き出さない）" {
        $pending = [PendingPublish]::new()
        $pending.Add("営業\b.xlsx", $false)
        [void]$pending.TakeFlushable(@())
        $pending.TakeFlushable(@()).Count | Should -Be 0
    }

    It "前回のインデックス作成で本文インデックスに入れていないフォルダ（MarkFolder）も書き出し待ちにする" {
        $pending = [PendingPublish]::new()
        $pending.MarkFolder("C:\work\content_index\営業")
        $flush = $pending.TakeFlushable(@())
        $flush.ContainsKey("C:\work\content_index\営業") | Should -Be $true
        $flush["C:\work\content_index\営業"].Count | Should -Be 0
    }
}
