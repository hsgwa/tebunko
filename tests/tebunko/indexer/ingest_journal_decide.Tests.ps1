# 取り込みのジャーナルの行の組み立てと読み取り（tebunko\indexer\indexer_decide.ps1 の convertToIngestJournalLine・convertFromIngestJournal）のテスト
BeforeDiscovery {
    # テストの表（-TestCases）は探索のときに作るため、ここで行を組み立てる（本番の関数は読み込まれていない）
    function j([string]$kind, [object[]]$fields = @()) {
        return (@($kind) + @($fields | ForEach-Object { [string]$_ })) -join "`t"
    }
    # 版・回（日時・行数）から始まる 1 回分のジャーナル
    function round([int]$rows, [string[]]$records) {
        return @((j "版" @(2)), (j "回" @("2026/10/17 10:00:00", $rows))) + $records
    }
}

BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\indexer\indexer_decide.ps1"
}

Describe "convertToIngestJournalLine" -Tag Unit {
    It "種類と列をタブでつなぐ。列のタブ・改行は 1 行を保つためスペースにする" {
        (convertToIngestJournalLine "始" @("営業\a.xlsx", 2)) | Should -Be "始`t営業\a.xlsx`t2"
        (convertToIngestJournalLine "終") | Should -Be "終"
        (convertToIngestJournalLine "出" @("営業`tx\y`r`nz")) | Should -Be "出`t営業 x\y z"
    }
}

Describe "convertFromIngestJournal" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "終がある（きれいに終わった）なら、復旧するものは無い"
           lines = (round 0 @((j "始" @("営業\a.xlsx", 1)), (j "出" @("営業")), (j "終"))); rows = @()
           folders = @(); resume = @(); finished = $true }
        @{ name = "終があっても、書き出していないフォルダ（出が無い）は書き出す"
           lines = (round 0 @((j "始" @("営業\a.xlsx", 1)), (j "終"))); rows = @("営業\a.xlsx")
           folders = @("営業"); resume = @(); finished = $true }
        @{ name = "始だけで、この回の行が無い仕事は、始の回数のままやり直す（初めて止まった）"
           lines = (round 1 @((j "始" @("営業\a.xlsx", 1)))); rows = @("営業\old.xlsx")
           folders = @("営業"); resume = @("営業\a.xlsx=1") }
        @{ name = "続けて止まった回数（始の回数）は、そのまま返す（上限の判定は呼び出し側）"
           lines = (round 1 @((j "始" @("営業\a.xlsx", 2)))); rows = @("営業\old.xlsx")
           folders = @("営業"); resume = @("営業\a.xlsx=2") }
        @{ name = "この回の行（回の行数より後ろ）がある仕事は済んでいるので、やり直さない"
           lines = (round 1 @((j "始" @("営業\a.xlsx", 1)), (j "始" @("営業\b.xlsx", 1)))); rows = @("営業\old.xlsx", "営業\a.xlsx")
           folders = @("営業"); resume = @("営業\b.xlsx=1") }
        @{ name = "前の回の行（回の行数より前）に同じ相対パスがあっても、この回の行ではないので、やり直す（後回しの行・前回と同じ日時の行）"
           lines = (round 1 @((j "始" @("営業\a.xlsx", 1)))); rows = @("営業\a.xlsx")
           folders = @("営業"); resume = @("営業\a.xlsx=1") }
        @{ name = "出が付いたフォルダは、書き出し直さない。済んでいない仕事はやり直す"
           lines = (round 0 @((j "始" @("営業\a.xlsx", 1)), (j "出" @("営業")))); rows = @()
           folders = @(); resume = @("営業\a.xlsx=1") }
        @{ name = "同じ仕事の始が何度もあれば、最後の行を使う"
           lines = (round 0 @((j "始" @("営業\a.xlsx", 1)), (j "始" @("営業\a.xlsx", 2)))); rows = @()
           folders = @("営業"); resume = @("営業\a.xlsx=2") }
        @{ name = "持だけ（まだ始めていない持ち越し）は、回数のまま返す。始めたものは始の回数を使う"
           lines = (round 0 @((j "持" @("営業\a.xlsx", 1)), (j "持" @("営業\b.xlsx", 1)), (j "始" @("営業\b.xlsx", 2)))); rows = @()
           folders = @("営業"); resume = @("営業\b.xlsx=2", "営業\a.xlsx=1") }
        @{ name = "書の印があるときは、行を見ずに、始の仕事を全部返す"
           lines = (round 1 @((j "始" @("営業\a.xlsx", 1)), (j "書"))); rows = @("営業\old.xlsx", "営業\a.xlsx")
           folders = @("営業"); resume = @("営業\a.xlsx=1") }
        @{ name = "回の行数が取り込み一覧の行数より多い（外から書き換えられた）ときも、始の仕事を全部返す"
           lines = (round 5 @((j "始" @("営業\a.xlsx", 1)))); rows = @("営業\a.xlsx")
           folders = @("営業"); resume = @("営業\a.xlsx=1") }
        @{ name = "回の行数が数でないときも、始の仕事を全部返す"
           lines = @((j "版" @(2)), (j "回" @("2026/10/17 10:00:00", "x")), (j "始" @("営業\a.xlsx", 1))); rows = @("営業\a.xlsx")
           folders = @("営業"); resume = @("営業\a.xlsx=1") }
        @{ name = "回が複数あれば、最後の回より後ろだけを読む"
           lines = ((round 0 @((j "始" @("営業\old.xlsx", 1)))) + @((j "回" @("2026/10/17 11:00:00", 0)), (j "始" @("営業\new.xlsx", 1)))); rows = @()
           folders = @("営業"); resume = @("営業\new.xlsx=1") }
        @{ name = "使えない相対パス（.. で外へ出る・名前だけ）・回数が数でない始と持は読み飛ばす"
           lines = (round 0 @((j "始" @("..\外\a.xlsx", 1)), (j "始" @("営業", 1)), (j "始" @("営業\a.xlsx", "x")), (j "持" @("営業\b.xlsx", 0)), (j "始" @("営業\c.xlsx", 1)))); rows = @()
           folders = @("営業"); resume = @("営業\c.xlsx=1") }
        @{ name = "知らない種類の行と、余分な列は読み飛ばす"
           lines = (round 0 @("未来`tx", ((j "始" @("営業\a.xlsx", 1))) + "`t余分", "")); rows = @()
           folders = @("営業"); resume = @("営業\a.xlsx=1") }
        @{ name = "サブフォルダの仕事は、元のファイルのフォルダ（インデックスのフォルダからの相対）を書き出す"
           lines = (round 0 @((j "始" @("営業\サブ\a.xlsx", 1)), (j "始" @("営業\サブ\b.xlsx", 1)), (j "始" @("営業\c.xlsx", 1)))); rows = @()
           folders = @("営業\サブ", "営業"); resume = @("営業\サブ\a.xlsx=1", "営業\サブ\b.xlsx=1", "営業\c.xlsx=1") }
        @{ name = "始の無い回（取り込みを始める前に止まった）は、何も復旧しない"
           lines = (round 0 @()); rows = @()
           folders = @(); resume = @() }
    ) {
        param ($name, $lines, $rows, $folders, $resume, $finished)
        $result = convertFromIngestJournal ([string[]]$lines) ([string[]]$rows) 2
        $result.Valid | Should -Be $true
        [bool]$result.Finished | Should -Be ([bool]$finished)
        @($result.Folders) | Should -Be @($folders)
        @($result.Resume | ForEach-Object { "$($_.RelPath)=$($_.Count)" }) | Should -Be @($resume)
    }

    It "<name>は、読まない（Valid が偽で、復旧するものは無い）" -TestCases @(
        @{ name = "空"; lines = @() }
        @{ name = "版が違う"; lines = @((j "版" @(1)), (j "回" @("2026/10/17 10:00:00", 0)), (j "始" @("営業\a.xlsx", 1))) }
        @{ name = "版の行が先頭に無い"; lines = @((j "回" @("2026/10/17 10:00:00", 0)), (j "版" @(2)), (j "始" @("営業\a.xlsx", 1))) }
        @{ name = "回が無い"; lines = @((j "版" @(2)), (j "始" @("営業\a.xlsx", 1))) }
    ) {
        param ($name, $lines)
        $result = convertFromIngestJournal ([string[]]$lines) ([string[]]@()) 2
        $result.Valid | Should -Be $false
        @($result.Folders).Count | Should -Be 0
        @($result.Resume).Count | Should -Be 0
    }
}
