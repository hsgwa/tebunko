# 段階ごと・種類ごとの時間の集計（tebunko\indexer\ingest_timing.ps1 の IngestTiming）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\indexer\ingest_timing.ps1"
}

Describe "IngestTiming（部品の読み直し）" -Tag Unit {
    It "同じセッションで部品を読み直しても、ログの行を作れる（計測が、読み込み済みのセッションで動かすため）" {
        . "${scriptsDir}\tebunko\indexer\ingest_timing.ps1"
        $t = [IngestTiming]::new()
        $t.AddStage("Ingest", 1000)
        $t.Finish(1500)
        $lines = @($t.FormatLogLines())
        $lines[0] | Should -BeLike "段階ごとの時間（秒）: *全体 1.5*"
    }
}

Describe "IngestTiming（結果と種類の分け方）" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "成功は Ok"; result = @{ Ok = $true }; outcome = "Ok" }
        @{ name = "回し直しは Reroute"; result = @{ Ok = $false; Reroute = $true }; outcome = "Reroute" }
        @{ name = "後回しは Postponed"; result = @{ Ok = $false; Postponed = $true }; outcome = "Postponed" }
        @{ name = "中止で残した分も Postponed"; result = @{ Ok = $false; Stopped = $true }; outcome = "Postponed" }
        @{ name = "制限時間を過ぎた分は Failed"; result = @{ Ok = $false; TimedOut = $true }; outcome = "Failed" }
        @{ name = "どれでもなければ Failed"; result = @{ Ok = $false }; outcome = "Failed" }
    ) {
        param($name, $result, $outcome)
        [IngestTiming]::GetOutcome($result) | Should -Be $outcome
    }

    It "種類は拡張子（小文字）で、無ければ (なし)" -TestCases @(
        @{ path = "a\b.XLSX"; kind = ".xlsx" }
        @{ path = "c.docx"; kind = ".docx" }
        @{ path = "memo"; kind = "(なし)" }
    ) {
        param($path, $kind)
        [IngestTiming]::GetKind($path) | Should -Be $kind
    }
}

Describe "IngestTiming（段階の足し合わせ）" -Tag Unit {
    It "同じ段階は足し、その他は全体から段の合計を引いた残りになる" {
        $t = [IngestTiming]::new()
        $t.AddStage("Startup", 1000)
        $t.AddStage("Crawl", 500)
        $t.AddStage("Ingest", 2000)
        $t.AddStage("Ingest", 1000)
        $t.AddStage("Gram", 500)
        $t.Finish(5000)
        $s = $t.ToHashtable().Stages
        $s.Ingest | Should -Be 3.0
        $s.Other | Should -Be 0.0
        $s.Total | Should -Be 5.0
        $t2 = [IngestTiming]::new()
        $t2.AddStage("Crawl", 500)
        $t2.Finish(2000)
        $s2 = $t2.ToHashtable().Stages
        $s2.Other | Should -Be 1.5
        $s2.Total | Should -Be 2.0
    }

    It "取り込み中・起動中のまとめは内訳で、段の合計に足さない" {
        $t = [IngestTiming]::new()
        $t.AddStage("Ingest", 3000)
        $t.AddStage("IngestFlush", 1000)
        $t.AddStage("StartupFlush", 200)
        $t.Finish(3500)
        $s = $t.ToHashtable().Stages
        $s.IngestFlush | Should -Be 1.0
        $s.StartupFlush | Should -Be 0.2
        $s.Other | Should -Be 0.5
        $s.Total | Should -Be 3.5
    }

    It "画面の確認待ちは全体に含めない" {
        $t = [IngestTiming]::new()
        $t.AddStage("Crawl", 1000)
        $t.AddStage("Wait", 4000)
        $t.Finish(6000)
        $s = $t.ToHashtable().Stages
        $s.Wait | Should -Be 4.0
        $s.Total | Should -Be 2.0
        $s.Other | Should -Be 1.0
    }

    It "Lap は前の印からの時間を足し、Mark は足さない" {
        $t = [IngestTiming]::new()
        $t.Start()
        Start-Sleep -Milliseconds 60
        $t.Mark()
        $t.Lap("Crawl")
        $t.Lap("Gram")
        $t.Finish()
        $s = $t.ToHashtable().Stages
        $s.Crawl | Should -BeLessThan 0.05
        $s.Total | Should -BeGreaterThan 0.05
    }
}

Describe "IngestTiming（種類ごとの足し合わせ）" -Tag Unit {
    It "成功・失敗・回し直し・後回しを数え、1 件の平均は成功の分だけで出す" {
        $t = [IngestTiming]::new()
        $t.AddFile(".xlsx", "Ok", 10, 100, 10)
        $t.AddFile(".xlsx", "Ok", 20, 200, 20)
        $t.AddFile(".xlsx", "Failed", 5, 50, 0)
        $t.AddFile(".xlsx", "Reroute", 1, 1, 0)
        $t.AddPostponed(".xlsx")
        $t.AddFile(".txt", "Failed", 1, 1, 0)
        $k = $t.ToHashtable().Kinds
        $x = $k[".xlsx"]
        $x.Count | Should -Be 3
        $x.Ok | Should -Be 2
        $x.Failed | Should -Be 1
        $x.Reroute | Should -Be 1
        $x.Postponed | Should -Be 1
        $x.ExtractSeconds | Should -Be 0.351
        $x.CopySeconds | Should -Be 0.036
        $x.PublishSeconds | Should -Be 0.03
        $x.AvgMs | Should -Be 180.0
        $k[".txt"].AvgMs | Should -BeNullOrEmpty
    }
}

Describe "IngestTiming（ログの行）" -Tag Unit {
    BeforeAll {
        $script:t = [IngestTiming]::new()
        $script:t.AddStage("Startup", 1400)
        $script:t.AddStage("StartupFlush", 200)
        $script:t.AddStage("Crawl", 500)
        $script:t.AddStage("Ingest", 212000)
        $script:t.AddStage("IngestFlush", 2900)
        $script:t.AddStage("FinishFlush", 300)
        $script:t.AddStage("Gram", 1800)
        $script:t.Finish(219000)
        $script:t.AddFile(".txt", "Ok", 1000, 2000, 500)
        $script:t.AddFile(".xlsx", "Ok", 2000, 150000, 900)
        $script:t.AddPostponed(".xlsx")
    }

    It "段階の 1 行と、抽出の長い順の種類ごとの行を出す" {
        $lines = $script:t.FormatLogLines()
        $lines.Count | Should -Be 3
        $lines[0] | Should -Be "段階ごとの時間（秒）: 起動 1.4（うちまとめ 0.2） / クロール 0.5 / 取り込み 212.0（うちまとめ 2.9） / 仕上げのまとめ 0.3 / gram 索引 1.8 / その他 3.0 / 全体 219.0"
        $lines[1] | Should -Be "  .xlsx 1 件: 抽出 150.0 / コピー 2.0 / 置き場への移し 0.9 / 1 件 152900 ms（失敗 0・回し直し 0・後回し 1）"
        $lines[2] | Should -Be "  .txt 1 件: 抽出 2.0 / コピー 1.0 / 置き場への移し 0.5 / 1 件 3500 ms（失敗 0・回し直し 0・後回し 0）"
    }

    It "成功が 0 件の種類の 1 件は - にし、確認待ちがあれば行に足す" {
        $u = [IngestTiming]::new()
        $u.AddStage("Wait", 3000)
        $u.Finish(4000)
        $u.AddFile(".pdf", "Failed", 0, 10, 0)
        $lines = $u.FormatLogLines()
        $lines[0] | Should -BeLike "*/ 全体 1.0 / 画面の確認待ち 3.0（全体に含めない）"
        $lines[1] | Should -BeLike "*1 件 -（失敗 1・*"
    }
}
