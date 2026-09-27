# 検索と pack の作成の速さの回帰テスト（tools\measure_perf.ps1 で測り、上限と比べる）
#
# Slow は tebunko-perfdata（データを作るスクリプトのリポジトリ）が要る。場所は環境変数 TEBUNKO_PERFDATA、
# 無ければリポジトリと並んだ tebunko-perfdata（git worktree のときは、本体のチェックアウトと並んだもの）。
# CI では perf-check.yml が PERFDATA_SHA のコミットを取り出す。
#
#   .\tests\run.ps1 -Tag Slow -Path tests\tools\perf_search.Tests.ps1
#
# 上限は、ランナー（GitHub の windows-latest。4 コア）で perf-check.yml と同じ構成で 5 回測った比べる値の最大に余裕を掛けて決めた。
# 検索は 1.5 倍を 50 ms 単位、pack の作成は 2 倍を 5 秒単位で切り上げる。決め方と変え方は docs\design\testing\ci.md。
#   語ごとの上限（検索の中央値。ms）と、期待する件数は scale 0.1・種 1 のときの値（words.tsv の「件数」）
#   pack の作成の上限（1 回の秒）
BeforeAll {
    . "$PSScriptRoot\..\..\tools\perf\perf_common.ps1"

    $script:searchLimitsMs = @{ "0 件" = 10000; "まれ" = 10000; "大量" = 10000; "正規表現" = 10000 }
    $script:packLimitSeconds = 600
}

Describe "検索と pack の作成の速さ" -Tag Slow {
    BeforeAll {
        $repo = (Resolve-Path "$PSScriptRoot\..\..").ProviderPath

        # tebunko-perfdata の場所を探す
        $candidates = New-Object System.Collections.Generic.List[string]
        if ($env:TEBUNKO_PERFDATA) { $candidates.Add($env:TEBUNKO_PERFDATA) }
        $candidates.Add((Join-Path (Split-Path $repo -Parent) "tebunko-perfdata"))
        $common = & git -C $repo rev-parse --path-format=absolute --git-common-dir 2>$null
        if ($common) { $candidates.Add((Join-Path (Split-Path (Split-Path ([string]$common) -Parent) -Parent) "tebunko-perfdata")) }
        $perfdata = $candidates | Where-Object { Test-Path -LiteralPath (Join-Path $_ "tools\new_index.ps1") } | Select-Object -First 1
        if (!$perfdata) {
            throw ("tebunko-perfdata が見つかりません。次のように取ってくるか、環境変数 TEBUNKO_PERFDATA に置き場所を指定してください。`n" +
                "  git clone https://github.com/hsgwa/tebunko-perfdata `"$(Join-Path (Split-Path $repo -Parent) 'tebunko-perfdata')`"")
        }
        $perfdata = (Resolve-Path -LiteralPath $perfdata).ProviderPath
        $commit = & git -C $perfdata rev-parse HEAD 2>$null
        if ($LASTEXITCODE -or !$commit) { $commit = "不明" }

        # スクリプトは同じプロセスで動かす（Windows PowerShell 5.1 は -File で起動すると、param の既定値の $PSScriptRoot が空になるため）
        $global:LASTEXITCODE = 0
        $index = Join-Path $TestDrive "index"
        $work = Join-Path $TestDrive "work"
        $out = Join-Path $repo "work\test\perf-search"
        if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Recurse -Force }
        $words = Join-Path $perfdata "words.tsv"

        & (Join-Path $perfdata "tools\new_index.ps1") -Dest $index -Scale 0.1 | Out-Host
        if ($LASTEXITCODE) { throw "テスト用のデータを作れませんでした（終了コード $LASTEXITCODE）。" }
        & (Join-Path $repo "tools\measure_perf.ps1") -Index $index -Work $work -Words $words -Count 20 -Label "tebunko-perfdata $commit" -Out $out | Out-Host
        if ($LASTEXITCODE) { throw "measure_perf.ps1 が失敗しました（終了コード $LASTEXITCODE）。" }

        $result = [System.IO.File]::ReadAllText((Join-Path $out "result.json"), [System.Text.Encoding]::UTF8) | ConvertFrom-Json
        $rows = @(Import-Csv -LiteralPath (Join-Path $out "searches.csv") -Encoding UTF8)
        # 期待する件数は words.tsv の 4 列目（1 行目は見出し）
        $expected = @{}
        foreach ($line in @([System.IO.File]::ReadAllLines($words, [System.Text.Encoding]::UTF8) | Select-Object -Skip 1)) {
            $f = $line.Split("`t")
            if ($f.Count -ge 4) { $expected[$f[0]] = $f[3] }
        }
    }

    It "検索の中央値と pack の作成が上限以内で、件数と pack の数が合う" {
        $problems = getSearchPerfProblems $result $rows $searchLimitsMs $expected $packLimitSeconds
        ($problems -join "`n") | Should -BeNullOrEmpty
    }
}

Describe "検索と pack の作成の判定（getSearchPerfProblems）" -Tag Unit {
    BeforeAll {
        $script:limits = @{ "0 件" = 450; "まれ" = 550; "大量" = 2100; "正規表現" = 2200 }
        $script:expected = @{ "0 件" = "0"; "まれ" = "3"; "大量" = "10000+"; "正規表現" = "10000+" }

        # 合格になる result.json（の一部）と searches.csv の行を作る
        function newFixture {
            $hits = @{ "0 件" = 0; "まれ" = 3; "大量" = 10000; "正規表現" = 10000 }
            $search = foreach ($name in $limits.Keys) {
                [pscustomobject]@{ Name = $name; TotalMs = [pscustomobject]@{ Median = 300 } }
            }
            $rows = foreach ($name in $limits.Keys) {
                foreach ($n in 1..3) {
                    [pscustomobject]@{ Word = $name; N = $n; Hits = $hits[$name]; Truncated = $(if ($hits[$name] -eq 10000) { "True" } else { "False" }); Packs = 521 }
                }
            }
            $result = [pscustomobject]@{
                Run = [pscustomobject]@{ SearchMode = "service" }
                Index = [pscustomobject]@{ Pack = [pscustomobject]@{ Seconds = 40; Packs = 521 } }
                Search = @($search)
            }
            return @{ Result = $result; Rows = @($rows) }
        }
        function getEntry($fixture, [string]$name) { @($fixture.Result.Search | Where-Object { $_.Name -eq $name })[0] }
        function getRows($fixture, [string]$name) { @($fixture.Rows | Where-Object { $_.Word -eq $name }) }
    }

    It "上限より小さい・ちょうど上限なら合格（一覧が空）" {
        $f = newFixture
        (getEntry $f "0 件").TotalMs.Median = 449
        (getEntry $f "まれ").TotalMs.Median = 550
        $f.Result.Index.Pack.Seconds = 90
        getSearchPerfProblems $f.Result $f.Rows $limits $expected 90 | Should -BeNullOrEmpty
    }

    It "<name>" -TestCases @(
        @{ name = "中央値が上限を超えたら、語・中央値・上限を返す"; change = { param($f) (getEntry $f "大量").TotalMs.Median = 2101 }; expect = "検索 大量: 中央値 2,101 ms（上限 2,100 ms）" }
        @{ name = "語が結果に無ければ返す"; change = { param($f) $f.Result.Search = @($f.Result.Search | Where-Object { $_.Name -ne "まれ" }) }; expect = "検索 まれ: 結果がありません" }
        @{ name = "語の 1 回ごとの記録が無ければ返す"; change = { param($f) $f.Rows = @($f.Rows | Where-Object { $_.Word -ne "まれ" }) }; expect = "検索 まれ: 1 回ごとの記録がありません" }
        @{ name = "件数が違えば返す"; change = { param($f) (getRows $f "まれ")[0].Hits = 2 }; expect = "検索 まれ: 件数が期待の 3 と違う回が 1 回（最初は 2 件）" }
        @{ name = "2 回目以降だけ件数が違っても返す"; change = { param($f) (getRows $f "0 件")[2].Hits = 1 }; expect = "検索 0 件: 件数が期待の 0 と違う回が 1 回（最初は 1 件）" }
        @{ name = "10000+ なのに打ち切りでなければ返す"; change = { param($f) (getRows $f "大量")[1].Truncated = "False" }; expect = "検索 大量: 件数が期待の 10000+ と違う回が 1 回（最初は 10000 件）" }
        @{ name = "10000+ なのに 1 万件でなければ返す"; change = { param($f) (getRows $f "正規表現")[0].Hits = 9000 }; expect = "検索 正規表現: 件数が期待の 10000+ と違う回が 1 回（最初は 9000 件）" }
        @{ name = "pack の数が少ない回があれば返す"; change = { param($f) (getRows $f "0 件")[1].Packs = 300 }; expect = "検索 0 件: 照合した pack の数が 521 と違う回が 1 回（最初は 300）" }
        @{ name = "照合した pack が 0 の回があれば返す"; change = { param($f) (getRows $f "まれ")[0].Packs = 0 }; expect = "検索 まれ: 照合した pack の数が 521 と違う回が 1 回（最初は 0）" }
        @{ name = "SearchMode が runspace なら返す"; change = { param($f) $f.Result.Run.SearchMode = "runspace" }; expect = "検索の流れが service ではありません（runspace）" }
        @{ name = "pack の作成が上限を超えたら返す"; change = { param($f) $f.Result.Index.Pack.Seconds = 90.5 }; expect = "pack の作成: 90.5 秒（上限 90 秒）" }
        @{ name = "pack が無ければ返す"; change = { param($f) $f.Result.Index.Pack.Packs = 0 }; expect = "pack がありません" }
        @{ name = "pack の作成を測っていなければ返す"; change = { param($f) $f.Result.Index.Pack = $null }; expect = "pack がありません" }
    ) {
        param ($name, $change, $expect)
        $f = newFixture
        & $change $f
        $problems = getSearchPerfProblems $f.Result $f.Rows $limits $expected 90
        $problems | Should -Contain $expect -Because ($problems -join "; ")
    }
}
