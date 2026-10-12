# 選択行のプレビューの速さとメモリの回帰テスト（tools\measure_perf.ps1 -Preview で測り、上限と比べる）
#
# Slow は tebunko-perfdata（データを作るスクリプトのリポジトリ）が要る。場所は環境変数 TEBUNKO_PERFDATA、
# 無ければリポジトリと並んだ tebunko-perfdata（git worktree のときは、本体のチェックアウトと並んだもの）。
# CI では perf-check.yml が PERFDATA_SHA のコミットを取り出す。
#
#   .\tests\run.ps1 -Tag Slow -ExcludeTag Manual -Path tests\tools\perf_preview.Tests.ps1
#
# 上限は、ランナー（GitHub の windows-latest。4 コア）で perf-check.yml と同じ構成で 5 回測った値の最大に余裕を掛けて決める。
# 読み込みは 1.5 倍を 10 ms 単位、1 回あたりに残るメモリは 2 倍を 0.1 MB 単位で切り上げる。決め方と変え方は docs\design\testing\ci.md。
#   語ごとの上限（キャッシュ無しの読み込みの中央値。ms）は scale 0.1・種 1 のときの値
#   1 回あたりに残るメモリの上限（GC のあとにマネージドヒープに残る増え。MB）
BeforeAll {
    . "$PSScriptRoot\..\..\tools\perf\perf_common.ps1"

    $script:previewCount = 20
    $script:readLimitsMs = @{ "まれ" = 500; "大量" = 500; "正規表現" = 500 }
    $script:retainedPerPreviewLimitMB = 5.0
}

Describe "選択行のプレビューの速さとメモリ" -Tag Slow {
    BeforeAll {
        $repo = (Resolve-Path "$PSScriptRoot\..\..").ProviderPath

        # git が使えない・リポジトリでないときも失敗にしない（標準エラーが例外になるのを避ける）
        function tryGit {
            $saved = $ErrorActionPreference
            $ErrorActionPreference = "Continue"
            try { $gitOutput = & git @args 2>$null; if ($LASTEXITCODE) { return $null }; return $gitOutput } catch { return $null } finally { $ErrorActionPreference = $saved }
        }

        # tebunko-perfdata の場所を探す
        $candidates = New-Object System.Collections.Generic.List[string]
        if ($env:TEBUNKO_PERFDATA) { $candidates.Add($env:TEBUNKO_PERFDATA) }
        $candidates.Add((Join-Path (Split-Path $repo -Parent) "tebunko-perfdata"))
        $common = tryGit -C $repo rev-parse --path-format=absolute --git-common-dir
        if ($common) { $candidates.Add((Join-Path (Split-Path (Split-Path ([string]$common) -Parent) -Parent) "tebunko-perfdata")) }
        $perfdata = $candidates | Where-Object { Test-Path -LiteralPath (Join-Path $_ "tools\new_index.ps1") } | Select-Object -First 1
        if (!$perfdata) {
            throw ("tebunko-perfdata が見つかりません。次のように取ってくるか、環境変数 TEBUNKO_PERFDATA に置き場所を指定してください。`n" +
                "  git clone https://github.com/hsgwa/tebunko-perfdata `"$(Join-Path (Split-Path $repo -Parent) 'tebunko-perfdata')`"")
        }
        $perfdata = (Resolve-Path -LiteralPath $perfdata).ProviderPath
        $commit = tryGit -C $perfdata rev-parse HEAD
        if (!$commit) { $commit = "不明" }

        # スクリプトは同じプロセスで動かす（Windows PowerShell 5.1 は -File で起動すると、param の既定値の $PSScriptRoot が空になるため）
        $global:LASTEXITCODE = 0
        $index = Join-Path $TestDrive "index"
        $work = Join-Path $TestDrive "work"
        # 結果は毎回、同じ名前のファイルを上書きして書く
        $out = Join-Path $repo "work\test\perf-preview"
        $words = Join-Path $perfdata "words.tsv"

        & (Join-Path $perfdata "tools\new_index.ps1") -Dest $index -Scale 0.1 | Out-Host
        if ($LASTEXITCODE) { throw "テスト用のデータを作れませんでした（終了コード $LASTEXITCODE）。" }
        & (Join-Path $repo "tools\measure_perf.ps1") -Index $index -Work $work -Words $words -Count 1 -Preview $previewCount -Label "tebunko-perfdata $commit" -Out $out | Out-Host
        if ($LASTEXITCODE) { throw "measure_perf.ps1 が失敗しました（終了コード $LASTEXITCODE）。" }

        $result = [System.IO.File]::ReadAllText((Join-Path $out "result.json"), [System.Text.Encoding]::UTF8) | ConvertFrom-Json
    }

    It "ヒットのある語のプレビューが、キャッシュ無しの読み込みの上限以内で、残るメモリが上限以内" {
        $problems = New-Object System.Collections.Generic.List[string]
        foreach ($name in @($readLimitsMs.Keys | Sort-Object)) {
            $entry = @($result.Preview.Words | Where-Object { $_.Name -eq $name }) | Select-Object -First 1
            $searched = @($result.Search | Where-Object { $_.Name -eq $name }) | Select-Object -First 1
            if ($null -eq $entry -or $null -eq $searched) { $problems.Add("プレビュー ${name}: 結果がありません"); continue }
            $expectedCount = [Math]::Min($previewCount, [int]$searched.Hits)
            foreach ($problem in @(getPreviewPerfProblems $entry $readLimitsMs[$name] $retainedPerPreviewLimitMB $expectedCount)) { $problems.Add("プレビュー ${name}: $problem") }
        }
        ($problems -join "`n") | Should -BeNullOrEmpty
    }

    It "ヒットの無い語は、プレビューを測らずに飛ばす" {
        @($result.Preview.Skipped) | Should -Contain "0 件"
    }
}

Describe "プレビューの描画の計測（-Render）" -Tag Gui {
    # 画面は出さず、プレビューの表を画面の部品に入れて Measure・Arrange する時間を測る
    BeforeAll {
        function writeTsv([string]$path, [string[]]$lines) {
            [void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path))
            [System.IO.File]::WriteAllText($path, (($lines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
        }
        $index = Join-Path $TestDrive "work\index"
        writeTsv "$index\営業\見積.xlsx\シート1.tsv" @("1`t見積書`t山田", "2`tりんご`t佐藤")
        $words = Join-Path $TestDrive "words.tsv"
        writeTsv $words @("名前`t語`t正規表現`t件数", "まれ`t見積書`tfalse`t1")
        $out = Join-Path $TestDrive "work\result"
        & "$PSScriptRoot\..\..\tools\measure_perf.ps1" -Index $index -Work (Join-Path $TestDrive "work") -Words $words -Count 1 -Preview 2 -Render -SampleMs 50 6>$null | Out-Null
        $result = [System.IO.File]::ReadAllText("$out\result.json") | ConvertFrom-Json
    }

    It "描画の時間を 3 回分と、セルの数・行・列を書く" {
        $result.Preview.Render | Should -Be $true
        $entry = @($result.Preview.Words)[0]
        $entry.Render.Cells | Should -BeGreaterThan 0
        $entry.Render.Rows | Should -BeGreaterThan 0
        $entry.Render.Columns | Should -BeGreaterThan 0
        @($entry.Render.RenderMs).Count | Should -Be 3
        $md = [System.IO.File]::ReadAllText("$out\summary.md")
        $md | Should -Match "### プレビューの描画"
        @(Import-Csv "$out\metrics.csv" | Where-Object { $_.metric -eq "preview_render_ms" }).Count | Should -Be 1
    }
}
