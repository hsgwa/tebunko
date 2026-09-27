# 取り込み（.docx・.pptx。Office を使わずに読むファイル）の速さの回帰テスト（tools\measure_perf.ps1 で測り、上限と比べる）
#
# Slow はテストデータ（tests\testdata\office）の複製を作って測る。tebunko-perfdata は要らない。
# Excel・Word・PowerPoint を使う形式（.xlsx・.doc・.ppt）は、機械・Defender・Office の版で大きく揺れるので固定の上限を置かない。
# 手元で main と続けて測って比べる（docs\design\testing\ci.md の「Office を使う形式の比べ方」）。
#
#   .\tests\run.ps1 -Tag Slow -ExcludeTag Manual -Path tests\tools\perf_ingest.Tests.ps1
#
# 上限は、ランナー（GitHub の windows-latest。4 コア）で perf-check.yml と同じ構成で 5 回測った比べる値の最大に 1.5 倍の余裕を掛けて決めた。
# 1 ファイルあたりは 50 ms 単位、全体は 10 秒単位で切り上げる。決め方と変え方は docs\design\testing\ci.md。
#   取り込むファイルの数（.docx・.pptx を 50 個ずつ）
#   1 ファイルあたりの中央値の上限（ms）・全体の中央値の上限（秒）
BeforeAll {
    . "$PSScriptRoot\..\..\tools\perf\perf_common.ps1"

    $script:ingestFiles = 100
    $script:perFileLimitMs = 1200
    $script:secondsLimit = 130
}

Describe "取り込み（.docx・.pptx）の速さ" -Tag Slow {
    BeforeAll {
        $repo = (Resolve-Path "$PSScriptRoot\..\..").ProviderPath
        # スクリプトは同じプロセスで動かす（Windows PowerShell 5.1 は -File で起動すると、param の既定値の $PSScriptRoot が空になるため）
        $global:LASTEXITCODE = 0
        $office = Join-Path $TestDrive "office"
        $work = Join-Path $TestDrive "work"
        $out = Join-Path $repo "work\test\perf-ingest"
        if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Recurse -Force }

        & (Join-Path $repo "tools\perf\new_ingest_data.ps1") -Dest $office -Docx 50 -Pptx 50 | Out-Host
        if ($LASTEXITCODE) { throw "取り込みのデータを作れませんでした（終了コード $LASTEXITCODE）。" }
        # -Index は渡さない（取り込みだけを測る）。1 回ごとに新しいプロセス・空のワークスペースで取り込む
        & (Join-Path $repo "tools\measure_perf.ps1") -Office $office -Work $work -Threads 2 -Repeat 3 -Out $out | Out-Host
        if ($LASTEXITCODE) { throw "measure_perf.ps1 が失敗しました（終了コード $LASTEXITCODE）。" }

        $result = [System.IO.File]::ReadAllText((Join-Path $out "result.json"), [System.Text.Encoding]::UTF8) | ConvertFrom-Json
        # 最後の回のワークスペース（measure_ingest.ps1 は次の回の始めまで消さない）にできた集約ファイル。サブフォルダの下にもできる
        $files = @(Get-ChildItem -LiteralPath (Join-Path $work "ingest\ws") -Recurse -File -Filter "content.*.tsv" -ErrorAction SilentlyContinue)
        $aggregate = @{ Count = $files.Count; Bytes = [long](($files | Measure-Object Length -Sum).Sum) }
    }

    It "1 ファイルあたりと全体の中央値が上限以内で、全部取り込まれ、集約ファイルができている" {
        $problems = getIngestPerfProblems $result.Ingest $aggregate $perFileLimitMs $secondsLimit $ingestFiles
        ($problems -join "`n") | Should -BeNullOrEmpty
    }
}

Describe "取り込みの判定（getIngestPerfProblems）" -Tag Unit {
    BeforeAll {
        function newIngest {
            [pscustomobject]@{
                Total = 100; Done = 100; Failed = 0
                Seconds = [pscustomobject]@{ Median = 80.0 }
                PerFileMs = [pscustomobject]@{ Median = 800.0 }
            }
        }
        $script:goodAggregate = @{ Count = 4; Bytes = 12345 }
    }

    It "<name>" -TestCases @(
        @{ name = "上限より小さければ合格"; perFile = 800; seconds = 80 }
        @{ name = "ちょうど上限なら合格"; perFile = 1200; seconds = 120 }
    ) {
        param ($name, $perFile, $seconds)
        $ingest = newIngest
        $ingest.PerFileMs.Median = $perFile
        $ingest.Seconds.Median = $seconds
        getIngestPerfProblems $ingest $goodAggregate 1200 120 100 | Should -BeNullOrEmpty
    }

    It "<name>" -TestCases @(
        @{ name = "1 ファイルあたりの中央値が上限を超えたら返す"; change = { param($i) $i.PerFileMs.Median = 1201 }; aggregate = @{ Count = 4; Bytes = 1 }; expect = "取り込み 1 ファイルあたり: 中央値 1,201 ms（上限 1,200 ms）" }
        @{ name = "全体の中央値が上限を超えたら返す"; change = { param($i) $i.Seconds.Median = 121.5 }; aggregate = @{ Count = 4; Bytes = 1 }; expect = "取り込み全体: 中央値 121.5 秒（上限 120 秒）" }
        @{ name = "Done が少なければ返す"; change = { param($i) $i.Done = 99 }; aggregate = @{ Count = 4; Bytes = 1 }; expect = "取り込みが済んだのは 99 件（期待 100 件）" }
        @{ name = "Total がファイルの数と違えば返す"; change = { param($i) $i.Total = 50 }; aggregate = @{ Count = 4; Bytes = 1 }; expect = "取り込んだファイルが 50 件（期待 100 件）" }
        @{ name = "Failed が 1 以上なら返す"; change = { param($i) $i.Failed = 1 }; aggregate = @{ Count = 4; Bytes = 1 }; expect = "取り込みに失敗したのは 1 件" }
        @{ name = "1 ファイルあたりの時間が無ければ返す"; change = { param($i) $i.PerFileMs = $null }; aggregate = @{ Count = 4; Bytes = 1 }; expect = "1 ファイルあたりの時間がありません" }
        @{ name = "全体の時間が無ければ返す"; change = { param($i) $i.Seconds = $null }; aggregate = @{ Count = 4; Bytes = 1 }; expect = "全体の時間がありません" }
        @{ name = "集約ファイルが無ければ返す"; change = { param($i) }; aggregate = @{ Count = 0; Bytes = 0 }; expect = "集約ファイルがありません" }
        @{ name = "集約ファイルを調べていなければ返す"; change = { param($i) }; aggregate = $null; expect = "集約ファイルがありません" }
        @{ name = "集約ファイルの合計の大きさが 0 なら返す"; change = { param($i) }; aggregate = @{ Count = 4; Bytes = 0 }; expect = "集約ファイルの合計の大きさが 0 です" }
    ) {
        param ($name, $change, $aggregate, $expect)
        $ingest = newIngest
        & $change $ingest
        $problems = getIngestPerfProblems $ingest $aggregate 1200 120 100
        $problems | Should -Contain $expect -Because ($problems -join "; ")
    }

    It "数の書き方は、今のカルチャに左右されない（de-DE でも 1,201・121.5 のまま）" {
        $old = [System.Threading.Thread]::CurrentThread.CurrentCulture
        try {
            [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo("de-DE")
            $ingest = newIngest
            $ingest.PerFileMs.Median = 1201
            $ingest.Seconds.Median = 121.5
            $problems = getIngestPerfProblems $ingest $goodAggregate 1200 120 100
            $problems | Should -Contain "取り込み 1 ファイルあたり: 中央値 1,201 ms（上限 1,200 ms）"
            $problems | Should -Contain "取り込み全体: 中央値 121.5 秒（上限 120 秒）"
        } finally {
            [System.Threading.Thread]::CurrentThread.CurrentCulture = $old
        }
    }

    It "Ingest が無ければ返す" {
        $problems = getIngestPerfProblems $null $goodAggregate 1200 120 100
        $problems | Should -Be @("取り込みの結果がありません")
    }
}
