# 性能の計測で使うブック（tools\perf\xlsx_data.ps1。Office を使わず ZIP から作る）が読める形になっていることと、それを測る tools\perf\measure_xlsx.ps1 が動くことを確かめる
BeforeAll {
    . "$PSScriptRoot\..\..\tools\perf\xlsx_data.ps1"
    . "$PSScriptRoot\..\..\scripts\tebunko\indexer\indexer_lib.ps1"
}

Describe "newXlsxPerfBook" -Tag Io {
    It "<sheets> 個の <sheet> が同じシートの部品を指し、ヘッダーとフッターを読める" -TestCases @(
        @{ sheets = 1 }
        @{ sheets = 5 }
    ) {
        $path = Join-Path $TestDrive "book$sheets.xlsx"
        newXlsxPerfBook $path $sheets 20
        $failures = New-Object System.Collections.Generic.List[string]
        $units = readXlsxObjectUnits $path $failures
        $failures.Count | Should -Be 0
        $units.Count | Should -Be $sheets
        @($units["Sheet1[ヘッダー・フッター]"]) | Should -Be @("Perf Header", "Perf Footer")
    }

    It "<name> は止める" -TestCases @(
        @{ name = "sheets が 0"; sheets = 0; rows = 1; message = "*sheets*" }
        @{ name = "rows が負"; sheets = 1; rows = -1; message = "*rows*" }
    ) {
        { newXlsxPerfBook (Join-Path $TestDrive "x.xlsx") $sheets $rows } | Should -Throw $message
    }
}

Describe "measure_xlsx.ps1" -Tag Io {
    It "小さなブックで最後まで動き、読んだ場所の数と失敗の数と段階ごとの記録を JSON に書く" {
        $script = "$PSScriptRoot\..\..\tools\perf\measure_xlsx.ps1"
        $tool = (Resolve-Path "$PSScriptRoot\..\..").ProviderPath
        $out = Join-Path $TestDrive "xlsx.json"
        & $script -Tool $tool -Work (Join-Path $TestDrive "work") -Out $out -Rows 50 -ManySheets 3 -ManyRows 20 -Repeat 2 -SampleMs 20 6>$null | Out-Null
        $r = [System.IO.File]::ReadAllText($out) | ConvertFrom-Json
        @($r.Normal).Count | Should -Be 2
        $r.Normal[0].Units | Should -Be 1
        $r.Normal[0].Lines | Should -Be 2
        $r.Many.Lines | Should -Be 6
        $r.Many.Units | Should -Be 3
        $r.Normal[0].Failures | Should -Be 0
        $r.Normal[1].Failures | Should -Be 0
        $r.Many.Failures | Should -Be 0
        $phases = @($r.Samples | ForEach-Object { $_.Phase })
        $phases | Should -Contain "普通のブック"
        $phases | Should -Contain "同じ部品を指すブック"
    }
}
