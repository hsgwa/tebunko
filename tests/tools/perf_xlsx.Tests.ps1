# 性能の計測で使うブック（tools\perf\xlsx_data.ps1。Office を使わず ZIP から作る）が、読める形になっていることを確かめる
BeforeAll {
    . "$PSScriptRoot\..\..\tools\perf\xlsx_data.ps1"
    . "$PSScriptRoot\..\..\scripts\tebunko\indexer\indexer_lib.ps1"
}

Describe "newXlsxPerfBook" -Tag Unit {
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

    It "rows と sheets の範囲外は止める" {
        { newXlsxPerfBook (Join-Path $TestDrive "x.xlsx") 0 1 } | Should -Throw "*sheets*"
        { newXlsxPerfBook (Join-Path $TestDrive "x.xlsx") 1 -1 } | Should -Throw "*rows*"
    }
}