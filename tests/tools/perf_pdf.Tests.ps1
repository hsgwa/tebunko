# PDF の読み取りの速さを測る道具（tools\perf\new_pdf_data.ps1・measure_pdf_version.ps1）が動くことの確かめ。速さは測らない
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    $script:repo = (Resolve-Path "$PSScriptRoot\..\..").ProviderPath
}

Describe "PDF の計測の道具" -Tag Io {
    BeforeAll {
        $script:pdfPath = Join-Path $TestDrive "perf.pdf"
        & "$script:repo\tools\perf\new_pdf_data.ps1" -Dest $script:pdfPath -Scale 0.02
    }

    It "作った PDF を読めて、文字・図形・ToUnicode のページの文字が出る" {
        $pages = readPdfFile $script:pdfPath
        $pages.Count | Should -Be 3
        $pages[0].Length | Should -BeGreaterThan 100
        $pages[1].Length | Should -BeGreaterThan 0
        $pages[2].Length | Should -BeGreaterThan 100
    }

    It "同じ引数からは同じ PDF ができる" {
        $again = Join-Path $TestDrive "again.pdf"
        & "$script:repo\tools\perf\new_pdf_data.ps1" -Dest $again -Scale 0.02
        [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($again)) | Should -BeExactly ([Convert]::ToBase64String([System.IO.File]::ReadAllBytes($script:pdfPath)))
    }

    It "版ごとの計測が、時間・文字の検査値・関数ごとの時間を出す" {
        $json = Join-Path $TestDrive "version.json"
        & "$script:repo\tools\perf\measure_pdf_version.ps1" -Scripts "$script:repo\scripts" -Pdf $script:pdfPath -Repeat 1 -Profile -Out $json
        $result = [System.IO.File]::ReadAllText($json, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
        @($result.Seconds).Count | Should -Be 1
        $result.Pages | Should -Be 3
        $result.Hash.Length | Should -Be 64
        (@($result.Profile | Where-Object { $_.Name -eq "readPdfPageText" })).Calls | Should -Be 3
    }
}
