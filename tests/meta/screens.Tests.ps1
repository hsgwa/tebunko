# 画面の写真（docs\images\screens\・tools\capture_screens.ps1 で撮る）のテスト。
# 写真そのものを撮る動きには触らない（実機の確かめは PR 本文に書く）。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$here\..\tools\capture\capture_common.ps1"
    $script:screensDir = "$here\..\docs\images\screens"
    $script:docsDir = "$here\..\docs\design\gui\screens"
}

Describe "画面の写真" -Tag Meta {
    It "状態の ID ごとに写真がある" {
        $missing = @(${captureIds} | Where-Object { !(Test-Path -LiteralPath (getCaptureImagePath -Id $_ -OutDir $script:screensDir)) })
        ($missing -join "、") | Should -Be ""
    }

    It "写真は 1 枚 200 KB 以下" {
        $files = @(Get-ChildItem -LiteralPath $script:screensDir -Recurse -Filter "*.png" -ErrorAction SilentlyContinue)
        $files.Count | Should -BeGreaterThan 0
        $over = @($files | Where-Object { !(testCaptureImageSize -Bytes $_.Length) } | ForEach-Object { "$($_.Name)：$([Math]::Round($_.Length / 1KB)) KB" })
        ($over -join "、") | Should -Be ""
    }

    It "写真は全部で 5 MB 以下" {
        $files = @(Get-ChildItem -LiteralPath $script:screensDir -Recurse -Filter "*.png" -ErrorAction SilentlyContinue)
        $total = ($files | Measure-Object -Property Length -Sum).Sum
        $total | Should -BeLessOrEqual 5MB
    }

    It "写真は画面ごとのページ（docs\design\gui\screens）から参照されている" {
        $files = @(Get-ChildItem -LiteralPath $script:screensDir -Recurse -Filter "*.png" -ErrorAction SilentlyContinue)
        $pages = @(Get-ChildItem -LiteralPath $script:docsDir -Filter "*.md" -ErrorAction SilentlyContinue)
        $text = ($pages | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8 }) -join "`n"
        $unused = @($files | Where-Object { $text -notlike "*$($_.Name)*" } | ForEach-Object { $_.FullName.Substring($script:screensDir.Length + 1) })
        ($unused -join "、") | Should -Be ""
    }

    It "ID の形（<画面>/<状態>）は、画面のページのファイル名（<画面>.md）と対応する" {
        $screens = @(${captureIds} | ForEach-Object { ($_ -split "/", 2)[0] } | Select-Object -Unique)
        $pages = @(Get-ChildItem -LiteralPath $script:docsDir -Filter "*.md" -ErrorAction SilentlyContinue | ForEach-Object { $_.BaseName })
        $missing = @($screens | Where-Object { $pages -notcontains $_ })
        ($missing -join "、") | Should -Be ""
    }
}
