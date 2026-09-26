# Markdown のリンク切れを確かめるスクリプト（tools\check_markdown_links.ps1）のテスト
BeforeAll {
    $check = "$PSScriptRoot\..\..\tools\check_markdown_links.ps1"

    function writeText([string]$path, [string]$text) {
        [void](New-Item -ItemType Directory -Force (Split-Path $path))
        [System.IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding($false)))
    }

    # docs\page.md に本文を書いて調べ、切れたリンクの行（「  docs/page.md:行: 理由」）を返す
    function checkPage([string]$markdown) {
        writeText "$TestDrive\repo\docs\page.md" $markdown
        $output = @(& $check -Root "$TestDrive\repo" -Path "docs/page.md" 6>&1 | ForEach-Object { "$_" })
        $problems = @($output | Where-Object { $_.StartsWith("  ") } | ForEach-Object { $_.Trim() })
        # 終了コードと出力が食い違わないこと
        $null = ($LASTEXITCODE -eq 1) | Should -Be ($problems.Count -gt 0)
        return $problems
    }
}

Describe "check_markdown_links.ps1" -Tag Io {
    # リンク先にするファイル
    BeforeAll {
        writeText "$TestDrive\repo\README.md" "# 使い方`n"
        writeText "$TestDrive\repo\docs\images\図.png" ""
        writeText "$TestDrive\repo\docs\target.md" (@(
            "# 1 概要"
            "## 4.3 Excel の抽出処理（``ExtractWorkbook``）"
            "## 6.3 TSV 整形仕様（``PrettyTsv`` / ``FormatTsv``）"
            "## [リンク](README.md) の見出し"
            "## 重複"
            "## 重複"
            "``````"
            "## コードブロックの中"
            "``````"
            "<a id=""手で付けた""></a>"
        ) -join "`n")
    }

    It "あるファイル・フォルダ・見出しへのリンクを通す: <link>" -TestCases @(
        @{ link = "[a](target.md)" }
        @{ link = "[a](./target.md)" }
        @{ link = "[a](../README.md)" }
        @{ link = "[a](/README.md)" }
        @{ link = "[a](images/)" }
        @{ link = "![図](images/図.png)" }
        @{ link = "![図](images/%E5%9B%B3.png)" }
        @{ link = "[a](<images/図.png> ""題"")" }
        @{ link = "[a](target.md#1-概要)" }
        @{ link = "[a](target.md#43-excel-の抽出処理extractworkbook)" }
        @{ link = "[a](target.md#63-tsv-整形仕様prettytsv--formattsv)" }
        @{ link = "[a](target.md#リンク-の見出し)" }
        @{ link = "[a](target.md#重複-1)" }
        @{ link = "[a](target.md#手で付けた)" }
        @{ link = "[a](#自分の見出し)`n## 自分の見出し" }
        @{ link = "[名前]: target.md#1-概要" }
        @{ link = "<a href=""../README.md"">a</a>" }
        @{ link = "[a](https://example.com/no/such/page) [b](mailto:a@example.com)" }
    ) {
        param($link)
        checkPage $link | Should -BeNullOrEmpty
    }

    It "コードの中のリンクは調べない" {
        checkPage "``[a](none.md)```n``````md`n[a](none.md)`n```````n~~~`n[a](none.md)`n~~~`n" | Should -BeNullOrEmpty
    }

    It "切れたリンクを止める: <link>" -TestCases @(
        @{ link = "[a](none.md)"; reason = "ファイルが無い" }
        @{ link = "[a](Target.md)"; reason = "ファイルが無い" }
        @{ link = "![図](Images/図.png)"; reason = "ファイルが無い" }
        @{ link = "[名前]: none.md"; reason = "ファイルが無い" }
        @{ link = "<img src=""images/none.png"">"; reason = "ファイルが無い" }
        @{ link = "[a](../../outside.md)"; reason = "リポジトリの外" }
        @{ link = "[a](target.md#無い見出し)"; reason = "見出し" }
        @{ link = "[a](target.md#重複-2)"; reason = "見出し" }
        @{ link = "[a](target.md#コードブロックの中)"; reason = "見出し" }
        @{ link = "[a](#無い見出し)"; reason = "見出し" }
    ) {
        param($link, $reason)
        $problems = @(checkPage "# 題`n`n$link`n")
        $problems.Count | Should -Be 1
        $problems[0] | Should -Match "^docs/page\.md:3: .*$reason"
    }

    It "1 行に複数のリンクがあれば、それぞれ調べる" {
        @(checkPage "[a](none1.md) [b](target.md) [c](none2.md)").Count | Should -Be 2
    }
}
