# PSScriptAnalyzer の全体の検査（tools\check_analyzer.ps1）の設定とベースラインの形を確かめる。
# PSScriptAnalyzer を実際にかけるのは CI の手順だけにする（全部で 90 秒ほどかかり、コミット前のフックで待たせないため）。
# ここで確かめるのは、設定から規則を消したり、プロファイルを空にしたりして「指摘 0 件」にする抜け道が無いこと。
Describe "PSScriptAnalyzer の全体の検査の設定（docs/safety/scans.md「静的解析: PSScriptAnalyzer（Microsoft）」）" -Tag Meta {
    BeforeAll {
        $metaDir = $PSScriptRoot
        $rootDir = (Resolve-Path "$PSScriptRoot\..\..").Path
        $main = Import-LocalizedData -BaseDirectory $metaDir -FileName "PSScriptAnalyzer.psd1"
        $testsConfig = Import-LocalizedData -BaseDirectory $metaDir -FileName "PSScriptAnalyzer.tests.psd1"
        $common = @("PSAvoidUsingPositionalParameters", "PSReviewUnusedParameter", "PSAvoidAssignmentToAutomaticVariable", "PSUseDeclaredVarsMoreThanAssignments", "TypeNotFound")
        $scans = [System.IO.File]::ReadAllText("$rootDir\docs\safety\scans.md", [System.Text.Encoding]::UTF8)
    }

    It "除く規則は決めた一覧のとおりで、scans.md に理由が載っている: <name>" -TestCases @(
        @{ name = "scripts・tools 用"; key = "main"; extra = @() }
        @{ name = "tests 用"; key = "tests"; extra = @("PSAvoidUsingWriteHost", "PSAvoidGlobalVars") }
    ) {
        param ($name, $key, $extra)
        $config = if ($key -eq "main") { $main } else { $testsConfig }
        $expected = @($common) + @($extra)
        (@($config.ExcludeRules) | Sort-Object) -join "," | Should -Be (($expected | Sort-Object) -join ",")
        $config.IncludeDefaultRules | Should -Be $true
        foreach ($rule in $expected) {
            if ($rule -eq "TypeNotFound") { continue }
            $scans.Contains("``$rule``") | Should -Be $true -Because "$rule を除く理由を docs\safety\scans.md に書く"
        }
    }

    It "5.1 に無い構文を止める規則が、対象 5.1 で有効" -TestCases @(
        @{ name = "scripts・tools 用"; key = "main" }
        @{ name = "tests 用"; key = "tests" }
    ) {
        param ($name, $key)
        $config = if ($key -eq "main") { $main } else { $testsConfig }
        $config.Rules.PSUseCompatibleSyntax.Enable | Should -Be $true
        (@($config.Rules.PSUseCompatibleSyntax.TargetVersions) -join ",") | Should -Be "5.1"
        foreach ($rule in "PSAvoidSemicolonsAsLineTerminators", "PSPlaceOpenBrace") {
            $config.Rules[$rule].Enable | Should -Be $true
        }
    }

    It "5.1 に無いコマンドを止める規則が、Windows PowerShell 5.1 のプロファイルで有効（tests 用は入れない）" {
        $main.Rules.PSUseCompatibleCommands.Enable | Should -Be $true
        (@($main.Rules.PSUseCompatibleCommands.TargetProfiles) -join ",") | Should -Be "win-48_x64_10.0.17763.0_5.1.17763.316_x64_4.0.30319.42000_framework"
        $testsConfig.Rules.ContainsKey("PSUseCompatibleCommands") | Should -Be $false
    }

    It "制限言語モードの規則は足さない（safety.Tests.ps1 が受け持つ）" {
        $main.Rules.ContainsKey("PSUseConstrainedLanguageMode") | Should -Be $false
        $testsConfig.Rules.ContainsKey("PSUseConstrainedLanguageMode") | Should -Be $false
    }

    It "ベースラインは 3 列（ファイル・規則・正の整数の件数）で、ファイル・規則の順に並び、重複が無い" {
        $text = [System.IO.File]::ReadAllText("$metaDir\PSScriptAnalyzer.baseline.tsv", [System.Text.Encoding]::UTF8)
        $lines = @($text -split "\r?\n" | Where-Object { $_ })
        $lines.Count | Should -BeGreaterThan 0
        $keys = New-Object System.Collections.Generic.List[string]
        foreach ($line in $lines) {
            $cols = $line.Split("`t")
            $cols.Count | Should -Be 3 -Because $line
            $cols[2] | Should -Match '^[1-9][0-9]*$' -Because $line
            $cols[0] | Should -Match '^(scripts|tools|tests)/' -Because $line
            $cols[0].Contains("\") | Should -Be $false -Because $line
            $keys.Add("$($cols[0])`t$($cols[1])")
        }
        $sorted = @($keys); [array]::Sort($sorted, [System.StringComparer]::Ordinal)
        ($keys -join "|") | Should -Be ($sorted -join "|")
        @($keys | Select-Object -Unique).Count | Should -Be $keys.Count
    }

    It "ベースラインのファイルが、どれも存在する" {
        $text = [System.IO.File]::ReadAllText("$metaDir\PSScriptAnalyzer.baseline.tsv", [System.Text.Encoding]::UTF8)
        $missing = @($text -split "\r?\n" | Where-Object { $_ } | ForEach-Object { $_.Split("`t")[0] } | Select-Object -Unique | Where-Object { !(Test-Path -LiteralPath "$rootDir\$_") })
        $missing -join "," | Should -Be ""
    }
}
