# PSScriptAnalyzer の指摘をベースラインと比べる関数（tools\analyzer_baseline.ps1。tools\check_analyzer.ps1 が使う）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\tools\analyzer_baseline.ps1"

    $script:root = "C:\repo"

    function newFinding([string]$file, [string]$rule, [int]$line = 1) {
        [pscustomobject]@{ ScriptPath = "$script:root\$file"; RuleName = $rule; Line = $line; Message = "内容" }
    }

    # 「ファイル<TAB>規則」→ 件数 のベースラインを作る
    function newBaseline([string]$text) { readAnalyzerBaseline $text }
}

Describe "getAnalyzerCounts" -Tag Unit {
    It "ファイル・規則ごとに数え、パスはリポジトリからの相対で / 区切りにする" {
        $counts = getAnalyzerCounts @(
            (newFinding "scripts\a.ps1" "PSX" 3)
            (newFinding "scripts\a.ps1" "PSX" 9)
            (newFinding "scripts\a.ps1" "PSY" 1)
            (newFinding "tools\b.ps1" "PSX" 2)
        ) $script:root
        $counts.Count | Should -Be 3
        $counts["scripts/a.ps1`tPSX"].Count | Should -Be 2
        $counts["scripts/a.ps1`tPSX"][0] | Should -BeLike "PSX scripts/a.ps1:3 *"
        $counts["tools/b.ps1`tPSX"].Count | Should -Be 1
    }

    It "指摘が無ければ空" {
        (getAnalyzerCounts @() $script:root).Count | Should -Be 0
    }
}

Describe "ベースラインの読み書き" -Tag Unit {
    It "書いたものを読み直すと同じ件数になり、並びはファイル・規則の順で CRLF 区切り" {
        $counts = getAnalyzerCounts @(
            (newFinding "tools\b.ps1" "PSY")
            (newFinding "scripts\a.ps1" "PSZ")
            (newFinding "scripts\a.ps1" "PSX")
            (newFinding "scripts\a.ps1" "PSX")
        ) $script:root
        $text = formatAnalyzerBaseline $counts
        $text | Should -Be "scripts/a.ps1`tPSX`t2`r`nscripts/a.ps1`tPSZ`t1`r`ntools/b.ps1`tPSY`t1`r`n"
        $back = readAnalyzerBaseline $text
        $back.Count | Should -Be 3
        $back["scripts/a.ps1`tPSX"] | Should -Be 2
    }

    It "空のときは空の本文" {
        formatAnalyzerBaseline @{} | Should -Be ""
        (readAnalyzerBaseline "").Count | Should -Be 0
    }

    It "3 列でない行は読めない" {
        { readAnalyzerBaseline "a`tb" } | Should -Throw
    }
}

Describe "compareAnalyzerBaseline" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "同じ件数なら何も出さない"; baseline = "scripts/a.ps1`tPSX`t2"; findings = @(@("scripts\a.ps1", "PSX"), @("scripts\a.ps1", "PSX")); increased = 0; decreased = 0 }
        @{ name = "増えたら、そのファイル・規則の指摘を全部出す"; baseline = "scripts/a.ps1`tPSX`t1"; findings = @(@("scripts\a.ps1", "PSX"), @("scripts\a.ps1", "PSX")); increased = 2; decreased = 0 }
        @{ name = "減ったら、書き直す行を出す"; baseline = "scripts/a.ps1`tPSX`t2"; findings = @(, @("scripts\a.ps1", "PSX")); increased = 0; decreased = 1 }
        @{ name = "指摘が無くなったら減ったとみなす"; baseline = "scripts/a.ps1`tPSX`t1"; findings = @(); increased = 0; decreased = 1 }
        @{ name = "ベースラインに無いファイルは増えたとみなす"; baseline = "scripts/a.ps1`tPSX`t1"; findings = @(@("scripts\a.ps1", "PSX"), @("scripts\b.ps1", "PSX")); increased = 1; decreased = 0 }
        @{ name = "ベースラインに無い規則は増えたとみなす"; baseline = "scripts/a.ps1`tPSX`t1"; findings = @(@("scripts\a.ps1", "PSX"), @("scripts\a.ps1", "PSY")); increased = 1; decreased = 0 }
        @{ name = "ベースラインが空なら、指摘は全部増えたとみなす"; baseline = ""; findings = @(, @("tools\c.ps1", "PSX")); increased = 1; decreased = 0 }
    ) {
        param ($name, $baseline, $findings, $increased, $decreased)
        $list = @($findings | ForEach-Object { newFinding $_[0] $_[1] 5 })
        $result = compareAnalyzerBaseline (newBaseline $baseline) (getAnalyzerCounts $list $script:root)
        $result.Increased.Count | Should -Be $increased
        $result.Decreased.Count | Should -Be $decreased
    }

    It "増えた指摘は「規則 ファイル:行」を含み、減った指摘は -UpdateBaseline を案内する" {
        $up = compareAnalyzerBaseline (newBaseline "") (getAnalyzerCounts @(newFinding "scripts\a.ps1" "PSX" 7) $script:root)
        $up.Increased[0] | Should -BeLike "PSX scripts/a.ps1:7 *"
        $down = compareAnalyzerBaseline (newBaseline "scripts/a.ps1`tPSX`t1") @{}
        $down.Decreased[0] | Should -BeLike "*-UpdateBaseline*"
    }
}
