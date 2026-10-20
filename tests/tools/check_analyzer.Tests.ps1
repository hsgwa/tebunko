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

Describe "selectAnalyzerFindings" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "TypeNotFound は除く"; file = "scripts\a.ps1"; rule = "TypeNotFound"; kept = 0 }
        @{ name = "テキスト抽出の見本は除く"; file = "tests\testdata\text\TXT41.psm1"; rule = "PSX"; kept = 0 }
        @{ name = "区切りが / のパスの見本も除く"; file = "tests/testdata/text/TXT42.psd1"; rule = "PSX"; kept = 0 }
        @{ name = "大文字小文字が違うパスの見本も除く"; file = "Tests\TestData\Text\a.ps1"; rule = "PSX"; kept = 0 }
        @{ name = "ほかの testdata は残す"; file = "tests\testdata\office\a.ps1"; rule = "PSX"; kept = 1 }
        @{ name = "ふつうの指摘は残す"; file = "scripts\a.ps1"; rule = "PSX"; kept = 1 }
    ) {
        param ($name, $file, $rule, $kept)
        $f = [pscustomobject]@{ ScriptPath = "$script:root\$file"; RuleName = $rule; Line = 1; Message = "内容" }
        @(selectAnalyzerFindings @($f)).Count | Should -Be $kept
    }

    It "指摘が無ければ空の配列を返す" {
        @(selectAnalyzerFindings @()).Count | Should -Be 0
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

Describe "selectAnalyzerFiles" -Tag Unit {
    BeforeAll {
        $script:work = Join-Path ([System.IO.Path]::GetTempPath()) ("analyzer_files_" + [guid]::NewGuid().ToString("N"))
        foreach ($d in "scripts", "tools", "tests", "docs") { New-Item -ItemType Directory -Path (Join-Path $script:work $d) -Force | Out-Null }
        foreach ($f in "scripts\a.ps1", "tools\b.ps1", "tests\c.Tests.ps1", "scripts\x.xaml", "docs\d.ps1") { Set-Content -LiteralPath (Join-Path $script:work $f) -Value "" }
    }
    AfterAll { Remove-Item -LiteralPath $script:work -Recurse -Force -ErrorAction SilentlyContinue }

    It "<name>" -TestCases @(
        @{ name = "scripts・tools・tests の .ps1 を / 区切りで返す"; files = @("scripts\a.ps1", "tools/b.ps1", "tests/c.Tests.ps1"); expected = "scripts/a.ps1|tools/b.ps1|tests/c.Tests.ps1" }
        @{ name = "ほかのフォルダ・.ps1 でないものは飛ばす"; files = @("docs/d.ps1", "scripts/x.xaml", "README.md"); expected = "" }
        @{ name = "消した・名前を変えて無いファイルは飛ばす"; files = @("scripts/gone.ps1"); expected = "" }
        @{ name = "同じファイルは 1 回だけ"; files = @("scripts/a.ps1", "scripts\a.ps1"); expected = "scripts/a.ps1" }
    ) {
        param($name, $files, $expected)
        (@(selectAnalyzerFiles $files $script:work) -join "|") | Should -Be $expected
    }
}

Describe "filterAnalyzerBaseline" -Tag Unit {
    It "指定したファイルの行だけ残す（区切りと大文字小文字は問わない）" {
        $baseline = newBaseline "scripts/a.ps1`tPSX`t2`r`nscripts/a.ps1`tPSY`t1`r`ntools/b.ps1`tPSX`t3"
        $filtered = filterAnalyzerBaseline $baseline @("Scripts\A.ps1")
        $filtered.Count | Should -Be 2
        $filtered["scripts/a.ps1`tPSX"] | Should -Be 2
    }

    It "指定が無ければ空" {
        (filterAnalyzerBaseline (newBaseline "scripts/a.ps1`tPSX`t2") @()).Count | Should -Be 0
    }
}

# check_analyzer.ps1 自身を、一時フォルダの小さなリポジトリに対して powershell.exe で流す
Describe "check_analyzer.ps1 -Path" -Tag Io {
    BeforeAll {
        $script:hasAnalyzer = [bool](Get-Module -ListAvailable -Name PSScriptAnalyzer)
        $script:checker = (Resolve-Path "$PSScriptRoot\..\..\tools\check_analyzer.ps1").Path
        $metaDir = (Resolve-Path "$PSScriptRoot\..\meta").Path
        $script:work = Join-Path ([System.IO.Path]::GetTempPath()) ("analyzer_path_" + [guid]::NewGuid().ToString("N"))
        New-Item -ItemType Directory -Path (Join-Path $script:work "tests\meta") -Force | Out-Null
        foreach ($d in "scripts", "tools") { New-Item -ItemType Directory -Path (Join-Path $script:work $d) -Force | Out-Null }
        Copy-Item "$metaDir\PSScriptAnalyzer.psd1", "$metaDir\PSScriptAnalyzer.tests.psd1" (Join-Path $script:work "tests\meta")
        $script:baselinePath = Join-Path $script:work "tests\meta\PSScriptAnalyzer.baseline.tsv"

        # a.ps1 は Write-Host 1 件、b.ps1 は 2 件の指摘がある
        Set-Content -LiteralPath (Join-Path $script:work "scripts\a.ps1") -Value 'Write-Host "x"'
        Set-Content -LiteralPath (Join-Path $script:work "tools\b.ps1") -Value "Write-Host `"x`"`r`nWrite-Host `"y`""

        function runChecker([string]$baseline, [string[]]$extra) {
            [System.IO.File]::WriteAllText($script:baselinePath, $baseline)
            $null = & powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File $script:checker -Root $script:work @extra
            return $LASTEXITCODE
        }
    }
    AfterAll { Remove-Item -LiteralPath $script:work -Recurse -Force -ErrorAction SilentlyContinue }

    It "<name>" -TestCases @(
        @{ name = "-Path のファイルの件数がベースラインどおりなら 0"; baseline = "scripts/a.ps1`tPSAvoidUsingWriteHost`t1`r`ntools/b.ps1`tPSAvoidUsingWriteHost`t2`r`n"; extra = @("-Path", "scripts/a.ps1"); code = 0 }
        @{ name = "-Path のファイルの指摘が増えたら 1"; baseline = "tools/b.ps1`tPSAvoidUsingWriteHost`t2`r`n"; extra = @("-Path", "scripts/a.ps1"); code = 1 }
        @{ name = "-Path のファイルの件数が減ったら 1"; baseline = "scripts/a.ps1`tPSAvoidUsingWriteHost`t2`r`n"; extra = @("-Path", "scripts/a.ps1"); code = 1 }
        @{ name = "-Path に無いファイルの増減は見ない"; baseline = "scripts/a.ps1`tPSAvoidUsingWriteHost`t1`r`n"; extra = @("-Path", "scripts/a.ps1"); code = 0 }
        @{ name = "カンマで分けた複数のファイルも検査する"; baseline = "scripts/a.ps1`tPSAvoidUsingWriteHost`t1`r`n"; extra = @("-Path", "scripts/a.ps1,tools/b.ps1"); code = 1 }
        @{ name = "無いファイルは飛ばす"; baseline = ""; extra = @("-Path", "scripts/gone.ps1"); code = 0 }
        @{ name = "-Path なしは全体を比べる（今までどおり）"; baseline = "scripts/a.ps1`tPSAvoidUsingWriteHost`t1`r`n"; extra = @(); code = 1 }
        @{ name = "-Path なしで全体がベースラインどおりなら 0"; baseline = "scripts/a.ps1`tPSAvoidUsingWriteHost`t1`r`ntools/b.ps1`tPSAvoidUsingWriteHost`t2`r`n"; extra = @(); code = 0 }
    ) {
        param($name, $baseline, $extra, $code)
        if (!$script:hasAnalyzer) { Set-ItResult -Skipped -Because "PSScriptAnalyzer が入っていない"; return }
        (runChecker $baseline $extra) | Should -Be $code
    }

    It "-Path と -UpdateBaseline を一緒に渡すと 2 で、ベースラインを書き換えない" {
        $text = "scripts/a.ps1`tPSAvoidUsingWriteHost`t9`r`n"
        (runChecker $text @("-Path", "scripts/a.ps1", "-UpdateBaseline")) | Should -Be 2
        [System.IO.File]::ReadAllText($script:baselinePath) | Should -Be $text
    }
}
