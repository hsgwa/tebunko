# 開発の準備のスクリプト（tools\setup_dev.ps1）のテスト。
# 本物のインストールは流さない（Install-Module・フック・Python は Mock）。
BeforeAll {
    . "$PSScriptRoot\..\..\tools\setup_dev.ps1"

    function newWorkflow([string]$text) {
        $path = Join-Path $TestDrive ("test_" + [guid]::NewGuid().ToString("N") + ".yml")
        [System.IO.File]::WriteAllText($path, $text)
        return $path
    }
    $script:yml = "      PSSA_VERSION: 1.25.0`r`n        env:`r`n          PESTER_VERSION: 5.9.0`r`n"

    # runSetupDev の出力（情報・警告ストリーム）と終了コードを集める
    function runSetup([string]$workflow, [bool]$dryRun, [bool]$skipDocs) {
        $rc = $null
        $out = & { $script:rc = runSetupDev $TestDrive $workflow $dryRun $skipDocs } *>&1 | Out-String
        return @{ Out = $out; Rc = $script:rc }
    }
}

Describe "setup_dev.ps1" -Tag Unit {
    BeforeEach {
        Mock Install-Module { }
        Mock Get-Module { @() }
        Mock runHooksInstall { }
        Mock findPython { "" }
    }

    It "DryRun は 4 つの手順を出し、何も入れない" {
        Mock setupDocsVenv { }
        $r = runSetup (newWorkflow $script:yml) $true $false
        $r.Rc | Should -Be 0
        $r.Out | Should -Match "Pester"
        $r.Out | Should -Match "PSScriptAnalyzer"
        $r.Out | Should -Match "install_hooks"
        $r.Out | Should -Match "mkdocs"
        Should -Invoke Install-Module -Times 0 -Exactly
        Should -Invoke runHooksInstall -Times 0 -Exactly
        Should -Invoke setupDocsVenv -Times 0 -Exactly
    }

    It "版は test.yml から読む（PESTER_VERSION・PSSA_VERSION）" {
        $yml = "PSSA_VERSION: 9.8.7`nPESTER_VERSION: 6.5.4`n"
        $r = runSetup (newWorkflow $yml) $true $false
        $r.Out | Should -Match "Pester 6\.5\.4"
        $r.Out | Should -Match "PSScriptAnalyzer 9\.8\.7"
        $r.Out | Should -Not -Match "5\.9\.0"
    }

    It "test.yml に版の行が無ければ止まる" {
        $r = runSetup (newWorkflow "PESTER_VERSION: 5.9.0`n") $false $false
        $r.Rc | Should -Not -Be 0
        Should -Invoke Install-Module -Times 0 -Exactly
    }

    It "test.yml が無ければ止まる" {
        $r = runSetup (Join-Path $TestDrive "none.yml") $false $false
        $r.Rc | Should -Not -Be 0
    }

    It "入っている版と同じなら入れ直さない" {
        Mock Get-Module { [pscustomobject]@{ Name = $Name; Version = [version]"5.9.0" }, [pscustomobject]@{ Name = $Name; Version = [version]"1.25.0" } }
        $r = runSetup (newWorkflow $script:yml) $false $true
        $r.Rc | Should -Be 0
        Should -Invoke Install-Module -Times 0 -Exactly
        $r.Out | Should -Match "入っています"
    }

    It "入っていなければ版を指定して入れ、フックも有効にする" {
        Mock setupDocsVenv { }
        $r = runSetup (newWorkflow $script:yml) $false $true
        $r.Rc | Should -Be 0
        Should -Invoke Install-Module -Times 1 -Exactly -ParameterFilter { $Name -eq "Pester" -and $RequiredVersion -eq "5.9.0" }
        Should -Invoke Install-Module -Times 1 -Exactly -ParameterFilter { $Name -eq "PSScriptAnalyzer" -and $RequiredVersion -eq "1.25.0" }
        Should -Invoke runHooksInstall -Times 1 -Exactly
        Should -Invoke setupDocsVenv -Times 0 -Exactly
    }

    It "Python が無ければ MkDocs だけ注意を出して飛ばし、ほかは続ける" {
        $r = runSetup (newWorkflow $script:yml) $false $false
        $r.Rc | Should -Be 0
        $r.Out | Should -Match "Python が見つからない"
        Should -Invoke Install-Module -Times 2 -Exactly
        Should -Invoke runHooksInstall -Times 1 -Exactly
    }
}

Describe "setup_dev.ps1 の readPinnedVersion" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "インデントと CRLF があっても読む"; text = "    PSSA_VERSION: 1.25.0`r`n"; key = "PSSA_VERSION"; expected = "1.25.0" }
        @{ name = "別の名前の行は読まない"; text = "PESTER_VERSION: 5.9.0`n"; key = "PSSA_VERSION"; expected = "" }
        @{ name = "版でない値は読まない"; text = "PSSA_VERSION: latest`n"; key = "PSSA_VERSION"; expected = "" }
    ) {
        param($name, $text, $key, $expected)
        readPinnedVersion $text $key | Should -Be $expected
    }
}
