# 配布物の部品表を作る tools\new_sbom.ps1 のテスト
BeforeAll {
    $rootDir = (Resolve-Path "$PSScriptRoot\..\..").Path
    $newSbom = "$rootDir\tools\new_sbom.ps1"
    $time = [DateTimeOffset]::FromUnixTimeSeconds(1790000000)
    $sha = "0123456789abcdef0123456789abcdef01234567"
    function New-Entries {
        $entries = [ordered]@{}
        $entries["scripts/a.ps1"] = [System.Text.Encoding]::UTF8.GetBytes("a")
        $entries["tebunko.bat"] = [System.Text.Encoding]::UTF8.GetBytes("b")
        $entries
    }
    function Get-Sbom([string]$Version, [string]$Sha) {
        $bytes = & $newSbom -Version $Version -Sha $Sha -Timestamp $time -Entries (New-Entries)
        [System.Text.Encoding]::UTF8.GetString($bytes) | ConvertFrom-Json
    }
}

Describe "new_sbom.ps1" -Tag Unit {
    It "同じ入力なら同じバイト列になる（BOM 無し UTF-8）" {
        $first = & $newSbom -Version "v1.0.0" -Sha $sha -Timestamp $time -Entries (New-Entries)
        $second = & $newSbom -Version "v1.0.0" -Sha $sha -Timestamp $time -Entries (New-Entries)
        [System.Convert]::ToBase64String($first) | Should -Be ([System.Convert]::ToBase64String($second))
        $first[0] | Should -Be ([byte][char]"{")
    }

    It "版が違えば serialNumber が違う" {
        (Get-Sbom "v1.0.0" $sha).serialNumber | Should -Not -Be (Get-Sbom "v1.0.1" $sha).serialNumber
    }

    It "SHA が違えば serialNumber が違う" {
        (Get-Sbom "v1.0.0" $sha).serialNumber | Should -Not -Be (Get-Sbom "v1.0.0" ("f" * 40)).serialNumber
    }

    It "serialNumber が CycloneDX 1.6 の形に合う（<Version> / <Sha>）" -TestCases @(
        @{ Version = "v0.2.0"; Sha = "0123456789abcdef0123456789abcdef01234567" }
        @{ Version = "v1.0.0"; Sha = "ffffffffffffffffffffffffffffffffffffffff" }
        @{ Version = "v9.9.9"; Sha = "0000000000000000000000000000000000000000" }
        @{ Version = "v10.20.30"; Sha = "deadbeefdeadbeefdeadbeefdeadbeefdeadbeef" }
        @{ Version = "v0.0.1-rc.1"; Sha = "89abcdef0123456789abcdef0123456789abcdef" }
    ) {
        (Get-Sbom $Version $Sha).serialNumber -cmatch '^urn:uuid:[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$' | Should -Be $true
    }

    It "形式・版・時刻・ライセンスと、components・dependencies の対応が正しい" {
        $sbom = Get-Sbom "v1.0.0" $sha
        $sbom.bomFormat | Should -Be "CycloneDX"
        $sbom.specVersion | Should -Be "1.6"
        $sbom.metadata.timestamp | Should -Be $time.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ", [System.Globalization.CultureInfo]::InvariantCulture)
        $sbom.metadata.component.version | Should -Be "v1.0.0"
        $sbom.metadata.component.licenses[0].license.id | Should -Be "MIT"
        @($sbom.metadata.properties | Where-Object { $_.name -eq "tebunko:prerequisite" }).Count -gt 0 | Should -Be $true
        @($sbom.components | ForEach-Object { $_."bom-ref" }) | Should -Be @("scripts/a.ps1", "tebunko.bat")
        @($sbom.dependencies).Count | Should -Be 1
        $sbom.dependencies[0].ref | Should -Be $sbom.metadata.component."bom-ref"
        @($sbom.dependencies[0].dependsOn) | Should -Be @($sbom.components | ForEach-Object { $_."bom-ref" })
        $sbom.components[0].hashes[0].alg | Should -Be "SHA-256"
        # "a" の SHA-256（小文字）
        $sbom.components[0].hashes[0].content | Should -Be "ca978112ca1bbdcafac231b39a23dc4da786eff8147c4e72b9807785afee48bb"
    }

    It "現在のカルチャに依らず、timestamp の書式が同じになる（<Culture>）" -TestCases @(
        @{ Culture = "th-TH" }
        @{ Culture = "ja-JP" }
        @{ Culture = "ar-SA" }
    ) {
        $original = [System.Threading.Thread]::CurrentThread.CurrentCulture
        try {
            [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo($Culture)
            (Get-Sbom "v1.0.0" $sha).metadata.timestamp | Should -Be $time.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ", [System.Globalization.CultureInfo]::InvariantCulture)
        } finally {
            [System.Threading.Thread]::CurrentThread.CurrentCulture = $original
        }
    }

    It "第三者の部品（purl を持つもの・group が tebunko 以外）を 1 件も含まない" {
        $sbom = Get-Sbom "v1.0.0" $sha
        (@($sbom.components).Count -gt 0) | Should -Be $true
        @($sbom.components | Where-Object { $_.group -ne "tebunko" }).Count | Should -Be 0
        @($sbom.components | Where-Object { $_.purl }).Count | Should -Be 0
    }
}
