# 配布する zip（tools\new_release_package.ps1）と、その中の VERSION.txt（tools\new_version_text.ps1）のテスト
BeforeAll {
    $rootDir = (Resolve-Path "$PSScriptRoot\..\..").Path
    $newVersionText = "$rootDir\tools\new_version_text.ps1"
    $newReleasePackage = "$rootDir\tools\new_release_package.ps1"
    $headSha = (& git -C $rootDir rev-parse HEAD).Trim()
}

Describe "new_version_text.ps1" -Tag Io {
    It "タグ名とコミットの SHA の2行（BOM 付き UTF-8・CRLF）を返す" {
        $bytes = & $newVersionText -Version "v9.9.9"
        ($bytes[0], $bytes[1], $bytes[2]) | Should -Be @(0xEF, 0xBB, 0xBF)
        $text = [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
        $text | Should -Be "v9.9.9`r`n$headSha`r`n"
    }

    It "バージョンに使えない文字が含まれていれば throw する" {
        { & $newVersionText -Version "bad version" } | Should -Throw
    }
}

Describe "new_release_package.ps1" -Tag Io {
    BeforeAll {
        $outDir = Join-Path $TestDrive "out"
        & $newReleasePackage -Version "v9.9.9" -OutDir $outDir | Out-Null
        $zipPath = Join-Path $outDir "tebunko-v9.9.9.zip"

        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $extractDir = Join-Path $TestDrive "extract"
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $extractDir)
    }

    It "zip ができる" {
        Test-Path -LiteralPath $zipPath | Should -Be $true
    }

    It "VERSION.txt に、タグ名と現在のコミットの SHA が入っている（BOM 付き UTF-8・CRLF）" {
        $versionPath = Join-Path $extractDir "tebunko\VERSION.txt"
        Test-Path -LiteralPath $versionPath | Should -Be $true
        $bytes = [System.IO.File]::ReadAllBytes($versionPath)
        ($bytes[0], $bytes[1], $bytes[2]) | Should -Be @(0xEF, 0xBB, 0xBF)
        $text = [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
        $text | Should -Be "v9.9.9`r`n$headSha`r`n"
    }

    It "既存の中身（tebunko.bat・scripts\・README.md・LICENSE）もそのまま入っている" {
        Test-Path -LiteralPath (Join-Path $extractDir "tebunko\tebunko.bat") | Should -Be $true
        Test-Path -LiteralPath (Join-Path $extractDir "tebunko\scripts") | Should -Be $true
        Test-Path -LiteralPath (Join-Path $extractDir "tebunko\README.md") | Should -Be $true
        Test-Path -LiteralPath (Join-Path $extractDir "tebunko\LICENSE") | Should -Be $true
    }

    It "実行後、リポジトリ直下に VERSION.txt ができていない" {
        Test-Path -LiteralPath (Join-Path $rootDir "VERSION.txt") | Should -Be $false
    }
}
