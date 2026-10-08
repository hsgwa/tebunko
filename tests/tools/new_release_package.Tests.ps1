# 配布する zip（tools\new_release_package.ps1）と、その中の VERSION.txt（tools\new_version_text.ps1）のテスト
BeforeAll {
    $rootDir = (Resolve-Path "$PSScriptRoot\..\..").Path
    $newVersionText = "$rootDir\tools\new_version_text.ps1"
    $newReleasePackage = "$rootDir\tools\new_release_package.ps1"
    $headSha = (& git -C $rootDir rev-parse HEAD).Trim()
    $commitTime = [DateTimeOffset]::FromUnixTimeSeconds([long](& git -C $rootDir log -1 --format=%ct HEAD))
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
        $zipFile = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
        try {
            $zipEntries = @($zipFile.Entries | ForEach-Object { [pscustomobject]@{ FullName = $_.FullName; LastWriteTime = $_.LastWriteTime } })
        } finally {
            $zipFile.Dispose()
        }
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

    It "zip のエントリーが、git で追跡している scripts\ と固定のファイルに過不足なく一致し、序数の順に並ぶ" {
        $expected = @(@(& git -C $rootDir ls-files scripts) + @("tebunko.bat", "README.md", "LICENSE", "VERSION.txt") | ForEach-Object { "tebunko/$_" })
        $expectedSorted = [string[]]$expected
        [Array]::Sort($expectedSorted, [StringComparer]::Ordinal)
        @($zipEntries | ForEach-Object { $_.FullName }) | Should -Be $expectedSorted
    }

    It "全エントリーの時刻が、コミットの時刻の UTC の時計の値（2 秒以内。手元の時差に依らない）になっている" {
        $expectedClock = $commitTime.UtcDateTime
        @($zipEntries | Where-Object { [Math]::Abs(($_.LastWriteTime.DateTime - $expectedClock).TotalSeconds) -gt 2 }).Count | Should -Be 0
    }

    It "<Name> の改行は CRLF にそろっている（CR の無い LF も、CR CR も無い）" -TestCases @(
        @{ Name = "LICENSE" }
        @{ Name = "README.md" }
    ) {
        $text = [System.Text.Encoding]::UTF8.GetString([System.IO.File]::ReadAllBytes((Join-Path $extractDir "tebunko\$Name")))
        $text.Contains("`n") | Should -Be $true
        [regex]::IsMatch($text, "(?<!`r)`n") | Should -Be $false
        $text.Contains("`r`r") | Should -Be $false
    }

    It "同じコミットで時間を空けて 2 回作ると、zip・sbom.cdx.json・SHA256SUMS.txt がそれぞれ同じ中身になる" {
        # DOS 時刻は 2 秒単位のため、2 秒以上空けて偶然そろうことを避ける
        Start-Sleep -Seconds 3
        $outDir2 = Join-Path $TestDrive "out2"
        & $newReleasePackage -Version "v9.9.9" -OutDir $outDir2 | Out-Null
        foreach ($name in @("tebunko-v9.9.9.zip", "sbom.cdx.json", "SHA256SUMS.txt")) {
            (Get-FileHash -LiteralPath (Join-Path $outDir2 $name) -Algorithm SHA256).Hash |
                Should -Be (Get-FileHash -LiteralPath (Join-Path $outDir $name) -Algorithm SHA256).Hash
        }
    }

    It "現在のカルチャが th-TH でも、sbom.cdx.json・SHA256SUMS.txt が同じ中身になる" {
        $original = [System.Threading.Thread]::CurrentThread.CurrentCulture
        $outDir3 = Join-Path $TestDrive "out3"
        try {
            [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo("th-TH")
            & $newReleasePackage -Version "v9.9.9" -OutDir $outDir3 | Out-Null
        } finally {
            [System.Threading.Thread]::CurrentThread.CurrentCulture = $original
        }
        foreach ($name in @("sbom.cdx.json", "SHA256SUMS.txt")) {
            (Get-FileHash -LiteralPath (Join-Path $outDir3 $name) -Algorithm SHA256).Hash |
                Should -Be (Get-FileHash -LiteralPath (Join-Path $outDir $name) -Algorithm SHA256).Hash
        }
    }

    It "部品表の components が zip のエントリーと過不足なく一致し、各 SHA-256 が zip の中身と一致する" {
        $sbom = [System.IO.File]::ReadAllText((Join-Path $outDir "sbom.cdx.json")) | ConvertFrom-Json
        $sbom.bomFormat | Should -Be "CycloneDX"
        $sbom.specVersion | Should -Be "1.6"
        $sbom.version | Should -Be 1
        $sbom.metadata.component.version | Should -Be "v9.9.9"
        ($sbom.serialNumber -cmatch '^urn:uuid:[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$') | Should -Be $true
        $sbom.metadata.timestamp | Should -Be $commitTime.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ", [System.Globalization.CultureInfo]::InvariantCulture)
        $fileComponents = @($sbom.components | Where-Object { $_.type -eq "file" })
        @($fileComponents | ForEach-Object { "tebunko/" + $_.name }) | Should -Be @($zipEntries | ForEach-Object { $_.FullName })
        foreach ($component in $fileComponents) {
            $path = Join-Path $extractDir ("tebunko\" + $component.name.Replace("/", "\"))
            $component.hashes[0].alg | Should -Be "SHA-256"
            $component.hashes[0].content | Should -Be (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLower()
        }
    }

    It "SHA256SUMS.txt の行が zip のエントリーと sbom.cdx.json に過不足なく対応し、ハッシュが中身と一致する" {
        $sumsPath = Join-Path $outDir "SHA256SUMS.txt"
        $lines = @([System.IO.File]::ReadAllLines($sumsPath) | Where-Object { $_ -and !$_.StartsWith("#") })
        $expectedPaths = @(@($zipEntries | ForEach-Object { $_.FullName.Substring("tebunko/".Length).Replace("/", "\") }) + @("sbom.cdx.json"))
        @($lines | ForEach-Object { $_.Substring(66) }) | Should -Be $expectedPaths
        foreach ($line in $lines) {
            $relative = $line.Substring(66)
            $path = if ($relative -eq "sbom.cdx.json") { Join-Path $outDir $relative } else { Join-Path $extractDir ("tebunko\" + $relative) }
            $line.Substring(0, 64) | Should -Be (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        }
        [System.IO.File]::ReadAllText($sumsPath).Contains($commitTime.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ", [System.Globalization.CultureInfo]::InvariantCulture)) | Should -Be $true
    }

    It "実行後、リポジトリ直下に VERSION.txt ができていない" {
        Test-Path -LiteralPath (Join-Path $rootDir "VERSION.txt") | Should -Be $false
    }
}
