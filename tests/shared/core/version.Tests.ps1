# 版の記録（shared\core\version.ps1 の readVersionFile）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    function writeVersionFileBytes {
        # BOM 付き UTF-8 で $lines（CRLF 区切り）を書く。$bom を $false にすると BOM なしで書ける
        param (
            [string]$path,
            [string[]]$lines,
            [bool]$bom = $true,
            [string]$newLine = "`r`n"
        )

        $content = ($lines -join $newLine)
        if ($lines.Count -gt 0) {
            $content += $newLine
        }
        $encoding = New-Object System.Text.UTF8Encoding($bom)
        [System.IO.File]::WriteAllBytes($path, $encoding.GetBytes($content))
    }
}

Describe "readVersionFile" -Tag Io {
    It "ファイルが無ければ `$null" {
        readVersionFile (Join-Path $TestDrive "無い\VERSION.txt") | Should -BeNullOrEmpty
    }

    It "<name>なら `$null" -TestCases @(
        @{ name = "空のファイル"; lines = @() }
        @{ name = "1 行だけ"; lines = @("v1.0.0") }
        @{ name = "3 行ある"; lines = @("v1.0.0", "0123456789abcdef0123456789abcdef01234567", "") }
        @{ name = "1 行目（タグ名）の形が違う"; lines = @("v1.0.0 だめ", "0123456789abcdef0123456789abcdef01234567") }
        @{ name = "2 行目（SHA）の形が違う"; lines = @("v1.0.0", "abc") }
    ) {
        $path = Join-Path $TestDrive "version.txt"
        writeVersionFileBytes $path $lines
        readVersionFile $path | Should -BeNullOrEmpty
    }

    It "<name>でも、2 行の形が正しければ Tag・Sha を返す" -TestCases @(
        @{ name = "BOM 付き UTF-8・CRLF"; bom = $true; newLine = "`r`n" }
        @{ name = "BOM が無く LF 区切り"; bom = $false; newLine = "`n" }
    ) {
        $path = Join-Path $TestDrive "ok.txt"
        writeVersionFileBytes $path @("v1.0.0", "0123456789abcdef0123456789abcdef01234567") $bom $newLine
        $result = readVersionFile $path
        $result.Tag | Should -Be "v1.0.0"
        $result.Sha | Should -Be "0123456789abcdef0123456789abcdef01234567"
    }
}
