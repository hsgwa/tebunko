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

    It "空のファイルなら `$null" {
        $path = Join-Path $TestDrive "empty.txt"
        [System.IO.File]::WriteAllBytes($path, [byte[]]@())
        readVersionFile $path | Should -BeNullOrEmpty
    }

    It "1 行だけなら `$null" {
        $path = Join-Path $TestDrive "oneline.txt"
        writeVersionFileBytes $path @("v1.0.0")
        readVersionFile $path | Should -BeNullOrEmpty
    }

    It "3 行あれば `$null" {
        $path = Join-Path $TestDrive "threelines.txt"
        writeVersionFileBytes $path @("v1.0.0", "0123456789abcdef0123456789abcdef01234567", "")
        readVersionFile $path | Should -BeNullOrEmpty
    }

    It "1 行目（タグ名）の形が違えば `$null" {
        $path = Join-Path $TestDrive "badtag.txt"
        writeVersionFileBytes $path @("v1.0.0 だめ", "0123456789abcdef0123456789abcdef01234567")
        readVersionFile $path | Should -BeNullOrEmpty
    }

    It "2 行目（SHA）の形が違えば `$null" {
        $path = Join-Path $TestDrive "badsha.txt"
        writeVersionFileBytes $path @("v1.0.0", "abc")
        readVersionFile $path | Should -BeNullOrEmpty
    }

    It "正しい形（BOM 付き UTF-8・CRLF）なら Tag・Sha を返す" {
        $path = Join-Path $TestDrive "ok.txt"
        writeVersionFileBytes $path @("v1.0.0", "0123456789abcdef0123456789abcdef01234567")
        $result = readVersionFile $path
        $result.Tag | Should -Be "v1.0.0"
        $result.Sha | Should -Be "0123456789abcdef0123456789abcdef01234567"
    }

    It "BOM が無く LF 区切りでも、2 行の形が正しければ返す" {
        $path = Join-Path $TestDrive "nobombutlf.txt"
        writeVersionFileBytes $path @("v1.0.0", "0123456789abcdef0123456789abcdef01234567") $false "`n"
        $result = readVersionFile $path
        $result.Tag | Should -Be "v1.0.0"
        $result.Sha | Should -Be "0123456789abcdef0123456789abcdef01234567"
    }
}
