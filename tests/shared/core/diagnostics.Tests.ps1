# 診断情報の収集とログの世代の入れ替え（shared\core\diagnostics.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "getPreviousLogPath" -Tag Io {
    It "同じフォルダの <名前>_previous<拡張子>" {
        getPreviousLogPath 'C:\w\gui_error_log.txt' | Should -Be 'C:\w\gui_error_log_previous.txt'
    }
}

Describe "rotateLogFile" -Tag Io {
    BeforeEach {
        $script:log = Join-Path $TestDrive "x_log.txt"
        $script:prev = Join-Path $TestDrive "x_log_previous.txt"
        Remove-Item -LiteralPath $script:log, $script:prev -Force -ErrorAction SilentlyContinue
    }

    It "上限未満なら移さない" {
        [System.IO.File]::WriteAllText($script:log, "abc")
        rotateLogFile $script:log 100 | Should -Be $false
        [System.IO.File]::Exists($script:log) | Should -Be $true
        [System.IO.File]::Exists($script:prev) | Should -Be $false
    }

    It "上限以上なら前回の分へ移す" {
        [System.IO.File]::WriteAllText($script:log, "abcde")
        rotateLogFile $script:log 5 | Should -Be $true
        [System.IO.File]::Exists($script:log) | Should -Be $false
        [System.IO.File]::ReadAllText($script:prev) | Should -Be "abcde"
    }

    It "前回の分があれば上書きする" {
        [System.IO.File]::WriteAllText($script:prev, "古い")
        [System.IO.File]::WriteAllText($script:log, "新しい")
        rotateLogFile $script:log 1 | Should -Be $true
        [System.IO.File]::ReadAllText($script:prev) | Should -Be "新しい"
    }

    It "上限 0 はいつも移す" {
        [System.IO.File]::WriteAllText($script:log, "a")
        rotateLogFile $script:log 0 | Should -Be $true
        [System.IO.File]::Exists($script:prev) | Should -Be $true
    }

    It "元が無ければ何もしない" {
        rotateLogFile $script:log 0 | Should -Be $false
        [System.IO.File]::Exists($script:prev) | Should -Be $false
    }

    It "移せない（開いたまま）ときは例外を出さずに `$false" {
        [System.IO.File]::WriteAllText($script:log, "abc")
        $stream = [System.IO.File]::Open($script:log, "Open", "Read", "None")
        try {
            rotateLogFile $script:log 0 | Should -Be $false
        } finally {
            $stream.Dispose()
        }
        [System.IO.File]::Exists($script:log) | Should -Be $true
    }

    It "前回の分が読み取り専用で消せないときは例外を出さずに `$false" {
        [System.IO.File]::WriteAllText($script:prev, "古い")
        [System.IO.File]::SetAttributes($script:prev, [System.IO.FileAttributes]::ReadOnly)
        [System.IO.File]::WriteAllText($script:log, "新しい")
        try {
            rotateLogFile $script:log 0 | Should -Be $false
        } finally {
            [System.IO.File]::SetAttributes($script:prev, [System.IO.FileAttributes]::Normal)
        }
    }
}

Describe "getEnvironmentInfo" -Tag Io {
    It "OS・PowerShell の項目が空でない" {
        $info = getEnvironmentInfo
        $info.OsVersion | Should -Not -BeNullOrEmpty
        $info.PsVersion | Should -Not -BeNullOrEmpty
        $info.UICulture | Should -Not -BeNullOrEmpty
    }

    It "見出しの全文に、利用者名・コンピューター名・ドメイン・プロファイルのパスが含まれない" {
        $text = (formatDiagnosticHeader "tebunko" $null (getEnvironmentInfo) ([System.DateTimeOffset]::Now)) -join "`n"
        foreach ($value in @($env:USERNAME, $env:COMPUTERNAME, $env:USERDOMAIN)) {
            if ($value -and $value.Length -ge 3) {
                $text.IndexOf($value, [System.StringComparison]::OrdinalIgnoreCase) | Should -Be -1
            }
        }
        $text.IndexOf($env:USERPROFILE, [System.StringComparison]::OrdinalIgnoreCase) | Should -Be -1
    }
}
