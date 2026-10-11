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

    It "前回の分が読み取り専用で置き換えられないときは例外を出さずに `$false" {
        [System.IO.File]::WriteAllText($script:prev, "古い")
        [System.IO.File]::SetAttributes($script:prev, [System.IO.FileAttributes]::ReadOnly)
        [System.IO.File]::WriteAllText($script:log, "新しい")
        try {
            rotateLogFile $script:log 0 | Should -Be $false
        } finally {
            [System.IO.File]::SetAttributes($script:prev, [System.IO.FileAttributes]::Normal)
        }
        [System.IO.File]::ReadAllText($script:prev) | Should -Be "古い"
    }

    It "今のログがほかに読み取りの共有だけで開かれていて消せないとき、前回の分は変わらず元も残り、一時ファイルを残さない" {
        [System.IO.File]::WriteAllText($script:prev, "古い")
        [System.IO.File]::WriteAllText($script:log, "新しい")
        $stream = [System.IO.File]::Open($script:log, "Open", "Read", "ReadWrite")
        try {
            rotateLogFile $script:log 0 | Should -Be $false
        } finally {
            $stream.Dispose()
        }
        [System.IO.File]::ReadAllText($script:prev) | Should -Be "古い"
        [System.IO.File]::ReadAllText($script:log) | Should -Be "新しい"
        [System.IO.File]::Exists($script:prev + ".tmp") | Should -Be $false
    }

    It "今のログが開いたままで移せないとき、前回の分は消えずに残る" {
        [System.IO.File]::WriteAllText($script:prev, "古い")
        [System.IO.File]::WriteAllText($script:log, "新しい")
        $stream = [System.IO.File]::Open($script:log, "Open", "Read", "None")
        try {
            rotateLogFile $script:log 0 | Should -Be $false
        } finally {
            $stream.Dispose()
        }
        [System.IO.File]::ReadAllText($script:prev) | Should -Be "古い"
        [System.IO.File]::ReadAllText($script:log) | Should -Be "新しい"
    }
}

Describe "findOfficeApps" -Tag Io {
    BeforeAll {
        function newExe($base, $sub, $name) {
            $dir = Join-Path $base $sub
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            [System.IO.File]::WriteAllText((Join-Path $dir $name), "x")
        }
    }

    BeforeEach {
        $script:pf64 = Join-Path $TestDrive "pf64"
        $script:pf86 = Join-Path $TestDrive "pf86"
        Remove-Item -LiteralPath $script:pf64, $script:pf86 -Recurse -Force -ErrorAction SilentlyContinue
    }

    It "何も無ければ空" {
        @(findOfficeApps -programFiles64 $script:pf64 -programFilesX86 $script:pf86).Count | Should -Be 0
    }

    It "64 ビット側の Click-to-Run の Excel を見つける" {
        newExe $script:pf64 "Microsoft Office\root\Office16" "EXCEL.EXE"
        $apps = @(findOfficeApps -programFiles64 $script:pf64 -programFilesX86 $script:pf86)
        $apps.Count | Should -Be 1
        $apps[0].Name | Should -Be "Excel"
        $apps[0].Kind | Should -Be "Click-to-Run"
        $apps[0].Bits | Should -Be 64
    }

    It "x86 側の MSI の Word は 32 ビット" {
        newExe $script:pf86 "Microsoft Office\Office15" "WINWORD.EXE"
        $apps = @(findOfficeApps -programFiles64 $script:pf64 -programFilesX86 $script:pf86)
        $apps.Count | Should -Be 1
        $apps[0].Name | Should -Be "Word"
        $apps[0].Kind | Should -Be "MSI"
        $apps[0].Bits | Should -Be 32
    }

    It "32 ビットの Windows では Program Files の Office を 32 ビットとする" {
        newExe $script:pf64 "Microsoft Office\Office16" "POWERPNT.EXE"
        $apps = @(findOfficeApps -programFiles64 $script:pf64 -programFilesX86 $script:pf86 -os64 $false)
        $apps.Count | Should -Be 1
        $apps[0].Bits | Should -Be 32
    }
}

Describe "getEnvironmentInfo（Office を調べ損ねたとき）" -Tag Io {
    It "findOfficeApps が例外なら Office は `$null（見つからない `@() とは区別する）" {
        Mock findOfficeApps { throw "調べ損ね" }
        (getEnvironmentInfo).Office | Should -Be $null
    }

    It "findOfficeApps が空を返すなら Office は空の配列" {
        Mock findOfficeApps { return @() }
        $info = getEnvironmentInfo
        ($null -eq $info.Office) | Should -Be $false
        @($info.Office).Count | Should -Be 0
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
        # 見出しの決まった語と重なる名前（利用者名が見出しの語の一部など）は、取り込みの有無を見分けられないので除く
        $fixedText = (formatDiagnosticHeader "tebunko" $null @{ Office = @() } ([System.DateTimeOffset]::Now)) -join "`n"
        foreach ($value in @($env:USERNAME, $env:COMPUTERNAME, $env:USERDOMAIN)) {
            if ($value -and $value.Length -ge 3 -and $fixedText.IndexOf($value, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
                $text.IndexOf($value, [System.StringComparison]::OrdinalIgnoreCase) | Should -Be -1
            }
        }
        $text.IndexOf($env:USERPROFILE, [System.StringComparison]::OrdinalIgnoreCase) | Should -Be -1
    }
}
