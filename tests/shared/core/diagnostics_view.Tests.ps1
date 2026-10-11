# 診断情報の見出しの組み立て（shared\core\diagnostics_view.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "getOfficeKind" -Tag Unit {
    It "<path> は <kind>・<bits> ビット" -ForEach @(
        @{ path = 'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE'; kind = "Click-to-Run"; bits = 64 }
        @{ path = 'C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE'; kind = "Click-to-Run"; bits = 32 }
        @{ path = 'C:\Program Files\Microsoft Office\Office16\WINWORD.EXE'; kind = "MSI"; bits = 64 }
        @{ path = 'C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE'; kind = "MSI"; bits = 32 }
        @{ path = 'C:\Program Files (x86)\Microsoft Office\Office15\POWERPNT.EXE'; kind = "MSI"; bits = 32 }
        @{ path = 'c:\program files (x86)\microsoft office\root\office16\excel.exe'; kind = "Click-to-Run"; bits = 32 }
    ) {
        $result = getOfficeKind $path 'C:\Program Files' 'C:\Program Files (x86)'
        $result.Kind | Should -Be $kind
        $result.Bits | Should -Be $bits
    }
}

Describe "hideUserProfile" -Tag Unit {
    It "<name>" -ForEach @(
        @{ name = "プロファイルを置き換える"; text = 'C:\Users\test\Documents\a.xlsx'; expected = '%USERPROFILE%\Documents\a.xlsx' }
        @{ name = "大文字小文字は区別しない"; text = 'c:\USERS\Test\a.txt'; expected = '%USERPROFILE%\a.txt' }
        @{ name = "プロファイルの外のパスはそのまま"; text = 'D:\共有\営業部\a.txt'; expected = 'D:\共有\営業部\a.txt' }
        @{ name = "前方だけ一致する別のフォルダは置き換えない"; text = 'C:\Users\test2\a.txt'; expected = 'C:\Users\test2\a.txt' }
        @{ name = "文の途中にあっても置き換える"; text = 'パス C:\Users\test を開けない'; expected = 'パス %USERPROFILE% を開けない' }
        @{ name = "空の文字列"; text = ''; expected = '' }
    ) {
        hideUserProfile $text 'C:\Users\test' | Should -Be $expected
    }
}

Describe "formatDiagnosticHeader" -Tag Unit {
    BeforeAll {
        $script:now = [System.DateTimeOffset]::new(2026, 10, 11, 9, 30, 0, [System.TimeSpan]::FromHours(9))
        $script:env = @{
            OsVersion = "10.0.26200"; OsFileVersion = "10.0.26100.1234"; Os64 = $true
            PsVersion = "5.1.26100.1"; Ps64 = $true; LanguageMode = "FullLanguage"; UICulture = "ja-JP"
            Office = @(@{ Name = "Excel"; Version = "16.0.17000.1"; Kind = "Click-to-Run"; Bits = 64 })
        }
    }

    It "版があれば、タグとコミット 7 桁を書く" {
        $lines = formatDiagnosticHeader "tebunko" @{ Tag = "v0.5.0"; Sha = ("a" * 40) } $script:env $script:now
        $text = $lines -join "`n"
        $text | Should -Match "tebunko の版: v0.5.0 \(aaaaaaa\)"
        $text | Should -Match "日時: 2026-10-11 09:30:00 \+09:00"
        $text | Should -Match "OS: Windows 10.0.26200（ntoskrnl 10.0.26100.1234・64 ビット）"
        $text | Should -Match "PowerShell: 5.1.26100.1（64 ビット・FullLanguage）"
        $text | Should -Match "Office: Excel 16.0.17000.1 Click-to-Run 64 ビット"
        $text | Should -Match "表示言語: ja-JP"
    }

    It "版が無ければ開発版・Office が無ければ見つからない" {
        $info = $script:env.Clone()
        $info.Office = @()
        $text = (formatDiagnosticHeader "tebunko" $null $info $script:now) -join "`n"
        $text | Should -Match "(?m)tebunko の版: 開発版\r?$"
        $text | Should -Match "Office: 見つからない"
    }

    It "調べられない項目は 不明 と書く" {
        $info = @{ Office = @() }
        $text = (formatDiagnosticHeader "tebunko" $null $info $script:now) -join "`n"
        $text | Should -Match "OS: Windows 不明（ntoskrnl 不明・不明）"
        $text | Should -Match "PowerShell: 不明（不明・不明）"
        $text | Should -Match "表示言語: 不明"
    }
}
