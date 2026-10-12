# 「バージョン情報」の表示（tebunko\ui\about_view.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\about_view.ps1"
}

Describe "getAboutView" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "版があれば、その版と先頭7桁のコミットを出す"; info = @{ Tag = "v1.2.3"; Sha = "0123456789abcdef0123456789abcdef01234567" }; version = "v1.2.3"; commit = "0123456" }
        @{ name = '$null（VERSION.txt が無い・形が違う）なら開発版とし、コミットは出さない'; info = $null; version = "開発版"; commit = "" }
    ) {
        $view = getAboutView $info
        $view.Version | Should -Be $version
        $view.Commit | Should -Be $commit
    }
}

Describe "getAboutLinks・testAboutLinkUrl" -Tag Unit {
    It "リンクは GitHub のリポジトリとリリースの 2 つで、どれも https の GitHub のアドレス" {
        $links = @(getAboutLinks)
        ($links | ForEach-Object { $_.Id }) -join "," | Should -Be "GitHubLink,ReleasesLink"
        ($links | ForEach-Object { $_.Url }) -join "," | Should -Be "https://github.com/hsgwa/tebunko,https://github.com/hsgwa/tebunko/releases"
        foreach ($link in $links) { $link.Text | Should -Not -BeNullOrEmpty }
    }

    It "<name>" -TestCases @(
        @{ name = "載せたアドレスは開いてよい"; url = "https://github.com/hsgwa/tebunko/releases"; expected = $true }
        @{ name = "載せていないアドレスは開かない"; url = "https://example.com/"; expected = $false }
        @{ name = "大文字小文字が違うものは開かない"; url = "https://GitHub.com/hsgwa/tebunko"; expected = $false }
        @{ name = "後ろに足したものは開かない"; url = "https://github.com/hsgwa/tebunko/../x"; expected = $false }
        @{ name = "ファイルのパスや実行ファイルは開かない"; url = "C:\Windows\System32\cmd.exe"; expected = $false }
        @{ name = "空は開かない"; url = ""; expected = $false }
        @{ name = "末尾に / を足したものは開かない"; url = "https://github.com/hsgwa/tebunko/"; expected = $false }
        @{ name = "前後に空白があるものは開かない"; url = " https://github.com/hsgwa/tebunko "; expected = $false }
        @{ name = "クエリを付けたものは開かない"; url = "https://github.com/hsgwa/tebunko?x=1"; expected = $false }
    ) {
        testAboutLinkUrl $url | Should -Be $expected
    }
}

Describe "getAboutLicenses" -Tag Unit {
    It "tebunko の MIT と、同梱するアイコンの形（Lucide）の ISC（一部 MIT）を出す" {
        $lines = @(getAboutLicenses)
        $lines.Count | Should -Be 2
        $lines[0] | Should -BeLike "tebunko*MIT*"
        $lines[1] | Should -BeLike "*Lucide*ISC*一部 MIT*"
    }
}
