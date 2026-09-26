# 「tebunko について」の表示（tebunko\ui\about_view.ps1）のテスト
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
