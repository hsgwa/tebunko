# 「tebunko について」の表示（tebunko\ui\about_view.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\about_view.ps1"
}

Describe "getAboutView" -Tag Unit {
    It "版があれば、その版と先頭7桁のコミットを出す" {
        $view = getAboutView @{ Tag = "v1.2.3"; Sha = "0123456789abcdef0123456789abcdef01234567" }
        $view.Version | Should -Be "v1.2.3"
        $view.Commit | Should -Be "0123456"
    }

    It '$null（VERSION.txt が無い・形が違う）なら開発版とし、コミットは出さない' {
        $view = getAboutView $null
        $view.Version | Should -Be "開発版"
        $view.Commit | Should -Be ""
    }
}
