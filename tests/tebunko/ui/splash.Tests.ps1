# 起動中の表示（ui\splash.ps1）のテスト。本物の窓を作るので Gui（STA のセッションで流す）
BeforeAll {
    $rootDir = (Resolve-Path "$PSScriptRoot\..\..\..").Path
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    . "$rootDir\scripts\tebunko\ui\splash.ps1"
    $xamlPath = "$rootDir\scripts\tebunko\xaml\splash.xaml"
}

Describe "showSplash" -Tag Gui {
    AfterEach {
        closeSplash
    }

    It "xaml を読んで窓を出し、見える状態で返す（失敗を握りつぶして `$null を返すと、起動画面が黙って出なくなる）" {
        $window = showSplash $xamlPath
        $window | Should -Not -BeNullOrEmpty
        $window.IsVisible | Should -BeTrue
    }

    It "stepSplash で進み具合が変わり、closeSplash で閉じる" {
        $window = showSplash $xamlPath
        stepSplash 40
        $window.FindName("SplashProgress").Value | Should -Be 40
        closeSplash
        $window.IsVisible | Should -BeFalse
    }

    It "単一 .ps1 版のように、xaml が文字列で埋め込まれていても窓を出せる" {
        $fullPath = [System.IO.Path]::GetFullPath($xamlPath)
        $text = [System.IO.File]::ReadAllText($fullPath, (New-Object System.Text.UTF8Encoding($true)))
        $bundledXaml = @{ $fullPath = $text }
        $window = showSplash "$rootDir\scripts\tebunko\xaml\..\xaml\splash.xaml"
        $window | Should -Not -BeNullOrEmpty
        $window.IsVisible | Should -BeTrue
    }

    It "xaml が読めなければ `$null を返し、stepSplash・closeSplash は何もしない" {
        showSplash (Join-Path $TestDrive "none.xaml") | Should -BeNullOrEmpty
        { stepSplash 10; closeSplash } | Should -Not -Throw
    }
}
