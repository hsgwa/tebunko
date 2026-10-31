# tools\capture_screens.ps1 の判断だけの関数（tools\capture\capture_common.ps1）のテスト。
# 画面を動かす部分（本体）は流さない（実機の確かめは PR 本文に書く）。
BeforeAll {
    . "$PSScriptRoot\..\..\tools\capture\capture_common.ps1"
}

Describe "captureIds" -Tag Unit {
    It "46 件あり、重複が無い" {
        ${captureIds}.Count | Should -Be 46
        (${captureIds} | Select-Object -Unique).Count | Should -Be 46
    }

    It "すべて <画面>/<状態> の形" {
        ${captureIds} | ForEach-Object { $_ | Should -Match "^[a-z-]+/[a-z-]+$" }
    }
}

Describe "resolveCaptureIds" -Tag Unit {
    It "指定が無ければ、一覧をそのまま返す" {
        resolveCaptureIds -Only @() -Ids @("window/menu", "index-tab/empty") | Should -Be @("window/menu", "index-tab/empty")
    }

    It "指定した ID だけを、一覧の表記で返す" {
        resolveCaptureIds -Only @("index-tab/empty") -Ids @("window/menu", "index-tab/empty", "index-tab/normal") | Should -Be @("index-tab/empty")
    }

    It "知らない ID を指定すると例外にする" {
        { resolveCaptureIds -Only @("index-tab/no-such") -Ids @("index-tab/empty") } | Should -Throw "*index-tab/no-such*"
    }
}

Describe "getCaptureImagePath" -Tag Unit {
    It "<id> は <path>" -TestCases @(
        @{ id = "index-tab/empty"; path = "C:\out\index-tab\empty.png" }
        @{ id = "window/close-confirm"; path = "C:\out\window\close-confirm.png" }
    ) {
        param ($id, $path)
        getCaptureImagePath -Id $id -OutDir "C:\out" | Should -Be $path
    }

    It "<id> は ID の形が正しくないので例外にする" -TestCases @(
        @{ id = "index-tab" }
        @{ id = "" }
        @{ id = "index-tab/" }
        @{ id = "/empty" }
    ) {
        param ($id)
        { getCaptureImagePath -Id $id -OutDir "C:\out" } | Should -Throw
    }
}

Describe "getCaptureEnvironmentProblems" -Tag Unit {
    It "そろえる条件にすべて合えば空" {
        getCaptureEnvironmentProblems -AppliedDpi 96 -LightTheme $true -ScreenCount 1 -OfficeProcessCount 0 | Should -BeNullOrEmpty
    }

    It "<what> が合わないと、その理由が入る" -TestCases @(
        @{ what = "倍率"; dpi = 120; theme = $true; screens = 1; office = 0; like = "*倍率*" }
        @{ what = "テーマ"; dpi = 96; theme = $false; screens = 1; office = 0; like = "*ライトモード*" }
        @{ what = "画面の数"; dpi = 96; theme = $true; screens = 2; office = 0; like = "*画面が 1 つ*" }
        @{ what = "Office"; dpi = 96; theme = $true; screens = 1; office = 2; like = "*Excel*" }
    ) {
        param ($dpi, $theme, $screens, $office, $like)
        $problems = getCaptureEnvironmentProblems -AppliedDpi $dpi -LightTheme $theme -ScreenCount $screens -OfficeProcessCount $office
        $problems | Where-Object { $_ -like $like } | Should -Not -BeNullOrEmpty
    }

    It "合わない条件が複数あれば、理由も複数になる" {
        (getCaptureEnvironmentProblems -AppliedDpi 120 -LightTheme $false -ScreenCount 2 -OfficeProcessCount 1).Count | Should -Be 4
    }
}

Describe "testCaptureSensitiveText" -Tag Unit {
    It "利用者名・コンピューター名・利用者のフォルダのパスのどれかを含めば真" -TestCases @(
        @{ text = "既定の場所は「C:\Users\a\Documents\tebunko_ws」です。"; contains = $true }
        @{ text = "設定ファイルの場所: C:\Users\a\AppData"; contains = $true }
        @{ text = "コンピューター名: TEST-PC の共有"; contains = $true }
        @{ text = "既定の場所（ドキュメントの tebunko）です。"; contains = $false }
        @{ text = ""; contains = $false }
        @{ text = $null; contains = $false }
    ) {
        param ($text, $contains)
        testCaptureSensitiveText -Text $text -UserName "a" -ComputerName "TEST-PC" -UserProfile "C:\Users\a" | Should -Be $contains
    }
}

Describe "getCaptureRedactedText" -Tag Unit {
    It "利用者のフォルダのパスを C:\Users\test に置き換える" {
        getCaptureRedactedText -Text "既定の場所は「C:\Users\a\Documents\tebunko_ws」です。" -UserProfile "C:\Users\a" |
            Should -Be "既定の場所は「C:\Users\test\Documents\tebunko_ws」です。"
    }

    It "大文字小文字が違っても置き換える" {
        getCaptureRedactedText -Text "C:\USERS\a\Documents" -UserProfile "C:\Users\a" | Should -Be "C:\Users\test\Documents"
    }

    It "利用者のフォルダのパスを含まなければ、そのまま返す" {
        getCaptureRedactedText -Text "既定の場所（ドキュメントの tebunko）です。" -UserProfile "C:\Users\a" |
            Should -Be "既定の場所（ドキュメントの tebunko）です。"
    }

    It "利用者のフォルダのパスの外に出た利用者名だけも置き換える" {
        getCaptureRedactedText -Text "共有: \\山田太郎-PC\共有" -UserName "山田太郎" -ComputerName "山田太郎-PC" |
            Should -Be "共有: \\TEST-PC\共有"
    }

    It "コンピューター名だけを含む部品も置き換える" {
        getCaptureRedactedText -Text "\\yamada-pc\営業部" -ComputerName "yamada-pc" | Should -Be "\\TEST-PC\営業部"
    }

    It "利用者のフォルダのパスと、パスの外の利用者名の両方を置き換える" {
        getCaptureRedactedText -Text "C:\Users\a\Documents（a の共有）" -UserProfile "C:\Users\a" -UserName "a" |
            Should -Be "C:\Users\test\Documents（test の共有）"
    }
}

Describe "testCaptureImageSize" -Tag Unit {
    It "<bytes> バイトは、200 KB 以下なら <ok>" -TestCases @(
        @{ bytes = 100KB; ok = $true }
        @{ bytes = 200KB; ok = $true }
        @{ bytes = 201KB; ok = $false }
    ) {
        param ($bytes, $ok)
        testCaptureImageSize -Bytes $bytes | Should -Be $ok
    }
}

Describe "撮る場面（capture_screens.ps1）" -Tag Unit {
    BeforeAll {
        $script:captureScript = [System.IO.File]::ReadAllText("$PSScriptRoot\..\..\tools\capture_screens.ps1")
    }

    It "場面の関数は、すべて撮影の本体から呼ばれている（足した場面を呼び忘れない）。captureTopOnlyScene も含む" {
        $scenes = @([regex]::Matches($script:captureScript, '(?m)^function (capture\w+Scene) \{') | ForEach-Object { $_.Groups[1].Value })
        $scenes | Should -Contain "captureTopOnlyScene"
        $scenes | Should -Contain "captureTextExtensionsScene"
        foreach ($scene in $scenes) {
            [regex]::Matches($script:captureScript, "(?m)^\s+$scene\s+\`$ids ").Count | Should -Be 1 -Because "$scene を本体から 1 回呼ぶ"
        }
    }

    It "状態の ID <id> は、場面の中で撮っている（一覧に足して撮り忘れない）" -TestCases @(
        @{ id = "index-tab/top-only" }
        @{ id = "index-tab/start-confirm-top-only" }
    ) {
        param ($id)
        $script:captureScript | Should -Match ('captureGuiState -S \$S -Id "' + [regex]::Escape($id) + '"')
    }
}
