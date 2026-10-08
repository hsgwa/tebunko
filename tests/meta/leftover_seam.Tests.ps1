# 残り Office の確認の「偽の行」の継ぎ目（写真・画面のテスト用）が、本物の動きに混ざらないことを確かめるテスト。
# 継ぎ目は環境変数 TEBUNKO_GUI_LEFTOVER_FILE で、画面層の leftover_dialog.ps1 だけが読む。既定（未設定）は本物。
BeforeAll {
    $here = (Resolve-Path "$PSScriptRoot\..").Path
    $scriptsDir = (Resolve-Path "$here\..\scripts").Path
    $script:dialogPath = "$scriptsDir\tebunko\ui\leftover_dialog.ps1"
}

Describe "残り Office の偽の行は、本物の動きに混ざらない" -Tag Meta {
    It "環境変数 TEBUNKO_GUI_LEFTOVER_FILE を読むのは（コメントを除き）、leftover_dialog.ps1 だけ" {
        $users = @(Get-ChildItem -LiteralPath $scriptsDir -Recurse -Filter "*.ps1" |
            Where-Object { @(Get-Content -LiteralPath $_.FullName -Encoding UTF8 | Where-Object { $_ -notmatch '^\s*#' -and $_ -match "TEBUNKO_GUI_LEFTOVER_FILE" }).Count -gt 0 } |
            ForEach-Object { $_.FullName })
        $users | Should -Be @($script:dialogPath)
    }

    It "画面層の継ぎ目は、未設定なら本物（Real）になる向きで書かれている" {
        $text = Get-Content -LiteralPath $script:dialogPath -Raw -Encoding UTF8
        $text | Should -Match 'getLeftoverSourceMode'
        $text | Should -Match '\$script:leftoverFakeFile\s*=\s*\$env:TEBUNKO_GUI_LEFTOVER_FILE'
    }

    It "継ぎ目のあるファイルには、プロセスを止める呼び出しが無い（止めるのは stopOfficeProcesses だけ）" {
        $text = Get-Content -LiteralPath $script:dialogPath -Raw -Encoding UTF8
        $text | Should -Not -Match 'Stop-Process|\.Kill\('
    }

    It "偽の行の関数は、leftover_dialog.ps1 と leftover_view.ps1 のほかから呼ばれない" {
        $text = Get-Content -LiteralPath $script:dialogPath -Raw -Encoding UTF8
        foreach ($name in "convertLeftoverFakeRows", "getLeftoverFakeStopResults") {
            $others = @(Get-ChildItem -LiteralPath $scriptsDir -Recurse -Filter "*.ps1" |
                Where-Object { $_.Name -notin @("leftover_dialog.ps1", "leftover_view.ps1") -and (Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8) -match $name })
            $others.Count | Should -Be 0
            $text | Should -Match $name
        }
    }
}
