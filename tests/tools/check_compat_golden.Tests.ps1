# 見本（tests/testdata/compat/）の変更に ! を求めるスクリプト（tools/check_compat_golden.ps1）のテスト
BeforeAll {
    $check = "$PSScriptRoot\..\..\tools\check_compat_golden.ps1"

    # 終了コードを返す（Write-Host の出力は捨てる）。git は呼ばず、-DiffLines に直接 name-status の行を渡す
    function checkDiff([string]$title, [string[]]$diffLines) {
        & $check -Title $title -Base "base" -DiffLines $diffLines 6>$null
        return $LASTEXITCODE
    }
}

Describe "check_compat_golden.ps1" -Tag Unit {
    It "通す: <label>" -TestCases @(
        @{ label = "compat/ に A だけ・! なし"; title = "test: 見本を足す"; lines = @("A`ttests/testdata/compat/index/v1/a.txt") }
        @{ label = "M で ! あり"; title = "test!: 見本を直す"; lines = @("M`ttests/testdata/compat/index/v1/a.txt") }
        @{ label = "D で ! あり"; title = "feat(index)!: 形式を変える"; lines = @("D`ttests/testdata/compat/index/v1/a.txt") }
        @{ label = "R で ! あり"; title = "test!: 見本を改名する"; lines = @("R100`ttests/testdata/compat/index/v1/a.txt`ttests/testdata/compat/index/v1/b.txt") }
        @{ label = "日本語を含むパスの A（core.quotepath=off の出力の形）"; title = "test: 見本を足す"; lines = @("A`ttests/testdata/compat/index/v1/資料・案内[確定].txt") }
        @{ label = "compat/ の外だけの M"; title = "docs: 設計書を直す"; lines = @("M`tdocs/design/index-data/format.md") }
        @{ label = "差分が無い"; title = "test: 見本を足す"; lines = @() }
    ) {
        param($title, $lines)
        checkDiff $title $lines | Should -Be 0
    }

    It "止める: <label>" -TestCases @(
        @{ label = "M で ! なし"; title = "test: 見本を直す"; lines = @("M`ttests/testdata/compat/index/v1/a.txt") }
        @{ label = "D で ! なし"; title = "fix: 直す"; lines = @("D`ttests/testdata/compat/index/v1/a.txt") }
        @{ label = "R で ! なし"; title = "test: 見本を改名する"; lines = @("R100`ttests/testdata/compat/index/v1/a.txt`ttests/testdata/compat/index/v1/b.txt") }
        @{ label = "日本語を含むパスの M（core.quotepath=off の出力の形）"; title = "test: 見本を直す"; lines = @("M`ttests/testdata/compat/index/v1/資料・案内[確定].txt") }
        @{ label = "A と M が混ざる（1 つでも A 以外があれば ! が要る）"; title = "test: 見本を足す・直す"; lines = @("A`ttests/testdata/compat/index/v1/new.txt", "M`ttests/testdata/compat/index/v1/a.txt") }
    ) {
        param($title, $lines)
        checkDiff $title $lines | Should -Be 1
    }
}
