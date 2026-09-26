# コミットメッセージ・タイトルの形を確かめるスクリプト（tools\check_commit_message.ps1）のテスト
BeforeAll {
    $check = "$PSScriptRoot\..\..\tools\check_commit_message.ps1"

    # 終了コードを返す（Write-Host の出力は捨てる）
    function checkTitle([string]$title) {
        & $check -Title $title 6>$null
        return $LASTEXITCODE
    }

    function checkFile([string]$text) {
        $path = Join-Path $TestDrive "COMMIT_EDITMSG"
        [System.IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding($false)))
        & $check -Path $path 6>$null
        return $LASTEXITCODE
    }
}

Describe "check_commit_message.ps1 のタイトルの判定" -Tag Unit {
    It "形が合うものを通す: <title>" -TestCases @(
        @{ title = "feat: Excel の図形の文字を検索できるようにする" }
        @{ title = "fix(gui): 検索結果の件数が 0 のままになるのを直す" }
        @{ title = "feat!: インデックスの形式を変える" }
        @{ title = "refactor(shared/core)!: 設定の読み込みを分ける" }
        @{ title = "ci(deps): bump actions/checkout from 7.0.0 to 7.0.1" }
        @{ title = "docs: 設計書の誤字を直す (#12)" }
    ) {
        param($title)
        checkTitle $title | Should -Be 0
    }

    It "git が自動で作るメッセージは調べない: <title>" -TestCases @(
        @{ title = "Merge remote-tracking branch 'origin/main' into worktree-a" }
        @{ title = 'Revert "feat: Excel の図形の文字を検索できるようにする"' }
        @{ title = "fixup! feat: Excel の図形の文字を検索できるようにする" }
        @{ title = "squash! fix: 直す" }
    ) {
        param($title)
        checkTitle $title | Should -Be 0
    }

    It "形が違うものを止める: <title>" -TestCases @(
        @{ title = "" }
        @{ title = "Excel の図形の文字を検索できるようにする" }
        @{ title = "feature: 型の名前が違う" }
        @{ title = "Feat: 型が大文字" }
        @{ title = "feat:コロンの後ろに空白が無い" }
        @{ title = "feat：全角のコロン" }
        @{ title = "feat:  空白が 2 つ" }
        @{ title = "feat: " }
        @{ title = "feat (gui): 型と範囲の間に空白" }
        @{ title = "feat(): 範囲が空" }
        @{ title = "feat(g ui): 範囲に空白" }
        @{ title = "Revert: 自動で作るものに似せただけ" }
    ) {
        param($title)
        checkTitle $title | Should -Be 1
    }
}

Describe "check_commit_message.ps1 のコミットメッセージのファイル" -Tag Io {
    It "コメント行と空行を飛ばして、最初の行だけを調べる" {
        checkFile "`n# コメント`nfeat: 足す`n`n本文は調べない`nCo-Authored-By: test <test@example.com>`n" | Should -Be 0
        checkFile "# コメント`n足す`n" | Should -Be 1
    }

    It "CRLF のファイルも読める" {
        checkFile "fix: 直す`r`n`r`n本文`r`n" | Should -Be 0
    }

    It "空のメッセージは通す（git がコミットを中止する）" {
        checkFile "# コメントだけ`n`n" | Should -Be 0
    }
}
