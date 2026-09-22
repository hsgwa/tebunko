# Signed-off-by を確かめるスクリプト（tools\check_signoff.ps1）のテスト
$check = "$PSScriptRoot\..\..\tools\check_signoff.ps1"
$mail = "test@example.com"

# 終了コードを返す（Write-Host の出力は捨てる）
function checkFile([string]$text, [string]$author = $mail) {
    $path = Join-Path $TestDrive "COMMIT_EDITMSG"
    [System.IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding($false)))
    & $check -Path $path -AuthorEmail $author 6>$null
    return $LASTEXITCODE
}

Describe "check_signoff.ps1 のコミットメッセージのファイル" -Tag Io {
    It "作者の Signed-off-by があれば通す" {
        checkFile "feat: 足す`n`nSigned-off-by: test <test@example.com>`n" | Should Be 0
        checkFile "feat: 足す`r`n`r`nCo-Authored-By: a <a@example.com>`r`nSigned-off-by: test <Test@Example.com>`r`n" | Should Be 0
    }

    It "無いとき・作者と違うメールアドレスのときは止める" {
        checkFile "feat: 足す`n`n本文`n" | Should Be 1
        checkFile "feat: 足す`n`nSigned-off-by: other <other@example.com>`n" | Should Be 1
        checkFile "feat: 足す`n`nSigned-off-by: test`n" | Should Be 1
    }

    It "コメント行と切り取り線より後ろ（git commit -v の差分）の Signed-off-by は数えない" {
        checkFile "feat: 足す`n# Signed-off-by: test <test@example.com>`n" | Should Be 1
        checkFile "feat: 足す`n# ------------------------ >8 ------------------------`n+Signed-off-by: test <test@example.com>`nSigned-off-by: test <test@example.com>`n" | Should Be 1
    }

    It "マージコミット・bot の作者・空のメッセージは調べない" {
        checkFile "Merge remote-tracking branch 'origin/main' into worktree-a`n" | Should Be 0
        checkFile "ci(deps): bump a`n" "49699333+dependabot[bot]@users.noreply.github.com" | Should Be 0
        checkFile "# コメントだけ`n`n" | Should Be 0
    }
}

Describe "check_signoff.ps1 の範囲のコミット" -Tag Io {
    $repo = Join-Path $TestDrive "repo"

    # 作者を決めてコミットする（-m は段落ごと）
    function commit([string]$author, [string[]]$paragraphs) {
        $arguments = @("-C", $repo, "-c", "user.name=test", "-c", "user.email=$author", "commit", "-q", "--allow-empty", "--no-verify")
        foreach ($p in $paragraphs) { $arguments += @("-m", $p) }
        & git @arguments
    }

    function checkRange([string]$range) {
        Push-Location $repo
        try {
            & $check -Range $range 6>$null
            return $LASTEXITCODE
        } finally {
            Pop-Location
        }
    }

    BeforeEach {
        Remove-Item $repo -Recurse -Force -ErrorAction SilentlyContinue
        git init -q $repo
        commit $mail @("chore: 最初")
        git -C $repo tag base
    }

    It "範囲のコミットすべてに作者の Signed-off-by があれば通す" {
        commit $mail @("feat: 一つ目", "Signed-off-by: test <test@example.com>")
        commit $mail @("fix: 二つ目の説明は日本語", "Signed-off-by: test <test@example.com>")
        checkRange "base..HEAD" | Should Be 0
    }

    It "1 つでも無ければ止める" {
        commit $mail @("feat: 一つ目", "Signed-off-by: test <test@example.com>")
        commit $mail @("fix: 二つ目")
        checkRange "base..HEAD" | Should Be 1
    }

    It "範囲より前のコミットは調べない" {
        checkRange "base..HEAD" | Should Be 0
    }

    It "マージコミットと bot のコミットは調べない" {
        git -C $repo checkout -q -b side
        commit "49699333+dependabot[bot]@users.noreply.github.com" @("ci(deps): bump a")
        git -C $repo checkout -q -
        commit $mail @("feat: 足す", "Signed-off-by: test <test@example.com>")
        git -C $repo -c user.name=test -c user.email=$mail merge -q --no-ff --no-verify --no-edit side
        checkRange "base..HEAD" | Should Be 0
    }
}
