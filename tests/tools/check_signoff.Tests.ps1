# Signed-off-by を確かめるスクリプト（tools\check_signoff.ps1）のテスト
BeforeAll {
    $check = "$PSScriptRoot\..\..\tools\check_signoff.ps1"
    $mail = "test@example.com"

    # 終了コードを返す（Write-Host の出力は捨てる）
    function checkFile([string]$text, [string]$author = $mail) {
        $path = Join-Path $TestDrive "COMMIT_EDITMSG"
        [System.IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding($false)))
        & $check -Path $path -AuthorEmail $author 6>$null
        return $LASTEXITCODE
    }

    # 画面に出した文字列を返す
    function checkFileOutput([string]$text, [string]$author = $mail) {
        $path = Join-Path $TestDrive "COMMIT_EDITMSG"
        [System.IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding($false)))
        return (& $check -Path $path -AuthorEmail $author 6>&1 | Out-String)
    }
}

Describe "check_signoff.ps1 のコミットメッセージのファイル" -Tag Io {
    It "作者の Signed-off-by があれば通す" {
        checkFile "feat: 足す`n`nSigned-off-by: test <test@example.com>`n" | Should -Be 0
        checkFile "feat: 足す`r`n`r`nCo-Authored-By: a <a@example.com>`r`nSigned-off-by: test <Test@Example.com>`r`n" | Should -Be 0
    }

    It "無いとき・作者と違うメールアドレスのときは止める" {
        checkFile "feat: 足す`n`n本文`n" | Should -Be 1
        checkFile "feat: 足す`n`nSigned-off-by: other <other@example.com>`n" | Should -Be 1
        checkFile "feat: 足す`n`nSigned-off-by: test`n" | Should -Be 1
    }

    It "作者の行に加えて作者と違うメールアドレスの行があれば止める（<Name>）" -TestCases @(
        @{ Name = "作者の行のあとに違う行"; Text = "feat: 足す`n`nSigned-off-by: test <test@example.com>`nSigned-off-by: other <other@example.org>`n" }
        @{ Name = "違う行のあとに作者の行"; Text = "feat: 足す`n`nSigned-off-by: other <other@example.org>`nSigned-off-by: test <test@example.com>`n" }
        @{ Name = "違う行だけ"; Text = "feat: 足す`n`nSigned-off-by: other <other@example.org>`n" }
        @{ Name = "別の noreply の行"; Text = "feat: 足す`n`nSigned-off-by: test <test@example.com>`nSigned-off-by: a <1+a@users.noreply.github.com>`n" }
    ) {
        checkFile $Text | Should -Be 1
    }

    It "大文字・小文字だけ違う作者の行は通し、切り取り線より後ろの違う行は数えない" {
        checkFile "feat: 足す`n`nSigned-off-by: test <TEST@example.com>`nSigned-off-by: Test <test@Example.com>`n" | Should -Be 0
        checkFile "feat: 足す`n`nSigned-off-by: test <test@example.com>`n# ------------------------ >8 ------------------------`n+Signed-off-by: other <other@example.org>`n" | Should -Be 0
    }

    It "bot の作者・マージコミットは、違う行があっても調べない" {
        checkFile "ci(deps): bump a`n`nSigned-off-by: other <other@example.org>`n" "49699333+dependabot[bot]@users.noreply.github.com" | Should -Be 0
        checkFile "Merge branch 'a'`n`nSigned-off-by: other <other@example.org>`n" | Should -Be 0
    }

    It "止めたときの出力に、違う行のメールアドレスを含めず、直し方を出す" {
        $out = checkFileOutput "feat: 足す`n`nSigned-off-by: test <test@example.com>`nSigned-off-by: other <other@example.org>`n"
        $out | Should -Not -Match "other@example\.org"
        $out | Should -Match "git commit --amend -s"
        $out | Should -Match "git rebase origin/main -x"
    }

    It "コメント行と切り取り線より後ろ（git commit -v の差分）の Signed-off-by は数えない" {
        checkFile "feat: 足す`n# Signed-off-by: test <test@example.com>`n" | Should -Be 1
        checkFile "feat: 足す`n# ------------------------ >8 ------------------------`n+Signed-off-by: test <test@example.com>`nSigned-off-by: test <test@example.com>`n" | Should -Be 1
    }

    It "マージコミット・bot の作者・空のメッセージは調べない" {
        checkFile "Merge remote-tracking branch 'origin/main' into worktree-a`n" | Should -Be 0
        checkFile "ci(deps): bump a`n" "49699333+dependabot[bot]@users.noreply.github.com" | Should -Be 0
        checkFile "# コメントだけ`n`n" | Should -Be 0
    }
}

Describe "check_signoff.ps1 の範囲のコミット" -Tag Io {
    BeforeAll {
        # コミット前のフックから動くときは、GIT_DIR などがフックを動かしたリポジトリを、GIT_AUTHOR_* などがそのコミットの作者を指している。外して、下で作るリポジトリを使う
        $script:savedGitEnv = @{}
        foreach ($name in @("GIT_DIR", "GIT_WORK_TREE", "GIT_INDEX_FILE", "GIT_PREFIX", "GIT_OBJECT_DIRECTORY", "GIT_COMMON_DIR", "GIT_AUTHOR_NAME", "GIT_AUTHOR_EMAIL", "GIT_AUTHOR_DATE", "GIT_COMMITTER_NAME", "GIT_COMMITTER_EMAIL", "GIT_COMMITTER_DATE")) {
            $script:savedGitEnv[$name] = [Environment]::GetEnvironmentVariable($name)
            [Environment]::SetEnvironmentVariable($name, $null)
        }

        $repo = Join-Path $TestDrive "repo"

        # 作者を決めてコミットする（-m は段落ごと）
        function commit([string]$author, [string[]]$paragraphs) {
            $arguments = @("-C", $repo, "-c", "user.name=test", "-c", "user.email=$author", "commit", "-q", "--allow-empty", "--no-verify")
            foreach ($p in $paragraphs) { $arguments += @("-m", $p) }
            & git @arguments
        }

        function checkRangeOutput([string]$range) {
            Push-Location $repo
            try {
                return (& $check -Range $range 6>&1 | Out-String)
            } finally {
                Pop-Location
            }
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
        checkRange "base..HEAD" | Should -Be 0
    }

    It "1 つでも無ければ止める" {
        commit $mail @("feat: 一つ目", "Signed-off-by: test <test@example.com>")
        commit $mail @("fix: 二つ目")
        checkRange "base..HEAD" | Should -Be 1
    }

    It "作者の行に加えて作者と違う行があるコミットが 1 つでもあれば止め、出力にそのメールアドレスを含めない" {
        commit $mail @("feat: 一つ目", "Signed-off-by: test <test@example.com>")
        commit $mail @("fix: 二つ目", "Signed-off-by: test <test@example.com>`nSigned-off-by: other <other@example.org>")
        checkRange "base..HEAD" | Should -Be 1
        $out = checkRangeOutput "base..HEAD"
        $out | Should -Not -Match "other@example\.org"
        $out | Should -Match "fix: 二つ目"
        $out | Should -Match "git rebase origin/main -x"
    }

    It "違う行だけのコミットも止める" {
        commit $mail @("feat: 一つ目", "Signed-off-by: other <other@example.org>")
        checkRange "base..HEAD" | Should -Be 1
    }

    It "bot のコミットは、作者と違う行があっても調べない" {
        commit "49699333+dependabot[bot]@users.noreply.github.com" @("ci(deps): bump a", "Signed-off-by: other <other@example.org>")
        checkRange "base..HEAD" | Should -Be 0
    }

    It "マージコミットと bot のコミットは調べない" {
        git -C $repo checkout -q -b side
        commit "49699333+dependabot[bot]@users.noreply.github.com" @("ci(deps): bump a")
        git -C $repo checkout -q -
        commit $mail @("feat: 足す", "Signed-off-by: test <test@example.com>")
        git -C $repo -c user.name=test -c user.email=$mail merge -q --no-ff --no-verify --no-edit side
        checkRange "base..HEAD" | Should -Be 0
    }

    AfterAll {
        foreach ($name in $script:savedGitEnv.Keys) {
            [Environment]::SetEnvironmentVariable($name, $script:savedGitEnv[$name])
        }
    }
}
