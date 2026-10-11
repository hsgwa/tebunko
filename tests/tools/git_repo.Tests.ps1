# テストで一時の git リポジトリを作る道具（tests\helpers\git_repo.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\helpers\git_repo.ps1"
}

Describe "git_repo.ps1 の一時のリポジトリの守り" -Tag Io {
    It "git の環境変数が残っていると、リポジトリを作らずに throw する" {
        $saved = clearGitEnv
        try {
            [Environment]::SetEnvironmentVariable("GIT_DIR", (Join-Path $TestDrive "gitdir"))
            $repo = Join-Path $TestDrive "repo"
            { initTestRepo $repo $TestDrive } | Should -Throw "*GIT_DIR*"
            Test-Path $repo | Should -BeFalse
        } finally {
            restoreGitEnv $saved
        }
    }

    It "場所が root の外なら、リポジトリを作らずに throw する" {
        $saved = clearGitEnv
        try {
            $repo = Join-Path $TestDrive "repo"
            { initTestRepo $repo (Join-Path $TestDrive "other") } | Should -Throw "*の外*"
            Test-Path $repo | Should -BeFalse
        } finally {
            restoreGitEnv $saved
        }
    }

    It "環境変数を外していれば root の下にリポジトリを作る" {
        $saved = clearGitEnv
        try {
            $repo = Join-Path $TestDrive "repo"
            initTestRepo $repo $TestDrive
            Test-Path (Join-Path $repo ".git") | Should -BeTrue
        } finally {
            restoreGitEnv $saved
        }
    }

    It "clearGitEnv で外した値は restoreGitEnv で戻る" {
        $saved = clearGitEnv
        try {
            [Environment]::SetEnvironmentVariable("GIT_INDEX_FILE", "index-a")
            [Environment]::SetEnvironmentVariable("GIT_AUTHOR_EMAIL", "a@example.com")
            $inner = clearGitEnv
            [Environment]::GetEnvironmentVariable("GIT_INDEX_FILE") | Should -BeNullOrEmpty
            [Environment]::GetEnvironmentVariable("GIT_AUTHOR_EMAIL") | Should -BeNullOrEmpty
            restoreGitEnv $inner
            [Environment]::GetEnvironmentVariable("GIT_INDEX_FILE") | Should -Be "index-a"
            [Environment]::GetEnvironmentVariable("GIT_AUTHOR_EMAIL") | Should -Be "a@example.com"
        } finally {
            [Environment]::SetEnvironmentVariable("GIT_INDEX_FILE", $null)
            [Environment]::SetEnvironmentVariable("GIT_AUTHOR_EMAIL", $null)
            restoreGitEnv $saved
        }
    }
}
