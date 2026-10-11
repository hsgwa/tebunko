# テストで一時の git リポジトリを作る道具（tests\helpers\git_repo.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\helpers\git_repo.ps1"
}

Describe "git_repo.ps1 の一時のリポジトリの守り" -Tag Io {
    It "git の環境変数が残っていると、リポジトリを作らずに throw する" {
        $saved = clearGitEnv
        try {
            [Environment]::SetEnvironmentVariable("GIT_DIR", (Join-Path $TestDrive "gitdir"))
            $repo = Join-Path $TestDrive "envleft"
            { initTestRepo $repo $TestDrive } | Should -Throw "*GIT_DIR*"
            Test-Path $repo | Should -BeFalse
        } finally {
            restoreGitEnv $saved
        }
    }

    It "場所が root の外（<name>）なら、リポジトリを作らずに throw する" {
        param($name, $rootRel, $pathRel)
        $root = Join-Path $TestDrive $rootRel
        $path = Join-Path $TestDrive $pathRel
        $saved = clearGitEnv
        try {
            { initTestRepo $path $root } | Should -Throw "*の外*"
            Test-Path $path | Should -BeFalse
        } finally {
            restoreGitEnv $saved
        }
    } -TestCases @(
        @{ name = "別のフォルダ"; rootRel = "other"; pathRel = "outside1" }
        @{ name = "名前が root で始まる隣のフォルダ"; rootRel = "rep"; pathRel = "repo" }
        @{ name = ".. で root の外へ出る"; rootRel = "a\sub"; pathRel = "a\sub\..\x" }
    )

    It "環境変数を外していれば root の下にリポジトリを作る" {
        $saved = clearGitEnv
        try {
            $repo = Join-Path $TestDrive "created"
            initTestRepo $repo $TestDrive
            Test-Path (Join-Path $repo ".git") | Should -BeTrue
        } finally {
            restoreGitEnv $saved
        }
    }

    It "clearGitEnv で外した値は restoreGitEnv で戻る" {
        $saved = clearGitEnv
        $indexA = Join-Path $TestDrive "index-a"
        try {
            [Environment]::SetEnvironmentVariable("GIT_INDEX_FILE", $indexA)
            [Environment]::SetEnvironmentVariable("GIT_AUTHOR_EMAIL", "a@example.com")
            $inner = clearGitEnv
            [Environment]::GetEnvironmentVariable("GIT_INDEX_FILE") | Should -BeNullOrEmpty
            [Environment]::GetEnvironmentVariable("GIT_AUTHOR_EMAIL") | Should -BeNullOrEmpty
            restoreGitEnv $inner
            [Environment]::GetEnvironmentVariable("GIT_INDEX_FILE") | Should -Be $indexA
            [Environment]::GetEnvironmentVariable("GIT_AUTHOR_EMAIL") | Should -Be "a@example.com"
        } finally {
            [Environment]::SetEnvironmentVariable("GIT_INDEX_FILE", $null)
            [Environment]::SetEnvironmentVariable("GIT_AUTHOR_EMAIL", $null)
            restoreGitEnv $saved
        }
    }
}
