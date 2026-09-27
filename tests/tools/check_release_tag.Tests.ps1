# リリースのタグの検査（tools\check_release_tag.ps1）のテスト
BeforeAll {
    $check = "$PSScriptRoot\..\..\tools\check_release_tag.ps1"

    # 終了コードを返す（Write-Host の出力は捨てる）
    function checkTag([string]$tag, [string]$sha, [string]$base = "main") {
        & $check -Tag $tag -Sha $sha -Base $base 6>$null
        return $LASTEXITCODE
    }

    # コミット前のフックから動くときは、GIT_DIR などがフックを動かしたリポジトリを指している。外して、下で作るリポジトリを使う
    $script:savedGitEnv = @{}
    foreach ($name in @("GIT_DIR", "GIT_WORK_TREE", "GIT_INDEX_FILE", "GIT_PREFIX", "GIT_OBJECT_DIRECTORY", "GIT_COMMON_DIR")) {
        $script:savedGitEnv[$name] = [Environment]::GetEnvironmentVariable($name)
        [Environment]::SetEnvironmentVariable($name, $null)
    }

    # main に 1 コミット、別のブランチに main に入っていないコミットがあるリポジトリを作る
    $repo = Join-Path $TestDrive "repo"
    New-Item -ItemType Directory -Path $repo | Out-Null
    Push-Location $repo
    git init -q -b main
    git config user.email "test@example.com"
    git config user.name "test"
    git config commit.gpgsign false
    "a" | Set-Content a.txt
    git add a.txt
    git commit -q -m "a"
    $script:onMain = (git rev-parse HEAD).Trim()
    git checkout -q -b side
    "b" | Set-Content b.txt
    git add b.txt
    git commit -q -m "b"
    $script:offMain = (git rev-parse HEAD).Trim()
    git checkout -q main
}

AfterAll {
    Pop-Location
    foreach ($name in $script:savedGitEnv.Keys) { [Environment]::SetEnvironmentVariable($name, $script:savedGitEnv[$name]) }
    Remove-Item $repo -Recurse -Force -ErrorAction SilentlyContinue
}

Describe "check_release_tag.ps1 のタグの形" -Tag Io {
    It "<tag> は通す" -TestCases @(
        @{ tag = "v0.2.0" }
        @{ tag = "v10.20.30" }
        @{ tag = "v0.0.0" }
    ) {
        checkTag $tag $script:onMain | Should -Be 0
    }

    It "<tag> は形が違うので止める" -TestCases @(
        @{ tag = "v1.0" }
        @{ tag = "v1.0.0-rc1" }
        @{ tag = "1.0.0" }
        @{ tag = "V1.0.0" }
        @{ tag = "v1.0.0.0" }
        @{ tag = "v01.0.0" }
        @{ tag = "v1.00.0" }
        @{ tag = "v1.0.0`n" }
        @{ tag = ([string][char]0xFF11 + ".0.0").Insert(0, "v") }
        @{ tag = "" }
    ) {
        checkTag $tag $script:onMain | Should -Be 1
    }
}

Describe "check_release_tag.ps1 のコミット" -Tag Io {
    It "main にあるコミットは通す" {
        checkTag "v0.2.0" $script:onMain | Should -Be 0
    }

    # 止めるときは終了コード 1 と、理由の文言を出す。main に無いときと、参照を解決できないときで文言が違う
    It "<name>は止め、「<reason>」と出す" -TestCases @(
        @{ name = "main に無いコミット"; sha = "off"; base = "main"; reason = "履歴にありません" }
        @{ name = "存在しないコミット"; sha = ("0" * 40); base = "main"; reason = "解決できませんでした" }
        @{ name = "存在しない基準"; sha = "on"; base = "origin/nothing"; reason = "解決できませんでした" }
    ) {
        if ($sha -eq "on") { $sha = $script:onMain }
        if ($sha -eq "off") { $sha = $script:offMain }
        $out = & $check -Tag "v0.2.0" -Sha $sha -Base $base 6>&1 | Out-String
        $LASTEXITCODE | Should -Be 1
        $out | Should -Match $reason
    }
}
