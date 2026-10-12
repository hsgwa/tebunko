# 版の記録（shared\core\version.ps1 の readVersionFile）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    function writeVersionFileBytes {
        # BOM 付き UTF-8 で $lines（CRLF 区切り）を書く。$bom を $false にすると BOM なしで書ける
        param (
            [string]$path,
            [string[]]$lines,
            [bool]$bom = $true,
            [string]$newLine = "`r`n"
        )

        $content = ($lines -join $newLine)
        if ($lines.Count -gt 0) {
            $content += $newLine
        }
        $encoding = New-Object System.Text.UTF8Encoding($bom)
        [System.IO.File]::WriteAllBytes($path, $encoding.GetBytes($content))
    }
}

Describe "readVersionFile" -Tag Io {
    It "ファイルが無ければ `$null" {
        readVersionFile (Join-Path $TestDrive "無い\VERSION.txt") | Should -BeNullOrEmpty
    }

    It "<name>なら `$null" -TestCases @(
        @{ name = "空のファイル"; lines = @() }
        @{ name = "1 行だけ"; lines = @("v1.0.0") }
        @{ name = "3 行ある"; lines = @("v1.0.0", "0123456789abcdef0123456789abcdef01234567", "") }
        @{ name = "1 行目（タグ名）の形が違う"; lines = @("v1.0.0 だめ", "0123456789abcdef0123456789abcdef01234567") }
        @{ name = "2 行目（SHA）の形が違う"; lines = @("v1.0.0", "abc") }
    ) {
        $path = Join-Path $TestDrive "version.txt"
        writeVersionFileBytes $path $lines
        readVersionFile $path | Should -BeNullOrEmpty
    }

    It "<name>でも、2 行の形が正しければ Tag・Sha を返す" -TestCases @(
        @{ name = "BOM 付き UTF-8・CRLF"; bom = $true; newLine = "`r`n" }
        @{ name = "BOM が無く LF 区切り"; bom = $false; newLine = "`n" }
    ) {
        $path = Join-Path $TestDrive "ok.txt"
        writeVersionFileBytes $path @("v1.0.0", "0123456789abcdef0123456789abcdef01234567") $bom $newLine
        $result = readVersionFile $path
        $result.Tag | Should -Be "v1.0.0"
        $result.Sha | Should -Be "0123456789abcdef0123456789abcdef01234567"
    }

    It "`${bundledVersion} があれば、ファイルを見ずにその値を返す（単一 .ps1 版）" {
        try {
            $global:bundledVersion = @{ Tag = "v9.9.9"; Sha = "9999999999999999999999999999999999999999" }
            $result = readVersionFile (Join-Path $TestDrive "無い\VERSION.txt")
            $result.Tag | Should -Be "v9.9.9"
            $result.Sha | Should -Be "9999999999999999999999999999999999999999"
        } finally {
            Remove-Variable -Name bundledVersion -Scope Global -ErrorAction SilentlyContinue
        }
    }
}

Describe "readDevVersion" -Tag Io {
    BeforeAll {
        $script:sha1 = "1" * 40
        $script:sha2 = "2" * 40
        $script:sha3 = "3" * 40

        function newGitDir {
            # $root\.git の下に、HEAD・タグを置いた最小の .git を作る
            param (
                [string]$root,
                [string]$head,
                [hashtable]$looseTags = @{},
                [string[]]$packedLines = @(),
                [hashtable]$looseRefs = @{}
            )

            $gitDir = Join-Path $root ".git"
            [System.IO.Directory]::CreateDirectory((Join-Path $gitDir "refs\tags")) | Out-Null
            [System.IO.File]::WriteAllText((Join-Path $gitDir "HEAD"), $head + "`n")
            foreach ($name in $looseTags.Keys) {
                [System.IO.File]::WriteAllText((Join-Path $gitDir "refs\tags\$name"), $looseTags[$name] + "`n")
            }
            foreach ($ref in $looseRefs.Keys) {
                $path = Join-Path $gitDir ($ref.Replace("/", "\"))
                [System.IO.Directory]::CreateDirectory((Split-Path $path -Parent)) | Out-Null
                [System.IO.File]::WriteAllText($path, $looseRefs[$ref] + "`n")
            }
            if ($packedLines.Count -gt 0) {
                [System.IO.File]::WriteAllText((Join-Path $gitDir "packed-refs"), (@("# pack-refs with: peeled fully-peeled sorted") + $packedLines) -join "`n")
            }
            return $gitDir
        }
    }

    It "HEAD のコミットにタグが付いていれば、そのタグを返す（refs\tags のファイル）" {
        $root = "$TestDrive\dev_loose"
        [void](newGitDir $root "ref: refs/heads/main" @{ "v0.4.0" = $script:sha2; "v0.5.0" = $script:sha1 } @() @{ "refs/heads/main" = $script:sha1 })

        $info = readDevVersion $root
        $info.Tag | Should -Be "v0.5.0"
        $info.AtHead | Should -Be $true
        $info.Sha | Should -Be $script:sha1
    }

    It "HEAD のコミットにタグが付いていれば、そのタグを返す（packed-refs だけ）" {
        $root = "$TestDrive\dev_packed"
        [void](newGitDir $root "ref: refs/heads/main" @{} @("$($script:sha1) refs/heads/main", "$($script:sha2) refs/tags/v0.4.0", "$($script:sha1) refs/tags/v0.5.0"))

        $info = readDevVersion $root
        $info.Tag | Should -Be "v0.5.0"
        $info.AtHead | Should -Be $true
        $info.Sha | Should -Be $script:sha1
    }

    It "HEAD にタグが無ければ、版の並びで最も新しいタグを、より後として返す（0.10.0 は 0.9.0 より新しい）" {
        $root = "$TestDrive\dev_after"
        [void](newGitDir $root "ref: refs/heads/main" @{ "v0.9.0" = $script:sha2 } @("$($script:sha2) refs/tags/v0.10.0") @{ "refs/heads/main" = $script:sha3 })

        $info = readDevVersion $root
        $info.Tag | Should -Be "v0.10.0"
        $info.AtHead | Should -Be $false
        $info.Sha | Should -Be $script:sha3
    }

    It "リリースのタグでない名前（v 無し・数が足りない）は無視する" {
        $root = "$TestDrive\dev_other_tags"
        [void](newGitDir $root "ref: refs/heads/main" @{ "backup" = $script:sha1; "v1.0" = $script:sha1; "v0.1.0" = $script:sha2 } @() @{ "refs/heads/main" = $script:sha1 })

        $info = readDevVersion $root
        $info.Tag | Should -Be "v0.1.0"
        $info.AtHead | Should -Be $false
    }

    It "注釈付きタグは、packed-refs の ^ の行（指す先のコミット）で HEAD と照らす" {
        $root = "$TestDrive\dev_annotated"
        [void](newGitDir $root "ref: refs/heads/main" @{} @("$($script:sha2) refs/tags/v0.5.0", "^$($script:sha1)") @{ "refs/heads/main" = $script:sha1 })

        $info = readDevVersion $root
        $info.Tag | Should -Be "v0.5.0"
        $info.AtHead | Should -Be $true
    }

    It "detached HEAD（コミットの SHA が直接書いてある）も読む" {
        $root = "$TestDrive\dev_detached"
        [void](newGitDir $root $script:sha1 @{ "v0.5.0" = $script:sha1 })

        $info = readDevVersion $root
        $info.Tag | Should -Be "v0.5.0"
        $info.AtHead | Should -Be $true
        $info.Sha | Should -Be $script:sha1
    }

    It "タグが無ければ、コミットだけを返す" {
        $root = "$TestDrive\dev_no_tags"
        [void](newGitDir $root "ref: refs/heads/main" @{} @() @{ "refs/heads/main" = $script:sha1 })

        $info = readDevVersion $root
        $info.Tag | Should -Be ""
        $info.AtHead | Should -Be $false
        $info.Sha | Should -Be $script:sha1
    }

    It "worktree（.git がファイルで gitdir: を指し、タグと packed-refs は commondir にある）も読む" {
        $root = "$TestDrive\dev_worktree"
        $common = newGitDir "$TestDrive\dev_worktree_main" "ref: refs/heads/main" @{} @("$($script:sha1) refs/tags/v0.5.0") @{ "refs/heads/main" = $script:sha3 }
        $gitDir = Join-Path $common "worktrees\gui"
        [System.IO.Directory]::CreateDirectory($gitDir) | Out-Null
        [System.IO.Directory]::CreateDirectory($root) | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $gitDir "HEAD"), "ref: refs/heads/worktree-gui`n")
        [System.IO.File]::WriteAllText((Join-Path $gitDir "commondir"), "..\..`n")
        [System.IO.File]::WriteAllText((Join-Path $common "refs\heads_dummy"), "")
        [System.IO.Directory]::CreateDirectory((Join-Path $common "refs\heads")) | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $common "refs\heads\worktree-gui"), $script:sha1 + "`n")
        [System.IO.File]::WriteAllText((Join-Path $root ".git"), "gitdir: $gitDir`n")

        $info = readDevVersion $root
        $info.Tag | Should -Be "v0.5.0"
        $info.AtHead | Should -Be $true
        $info.Sha | Should -Be $script:sha1
    }

    It "worktree の gitdir が相対パスでも読む" {
        $root = "$TestDrive\dev_worktree_rel\wt"
        $common = newGitDir "$TestDrive\dev_worktree_rel\main" "ref: refs/heads/main" @{ "v0.1.0" = $script:sha2 } @() @{ "refs/heads/main" = $script:sha2 }
        $gitDir = Join-Path $common "worktrees\gui"
        [System.IO.Directory]::CreateDirectory($gitDir) | Out-Null
        [System.IO.Directory]::CreateDirectory($root) | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $gitDir "HEAD"), $script:sha3 + "`n")
        [System.IO.File]::WriteAllText((Join-Path $gitDir "commondir"), "../..`n")
        [System.IO.File]::WriteAllText((Join-Path $root ".git"), "gitdir: ../main/.git/worktrees/gui`n")

        $info = readDevVersion $root
        $info.Tag | Should -Be "v0.1.0"
        $info.AtHead | Should -Be $false
        $info.Sha | Should -Be $script:sha3
    }

    It "<name> は `$null を返す（例外にしない）" -TestCases @(
        @{ name = ".git が無い"; setup = { } }
        @{ name = ".git のファイルが gitdir: でない"; setup = { param ($root) [System.IO.File]::WriteAllText("$root\.git", "でたらめ`n") } }
        @{ name = "gitdir の先が無い"; setup = { param ($root) [System.IO.File]::WriteAllText("$root\.git", "gitdir: $root\無い`n") } }
        @{ name = "HEAD もタグも無い .git"; setup = { param ($root) [System.IO.Directory]::CreateDirectory("$root\.git") | Out-Null } }
        @{ name = "HEAD の参照が不正（..を含む）で、タグも無い"; setup = { param ($root) [System.IO.Directory]::CreateDirectory("$root\.git") | Out-Null; [System.IO.File]::WriteAllText("$root\.git\HEAD", "ref: refs/../../x`n") } }
    ) {
        param ($name, $setup)
        $root = "$TestDrive\dev_null_$([guid]::NewGuid().ToString('N').Substring(0, 6))"
        [System.IO.Directory]::CreateDirectory($root) | Out-Null
        & $setup $root

        readDevVersion $root | Should -BeNullOrEmpty
    }

    It "外部プロセス（git）を起動しない" {
        $code = Get-Content -LiteralPath "${scriptsDir}\shared\core\version.ps1" -Raw -Encoding UTF8
        $code | Should -Not -Match '(?m)^\s*(&\s*)?git(\.exe)?\s'
        $code | Should -Not -Match 'Start-Process|ProcessStartInfo'
    }
}