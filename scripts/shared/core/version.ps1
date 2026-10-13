# 配布物の版の記録（VERSION.txt）の読み取り（状態層）。
# 中身は BOM 付き UTF-8・CRLF の2行（1行目: タグ名 / 2行目: コミットの SHA）。
# tools/new_version_text.ps1 が作り、tools/new_release_package.ps1（zip）・tools/new_installer.ps1（インストーラー）が書き込む。
# 作業ツリー（git のチェックアウト）には無く、配布物にだけある。
#
# 単一 .ps1 版（展開せずに動く試験版）は VERSION.txt を持てないため、結合の道具（tools/new_single_script.ps1）が
# 版の文字列を ${bundledVersion}（@{ Tag; Sha }）としてこのファイルより前に埋める。あればファイルを読まずそれを返す

function readVersionFile {
    # ファイルが無い・読めない・形が違うときは $null。あれば @{ Tag; Sha }
    param (
        [string]$path
    )

    $bundled = Get-Variable -Name bundledVersion -ErrorAction SilentlyContinue
    if ($null -ne $bundled) {
        return $bundled.Value
    }

    if (!(Test-Path -LiteralPath $path)) {
        return $null
    }
    try {
        $lines = [System.IO.File]::ReadAllLines($path)
    } catch {
        return $null
    }
    if ($lines.Count -ne 2) {
        return $null
    }
    if ($lines[0] -notmatch '^[A-Za-z0-9._-]+$' -or $lines[1] -notmatch '^[0-9a-f]{40}$') {
        return $null
    }
    return @{ Tag = $lines[0]; Sha = $lines[1] }
}

function resolveGitRef {
    # git のリファレンス（refs/heads/main など）が指すコミットの SHA を返す。無ければ空文字。
    # ファイルを読むだけ（外部プロセスの git は起動しない）。loose（refs\ の下のファイル）を先に見て、無ければ packed-refs を見る。
    #   dirs: 探すフォルダ（worktree の gitdir・共通の gitdir の順）
    param (
        [string]$ref,
        [string[]]$dirs
    )

    if ($ref -notmatch '^refs/[A-Za-z0-9._/-]+$' -or $ref.Contains("..")) {
        return ""
    }
    foreach ($dir in $dirs) {
        $loose = Join-Path $dir ($ref.Replace("/", "\"))
        if (Test-Path -LiteralPath $loose -PathType Leaf) {
            $text = ([System.IO.File]::ReadAllText($loose)).Trim()
            if ($text -match '^[0-9a-f]{40}$') {
                return $text
            }
        }
    }
    foreach ($dir in $dirs) {
        $refListFile = Join-Path $dir "packed-refs"
        if (!(Test-Path -LiteralPath $refListFile -PathType Leaf)) {
            continue
        }
        foreach ($line in [System.IO.File]::ReadAllLines($refListFile)) {
            if ($line -match '^([0-9a-f]{40}) (\S+)$' -and $Matches[2] -ceq $ref) {
                return $Matches[1]
            }
        }
    }
    return ""
}

function readDevVersion {
    # 配布物でない（VERSION.txt が無い）ときの版の手がかりを、作業ツリーの .git のファイルだけから読む。
    # 外部プロセス（git）は起動せず、ファイルを読むだけ。.git が無い・読めない・形が違うときは $null。
    # 返すもの: @{ Tag; AtHead; Sha }
    #   Sha    : HEAD のコミット（分からなければ空）
    #   Tag    : リリースのタグ（v<数>.<数>.<数>）。HEAD のコミットに付いたタグがあればそれ（複数なら版の並びで最も新しい）、
    #            無ければ全タグのうち版の並びで最も新しいもの（無ければ空）
    #   AtHead : Tag が HEAD のコミットに付いているか
    # 「最も新しいタグ」は版の並びで決める。.git のファイルだけでは、タグが HEAD の祖先かどうかは分からず、
    # 別のブランチのタグを拾うことがある（目安の表示のため、それでよい）。
    # 注釈付きタグは、packed-refs の ^ の行（指す先のコミット）で HEAD と照らす。refs\tags の下に置かれた注釈付きタグは、
    # タグの実体を読まないと指す先が分からないため、HEAD に付いた印にはならない（版の並びの比べには使う）。
    # worktree は .git がファイル（gitdir: <パス>）で、HEAD は gitdir に、タグ・packed-refs は commondir（共通の gitdir）にある
    param (
        [string]$rootDir
    )

    try {
        $dotGit = Join-Path $rootDir ".git"
        $gitDir = ""
        if (Test-Path -LiteralPath $dotGit -PathType Container) {
            $gitDir = $dotGit
        } elseif (Test-Path -LiteralPath $dotGit -PathType Leaf) {
            $first = [System.IO.File]::ReadAllLines($dotGit) | Select-Object -First 1
            if ($first -notmatch '^gitdir:\s*(.+)$') {
                return $null
            }
            $target = $Matches[1].Trim()
            if (![System.IO.Path]::IsPathRooted($target)) {
                $target = Join-Path $rootDir $target
            }
            $gitDir = [System.IO.Path]::GetFullPath($target)
        }
        if (!$gitDir -or !(Test-Path -LiteralPath $gitDir -PathType Container)) {
            return $null
        }
        $commonDir = $gitDir
        $commonFile = Join-Path $gitDir "commondir"
        if (Test-Path -LiteralPath $commonFile -PathType Leaf) {
            $common = ([System.IO.File]::ReadAllText($commonFile)).Trim()
            if ($common) {
                $commonDir = [System.IO.Path]::GetFullPath($(if ([System.IO.Path]::IsPathRooted($common)) { $common } else { Join-Path $gitDir $common }))
            }
        }

        # HEAD のコミット
        $sha = ""
        $headFile = Join-Path $gitDir "HEAD"
        if (Test-Path -LiteralPath $headFile -PathType Leaf) {
            $head = ([System.IO.File]::ReadAllLines($headFile) | Select-Object -First 1)
            if ($head -match '^ref:\s*(\S+)$') {
                $sha = resolveGitRef $Matches[1] @($gitDir, $commonDir)
            } elseif ($head -match '^[0-9a-f]{40}$') {
                $sha = $head
            }
        }

        # タグ（名前 → @{ Commit }）。packed-refs（注釈付きタグは次の行の ^ が指す先）を読み、loose のタグで上書きする
        $tags = @{}
        $refListFile = Join-Path $commonDir "packed-refs"
        if (Test-Path -LiteralPath $refListFile -PathType Leaf) {
            $last = ""
            foreach ($line in [System.IO.File]::ReadAllLines($refListFile)) {
                if ($line -match '^([0-9a-f]{40}) refs/tags/(v\d+\.\d+\.\d+)$') {
                    $last = $Matches[2]
                    $tags[$last] = $Matches[1]
                } elseif ($line -match '^\^([0-9a-f]{40})$' -and $last) {
                    $tags[$last] = $Matches[1]
                } else {
                    $last = ""
                }
            }
        }
        $tagsDir = Join-Path $commonDir "refs\tags"
        if (Test-Path -LiteralPath $tagsDir -PathType Container) {
            foreach ($file in @(Get-ChildItem -LiteralPath $tagsDir -File)) {
                if ($file.Name -match '^v\d+\.\d+\.\d+$') {
                    $text = ([System.IO.File]::ReadAllText($file.FullName)).Trim()
                    if ($text -match '^[0-9a-f]{40}$') {
                        $tags[$file.Name] = $text
                    }
                }
            }
        }

        $atHead = ""
        $latest = ""
        foreach ($name in @($tags.Keys)) {
            $isNewer = { param($a, $b) (-not $b) -or ([version]$a.Substring(1)) -gt ([version]$b.Substring(1)) }
            if (& $isNewer $name $latest) {
                $latest = $name
            }
            if ($sha -and $tags[$name] -eq $sha -and (& $isNewer $name $atHead)) {
                $atHead = $name
            }
        }
        if ($atHead) {
            return @{ Tag = $atHead; AtHead = $true; Sha = $sha }
        }
        if (!$latest -and !$sha) {
            return $null
        }
        return @{ Tag = $latest; AtHead = $false; Sha = $sha }
    } catch {
        return $null
    }
}
