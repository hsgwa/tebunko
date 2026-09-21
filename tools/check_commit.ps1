# 公開してよい内容かを確かめる（AGENTS.md「個人情報を書かない」と、スクリプトの文字コードの決まり）。
#
#   .\tools\check_commit.ps1           追跡しているファイル全部と、HEAD までのコミットの作者を調べる（CI）
#   .\tools\check_commit.ps1 -Staged   ステージしたファイルと、これから作るコミットの作者を調べる（pre-commit フック）
#
# 調べること:
#   - 利用者名を含む絶対パス（C:\Users\<名前>。test などの例示用の名前は除く）
#   - メールアドレス（example.com などの例示用のドメインと、GitHub の noreply は除く）
#   - Office ファイルの作成者・最終更新者（test か空であること）と、コメント等に付く Microsoft アカウント ID
#     （Office ファイルは ZIP を展開して中の XML も調べる。旧形式・PDF はバイト列をそのまま調べる）
#   - 手元で実行したときは、Windows のユーザー名そのもの
#   - .ps1・.xaml が BOM 付き UTF-8・CRLF であること
#   - コミットの作者・コミッターのメールアドレスが GitHub の noreply であること
#
# ファイルの中身は作業ツリーではなく git に記録される内容（インデックス。-Staged のときはステージした分だけ）を読む
# （改行だけは作業ツリーで調べる。findEncodingProblems）。
# 見つかったら一覧を出して終了コード 1 で終わる。
param (
    [switch]$Staged
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression

$rootDir = Split-Path $PSScriptRoot -Parent

# 例示用として書いてよい名前（C:\Users\<名前> と Office の作成者名）
$allowedNames = @("", "test", "a", "public", "default")
# 例示用として書いてよいメールアドレスのドメイン
$allowedMailDomain = '(?i)(^|\.)(example\.(com|org|net|jp|co\.jp)|users\.noreply\.github\.com)$'
# コミットの作者として認めるメールアドレス（GitHub 上でマージしたときのコミッターは noreply@github.com になる）
$allowedCommitMail = '(?i)(@users\.noreply\.github\.com|^noreply@github\.com)$'

$utf8 = New-Object System.Text.UTF8Encoding($false)

# git を実行し、標準出力をバイト列で返す（-z の出力や、バイナリの中身をそのまま受け取るため）
function invokeGit([string]$arguments, [byte[]]$stdin) {
    $info = New-Object System.Diagnostics.ProcessStartInfo "git", $arguments
    $info.WorkingDirectory = $rootDir
    $info.UseShellExecute = $false
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardInput = $true
    $info.RedirectStandardError = $true
    $process = [System.Diagnostics.Process]::Start($info)
    $buffer = New-Object System.IO.MemoryStream
    # 標準出力を並行して読む（読まずに標準入力を書き続けると、パイプが詰まって止まる）
    $copy = $process.StandardOutput.BaseStream.CopyToAsync($buffer)
    $errors = $process.StandardError.ReadToEndAsync()
    if ($stdin) {
        $process.StandardInput.BaseStream.Write($stdin, 0, $stdin.Length)
    }
    $process.StandardInput.Close()
    $copy.Wait()
    $process.WaitForExit()
    if ($process.ExitCode -ne 0) {
        throw "git $arguments が失敗しました: $($errors.Result)"
    }
    return ,$buffer.ToArray()
}

function gitText([string]$arguments) {
    return $utf8.GetString((invokeGit $arguments $null))
}

# 調べるファイルの一覧（パスと blob の SHA）
function getTargets {
    $targets = New-Object System.Collections.Generic.List[object]
    if ($Staged) {
        # 追加・変更したファイルだけ。削除したファイルは調べない
        $fields = (gitText "-c core.quotepath=off diff --cached --raw -z --no-renames --no-abbrev --diff-filter=ACMRT").Split([char]0)
        for ($i = 0; $i + 1 -lt $fields.Length; $i += 2) {
            $meta = $fields[$i].Split(" ")
            # サブモジュール（160000）は中身を持たない
            if ($meta[1] -eq "160000") { continue }
            $targets.Add([pscustomobject]@{ Path = $fields[$i + 1]; Sha = $meta[3] })
        }
    } else {
        foreach ($record in (gitText "-c core.quotepath=off ls-files -s -z").Split([char]0)) {
            if (!$record) { continue }
            $meta, $path = $record.Split("`t", 2)
            $parts = $meta.Split(" ")
            if ($parts[0] -eq "160000") { continue }
            $targets.Add([pscustomobject]@{ Path = $path; Sha = $parts[1] })
        }
    }
    return $targets
}

# blob の中身をまとめて読む（git cat-file --batch。1 回の起動で済ませる）
function readBlobs($targets) {
    $contents = @{}
    if ($targets.Count -eq 0) { return $contents }
    # 先頭に空行を 1 つ送る。.NET Framework は標準入力の先頭に BOM を書いてしまうため、その行は missing として読み捨てる
    $request = $utf8.GetBytes("`n" + (($targets | ForEach-Object { $_.Sha }) -join "`n") + "`n")
    $bytes = invokeGit "cat-file --batch" $request
    $position = 0
    while ($position -lt $bytes.Length) {
        $lineEnd = [Array]::IndexOf($bytes, [byte]10, $position)
        $header = [System.Text.Encoding]::ASCII.GetString($bytes, $position, $lineEnd - $position).Split(" ")
        if ($header[-1] -eq "missing") {
            $position = $lineEnd + 1
            continue
        }
        $size = [int]$header[2]
        $content = New-Object byte[] $size
        [Array]::Copy($bytes, $lineEnd + 1, $content, 0, $size)
        $contents[$header[0]] = $content
        $position = $lineEnd + 1 + $size + 1
    }
    return $contents
}

function isAllowedName([string]$name) {
    $name = $name.Trim()
    # %USERNAME% や <利用者名> のような置き換え用の表記
    if ($name -match '^[%<$]') { return $true }
    # test に置き換えた名前（バイナリ内では元の長さに揃えて test_ のようにしてある）と、0 で埋めたアカウント ID
    if ($name -match '^(?i)test[_0-9]*$' -or $name -match '^0+$') { return $true }
    return $allowedNames -contains $name.ToLowerInvariant()
}

# 文字列から公開してはいけないものを探す。見つけたものの説明を返す
function findInText([string]$text) {
    $found = New-Object System.Collections.Generic.List[string]

    foreach ($match in [regex]::Matches($text, '(?i)\b[A-Z]:[\\/]{1,2}Users[\\/]{1,2}([^\\/"''<>|\s:*?]+)')) {
        if (!(isAllowedName $match.Groups[1].Value)) {
            $found.Add("利用者名を含むパス: $($match.Value)")
        }
    }
    foreach ($match in [regex]::Matches($text, '[A-Za-z0-9._%+-]+@([A-Za-z0-9-]+(\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,})')) {
        if ($match.Groups[1].Value -notmatch $allowedMailDomain -and $match.Value -notmatch $allowedCommitMail) {
            $found.Add("メールアドレス: $($match.Value)")
        }
    }
    # Office の作成者・最終更新者（OOXML の docProps/core.xml、ODF の meta.xml、PDF の /Author）
    $authorPattern = '<(dc:creator|cp:lastModifiedBy|meta:initial-creator)>([^<]*)</\1>|/Author\s*\(([^)]*)\)'
    foreach ($match in [regex]::Matches($text, $authorPattern)) {
        $name = if ($match.Groups[1].Success) { $match.Groups[2].Value } else { $match.Groups[3].Value }
        if (!(isAllowedName $name)) {
            $found.Add("作成者名: $($match.Value)")
        }
    }
    # コメント・変更履歴の作成者に付く Microsoft アカウント ID（word/people.xml・ppt/authors.xml など）
    foreach ($match in [regex]::Matches($text, '(?i)\b(w15:)?userId="([^"<>()\[\]]*)"')) {
        if (!(isAllowedName $match.Groups[2].Value)) {
            $found.Add("アカウント ID: $($match.Value)")
        }
    }
    # 手元で実行したときは、Windows のユーザー名そのものも探す（CI のランナーのユーザー名は意味が無い）
    if ($env:GITHUB_ACTIONS -ne "true" -and $env:USERNAME.Length -ge 3 -and !(isAllowedName $env:USERNAME)) {
        if ($text.IndexOf($env:USERNAME, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            $found.Add("Windows のユーザー名: $env:USERNAME")
        }
    }
    return $found
}

# バイト列を調べる。UTF-8 として読み、NUL を含む（バイナリ・UTF-16 の）ものは UTF-16LE としても読む
function findInBytes([byte[]]$bytes) {
    $found = @(findInText $utf8.GetString($bytes))
    if ([Array]::IndexOf($bytes, [byte]0) -ge 0) {
        $found += @(findInText ([System.Text.Encoding]::Unicode.GetString($bytes)))
    }
    return $found
}

function isZip([byte[]]$bytes) {
    return $bytes.Length -ge 4 -and $bytes[0] -eq 0x50 -and $bytes[1] -eq 0x4B -and $bytes[2] -eq 0x03 -and $bytes[3] -eq 0x04
}

# ZIP（Office の新形式・ODF）は展開して、中のファイルを 1 つずつ調べる
function findInZip([byte[]]$bytes) {
    $found = New-Object System.Collections.Generic.List[string]
    $archive = New-Object System.IO.Compression.ZipArchive((New-Object System.IO.MemoryStream(, $bytes)), [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        foreach ($entry in $archive.Entries) {
            $buffer = New-Object System.IO.MemoryStream
            $stream = $entry.Open()
            try { $stream.CopyTo($buffer) } finally { $stream.Dispose() }
            foreach ($item in (findInBytes $buffer.ToArray())) {
                $found.Add("$($entry.FullName) の $item")
            }
        }
    } finally {
        $archive.Dispose()
    }
    return $found
}

# .ps1・.xaml は BOM 付き UTF-8・CRLF（PowerShell 5.1 は BOM が無いと CP932 で読むため）。
# git は改行を LF にして記録する（core.autocrlf）ため、BOM は記録された内容で、改行は作業ツリーのファイルで調べる
function findEncodingProblems([string]$path, [byte[]]$bytes) {
    if ($path -notmatch '(?i)\.(ps1|xaml)$') { return @() }
    $problems = @()
    if (!($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)) {
        $problems += "BOM 付き UTF-8 ではない"
    }
    $workingFile = Join-Path $rootDir $path
    if ((Test-Path -LiteralPath $workingFile) -and
        [regex]::IsMatch([System.IO.File]::ReadAllText($workingFile), "(?<!`r)`n")) {
        $problems += "改行が CRLF ではない"
    }
    return $problems
}

$problems = New-Object System.Collections.Generic.List[string]

$targets = @(getTargets)
$contents = readBlobs $targets
foreach ($target in $targets) {
    $bytes = $contents[$target.Sha]
    $items = @()
    if (isZip $bytes) {
        try {
            $items += @(findInZip $bytes)
        } catch {
            # 壊れた ZIP（テストデータの異常系など）はバイト列として調べる
            $items += @(findInBytes $bytes)
        }
    } else {
        $items += @(findInBytes $bytes)
    }
    $items += @(findEncodingProblems $target.Path $bytes)
    foreach ($item in ($items | Select-Object -Unique)) {
        $problems.Add("$($target.Path): $item")
    }
}

# コミットの作者・コミッター
if ($Staged) {
    $identities = @("GIT_AUTHOR_IDENT", "GIT_COMMITTER_IDENT" | ForEach-Object { gitText "var $_" })
    $mails = @($identities | ForEach-Object { if ($_ -match '<([^>]*)>') { $Matches[1] } })
    $label = "これから作るコミット"
} else {
    $mails = @((gitText "log --format=%ae%n%ce HEAD").Split("`n") | Where-Object { $_ })
    $label = "HEAD までのコミット"
}
foreach ($mail in ($mails | Select-Object -Unique)) {
    if ($mail -notmatch $allowedCommitMail) {
        $problems.Add("${label}の作者・コミッター: $mail（git config user.email <ID>+<アカウント名>@users.noreply.github.com にする）")
    }
}

if ($problems.Count -gt 0) {
    Write-Host "公開してはいけない内容が見つかりました（$($problems.Count) 件）。" -ForegroundColor Red
    $problems | ForEach-Object { Write-Host "  $_" }
    exit 1
}
Write-Host "個人情報・文字コードの検査: 問題なし（$($targets.Count) ファイル）"
