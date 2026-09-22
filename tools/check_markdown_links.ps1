# Markdown のリンク切れを確かめる（Issue #60）。
#
#   .\tools\check_markdown_links.ps1                          git で管理している .md すべて（tests\meta\links.Tests.ps1 から呼ぶ）
#   .\tools\check_markdown_links.ps1 -Path README.md, docs\a.md    指定したファイルだけ（テスト用）
#
# 調べるもの: 本文の相対リンク [文字](先)・画像 ![文字](先)・参照リンクの定義 [名前]: 先・HTML の href="先" / src="先"。
#   - 先のファイル・フォルダがあること。大文字・小文字も区別する（Windows では開けても、GitHub では切れるため）
#   - 先が .md で #アンカーが付いているとき、その見出しがあること。アンカーは GitHub と同じ形で作る（下の getSlug）
#   - リポジトリの外を指していないこと
# 調べないもの: 外部の URL（https: mailto: など。通信が要り、結果が安定しないため）、コードブロック・インラインコードの中、
#   .md 以外のファイルに付いたアンカー。
# docs\ の中は mkdocs build --strict（.github\workflows\docs.yml）でも確かめるが、docs\ の外と、docs\ から外へのリンクはここでしか確かめない。
# 切れたリンクがあれば「ファイル:行: 内容」を出して終了コード 1 で終わる。
param (
    [string[]]$Path,
    [string]$Root = (Join-Path $PSScriptRoot "..")
)

$ErrorActionPreference = "Stop"

$rootDir = [System.IO.Path]::GetFullPath($Root).TrimEnd("\")
$utf8 = New-Object System.Text.UTF8Encoding($false)

# git で管理している .md の一覧（リポジトリの直下からの相対パス）
function getTrackedMarkdown {
    $info = New-Object System.Diagnostics.ProcessStartInfo "git", "-c core.quotepath=off ls-files -z -- *.md"
    $info.WorkingDirectory = $rootDir
    $info.UseShellExecute = $false
    $info.RedirectStandardOutput = $true
    $info.StandardOutputEncoding = $utf8
    $process = [System.Diagnostics.Process]::Start($info)
    $text = $process.StandardOutput.ReadToEnd()
    $process.WaitForExit()
    if ($process.ExitCode -ne 0) { throw "git ls-files が失敗しました" }
    return @($text.Split([char]0) | Where-Object { $_ })
}

# 見出しの文字から GitHub のアンカーを作る（github-slugger と同じ）。
# 英字を小文字にし、文字・数字・_・-・空白以外（記号・全角の括弧・中黒など）を除き、空白を - にする。
# リンク・画像は表示される文字に、HTML のタグは除いてから作る
function getSlug([string]$heading) {
    $text = $heading -replace '!?\[([^\]]*)\]\([^)]*\)', '$1'
    $text = $text -replace '<[^>]+>', ''
    $text = $text.Trim().ToLowerInvariant()
    $text = $text -replace '[^\p{L}\p{M}\p{N}_\- ]', ''
    return $text.Replace(" ", "-")
}

# コードブロックの外の行を返す。Text はインラインコードを空白に置き換えたもの（中の [a](b) をリンクと見ない）。Raw は元の行
function getLines([string]$file) {
    $lines = [System.IO.File]::ReadAllLines($file, $utf8)
    $fence = $null
    $result = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $lines.Length; $i++) {
        $line = $lines[$i]
        if ($line -match '^\s{0,3}(`{3,}|~{3,})') {
            $mark = $Matches[1]
            if (!$fence) {
                $fence = $mark
                continue
            }
            # 閉じるのは、開いたときと同じ文字で、同じ数以上並べた行
            if ($mark[0] -eq $fence[0] -and $mark.Length -ge $fence.Length -and $line.Trim() -eq $mark) {
                $fence = $null
                continue
            }
        }
        if ($fence) { continue }
        $plain = [regex]::Replace($line, '(`+)(?:(?!\1).)+?\1', { param($m) " " * $m.Length })
        $result.Add([pscustomobject]@{ Number = $i + 1; Raw = $line; Text = $plain })
    }
    return $result
}

# .md のアンカーの一覧（見出しと <a id="..."> / <a name="...">）
$anchorCache = @{}
function getAnchors([string]$file) {
    if ($anchorCache.ContainsKey($file)) { return $anchorCache[$file] }
    $anchors = New-Object System.Collections.Generic.HashSet[string]
    $count = @{}
    foreach ($line in (getLines $file)) {
        if ($line.Raw -match '^\s{0,3}#{1,6}\s+(.*?)(?:\s+#+)?\s*$') {
            $slug = getSlug $Matches[1]
            # 同じ見出しが 2 つ目からは -1, -2 … が付く
            if ($count.ContainsKey($slug)) {
                $count[$slug]++
                [void]$anchors.Add("$slug-$($count[$slug])")
            } else {
                $count[$slug] = 0
                [void]$anchors.Add($slug)
            }
        }
        foreach ($m in [regex]::Matches($line.Text, '<a\s[^>]*\b(?:id|name)\s*=\s*"([^"]+)"')) {
            [void]$anchors.Add($m.Groups[1].Value)
        }
    }
    $anchorCache[$file] = $anchors
    return $anchors
}

# フォルダの中の名前の一覧（大文字・小文字を区別して比べるため、実際の名前を取る）
$entryCache = @{}
function testExactPath([string]$fullPath) {
    $relative = $fullPath.Substring($rootDir.Length).TrimStart("\")
    if (!$relative) { return $true }
    $dir = $rootDir
    foreach ($part in $relative.Split("\")) {
        if (!$entryCache.ContainsKey($dir)) {
            $names = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::Ordinal)
            if ([System.IO.Directory]::Exists($dir)) {
                foreach ($entry in [System.IO.Directory]::GetFileSystemEntries($dir)) {
                    [void]$names.Add([System.IO.Path]::GetFileName($entry))
                }
            }
            $entryCache[$dir] = $names
        }
        if (!$entryCache[$dir].Contains($part)) { return $false }
        $dir = Join-Path $dir $part
    }
    return $true
}

# リンクの先を 1 つ調べ、問題があれば理由を返す
function testLink([string]$file, [string]$target) {
    # 外部の URL（https: mailto: など）と、ホストから始まる //example.com は調べない
    if ($target -match '^[A-Za-z][A-Za-z0-9+.\-]*:' -or $target.StartsWith("//")) { return $null }
    $pathPart, $anchor = $target.Split("#", 2)
    $pathPart = [uri]::UnescapeDataString($pathPart.Split("?")[0])
    if (!$pathPart) {
        $full = $file
    } elseif ($pathPart.StartsWith("/")) {
        $full = [System.IO.Path]::GetFullPath((Join-Path $rootDir $pathPart.TrimStart("/")))
    } else {
        $full = [System.IO.Path]::GetFullPath((Join-Path (Split-Path $file) $pathPart))
    }
    $full = $full.TrimEnd("\")
    if ($full -ne $rootDir -and !$full.StartsWith("$rootDir\", [System.StringComparison]::OrdinalIgnoreCase)) {
        return "リポジトリの外を指している: $target"
    }
    if (!(testExactPath $full)) {
        return "リンク先のファイルが無い（大文字・小文字も区別する）: $target"
    }
    if ($anchor -and $full -like "*.md" -and [System.IO.File]::Exists($full)) {
        $anchor = [uri]::UnescapeDataString($anchor)
        if (!(getAnchors $full).Contains($anchor)) {
            return "リンク先に見出し（アンカー）が無い: $target"
        }
    }
    return $null
}

# 行からリンクの先を取り出す
$inlineLink = '!?\[(?:[^\[\]]|\[[^\[\]]*\])*\]\(\s*(?:<(?<target>[^<>]*)>|(?<target>[^()\s]*(?:\([^()\s]*\)[^()\s]*)*))(?:\s+(?:"[^"]*"|''[^'']*''|\([^)]*\)))?\s*\)'
$definition = '^\s{0,3}\[[^\]]+\]:\s*<?(?<target>[^\s>]+)>?'
$htmlLink = '\b(?:href|src)\s*=\s*"(?<target>[^"]*)"'

if (!$Path) { $Path = getTrackedMarkdown }

$problems = New-Object System.Collections.Generic.List[string]
foreach ($name in $Path) {
    $file = [System.IO.Path]::GetFullPath((Join-Path $rootDir $name))
    foreach ($line in (getLines $file)) {
        foreach ($pattern in $inlineLink, $definition, $htmlLink) {
            foreach ($m in [regex]::Matches($line.Text, $pattern)) {
                $target = $m.Groups["target"].Value
                if (!$target) { continue }
                $problem = testLink $file $target
                if ($problem) { $problems.Add("${name}:$($line.Number): $problem") }
            }
        }
    }
}

if ($problems.Count -gt 0) {
    Write-Host "Markdown のリンクが切れています（$($problems.Count) 件）。" -ForegroundColor Red
    $problems | ForEach-Object { Write-Host "  $_" }
    exit 1
}
Write-Host "Markdown のリンクの検査: 問題なし（$($Path.Count) ファイル）"
exit 0
