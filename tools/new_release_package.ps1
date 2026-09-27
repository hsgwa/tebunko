# 配布する zip を作る（GitHub Actions の .github\workflows\release.yml が使う。手元でも実行できる）。
#
#   .\tools\new_release_package.ps1 -Version v1.0.0     work\release\ に zip と、zip の横に並べるファイルを作る
#
# zip の中身（展開すると tebunko\ フォルダになる）。展開したときに、起動するもの（tebunko.bat）と使い方がすぐ分かるものだけにする:
#   tebunko.bat scripts\                              ツール本体
#   README.md                                         使い方。相対リンクと画像は、その版の GitHub の URL に書き換える
#                                                     （docs\ や画像は zip に入れないため。ページ内のリンク #… はそのまま）
#   LICENSE                                           ライセンス（MIT。写しに許諾表示を含めるため同梱する）
#   VERSION.txt                                       版とコミットの SHA（tools\new_version_text.ps1 が作る。画面の「tebunko について」）
#
# zip の横に並べて、GitHub Release に載せるもの（release.yml）:
#   tebunko-<版>.zip
#   tebunko.cat SHA256SUMS.txt                        改ざんの確認用（tools\new_release_files.ps1 が作る）
#   sbom.cdx.json                                     部品表（tools\new_sbom.ps1 が zip の中身から作る。リポジトリの sbom.cdx.json は雛形）
# SECURITY は zip に入れず、README とリリースの説明からリンクする。
#
# 同じコミット・同じ版の名前で作り直すと、zip の中身・SHA256SUMS.txt・sbom.cdx.json が同じになる（docs\safety\scans.md「再現の条件」）。
#   ファイルの時刻はコミットの時刻（UTC の時計の値）、ファイルの一覧は git で追跡しているものだけを序数の順、
#   テキストの改行は CRLF にそろえる。tebunko.cat・インストーラーは時刻を埋め込むため対象外。
#
# work\・setting.config は利用者ごとに作られるため入れない。docs\・tests\ も配布しない。
param (
    [Parameter(Mandatory = $true)]
    [string]$Version,
    [string]$OutDir,
    # README のリンクの書き換え先（GitHub の <持ち主>/<リポジトリ>）
    [string]$Repository = "hsgwa/tebunko"
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression

$rootDir = Split-Path $PSScriptRoot -Parent
if (!$OutDir) {
    $OutDir = Join-Path $rootDir "work\release"
}
if ($Version -notmatch '^[A-Za-z0-9._-]+$') {
    throw "バージョンに使えない文字が含まれています: $Version"
}

# コミットの時刻と SHA（取れないときは、ここで止めて何も作らない）。
# zip のファイルの時刻・SHA256SUMS.txt・部品表の時刻は、これに固定する（作った時刻を使うと、作るたびに中身が変わる）
$commitTime = & git -C $rootDir log -1 --format=%ct HEAD 2>$null
if ($LASTEXITCODE -ne 0 -or $commitTime -notmatch '^[0-9]+$') {
    throw "コミットの時刻を取得できませんでした（git log -1 --format=%ct HEAD）。git のリポジトリで実行してください。"
}
# オフセット 0 の値。zip の時刻（DOS 時刻）は時差を持たず、この時計の値がそのまま書かれる（手元の時差で値が変わらない）
$timestamp = [DateTimeOffset]::FromUnixTimeSeconds([long]$commitTime)
$sha = (& git -C $rootDir rev-parse HEAD).Trim()

# VERSION.txt の中身を先に作る（git の SHA が取れないときは、ここで止めて何も作らない）。
# リポジトリの作業ツリーには書かず、下の zip のエントリーへバイト列のまま直接書く
$versionBytes = & (Join-Path $PSScriptRoot "new_version_text.ps1") -Version $Version

# テキストの改行を CRLF にそろえる。いったん LF にしてから CRLF にする（CRLF のまま置き換えると \r\r\n になるため）
function ConvertTo-Crlf([string]$Text) {
    $Text.Replace("`r`n", "`n").Replace("`n", "`r`n")
}

# README の相対リンク（](docs/…)）と画像（src="docs/…"）を、その版の GitHub の URL にする。
# 展開したフォルダには docs\ などが無く、そのままでは画像が出ずリンクも切れるため
$readme = [System.IO.File]::ReadAllText((Join-Path $rootDir "README.md"), [System.Text.Encoding]::UTF8)
$readme = [regex]::Replace($readme, '\]\((?!https?:|#|mailto:)([^)\s]+)\)', "](https://github.com/$Repository/blob/$Version/`$1)")
$readme = [regex]::Replace($readme, 'src="(?!https?:)([^"]+)"', "src=`"https://raw.githubusercontent.com/$Repository/$Version/`$1`"")
$readmeBytes = (New-Object System.Text.UTF8Encoding($true)).GetPreamble() + (New-Object System.Text.UTF8Encoding($false)).GetBytes((ConvertTo-Crlf $readme))

# LICENSE は元の BOM の有無を保ったまま、改行だけそろえる
$licenseRaw = [System.IO.File]::ReadAllBytes((Join-Path $rootDir "LICENSE"))
$hasBom = $licenseRaw.Length -ge 3 -and $licenseRaw[0] -eq 0xEF -and $licenseRaw[1] -eq 0xBB -and $licenseRaw[2] -eq 0xBF
$licenseText = ConvertTo-Crlf ([System.Text.Encoding]::UTF8.GetString($licenseRaw, $(if ($hasBom) { 3 } else { 0 }), $licenseRaw.Length - $(if ($hasBom) { 3 } else { 0 })))
$licenseBytes = (New-Object System.Text.UTF8Encoding($hasBom)).GetPreamble() + (New-Object System.Text.UTF8Encoding($false)).GetBytes($licenseText)

# zip に入れるファイル（zip 内のパス → 入れるバイト列）。scripts\ は git で追跡しているものだけ
# （手元の追跡していないファイルを入れない）。並びは序数（機械の言語設定に依らない）
$paths = New-Object System.Collections.Generic.List[string]
foreach ($file in @(& git -C $rootDir -c core.quotepath=false ls-files scripts)) {
    $paths.Add($file)
}
if ($LASTEXITCODE -ne 0 -or $paths.Count -eq 0) {
    throw "scripts のファイルの一覧を取得できませんでした（git ls-files scripts）。"
}
foreach ($name in @("tebunko.bat", "README.md", "LICENSE", "VERSION.txt")) {
    $paths.Add($name)
}
$sorted = $paths.ToArray()
[Array]::Sort($sorted, [StringComparer]::Ordinal)
$entries = [ordered]@{}
foreach ($name in $sorted) {
    if ($name -eq "README.md") {
        $entries[$name] = [byte[]]$readmeBytes
    } elseif ($name -eq "LICENSE") {
        $entries[$name] = [byte[]]$licenseBytes
    } elseif ($name -eq "VERSION.txt") {
        $entries[$name] = [byte[]]$versionBytes
    } else {
        $entries[$name] = [System.IO.File]::ReadAllBytes((Join-Path $rootDir $name))
    }
}

# zip の横に並べるもの: 部品表（zip に入れるバイト列から作る）と、カタログ・ハッシュ一覧
[System.IO.Directory]::CreateDirectory($OutDir) | Out-Null
$sbomBytes = & (Join-Path $PSScriptRoot "new_sbom.ps1") -Version $Version -Sha $sha -Timestamp $timestamp -Entries $entries
[System.IO.File]::WriteAllBytes((Join-Path $OutDir "sbom.cdx.json"), [byte[]]$sbomBytes)
& (Join-Path $PSScriptRoot "new_release_files.ps1") -OutDir $OutDir -Entries $entries -SbomBytes $sbomBytes -Timestamp $timestamp

$zipPath = Join-Path $OutDir "tebunko-$Version.zip"
if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}
# Compress-Archive（5.1）は zip 内の区切りを \ にするため使わず、/ で書く
$stream = [System.IO.File]::Open($zipPath, [System.IO.FileMode]::CreateNew)
$archive = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($name in $entries.Keys) {
        $entry = $archive.CreateEntry("tebunko/" + $name, [System.IO.Compression.CompressionLevel]::Optimal)
        $entry.LastWriteTime = $timestamp
        $bytes = [byte[]]$entries[$name]
        $writer = $entry.Open()
        try {
            $writer.Write($bytes, 0, $bytes.Length)
        } finally {
            $writer.Dispose()
        }
    }
} finally {
    $archive.Dispose()
    $stream.Dispose()
}

$hash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash
Write-Host "配布物: $zipPath（$($entries.Count) ファイル）"
Write-Host "SHA256: $hash"
