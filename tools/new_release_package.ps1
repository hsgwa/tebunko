# 配布する zip を作る（GitHub Actions の .github\workflows\release.yml が使う。手元でも実行できる）。
#
#   .\tools\new_release_package.ps1 -Version v1.0.0     work\release\ に zip と、zip の横に並べるファイルを作る
#
# zip の中身（展開すると tebunko\ フォルダになる）。展開したときに、起動するもの（tebunko.bat）と使い方がすぐ分かるものだけにする:
#   tebunko.bat scripts\                              ツール本体
#   README.md                                         使い方。相対リンクと画像は、その版の GitHub の URL に書き換える
#                                                     （docs\ や画像は zip に入れないため。ページ内のリンク #… はそのまま）
#   LICENSE                                           ライセンス（MIT。写しに許諾表示を含めるため同梱する）
#
# zip の横に並べて、GitHub Release に載せるもの（release.yml）:
#   tebunko-<版>.zip
#   tebunko.cat SHA256SUMS.txt                        改ざんの確認用（tools\new_release_files.ps1 が作る）
#   sbom.cdx.json                                     部品表
# SECURITY は zip に入れず、README とリリースの説明からリンクする。
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

# zip の横に並べるもの: カタログとハッシュ一覧（配布物と同じ中身から作る）と部品表
& (Join-Path $PSScriptRoot "new_release_files.ps1") -OutDir $OutDir
Copy-Item -LiteralPath (Join-Path $rootDir "sbom.cdx.json") -Destination $OutDir -Force

# zip に入れるファイル（zip 内のパス → 元のファイル）
$entries = [ordered]@{}
foreach ($file in @(Get-ChildItem -LiteralPath (Join-Path $rootDir "scripts") -Recurse -File | Sort-Object FullName)) {
    $entries[$file.FullName.Substring($rootDir.Length + 1)] = $file.FullName
}
foreach ($name in @("tebunko.bat", "README.md", "LICENSE")) {
    $entries[$name] = Join-Path $rootDir $name
}

# README の相対リンク（](docs/…)）と画像（src="docs/…"）を、その版の GitHub の URL にする。
# 展開したフォルダには docs\ などが無く、そのままでは画像が出ずリンクも切れるため
$readmePath = $entries["README.md"]
$readme = [System.IO.File]::ReadAllText($readmePath, [System.Text.Encoding]::UTF8)
$readme = [regex]::Replace($readme, '\]\((?!https?:|#|mailto:)([^)\s]+)\)', "](https://github.com/$Repository/blob/$Version/`$1)")
$readme = [regex]::Replace($readme, 'src="(?!https?:)([^"]+)"', "src=`"https://raw.githubusercontent.com/$Repository/$Version/`$1`"")
$readmeBytes = (New-Object System.Text.UTF8Encoding($true)).GetPreamble() + (New-Object System.Text.UTF8Encoding($false)).GetBytes($readme)

$zipPath = Join-Path $OutDir "tebunko-$Version.zip"
if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}
# Compress-Archive（5.1）は zip 内の区切りを \ にするため使わず、/ で書く
$stream = [System.IO.File]::Open($zipPath, [System.IO.FileMode]::CreateNew)
$archive = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($name in $entries.Keys) {
        $entry = $archive.CreateEntry("tebunko/" + $name.Replace("\", "/"), [System.IO.Compression.CompressionLevel]::Optimal)
        $entry.LastWriteTime = (Get-Item -LiteralPath $entries[$name]).LastWriteTime
        $writer = $entry.Open()
        try {
            if ($name -eq "README.md") {
                $bytes = $readmeBytes
            } else {
                $bytes = [System.IO.File]::ReadAllBytes($entries[$name])
            }
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
