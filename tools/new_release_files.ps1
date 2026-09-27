# 配布物の完全性を確かめるためのファイルを作る（docs\safety\scans.md「配布物の完全性（カタログ・ハッシュ一覧・来歴の署名）」）。
#
#   .\tools\new_release_files.ps1 -Entries $entries -SbomBytes $sbomBytes                 work\release\ に書き出す
#   .\tools\new_release_files.ps1 -Entries $entries -SbomBytes $sbomBytes -OutDir .\dist
#
# -Entries は zip 内のパス（tebunko\ を除き / 区切り）→ zip に入れたバイト列（順序付き）、-SbomBytes は作った部品表のバイト列。
# どちらも tools\new_release_package.ps1 が作って渡す。-Timestamp を省くと、コミットの時刻（git log -1 --format=%ct）にする。
#
# 出力:
#   tebunko.cat       カタログ（各ファイルの SHA256。Test-FileCatalog で検証する。作った時刻を含むため、作るたびに中身が変わる）
#   SHA256SUMS.txt     ファイルごとの SHA256（テキストで目視・比較できる形式）。zip に入れたバイト列から計算する。
#                      同じコミット・同じ版なら同じ中身になる（見出しの時刻はコミットの時刻）
#
# 2 つとも配布する zip には入れず、zip と並べて GitHub Release に載せる（tools\new_release_package.ps1・release.yml）。
# 受け取った側は、zip を展開したフォルダで次のコマンドを実行すると、配布時点から
# 1 バイトも変わっていないことを自分で確認できる（証明書は要らない）。
#
#   Test-FileCatalog -Path .\scripts, .\tebunko.bat -CatalogFilePath <ダウンロードした tebunko.cat> -Detailed
#
# Status が Valid なら改ざんなし。ValidationFailed なら、どのファイルが違うかが表示される。
# コードサイニング証明書がある場合は、カタログに署名すると発行者の保証も付く。
#   Set-AuthenticodeSignature -FilePath .\tebunko.cat -Certificate $cert
param (
    [Parameter(Mandatory = $true)]
    [System.Collections.Specialized.IOrderedDictionary]$Entries,
    [Parameter(Mandatory = $true)]
    [byte[]]$SbomBytes,
    [string]$OutDir,
    # SHA256SUMS.txt の見出しの時刻（省略時はコミットの時刻）
    [DateTimeOffset]$Timestamp
)

$ErrorActionPreference = "Stop"

$rootDir = Split-Path $PSScriptRoot -Parent
if (!$OutDir) {
    $OutDir = Join-Path $rootDir "work\release"
}
[System.IO.Directory]::CreateDirectory($OutDir) | Out-Null

if (!$PSBoundParameters.ContainsKey("Timestamp")) {
    $commitTime = & git -C $rootDir log -1 --format=%ct HEAD 2>$null
    if ($LASTEXITCODE -ne 0 -or $commitTime -notmatch '^[0-9]+$') {
        throw "コミットの時刻を取得できませんでした（git log -1 --format=%ct HEAD）。git のリポジトリで実行してください。"
    }
    $Timestamp = [DateTimeOffset]::FromUnixTimeSeconds([long]$commitTime)
}

# 配布物のうち、内容が固定されているもの（work・setting.config は利用者ごとに変わるため含めない）
$targets = @(
    (Join-Path $rootDir "scripts"),
    (Join-Path $rootDir "tebunko.bat")
)

$catalogPath = Join-Path $OutDir "tebunko.cat"
if (Test-Path -LiteralPath $catalogPath) {
    Remove-Item -LiteralPath $catalogPath -Force
}
# CatalogVersion 2 = SHA256（1 は SHA1 のため使わない）
New-FileCatalog -Path $targets -CatalogFilePath $catalogPath -CatalogVersion 2 | Out-Null

$result = Test-FileCatalog -Path $targets -CatalogFilePath $catalogPath -Detailed
if ($result.Status -ne "Valid") {
    throw "作ったカタログの検証に失敗しました（Status: $($result.Status)）。"
}

$sumsPath = Join-Path $OutDir "SHA256SUMS.txt"
$sha256 = [System.Security.Cryptography.SHA256]::Create()
try {
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("# tebunko 配布物の SHA256（$($Timestamp.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ')) 時点）")
    $lines.Add("# 確認: Get-FileHash <ファイル> -Algorithm SHA256")
    # パスは zip を展開した tebunko\ から見た相対パス。zip に入るファイルの後に、zip の横に置く部品表を並べる
    foreach ($path in $Entries.Keys) {
        $hash = -join ($sha256.ComputeHash([byte[]]$Entries[$path]) | ForEach-Object { $_.ToString("X2") })
        $lines.Add("$hash  $($path.Replace('/', '\'))")
    }
    $sbomHash = -join ($sha256.ComputeHash($SbomBytes) | ForEach-Object { $_.ToString("X2") })
    $lines.Add("$sbomHash  sbom.cdx.json")
} finally {
    $sha256.Dispose()
}
# 改行は CRLF に固定する（実行する環境に依らない）
$sumsBytes = (New-Object System.Text.UTF8Encoding($true)).GetPreamble() + (New-Object System.Text.UTF8Encoding($false)).GetBytes(($lines -join "`r`n") + "`r`n")
[System.IO.File]::WriteAllBytes($sumsPath, [byte[]]$sumsBytes)

Write-Host "カタログ: $catalogPath（$(@($result.CatalogItems.Keys).Count) ファイル / Status: $($result.Status)）"
Write-Host "ハッシュ一覧: $sumsPath"
