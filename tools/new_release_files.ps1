# 配布物の完全性を確かめるためのファイルを作る（docs\04_安全性.md 5.1）。
#
#   .\tools\new_release_files.ps1              work\release\ に書き出す
#   .\tools\new_release_files.ps1 -OutDir .\dist
#
# 出力:
#   windox.cat       カタログ（各ファイルの SHA256。Test-FileCatalog で検証する）
#   SHA256SUMS.txt     ファイルごとの SHA256（テキストで目視・比較できる形式）
#
# 配布する zip に 2 つとも同梱する。受け取った側は次のコマンドで、配布時点から
# 1 バイトも変わっていないことを自分で確認できる（証明書は要らない）。
#
#   Test-FileCatalog -Path .\scripts, .\windox_grep.bat -CatalogFilePath .\windox.cat -Detailed
#
# Status が Valid なら改ざんなし。ValidationFailed なら、どのファイルが違うかが表示される。
# コードサイニング証明書がある場合は、カタログに署名すると発行者の保証も付く。
#   Set-AuthenticodeSignature -FilePath .\windox.cat -Certificate $cert
param (
    [string]$OutDir
)

$ErrorActionPreference = "Stop"

$rootDir = Split-Path $PSScriptRoot -Parent
if (!$OutDir) {
    $OutDir = Join-Path $rootDir "work\release"
}
[System.IO.Directory]::CreateDirectory($OutDir) | Out-Null

# 配布物のうち、内容が固定されているもの（work・setting.config は利用者ごとに変わるため含めない）
$targets = @(
    (Join-Path $rootDir "scripts"),
    (Join-Path $rootDir "windox_grep.bat")
)

$catalogPath = Join-Path $OutDir "windox.cat"
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
$lines = New-Object System.Collections.Generic.List[string]
$lines.Add("# windox_grep 配布物の SHA256（$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss') 時点）")
$lines.Add("# 確認: Get-FileHash <ファイル> -Algorithm SHA256")
foreach ($file in @(Get-ChildItem -LiteralPath (Join-Path $rootDir "scripts") -Recurse -File | Sort-Object FullName)) {
    $relative = $file.FullName.Substring($rootDir.Length + 1)
    $lines.Add("$((Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash)  $relative")
}
foreach ($name in @("windox_grep.bat", "sbom.cdx.json", "LICENSE")) {
    $path = Join-Path $rootDir $name
    if (Test-Path -LiteralPath $path) {
        $lines.Add("$((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash)  $name")
    }
}
[System.IO.File]::WriteAllLines($sumsPath, $lines, (New-Object System.Text.UTF8Encoding($true)))

Write-Host "カタログ: $catalogPath（$(@($result.CatalogItems.Keys).Count) ファイル / Status: $($result.Status)）"
Write-Host "ハッシュ一覧: $sumsPath"
