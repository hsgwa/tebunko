# 配布物の部品表（CycloneDX 1.6）を作る（docs\safety\scans.md「配布物の完全性」・supply-chain.md）。
# tools\new_release_package.ps1 が、zip に入れるバイト列をそろえた後に呼ぶ。作業ツリーには書かず、部品表のバイト列を返す。
#
#   $bytes = & .\tools\new_sbom.ps1 -Version v1.0.0 -Sha <コミットの SHA> -Timestamp $time -Entries $entries
#
# 雛形（リポジトリ直下の sbom.cdx.json。本体の説明・ライセンス・前提ソフトウェア・注記だけを持つ）に、次を足す。
#   serialNumber        版とコミットの SHA から決まる UUID（同じコミット・同じ版なら同じ値。CycloneDX 1.6 の形に合わせる）
#   metadata.timestamp  コミットの時刻（UTC）
#   metadata.component.version  版
#   components          zip に入る全ファイル（zip 内のパス・SHA-256）と、同梱する第三者の部品（フォント・アイコンの形。下の $thirdParty）
#   dependencies        本体 → 全ファイル・第三者の部品
# -Entries は zip 内のパス（tebunko\ を除き / 区切り）→ zip に入れるバイト列（順序付き）。並びはそのまま components の順になる。
# 返すのは BOM 無し UTF-8 の JSON。字下げ・エスケープは Windows PowerShell 5.1 の ConvertTo-Json のもので、
# バイト列が一致するのは 5.1 で作ったとき（既定の -Depth 2 では深いところが切れるため 10 を指定する）。
param (
    [Parameter(Mandatory = $true)]
    [string]$Version,
    [Parameter(Mandatory = $true)]
    [string]$Sha,
    [Parameter(Mandatory = $true)]
    [DateTimeOffset]$Timestamp,
    [Parameter(Mandatory = $true)]
    [System.Collections.Specialized.IOrderedDictionary]$Entries
)

$ErrorActionPreference = "Stop"

# 同梱する第三者の部品（実行されるコードではなく、purl（パッケージ識別子）は持たない）。
# 版・ライセンスを変えたら、ここと scripts/shared/fonts/ のライセンス文・docs/safety/supply-chain.md をそろえる
$thirdParty = @(
    [ordered]@{
        "type"        = "data"
        "bom-ref"     = "third-party/rethink-sans"
        "group"       = "Rethink Sans Project"
        "name"        = "Rethink Sans"
        "version"     = "20d5980cd14ce827e82d7fc58d758f7cc5086c91"   # 上流にタグが無いため、同梱の版と 1 バイトも違わない上流のコミット（2023-10-11）
        "description" = "画面のフォント（scripts/shared/fonts/RethinkSans-wght.ttf・RethinkSans-Italic-wght.ttf）"
        "scope"       = "required"
        "licenses"    = @([ordered]@{ "license" = [ordered]@{ "id" = "OFL-1.1" } })
        "externalReferences" = @([ordered]@{ "type" = "website"; "url" = "https://github.com/hans-thiessen/Rethink-Sans" })
    }
    [ordered]@{
        "type"        = "data"
        "bom-ref"     = "third-party/lucide"
        "group"       = "Lucide Icons and Contributors"
        "name"        = "Lucide"
        "version"     = "1.50.0"
        "description" = "アイコンの形（scripts/shared/xaml/theme.xaml の Icon.* に図形のデータとして写している）"
        "scope"       = "required"
        "licenses"    = @([ordered]@{ "license" = [ordered]@{ "id" = "ISC" } })
        "externalReferences" = @([ordered]@{ "type" = "website"; "url" = "https://github.com/lucide-icons/lucide" })
    }
)

$rootDir = Split-Path $PSScriptRoot -Parent
$template = [System.IO.File]::ReadAllText((Join-Path $rootDir "sbom.cdx.json"), [System.Text.Encoding]::UTF8) | ConvertFrom-Json

# serialNumber: 版と SHA の SHA-256 の先頭 16 バイトを UUID にする。
# スキーマの形（版の桁が 1〜5・variant が 8〜b）に合わせ、7 バイト目の上位 4 ビットを 5、9 バイト目の上位 2 ビットを 10 にする
$sha256 = [System.Security.Cryptography.SHA256]::Create()
try {
    $digest = $sha256.ComputeHash([System.Text.Encoding]::UTF8.GetBytes("${Version}:${Sha}"))
    $uuid = [byte[]]$digest[0..15]
    $uuid[6] = ($uuid[6] -band 0x0F) -bor 0x50
    $uuid[8] = ($uuid[8] -band 0x3F) -bor 0x80
    $hex = -join ($uuid | ForEach-Object { $_.ToString("x2") })
    $serialNumber = "urn:uuid:" + $hex.Substring(0, 8) + "-" + $hex.Substring(8, 4) + "-" + $hex.Substring(12, 4) + "-" + $hex.Substring(16, 4) + "-" + $hex.Substring(20, 12)

    $components = New-Object System.Collections.Generic.List[object]
    $refs = New-Object System.Collections.Generic.List[string]
    foreach ($path in $Entries.Keys) {
        $fileHash = -join ($sha256.ComputeHash([byte[]]$Entries[$path]) | ForEach-Object { $_.ToString("x2") })
        $components.Add([ordered]@{
            "type"    = "file"
            "bom-ref" = $path
            "group"   = "tebunko"
            "name"    = $path
            "hashes"  = @([ordered]@{ "alg" = "SHA-256"; "content" = $fileHash })
        })
        $refs.Add($path)
    }
    foreach ($part in $thirdParty) {
        $components.Add($part)
        $refs.Add($part["bom-ref"])
    }
} finally {
    $sha256.Dispose()
}

# 雛形の本体の説明に版を足す（name の次に置く）
$component = [ordered]@{}
foreach ($property in $template.metadata.component.PSObject.Properties) {
    $component[$property.Name] = $property.Value
    if ($property.Name -eq "name") {
        $component["version"] = $Version
    }
}
$metadata = [ordered]@{
    "timestamp" = $Timestamp.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ", [System.Globalization.CultureInfo]::InvariantCulture)
    "component" = $component
}
foreach ($property in $template.metadata.PSObject.Properties) {
    if ($property.Name -ne "component") {
        $metadata[$property.Name] = $property.Value
    }
}

$sbom = [ordered]@{
    "bomFormat"    = $template.bomFormat
    "specVersion"  = $template.specVersion
    "serialNumber" = $serialNumber
    "version"      = $template.version
    "metadata"     = $metadata
    "components"   = $components.ToArray()
    "dependencies" = @([ordered]@{ "ref" = $component["bom-ref"]; "dependsOn" = $refs.ToArray() })
}

$json = ConvertTo-Json -InputObject $sbom -Depth 10
, (New-Object System.Text.UTF8Encoding($false)).GetBytes($json + "`r`n")
