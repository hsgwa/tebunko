# 公開する前に、配布 zip の中身を検査する（.github\workflows\release.yml が、配布物を作った後・来歴に署名する前に実行する）。
#
#   .\tools\check_release_package.ps1 -ZipPath work\release\tebunko-v1.0.0.zip -OutDir work\release -Version v1.0.0
#
# -OutDir は zip と並べる tebunko.cat・SHA256SUMS.txt・sbom.cdx.json を置いたフォルダ。リポジトリの中（git ls-files を引ける場所）で、
# Windows PowerShell 5.1 で実行する（Test-FileCatalog が要る）。zip は一時フォルダに展開して確かめ、終わったら消す。
# 通らなかった項目はすべて列挙してから、終了コード 1 で終わる。
#
# 確かめること:
#   1. zip のエントリーが、git ls-files scripts と固定のファイル（tebunko.bat・README.md・LICENSE・VERSION.txt）に過不足なく一致する
#   2. 読み込み口（scripts\tebunko\gui.ps1・indexer.ps1）から dot-source でたどれる先がすべて存在する
#   3. .ps1 が構文エラーなく解析でき、.xaml が XML として読める
#   4. Test-FileCatalog -Path .\scripts, .\tebunko.bat が Valid になる
#   5. SHA256SUMS.txt・sbom.cdx.json のハッシュが、展開したファイルと一致する
#   6. VERSION.txt が、タグ名と HEAD のコミットの SHA になっている
param (
    [Parameter(Mandatory = $true)]
    [string]$ZipPath,
    [Parameter(Mandatory = $true)]
    [string]$OutDir,
    [Parameter(Mandatory = $true)]
    [string]$Version,
    # 読み込み口（zip の tebunko\ から見たパス）
    [string[]]$EntryScripts = @("scripts/tebunko/gui.ps1", "scripts/tebunko/indexer.ps1")
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$rootDir = Split-Path $PSScriptRoot -Parent
$failures = New-Object System.Collections.Generic.List[string]

# ファイルが dot-source している先を返す（. "$PSScriptRoot\..." の形と、ui\ 配下のファイルが使う
# . "$TebunkoDir\..." の形（起動口 gui.ps1 から渡される tebunko\ 直下）を見る）。存在しない先も返す
# （tests\meta\layers.Tests.ps1 の getSourcedFiles は、存在するものだけを返し、tools から tests を読み込まないため、ここに持つ）
function Get-DotSourceTargets([string]$Path, [string]$TebunkoDir) {
    $dir = Split-Path $Path -Parent
    $text = [System.IO.File]::ReadAllText($Path)
    foreach ($match in [regex]::Matches($text, '(?m)^\s*\.\s+"\$(PSScriptRoot|TebunkoDir)\\([^"]+)"')) {
        $base = if ($match.Groups[1].Value -eq "TebunkoDir") { $TebunkoDir } else { $dir }
        [System.IO.Path]::GetFullPath((Join-Path $base $match.Groups[2].Value))
    }
}

function Get-Sha256Hex([byte[]]$Bytes) {
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        -join ($sha256.ComputeHash($Bytes) | ForEach-Object { $_.ToString("X2") })
    } finally {
        $sha256.Dispose()
    }
}

$extractDir = Join-Path ([System.IO.Path]::GetTempPath()) ("tebunko-check-" + [Guid]::NewGuid().ToString("N"))
try {
    if (!(Test-Path -LiteralPath $ZipPath)) {
        throw "zip がありません: $ZipPath"
    }

    # --- 1. エントリーの一覧 ---
    $zipFile = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path -LiteralPath $ZipPath).Path)
    try {
        $actual = @($zipFile.Entries | ForEach-Object { $_.FullName })
    } finally {
        $zipFile.Dispose()
    }
    $tracked = @(& git -C $rootDir -c core.quotepath=false ls-files scripts)
    if ($LASTEXITCODE -ne 0 -or $tracked.Count -eq 0) {
        throw "scripts のファイルの一覧を取得できませんでした（git ls-files scripts）。"
    }
    $expected = @(@($tracked) + @("tebunko.bat", "README.md", "LICENSE", "VERSION.txt") | ForEach-Object { "tebunko/$_" })
    foreach ($name in $expected) {
        if ($actual -cnotcontains $name) { $failures.Add("zip に無いファイル: $name") }
    }
    foreach ($name in $actual) {
        if ($expected -cnotcontains $name) { $failures.Add("zip に余分なファイル: $name") }
    }

    [System.IO.Compression.ZipFile]::ExtractToDirectory((Resolve-Path -LiteralPath $ZipPath).Path, $extractDir)
    $pkgDir = Join-Path $extractDir "tebunko"
    $tebunkoDir = Join-Path $pkgDir "scripts\tebunko"

    # --- 2. dot-source の先 ---
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    $queue = New-Object System.Collections.Queue
    foreach ($entry in $EntryScripts) {
        $full = [System.IO.Path]::GetFullPath((Join-Path $pkgDir $entry.Replace("/", "\")))
        if (!(Test-Path -LiteralPath $full)) {
            $failures.Add("読み込み口がありません: $entry")
        } elseif ($seen.Add($full)) {
            $queue.Enqueue($full)
        }
    }
    while ($queue.Count -gt 0) {
        $file = $queue.Dequeue()
        foreach ($target in @(Get-DotSourceTargets $file $tebunkoDir)) {
            if (!(Test-Path -LiteralPath $target)) {
                $failures.Add("dot-source の先がありません: $($file.Substring($pkgDir.Length + 1)) -> $($target.Substring($pkgDir.Length + 1))")
            } elseif ($seen.Add($target)) {
                $queue.Enqueue($target)
            }
        }
    }

    # --- 3. 構文 ---
    foreach ($file in @(Get-ChildItem -LiteralPath (Join-Path $pkgDir "scripts") -Recurse -File)) {
        $relative = $file.FullName.Substring($pkgDir.Length + 1)
        if ($file.Extension -ieq ".ps1") {
            $errors = $null
            [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$errors) | Out-Null
            # 別のファイルで定義したクラスを継承・参照している箇所は、1 つのファイルだけでは型を解決できない（TypeNotFound）ため、構文エラーに数えない
            foreach ($e in @($errors | Where-Object { $_.ErrorId -ne "TypeNotFound" })) {
                $failures.Add("構文エラー: ${relative}:$($e.Extent.StartLineNumber) $($e.Message)")
            }
        } elseif ($file.Extension -ieq ".xaml") {
            try {
                $doc = New-Object System.Xml.XmlDocument
                $doc.XmlResolver = $null
                $doc.Load($file.FullName)
            } catch {
                $failures.Add("XML として読めません: $relative ($($_.Exception.Message))")
            }
        }
    }

    # --- 4. カタログ ---
    $catalogPath = Join-Path $OutDir "tebunko.cat"
    if (!(Test-Path -LiteralPath $catalogPath)) {
        $failures.Add("tebunko.cat がありません")
    } else {
        $result = Test-FileCatalog -Path (Join-Path $pkgDir "scripts"), (Join-Path $pkgDir "tebunko.bat") -CatalogFilePath $catalogPath -Detailed
        if ($result.Status -ne "Valid") {
            $failures.Add("カタログの検証に失敗しました（Status: $($result.Status)）")
        }
    }

    # --- 5. ハッシュ一覧と部品表 ---
    $sumsPath = Join-Path $OutDir "SHA256SUMS.txt"
    if (!(Test-Path -LiteralPath $sumsPath)) {
        $failures.Add("SHA256SUMS.txt がありません")
    } else {
        $listed = @{}
        foreach ($line in [System.IO.File]::ReadAllLines($sumsPath)) {
            if (!$line -or $line.StartsWith("#")) { continue }
            if ($line -cmatch '^([0-9A-F]{64})  (.+)$') {
                $listed[$Matches[2]] = $Matches[1]
            } else {
                $failures.Add("SHA256SUMS.txt の行の形が違います: $line")
            }
        }
        foreach ($path in @($listed.Keys)) {
            $file = if ($path -eq "sbom.cdx.json") { Join-Path $OutDir $path } else { Join-Path $pkgDir $path }
            if (!(Test-Path -LiteralPath $file)) {
                $failures.Add("SHA256SUMS.txt に載っているファイルがありません: $path")
            } elseif ((Get-Sha256Hex ([System.IO.File]::ReadAllBytes($file))) -ne $listed[$path]) {
                $failures.Add("SHA256SUMS.txt のハッシュと一致しません: $path")
            }
        }
        foreach ($name in $actual) {
            $path = $name.Substring("tebunko/".Length).Replace("/", "\")
            if (!$listed.ContainsKey($path)) { $failures.Add("SHA256SUMS.txt に載っていません: $path") }
        }
        if (!$listed.ContainsKey("sbom.cdx.json")) { $failures.Add("SHA256SUMS.txt に sbom.cdx.json が載っていません") }
    }

    $sbomPath = Join-Path $OutDir "sbom.cdx.json"
    if (!(Test-Path -LiteralPath $sbomPath)) {
        $failures.Add("sbom.cdx.json がありません")
    } else {
        $sbom = [System.IO.File]::ReadAllText($sbomPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
        $components = @($sbom.components)
        foreach ($component in $components) {
            $file = Join-Path $pkgDir $component.name.Replace("/", "\")
            if (!(Test-Path -LiteralPath $file)) {
                $failures.Add("部品表に載っているファイルがありません: $($component.name)")
            } elseif ((Get-Sha256Hex ([System.IO.File]::ReadAllBytes($file))).ToLower() -ne $component.hashes[0].content) {
                $failures.Add("部品表のハッシュと一致しません: $($component.name)")
            }
        }
        $componentNames = @($components | ForEach-Object { "tebunko/" + $_.name })
        foreach ($name in $actual) {
            if ($componentNames -cnotcontains $name) { $failures.Add("部品表に載っていません: $name") }
        }
    }

    # --- 6. VERSION.txt ---
    $sha = (& git -C $rootDir rev-parse HEAD).Trim()
    $versionPath = Join-Path $pkgDir "VERSION.txt"
    if (Test-Path -LiteralPath $versionPath) {
        $expectedBytes = [byte[]]((New-Object System.Text.UTF8Encoding($true)).GetPreamble() + (New-Object System.Text.UTF8Encoding($false)).GetBytes("$Version`r`n$sha`r`n"))
        $actualBytes = [System.IO.File]::ReadAllBytes($versionPath)
        if ([System.Convert]::ToBase64String($actualBytes) -ne [System.Convert]::ToBase64String($expectedBytes)) {
            $failures.Add("VERSION.txt がタグ名とコミットの SHA になっていません（期待: $Version と $sha）")
        }
    }
} finally {
    if (Test-Path -LiteralPath $extractDir) {
        Remove-Item -LiteralPath $extractDir -Recurse -Force
    }
}

if ($failures.Count -gt 0) {
    Write-Host "配布物の検査に通らなかった項目（$($failures.Count) 件）。公開しません。"
    foreach ($failure in $failures) { Write-Host "  - $failure" }
    exit 1
}
Write-Host "配布物の検査に通りました: $ZipPath"
exit 0
