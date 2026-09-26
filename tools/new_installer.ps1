# インストーラー（tebunko-setup-<版>.exe）を作る（GitHub Actions の .github\workflows\release.yml が使う。手元でも実行できる）。
#
#   .\tools\new_installer.ps1 -Version v1.0.0                       work\release\ に tebunko-setup-v1.0.0.exe を作る
#   .\tools\new_installer.ps1 -Version v1.0.0 -Iscc <ISCC.exe のパス>
#
# 手順:
#   1. 起動口 tebunko.exe（installer\tebunko.cs）を、Windows 標準の .NET Framework の csc.exe でビルドする
#   2. tebunko.exe・scripts\・LICENSE・VERSION.txt（tools\new_version_text.ps1）を work\release\installer\stage\ に並べる
#   3. Inno Setup 7 のコンパイラ（ISCC.exe）で installer\tebunko.iss をビルドする
#
# Inno Setup は -Iscc で指定するか、既定の場所（Program Files (x86)・%LOCALAPPDATA%\Programs の Inno Setup 7）に入れておく。
param (
    [Parameter(Mandatory = $true)]
    [string]$Version,
    [string]$OutDir,
    [string]$Iscc
)

$ErrorActionPreference = "Stop"

$rootDir = Split-Path $PSScriptRoot -Parent
if (!$OutDir) {
    $OutDir = Join-Path $rootDir "work\release"
}
if ($Version -notmatch '^[A-Za-z0-9._-]+$') {
    throw "バージョンに使えない文字が含まれています: $Version"
}

# VERSION.txt の中身を先に作る（git の SHA が取れないときは、ここで止めてビルドを始めない）
$versionBytes = & (Join-Path $PSScriptRoot "new_version_text.ps1") -Version $Version

function getNumericVersion {
    # exe の版の情報に書く数字だけの版（v1.2.3 → 1.2.3.0）。数字の版でないとき（手元での試し）は 0.0.0.0
    param (
        [string]$version
    )

    if ($version -match '^v?(\d+)\.(\d+)\.(\d+)$') {
        return "$($Matches[1]).$($Matches[2]).$($Matches[3]).0"
    }
    return "0.0.0.0"
}

function findIscc {
    param (
        [string]$path
    )

    if ($path) {
        if (!(Test-Path -LiteralPath $path)) {
            throw "ISCC.exe が見つかりません: $path"
        }
        return $path
    }
    $candidates = @(
        (Join-Path ${env:ProgramFiles(x86)} "Inno Setup 7\ISCC.exe"),
        (Join-Path $env:ProgramFiles "Inno Setup 7\ISCC.exe"),
        (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 7\ISCC.exe")
    )
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }
    throw "Inno Setup 7 の ISCC.exe が見つかりません。Inno Setup 7 を入れるか、-Iscc で指定してください。"
}

$appVersion = $Version -replace '^v', ''
$numericVersion = getNumericVersion $Version
$isccPath = findIscc $Iscc

$buildDir = Join-Path $OutDir "installer"
$stageDir = Join-Path $buildDir "stage"
if (Test-Path -LiteralPath $buildDir) {
    Remove-Item -LiteralPath $buildDir -Recurse -Force
}
[System.IO.Directory]::CreateDirectory($stageDir) | Out-Null

# 1. tebunko.exe をビルドする。版の情報は別のファイルにして足す（tebunko.cs には書かない）
$csc = Join-Path ([System.Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()) "csc.exe"
$versionFile = Join-Path $buildDir "version.cs"
$versionSource = @(
    "[assembly: System.Reflection.AssemblyTitle(`"tebunko`")]"
    "[assembly: System.Reflection.AssemblyProduct(`"tebunko`")]"
    "[assembly: System.Reflection.AssemblyVersion(`"$numericVersion`")]"
    "[assembly: System.Reflection.AssemblyFileVersion(`"$numericVersion`")]"
    "[assembly: System.Reflection.AssemblyInformationalVersion(`"$appVersion`")]"
) -join "`r`n"
[System.IO.File]::WriteAllText($versionFile, $versionSource, (New-Object System.Text.UTF8Encoding($true)))
$icon = Join-Path $rootDir "scripts\tebunko\tebunko.ico"
$exe = Join-Path $stageDir "tebunko.exe"
& $csc -nologo -codepage:65001 -target:winexe -optimize+ -warnaserror+ "-win32icon:$icon" "-out:$exe" `
    (Join-Path $rootDir "installer\tebunko.cs") $versionFile
if ($LASTEXITCODE -ne 0) {
    throw "tebunko.exe のビルドに失敗しました（csc.exe の終了コード $LASTEXITCODE）"
}

# 2. 入れるファイルを並べる（zip と同じく work\・setting.config は入れない）
Copy-Item -LiteralPath (Join-Path $rootDir "scripts") -Destination $stageDir -Recurse
Copy-Item -LiteralPath (Join-Path $rootDir "LICENSE") -Destination $stageDir
[System.IO.File]::WriteAllBytes((Join-Path $stageDir "VERSION.txt"), $versionBytes)

# 3. インストーラーをビルドする
& $isccPath /Q "/O$OutDir" "/DAppVersion=$appVersion" "/DSetupVersion=$Version" "/DNumericVersion=$numericVersion" "/DStageDir=$stageDir" "/DIconFile=$icon" `
    (Join-Path $rootDir "installer\tebunko.iss")
if ($LASTEXITCODE -ne 0) {
    throw "インストーラーのビルドに失敗しました（ISCC.exe の終了コード $LASTEXITCODE）"
}

$setupPath = Join-Path $OutDir "tebunko-setup-$Version.exe"
$hash = (Get-FileHash -LiteralPath $setupPath -Algorithm SHA256).Hash
Write-Host "インストーラー: $setupPath"
Write-Host "SHA256: $hash"
