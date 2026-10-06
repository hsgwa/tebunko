# 前の版のファイル（見本・golden。tests/testdata/compat/settings/<見本の名前>/setting.config）を作るときと、
# 今の版でも同じ形を書けることを確かめるとき（tests/tebunko/core/settings_compat.Tests.ps1）の、両方から呼ぶ。
# load.ps1 に頼らず、settings.ps1 が公開する書く関数だけを呼ぶ（見本の版の settings.ps1 の上でも動かすため）。

function writeGoldenSettings {
    # 既定値ではない値を、setting.config のすべてのキーに書く
    param (
        [Parameter(Mandatory = $true)]
        [string]$path
    )

    writeTargetFolders @(
        [pscustomobject]@{ Name = "資料"; Path = "C:\tebunko_golden\source\資料"; Enabled = $true }
        [pscustomobject]@{ Name = "(株)山田商事"; Path = "C:\tebunko_golden\source\(株)山田商事"; Enabled = $false }
    ) $path

    writeIndexSources @(
        [pscustomobject]@{ Name = "共有資料"; Path = "\\golden-server\共有資料" }
    ) $path

    writeSearchExcludes @(
        [pscustomobject]@{ Path = "C:\tebunko_golden\source\資料\除外A"; Subfolders = $true }
        [pscustomobject]@{ Path = "C:\tebunko_golden\source\資料\除外B"; Subfolders = $false }
    ) $path

    writeSearchOption @{
        UseRegex        = $true
        CaseSensitive   = $true
        FileFilter      = "*.xlsx;!~$*"
        IncludeShapes   = $false
        IncludeComments = $false
    } $path

    # 検索の対象にするファイルの種類（この版より前の版には無いので、関数があるときだけ書く）
    if (Get-Command writeFileKinds -ErrorAction SilentlyContinue) {
        writeFileKinds @("excel", "text") $path
    }

    writeOpenMode "readOnly" $path
    writeWorkspaceFolder "C:\tebunko_golden\ws" $path
    updateSettings "ingestThreads" 2 $path
}
