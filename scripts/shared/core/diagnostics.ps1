# 診断情報の収集と、ログの世代の入れ替え（状態層）。
# レジストリ・WMI・COM は使わず、Office も起動しない。ファイルの版情報と .NET・PowerShell の値だけを読む。
# 見出しの文面は diagnostics_view.ps1。

function getEnvironmentInfo {
    # OS・PowerShell・Office の情報を返す。調べられない項目は $null
    $info = @{ OsVersion = $null; OsFileVersion = $null; Os64 = $null; PsVersion = $null; Ps64 = $null; LanguageMode = $null; UICulture = $null; Office = @() }
    try { $info.OsVersion = [Environment]::OSVersion.Version.ToString() } catch {}
    try { $info.Os64 = [Environment]::Is64BitOperatingSystem } catch {}
    try { $info.Ps64 = [Environment]::Is64BitProcess } catch {}
    try { $info.PsVersion = $PSVersionTable.PSVersion.ToString() } catch {}
    try { $info.LanguageMode = [string]$ExecutionContext.SessionState.LanguageMode } catch {}
    try { $info.UICulture = [System.Globalization.CultureInfo]::CurrentUICulture.Name } catch {}
    try {
        $kernel = [System.Diagnostics.FileVersionInfo]::GetVersionInfo("$env:SystemRoot\System32\ntoskrnl.exe")
        $info.OsFileVersion = "$($kernel.FileMajorPart).$($kernel.FileMinorPart).$($kernel.FileBuildPart).$($kernel.FilePrivatePart)"
    } catch {}
    try { $info.Office = @(findOfficeApps) } catch {}
    return $info
}

function findOfficeApps {
    # 決まった場所にある Office の実行ファイル（EXCEL.EXE・WINWORD.EXE・POWERPNT.EXE）を探し、@{ Name; Version; Kind; Bits } の配列を返す。
    # Program Files は、Windows のあるドライブ（$env:SystemRoot のドライブ）の決まった名前のフォルダとする
    # （環境変数 ProgramFiles は 32 ビットの PowerShell で x86 側を指し、安全性の検査も SystemRoot 以外の環境変数を許さないため）
    $systemDrive = [System.IO.Path]::GetPathRoot($env:SystemRoot)
    if (-not $systemDrive) {
        return @()
    }
    $programFiles64 = Join-Path $systemDrive "Program Files"
    $programFilesX86 = Join-Path $systemDrive "Program Files (x86)"
    $subDirs = @("Microsoft Office\root\Office16", "Microsoft Office\Office16", "Microsoft Office\Office15")
    $apps = @(@("Excel", "EXCEL.EXE"), @("Word", "WINWORD.EXE"), @("PowerPoint", "POWERPNT.EXE"))
    $found = New-Object System.Collections.Generic.List[object]
    foreach ($app in $apps) {
        $hit = $null
        foreach ($base in @($programFiles64, $programFilesX86)) {
            if (-not $base -or $hit) { continue }
            foreach ($sub in $subDirs) {
                $candidate = Join-Path (Join-Path $base $sub) $app[1]
                if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                    $hit = $candidate
                    break
                }
            }
        }
        if ($hit) {
            $kind = getOfficeKind $hit $programFiles64 $programFilesX86
            $version = $null
            try { $version = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($hit).FileVersion } catch {}
            $found.Add(@{ Name = $app[0]; Version = $version; Kind = $kind.Kind; Bits = $kind.Bits })
        }
    }
    return $found.ToArray()
}

function getPreviousLogPath {
    # ログの前回の分の名前（<名前>_previous<拡張子>）。同じフォルダに置く
    param (
        [string]$path
    )

    $dir = [System.IO.Path]::GetDirectoryName($path)
    $name = [System.IO.Path]::GetFileNameWithoutExtension($path) + "_previous" + [System.IO.Path]::GetExtension($path)
    return (Join-Path $dir $name)
}

function rotateLogFile {
    # path が maxBytes 以上なら、前回の分（getPreviousLogPath。あれば上書き）に移す。0 はいつも移す。
    # 移したら $true。元が無い・上限未満・移せない（読み取り専用・開いたまま）は $false（例外は出さない）
    param (
        [string]$path,
        [long]$maxBytes
    )

    try {
        if (-not [System.IO.File]::Exists($path)) {
            return $false
        }
        if ($maxBytes -gt 0 -and (New-Object System.IO.FileInfo($path)).Length -lt $maxBytes) {
            return $false
        }
        $previous = getPreviousLogPath $path
        if ([System.IO.File]::Exists($previous)) {
            [System.IO.File]::Delete($previous)
        }
        [System.IO.File]::Move($path, $previous)
        return $true
    } catch {
        return $false
    }
}
