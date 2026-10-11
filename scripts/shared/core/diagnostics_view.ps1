# 診断情報（ログの先頭の見出し）の組み立て（判断層）。
# ファイル・レジストリ・画面に触らない。調べる側は diagnostics.ps1 の getEnvironmentInfo。
# 利用者名・コンピューター名・ドメイン・パス・ネットワークの情報は、見出しに書かない。

function getVersionLabel {
    # readVersionFile の結果から、版・コミットの表示を決める: @{ Version; Commit（出さないときは空） }
    #   versionInfo: readVersionFile の戻り値（$null なら開発中に git から直接起動したときの表示にする）
    param (
        $versionInfo
    )

    if ($null -eq $versionInfo) {
        return @{ Version = "開発版"; Commit = "" }
    }
    return @{ Version = $versionInfo.Tag; Commit = $versionInfo.Sha.Substring(0, 7) }
}

function getOfficeKind {
    # 見つかった Office の実行ファイルのパスから、形態とビット数を決める: @{ Kind = "Click-to-Run" / "MSI"; Bits = 64 / 32 }
    #   programFiles64: 64 ビット側の Program Files / programFilesX86: x86 側の Program Files
    param (
        [string]$exePath,
        [string]$programFiles64,
        [string]$programFilesX86
    )

    $kind = "MSI"
    if ($exePath -match '(?i)\\root\\Office\d+\\') {
        $kind = "Click-to-Run"
    }
    $bits = 64
    if ($programFilesX86 -and $exePath.StartsWith($programFilesX86.TrimEnd("\") + "\", [System.StringComparison]::OrdinalIgnoreCase)) {
        $bits = 32
    }
    return @{ Kind = $kind; Bits = $bits }
}

function hideUserProfile {
    # 文字列の中の利用者のプロファイルのフォルダ（C:\Users\<利用者名>）を %USERPROFILE% に置き換える（大文字小文字は区別しない）。
    # 名前の前方だけが同じ別のフォルダ（末尾に文字が続く名前）は置き換えない
    param (
        [string]$text,
        [string]$profileDir
    )

    if (-not $text -or -not $profileDir) {
        return $text
    }
    $pattern = [regex]::Escape($profileDir.TrimEnd("\")) + '(?![\w.\-])'
    return [regex]::Replace($text, $pattern, '%USERPROFILE%', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
}

function formatDiagnosticHeader {
    # ログの先頭に書く見出しの行（配列）を返す
    #   toolName: ツールの名前 / versionInfo: readVersionFile の結果 / envInfo: getEnvironmentInfo の結果 / now: 日時（DateTimeOffset）
    param (
        [string]$toolName,
        $versionInfo,
        $envInfo,
        [System.DateTimeOffset]$now
    )

    $unknown = "不明"
    $orUnknown = { param($value) if ($value) { [string]$value } else { $unknown } }
    $label = getVersionLabel $versionInfo
    $versionText = $label.Version
    if ($label.Commit) {
        $versionText = "$($label.Version) ($($label.Commit))"
    }
    $bitsText = { param($is64) if ($null -eq $is64) { $unknown } elseif ($is64) { "64 ビット" } else { "32 ビット" } }

    $officeText = "見つからない"
    $officeItems = @($envInfo.Office)
    if ($null -eq $envInfo.Office) {
        $officeText = $unknown
    } elseif ($officeItems.Count -gt 0) {
        $officeText = (@($officeItems | ForEach-Object {
            "$($_.Name) $(& $orUnknown $_.Version) $($_.Kind) $($_.Bits) ビット"
        })) -join " / "
    }

    return @(
        "==== $toolName の診断情報 ===="
        "日時: $($now.ToString('yyyy-MM-dd HH:mm:ss zzz'))"
        "$toolName の版: $versionText"
        "OS: Windows $(& $orUnknown $envInfo.OsVersion)（ntoskrnl $(& $orUnknown $envInfo.OsFileVersion)・$(& $bitsText $envInfo.Os64)）"
        "PowerShell: $(& $orUnknown $envInfo.PsVersion)（$(& $bitsText $envInfo.Ps64)・$(& $orUnknown $envInfo.LanguageMode)）"
        "Office: $officeText"
        "表示言語: $(& $orUnknown $envInfo.UICulture)"
        "===="
    )
}
