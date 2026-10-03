# 配布物の版の記録（VERSION.txt）の読み取り（状態層）。
# 中身は BOM 付き UTF-8・CRLF の2行（1行目: タグ名 / 2行目: コミットの SHA）。
# tools/new_version_text.ps1 が作り、tools/new_release_package.ps1（zip）・tools/new_installer.ps1（インストーラー）が書き込む。
# 作業ツリー（git のチェックアウト）には無く、配布物にだけある。
#
# 単一 .ps1 版（展開せずに動く試験版）は VERSION.txt を持てないため、結合の道具（tools/new_single_script.ps1）が
# 版の文字列を ${bundledVersion}（@{ Tag; Sha }）としてこのファイルより前に埋める。あればファイルを読まずそれを返す

function readVersionFile {
    # ファイルが無い・読めない・形が違うときは $null。あれば @{ Tag; Sha }
    param (
        [string]$path
    )

    $bundled = Get-Variable -Name bundledVersion -ErrorAction SilentlyContinue
    if ($null -ne $bundled) {
        return $bundled.Value
    }

    if (!(Test-Path -LiteralPath $path)) {
        return $null
    }
    try {
        $lines = [System.IO.File]::ReadAllLines($path)
    } catch {
        return $null
    }
    if ($lines.Count -ne 2) {
        return $null
    }
    if ($lines[0] -notmatch '^[A-Za-z0-9._-]+$' -or $lines[1] -notmatch '^[0-9a-f]{40}$') {
        return $null
    }
    return @{ Tag = $lines[0]; Sha = $lines[1] }
}
