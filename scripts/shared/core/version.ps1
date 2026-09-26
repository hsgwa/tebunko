# 配布物の版の記録（VERSION.txt）の読み取り（状態層）。
# 中身は BOM 付き UTF-8・CRLF の2行（1行目: タグ名 / 2行目: コミットの SHA）。
# tools/new_version_text.ps1 が作り、tools/new_release_package.ps1（zip）・tools/new_installer.ps1（インストーラー）が書き込む。
# 作業ツリー（git のチェックアウト）には無く、配布物にだけある。

function readVersionFile {
    # ファイルが無い・読めない・形が違うときは $null。あれば @{ Tag; Sha }
    param (
        [string]$path
    )

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
