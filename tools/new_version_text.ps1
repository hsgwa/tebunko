# 配布物に入れる VERSION.txt の中身（タグ名とコミットの SHA の2行）を作る。
# tools\new_release_package.ps1（zip）と tools\new_installer.ps1（インストーラーのステージ）が共通で呼ぶ。
# 作業ツリー（git のチェックアウト）には書かず、呼び出し側が zip のエントリー・ステージのファイルへ直接書き込む。
#
#   $bytes = & .\tools\new_version_text.ps1 -Version v1.0.0
#
# 中身は BOM 付き UTF-8・CRLF の2行（1行目: タグ名 / 2行目: コミットの SHA）。
# SHA は `git rev-parse HEAD` で取る。git が使えない・リポジトリの外など、SHA を取得できないときは throw して止める
# （黙って空で入れると、読む側（scripts\shared\core\version.ps1 の readVersionFile）で「形が違う → 開発版」になり、
#   配布物なのに開発版と表示されてしまうため）。
param (
    [Parameter(Mandatory = $true)]
    [string]$Version
)

$ErrorActionPreference = "Stop"

if ($Version -notmatch '^[A-Za-z0-9._-]+$') {
    throw "バージョンに使えない文字が含まれています: $Version"
}

$rootDir = Split-Path $PSScriptRoot -Parent
$sha = & git -C $rootDir rev-parse HEAD 2>$null
if ($LASTEXITCODE -ne 0 -or $sha -notmatch '^[0-9a-f]{40}$') {
    throw "コミットの SHA を取得できませんでした（git rev-parse HEAD）。git のリポジトリで実行してください。"
}

$content = "$Version`r`n$sha`r`n"
(New-Object System.Text.UTF8Encoding($true)).GetPreamble() + (New-Object System.Text.UTF8Encoding($false)).GetBytes($content)
