# git のフック（tools\hooks\）を有効にする。clone したら 1 回だけ実行する。
#
#   .\tools\install_hooks.ps1             有効にする
#   .\tools\install_hooks.ps1 -Uninstall  やめる
#
# core.hooksPath を tools/hooks にするだけで、.git\hooks\ には何もコピーしない。
# フックの中身はリポジトリで管理されるため、更新は pull するだけで反映される。
# 相対パスは各作業ツリーの最上位から解決されるため、git worktree で作ったツリーでもそのツリーのフックが動く。
param (
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"

$rootDir = Split-Path $PSScriptRoot -Parent

if ($Uninstall) {
    git -C $rootDir config --unset core.hooksPath
    Write-Host "フックを無効にしました。"
    return
}

git -C $rootDir config core.hooksPath tools/hooks
if ($LASTEXITCODE -ne 0) {
    throw "core.hooksPath を設定できませんでした。"
}
Write-Host "フックを有効にしました（core.hooksPath = tools/hooks）。"
Write-Host "コミットのたびに、個人情報・文字コードの検査と速いテスト（Unit・Meta）、コミットメッセージの形（feat: ... など）と Signed-off-by の検査が走ります。"
Write-Host "コミットは git commit -s（Signed-off-by を付ける）で作ってください。"

$mail = git -C $rootDir config user.email
if ($mail -notmatch '@users\.noreply\.github\.com$') {
    Write-Host "コミットの作者のメールアドレスが GitHub の noreply ではありません（$mail）。次のように設定してください。" -ForegroundColor Yellow
    Write-Host "  git config user.email <ID>+<アカウント名>@users.noreply.github.com"
}
