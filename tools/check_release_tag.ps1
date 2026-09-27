# リリースのタグを公開してよいかを確かめる（.github\workflows\release.yml の最初の guard ジョブが実行する）。
#
#   .\tools\check_release_tag.ps1 -Tag v0.2.0 -Sha <タグが指すコミット> -Base origin/main
#
# 次のどちらかに当たれば理由を出して終了コード 1 で終わる。
#   - タグの形が「v<メジャー>.<マイナー>.<パッチ>」（数字は 0 か 0 始まりでない半角の数字）でない
#   - タグが指すコミットが、-Base（既定は origin/main）の履歴に無い
# 形の判定は大文字と小文字を区別する（大文字の V・全角の数字・0 始まり・v1.0.0-rc1 のような接尾辞を通さない）。
# Windows PowerShell 5.1 と PowerShell 7（CI の ubuntu）の両方で動かす。リポジトリの中（-Base を引ける場所）で実行する。
param (
    [Parameter(Mandatory = $true)]
    [AllowEmptyString()]
    [string]$Tag,
    [Parameter(Mandatory = $true)]
    [AllowEmptyString()]
    [string]$Sha,
    [string]$Base = "origin/main"
)

$ErrorActionPreference = "Stop"

if ($Tag -cnotmatch '^v(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\z') {
    Write-Host "タグ '$Tag' は v<メジャー>.<マイナー>.<パッチ> の形ではありません（例 v0.2.0。大文字の V・全角の数字・0 始まり・接尾辞は使えません）。リリースを公開しません。"
    exit 1
}

# 0 が main の履歴にある、1 が無い、128 が参照を解決できないなどの失敗
$ErrorActionPreference = "Continue"
& git merge-base --is-ancestor $Sha $Base 2>$null
$code = $LASTEXITCODE
if ($code -eq 0) { exit 0 }
if ($code -eq 1) {
    Write-Host "タグ '$Tag' が指すコミット $Sha は $Base の履歴にありません。main に入ったコミットに付けたタグからだけリリースします。"
    exit 1
}
Write-Host "コミット '$Sha' または '$Base' を解決できませんでした（git の終了コード $code）。履歴を全部取得しているか（fetch-depth: 0）を確かめてください。"
exit 1
