# 起動口（tebunko_grep.bat から -File で起動する）。いつもの画面（gui.ps1）で起動できるかを調べる。
#
# 制限言語モード（CLM）でも読み込めて動く書き方だけで書く。AppLocker・WDAC の PC では、許可されていないスクリプトが
# CLM で動き、gui.ps1 は 1 行目（Add-Type）で止まる。bat は画面を隠して起動するため、そのままでは利用者に何も見えない。
# そこで先にこのスクリプトを見えるコンソールで動かし、画面を開けないときは理由を表示して制限モードで起動する。
#
# 終了コード（tebunko_grep.bat が見る）:
#   10 … いつもの画面で起動できる（bat が gui.ps1 を起動する）
#   11 … 制限モードで起動する旨を表示した（-CheckOnly のとき。テストと CI で使う）
#   12 … 確かめている途中でエラーが起きた（内容はここで表示した）
#    0 … 制限モードを終えた
#   それ以外 … このスクリプト自体を動かせなかった（グループポリシーの実行ポリシーなど。bat が start_failed.txt を表示する）
param (
    [switch]$CheckOnly  # 制限モードの説明を出したところで終える（制限モードの中身には入らない）
)

$ErrorActionPreference = "Stop"

trap {
    Write-Host "起動の確認の途中でエラーが起きました: $($_.Exception.Message)" -ForegroundColor Red
    exit 12
}

. "$PSScriptRoot\restricted\startup_view.ps1"
. "$PSScriptRoot\restricted\startup_check.ps1"

# このファイルは scripts\tebunko_grep\ に置く。2 つ上がリポジトリ直下になる
# （shared\core\paths.ps1 は CLM で使えない型を使うため、ここでは読み込まない）
$rootDir = (Resolve-Path -LiteralPath "$PSScriptRoot\..\..").Path
$facts = getStartupFacts (Join-Path $rootDir "work")

if ((getStartupMode $facts) -eq ${startupModeNormal}) {
    exit 10
}

foreach ($line in (getRestrictedStartupLines $facts)) {
    Write-Host $line
}

# 制限モードの部品（-CheckOnly でも読み込み、制限言語モードで読み込めることを確かめる）
. "$PSScriptRoot\restricted\lib_restricted.ps1"

if ($CheckOnly) {
    exit 11
}

Write-Host ""
Write-Host "制限モードの検索は、まだ使えません（準備中）。"
Read-Host "Enter キーで閉じます" | Out-Null
exit 0
