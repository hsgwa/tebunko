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
# （shared\core\paths.ps1 は起動の判断より後に読み込む。ここでは setting.config の startupMode だけを自分で読む）
$rootDir = (Resolve-Path -LiteralPath "$PSScriptRoot\..\..").Path
$facts = getStartupFacts (Join-Path $rootDir "work") (Join-Path $rootDir "setting.config")

if ((getStartupMode $facts) -eq ${startupModeNormal}) {
    exit 10
}

# 設定（startupMode = clm）で、模擬の制限言語モードでの確認を選んでいるとき。
# 言語モードはセッションの初めにしか変えられないため、自分をもう一度起動して、そこで制限モードを動かす。
# -File では変える前に読み込まれてしまうため -Command で起動する（本物の制限言語モードでは、すでに CLM なので起動し直さない）
if ((testStartupSettingClm $facts.Setting) -and $facts.LanguageMode -eq "FullLanguage") {
    Write-Host "setting.config の startupMode が clm のため、模擬の制限言語モードで起動し直します。"
    $script = $PSCommandPath.Replace("'", "''")
    $arguments = $(if ($CheckOnly) { " -CheckOnly" } else { "" })
    # 最後の exit で、起動し直した側の終了コード（10・11・12・0）をそのまま返す（bat が見るため）
    $command = "`$ExecutionContext.SessionState.LanguageMode = 'ConstrainedLanguage'; & '${script}'${arguments}; exit `$LASTEXITCODE"
    & powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -Command $command
    exit $LASTEXITCODE
}

foreach ($line in (getRestrictedStartupLines $facts)) {
    Write-Host $line
}

# 制限モードの部品（-CheckOnly でも読み込み、制限言語モードで読み込めることを確かめる）
. "$PSScriptRoot\restricted\lib_restricted.ps1"

if ($CheckOnly) {
    exit 11
}

runRestrictedConsole $facts.HasExcel
exit 0
