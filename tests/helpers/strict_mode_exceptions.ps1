# StrictMode（Version 2）で流せないテストファイルの一覧。tests\run.ps1 は全部のテストを StrictMode 2 で流す。
# ここに載せたファイルは、最上位の BeforeAll の先頭で Set-StrictMode -Off にする。増やさない（まず本番かテストを直す）。
# tests\meta\strict_mode.Tests.ps1 が、この一覧と「Set-StrictMode -Off を書いたテストファイル」の一致を確かめる。
#
# 載せるとき: @{ Path = "tests/<パス>.Tests.ps1"; Reason = "<直せない理由>" }
function getStrictModeExceptions {
    @()
}
