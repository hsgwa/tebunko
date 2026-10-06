# カバレッジの計測の対象にするファイルの一覧（tests/run.ps1 と tests/meta/structure.Tests.ps1 が使う）。
# 対象は判断層・状態層だけ。画面層は自動テストの対象外にする。外すのは名前で決める画面層だけで、
# ui\ の下でも、テストのあるもの（result_list・preview・index_tree・open_source・types・startup_error など）は対象に入れる。
#   ・起動口と画面の枠: gui.ps1・gui_main.ps1・shell.ps1・app_host.ps1・splash.ps1・nav.ps1（ui\shell\ のナビと画面の切り替え）
#   ・タブ: *_tab.ps1
#   ・ダイアログ: *_dialog.ps1
# 画面層のファイルを足したら、上のどれかの名前にするか、ここに足して、分母から外れていることを確かめる。

function getCoverageTargets {
    param (
        [string]$scriptsRoot
    )

    return @(Get-ChildItem -LiteralPath $scriptsRoot -Recurse -Filter "*.ps1" |
        Where-Object {
            $_.Name -notmatch "^(gui|gui_main|shell|app_host|splash|nav)\.ps1$" -and
            $_.Name -notmatch "_tab\.ps1$" -and
            $_.Name -notmatch "_dialog\.ps1$"
        } |
        ForEach-Object { $_.FullName })
}
