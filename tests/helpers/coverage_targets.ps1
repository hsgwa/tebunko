# カバレッジの計測の対象にするファイルの一覧（tests/run.ps1 と tests/meta/structure.Tests.ps1 が使う）。
# 対象は判断層・状態層だけ。画面層は自動テストの対象外にする。外すのは名前で決める画面層だけで、
# ui\ の下でも、テストのあるもの（result_list・preview・index_tree・open_source・types・startup_error など）は対象に入れる。
#   ・起動口と画面の枠: gui.ps1・gui_main.ps1・shell.ps1・app_host.ps1・splash.ps1・trimmed_tooltip.ps1
#   ・タブ: *_tab.ps1
#   ・ダイアログ: *_dialog.ps1
#   ・画面ごとのフォルダ（tebunko\ui\search\）の画面層: 同じ名前のテスト（tests\tebunko\ui\search\<名前>.Tests.ps1）が無いもの。
#     名前だけで決めず、フォルダとテストの有無で決める（判断層の *_view.ps1 は、テストの有無にかかわらず必ず分母に入る）
# 画面層のファイルを足したら、上のどれかの名前にするか、ここに足して、分母から外れていることを確かめる。

function getCoverageTargets {
    param (
        [string]$scriptsRoot
    )

    $scriptsRoot = [System.IO.Path]::GetFullPath($scriptsRoot)
    $testsRoot = Join-Path (Split-Path -Parent $scriptsRoot) "tests"
    $screenFolders = @("tebunko\ui\search", "tebunko\ui\index", "tebunko\ui\settings")
    return @(Get-ChildItem -LiteralPath $scriptsRoot -Recurse -Filter "*.ps1" |
        Where-Object {
            $relative = $_.FullName.Substring($scriptsRoot.TrimEnd("\").Length + 1)
            $folder = Split-Path -Parent $relative
            if ($screenFolders -contains $folder) {
                # 判断層（*_view.ps1）は、同名のテストが無くても必ず分母に入れる（テストが無いまま黙って外れないように）
                if ($_.Name -match "_view\.ps1$") {
                    return $true
                }
                return (Test-Path -LiteralPath (Join-Path $testsRoot "$folder\$($_.BaseName).Tests.ps1"))
            }
            return $true
        } |
        Where-Object {
            $_.Name -notmatch "^(gui|gui_main|shell|app_host|splash|trimmed_tooltip)\.ps1$" -and
            $_.Name -notmatch "_tab\.ps1$" -and
            $_.Name -notmatch "_dialog\.ps1$"
        } |
        ForEach-Object { $_.FullName })
}
