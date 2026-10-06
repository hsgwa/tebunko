# 画面（WPF）の起動口。
#
# ［1 インデックス管理］［2 検索］［8 設定］［9 プロセス停止］の4タブ。画面の定義は xaml\tebunko.xaml。
# インデックス作成は indexer.ps1 をウィンドウを出さずに起動して進み具合を表示し、検索・プロセス停止は画面内で行う。
#
# このファイルは起動口。画面の中身は ui\gui_main.ps1（startGui）と、そこから読み込む ui\ 配下・..\shared\ui\ 配下に分けてある。
# 起動そのものに失敗したとき（Add-Type・読み込み・画面の組み立てで例外）は、ui\startup_error.ps1 の reportStartupFailure が知らせる。

. "$PSScriptRoot\ui\startup_error_view.ps1"
. "$PSScriptRoot\ui\startup_error.ps1"

try {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

    # ---- 起動中の表示 ----
    # スクリプトの読み込みに数秒かかるため、先に小さなウィンドウ（xaml\splash.xaml）を出して、起動していることを知らせる。
    # 画面（$window）を描き終わったら閉じる（ContentRendered）。多重起動の判定より前に出すため、2 つ目の起動でも一瞬出る
    . "$PSScriptRoot\ui\splash.ps1"
    $script:splash = showSplash "$PSScriptRoot\xaml\splash.xaml" "$PSScriptRoot\tebunko.ico"
    stepSplash 5

    # zip 展開で付く Mark-of-the-Web（外部由来の印）を、scripts 配下から消す。印が残っていると
    # RemoteSigned でスクリプトの読み込みがブロックされるため。通常は tebunko.bat が起動前に消すが、
    # ショートカットから直接起動したときや、あとでファイルを差し替えたときのために、ここでも消しておく。
    # （この gui.ps1 自身が印付きだと、この行に来る前にブロックされる。その場合は tebunko.bat から起動する）
    # 単一 .ps1 版（${bundledScriptPath} あり）ではしない。動き始めた後に自分の印を消しても起動には効かず、
    # EDR からは防御の回避に見えるため
    if ($null -eq (Get-Variable -Name bundledScriptPath -ErrorAction SilentlyContinue)) {
        try {
            Get-ChildItem -LiteralPath (Split-Path $PSScriptRoot -Parent) -Recurse -File -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue
        } catch { }
    }

    . "$PSScriptRoot\..\shared\shared.ps1"
    . "$PSScriptRoot\core\settings.ps1"
    . "$PSScriptRoot\lib.ps1"
    stepSplash 40

    . "$PSScriptRoot\..\shared\ui\types.ps1"
    . "$PSScriptRoot\ui\types.ps1"
    . "$PSScriptRoot\..\shared\ui\app_host.ps1"
    stepSplash 50

    . "$PSScriptRoot\ui\gui_main.ps1"
    startGui -TebunkoDir $PSScriptRoot
} catch {
    exit (reportStartupFailure $_)
}
