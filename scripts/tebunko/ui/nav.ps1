# ナビ（左の欄）と画面の切り替え（画面層）。gui_main.ps1 が読み込む。
# 画面は、ナビの項目（SearchTab・IndexTab・SettingsTab・KillTab）の名前で呼ぶ。ナビのクリック・キー操作・ほかの画面からの切り替えは、
# すべて selectScreen を通す（ナビの部品から直に中身を差し替えない）。
# 画面の中身は起動のときに読み込んで $script:screenContents に持ち、選んだものを ContentHost に差す。

$script:currentScreen = $null

function getCurrentScreen {
    # いま出している画面の名前（まだ選んでいなければ $null）
    return $script:currentScreen
}

function selectScreen {
    param (
        [string]$name
    )

    if (!(isScreenName $name) -or $script:currentScreen -eq $name) {
        return
    }
    $script:currentScreen = $name
    $ui.ContentHost.Content = $script:screenContents[$name]
    if ($ui.NavList.SelectedItem -ne $ui[$name]) {
        $ui.NavList.SelectedItem = $ui[$name]
    }
    # 起動時の読み込み（loadStartupData）が済むまでは、切り替えたときの読み直しをしない
    # （起動時の画面を選んだときにも呼ばれ、同じ読み込みが重なるため）
    if (!$script:startupLoaded) {
        return
    }
    safe {
        if ($name -eq "KillTab") {
            refreshProcesses
            $script:processTimer.Start()
        } else {
            $script:processTimer.Stop()
        }
        if ($name -eq "IndexTab") {
            refreshIndexingState
        }
    }
}

$ui.NavList.Add_SelectionChanged({
    param ($sender, $e)
    # 中の項目の選択変更だけを扱う
    $item = $sender.SelectedItem
    if ($null -ne $item) {
        selectScreen $item.Name
    }
})
