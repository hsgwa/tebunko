# 画面の上の知らせの帯（画面層）。gui_main.ps1 が読み込む。
# 検索の画面の上（SearchNotice。結果を保存できなかったとき）・インデックス管理の画面の上（IndexNotice。まとめてのエクスポート・削除の失敗）・
# 窓の上（WorkspaceNotice。ワークスペースが使えないとき）の 3 つ。どれも同じ作りで、部品の名前は <接頭辞>Panel・Icon・Text・Button・Close。
# 文言と色の区分・閉じられるか・ボタンは判断層（getNoticeBannerView）が決め、ここは結果を画面に当てるだけ。

function showNoticeBanner {
    # 帯を出す（message が空なら閉じる）。kind は getNoticeBannerView の種類
    param (
        [string]$prefix,
        [string]$kind,
        [string]$message
    )

    $view = getNoticeBannerView $kind $message
    $panel = $ui["${prefix}Panel"]
    if (!$view.Show) {
        $panel.Visibility = "Collapsed"
        return
    }
    # 地・左端の線・文字・アイコンの色（warn は橙、ng は赤。更新の帯 setIndexingBanner と同じ色）
    $looks = @{
        warn = @("Warn.Soft", "Warn", "Banner.Warn.Text")
        ng = @("Danger.Soft", "Danger.Text", "Banner.Ng.Text")
    }
    $look = $looks[$view.Level]
    $line = themeBrush $look[1]
    $text = themeBrush $look[2]
    $panel.Background = themeBrush $look[0]
    $panel.BorderBrush = $line
    $ui["${prefix}Icon"].Stroke = $line
    $ui["${prefix}Text"].Foreground = $text
    $ui["${prefix}Text"].Text = $view.Text
    $ui["${prefix}Text"].ToolTip = $view.ToolTip
    $ui["${prefix}Close"].Foreground = $text
    $ui["${prefix}Close"].Visibility = if ($view.Closable) { "Visible" } else { "Collapsed" }
    $button = $ui["${prefix}Button"]
    $button.Content = $view.ActionText
    $button.BorderBrush = $line
    $button.Foreground = $line
    $button.Visibility = if ($view.ActionText) { "Visible" } else { "Collapsed" }
    $panel.Visibility = "Visible"
}

function hideNoticeBanner {
    param ([string]$prefix)

    $ui["${prefix}Panel"].Visibility = "Collapsed"
}

function updateWorkspaceBanner {
    # ワークスペースが使えないとき（既定の場所にほかのファイルがある）は窓の上に赤い帯を出し、使えるようになれば消す。
    # 起動時と、ワークスペースを切り替えたときに呼ぶ
    showNoticeBanner "WorkspaceNotice" "WorkspaceBlock" ([string]$script:workspaceBlock)
}

$ui.SearchNoticeClose.Add_Click({ safe { hideNoticeBanner "SearchNotice" } })
$ui.IndexNoticeClose.Add_Click({ safe { hideNoticeBanner "IndexNotice" } })
$ui.WorkspaceNoticeButton.Add_Click({ safe { selectScreen "SettingsTab" } })
