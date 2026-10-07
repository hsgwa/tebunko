# ステータスバーの更新中の 1 行（画面層）。gui_main.ps1 が読み込む。
# 更新中だけ、左に「更新中: <インデックス名>　205 / 455 件（失敗 1 件）・残り約 3 分」を出し、いつもの文（StatusText）は隠す。
# 件数は更新全体の進み（getIndexingProgress）、残り時間は［インデックス管理］の進み表示と同じ文言を読む。
# 文言は判断層（getIndexingStatusLine）。

function updateIndexingStatusLine {
    if (!(isIndexing)) {
        $ui.IndexingStatusText.Visibility = "Collapsed"
        $ui.StatusText.Visibility = "Visible"
        return
    }

    $progress = getIndexingProgress
    # 更新しているインデックスの名前は、更新中の最初の行から取る（取れなければ名前なし）
    $name = ""
    foreach ($item in $script:targetItems) {
        if ($item.IndexLevel -eq "Run") {
            $name = [string]$item.Name
            break
        }
    }
    $text = getIndexingStatusLine $name $progress.Processed ($progress.Processed + $progress.Remaining) $progress.Failed ([string]$ui.IndexingProgressEta.Text)
    $ui.IndexingStatusText.Text = $text
    $ui.IndexingStatusText.ToolTip = $text
    $ui.IndexingStatusText.Visibility = "Visible"
    $ui.StatusText.Visibility = "Hidden"
}

$script:statusLineTimer = newTimer 1000 { safe { updateIndexingStatusLine } }
$script:statusLineTimer.Start()
