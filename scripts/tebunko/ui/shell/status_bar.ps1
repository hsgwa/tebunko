# ステータスバーの 1 行（画面層）。gui_main.ps1 が読み込む。
# 更新中だけ、左に「更新中: <インデックス名>　205 / 455 件（失敗 1 件）・残り約 3 分」を出し、いつもの文（StatusText）は隠す。
# 件数は更新全体の進み（getIndexingProgress）、残り時間は［インデックス管理］の進み表示と同じ値（$script:indexingView）を読む。
# 何をいつ出すかは判断層（getStatusBarView・getIndexingStatusLine）。文が長くて切れたときは、切れているときだけツールチップで全文を見せる。

addTrimmedToolTip $ui.StatusText
addTrimmedToolTip $ui.IndexingStatusText

function updateIndexingStatusLine {
    $line = ""
    if (isIndexing) {
        # 名前・件数・残り時間は、更新の進みの更新が値で持つ $script:indexingView から読む（最初の進みが来るまでは名前・残り時間なし）
        $view = $script:indexingView
        if ($null -eq $view) {
            $progress = getIndexingProgress
            $view = @{ Name = ""; Processed = $progress.Processed; Total = $progress.Processed + $progress.Remaining; Failed = $progress.Failed; Eta = "" }
        }
        $line = getIndexingStatusLine ([string]$view.Name) $view.Processed $view.Total $view.Failed ([string]$view.Eta)
    }
    $statusView = getStatusBarView (isIndexing) $line ([string]$ui.StatusText.Text)
    if ($statusView.Source -ne "Indexing") {
        $ui.IndexingStatusText.Visibility = "Collapsed"
        $ui.StatusText.Visibility = "Visible"
        return
    }
    $ui.IndexingStatusText.Text = $statusView.Text
    $ui.IndexingStatusText.Visibility = "Visible"
    $ui.StatusText.Visibility = "Hidden"
}

$script:statusLineTimer = newTimer 1000 { safe { updateIndexingStatusLine } }
$script:statusLineTimer.Start()
