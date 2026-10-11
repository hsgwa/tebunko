# 画面の上に出す知らせの帯（検索の結果を保存できなかった・インデックスの削除などに失敗した・ワークスペースが使えない）の判断層。
# 画面に触らないため、そのままテストできる（tests/tebunko/ui/shell/notice_banner_view.Tests.ps1）。

function getNoticeBannerView {
    # 帯に出す文言と、色の区分・閉じられるか・自動で消えるか・ボタンを決める。
    #   kind: ResultSaveFailed（検索の結果を保存できなかった。検索の画面の上）/
    #         IndexOpFailed（インデックスの削除、まとめてのエクスポート・削除で失敗があった。インデックス管理の画面の上）/
    #         WorkspaceBlock（ワークスペースが使えない。窓の上）
    #   message: 出す文（複数行なら 1 行につなぎ、全文はツールヒントに出す）。空なら帯を出さない
    # 戻り値: @{ Show; Text; ToolTip; Level（warn / ng。getIndexingBannerBehavior と同じ区分）; Closable; AutoCloseSeconds（0 なら自動では消えない）; ActionText（ボタンの文言。無ければ空）}
    # どの帯も、閉じるまで（ワークスペースが使えない帯は、使えるようになるまで）残る。自動では消さない
    param (
        [string]$kind,
        [string]$message = ""
    )

    $lines = @(($message -split "`r?`n") | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
    $show = ($lines.Count -gt 0) -and ($kind -in @("ResultSaveFailed", "IndexOpFailed", "WorkspaceBlock"))
    if (!$show) {
        return @{ Show = $false; Text = ""; ToolTip = ""; Level = "warn"; Closable = $true; AutoCloseSeconds = 0; ActionText = "" }
    }
    $view = @{
        Show = $true; Text = ($lines -join "  "); ToolTip = ($lines -join "`n")
        Level = "warn"; Closable = $true; AutoCloseSeconds = 0; ActionText = ""
    }
    if ($kind -eq "WorkspaceBlock") {
        $view.Level = "ng"
        $view.Closable = $false
        $view.ActionText = "設定を開く"
    }
    return $view
}
