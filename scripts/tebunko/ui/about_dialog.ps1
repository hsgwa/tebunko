# ナビ下の「バージョン情報」と、そのダイアログ（画面層）。gui.ps1 が読み込む。
# 版・コミットの文字列は起動時に 1 回だけ組み立てる（gui.ps1 の $script:aboutView。ui/about_view.ps1 の getAboutView）。
# ダイアログを開くたびに VERSION.txt を読み直さない。

# ナビ下のリンク（Button の Click は、左クリックとキーボード（Enter・Space）の両方で発生する）
$ui.AboutLink.Add_Click({
    safe { showAboutDialog }
})

function showAboutDialog {
    # 「バージョン情報」ダイアログを開く
    $dialog = loadWindow "${xamlDir}\dialog_about.xaml" ${fontsDir}
    $dialog.Owner = $window
    $ctrl = @{}
    foreach ($name in @("AppIcon", "VersionText", "CommitText")) {
        $ctrl[$name] = $dialog.FindName($name)
    }
    # ベクターの絵なので、表示の大きさ・画面の倍率が変わってもぼやけない
    $icon = loadAppIcon
    if ($null -ne $icon) {
        $ctrl.AppIcon.Source = $icon
    }
    $ctrl.VersionText.Text = "バージョン $($script:aboutView.Version)"
    if ($script:aboutView.Commit -eq "") {
        $ctrl.CommitText.Visibility = "Collapsed"
    } else {
        $ctrl.CommitText.Text = "|　コミット $($script:aboutView.Commit)"
    }
    [void](showOwnedDialog $dialog)
}
