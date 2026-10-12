# ナビ下の「バージョン情報」と、そのダイアログ（画面層）。gui.ps1 が読み込む。
# 版・コミットの文字列は起動時に 1 回だけ組み立てる（gui.ps1 の $script:aboutView。ui/about_view.ps1 の getAboutView）。
# ダイアログを開くたびに VERSION.txt を読み直さない。

# ナビ下のリンク（Button の Click は、左クリックとキーボード（Enter・Space）の両方で発生する）
$ui.AboutLink.Add_Click({
    safe { showAboutDialog }
})

function openAboutLink {
    # リンクを既定のブラウザで開く（explorer.exe にアドレスを渡す）。getAboutLinks に載せたアドレス以外は開かない
    param (
        [string]$url
    )

    if (-not (testAboutLinkUrl $url)) {
        return
    }
    Start-Process -FilePath "explorer.exe" -ArgumentList "`"${url}`""
}

function showAboutDialog {
    # 「バージョン情報」ダイアログを開く
    $dialog = loadWindow "${xamlDir}\dialog_about.xaml"
    $dialog.Owner = $window
    $ctrl = @{}
    foreach ($name in @("AppIcon", "VersionText", "CommitText", "LicenseText", "GitHubLink", "ReleasesLink")) {
        $ctrl[$name] = $dialog.FindName($name)
    }
    # リンクは、アドレスをボタンの Tag に持たせる。開くのは getAboutLinks に載せたアドレスだけ（openAboutLink）
    foreach ($link in @(getAboutLinks)) {
        $button = $ctrl[$link.Id]
        $button.Content = $link.Text
        $button.Tag = $link.Url
        $button.ToolTip = $link.Url
        $button.Add_Click({
            param ($sender, $e)
            safe { openAboutLink $sender.Tag }
        })
    }
    $ctrl.LicenseText.Text = (@(getAboutLicenses) -join "`n")
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
