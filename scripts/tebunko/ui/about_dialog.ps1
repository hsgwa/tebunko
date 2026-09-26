# タブ右端の［⋯］メニューと「tebunko について」ダイアログ（画面層）。gui.ps1 が読み込む。
# 版・コミットの文字列は起動時に 1 回だけ組み立てる（gui.ps1 の $script:aboutView。ui/about_view.ps1 の getAboutView）。
# ダイアログを開くたびに VERSION.txt を読み直さない。

# ［⋯］は右クリックでしか開かない Button.ContextMenu を、左クリック・キーボード（Enter・Space）の
# どちらでも開けるようにする（Button の Click は両方で発生する）
$ui.MoreButton.Add_Click({
    param ($sender, $e)
    $menu = $sender.ContextMenu
    $menu.PlacementTarget = $sender
    $menu.Placement = "Bottom"
    $menu.IsOpen = $true
})

$ui.AboutMenuItem.Add_Click({
    safe { showAboutDialog }
})

function showAboutDialog {
    # 「tebunko について」ダイアログを開く
    $dialog = loadWindow "${xamlDir}\dialog_about.xaml"
    $dialog.Owner = $window
    $ctrl = @{}
    foreach ($name in @("AppIcon", "VersionText", "CommitText")) {
        $ctrl[$name] = $dialog.FindName($name)
    }
    if (Test-Path -LiteralPath ${iconFile}) {
        # .ico には大きさの違う絵が何枚も入っている。BitmapFrame.Create は先頭の 1 枚（16px）だけを返し、拡大されてぼやけるため、
        # 全フレームを読み、表示の大きさ × 画面の倍率に足りる最小のフレームを使う（足りるものが無ければ最大のもの）
        $decoder = New-Object System.Windows.Media.Imaging.IconBitmapDecoder(
            (New-Object Uri ${iconFile}),
            [System.Windows.Media.Imaging.BitmapCreateOptions]::None,
            [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
        $source = [System.Windows.PresentationSource]::FromVisual($window)
        $dpiScale = if ($source) { $source.CompositionTarget.TransformToDevice.M11 } else { 1.0 }
        $needed = [Math]::Ceiling($ctrl.AppIcon.Width * $dpiScale)
        $frames = @($decoder.Frames | Sort-Object PixelWidth)
        $fit = @($frames | Where-Object { $_.PixelWidth -ge $needed })
        $ctrl.AppIcon.Source = if ($fit.Count -gt 0) { $fit[0] } else { $frames[$frames.Count - 1] }
        [System.Windows.Media.RenderOptions]::SetBitmapScalingMode($ctrl.AppIcon, "HighQuality")
    }
    $ctrl.VersionText.Text = "版: $($script:aboutView.Version)"
    if ($script:aboutView.Commit -eq "") {
        $ctrl.CommitText.Visibility = "Collapsed"
    } else {
        $ctrl.CommitText.Text = "コミット: $($script:aboutView.Commit)"
    }
    [void]$dialog.ShowDialog()
}
