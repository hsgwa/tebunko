# 起動中の表示（xaml\splash.xaml）。
# スクリプトの読み込みに数秒かかるため、小さなウィンドウを出して起動していることを知らせる。
# 画面（$window）を描き終わったら closeSplash で閉じる（ContentRendered）。
# 読み込みの間は画面のスレッドが塞がるため、区切りごとに stepSplash で描画と入力の処理を進める（応答なしにしない）。

$script:splash = $null

# 出せなくても起動は続ける（$null を返す）
function showSplash {
    param (
        [string]$path
    )

    try {
        $stream = [System.IO.File]::OpenRead($path)
        try {
            $script:splash = [System.Windows.Markup.XamlReader]::Load($stream)
        } finally {
            $stream.Dispose()
        }
        $script:splash.Show()
    } catch {
        $script:splash = $null
    }
    return $script:splash
}

function stepSplash {
    param (
        [int]$percent
    )

    if ($null -eq $script:splash) {
        return
    }
    $script:splash.FindName("SplashProgress").Value = $percent
    # 自分のスレッドの Dispatcher に優先度 Background の空の仕事を渡し、それより優先度の高い描画・入力を先に処理させる
    $script:splash.Dispatcher.Invoke([action]{ }, [System.Windows.Threading.DispatcherPriority]::Background)
}

function closeSplash {
    if ($null -ne $script:splash) {
        $script:splash.Close()
        $script:splash = $null
    }
}
