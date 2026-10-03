# 起動中の表示（xaml\splash.xaml）。
# スクリプトの読み込みに数秒かかるため、小さなウィンドウを出して起動していることを知らせる。
# 画面（$window）を描き終わったら closeSplash で閉じる（ContentRendered）。
# 読み込みの間は画面のスレッドが塞がるため、区切りごとに stepSplash で描画と入力の処理を進める（応答なしにしない）。

$script:splash = $null

# 出せなくても起動は続ける（$null を返す）。
# shared/ui/app_host.ps1（getXamlText）はこの時点でまだ読み込んでいない（早く出すため）ので、
# 単一 .ps1 かどうかの見分け方（${bundledXaml}）だけをここでも行う
function showSplash {
    param (
        [string]$path
    )

    try {
        $bundled = Get-Variable -Name bundledXaml -ErrorAction SilentlyContinue
        # shared/ui/app_host.ps1 の getXamlText と同じく GetFullPath で正規化してから探す（鍵の決め方をそろえる）
        $fullPath = [System.IO.Path]::GetFullPath($path)
        $text = if ($null -ne $bundled -and $bundled.Value -and $bundled.Value.ContainsKey($fullPath)) {
            $bundled.Value[$fullPath]
        } else {
            [System.IO.File]::ReadAllText($path, (New-Object System.Text.UTF8Encoding($true)))
        }
        $stream = New-Object System.IO.MemoryStream([System.Text.Encoding]::UTF8.GetBytes($text))
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
