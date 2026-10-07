# 起動中の表示（xaml\splash.xaml）。
# スクリプトの読み込みに数秒かかるため、小さなウィンドウを出して起動していることを知らせる。
# 画面（$window）を描き終わったら closeSplash で閉じる（ContentRendered）。
# 読み込みの間は画面のスレッドが塞がるため、区切りごとに stepSplash で描画と入力の処理を進める（応答なしにしない）。

$script:splash = $null

# XAML の中身を文字列で返す。
# shared/ui/app_host.ps1（getXamlText）はこの時点でまだ読み込んでいない（早く出すため）ので、
# 単一 .ps1 かどうかの見分け方（${bundledXaml}）だけをここでも行う
function getSplashXamlText {
    param (
        [string]$path
    )

    $bundled = Get-Variable -Name bundledXaml -ErrorAction SilentlyContinue
    # shared/ui/app_host.ps1 の getXamlText と同じく GetFullPath で正規化してから探す（鍵の決め方をそろえる）
    $fullPath = [System.IO.Path]::GetFullPath($path)
    if ($null -ne $bundled -and $bundled.Value -and $bundled.Value.ContainsKey($fullPath)) {
        return $bundled.Value[$fullPath]
    }
    return [System.IO.File]::ReadAllText($path, (New-Object System.Text.UTF8Encoding($true)))
}

# 出せなくても起動は続ける（$null を返す）。
# iconXamlPath は、アイコン（ベクターの絵。xaml\app_icon.xaml）の XAML
function showSplash {
    param (
        [string]$path,
        [string]$iconXamlPath = ""
    )

    try {
        $text = getSplashXamlText $path
        # 配列は ,（単項）で 1 つの引数として渡す。付けないと New-Object が要素ごとの引数に展開して失敗し、下の catch で黙って握りつぶされる
        $stream = New-Object System.IO.MemoryStream -ArgumentList (,[System.Text.Encoding]::UTF8.GetBytes($text))
        try {
            $script:splash = [System.Windows.Markup.XamlReader]::Load($stream)
        } finally {
            $stream.Dispose()
        }
        setSplashIcon $script:splash $iconXamlPath
        $script:splash.Show()
    } catch {
        $script:splash = $null
    }
    return $script:splash
}

# アイコンは飾りなので、読めなくても何もしない
function setSplashIcon {
    param (
        $window,
        [string]$iconXamlPath
    )

    if (-not $iconXamlPath) {
        return
    }
    try {
        $stream = New-Object System.IO.MemoryStream -ArgumentList (,[System.Text.Encoding]::UTF8.GetBytes((getSplashXamlText $iconXamlPath)))
        try {
            $icon = [System.Windows.Markup.XamlReader]::Load($stream)
        } finally {
            $stream.Dispose()
        }
        $window.FindName("SplashIcon").Source = $icon
    } catch {
        $null = $_
    }
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
