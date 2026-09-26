# 画面の土台（XAML の読み込み・見た目の共通定義・エラーの記録）。
# 使う側が themeFile・iconFile と、エラーの記録先を返す関数 getGuiErrorLogFile を定義しておくこと。

# 見た目の共通定義を読み込む（画面自体は XAML の MergedDictionaries で読み込む）
function loadTheme {
    if ($null -eq ${script:theme}) {
        ${script:theme} = loadXaml ${themeFile}
    }
    return ${script:theme}
}

# XAML を読み込む。BaseUri にそのファイルの場所を渡し、XAML 内の相対パス
# （theme.xaml の MergedDictionaries）を解決できるようにする。
function loadXaml {
    param (
        [string]$path
    )

    $context = New-Object System.Windows.Markup.ParserContext
    $context.BaseUri = New-Object Uri $path
    $stream = [System.IO.File]::OpenRead($path)
    try {
        return [System.Windows.Markup.XamlReader]::Load($stream, $context)
    } finally {
        $stream.Dispose()
    }
}

function loadWindow {
    param (
        [string]$path
    )

    $loaded = loadXaml $path

    # アイコンは XAML に書かず、ここで読み込む（XamlReader.Load は XAML 内の相対パスを解決できないため）。
    # ファイルを掴んだままにしないよう OnLoad で読み切る。アイコンが無くても画面は開けるようにする。
    if (Test-Path ${iconFile}) {
        $loaded.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create(
            (New-Object Uri ${iconFile}),
            [System.Windows.Media.Imaging.BitmapCreateOptions]::None,
            [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
    }

    return $loaded
}

# 表のセルなど、コードから色を付ける箇所。色は theme.xaml のトークンから取り、画面と食い違わないようにする
function themeBrush {
    param (
        [string]$key
    )

    return (loadTheme)[$key]
}

# 予期しないエラーを getGuiErrorLogFile のファイルに残す。画面に出したメッセージだけでは、
# どこで起きたのかが後から分からないため（利用者に見せるのは従来どおりメッセージだけ）
function writeErrorLog {
    param (
        [string]$context,
        [System.Management.Automation.ErrorRecord]$record
    )

    try {
        $path = getGuiErrorLogFile
        $dir = Split-Path -Parent $path
        if (-not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Force -Path $dir | Out-Null
        }
        $text = @(
            "==== $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ${context} ===="
            "$($record.Exception.GetType().FullName): $($record.Exception.Message)"
            "$($record.InvocationInfo.PositionMessage)"
            "$($record.ScriptStackTrace)"
            if ($record.Exception.InnerException) { "内側: $($record.Exception.InnerException)" }
            ""
        ) -join "`r`n"
        [System.IO.File]::AppendAllText($path, $text, (New-Object System.Text.UTF8Encoding($true)))
    } catch {
        # 記録できなくても、画面の動作は止めない
    }
}
