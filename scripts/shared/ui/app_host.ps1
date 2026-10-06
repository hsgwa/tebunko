# 画面の土台（XAML の読み込み・見た目の共通定義・エラーの記録）。
# 使う側が themeFile・iconFile と、エラーの記録先を返す関数 getGuiErrorLogFile を定義しておくこと。

# 見た目の共通定義を読み込む（画面自体は XAML の MergedDictionaries で読み込む）
function loadTheme {
    if ($null -eq ${script:theme}) {
        ${script:theme} = loadXaml ${themeFile}
    }
    return ${script:theme}
}

# XAML の中身を文字列で返す。単一 .ps1 版（展開せずに動く試験版。結合の道具 tools/new_single_script.ps1 が
# xaml の中身を ${bundledXaml}[<このファイルを指す絶対パス>] に埋め込む）はそこから返し、
# 無ければ（zip 版）ファイルから読む。
# 探す前に GetFullPath で正規化する（".." を含む書き方でも、結合の道具が埋めた鍵と一致させるため。
# ファイルが実在しなくても使える＝単一 .ps1 版で theme.xaml が無くても引ける）
function getXamlText {
    param (
        [string]$path
    )

    $bundled = Get-Variable -Name bundledXaml -ErrorAction SilentlyContinue
    if ($null -ne $bundled -and $bundled.Value) {
        $full = [System.IO.Path]::GetFullPath($path)
        if ($bundled.Value.ContainsKey($full)) {
            return $bundled.Value[$full]
        }
    }
    return [System.IO.File]::ReadAllText($path, (New-Object System.Text.UTF8Encoding($true)))
}

# <ResourceDictionary Source="..." /> を、そのファイルの中身（ResourceDictionary の子要素）に差し替える。
# XamlReader に渡す前に XML としてテキストで組み立て直すことで、BaseUri（System.Windows.Application が無いと
# 相対パスを解決できない）に頼らずに済む。単一 .ps1 版で xaml がファイルとして存在しなくても働く
function inlineMergedDictionaries {
    param (
        [System.Xml.XmlDocument]$xml,
        [string]$basePath
    )

    foreach ($node in @($xml.SelectNodes("//*[local-name()='ResourceDictionary' and @Source]"))) {
        # GetFullPath は実在しないファイルでも正規化できる（Resolve-Path と違い、単一 .ps1 版で theme.xaml が
        # 実在しなくても働く）
        $refPath = [System.IO.Path]::GetFullPath((Join-Path (Split-Path -Parent $basePath) $node.Attributes["Source"].Value))
        $refXml = New-Object System.Xml.XmlDocument
        $refXml.LoadXml((getXamlText $refPath))
        inlineMergedDictionaries $refXml $refPath
        $replacement = $xml.CreateElement($node.LocalName, $node.NamespaceURI)
        foreach ($attribute in @($refXml.DocumentElement.Attributes)) {
            if ($attribute.Prefix -eq "xmlns" -or $attribute.Name -eq "xmlns") {
                continue
            }
            [void]$replacement.SetAttribute($attribute.Name, $attribute.NamespaceURI, $attribute.Value)
        }
        foreach ($child in @($refXml.DocumentElement.ChildNodes)) {
            [void]$replacement.AppendChild($xml.ImportNode($child, $true))
        }
        [void]$node.ParentNode.ReplaceChild($replacement, $node)
    }
}

# XAML を読み込む（Window・タブ・ダイアログの中身）。theme.xaml への参照は、読み込む前にテキストで差し替える
function loadXaml {
    param (
        [string]$path
    )

    $xml = New-Object System.Xml.XmlDocument
    $xml.LoadXml((getXamlText $path))
    inlineMergedDictionaries $xml $path

    $stream = New-Object System.IO.MemoryStream
    $xml.Save($stream)
    $stream.Position = 0
    try {
        $loaded = [System.Windows.Markup.XamlReader]::Load($stream)
    } finally {
        $stream.Dispose()
    }

    # 画面の既定のフォント（Font.Body）。XAML のルートが FontFamily="{DynamicResource Font.Body}" で使う。
    # XamlReader.Load は XAML 内の相対パスを解決できないため、同梱のフォントを指す FontFamily はここで作って入れる
    if ($loaded -is [System.Windows.FrameworkElement]) {
        # 型を明示する（PowerShell のラッパーのまま入れると、DynamicResource が文字列として解決して失敗する）
        $loaded.Resources["Font.Body"] = [System.Windows.Media.FontFamily](newAppFontFamily)
    }
    return $loaded
}

# 画面の既定のフォント。同梱のフォント（${fontsDir}）の Rethink Sans を先に、足りない文字（日本語）は Yu Gothic UI・Meiryo UI で表す。
# フォルダが無いとき（単一 .ps1 版など）は Yu Gothic UI・Meiryo UI だけにする
function newAppFontFamily {
    param (
        [string]$fontsFolder
    )

    if (-not $fontsFolder) {
        $variable = Get-Variable -Name fontsDir -ErrorAction SilentlyContinue
        if ($null -ne $variable) {
            $fontsFolder = [string]$variable.Value
        }
    }
    $fallback = "Yu Gothic UI, Meiryo UI"
    if ($fontsFolder -and (Test-Path -LiteralPath $fontsFolder -PathType Container)) {
        # 末尾が \ のフォルダを指す URI にする（そうしないと、./# の探し先がフォルダの 1 つ上になる）
        $baseUri = New-Object System.Uri (([System.IO.Path]::GetFullPath($fontsFolder).TrimEnd("\") + "\"))
        return New-Object System.Windows.Media.FontFamily -ArgumentList $baseUri, "./#Rethink Sans, $fallback"
    }
    return New-Object System.Windows.Media.FontFamily -ArgumentList $fallback
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
