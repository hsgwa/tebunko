# 起動そのものに失敗したとき（Add-Type・読み込み・画面の組み立てで例外）の知らせ方。
# gui.ps1 が try で囲み、失敗したら reportStartupFailure を呼んで終了コードを受け取る。
# app_host.ps1 を読み込む前（writeErrorLog がまだ使えない）に起きた失敗にも対応するため、
# 文言は startup_error_view.ps1（gui.ps1 が先に読み込む）。gui.ps1 の読み込みの一番最初（Add-Type より前）に、ここだけを読み込む。

${appTitle} = "tebunko"

# 画面で起きた予期しないエラーの記録先（app_host.ps1 の writeErrorLog が使う）。今のワークスペースの中に置く。
# startGui を抜けた後（gui.ps1 の外側の catch）でも使えるよう、スクリプト直下に置く。ワークスペースが決まる前は $null を返す
function getGuiErrorLogFile {
    if ($null -eq $script:workspace) {
        return $null
    }
    return $script:workspace.GuiErrorLogFile
}

# 実際に書けた記録のファイルだけを返す（writeErrorLog は書けなくても例外を出さないため）
function getExistingRecordFile {
    param (
        [string]$path
    )

    if ($path -and (Test-Path -LiteralPath $path)) {
        return $path
    }
    return $null
}

# writeErrorLog がまだ使えない（app_host.ps1 を読み込む前）ときに起きた失敗を、固定の場所に追記する
# （置き場所は tebunko.bat と同じ。docs/safety/disclosure.md「起動に失敗したときの知らせ（tebunko.bat）」）。
# 制限言語モードでも動くよう、コマンドレットだけで書く（.NET のメソッドを呼ばない）。
# 書けなければ次の候補へ。すべて書けなければ $null を返す（そのときはメッセージボックスにファイル名を添えない）
function writeStartupErrorFile {
    param (
        [string]$text
    )

    $candidates = @(
        (Join-Path $env:LOCALAPPDATA "tebunko\startup_error.txt"),
        (Join-Path $env:TEMP "tebunko_startup_error.txt")
    )
    foreach ($candidate in $candidates) {
        try {
            $dir = Split-Path -Parent $candidate
            if (-not (Test-Path -LiteralPath $dir)) {
                New-Item -ItemType Directory -Force -Path $dir -ErrorAction Stop | Out-Null
            }
            Add-Content -LiteralPath $candidate -Value $text -Encoding UTF8 -ErrorAction Stop
            return $candidate
        } catch {
            continue
        }
    }
    return $null
}

# 起動そのものに失敗したときの知らせ。画面を閉じ（起動中の表示が残っていれば）、記録してメッセージボックスで知らせる。
# 戻り値は終了コード（gui.ps1 が exit に渡す）。メッセージボックスも出せない（制限言語モード・WPF が読めない）ときは、
# 元の例外を投げ直す（tebunko.bat の catch か、tebunko.exe のエラー出力が受けて知らせる）
function reportStartupFailure {
    param (
        $err
    )

    $recordFile = $null
    # 記録できる状態（app_host.ps1 の読み込み後で、writeErrorLog が使える）なら今までどおり gui_error_log.txt に、無ければ固定の場所に記録する
    if (Get-Command writeErrorLog -ErrorAction SilentlyContinue) {
        writeErrorLog "起動・実行中" $err
        $recordFile = getGuiErrorLogFile
        $recordFile = getExistingRecordFile $recordFile
    } else {
        $detail = getStartupErrorDetail $err ([string]$ExecutionContext.SessionState.LanguageMode) (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        $recordFile = writeStartupErrorFile $detail
    }
    if (Get-Command closeSplash -ErrorAction SilentlyContinue) {
        closeSplash
    }
    try {
        [System.Windows.MessageBox]::Show((getStartupErrorMessage $err.Exception.Message $recordFile), ${appTitle}, "OK", "Error") | Out-Null
        return 1
    } catch {
        throw $err
    }
}
