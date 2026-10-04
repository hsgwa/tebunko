# 起動そのものに失敗したときの知らせの文言（判断層。画面・WPF に触らない）。
# 記録に書く内容と、メッセージボックスに出す文を決める。出す・書くのは startup_error.ps1。
# 制限言語モードでも読み込めるよう、.NET のメソッドを呼ばない。

# 固定の場所に書く記録の 1 件分（app_host.ps1 を読み込む前の失敗用）
function getStartupErrorDetail {
    param (
        $err,
        [string]$languageMode,
        [string]$timestamp
    )

    return "==== $timestamp 起動・実行中 ====`r`nLanguageMode: $languageMode`r`n$($err.Exception.Message)`r`n$($err.InvocationInfo.PositionMessage)`r`n`r`n"
}

# メッセージボックスの本文。記録できたファイルがあれば、そのファイル名を添える
function getStartupErrorMessage {
    param (
        [string]$message,
        [string]$recordFile
    )

    $suffix = if ($recordFile) { "`n`n詳しい内容は $(Split-Path -Leaf $recordFile) に残しています。" } else { "" }
    return "予期しないエラーが発生しました。`n${message}${suffix}"
}