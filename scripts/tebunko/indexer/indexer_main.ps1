# インデクサ（indexer.ps1）の本体。indexer.ps1 が読み込んで呼ぶ（引数はそのまま転送する）。
# 流れは docs/design/indexing/flow.md、スレッドの分け方は docs/design/structure/threads.md。

function invokeIndexerMain {
    # 画面なしで起動したとき（$Channel が無い）は、壊れた設定ファイルを退避して既定の設定で続ける
    # （知らせは invokeIndexer がログを開いた直後に出す）。画面から -Channel 付きで動くときは、画面が起動時に退避済み
    param (
        [switch]$RetryFailed,
        $Channel = $null
    )

    $script:settingsRecovery = ""
    if ($null -eq $Channel) {
        $script:settingsRecovery = repairBrokenSettings
    }

    $ErrorActionPreference = "Stop"
    initWorkspace

    if ($null -eq $Channel) {
        $Channel = newIndexerChannel -retryFailed ([bool]$RetryFailed) -includeCloud ((readCloudFiles) -eq ${cloudFilesDownload})
        $script:indexerEcho = $true
    }
    # 途中の処理が出力した値が混ざらないよう、最後の値（終了コード）を使う
    return [int]@(invokeIndexer $Channel)[-1]
}
