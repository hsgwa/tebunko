# インデックス作成の起動口。画面（gui.ps1）がウィンドウを出さずに別スレッドで実行する（indexing_session.ps1）。
# 本体は indexer\indexer_main.ps1（invokeIndexerMain）にある。

param (
    [switch]$RetryFailed,
    $Channel = $null
)

. "$PSScriptRoot\indexer\indexer_lib.ps1"
. "$PSScriptRoot\indexer\indexer_main.ps1"

exit (invokeIndexerMain -RetryFailed:$RetryFailed -Channel $Channel)
