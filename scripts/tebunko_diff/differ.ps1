# 比較の抽出プロセス（画面がウィンドウ無しで起動する。入力は求めない）
#
#   powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -WindowStyle Hidden -File differ.ps1 -JobDir <比較の作業フォルダ>
#
# 比較の作業フォルダ（job\diff_job.ps1）の 抽出要求.tsv に書かれたファイルを順に抽出し、
# 場所ごとの TSV を <作業フォルダ>\<側>\<番号>\ に置く。本体は job\diff_extract.ps1 の invokeDiffer。
#   終了コード: 0 = 完了（ファイルごとの失敗は 抽出結果.tsv）/ 1 = 続けられないエラー（比較エラー.txt）/ 2 = 中止

param (
    [Parameter(Mandatory = $true)]
    [string]$JobDir
)

. "$PSScriptRoot\lib.ps1"
. "$PSScriptRoot\..\shared\office\office_reader.ps1"
. "$PSScriptRoot\..\shared\office\office_app.ps1"
. "$PSScriptRoot\..\shared\office\office_extract.ps1"
. "$PSScriptRoot\job\diff_extract.ps1"

$ErrorActionPreference = "Stop"

trap {
    # 続けられないエラー。画面に伝えるため、メッセージをファイルに書いて終了する
    try {
        [System.IO.File]::WriteAllText((Join-Path $JobDir ${diffErrorFileName}), $_.Exception.Message, ${utf8Bom})
    } catch {}
    try { stopWatchdog } catch {}
    try { stopAllApps } catch {}
    exit 1
}

exit (invokeDiffer $JobDir)
