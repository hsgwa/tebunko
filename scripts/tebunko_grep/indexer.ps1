# Office → TSV インデックス作成の起動口
#
# 画面で設定したクロール対象フォルダ（setting.config。チェックなしのフォルダは取り込まない）配下の
# Excel・Word・PowerPoint ファイルを取り込み、work\index のフォルダごと・拡張子ごとの集約ファイルに入れる。
# 本体は indexer\indexer_run.ps1 の invokeIndexer（流れは docs/01_インデックス作成_4_メインフローと取り込み一覧.md、
# スレッドの分け方は docs/00_共通_4_プロセスとスレッド.md）。
#
# ・Excel は Excel で抽出する（セルの表示値を得るため）
# ・Word・PowerPoint（.docx / .pptx 等）は、ファイルを直接読む（Word・PowerPointは使わない）
# ・旧形式（.doc / .ppt）は、Word・PowerPointで新形式に変換してから読む
# ・1 ファイルの取り込みは、取り込みのスレッド（既定はコア数 − 1、最大 4）で並べて行う
# ・全ファイルの更新日時・サイズ・状態を work\取り込み一覧.tsv に記録し、
#   前回から更新されたファイル・未取り込みのファイルだけを取り込む
#
# 画面を使わずに実行するときは、そのまま実行する（確認は求めない。表示内容はコンソールと work\インデックス作成ログ.txt に出す）。
#   -RetryFailed : 前回失敗し、その後更新されていないファイルも再取り込みする
#   -Channel     : 画面（インデクサのスレッド）から実行するときの受け渡しの口（newIndexerChannel）。
#                  確認・中止・進み具合は、ここでやり取りする（-RetryFailed は使わず、口の RetryFailed を使う）
#   終了コード   : 0 = 完了（ファイルごとの失敗は取り込み一覧に記録）/ 1 = 続けられないエラー / 2 = 中止（確認で取りやめた場合を含む）

param (
    [switch]$RetryFailed,
    $Channel = $null
)

. "$PSScriptRoot\indexer\indexer_lib.ps1"

$ErrorActionPreference = "Stop"

if ($null -eq $Channel) {
    $Channel = newIndexerChannel -retryFailed ([bool]$RetryFailed)
    $script:indexerEcho = $true
}
# 途中の処理が出力した値が混ざらないよう、最後の値（終了コード）を使う
$exitCode = [int]@(invokeIndexer $Channel)[-1]
exit $exitCode
