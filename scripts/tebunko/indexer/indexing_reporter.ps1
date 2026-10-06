# インデックス作成の司令のスレッドの進み具合と、画面の確認待ちをまとめる（IndexingReporter）。
# 受け渡しの口（newIndexerChannel。hashtable のまま）を持ち、進み具合を書く（writeIndexingProgress を呼ぶ）・
# 画面の確認を待つ（waitForIndexingApproval だった処理をメソッドにした）。
# ログ（writeIndexerLog・$script:indexerLog）は部品として別に残し、ここには入れない（docs/design/structure/classes.md）。
# クラスの本体では $script: 変数を直接読まない（docs/design/structure/classes.md「クラスと関数の使い分け」）ため、
# インデックス作成の段階（${indexingPhase*}）は呼び出し元（司令）から引数で受け取る

class IndexingReporter {
    # 受け渡しの口（画面とやり取りする hashtable。newIndexerChannel）
    $Channel

    IndexingReporter($channel) {
        $this.Channel = $channel
    }

    [void] Progress([string]$phase, [int]$processed, [int]$remaining, [int]$failed, [string]$detail) {
        # インデックス作成の進み具合を受け渡しの口に書く（画面が 1 秒ごとに読む）
        writeIndexingProgress $phase $processed $remaining $failed $detail $this.Channel
    }

    [object] WaitForApproval([string]$phase, $plan, [int]$targetCount, [int]$failedCount, [int]$timeoutMinutes) {
        # 取り込み対象の件数を画面に渡し、［インデックス作成を開始］か［キャンセル］の返事を待つ。
        #   取り込む → @{ RetryFailed } / 取りやめ（中止を求められた場合を含む） → $null
        # 画面が返事をしないまま待ち続けないよう、timeoutMinutes で打ち切って取りやめる。
        # ローカル変数はプロパティ名 Channel と大文字・小文字だけの違いにしない（$ch にする。PowerShell のクラスは
        # 同名（大文字・小文字を区別しない）のローカル変数への代入をプロパティへの代入と見なし、$this. を要求してエラーになる）
        $ch = $this.Channel
        [void]$ch.Answered.Reset()
        $ch.Answer = $null
        $ch.Plan = @($plan)
        writeIndexingProgress $phase 0 $targetCount $failedCount "更新する内容を画面で確認しています…" $ch
        writeIndexerLog ""
        writeIndexerLog "取り込み対象を画面に表示しました。［インデックス作成を開始］が押されるまで待ちます。（${timeoutMinutes} 分待っても返事が無ければ取りやめます）"

        try {
            if (!$ch.Answered.WaitOne([TimeSpan]::FromMinutes($timeoutMinutes))) {
                writeIndexerLog "画面からの返事が ${timeoutMinutes} 分ありませんでした。インデックス作成を取りやめます。" "Yellow"
                return $null
            }
            if ($ch.Stop) {
                return $null
            }
            return $ch.Answer
        } finally {
            $ch.Plan = $null
        }
    }
}
