# 待つ間のダイアログ（shared\ui\shell.ps1 の showWaitDialog）に出す文言・秒読み・進み具合の棒の値の決定。
# 画面に触らないため、そのままテストできる（tests\tebunko\ui\wait_view.Tests.ps1）。

function getWaitDialogView {
    # 場面ごとの見出し・説明・押せないボタンの文字を返す。@{ Heading; Description; ButtonText }
    #   Stopping:  ［中止］を押したあと、更新中のファイルが終わるまで
    #   CloseWait: 窓を閉じるときに、インデックス作成が止まるまで（強制終了の秒読みつき）
    #   Moving:    ワークスペースの中身を移す間
    #   MovingClose: ワークスペースを移す間に窓を閉じようとしたとき（移し終わってから閉じる。秒読みは出さない）
    param (
        [string]$kind
    )

    switch ($kind) {
        "Stopping" {
            return @{ Heading = "中止しています…"; Description = "更新中のファイルが終わるまでお待ちください。"; ButtonText = "" }
        }
        "CloseWait" {
            return @{ Heading = "中止しています…"; Description = "更新中のファイルが終わると、画面を閉じます。"; ButtonText = "閉じない" }
        }
        "Moving" {
            return @{ Heading = "ワークスペースを移しています…"; Description = "中身の量によっては時間がかかります。終わるまでお待ちください。"; ButtonText = "" }
        }
        "MovingClose" {
            return @{ Heading = "ワークスペースを移しています…"; Description = "移し終わると、画面を閉じます。"; ButtonText = "" }
        }
        default {
            throw "待つ間のダイアログの場面が不明です: $kind"
        }
    }
}

function getForceCloseCountdownText {
    # 強制終了までの残りの秒から、秒読みの 1 行を作る。端数は切り上げ、0 秒より小さいときは 0 にする
    param (
        [double]$secondsLeft
    )

    $seconds = [int][Math]::Max(0, [Math]::Ceiling($secondsLeft))
    return "あと $seconds 秒で強制終了します"
}

function getWaitProgressValue {
    # 処理した件数と総数から、進み具合の棒の値を決める。@{ Indeterminate; Value }
    # 総数が 0 以下（まだ数え終わっていない・不明）のときは流れる棒（不確定）。分かるときは 0 から 1 の割合（1 を超えない）
    param (
        [double]$processed,
        [double]$total
    )

    if ($total -le 0) {
        return @{ Indeterminate = $true; Value = 0.0 }
    }
    $ratio = [Math]::Min(1.0, [Math]::Max(0.0, $processed / $total))
    return @{ Indeterminate = $false; Value = $ratio }
}
