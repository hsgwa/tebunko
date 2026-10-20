# 閉じる処理の記録（どの段まで進んだかを 1 行ずつ足して書く）。
#
# 画面を閉じたのにプロセスが終わらないとき、どの段で止まったかを後から読めるようにする（docs/design/structure/closing.md）。
# 記録先は引数で受ける（環境変数など、記録先を決める仕組みは呼び出し側が持つ）。記録先が空のときは何もしない。
# 書けなくても例外を出さない（記録のせいで閉じる処理を止めない）。

function writeCloseTrace {
    param (
        [string]$Path,
        [string]$Step
    )

    if ([string]::IsNullOrEmpty($Path)) {
        return
    }
    try {
        $line = "{0:yyyy-MM-dd HH:mm:ss.fff} {1}`r`n" -f (Get-Date), $Step
        [System.IO.File]::AppendAllText($Path, $line, (New-Object System.Text.UTF8Encoding($false)))
    } catch {
    }
}
