# ［2 比較］タブ（トグル・入力・要約・進み具合）。

function isDiffRunning {
    # 比較の抽出プロセスが動いているか（［9 プロセス停止］の注意に使う）
    return $false
}

function onDiffKeyDown {
    # ［2 比較］のキー操作。扱ったら $true を返す
    param (
        $e,
        $modifiers
    )

    return $false
}

function confirmDiffClosing {
    # ウィンドウを閉じる前。閉じてよければ $true
    return $true
}

function closeDiffPage {
    # ウィンドウを閉じた後の後始末
}

function startDiffPage {
    # 起動時
}
