# ナビ（左の欄）の画面の並びと、画面を順に切り替える決まり（判断層）。
# 画面に触らないため、そのままテストできる（tests/tebunko/ui/nav_view.Tests.ps1）。
# Office の停止の画面をやめるときは、getScreenOrder の表から 1 行を消すだけで済む。

function getScreenOrder {
    # 画面の名前を、ナビの上から順に並べて返す（Ctrl+Tab で切り替える順でもある）。名前はナビの項目の x:Name
    return @("SearchTab", "IndexTab", "SettingsTab", "KillTab")
}

function isScreenName {
    param (
        [string]$name
    )

    return (@(getScreenOrder) -contains $name)
}

function getNextScreen {
    # 今の画面から step 個先（負なら前）の画面の名前。端から先は反対の端に回る。
    # 今の画面が一覧に無いとき（まだ選んでいない）は、進むなら先頭、戻るなら末尾にする
    param (
        [string]$current,
        [int]$step
    )

    $order = @(getScreenOrder)
    $index = [Array]::IndexOf($order, $current)
    if ($index -lt 0) {
        return $(if ($step -ge 0) { $order[0] } else { $order[$order.Count - 1] })
    }
    $next = (($index + $step) % $order.Count + $order.Count) % $order.Count
    return $order[$next]
}
