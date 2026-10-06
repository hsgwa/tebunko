# ナビ（左の欄）の画面の並びと、画面を順に切り替える決まり（判断層）。
# 画面に触らないため、そのままテストできる（tests/tebunko/ui/shell/nav_view.Tests.ps1）。

function getScreenOrder {
    # 画面の名前を、ナビの上から順に並べて返す（Ctrl+Tab で切り替える順でもある）。名前はナビの項目の x:Name
    return @("SearchTab", "IndexTab", "SettingsTab")
}

function getScreenMargin {
    # 画面の中身（ContentHost）の外側の余白（左,上,右,下）。検索の画面は、白い面を窓いっぱいに敷いて帯の線を端まで引くため 0。
    # ほかの画面は、周りに余白を取る
    param (
        [string]$name
    )

    if ($name -eq "SearchTab") {
        return "0,0,0,0"
    }
    return "16,12,16,12"
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

function getShortcutAction {
    # キー操作（PreviewKeyDown）から、窓全体の動きを決める。画面のスレッドの外でも確かめられるよう、
    # キーの名前（Key の ToString）と Ctrl・Shift の有無だけを受ける。返す Action は、
    #   FocusSearchWord（Ctrl+F）・FocusFilter（Ctrl+Shift+F）・SwitchScreen（Ctrl+Tab・Ctrl+Shift+Tab。Screen に切り替え先）・
    #   Refresh（F5）・CancelSearch（Escape）・None（ここでは扱わない）
    param (
        [string]$key,
        [bool]$ctrl,
        [bool]$shift,
        [string]$current
    )

    if ($key -eq "F" -and $ctrl) {
        return @{ Action = $(if ($shift) { "FocusFilter" } else { "FocusSearchWord" }); Screen = "SearchTab" }
    }
    if ($key -eq "Tab" -and $ctrl) {
        return @{ Action = "SwitchScreen"; Screen = (getNextScreen $current $(if ($shift) { -1 } else { 1 })) }
    }
    if ($key -eq "F5") {
        return @{ Action = "Refresh"; Screen = $current }
    }
    if ($key -eq "Escape") {
        return @{ Action = "CancelSearch"; Screen = $current }
    }
    return @{ Action = "None"; Screen = $current }
}
