# ［インデックス管理］タブの判断: 追加・改名・削除と、その仕事の間の排他・閉じるときの判断。
# 画面に触らないため、そのままテストできる（tests\tebunko\ui\index\index_edit_view.Tests.ps1）。

function getIndexStoreJobStatus {
    # インデックスの削除・名前の変更を別スレッドで始めたときのステータス
    param (
        [string]$kind,      # "delete"（削除）か "rename"（名前の変更）
        [string]$name,
        [string]$newName
    )

    if ($kind -eq "rename") {
        return "インデックス [${name}] の名前を [${newName}] に変えています…（共有フォルダでは少し時間がかかります）"
    }
    return "インデックス [${name}] を削除しています…（件数によっては少し時間がかかります）"
}

function getIndexStoreJobFailedStatus {
    # 削除・名前の変更が失敗したときのステータス
    param (
        [string]$kind,
        [string]$name,
        [string]$errorText
    )

    $operation = if ($kind -eq "rename") { "名前の変更" } else { "削除" }
    return "インデックス [${name}] の${operation}に失敗しました：${errorText}"
}

function getBulkDeleteErrorStatus {
    # 削除が失敗した（1 件の削除の失敗と、複数削除が裏の仕事ごと失敗した場合。ワークスペースに届かないなど）ときのステータスと帯の文
    param (
        [string]$errorText
    )

    return "削除に失敗しました：${errorText}"
}

function getBulkDeleteFailedStatus {
    # 複数削除で、閉じるのを待っていて結果のダイアログを出さないときに、失敗があったことをステータスに残す文
    param (
        [int]$failedCount
    )

    return "${failedCount} 件のインデックスは削除できませんでした"
}

function getIndexStoreSourceFolderError {
    # 名前の変更の裏の仕事の戻り値（@{ SourceFolderError }）から、元のフォルダの記録を書けなかった理由を取り出す（書けたなら ""）
    param (
        $output
    )

    foreach ($item in @($output)) {
        if ($item -is [hashtable] -and $item.ContainsKey("SourceFolderError")) {
            return [string]$item.SourceFolderError
        }
    }
    return ""
}

function getIndexJobCloseAction {
    # インデックスの削除・名前の変更の途中で画面を閉じようとしたときの動き。
    #   "wait"    … 閉じずに、終わるまで待って閉じる（1 回目。閉じると設定への反映が抜けて、設定とインデックスのフォルダの名前が食い違うため）
    #   "confirm" … 待っている間にもう一度閉じようとした。待たずに閉じてよいかを確かめる
    #   "proceed" … 仕事が無い。ふつうの閉じ方に進む
    #   "waitMove" … ワークスペースを移している途中。閉じずに、移し終わるまで待って閉じる（途中で止めると移動元と移動先に分かれるため、確認は出さない）
    param (
        [bool]$indexBusy,
        [bool]$alreadyAsked,
        [bool]$workspaceMoving = $false
    )

    if ($workspaceMoving) {
        return "waitMove"
    }
    if (!$indexBusy) {
        return "proceed"
    }
    if ($alreadyAsked) {
        return "confirm"
    }
    return "wait"
}

function getIndexJobCloseConfirm {
    # 待っている間にもう一度閉じようとしたときの確認（showConfirm に渡す文言）
    return @{
        Title      = "変更の途中です"
        Heading    = "インデックスの変更の途中です。閉じますか？"
        Hint       = "いま閉じると、設定とインデックスのフォルダの名前が食い違うことがあります。終わるまで待てば、自動で閉じます。"
        CloseText  = "待たずに閉じる"
        CloseValue = "close"
        CancelText = "終わるまで待つ"
    }
}

function getIndexingCloseConfirm {
    # インデックスを更新中に画面を閉じようとしたときの確認（showConfirm に渡す文言）
    return @{
        Title      = "更新中です"
        Heading    = "インデックスを更新中です。中止して閉じますか？"
        Hint       = "更新したところまでは残ります。次に起動したときに続きから再開できます。"
        StopText   = "中止して閉じる"
        StopValue  = "stop"
        CancelText = "閉じない"
    }
}

function getIndexJobCloseAfterConfirm {
    # 確認から戻ったときの動き（"close" か "stay"）。聞いている間に仕事が終わったなら、答えにかかわらず閉じる
    # （終わったときの「終わったら閉じる」は、確認の間は動かさないため）
    param (
        [string]$answer,
        [bool]$indexBusy
    )

    if (!$indexBusy -or $answer -eq (getIndexJobCloseConfirm).CloseValue) {
        return "close"
    }
    return "stay"
}

function getIndexJobClosingStatus {
    # 削除・名前の変更の途中で閉じようとして、待っているときのステータス
    return "インデックスの変更が終わってから閉じます…（待たずに閉じるには、もう一度閉じる操作をします）"
}

function getWorkspaceMoveClosingStatus {
    # ワークスペースを移している途中で閉じようとして、待っているときのステータス
    return "ワークスペースを移し終わってから閉じます…"
}

function getSourceFolderFileFailedStatus {
    # 元のフォルダの記録（source_folder.txt）を書き直せなかったときのステータス
    param (
        [string]$errorText
    )

    return "元のフォルダの記録を書き直せませんでした：${errorText}"
}

function getIndexJobBlocker {
    # インデックス作成・削除・エクスポート・インポート・ワークスペースの変更は互いに排他（画面の可否の表）。
    # 動いているものがあれば、その名前を返す（無ければ空文字列。空なら操作してよい）
    param (
        [bool]$isIndexing,   # インデックス作成中
        [bool]$indexBusy,    # 前のインデックスの削除・名前の変更中
        [bool]$archiveBusy,  # エクスポート・インポート中
        [bool]$namingBusy = $false,  # ネットワークのワークスペースで、インデックスの名前を裏で決めている間（一覧はまだ空。この間に保存すると、ほかのインデックスが設定から消える）
        [bool]$workspaceMoving = $false   # ワークスペースの中身を移している間（別スレッド。移動元と移動先の両方を使っている）
    )

    if ($workspaceMoving) {
        return "ワークスペースの移動中"
    }
    if ($isIndexing) {
        return "インデックス作成中"
    }
    if ($namingBusy) {
        return "インデックス名の決定中"
    }
    if ($indexBusy) {
        return "削除・名前の変更中"
    }
    if ($archiveBusy) {
        return "エクスポート・インポート中"
    }
    return ""
}

function testIndexReadAllowed {
    # インデックス・設定を読む操作（検索を始める・F5 で読み直す・窓を前に出したときの読み直し）をしてよいか。
    # ワークスペースの中身を移している間は、移動元と移動先の両方を移す側が使っているため、キー操作でも始めない
    param (
        [bool]$workspaceMoving = $false
    )

    return -not $workspaceMoving
}

function testMainContentOperable {
    # ナビと中身（画面）の操作を受け付けてよいか。ワークスペースの中身を移している間は受け付けない
    # （マウスもキーも、設定・インデックス・検索の入口のどれからも書き込み・読み込みを始めさせない。待つ間のダイアログ自身は別の窓なので影響しない）
    param (
        [bool]$workspaceMoving = $false
    )

    return -not $workspaceMoving
}

function getIndexJobBlockedMessage {
    # 排他で操作をできないときのメッセージ（blocker は getIndexJobBlocker の結果。空なら空文字列）
    param (
        [string]$blocker,
        [string]$operation   # "追加" "編集" "削除" "エクスポート" "インポート" "ワークスペースの変更"
    )

    if ($blocker -eq "") {
        return ""
    }
    if ($operation -eq "ワークスペースの変更") {
        if ($blocker -eq "ワークスペースの移動中") {
            return "ワークスペースを移している間は、もう一度変えられません。終わるまでお待ちください。"
        }
        if ($blocker -eq "インデックス作成中") {
            return "更新中はワークスペースを変えられません。更新が終わるまでお待ちください（［中止］で止められます）。"
        }
        if ($blocker -eq "削除・名前の変更中") {
            return "前のインデックスの削除・名前の変更が終わるまでお待ちください。"
        }
        if ($blocker -eq "インデックス名の決定中") {
            return "インデックスの一覧を読み込んでいます。終わるまでお待ちください。"
        }
        return "エクスポート・インポートが終わるまでお待ちください。"
    }
    if ($blocker -eq "インデックス作成中") {
        return "更新中はインデックスを${operation}できません。更新が終わるまでお待ちください（［中止］で止められます）。"
    }
    return "${blocker}は${operation}できません。終わるまでお待ちください。"
}

function getIndexDeleteConfirmMessage {
    # インデックスを削除する確認の文。@{ Heading; Body }（showConfirm の heading と hint に渡す）
    #   1 件: 「営業」のインデックスを削除しますか？
    #   2 件以上: 選んだ N 件のインデックスを削除しますか？ ／ 名前を 3 つまで並べ、4 つ目からは「ほか N 件」にまとめる
    # 空の名前は数えない
    param (
        [string[]]$names
    )

    $names = @($names | Where-Object { $_ })
    $note = "元のファイルは削除されません。"
    if ($names.Count -le 1) {
        $name = if ($names.Count -eq 1) { $names[0] } else { "" }
        return @{ Heading = "「${name}」のインデックスを削除しますか？"; Body = $note }
    }
    $shown = @($names | Select-Object -First 3 | ForEach-Object { "「$_」" }) -join ""
    $others = $names.Count - 3
    $list = if ($others -ge 1) { "${shown}ほか {0:#,0} 件" -f $others } else { $shown }
    return @{
        Heading = "選んだ {0:#,0} 件のインデックスを削除しますか？" -f $names.Count
        Body = "${list}を削除します。${note}"
    }
}

function getIndexBulkResultView {
    # まとめてのエクスポート・削除の終わりに出す文言（仮）。results は 1 件ごとの @{ Name; Ok; Reason }
    #   operation: "エクスポート" / "削除"
    # @{ Status（ステータスバーの 1 行）; HasFailure; FailureHeading; FailureDetail（失敗した名前と理由の一覧）}
    param (
        [string]$operation,
        [object[]]$results
    )

    $results = @($results | Where-Object { $null -ne $_ })
    $failed = @($results | Where-Object { !$_.Ok })
    $okCount = $results.Count - $failed.Count
    if ($failed.Count -eq 0) {
        return @{ Status = "${okCount} 件のインデックスを${operation}しました"; HasFailure = $false; FailureHeading = ""; FailureDetail = "" }
    }
    return @{
        Status = "${operation}: ${okCount} 件成功・$($failed.Count) 件失敗"
        HasFailure = $true
        FailureHeading = "$($failed.Count) 件のインデックスを${operation}できませんでした（残りは終わっています）。"
        FailureDetail = ($failed | ForEach-Object { "「$($_.Name)」: $($_.Reason)" }) -join "`n"
    }
}

function testIndexEditInput {
    # 追加と詳細欄の編集の入力を調べ、直してほしい内容を返す（問題なければ空文字列）
    param (
        [string]$path,   # 入力された元のフォルダ
        [string]$name,   # 入力されたインデックス名
        $items,          # 今の一覧（Name・Path を持つ行）
        $current = $null # 編集中の行（重複の判定から外す）
    )

    $folder = normalizeFolderPath $path
    if ($folder -eq "") {
        return "元のフォルダを指定してください。"
    }
    $conflict = getIndexFolderConflict $folder @($items | Where-Object { $_ -ne $current })
    if ($conflict -ne "") {
        return $conflict
    }
    # @(getUsedIndexNames ...) と直接書くと集合が 1 要素の配列に入るだけなので、変数に受けてから配列にする
    $usedNames = getUsedIndexNames $items $current
    return (testIndexName $name.Trim() @($usedNames))
}
