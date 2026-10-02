# ［1 インデックス管理］タブの判断（入力の検査・名前の重複）。
# 画面に触らないため、そのままテストできる（tests\tebunko\ui\index_view.Tests.ps1）。

function getUsedIndexNames {
    # 一覧のインデックス名の集合（大文字・小文字を区別しない）。except に渡した行の名前は含めない
    param (
        $items,
        $except = $null
    )

    $used = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($item in $items) {
        if ($item -ne $except -and $item.Name) {
            [void]$used.Add($item.Name)
        }
    }
    # HashSet をそのまま return すると PowerShell が中身を展開してしまい（0 件なら $null、1 件なら文字列）、
    # 受け取った側の .Contains が落ちる・部分一致になる。, を付けて集合のまま返す
    return , $used
}

function getIndexAddedStatus {
    # インデックスを追加したときのステータス。フォルダの有無によらず同じ文言にする（有無は一覧の列で分かる）
    param (
        [string]$name
    )

    return "インデックス [${name}] を追加しました。［インデックス作成を開始］を押すと中身を取り込みます"
}

function getFailedFileCheckingStatus {
    # 取り込みに失敗したファイルのダブルクリックで、元のファイルを確かめている間のステータス
    param (
        [string]$path
    )

    return "元のファイルを確かめています…：${path}"
}

function getFailedFileUnreachableStatus {
    # 同じ操作で、接続できないと分かったときのステータス（見つからないときとは別の文言）
    param (
        [string]$path
    )

    return "元のフォルダに接続できません：${path}"
}

function getFailedFileOtherStatus {
    # 同じ操作で、接続できる・できないのどちらでもない理由（アクセス拒否・一覧に無いネットワークのエラーなど）で
    # 確かめられなかったときのステータス。フォルダをたどらず、文言だけ出す
    param (
        [string]$message
    )

    return "元のファイルを確かめられませんでした：${message}"
}

function getIndexJobBlocker {
    # インデックス作成・削除・エクスポート・インポート・ワークスペースの変更は互いに排他（画面の可否の表）。
    # 動いているものがあれば、その名前を返す（無ければ空文字列。空なら操作してよい）
    param (
        [bool]$isIndexing,   # インデックス作成中
        [bool]$indexBusy,    # 前のインデックスの削除中
        [bool]$archiveBusy   # エクスポート・インポート中
    )

    if ($isIndexing) {
        return "インデックス作成中"
    }
    if ($indexBusy) {
        return "削除中"
    }
    if ($archiveBusy) {
        return "エクスポート・インポート中"
    }
    return ""
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
        if ($blocker -eq "インデックス作成中") {
            return "インデックス作成中はワークスペースを変えられません。インデックス作成が終わるまでお待ちください（［中止］で止められます）。"
        }
        if ($blocker -eq "削除中") {
            return "前のインデックスの削除が終わるまでお待ちください。"
        }
        return "エクスポート・インポートが終わるまでお待ちください。"
    }
    if ($blocker -eq "インデックス作成中") {
        return "インデックス作成中はインデックスを${operation}できません。インデックス作成が終わるまでお待ちください（［中止］で止められます）。"
    }
    return "${blocker}は${operation}できません。終わるまでお待ちください。"
}

function getIndexTabButtonsEnabled {
    # 排他（getIndexJobBlocker の結果）と、一覧で選んでいる行の有無から、［1 インデックス管理］の各ボタンの可否を返す。
    #   New/Edit/Remove: ［追加…］［編集…］［削除］/ Export/Import: ［エクスポート…］［インポート…］
    # ［編集…］［削除］［エクスポート…］は、1 件選んでいるときだけ有効
    # （［インデックス作成を開始］は updateIndexingButton が、［8 設定］の［変更…］は押したときに testWorkspaceChangeable が、
    # 同じ getIndexJobBlocker の結果で止める）
    param (
        [string]$blocker,
        [bool]$hasSelection
    )

    $free = ($blocker -eq "")
    return @{
        New = $free; Edit = ($free -and $hasSelection); Remove = ($free -and $hasSelection)
        Export = ($free -and $hasSelection); Import = $free
    }
}

function getImportResultStatus {
    # インポートの結果（importIndex の戻り値）から、ステータスに出す文言を返す。
    # Warnings（同じ元のフォルダが別の名前で既に登録されている等）と、高速検索が次のインデックス作成の後に効くことを添える
    param (
        $result
    )

    $text = "インデックス [$($result.Name)] をインポートしました（$($result.Files) ファイル）。" +
        "高速検索は次のインデックス作成の後に効きます。"
    foreach ($warning in @($result.Warnings)) {
        $text += " ${warning}"
    }
    return $text
}

function testIndexImportInput {
    # インポートのダイアログの入力を調べ、直してほしい内容を返す（問題なければ空文字列）。
    # 名前が既にあるインデックスと重なることは断らない（上書き・別名・取りやめの確認に回す。getImportIndexName）。
    # 元のフォルダは、追加・編集と同じ決まり（getIndexFolderConflict）で調べる。同じ名前の行は上書きで置き換わる
    # （別名なら残る）ため、ここでは比べる相手から外す（別名で入れて重なったときは、インポートの側で止める）
    param (
        [string]$folder,   # 入力された元のフォルダ
        [string]$name,     # 入力されたインデックス名
        $items = @()       # 今の一覧（Name・Path を持つ行）
    )

    $normalized = normalizeFolderPath $folder
    if ($normalized -eq "") {
        return "元のフォルダを指定してください。"
    }
    $nameMessage = testIndexName $name.Trim() @()
    if ($nameMessage -ne "") {
        return $nameMessage
    }
    return (getIndexFolderConflict $normalized @($items | Where-Object { $_.Name -ine $name.Trim() }))
}

function getImportSuggestedName {
    # インポートのダイアログに初めに入れる名前。zip の名前が今の一覧に無ければそのまま、あれば「名前(2)」…にする。
    # usedNames は getUsedIndexNames の結果（集合。@(...) に入った 1 要素の配列でも、配列でもよい）
    param (
        [string]$indexName,
        $usedNames = $null
    )

    if (testImportNameCollision $indexName $usedNames) {
        return (newIndexName $indexName $usedNames)
    }
    return $indexName
}

function testImportNameCollision {
    # 入力された名前が今の一覧にあり、上書き・別名・取りやめの確認が要るか（大文字・小文字は区別しない）
    param (
        [string]$name,
        $usedNames = $null   # getUsedIndexNames の結果（集合・配列・文字列・$null。入れ子の集合でもよい）
    )

    foreach ($entry in @($usedNames)) {
        foreach ($used in @($entry)) {
            if ($used -and $used -ieq $name) {
                return $true
            }
        }
    }
    return $false
}

function getIndexImportNotice {
    # インポートのダイアログの説明（目録から読んだ合計の大きさ・ファイル数を添える）
    param (
        $info   # readIndexArchiveInfo の結果
    )

    $mb = [Math]::Max(0.1, [Math]::Round($info.Bytes / 1MB, 1))
    return "エクスポートされたインデックスを読み込みます（$($info.Files) ファイル・約 ${mb} MB）。" +
        "名前と、元のフォルダの場所を変えられます。同じ名前のインデックスが既にあるときは、後で上書き・別名・取りやめを選べます。"
}

function getIndexImportOverwriteConfirmMessage {
    # 名前が既にあるインデックスと重なったときの確認（choices は showConfirm に渡す。上書き・別名・取りやめ）
    param (
        [string]$name
    )

    return "インデックス「${name}」は既にあります。上書きしますか？（前のインデックスは置き換わります。別名で入れることもできます）"
}

function getIndexRowView {
    # 一覧の「ステータス」列（本文の取り込みの状態）の文言・ツールヒント・色の区分を返す。@{ Text; ToolTip; Level }
    #   stat     : getIndexStats のそのインデックスの値（Total; Done; Pending; Failed）。無ければ $null
    #   indexing : インデックス作成中か
    #   enabled  : 一覧でチェックが付いているか（インデックス作成で取り込む対象か）
    param (
        $stat,
        [bool]$indexing = $false,
        [bool]$enabled = $true
    )

    $notice = if (!$enabled) { "チェックが外れているため、インデックス作成では取り込まない（インデックスは残っている）" } else { "" }
    if ($indexing -and $enabled) {
        return @{ Text = "作成中"; Level = "Wait"; ToolTip = "インデックス作成中。終わると状態を表示する" }
    }
    if ($null -eq $stat -or $stat.Total -eq 0) {
        return @{ Text = "未作成"; Level = "None"; ToolTip = addIndexRowNotice "まだ取り込んでいない。チェックを付けて［インデックス作成を開始］を押すと作る" $notice }
    }
    if ($stat.Pending -ge 1) {
        return @{ Text = "途中"; Level = "Wait"; ToolTip = addIndexRowNotice "未取り込み $($stat.Pending) 件。次の［インデックス作成を開始］で続きから取り込む" $notice }
    }
    if ($stat.Failed -ge 1) {
        return @{ Text = "一部失敗"; Level = "Ng"; ToolTip = addIndexRowNotice "失敗 $($stat.Failed) 件。原因は下の「取り込みに失敗したファイル」で見られる。失敗したファイル以外は検索できる" $notice }
    }
    return @{ Text = "取り込み済"; Level = "Ok"; ToolTip = addIndexRowNotice "取り込み済み $($stat.Total) 件" $notice }
}

function addIndexRowNotice {
    # getIndexRowView のツールヒントに、チェックが外れているときの案内を改行で足す（無ければそのまま）
    param ([string]$text, [string]$notice)

    if ($notice -eq "") {
        return $text
    }
    return "${text}`n${notice}"
}

function getFastSearchRowView {
    # 一覧の「高速検索」列（システムインデックスが Windows Search にどこまで反映されたか）の
    # 文言・ツールヒント・色の区分を返す。@{ Text; ToolTip; Level }
    #   reason     : getWindowsSearchState の値（NoFolder; NoConnection; NotInScope; NotYet; Ok）。$null ならまだ確かめていない
    #   progress   : getSystemIndexProgress の値（$null なら読めなかった・問い合わせに失敗した）
    #   name       : インデックス名（progress.ByIndex を引く）
    #   hasContent : そのインデックスの getIndexStats の Done が 1 以上か
    #   checkedAt  : 最後に確かめ終えた時刻（DateTime。$null ならまだ）
    #   indexing   : インデックス作成中か。作成中に hasContent が真へ変わっても（高速検索の確かめは
    #                作成中は走らないため）、前の確かめの結果（NoFolder など）から「不可」にしない
    param (
        $reason,
        $progress,
        [string]$name,
        [bool]$hasContent = $false,
        $checkedAt = $null,
        [bool]$indexing = $false
    )

    if ($null -eq $reason) {
        return @{ Text = "確認中…"; Level = "None"; ToolTip = "Windows Search の状態を確かめている" }
    }
    if ($reason -eq "NoFolder" -and !$hasContent) {
        return @{ Text = "－"; Level = "None"; ToolTip = (addFastSearchCheckedAt "まだ作っていない。インデックス作成が終わると状態を表示する" $checkedAt) }
    }
    if ($reason -eq "NoConnection") {
        return newFastSearchRowResult "不可" "Ng" @("Windows Search に接続できない。Windows Search のサービスが動いているかを確かめる") $true $checkedAt $progress
    }
    if ($reason -eq "NotInScope") {
        return newFastSearchRowResult "不可" "Ng" @("ワークスペースが Windows Search の索引の対象外。［インデックスのオプション］でワークスペースの system_index を対象に加える（管理者の権限が要る PC では、PC の管理者に頼む）") $true $checkedAt $progress
    }
    if (($reason -eq "Ok" -or $reason -eq "NotYet") -and $null -eq $progress) {
        return @{ Text = "－"; Level = "None"; ToolTip = (addFastSearchCheckedAt "反映の進み具合を確かめられなかった。画面を前に出し直すと、もう一度確かめる" $checkedAt) }
    }
    $entry = $null
    if ($progress -and $progress.ByIndex -and $progress.ByIndex.ContainsKey($name)) {
        $entry = $progress.ByIndex[$name]
    }
    if ($null -eq $entry -or $entry.Folders -eq 0) {
        if ($hasContent -and $indexing) {
            return @{ Text = "－"; Level = "None"; ToolTip = (addFastSearchCheckedAt "インデックス作成中。終わると状態を表示する" $checkedAt) }
        }
        if ($hasContent) {
            return newFastSearchRowResult "不可" "Ng" @("このインデックスには高速検索用のデータがありません。") $false $checkedAt $null
        }
        return @{ Text = "－"; Level = "None"; ToolTip = (addFastSearchCheckedAt "まだ作っていない。インデックス作成が終わると状態を表示する" $checkedAt) }
    }
    if ($reason -eq "NotYet") {
        return newFastSearchRowResult "反映待ち" "Wait" @("Windows Search がまだ索引していない。対象に入っていれば、待つと使えるようになる（対象外のときは［インデックスのオプション］で加える）") $true $checkedAt $progress
    }
    # ここから reason は Ok
    if ($entry.Waiting -eq $entry.Folders) {
        return newFastSearchRowResult "反映待ち" "Wait" @("反映済み 0 / $($entry.Folders) フォルダ。Windows Search が索引すると反映中 N% に進む") $true $checkedAt $progress
    }
    if ($entry.Waiting -ge 1) {
        $percent = [Math]::Floor((($entry.Folders - $entry.Waiting) / [double]$entry.Folders) * 100)
        if ($percent -eq 0) { $percent = 1 }
        return newFastSearchRowResult "反映中 ${percent}%" "Wait" @("反映済み $($entry.Folders - $entry.Waiting) / $($entry.Folders) フォルダ（反映待ち $($entry.Waiting)）。反映済みのフォルダは高速検索で、反映待ちのフォルダはふつうの検索で調べる") $true $checkedAt $progress
    }
    return @{
        Text = "可"; Level = "Ok"
        ToolTip = addFastSearchCheckedAt "反映済み $($entry.Folders) / $($entry.Folders) フォルダ。高速検索に使える" $checkedAt
    }
}

function newFastSearchRowResult {
    # getFastSearchRowView の共通の組み立て: 理由の行＋（使っても結果は同じという注記）＋ ContentIndexed の案内＋最終確認
    param ([string]$text, [string]$level, [string[]]$lines, [bool]$withUsageNotice, $checkedAt, $progress)

    $all = New-Object 'System.Collections.Generic.List[string]'
    foreach ($line in $lines) { [void]$all.Add($line) }
    if ($withUsageNotice) {
        [void]$all.Add("使えなくても検索の結果は同じで、時間だけが違う")
        if ($progress -and $progress.ContentIndexed) {
            [void]$all.Add("content_index が Windows Search の対象から自動で外れなかった。対象から外すと反映が早くなる")
        }
    }
    $tooltip = addFastSearchCheckedAt ($all -join "`n") $checkedAt
    return @{ Text = $text; Level = $level; ToolTip = $tooltip }
}

function addFastSearchCheckedAt {
    # ツールヒントの最後に「最終確認 HH:mm」を足す（checkedAt が無ければそのまま）
    param ([string]$text, $checkedAt)

    if ($null -eq $checkedAt) {
        return $text
    }
    return "${text}`n最終確認 $($checkedAt.ToString('HH:mm'))"
}

function testIndexEditInput {
    # 追加・編集の入力を調べ、直してほしい内容を返す（問題なければ空文字列）
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
