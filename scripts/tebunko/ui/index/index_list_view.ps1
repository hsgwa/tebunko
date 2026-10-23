# ［インデックス管理］タブの判断: 一覧・選択・ボタンの可否・状態の文言。
# 画面に触らないため、そのままテストできる（tests\tebunko\ui\index\index_list_view.Tests.ps1）。

function getUsedIndexNames {
    # 付けられない名前の集合（大文字・小文字を区別しない）。一覧のインデックス名（except に渡した行の名前は含めない）に、
    # 設定から外れたまま残してあるインデックスの名前のうち folder には付けられないもの（getBlockedIndexNames）を足す。
    # 名前を決める入口（追加の自動の名前・インポートの提案など）はすべてこの集合を使う。入口ごとに別の集合を作らない
    #   leftovers: getWorkspaceIndexFolders の結果（@{ Name; Path }）。省くと一覧の名前だけ
    #   folder   : 名前を付けようとするフォルダ。省くと、残っているものはすべて付けられない
    param (
        $items,
        $except = $null,
        $leftovers = $null,
        [string]$folder = ""
    )

    $used = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($item in $items) {
        if ($item -ne $except -and $item.Name) {
            [void]$used.Add($item.Name)
        }
    }
    $blocked = getBlockedIndexNames $items $leftovers $folder
    # .Keys と書くと、名前が Keys のキーがあるとき、キーの一覧でなくその値が返る。メソッドで読む
    foreach ($name in $blocked.get_Keys()) {
        [void]$used.Add($name)
    }
    # HashSet をそのまま return すると PowerShell が中身を展開してしまい（0 件なら $null、1 件なら文字列）、
    # 受け取った側の .Contains が落ちる・部分一致になる。, を付けて集合のまま返す
    return , $used
}

function getIndexNavBadge {
    # ナビの［インデックス管理］の横に出す小さな印。更新中は「● N%」（割合が分からないうちは「●」）、
    # 中断中は「中断」、更新に失敗したファイルがあるときは「失敗」。どれでもなければ出さない（Text が空）。
    # Kind は色の種類（Run = 更新中 / Warn = 注意）
    param (
        [bool]$indexing,
        [double]$ratio,     # 更新の進み（0～1。分からないときは負の値）
        [int]$pending,      # まだ更新していないファイルの数
        [int]$failed        # 更新に失敗したファイルの数
    )

    if ($indexing) {
        $text = if ($ratio -ge 0) { "● {0}%" -f [int][Math]::Floor([Math]::Min($ratio, 1.0) * 100) } else { "●" }
        return @{ Text = $text; Kind = "Run"; ToolTip = "インデックスを更新しています" }
    }
    if ($pending -gt 0) {
        return @{ Text = "中断"; Kind = "Warn"; ToolTip = "更新が途中で止まっています（残り ${pending} 件）" }
    }
    if ($failed -gt 0) {
        return @{ Text = "失敗"; Kind = "Warn"; ToolTip = "更新に失敗したファイルがあります" }
    }
    return @{ Text = ""; Kind = ""; ToolTip = "" }
}

function getIndexAddedStatus {
    # インデックスを追加したときのステータス。フォルダの有無によらず同じ文言にする（有無は一覧の列で分かる）
    param (
        [string]$name
    )

    return "インデックス [${name}] を追加しました。［すべて更新］を押すと中身を更新します"
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

function getIndexNamingStatus {
    # 名前の決まっていないインデックスに名前を付けている間（ネットワークのワークスペースでは裏で行う）のステータス
    return "インデックス名を決めています…（共有フォルダに接続できないときは、しばらくかかります）"
}

function getIndexNamingFailedStatus {
    # インデックス名を決められなかったときのステータス
    param (
        [string]$errorText
    )

    return "インデックス名を決められませんでした：${errorText}"
}

function getIndexTabButtonsEnabled {
    # 排他（getIndexJobBlocker の結果）と、一覧で選んでいる行の有無から、［インデックス管理］の各ボタンの可否を返す。
    #   New/Edit/Remove/ChangeFolder: ［＋ フォルダを追加］・詳細の名前とフォルダパスの欄（Edit。欄で直接直す）・［削除］・フォルダパスの［変更…］（Edit と同じ）/ Export/Import: ［エクスポート…］［インポート…］
    # 名前とフォルダパスの欄・［削除］［エクスポート…］は、1 件選んでいるときだけ有効
    # （［すべて更新］は updateIndexingButton が、［設定］の［変更…］は押したときに testWorkspaceChangeable が、
    # 同じ getIndexJobBlocker の結果で止める）
    param (
        [string]$blocker,
        [bool]$hasSelection
    )

    $free = ($blocker -eq "")
    return @{
        New = $free; Edit = ($free -and $hasSelection); Remove = ($free -and $hasSelection)
        ChangeFolder = ($free -and $hasSelection)
        Export = ($free -and $hasSelection); Import = $free
    }
}

function getIndexStatusHelpText {
    # 一覧の「ステータス」列の見出しのツールヒント（状態の名前と意味。行のバッジの名前と同じ言葉を使う）
    return @(
        "検索の準備の状態です。"
        "・最新: 検索できます"
        "・要更新: 前回の更新のあとにファイルが変わっています"
        "・更新中: 更新しています"
        "・警告: 読み取れないファイルがあります（それ以外のファイルは検索できます）"
        "・エラー: フォルダが見つかりません"
        "・未作成: まだ更新していません"
    )
}

function getIndexCheckedItems {
    # 一覧でチェックを付けている行（［アクション ▼］の対象・「N 件を選択中」の数・全選択の状態の元）。
    # チェックは一時の選択で、保存しない（初めは全部外れ。設定の enabled とは別のもの）
    param (
        [object[]]$items
    )

    return @($items | Where-Object { $null -ne $_ -and $_.Checked })
}

function getIndexSelectionView {
    # 一覧の見出しの全選択と、見出しの横の「N / M 件を選択中」。@{ CountText; AllChecked }
    #   AllChecked: すべて選んでいれば $true、1 件も選んでいなければ $false、一部だけなら $null（横棒）
    param (
        [int]$total,    # 一覧の件数
        [int]$checked   # チェックを付けている件数
    )

    $countText = if ($checked -ge 1) { "{0:#,0} / {1:#,0} 件を選択中" -f $checked, $total } else { "" }
    $all = if ($total -eq 0 -or $checked -eq 0) { $false } elseif ($checked -ge $total) { $true } else { $null }
    return @{ CountText = $countText; AllChecked = $all }
}

function getIndexActionsEnabled {
    # ［アクション ▼］のメニューの項目の可否。@{ Update; Export; Import; Delete }
    # ［インポート…］は選ばなくても使える。［更新］［エクスポート…］［削除…］は 1 件以上チェックしていると使える（2 件以上はまとめて行う）。
    # 動いている処理があれば（blocker は getIndexJobBlocker の結果）、すべて使えない
    param (
        [string]$blocker,
        [int]$checked
    )

    $free = ($blocker -eq "")
    return @{
        Update = ($free -and $checked -ge 1)
        Export = ($free -and $checked -ge 1)
        Import = $free
        Delete = ($free -and $checked -ge 1)
    }
}

function getIndexRowMenuEnabled {
    # 行の右クリックのメニューの項目の可否。@{ Update; OpenFolder; Export; Delete }。対象は押した行（右クリックした行を選んでから開く）
    #   Update: 行の右端の［更新］と同じ判断（getIndexRowActions）。更新中の行・動いている処理があるときは使えない（失敗のある警告の行は使える）
    #   Export/Delete: 1 件に効く判断（getIndexTabButtonsEnabled）。OpenFolder: 動いている処理があっても使える（元のフォルダの場所があれば）
    # 行が無いとき（hasRow が false）は、すべて使えない
    param (
        [string]$blocker,
        [bool]$hasRow,
        [string]$level,       # 押した行の getIndexRowView の Level
        [bool]$hasPath        # 押した行の元のフォルダの場所があるか
    )

    $rowAction = getIndexRowActions $level $blocker
    $buttons = getIndexTabButtonsEnabled $blocker $hasRow
    return @{
        Update = ($hasRow -and $rowAction.Action -eq "Update" -and $rowAction.UpdateEnabled)
        OpenFolder = ($hasRow -and $hasPath)
        Export = $buttons.Export
        Delete = $buttons.Remove
    }
}

function getIndexSourceFolderStatus {
    # ［元のフォルダを開く］で、フォルダの場所を確かめた結果のステータス。kind は Empty（記録が無い）/ Checking / Missing / Unreachable / Other
    param (
        [string]$kind,
        [string]$path,
        [string]$message = ""
    )

    switch ($kind) {
        "Empty" { return "元のフォルダの場所が分かりません（元のフォルダの記録がありません）" }
        "Checking" { return "元のフォルダを確かめています…：${path}" }
        "Missing" { return "元のフォルダが見つかりません（移動・削除された可能性があります）：${path}" }
        "Unreachable" { return getFailedFileUnreachableStatus $path }
        default { return "元のフォルダを確かめられませんでした：${message}" }
    }
}

function getIndexSourceFolderKind {
    # ［元のフォルダを開く］で getPathState が返した結果を、次にすることに分ける。
    # Open（フォルダが見つかった。エクスプローラーで開く）/ Unreachable / Other / Missing（見つからない。ファイルだったときも）。
    # Open 以外は、そのまま getIndexSourceFolderStatus の kind に渡す
    param (
        $state
    )

    if ($state.State -eq ${pathStateFound}) {
        return $(if ($state.IsDirectory) { "Open" } else { "Missing" })
    }
    switch ($state.State) {
        ${pathStateUnreachable} { return "Unreachable" }
        ${pathStateOther} { return "Other" }
        default { return "Missing" }
    }
}

function getIndexRowActions {
    # 一覧の行の右端のボタン。@{ Action（Update / Stop / None）; UpdateEnabled }
    # 更新中の行には［中止］（回ごと止める）、エラー（Ng）の行には何も出さず（場所は空けておく）、ほかの行（失敗のある警告の行を含む）には［更新］を出す。
    # ［更新］は、動いている処理があれば（blocker は getIndexJobBlocker の結果）使えない（更新中はほかの行も押せない）
    param (
        [string]$level,           # getIndexRowView の Level
        [string]$blocker,
        [bool]$exporting = $false # エクスポート中の行か（止められないので、［中止］も［更新］も出さない）
    )

    if ($exporting) {
        return @{ Action = "None"; UpdateEnabled = $false }
    }
    $action = switch ($level) { "Run" { "Stop" } "Ng" { "None" } default { "Update" } }
    return @{ Action = $action; UpdateEnabled = ($blocker -eq "") }
}

function selectFailedRowsForIndex {
    # 「更新に失敗したファイル」に出す行を、詳細に出しているインデックスの分だけにする。
    # indexName が空（行を選んでいない・2 件以上にチェックしている）ときは、全部をそのまま返す。
    # 失敗の行の相対パス（work\index からの相対）の先頭の名前が、indexName と同じ（大文字・小文字は区別しない）ものだけを残す。
    # 前方一致では絞らない（「営業」で「営業部」の行を拾わない）
    param (
        [object[]]$failedRows,  # getIndexingState の FailedRows（相対パスを持つ行）
        [string]$indexName
    )

    $rows = @($failedRows | Where-Object { $null -ne $_ })
    if (!$indexName) { return $rows }
    $prefix = $indexName + "\"
    return @($rows | Where-Object { ([string]$_.相対パス).StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase) })
}

function getIndexRowView {
    # 一覧の「ステータス」列（本文の更新の状態）の文言・補足・ツールヒント・色の区分を返す。@{ Text; Sub; ToolTip; Level }
    #   Text は状態の名前だけ（数字を含めない）。Sub は、バッジの右に小さく出す補足（更新中の「45%」、途中で止まったときの「残り N 件」。無ければ空文字列）
    #   stat     : getIndexStats のそのインデックスの値（Total; Done; Pending; Failed）。無ければ $null
    #   indexing : インデックス作成中か
    #   enabled  : 設定の enabled（［すべて更新］で更新する対象か。画面にチェックは出さない）
    #   ratio    : 更新全体の進み（0～1。分からないときは負の値）。更新中の補足に「45%」と出し、棒の長さにする
    #   inRound  : 今の回で更新するインデックスか（選んだものだけの回では、選ばなかった行は $false）
    #   exporting: エクスポートで書き出している最中の行か（バッジは「エクスポート中」。更新中と同じ Run の色。棒は出さない）
    #   formatEntry: getIndexFormatState のそのインデックスの値（State; Searchable; Message）。無ければ $null。
    #                古い形式・作り直しの途中・新しい版のときは、本文の状態の代わりにその旨を出す
    # 返す値の Percent は棒の長さ（0～100）。更新中でなければ 0
    param (
        $stat,
        [bool]$indexing = $false,
        [bool]$enabled = $true,
        [double]$ratio = -1.0,
        [bool]$inRound = $true,
        [bool]$exporting = $false,
        $formatEntry = $null
    )

    if ($exporting) {
        return @{ Text = "エクスポート中"; Sub = ""; Level = "Run"; Percent = 0; ToolTip = "インデックスを zip に書き出しています。終わるまで、ほかの操作はできません。" }
    }
    $notice = if (!$enabled) { "設定で［すべて更新］の対象から外れています（インデックスは残っています。行の［更新］で更新できます）。" } else { "" }
    if ($indexing -and $enabled -and $inRound) {
        $percent = if ($ratio -ge 0) { [int][Math]::Floor([Math]::Min($ratio, 1.0) * 100) } else { 0 }
        $sub = if ($ratio -ge 0) { "${percent}%" } else { "" }
        return @{ Text = "更新中"; Sub = $sub; Level = "Run"; Percent = $percent; ToolTip = "インデックスを更新しています。終わると状態を表示します。" }
    }
    if ($null -ne $formatEntry -and [string]$formatEntry.Message -ne "") {
        $formatView = @{
            ${indexStateOld} = @{ Text = "古い形式"; Level = "Warn" }
            ${indexStateRebuilding} = @{ Text = "作り直し中"; Level = "Wait" }
            ${indexStateNewer} = @{ Text = "新しい版"; Level = "Ng" }
        }
        $entryView = $formatView[[string]$formatEntry.State]
        if ($null -ne $entryView) {
            return @{ Text = $entryView.Text; Sub = ""; Level = $entryView.Level; ToolTip = addIndexRowNotice $formatEntry.Message $notice }
        }
    }
    if ($null -eq $stat -or $stat.Total -eq 0) {
        return @{ Text = "未作成"; Sub = ""; Level = "None"; ToolTip = addIndexRowNotice "まだ更新していません。［すべて更新］か、行の［更新］で作ります。" $notice }
    }
    if ($stat.Pending -ge 1) {
        return @{ Text = "要更新"; Sub = "残り $($stat.Pending) 件"; Level = "Wait"; ToolTip = addIndexRowNotice "未更新 $($stat.Pending) 件です。次の更新で続きから更新します。" $notice }
    }
    if ($stat.Failed -ge 1) {
        return @{ Text = "警告"; Sub = ""; Level = "Warn"; ToolTip = addIndexRowNotice "失敗 $($stat.Failed) 件です。原因は下の「更新に失敗したファイル」で見られます。失敗したファイル以外は検索できます。" $notice }
    }
    return @{ Text = "最新"; Sub = ""; Level = "Ok"; ToolTip = addIndexRowNotice "更新済み $($stat.Total) 件です。" $notice }
}

function getIndexFooterView {
    # 一覧の下の帯の文言。@{ Folders; Files }
    param (
        [int]$folderCount,
        [int]$fileTotal
    )

    return @{
        Folders = "登録済み: {0:#,0} フォルダ" -f $folderCount
        Files = "合計: {0:#,0} ファイル" -f $fileTotal
    }
}

function addIndexRowNotice {
    # getIndexRowView のツールヒントに、［すべて更新］の対象から外れているときの案内を改行で足す（無ければそのまま）
    param ([string]$text, [string]$notice)

    if ($notice -eq "") {
        return $text
    }
    return "${text}`n${notice}"
}

function getIndexLocationView {
    # インデックスの保存先（ワークスペースのフォルダ）の表示。@{ Heading; Path; Note; Warning }
    # ［インデックス管理］の一覧の上と、インデックス更新の確認で同じものを使う。
    # Note は常に出す注意（本文がそのまま入る・アクセス権は引き継がれない・［設定］で変えられる）。
    # Warning は保存先がネットワーク（共有）のときだけ（無ければ空文字）。ネットワークかどうかは画面層が testNetworkPath で求めて渡す
    param (
        [string]$workDir,
        [bool]$isNetwork
    )

    $warning = ""
    if ($isNetwork) {
        $warning = "注意：保存先は共有フォルダにあります。この共有を読める人は、元のファイルを開く権限が無くても本文を読めます。読める人が元のフォルダと同じか、それより少ないか確かめてください。"
    }
    return @{
        Heading = "インデックスの保存先"
        Path = $workDir
        Note = "インデックスには元のファイルの本文がそのまま入り、元のファイルのアクセス権は引き継がれません。元のフォルダと同じかそれより厳しいアクセス権の場所に置いてください。会社の決まりで置いてよい場所があるときは、［設定］でその場所に変えられます。"
        Warning = $warning
    }
}

function getIndexLocationOpenStatus {
    # ［フォルダを開く］（保存先）で、場所を確かめた結果の文言。kind は getIndexSourceFolderKind の値のうち Missing / Unreachable / Other
    param (
        [string]$kind,
        [string]$path,
        [string]$message = ""
    )

    switch ($kind) {
        "Checking" { return "保存先のフォルダを確かめています…：${path}" }
        "Missing" { return "保存先のフォルダはまだありません（インデックスを作ると作られます）：${path}" }
        "Unreachable" { return "保存先のフォルダに接続できません：${path}" }
        default { return "保存先のフォルダを開けません：${path}（${message}）" }
    }
}