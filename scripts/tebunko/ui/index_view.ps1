# ［インデックス管理］タブの判断（入力の検査・名前の重複）。
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

function getIndexNavBadge {
    # ナビの［インデックス管理］の横に出す小さな印。更新中は「● N%」（割合が分からないうちは「●」）、
    # 中断中は「中断」、更新に失敗したファイルがあるときは「失敗」。どれでもなければ出さない（Text が空）。
    # Kind は色の種類（Run = 更新中 / Warn = 注意）
    param (
        [bool]$indexing,
        [double]$ratio,     # 更新の進み（0〜1。分からないときは負の値）
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
        [bool]$archiveBusy   # エクスポート・インポート中
    )

    if ($isIndexing) {
        return "インデックス作成中"
    }
    if ($indexBusy) {
        return "削除・名前の変更中"
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
            return "更新中はワークスペースを変えられません。更新が終わるまでお待ちください（［中止］で止められます）。"
        }
        if ($blocker -eq "削除・名前の変更中") {
            return "前のインデックスの削除・名前の変更が終わるまでお待ちください。"
        }
        return "エクスポート・インポートが終わるまでお待ちください。"
    }
    if ($blocker -eq "インデックス作成中") {
        return "更新中はインデックスを${operation}できません。更新が終わるまでお待ちください（［中止］で止められます）。"
    }
    return "${blocker}は${operation}できません。終わるまでお待ちください。"
}

function getIndexTabButtonsEnabled {
    # 排他（getIndexJobBlocker の結果）と、一覧で選んでいる行の有無から、［インデックス管理］の各ボタンの可否を返す。
    #   New/Edit/Remove/ChangeFolder: ［＋ フォルダを追加］［編集…］［削除］・詳細のフォルダパスの［...］（［編集…］と同じ）/ Export/Import: ［エクスポート…］［インポート…］
    # ［編集…］［削除］［エクスポート…］は、1 件選んでいるときだけ有効
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

function getIndexScreenInfoText {
    # 頭の ⓘ に出す、この画面の説明（3 行。画面は改行でつないでツールヒントにする）
    return @(
        "検索したいフォルダを登録する画面です。"
        "登録したフォルダは、［すべて更新］を押すと検索できるようになります。"
        "インデックスは、フォルダの中身を読み取って作る検索用のデータです。"
    )
}

function getIndexStatusHelpText {
    # 一覧の「ステータス」列の見出しのツールヒント（状態の名前と意味。行のバッジの名前と同じ言葉を使う）
    return @(
        "検索の準備の状態です。"
        "・最新: 検索できます"
        "・要更新: 前回の更新のあとにファイルが変わっています"
        "・更新中: 更新しています"
        "・エラー: フォルダが見つからないか、読み込めないファイルがあります"
        "・未作成: まだ更新していません"
    )
}

function getIndexCheckedItems {
    # 一覧でチェックを付けている行（［アクション ▾］の対象・「N 件を選択中」の数・全選択の状態の元）。
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
    # ［アクション ▾］のメニューの項目の可否。@{ Update; Edit; Export; Import; Delete }
    # ［インポート…］は選ばなくても使える。［更新］［エクスポート…］［削除…］は 1 件以上チェックしていると使える（2 件以上はまとめて行う）。
    # ［編集…］は 1 件にしか効かないので、チェックが 1 件ならその行、チェックが無ければ押した行（hasPressedRow）があるときだけ使える。
    # チェックが 2 件以上のときは使えない（詳細も「N 件を選択中」になり、どの行か決まらないため）。
    # 動いている処理があれば（blocker は getIndexJobBlocker の結果）、すべて使えない
    param (
        [string]$blocker,
        [int]$checked,
        [bool]$hasPressedRow = $false   # 一覧で押した行があるか
    )

    $free = ($blocker -eq "")
    return @{
        Update = ($free -and $checked -ge 1)
        Edit = ($free -and ($checked -eq 1 -or ($checked -eq 0 -and $hasPressedRow)))
        Export = ($free -and $checked -ge 1)
        Import = $free
        Delete = ($free -and $checked -ge 1)
    }
}

function getIndexRowMenuEnabled {
    # 行の右クリックのメニューの項目の可否。@{ Update; Edit; OpenFolder; Export; Delete }。対象は押した行（右クリックした行を選んでから開く）
    #   Edit: ［アクション ▾］の［編集…］と同じ判断（getIndexActionsEnabled。チェックが無く押した行がある場合）
    #   Update: 行の右端の［更新］と同じ判断（getIndexRowActions）。更新中の行・エラーの行・動いている処理があるときは使えない
    #   Export/Delete: 1 件に効く判断（getIndexTabButtonsEnabled）。OpenFolder: 動いている処理があっても使える（元のフォルダの場所があれば）
    # 行が無いとき（hasRow が false）は、すべて使えない
    param (
        [string]$blocker,
        [bool]$hasRow,
        [string]$level,       # 押した行の getIndexRowView の Level
        [bool]$hasPath        # 押した行の元のフォルダの場所があるか
    )

    $rowAction = getIndexRowActions $level $blocker
    $actions = getIndexActionsEnabled $blocker 0 $hasRow
    $buttons = getIndexTabButtonsEnabled $blocker $hasRow
    return @{
        Update = ($hasRow -and $rowAction.Action -eq "Update" -and $rowAction.UpdateEnabled)
        Edit = $actions.Edit
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
    # 更新中の行には［中止］（回ごと止める）、エラーの行には何も出さず（場所は空けておく）、ほかの行には［更新］を出す。
    # ［更新］は、動いている処理があれば（blocker は getIndexJobBlocker の結果）使えない（更新中はほかの行も押せない）
    param (
        [string]$level,           # getIndexRowView の Level
        [string]$blocker
    )

    $action = switch ($level) { "Run" { "Stop" } "Ng" { "None" } default { "Update" } }
    return @{ Action = $action; UpdateEnabled = ($blocker -eq "") }
}

function getIndexDetailItem {
    # 詳細欄に出す行（右クリック・二重クリックの［編集…］と［...］、右クリックの［削除］［エクスポート…］の対象も同じ行。［アクション ▾］は別で、［更新］［エクスポート…］［削除…］はチェックの行、［編集…］はチェックが 1 件ならその行・無ければ押した行が対象で、［インポート…］はチェックが要らない。可否は getIndexActionsEnabled が決める）。押した行があればそれ、
    # 無くてチェックが 1 件だけならその行、それ以外は $null。詳細を出す所と可否・対象を決める所で、この 1 つを使う
    param (
        [object]$selectedItem,   # 一覧で押した行（無ければ $null）
        [object[]]$checkedItems  # getIndexCheckedItems の結果
    )

    if ($null -ne $selectedItem) { return $selectedItem }
    $checked = @($checkedItems | Where-Object { $null -ne $_ })
    if ($checked.Count -eq 1) { return $checked[0] }
    return $null
}

function getIndexDetailMultiCount {
    # 詳細欄を「N 件を選択中」にするときの N（そうしないときは 0）。2 件以上チェックしているときだけ、その件数を返す。
    # 0 件・1 件のときは 0（詳細欄は押した行のまま）
    param (
        [int]$checked
    )

    if ($checked -ge 2) { return $checked }
    return 0
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

function getImportResultStatus {
    # インポートの結果（importIndex の戻り値）から、ステータスに出す文言を返す。
    # Warnings（同じ元のフォルダが別の名前で既に登録されている等）と、高速検索が次の更新の後に効くことを添える
    param (
        $result
    )

    $text = "インデックス [$($result.Name)] をインポートしました（$($result.Files) ファイル）。" +
        "高速検索は次の更新の後に効きます。"
    foreach ($warning in @($result.Warnings)) {
        $text += " ${warning}"
    }
    return $text
}

function testIndexExportInput {
    # エクスポートのダイアログの入力を調べ、直してほしい内容を返す（問題なければ空文字列）。
    # 書き出し先のフォルダが今あるかは、ここでは調べない（画面のスレッドでネットワークのパスを調べないため。書き出す仕事の中で調べる）
    param (
        [string]$folder   # 入力された書き出し先のフォルダ
    )

    if ((normalizeFolderPath $folder) -eq "") {
        return "書き出し先のフォルダを指定してください。"
    }
    return ""
}

function getIndexExportNotice {
    # エクスポートのダイアログの案内文（対象の名前・大きさは getIndexExportHeader の枠に出す）
    return "別の PC・ワークスペースでインポートして使えます。"
}

function formatArchiveSize {
    # 大きさを "94.0 MB" の形にする（0.1 MB に満たなくても 0.1 MB。区切りの記号は環境によらず "."）
    param ([long]$bytes)

    $mb = [Math]::Max(0.1, [Math]::Round($bytes / 1MB, 1))
    return ($mb.ToString("0.0", [System.Globalization.CultureInfo]::InvariantCulture) + " MB")
}

function getIndexExportHeader {
    # エクスポートのダイアログの上の枠（対象・大きさ・zip の名前）と注意。
    #   items: 書き出すインデックス（Name と FileCountText（"2,530"。無ければ "－" か空）を持つ行）。1 件でも 2 件以上でもよい
    # 戻り値: @{ Title（1 件は名前、2 件以上は「N 件のインデックス」）; Size（「2,530 ファイル」。数が分からなければ空。大きさは数えていないので出さない）;
    #            FileText（書き出し先の下の「<名前>_インデックス_<yyyyMMdd>.zip として保存します」）; Caution }
    param (
        [object[]]$items,
        [datetime]$now = (Get-Date)
    )

    $items = @($items | Where-Object { $null -ne $_ })
    $total = 0
    $known = $false
    foreach ($item in $items) {
        $parsed = 0
        if ([int]::TryParse((([string]$item.FileCountText) -replace ",", ""), [ref]$parsed)) {
            $total += $parsed
            $known = $true
        }
    }
    $size = if ($known) { "{0:#,0} ファイル" -f $total } else { "" }
    if ($items.Count -eq 1) {
        $title = [string]$items[0].Name
        $fileText = "$(getExportFileName $title @() $now) として保存します"
    } else {
        $title = "{0:#,0} 件のインデックス" -f $items.Count
        $fileText = "インデックスごとに <名前>_インデックス_$($now.ToString('yyyyMMdd')).zip として保存します"
    }
    return @{
        Title = $title; Size = $size; FileText = $fileText
        Caution = "このファイルには、元のファイルの本文と元のフォルダの場所が含まれます。元のファイルと同じように取り扱ってください。"
    }
}

function getIndexImportHeader {
    # インポートのダイアログの上の枠（zip の中身）。
    # 戻り値: @{ Name; Exported（「エクスポートした日時 2026/09/27 10:00」。分からなければ空）; Size（「2,530 ファイル・56.6 MB・v0.3.0」。版が無ければ版は省く） }
    param (
        $info   # readIndexArchiveInfo の結果
    )

    $exported = ""
    $at = [datetimeoffset]::MinValue
    if ($info.ExportedAt -and [datetimeoffset]::TryParse([string]$info.ExportedAt, [ref]$at)) {
        $exported = "エクスポートした日時 " + $at.LocalDateTime.ToString("yyyy/MM/dd HH:mm", [System.Globalization.CultureInfo]::InvariantCulture)
    }
    $parts = New-Object 'System.Collections.Generic.List[string]'
    [void]$parts.Add(("{0:#,0} ファイル" -f [int]$info.Files))
    [void]$parts.Add((formatArchiveSize ([long]$info.Bytes)))
    $version = ([string]$info.AppVersion).Trim().TrimStart("v", "V")
    if ($version) {
        [void]$parts.Add("v${version}")
    }
    return @{ Name = [string]$info.IndexName; Exported = $exported; Size = ($parts -join "・") }
}

function getImportSourceCheckText {
    # インポートの「元のフォルダ」の下に出す 1 行（文言と色の区分）。@{ Text; Kind }。Kind は Ok / Warn / Ng / None（None は何も出さない）。
    #   folder : 入力された元のフォルダ
    #   exists : フォルダがあるか（$true / $false。調べていない・分からないときは $null）。画面のスレッドでは調べない
    #   others : 重なりを比べる相手（Name・Path を持つ行。取り込む名前と同じ名前の行は呼び出し側で外す）
    param (
        [string]$folder,
        $exists = $null,
        $others = @()
    )

    $normalized = normalizeFolderPath $folder
    if ($normalized -eq "") {
        return @{ Text = ""; Kind = "None" }
    }
    foreach ($other in @($others)) {
        if ((testSameFolder $other.Path $normalized) -or (testFolderUnder $normalized $other.Path) -or (testFolderUnder $other.Path $normalized)) {
            return @{ Text = "「$($other.Name)」の元のフォルダと重なっています"; Kind = "Ng" }
        }
    }
    if ($exists -eq $true) {
        return @{ Text = "見つかりました"; Kind = "Ok" }
    }
    if ($exists -eq $false) {
        return @{ Text = "元のフォルダがこの PC に見つかりません。検索はできますが、元のフォルダを設定するまで更新できません。"; Kind = "Warn" }
    }
    return @{ Text = ""; Kind = "None" }
}

function getImportNameNoticeText {
    # インポートの名前の欄に、今の一覧にある名前を入れたときの案内（無ければ空）。確認は［インポート］を押したあと（getIndexImportOverwriteConfirmMessage）
    param (
        [string]$name,
        $usedNames = $null
    )

    if ($name.Trim() -ne "" -and (testImportNameCollision $name.Trim() $usedNames)) {
        return "同じ名前のインデックスがあります。［インポート］を押すと、上書きするか別の名前にするかを選べます。"
    }
    return ""
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
    # インポートのダイアログの説明
    # 中身（名前・日時・ファイル数・大きさ・版）は getIndexImportHeader の枠に出す
    return "名前と、元のフォルダの場所を変えられます。同じ名前のインデックスが既にあるときは、後で上書き・別名・取りやめを選べます。"
}

function getIndexImportOverwriteConfirmMessage {
    # 名前が既にあるインデックスと重なったときの確認（choices は showConfirm に渡す。上書き・別名・取りやめ）
    param (
        [string]$name
    )

    return "インデックス「${name}」は既にあります。上書きしますか？（前のインデックスは置き換わります。別名で入れることもできます）"
}

function getIndexRowView {
    # 一覧の「ステータス」列（本文の更新の状態）の文言・補足・ツールヒント・色の区分を返す。@{ Text; Sub; ToolTip; Level }
    #   Text は状態の名前だけ（数字を含めない）。Sub は、バッジの右に小さく出す補足（更新中の「45%」、途中で止まったときの「残り N 件」。無ければ空文字列）
    #   stat     : getIndexStats のそのインデックスの値（Total; Done; Pending; Failed）。無ければ $null
    #   indexing : インデックス作成中か
    #   enabled  : 設定の enabled（［すべて更新］で更新する対象か。画面にチェックは出さない）
    #   ratio    : 更新全体の進み（0〜1。分からないときは負の値）。更新中の補足に「45%」と出し、棒の長さにする
    #   inRound  : 今の回で更新するインデックスか（選んだものだけの回では、選ばなかった行は $false）
    # 返す値の Percent は棒の長さ（0〜100）。更新中でなければ 0
    param (
        $stat,
        [bool]$indexing = $false,
        [bool]$enabled = $true,
        [double]$ratio = -1.0,
        [bool]$inRound = $true
    )

    $notice = if (!$enabled) { "設定で［すべて更新］の対象から外れています（インデックスは残っています。行の［更新］で更新できます）。" } else { "" }
    if ($indexing -and $enabled -and $inRound) {
        $percent = if ($ratio -ge 0) { [int][Math]::Floor([Math]::Min($ratio, 1.0) * 100) } else { 0 }
        $sub = if ($ratio -ge 0) { "${percent}%" } else { "" }
        return @{ Text = "更新中"; Sub = $sub; Level = "Run"; Percent = $percent; ToolTip = "インデックスを更新しています。終わると状態を表示します。" }
    }
    if ($null -eq $stat -or $stat.Total -eq 0) {
        return @{ Text = "未作成"; Sub = ""; Level = "None"; ToolTip = addIndexRowNotice "まだ更新していません。［すべて更新］か、行の［更新］で作ります。" $notice }
    }
    if ($stat.Pending -ge 1) {
        return @{ Text = "要更新"; Sub = "残り $($stat.Pending) 件"; Level = "Wait"; ToolTip = addIndexRowNotice "未更新 $($stat.Pending) 件です。次の更新で続きから更新します。" $notice }
    }
    if ($stat.Failed -ge 1) {
        return @{ Text = "エラー"; Sub = ""; Level = "Ng"; ToolTip = addIndexRowNotice "失敗 $($stat.Failed) 件です。原因は下の「更新に失敗したファイル」で見られます。失敗したファイル以外は検索できます。" $notice }
    }
    return @{ Text = "最新"; Sub = ""; Level = "Ok"; ToolTip = addIndexRowNotice "更新済み $($stat.Total) 件です。" $notice }
}

function getIndexDetailRowPlan {
    # 詳細の行の高さをどうするか。1 件も無いとき（empty）は詳細と境目を畳み（Hide）、1 件以上になったら戻す（Show）。
    # 空かどうかが変わっていないとき（別の行を選んだだけ）は何もしない（Keep）。境目をドラッグして変えた高さを、選び直しで戻さないため。
    # wasEmpty は前回の空かどうか（まだ一度も決めていなければ $null）。
    param (
        $wasEmpty,
        [bool]$empty
    )

    if ($null -ne $wasEmpty -and [bool]$wasEmpty -eq $empty) { return "Keep" }
    if ($empty) { return "Hide" }
    return "Show"
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
        return @{ Text = "確認中…"; Level = "None"; ToolTip = "Windows Search の状態を確かめています。" }
    }
    if ($reason -eq "NoFolder" -and !$hasContent) {
        return @{ Text = "－"; Level = "None"; ToolTip = (addFastSearchCheckedAt "まだ作っていません。更新が終わると状態を表示します。" $checkedAt) }
    }
    if ($reason -eq "NoConnection") {
        return newFastSearchRowResult "不可" "Ng" @("Windows Search に接続できません。Windows Search のサービスが動いているかを確かめてください。") $true $checkedAt $progress
    }
    if ($reason -eq "NotInScope") {
        return newFastSearchRowResult "不可" "Ng" @("ワークスペースが Windows Search の索引の対象外です。［インデックスのオプション］でワークスペースの system_index を対象に加えてください（管理者の権限が要る PC では、PC の管理者に頼んでください）。") $true $checkedAt $progress
    }
    if (($reason -eq "Ok" -or $reason -eq "NotYet") -and $null -eq $progress) {
        return @{ Text = "－"; Level = "None"; ToolTip = (addFastSearchCheckedAt "反映の進み具合を確かめられませんでした。画面を前に出し直すと、もう一度確かめます。" $checkedAt) }
    }
    $entry = $null
    if ($progress -and $progress.ByIndex -and $progress.ByIndex.ContainsKey($name)) {
        $entry = $progress.ByIndex[$name]
    }
    if ($null -eq $entry -or $entry.Folders -eq 0) {
        if ($hasContent -and $indexing) {
            return @{ Text = "－"; Level = "None"; ToolTip = (addFastSearchCheckedAt "更新中です。終わると状態を表示します。" $checkedAt) }
        }
        if ($hasContent) {
            return newFastSearchRowResult "不可" "Ng" @("このインデックスには高速検索用のデータがありません。") $false $checkedAt $null
        }
        return @{ Text = "－"; Level = "None"; ToolTip = (addFastSearchCheckedAt "まだ作っていません。更新が終わると状態を表示します。" $checkedAt) }
    }
    if ($reason -eq "NotYet") {
        return newFastSearchRowResult "反映待ち" "Wait" @("Windows Search がまだ索引していません。対象に入っていれば、待つと使えるようになります（対象外のときは［インデックスのオプション］で加えてください）。") $true $checkedAt $progress
    }
    # ここから reason は Ok
    if ($entry.Waiting -eq $entry.Folders) {
        return newFastSearchRowResult "反映待ち" "Wait" @("反映済み 0 / $($entry.Folders) フォルダです。Windows Search が索引すると反映中 N% に進みます。") $true $checkedAt $progress
    }
    if ($entry.Waiting -ge 1) {
        $percent = [Math]::Floor((($entry.Folders - $entry.Waiting) / [double]$entry.Folders) * 100)
        if ($percent -eq 0) { $percent = 1 }
        return newFastSearchRowResult "反映中 ${percent}%" "Wait" @("反映済み $($entry.Folders - $entry.Waiting) / $($entry.Folders) フォルダです（反映待ち $($entry.Waiting)）。反映済みのフォルダは高速検索で、反映待ちのフォルダはふつうの検索で調べます。") $true $checkedAt $progress
    }
    return @{
        Text = "可"; Level = "Ok"
        ToolTip = addFastSearchCheckedAt "反映済み $($entry.Folders) / $($entry.Folders) フォルダです。高速検索に使えます。" $checkedAt
    }
}

function newFastSearchRowResult {
    # getFastSearchRowView の共通の組み立て: 理由の行＋（使っても結果は同じという注記）＋ ContentIndexed の案内＋最終確認
    param ([string]$text, [string]$level, [string[]]$lines, [bool]$withUsageNotice, $checkedAt, $progress)

    $all = New-Object 'System.Collections.Generic.List[string]'
    foreach ($line in $lines) { [void]$all.Add($line) }
    if ($withUsageNotice) {
        [void]$all.Add("使えなくても検索の結果は同じで、時間だけが違います。")
        if ($progress -and $progress.ContentIndexed) {
            [void]$all.Add("content_index が Windows Search の対象から自動で外れませんでした。対象から外すと反映が早くなります。")
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

function getIndexDetailView {
    # インデックスの詳細のパネルに出す内容を返す。1 つ選んでいるときだけ中身を出す
    # （何も選んでいない・複数のときは見出しだけ。一覧の下の合計は別の部品）。
    #   items         : 選んでいる行。@{ Name; Path; Enabled; FolderStatus; FolderMissing（元のフォルダが無いと分かったか）; IndexText; IndexLevel; IndexSub; FileCountText;
    #                   LastIngestedText; FastText; FastLevel; FastToolTip; FastCheckedText }
    #   fastEntry     : getSystemIndexProgress の ByIndex のそのインデックスの値（@{ Folders; Waiting }）。無ければ $null
    #   multiCount    : 2 以上なら、詳細は出さず「N 件を選択中」と案内だけにする（getIndexDetailMultiCount）
    #   running       : 更新の進み具合（@{ Ratio; Processed; Total; Failed; Eta; Current }）。更新中の行の「インデックス」の箱に出す。無ければ $null
    # 戻り値: @{ Title; Selected; Name; Path; FolderStatus; Badge（@{ Text; Level }）; Updated; Count;
    #            Rows（右の「インデックス情報」の @{ Label; Value } の配列）; RowsKey;
    #            Fast（@{ Shown（反映の進みの棒を出すか）; Value（0〜1）; Text; State; Level; Reason; Checked; Note }）;
    #            Multi（複数を選んでいるか）; Hint（Multi のときの案内）;
    #            Run（@{ Shown（更新中の棒・件数・ファイル名を出すか）; Value（0〜1）; CountText; FileText }）;
    #            Error・ErrorHint（元のフォルダが無いときの理由と直し方。無ければ空。バッジは「エラー」の Ng になる） }
    #   RowsKey は左右の表の中身から作る文字列。同じなら画面は行を置き直さない（高速検索の進みだけが変わったとき）
    param (
        [object[]]$items,
        $fastEntry = $null,
        [int]$multiCount = 0,
        $running = $null
    )

    $noRun = @{ Shown = $false; Value = 0.0; CountText = ""; FileText = "" }
    $emptyFast = @{ Shown = $false; Value = 0.0; Text = ""; State = ""; Level = "None"; Reason = ""; Checked = ""; Note = "" }
    $none = @{
        Title = "インデックスの状態"; Selected = $false; Name = ""; Path = ""; FolderStatus = ""
        Badge = @{ Text = ""; Level = "None" }; Updated = ""; Count = ""; Rows = @(); RowsKey = ""; Fast = $emptyFast
        Multi = $false; Hint = ""; Run = $noRun; Error = ""; ErrorHint = ""
    }
    if ($multiCount -ge 2) {
        $multi = $none.Clone()
        $multi.Title = "{0:#,0} 件を選択中" -f $multiCount
        $multi.Multi = $true
        $multi.Hint = "1 件だけ選ぶと、ここに詳細を表示します。"
        return $multi
    }
    $selected = @($items | Where-Object { $null -ne $_ })
    if ($selected.Count -ne 1) {
        return $none
    }
    $item = $selected[0]

    $hasCount = $item.FileCountText -and $item.FileCountText -ne "－"
    $count = if ($hasCount) { "$($item.FileCountText) ファイル" } else { "まだ更新していません" }
    $updated = if ($item.LastIngestedText) { "最終更新 $($item.LastIngestedText)" } else { "" }

    $rows = New-Object 'System.Collections.Generic.List[object]'
    [void]$rows.Add(@{ Label = "対象ファイル数"; Value = $(if ($hasCount) { "$($item.FileCountText) ファイル" } else { "－" }) })
    [void]$rows.Add(@{ Label = "最終更新"; Value = $(if ($item.LastIngestedText) { [string]$item.LastIngestedText } else { "－" }) })

    $fast = @{ Shown = $false; Value = 0.0; Text = ""; State = [string]$item.FastText; Level = [string]$item.FastLevel; Reason = ""; Checked = [string]$item.FastCheckedText; Note = "" }
    if ($item.FastLevel -eq "Ng" -and $item.FastToolTip) {
        $fast.Reason = ([string]$item.FastToolTip -split "`n")[0]
    }
    if ($null -ne $fastEntry -and $fastEntry.Folders -gt 0) {
        $done = [Math]::Max(0, [Math]::Min([int]$fastEntry.Folders, [int]$fastEntry.Folders - [int]$fastEntry.Waiting))
        $fast.Shown = $true
        $fast.Value = $done / [double]$fastEntry.Folders
        $fast.Text = "反映済み $done / $($fastEntry.Folders) フォルダ"
        if ($done -lt [int]$fastEntry.Folders) {
            $fast.Note = "反映待ちのフォルダは通常の検索で調べます。検索結果は変わりませんが、時間がかかります。"
        }
    }
    $run = $noRun
    if ($item.IndexLevel -eq "Run") {
        $run = @{ Shown = $true; Value = 0.0; CountText = ""; FileText = "" }
        if ($null -ne $running) {
            $parts = New-Object 'System.Collections.Generic.List[string]'
            if ([int]$running.Total -gt 0) {
                $counts = "{0:#,0} / {1:#,0} 件" -f [int]$running.Processed, [int]$running.Total
                if ([int]$running.Failed -gt 0) {
                    $counts += "（失敗 {0:#,0} 件）" -f [int]$running.Failed
                }
                [void]$parts.Add($counts)
            }
            if ($running.Eta) {
                [void]$parts.Add([string]$running.Eta)
            }
            $run.Value = [Math]::Max(0.0, [Math]::Min(1.0, [double]$running.Ratio))
            $run.CountText = $parts -join "・"
            $run.FileText = if ($running.Current) { "更新中のファイル：$($running.Current)" } else { "" }
        }
    }
    # 元のフォルダが見つからないとき（更新中は更新の表示を優先する）は、バッジを「エラー」にして、理由と直し方を出す
    $badge = @{ Text = [string]$item.IndexText; Level = [string]$item.IndexLevel; Sub = [string]$item.IndexSub }
    $errorText = ""
    $errorHint = ""
    if ($item.FolderMissing -and $item.IndexLevel -ne "Run") {
        $badge = @{ Text = "エラー"; Level = "Ng"; Sub = "" }
        $errorText = "フォルダが見つかりません。"
        $errorHint = "フォルダが移動・削除されたか、ネットワークにつながっていない可能性があります。`nフォルダパスの［...］でフォルダを選び直してください。"
    }
    $fixed = @(
        $item.Name, $item.Path, $item.FolderStatus, $badge.Text, $badge.Level, $badge.Sub, $updated, $count
        $fast.State, $fast.Level, $fast.Reason, $fast.Checked
    ) -join "`t"
    $rowsKey = $fixed + "`n" + (@($rows | ForEach-Object { "$($_.Label)`t$($_.Value)" }) -join "`n")
    return @{
        Title = "$($item.Name) - 詳細"; Selected = $true; Name = [string]$item.Name; Path = [string]$item.Path
        FolderStatus = [string]$item.FolderStatus
        Badge = $badge
        Updated = $updated; Count = $count; Rows = $rows.ToArray(); RowsKey = $rowsKey; Fast = $fast
        Multi = $false; Hint = ""; Run = $run; Error = $errorText; ErrorHint = $errorHint
    }
}
