# インデックス更新の確認ダイアログ・進み具合に出す文言の決定。
# 画面に触らないため、そのままテストできる（tests\tebunko\ui\indexing_view.Tests.ps1）。
#
# 色は Level（Ok / Wait / Ng / Run / None）で返し、実際の色は theme.xaml の Badge.* から引く（画面側）。

function newPlanViewRows {
    # 更新の予定（取り込み予定.tsv の行）を、確認のダイアログに出す形にする。
    # 行は Name・Path・TotalText（対象ファイル数）・StatusText（バッジの文言）・Level（バッジの色。Wait / Ok / None / Ng）・DetailText（バッジの ToolTip）
    #   onlyNames: 選んだものだけの回のインデックス名（空なら全部）。選ばなかったものは出さない（「対象外」とも出さない）
    #   subfolders: インデックス名 → 設定の subfolders（$false は直下のファイルだけを取り込む）。行の Subfolders になる。無い名前は $true（サブフォルダも含める）
    param (
        $plan,  # readIngestPlan の結果
        [string[]]$onlyNames = @(),
        [hashtable]$subfolders = @{}
    )

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($item in @($plan)) {
        # 空の配列を渡すと @($plan) に $null が 1 つ入るため、ここで外す
        if ($null -eq $item) { continue }
        if (@($onlyNames).Count -gt 0 -and $item.区分 -ne ${planKindDropped} -and @($onlyNames) -notcontains [string]$item.インデックス名) { continue }
        $row = @{ Name = $item.インデックス名; Path = $item.元のフォルダ; Subfolders = ($subfolders[[string]$item.インデックス名] -ne $false); Choosable = ($item.区分 -ne ${planKindDropped}) }
        if ($item.区分 -eq ${planKindDropped}) {
            # 選んだものだけの回でも出す（［更新を開始］で消えるため、利用者に見せずに消さない）
            $row.TotalText = "－"
            $row.StatusText = "削除予定"
            $row.Level = "Ng"
            $row.DetailText = "設定に無いため、［更新を開始］でこのインデックスを削除します"
        } elseif ($item.区分 -eq ${planKindUnchecked}) {
            $row.TotalText = "－"
            $row.StatusText = "対象外"
            $row.Level = "None"
            $row.DetailText = "設定で［すべて更新］の対象から外れているため更新しません（インデックスはそのまま残します）"
        } elseif ($item.区分 -eq ${planKindMissing}) {
            $row.TotalText = "－"
            $row.StatusText = "フォルダなし"
            $row.Level = "Ng"
            $row.DetailText = "元のフォルダが見つかりません（詳細のフォルダパスで場所を変えられます）"
        } else {
            $row.TotalText = "{0:#,0}" -f $item.ファイル数
            # 0 件の内訳は出さない（ふだんは「新規」「更新あり」だけになる）
            $parts = New-Object System.Collections.Generic.List[string]
            foreach ($pair in @(
                    @("新規", $item.新規),
                    @("更新あり", $item.更新あり),
                    @("前回未完了", $item.前回未完了),
                    @("インデックスが無い・壊れている", $item.インデックスなし),
                    @("前回失敗", $item.前回失敗),
                    @("クラウドにだけある", ($item.クラウド + $item.クラウド失敗)))) {
                if ($pair[1] -gt 0) {
                    $parts.Add("$($pair[0]) $('{0:#,0}' -f $pair[1]) 件")
                }
            }
            if ($item.取り込み対象 -gt 0) {
                $row.StatusText = "要更新"
                $row.Level = "Wait"
                $row.DetailText = "更新するファイル {0:#,0} 件（{1}）" -f $item.取り込み対象, ($parts -join " / ")
            } elseif ($item.クラウド -gt 0) {
                # 取り込み対象は無いが、クラウドにだけあるため残すファイルがある。「最新」とは言わず、注意の色で出す
                # （クラウドにだけある前回失敗は、更新し直す指定が無ければ取り込み対象にならない。手元の前回失敗と同じく「最新」のまま内訳に出す）
                $row.StatusText = "要確認"
                $row.Level = "Wait"
                $row.DetailText = $parts -join " / "
            } else {
                $row.StatusText = "最新"
                $row.Level = "Ok"
                $row.DetailText = if ($parts.Count -gt 0) { $parts -join " / " } else { "すべて最新です" }
            }
        }
        $rows.Add($row)
    }
    return , $rows.ToArray()
}

function getIndexingRecountView {
    # 確認ダイアログで選びを変えたときの見せ方。変えた行があれば、合計・説明・主ボタンを数え直しの文言にする。
    #   rowCount: 一覧の行の数 / changedCount: 選びを変えた行の数
    # 戻り値: @{ Changed; TotalText; HintText; ButtonText; RowStatusText; RowLevel }（Changed が $false のときほかは空。今までの文言のまま）
    param (
        [int]$rowCount,
        [int]$changedCount
    )

    if ($changedCount -le 0) {
        return @{ Changed = $false; TotalText = ""; HintText = ""; ButtonText = ""; RowStatusText = ""; RowLevel = "" }
    }
    return @{
        Changed = $true
        TotalText = "更新対象: {0:#,0} フォルダ（件数は数え直したあとに出ます）" -f $rowCount
        HintText = "選びを変えたので、件数を数え直します"
        ButtonText = "数え直して続ける"
        RowStatusText = "取り込み直し"
        RowLevel = "Wait"
    }
}

function getIndexingRecountingView {
    # 確認ダイアログを開いたまま数え直している間の見せ方。全部の行を「数え中」・件数は「…」にし、チェックは押せなくし、
    # 主ボタンは「数え直しています…」で押せなくする（［キャンセル］は押せる）。
    # 戻り値: @{ TotalText; HintText; ButtonText; ButtonEnabled; RowCountText; RowStatusText; RowLevel; RowEditable }
    return @{
        TotalText = "更新対象: 数え直しています…（［キャンセル］で取りやめ）"
        HintText = "件数を数え直しています…"
        ButtonText = "数え直しています…"
        ButtonEnabled = $false
        RowCountText = "…"
        RowStatusText = "数え中"
        RowLevel = "Run"
        RowEditable = $false
    }
}

function getIndexingConfirmChoiceChanges {
    # 確認ダイアログで、元の選びから変えた行を返す（［更新を開始］を押したとき、差があれば数え直しにする）。
    # 行は @{ Name; Subfolders } の配列（元は newPlanViewRows の行、今はダイアログで選び直したあとの行）。
    # 戻り値: 変えた行の @{ Name; Subfolders }（今の値）の配列。差が無ければ空
    param (
        [object[]]$original,
        [object[]]$current
    )

    $before = @{}
    foreach ($row in @($original | Where-Object { $null -ne $_ })) {
        $before[[string]$row.Name] = ($row.Subfolders -ne $false)
    }
    $changes = New-Object System.Collections.Generic.List[object]
    foreach ($row in @($current | Where-Object { $null -ne $_ })) {
        $name = [string]$row.Name
        $now = ($row.Subfolders -ne $false)
        if ($before.ContainsKey($name) -and $before[$name] -ne $now) {
            $changes.Add(@{ Name = $name; Subfolders = $now })
        }
    }
    return , $changes.ToArray()
}

function getIndexingRecountPlan {
    # ［更新を開始］で、数え直しにするかと、数え直した後にやり直す対象を決める。
    #   changes   : getIndexingConfirmChoiceChanges の結果
    #   onlyNames : 今の回の選んだインデックス名（［すべて更新］なら空。やり直しでもそのまま使う）
    #   canceled  : ［キャンセル］で取りやめたか（取りやめのときは何も書かず、数え直しにもしない）
    # 戻り値: @{ Recount（数え直すか）; OnlyNames; Changes（設定に書く選び。数え直さないときは空） }
    param (
        [object[]]$changes,
        [string[]]$onlyNames = @(),
        [bool]$canceled = $false
    )

    $changes = @($changes | Where-Object { $null -ne $_ })
    if ($canceled -or $changes.Count -eq 0) {
        return @{ Recount = $false; OnlyNames = @($onlyNames); Changes = @() }
    }
    return @{ Recount = $true; OnlyNames = @($onlyNames); Changes = $changes }
}

function getIndexingSkippedView {
    # 選んだものだけの回で、更新できなかった名前（インデクサの OnlySkipped。@{ Name; Reason } の配列）を知らせる文言。
    # 無ければ $null。@{ Heading; Detail }
    param (
        [object[]]$skipped
    )

    $skipped = @($skipped | Where-Object { $null -ne $_ })
    if ($skipped.Count -eq 0) { return $null }
    return @{
        Heading = "{0:#,0} 件のインデックスは更新できませんでした。" -f $skipped.Count
        Detail = ($skipped | ForEach-Object { "「$($_.Name)」: $($_.Reason)" }) -join "`n"
    }
}

function getIndexingCurrentName {
    # 取り込み中のファイル（<インデックス名>\<相対パス>）から、インデックス名を取り出す。取れなければ空文字列
    param (
        [string]$current
    )

    $index = $current.IndexOf("\")
    if ($index -lt 1) { return "" }
    return $current.Substring(0, $index)
}

function getIndexingBannerLevel {
    # 更新が終わった帯の色の種類（info/warn/ok/ng のうち、ここでは warn か ok）。
    # 注意が要る終わり方（更新そのものができなかった・中止・取りやめ・失敗したファイルがある）は橙の warn、
    # 失敗なく終わったときだけ緑の ok。赤（ng）は、使えない状態（ワークスペースが使えないなど）にだけ使う
    param (
        [int]$exitCode,  # インデクサの終了コード（0 完了、1 続けられないエラー、2 中止・取りやめ）
        [int]$failed     # 失敗したファイルの件数
    )

    if ($exitCode -ne 0) { return "warn" }
    if ($failed -gt 0) { return "warn" }
    return "ok"
}

function getIndexingBannerBehavior {
    # 帯がいつ消えるか。@{ Closable; AutoCloseSeconds }
    #   Closable: 閉じるボタン（×）を出すか
    #   AutoCloseSeconds: 出してから自動で消すまでの秒数（0 なら自動では消えない）
    # 帯はどの種類でも閉じられる（閉じるボタンを出す）。更新中（run）は、進み具合がステータスバーと行にも出るので閉じてよい。
    # 閉じるのは見せ方だけで、更新は止まらない（止めるのは帯の［中止］か行の［中止］。帯を閉じたあとは行の［中止］）。
    # 自動で消えるのは、残りの無い成功（ok）だけ（数秒）。更新中と、残りがあって更新していないとき（pending が 1 以上。
    # 内部では resume として扱う）は、［続きから再開］を載せているので、消さずに残す。それ以外（warn・info・ng）も閉じるまで残る
    param (
        [string]$level,  # run / ok / info / warn / ng
        [int]$pending = 0  # 取り込みの残りの件数
    )

    if ($level -ne "run" -and $pending -gt 0) { $level = "resume" }
    switch ($level) {
        { $_ -in @("run", "resume") } { return @{ Closable = $true; AutoCloseSeconds = 0 } }
        "ok" { return @{ Closable = $true; AutoCloseSeconds = 8 } }
        default { return @{ Closable = $true; AutoCloseSeconds = 0 } }
    }
}

function getIndexingBannerView {
    # 帯に何を出すか。@{ Level; Closable; AutoCloseSeconds }
    #   Level: 帯の色・振る舞いの種類。更新中の帯（kind が running）は run、それ以外は渡された level のまま
    #   Closable・AutoCloseSeconds: getIndexingBannerBehavior のとおり（閉じるボタンを出すか・自動で消えるまでの秒数）
    # 残りの件数は、いま更新していないときだけ数える（更新中は残りを 0 とみる）。数え方はここで決める
    param (
        [string]$kind,           # 帯の種類（running / interrupted / done / 空）
        [string]$level,          # 色の種類（info / warn / ok / ng）。更新中は使わない
        [int]$pending = 0,       # 取り込みの残りの件数
        [bool]$indexing = $false # いま更新しているか
    )

    $effectiveLevel = if ($kind -eq "running") { "run" } else { $level }
    $effectivePending = if ($indexing) { 0 } else { $pending }
    $behavior = getIndexingBannerBehavior $effectiveLevel $effectivePending
    return @{
        Level            = $effectiveLevel
        Closable         = [bool]$behavior.Closable
        AutoCloseSeconds = [int]$behavior.AutoCloseSeconds
    }
}

function getIndexingResumeBannerDecision {
    # 更新の集計を読み直したときの、帯の扱い。@{ Action; Dismissed; CanResume }
    #   Action: Show（中断の帯を出す）/ Text（出ている中断の帯の文だけ直す）/ Retime（終わりの帯の自動で消える時計を決め直す）/
    #           Hide（出ている中断の帯を消す）/ None（何もしない）
    #   CanResume: 残りがあり、いま更新していない（［続きから再開］を出せる）か
    #   Dismissed: 次に持ち越す「利用者が閉じた」印。残りが無くなる・更新を始めると戻す（閉じたあとの読み直しでは出し直さない）
    param (
        [int]$pending,             # 取り込みの残りの件数
        [bool]$indexing,           # いま更新しているか
        [bool]$panelVisible,       # 帯が出ているか
        [string]$kind,             # 出ている帯の種類（running / interrupted / done / 空）
        $dismissed = $false,       # 利用者が閉じた印（まだ一度も決めていないと $null）
        [string]$level = "",       # 出ている帯の level（getIndexingBannerView に渡す）
        [int]$currentAutoClose = 0 # いま掛けている自動で消えるまでの秒数
    )

    $dismissed = [bool]$dismissed
    $canResume = ($pending -gt 0) -and !$indexing
    if (!$canResume) { $dismissed = $false }
    $action = "None"
    if ($canResume -and !$panelVisible -and !$dismissed) {
        $action = "Show"
    } elseif ($canResume -and $kind -eq "interrupted") {
        $action = "Text"
    } elseif ($kind -eq "done" -and $panelVisible) {
        # 終わりの帯（完了・中止）は、出した時点の古い集計で時計を決めている。残りの有無が変わって秒数が違うときだけ決め直す
        $seconds = (getIndexingBannerView $kind $level $pending $indexing).AutoCloseSeconds
        if ($seconds -ne $currentAutoClose) { $action = "Retime" }
    } elseif (!$canResume -and $kind -eq "interrupted") {
        $action = "Hide"
    }
    return @{ Action = $action; Dismissed = $dismissed; CanResume = $canResume }
}

function getIndexingEndText {
    # インデックス作成が完了した（終了コード 0）ときの、進み具合の見出しと説明（@{ Text; Detail }）を返す。
    # 利用者のPowerPointのために取り込めなかったファイルは、失敗の件数に入っている
    param (
        [int]$success,
        [int]$failed,
        [int]$cloudSkipped = 0  # クラウドにだけあるため、ダウンロードせずに残した件数
    )

    $cloudText = if ($cloudSkipped -gt 0) { "クラウドにだけあるファイル {0:#,0} 件は、ダウンロードせずに残しました（次のインデックス作成でも確かめます）。" -f $cloudSkipped } else { "" }
    $processed = $success + $failed
    if ($processed -gt 0) {
        $counts = "成功 ${success} 件 / 失敗 ${failed} 件"
        $parts = New-Object System.Collections.Generic.List[string]
        if ($failed -gt 0) {
            $parts.Add("失敗したファイルと原因は「ファイルの一覧」（失敗のみ）で確認できます。")
        }
        if ($cloudText) {
            $parts.Add($cloudText)
        }
        return @{ Text = "更新が終わりました（${counts}）"; Detail = ($parts -join " ") }
    }
    if ($cloudSkipped -gt 0) {
        return @{ Text = $cloudText; Detail = "" }
    }
    return @{ Text = "更新が必要なファイルはありませんでした"; Detail = "" }
}

function getIndexingStateText {
    # ボタンの上の一言（取り込み一覧の「未取り込み」の残り件数から。中断でも後回しでも同じ状態のため見分けない）。
    # インデックス作成中は出さない（そのときは別の一言をボタンの下に出す）
    param (
        [int]$pending,
        [bool]$indexing
    )

    if ($pending -gt 0 -and !$indexing) {
        return "更新を中断しました（残り ${pending} 件）"
    }
    return ""
}

function getWorkspaceCheckingStatus {
    # インデックス作成を始める前に、ネットワークのワークスペースを確かめている間のステータス
    return "ワークスペースを確かめています…（共有フォルダに接続できないときは、しばらくかかります）"
}

function getWorkspaceUnreachableStatus {
    # ネットワークのワークスペースに接続できず、インデックス作成を始めなかったときのステータス
    param (
        [string]$dir
    )

    return "ワークスペースに接続できません：${dir}"
}

function getReingestConfirm {
    # ［すべて更新］の確かめ。前の版のしるしがあり content_index\ が空のときだけ確かめの文言を返し、
    # ほかの 3 通り（しるしが無い・空でない）では確かめを出さない（空を返す）
    param (
        [bool]$hasLegacyIndex,
        [bool]$contentEmpty
    )

    if ($hasLegacyIndex -and $contentEmpty) {
        return "前の版のインデックスは使えないため、元のファイルをすべて更新し直します。ファイルが多いと時間がかかります。始めますか？"
    }
    return ""
}

function testIndexingPlanSelected {
    # 確認に数える行か。選んだものだけの回（onlyNames）では、選んだインデックスの行だけ。
    # 取り込み予定は、選ばなかったインデックスの行も含むことがあるため、画面側の確認で絞る
    param (
        $item,
        [string[]]$onlyNames = @()
    )

    if ($null -eq $item) { return $false }
    $names = @($onlyNames)
    return ($names.Count -eq 0 -or $names -contains [string]$item.インデックス名)
}

function getIndexingConfirmFolderCount {
    # 確認の合計に出す「更新するフォルダの数」。取り込む対象（失敗分を含めるなら前回失敗も）が 1 件以上あるフォルダだけを数える
    param (
        [object[]]$plan,    # readIngestPlan の結果
        [bool]$retryFailed, # 失敗分も更新し直すか
        [bool]$includeCloud = $false,  # クラウドにだけあるファイルもダウンロードして取り込むか
        [string[]]$onlyNames = @()     # 選んだものだけの回のインデックス名（空なら全部）
    )

    $folders = 0
    foreach ($item in @($plan)) {
        if (-not (testIndexingPlanSelected $item $onlyNames) -or $item.区分 -ne ${planKindIngest}) { continue }
        $count = [int]$item.取り込み対象
        if ($retryFailed) { $count += [int]$item.前回失敗 }
        if ($includeCloud) {
            $count += [int]$item.クラウド
            if ($retryFailed) { $count += [int]$item.クラウド失敗 }
        }
        if ($count -gt 0) { $folders++ }
    }
    return $folders
}

function getIndexingDroppedCount {
    # 確認に出す「削除予定」（設定から外れたインデックス）の数。選んだものだけの回でも数える
    param (
        [object[]]$plan  # readIngestPlan の結果
    )

    return @(@($plan) | Where-Object { $null -ne $_ -and $_.区分 -eq ${planKindDropped} }).Count
}

function isIndexingConfirmNothing {
    # 確認に出すものが何も無いか（更新するファイルも、前回の失敗も、削除予定も無い）。真なら閉じるだけのダイアログにする。
    # 削除予定があるのに閉じるだけにすると、閉じることが承認扱いになり、見せたまま消してしまう
    param (
        [int]$targets,  # 更新対象のファイル数
        [int]$failed,   # 前回失敗したファイル数
        [int]$dropped,  # 削除予定のインデックスの数（getIndexingDroppedCount）
        [int]$cloud = 0 # クラウドにだけあるファイル数（1 件でもあれば、ダウンロードして取り込むかを選べるよう、閉じるだけにしない）
    )

    return ($targets -eq 0 -and $failed -eq 0 -and $dropped -eq 0 -and $cloud -eq 0)
}

function getIndexingConfirmIntro {
    # 確認の説明文。削除予定があるときだけ、消すことが分かる一文を返す（無ければ空文字列。ふだんは何も出さない）
    param (
        [int]$dropped = 0  # 削除予定のインデックスの数（getIndexingDroppedCount）
    )

    if ($dropped -gt 0) {
        return "［更新を開始］を押すと、設定に無いインデックスも削除します。"
    }
    return ""
}

function getIndexingConfirmText {
    # 「失敗分も更新し直す」のチェックに合わせた、合計の文言と主ボタンの文言
    param (
        [int]$targets,      # 更新対象のファイル数
        [int]$failed,       # 前回失敗したファイル数
        [bool]$retryFailed, # 失敗分も更新し直すか
        [int]$folders = 0,  # 更新するフォルダの数（失敗分を含めるかは呼び出し側で数える）
        [int]$dropped = 0,  # 削除予定のインデックスの数（getIndexingDroppedCount）
        [int]$cloud = 0,    # クラウドにだけあるファイル数（getIndexingCloudView。取り込み対象には数えていない）
        [bool]$includeCloud = $false  # クラウドにだけあるファイルもダウンロードして取り込むか
    )

    $total = $targets
    if ($retryFailed) {
        $total += $failed
    }
    if ($includeCloud) {
        $total += $cloud
    }
    $droppedText = if ($dropped -gt 0) { "設定に無いインデックス {0:#,0} 件を削除します。" -f $dropped } else { "" }
    $cloudText = if ($cloud -gt 0 -and -not $includeCloud) { "クラウドにだけあるファイル {0:#,0} 件は、ダウンロードせずに残します。" -f $cloud } else { "" }
    $tail = (@($droppedText, $cloudText) | Where-Object { $_ }) -join " "
    if ($total -gt 0) {
        return @{ Total = $total; Text = ("更新対象: {0:#,0} フォルダ / {1:#,0} ファイル（最新のフォルダは更新しません）" -f $folders, $total) + $(if ($tail) { " " + $tail } else { "" }); Button = "更新を開始" }
    }
    if ($dropped -gt 0) {
        return @{ Total = 0; Text = "更新が必要なファイルはありません。" + $tail; Button = "更新を開始" }
    }
    if ($cloud -gt 0) {
        # 閉じるだけにしない（［更新を開始］で、クラウドにだけあるため残すファイルの扱いを進める。取りやめは［キャンセル］）
        return @{ Total = 0; Text = "更新が必要なファイルはありません。" + $cloudText; Button = "更新を開始" }
    }
    return @{ Total = 0; Text = "更新が必要なファイルはありません（すべて最新です）。"; Button = "閉じる" }
}

# ---- クラウドにだけあるファイル（OneDrive・SharePoint のオンデマンド） ----

function formatCloudSize {
    # 合計サイズの書き方。KB・MB・GB のどれかで、小数点以下 1 桁（"約 12.3 MB"）。1 KB に満たなければ "1 KB 未満"
    param (
        [long]$bytes
    )

    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    if ($bytes -lt 1KB) { return "1 KB 未満" }
    # 小数点以下 1 桁に丸めて 1024 になるとき（1 MB の直前など）は、次の単位で書く（"約 1024.0 KB" にしない）
    $units = @("KB", "MB", "GB")
    $value = [double]$bytes / 1KB
    $index = 0
    while ($index -lt 2 -and [math]::Round($value, 1, [System.MidpointRounding]::AwayFromZero) -ge 1024) {
        $value = $value / 1024
        $index++
    }
    return "約 " + $value.ToString("0.0", $culture) + " " + $units[$index]
}

function getIndexingCloudView {
    # 確認に出す、クラウドにだけあるファイルの件数と合計サイズ（@{ Count; Bytes }）。
    # 前回失敗したファイルは、「失敗分も更新し直す」を選んだときだけ数える（選ばなければ、クラウドに関係なく取り込まないため）
    param (
        [object[]]$plan,    # readIngestPlan の結果
        [bool]$retryFailed,
        [string[]]$onlyNames = @()  # 選んだものだけの回のインデックス名。取り込み予定には選ばなかった行も含まれるため、ここで絞る
    )

    $count = 0
    [long]$bytes = 0
    foreach ($item in @($plan)) {
        if (-not (testIndexingPlanSelected $item $onlyNames) -or $item.区分 -ne ${planKindIngest}) { continue }
        $count += [int]$item.クラウド
        $bytes += [long]$item.クラウド容量
        if ($retryFailed) {
            $count += [int]$item.クラウド失敗
            $bytes += [long]$item.クラウド失敗容量
        }
    }
    return @{ Count = $count; Bytes = $bytes }
}

function getIndexingCloudWarning {
    # 確認に出す、クラウドにだけあるファイルの注意書き
    param (
        [int]$count,
        [long]$bytes
    )

    return ("クラウドにだけあるファイル {0:#,0} 件（合計 {1}）は、既定ではダウンロードせずに残します。取り込むと、そのすべてをこの PC にダウンロードします。登録するフォルダを、要るものだけに絞ることをおすすめします。" -f $count, (formatCloudSize $bytes))
}

function getIndexingCloudCheckText {
    param (
        [int]$count
    )

    return ("クラウドにだけあるファイル {0:#,0} 件もダウンロードして取り込む" -f $count)
}

${indexingCloudRememberText} = "次からこの確認を出さずにダウンロードする"

function getIndexingCloudConfirmView {
    # ダウンロードして取り込む前の、2 度目の確認の文言（showConfirm に渡す）
    param (
        [int]$count,
        [long]$bytes,
        [int]$timeoutMinutes  # 1 ファイルの制限時間（分。インデクサの ${ingestFileTimeoutMinutes}）
    )

    return @{
        Heading = ("{0:#,0} 件（合計 {1}）をダウンロードして取り込みますか？" -f $count, (formatCloudSize $bytes))
        Facts = @(
            "ダウンロードの間、ネットワークの帯域とこの PC のディスクの空きを使います。"
            "ダウンロードしたファイルは、OneDrive の［空き容量を増やす］を行うまで、この PC に残ります。"
            "ダウンロードも含めて 1 ファイルが ${timeoutMinutes} 分以内に終わらなければ、そのファイルは失敗になります。失敗したファイルは、手元に落としてから取り込み直してください。"
        )
        Choices = @(@{ Text = "ダウンロードして取り込む"; Value = "download" })
        CancelText = "キャンセル"
    }
}

function testIndexingCloudNeedsConfirm {
    # ダウンロードして取り込むときに、2 度目の確認を出すか。設定がすでに download（確認を出さない）なら出さない
    param (
        [bool]$cloudChecked,
        [string]$savedMode
    )

    return ($cloudChecked -and $savedMode -ne ${cloudFilesDownload})
}

function getIndexingCloudSavedMode {
    # 開始するときに設定へ保存する値。両方のチェックが付いていれば download、それ以外は ask
    param (
        [bool]$cloudChecked,
        [bool]$rememberChecked
    )

    if ($cloudChecked -and $rememberChecked) {
        return ${cloudFilesDownload}
    }
    return ${cloudFilesAsk}
}

function getIndexingCloudModeToSave {
    # 開始するときに設定の cloudFiles へ書く値。書かないとき（クラウドの欄を出していない・今の設定と同じ）は $null
    param (
        [bool]$cloudVisible,    # クラウドの欄を出したか（getIndexingConfirmView の CloudVisible）
        [bool]$cloudChecked,    # ダウンロードして取り込むを選んだか（同 Include）
        [bool]$rememberChecked,
        [string]$savedMode      # 開くときに読んだ cloudFiles
    )

    if (-not $cloudVisible) { return $null }
    $mode = getIndexingCloudSavedMode $cloudChecked $rememberChecked
    if ($mode -eq $savedMode) { return $null }
    return $mode
}

function getIndexingConfirmView {
    # 確認のダイアログの、チェックの値で変わる表示をまとめて決める。画面は結果を当てるだけにする。
    # クラウドの欄は、そのときの「失敗分も更新し直す」の値で数えたクラウドの件数が 1 件以上のときだけ出す
    # （クラウドにだけある前回失敗は、更新し直すを選んだときだけ数える）。件数が 0 なら、クラウドのチェックは付いていないものとして扱う。
    param (
        [object[]]$plan,          # readIngestPlan の結果
        [string[]]$onlyNames = @(),
        [bool]$retryChecked = $false,
        [bool]$cloudChecked = $false
    )

    $targets = 0
    $failed = 0       # 前回失敗（手元にある）
    $cloudFailed = 0  # 前回失敗（クラウドにだけある）
    foreach ($item in @($plan)) {
        if (-not (testIndexingPlanSelected $item $onlyNames) -or $item.区分 -ne ${planKindIngest}) { continue }
        $targets += [int]$item.取り込み対象
        $failed += [int]$item.前回失敗
        $cloudFailed += [int]$item.クラウド失敗
    }
    $dropped = getIndexingDroppedCount $plan

    $retryVisible = (($failed + $cloudFailed) -gt 0)
    $retry = ($retryChecked -and $retryVisible)
    $retryText = "前回更新に失敗し、その後変わっていないファイル {0:#,0} 件も更新し直す（パスワード付きなど）" -f ($failed + $cloudFailed)
    if ($cloudFailed -gt 0) {
        $retryText += "。うちクラウドにだけある {0:#,0} 件は、ダウンロードして取り込むを選んだときだけ" -f $cloudFailed
    }

    $cloud = getIndexingCloudView $plan $retry $onlyNames
    $cloudVisible = ($cloud.Count -gt 0)
    $include = ($cloudChecked -and $cloudVisible)
    $cloudAll = getIndexingCloudView $plan $true $onlyNames
    $folders = getIndexingConfirmFolderCount $plan $retry $include $onlyNames

    return @{
        Targets        = $targets
        Failed         = $failed
        Dropped        = $dropped
        Nothing        = (isIndexingConfirmNothing $targets ($failed + $cloudFailed) $dropped $cloudAll.Count)
        RetryVisible   = $retryVisible
        RetryText      = $retryText
        Retry          = $retry
        Cloud          = $cloud
        CloudVisible   = $cloudVisible
        Include        = $include
        CloudWarn      = $(if ($cloudVisible) { getIndexingCloudWarning $cloud.Count $cloud.Bytes } else { "" })
        CloudCheckText = $(if ($cloudVisible) { getIndexingCloudCheckText $cloud.Count } else { "" })
        Confirm        = (getIndexingConfirmText $targets $failed $retry $folders $dropped $cloud.Count $include)
    }
}
