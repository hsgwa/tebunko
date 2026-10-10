# インデックス更新の確認ダイアログ・進み具合に出す文言の決定。
# 画面に触らないため、そのままテストできる（tests\tebunko\ui\indexing_view.Tests.ps1）。
#
# 色は Level（Ok / Wait / Ng / Run / None）で返し、実際の色は theme.xaml の Badge.* から引く（画面側）。

function newPlanViewRows {
    # 更新の予定（取り込み予定.tsv の行）を、確認のダイアログに出す形にする。
    # 行は Name・Path・TotalText（対象ファイル数）・StatusText（バッジの文言）・Level（バッジの色。Wait / Ok / None / Ng）・DetailText（バッジの ToolTip）
    #   onlyNames: 選んだものだけの回のインデックス名（空なら全部）。選ばなかったものは出さない（「対象外」とも出さない）
    param (
        $plan,  # readIngestPlan の結果
        [string[]]$onlyNames = @()
    )

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($item in @($plan)) {
        # 空の配列を渡すと @($plan) に $null が 1 つ入るため、ここで外す
        if ($null -eq $item) { continue }
        if (@($onlyNames).Count -gt 0 -and $item.区分 -ne ${planKindDropped} -and @($onlyNames) -notcontains [string]$item.インデックス名) { continue }
        $row = @{ Name = $item.インデックス名; Path = $item.元のフォルダ }
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
                    @("前回失敗", $item.前回失敗))) {
                if ($pair[1] -gt 0) {
                    $parts.Add("$($pair[0]) $('{0:#,0}' -f $pair[1]) 件")
                }
            }
            if ($item.取り込み対象 -gt 0) {
                $row.StatusText = "要更新"
                $row.Level = "Wait"
                $row.DetailText = "更新するファイル {0:#,0} 件（{1}）" -f $item.取り込み対象, ($parts -join " / ")
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
    # 更新中（run）は終わるまで残し、閉じるボタンも出さない（［中止］がある）。
    # 残りがあって更新していないとき（resume。pending が 1 以上）は、［続きから再開］を載せているので、種類にかかわらず閉じられない。
    # 成功（ok）は数秒で消える。それ以外（warn・info・ng）は、閉じるまで残り、閉じてもよい
    param (
        [string]$level,  # run / resume / ok / info / warn / ng
        [int]$pending = 0  # 取り込みの残りの件数
    )

    if ($level -ne "run" -and $pending -gt 0) { $level = "resume" }
    switch ($level) {
        { $_ -in @("run", "resume") } { return @{ Closable = $false; AutoCloseSeconds = 0 } }
        "ok" { return @{ Closable = $true; AutoCloseSeconds = 8 } }
        default { return @{ Closable = $true; AutoCloseSeconds = 0 } }
    }
}

function getIndexingEndText {
    # インデックス作成が完了した（終了コード 0）ときの、進み具合の見出しと説明（@{ Text; Detail }）を返す。
    # 後回し（利用者のPowerPointが起動していて取り込まなかったファイル）がある場合の文言もここで決める
    param (
        [int]$success,
        [int]$failed,
        [int]$postponed,
        [string]$notice = ""
    )

    $processed = $success + $failed
    if ($processed -gt 0) {
        $counts = "成功 ${success} 件 / 失敗 ${failed} 件"
        if ($postponed -gt 0) {
            $counts += " / 残り ${postponed} 件"
        }
        $parts = New-Object System.Collections.Generic.List[string]
        if ($failed -gt 0) {
            $parts.Add("失敗したファイルと原因は「更新に失敗したファイル」の一覧で確認できます。")
        }
        if ($notice) {
            $parts.Add($notice)
        }
        return @{ Text = "更新が終わりました（${counts}）"; Detail = ($parts -join " ") }
    }
    if ($postponed -gt 0) {
        return @{ Text = "更新が終わりました（更新せずに残したファイル ${postponed} 件）"; Detail = $notice }
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

function getIndexingConfirmFolderCount {
    # 確認の合計に出す「更新するフォルダの数」。取り込む対象（失敗分を含めるなら前回失敗も）が 1 件以上あるフォルダだけを数える
    param (
        [object[]]$plan,    # readIngestPlan の結果
        [bool]$retryFailed  # 失敗分も更新し直すか
    )

    $folders = 0
    foreach ($item in @($plan)) {
        if ($null -eq $item -or $item.区分 -ne ${planKindIngest}) { continue }
        $count = [int]$item.取り込み対象
        if ($retryFailed) { $count += [int]$item.前回失敗 }
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
        [int]$dropped   # 削除予定のインデックスの数（getIndexingDroppedCount）
    )

    return ($targets -eq 0 -and $failed -eq 0 -and $dropped -eq 0)
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
        [int]$dropped = 0   # 削除予定のインデックスの数（getIndexingDroppedCount）
    )

    $total = $targets
    if ($retryFailed) {
        $total += $failed
    }
    $droppedText = if ($dropped -gt 0) { "設定に無いインデックス {0:#,0} 件を削除します。" -f $dropped } else { "" }
    if ($total -gt 0) {
        return @{ Total = $total; Text = ("更新対象: {0:#,0} フォルダ / {1:#,0} ファイル（最新のフォルダは更新しません）" -f $folders, $total) + $(if ($dropped -gt 0) { " " + $droppedText } else { "" }); Button = "更新を開始" }
    }
    if ($dropped -gt 0) {
        return @{ Total = 0; Text = "更新が必要なファイルはありません。" + $droppedText; Button = "更新を開始" }
    }
    return @{ Total = 0; Text = "更新が必要なファイルはありません（すべて最新です）。"; Button = "閉じる" }
}
