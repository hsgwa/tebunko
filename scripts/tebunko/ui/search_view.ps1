# ［2 検索］タブの判断（検索できるか・注意書き・検索条件の説明）。
# 画面に触らないため、そのままテストできる（tests\tebunko\ui\search_view.Tests.ps1）。

function describeSearchOption {
    # 既定から変えた検索条件を「大文字と小文字を区別・対象ファイル：*.xlsx」のように返す（無ければ空）
    param (
        [hashtable]$option
    )

    $items = @()
    if ($option.CaseSensitive) {
        $items += "大文字と小文字を区別"
    }
    if ($option.FileFilter) {
        $items += "対象ファイル：$($option.FileFilter)"
    }
    # 既定はどちらも検索する（項目が無い古い形の条件も、検索するものとみなす）
    if ($option.ContainsKey("IncludeShapes") -and -not $option.IncludeShapes) {
        $items += "図形を除く"
    }
    if ($option.ContainsKey("IncludeComments") -and -not $option.IncludeComments) {
        $items += "コメントを除く"
    }
    return ($items -join "・")
}

function getWordNotice {
    # 検索ワードの下に出す注意書き（出さないときは空文字列）
    param (
        [string]$word,
        [bool]$useRegex
    )

    if ($useRegex -and $word -ne "" -and !(isValidRegex $word)) {
        return "正規表現として不正なため、文字どおり検索します。"
    }
    return ""
}

function getFastSearchReason {
    # 高速検索の状態（getFastSearchView の status）から、Windows Search の理由（getWindowsSearchState の値）を取り出す。
    # まだ確かめていない（$null）ときは Ok（使えるものとして扱う）
    param (
        $status
    )

    if ($status -and $status.Reason) {
        return [string]$status.Reason
    }
    return "Ok"
}

function getFastSearchView {
    # 検索ワードの下に出す、高速検索（Windows Search で先に絞る）の使用可否と短い表示。
    #   status  : @{ Reason（getWindowsSearchState の値）; Progress（getSystemIndexProgress の値。$null は数えていない） }。
    #             $null はまだ確かめていない（使えるものとして扱う）
    #   checking: 確かめている間（表示を押した後など）。表示だけを「確認中…」にする（Usable は変えない）
    # 詳しいことは、押したときの画面（getFastSearchDetail）に出す。ワードの理由（正規表現・1 文字）は利用者がその場で直せるため、Windows Search の理由より先に出す
    param (
        $status,
        [bool]$useRegex,
        [string]$word,
        [bool]$checking = $false
    )

    $reason = getFastSearchReason $status
    $usable = testFastSearchUsable ($reason -eq "Ok") $useRegex $word
    if ($checking) {
        return @{ Usable = $usable; Text = "高速検索：確認中…" }
    }
    # ワードが空のときは、まだ入力していないだけなので、1 文字とは言わない
    if ($useRegex) {
        $text = "高速検索：使用不可（正規表現）"
    } elseif ($word -ne "" -and (getSearchGrams $word).Count -eq 0) {
        $text = "高速検索：使用不可（1 文字）"
    } elseif ($reason -eq "Ok") {
        $text = "高速検索：使用可"
        $progress = if ($status) { $status.Progress } else { $null }
        if ($progress -and $progress.Folders -gt 0 -and $progress.Waiting -gt 0) {
            # 反映待ちがあるうちは、進み具合を百分率で添える（数そのものは詳しい画面に出す）
            $percent = [int][Math]::Floor(($progress.Folders - $progress.Waiting) * 100 / $progress.Folders)
            $text = "高速検索：使用可（反映 ${percent}%）"
        }
    } else {
        $labels = @{
            NoFolder     = "インデックスがありません"
            NoConnection = "Windows Search に接続できません"
            NotInScope   = "Windows Search の対象外"
            NotYet       = "Windows Search の準備中"
        }
        $label = $labels[$reason]
        $text = if ($label) { "高速検索：使用不可（$label）" } else { "高速検索：使用不可" }
    }
    return @{ Usable = $usable; Text = $text }
}

function testFastSearchPreparing {
    # 高速検索が準備中か（画面が、準備中の間だけ確かめ直す（5 分おき）かどうかを決める）。
    # Windows Search がまだ索引していない（NotYet）か、反映待ちのフォルダがあるとき。準備が終われば偽になり、確かめ直しを止める
    param (
        $status
    )

    if ($null -eq $status) {
        return $false
    }
    if ($status.Reason -eq "NotYet") {
        return $true
    }
    return [bool]($status.Progress -and $status.Progress.Waiting -gt 0)
}

function getFastSearchDetail {
    # 高速検索の表示を押したときに出す詳しい画面の中身。@{ Title; Message } を返す（getSourceConnectFailureDialog と同じ形）。
    # 今の状態と理由・反映の進み具合・理由ごとの直し方・確かめた時刻を書く。
    # 画面には、内部の言葉（システムインデックス・集約ファイル・本文インデックス）を書かない（区別が要るときは「インデックス（高速検索用）」）。
    # フォルダの名前 system_index・content_index は、直し方の中でだけ書く
    #   status: getFastSearchView と同じ（Reason・Progress に加えて、確かめた時刻 CheckedAt があれば書く）
    param (
        $status,
        [bool]$useRegex,
        [string]$word
    )

    $reason = getFastSearchReason $status
    $progress = if ($status) { $status.Progress } else { $null }
    $lines = New-Object System.Collections.Generic.List[string]

    $states = @{
        Ok           = "Windows Search：使えます"
        NoFolder     = "Windows Search：インデックス（高速検索用）がありません"
        NoConnection = "Windows Search：接続できません"
        NotInScope   = "Windows Search：インデックス（高速検索用）が索引の対象外のようです"
        NotYet       = "Windows Search：インデックス（高速検索用）をまだ索引していません（準備中）"
    }
    $state = $states[$reason]
    $lines.Add($(if ($state) { $state } else { "Windows Search：使えません" }))
    if ($useRegex) {
        $lines.Add("［正規表現を使う］がオンのため、今の検索ではすべてを検索します。")
    } elseif ($word -ne "" -and (getSearchGrams $word).Count -eq 0) {
        $lines.Add("検索ワードに 2 文字以上の部分が無いため、今の検索ではすべてを検索します。")
    }
    if ($progress) {
        $folders = [int]$progress.Folders
        $waiting = [int]$progress.Waiting
        $lines.Add("インデックス（高速検索用）の反映：反映済み $(($folders - $waiting).ToString('N0')) / $($folders.ToString('N0')) フォルダ（反映待ち $($waiting.ToString('N0'))）")
    }

    $lines.Add("")
    $fixes = @{
        NoFolder     = "［1 インデックス管理］でインデックスを作成してください。作成が終わると、高速検索用のインデックスも作られます。ワークスペースを変えたときは、そのワークスペースにインデックスがあるかも確かめてください。"
        NoConnection = "Windows Search のサービス（WSearch）が動いているかを確かめてください。止まっているときは、サービスの管理画面で開始してください。"
        NotInScope   = "Windows の［インデックスのオプション］で、ワークスペースの system_index フォルダを索引の対象に加えてください。管理者の権限が要り変えられない PC では、PC の管理者に頼んでください。作ったばかりのワークスペースでは、対象でも「対象外」と出ることがあります。その場合は、しばらくしてからもう一度押してください。"
        NotYet       = "Windows Search が索引し終えるのを待ってください。終われば使えるようになります（この表示は 5 分おきに確かめ直します）。"
    }
    if ($fixes[$reason]) {
        $lines.Add("直し方：$($fixes[$reason])")
    } elseif ($progress -and $progress.Waiting -gt 0) {
        $lines.Add("反映待ちのフォルダは、Windows Search が索引し終えるまで、すべてを検索します。待てば使えるようになります。")
    }
    # 反映が終わっていない間だけ、遅い理由の案内を出す
    if ($progress -and $progress.ContentIndexed -and ($reason -eq "NotYet" -or ($reason -eq "Ok" -and $progress.Waiting -gt 0))) {
        $lines.Add("ワークスペースの content_index フォルダも Windows Search の索引の対象になっているため、反映が遅くなっています。［インデックスのオプション］で content_index フォルダを対象から外すと、早くなります。")
    }
    $lines.Add("高速検索が使えなくても、検索の結果は同じです。時間だけが違います。")
    if ($status -and $status.CheckedAt) {
        $lines.Add("")
        $lines.Add("確かめた時刻：$(([datetime]$status.CheckedAt).ToString('yyyy/MM/dd HH:mm:ss'))")
    }
    return @{ Title = "高速検索の状態"; Message = ($lines -join "`n") }
}

function getSearchProgressText {
    # 検索中の要約欄
    param (
        [int]$hits
    )

    return "検索中…　該当 $($hits.ToString('N0')) 件"
}

function getSearchSummaryText {
    # 検索が終わったときの要約欄（ヒットがあるとき）
    param (
        [int]$hits,
        [int]$files,
        [double]$seconds
    )

    return "該当 $($hits.ToString('N0')) 件（$($files.ToString('N0')) ファイル） ・ $($seconds.ToString('0.0')) 秒"
}

function newSearchButtonState {
    # ［検索］ボタンの文言と、押せるかどうか
    param (
        [bool]$searching,    # 検索中か
        [bool]$stopping,     # 中止を頼んだ後か
        [string]$word,       # 検索ワード
        [bool]$hasIndex,     # 検索できるインデックスがあるか
        [int]$targetCount    # 検索対象に選ばれている数
    )

    if ($searching) {
        return @{ Content = "中止"; Enabled = !$stopping }
    }
    return @{ Content = "検索"; Enabled = ($word -ne "" -and $hasIndex -and $targetCount -gt 0) }
}

# ---- 元のファイルを開く（ネットワークにあるときの文言） ----


function getSourceCheckingStatus {
    # 元のファイルの場所を確かめている間のステータス（ネットワークにあるときだけ出す。ローカルはその場で開くため出さない）
    param (
        [string]$path
    )

    return "元のファイルを確かめています…：${path}（共有フォルダに接続できないときは、しばらくかかります）"
}

function getSourceNotFoundStatus {
    # 見つからない・確認で選ばなかったときのステータス
    param (
        [string]$path
    )

    return "元のファイルが見つかりません：${path}"
}

function getSourceUnreachableStatus {
    # 接続できないと分かったときのステータス
    param (
        [string]$folder
    )

    return "元のフォルダに接続できません：${folder}"
}

function getSourceConnectFailureStatus {
    # 接続できない・その他のときの、ダイアログを出した直後のステータス
    param (
        [string]$state,
        [string]$folder,
        [string]$message
    )

    if ($state -eq "Unreachable") {
        return getSourceUnreachableStatus $folder
    }
    return "元のファイルを確かめられませんでした：${message}"
}

function getSourceConnectFailureDialog {
    # 接続できない・その他のときの確認ダイアログの中身。
    #   state: "Unreachable"（接続できない）・"Other"（その他。アクセス拒否・ログオンの失敗など）
    # 返すもの: @{ Heading; Title; Detail; Hint }（Heading はダイアログの見出し、Title・Detail は知らせの1行、Hint は選択肢の説明）
    param (
        [string]$book,
        [string]$state,
        [string]$folder,
        [string]$message
    )

    $heading = "${book} を開けません"
    if ($state -eq "Unreachable") {
        return @{
            Heading = $heading
            Title   = "元のフォルダに接続できません"
            Detail  = $folder
            Hint    = "ネットワーク・VPN の接続を確かめてから、もう一度開いてください"
        }
    }
    return @{
        Heading = $heading
        Title   = "元のファイルを確かめられませんでした"
        Detail  = $message
        Hint    = "アクセスの権限・サインインを確かめてください"
    }
}


# ---- ファイルごとにまとめた表示 ----

function getAppKind {
    # 元のファイル名の拡張子から、アプリの種類（Excel / Word / PowerPoint。どれでもなければ空）を返す
    param (
        [string]$book
    )

    $extension = [System.IO.Path]::GetExtension($book).ToLowerInvariant()
    if ($extension -match '^\.xls') {
        return "Excel"
    }
    if ($extension -match '^\.doc') {
        return "Word"
    }
    if ($extension -match '^\.ppt') {
        return "PowerPoint"
    }
    return ""
}

function describeFileLocations {
    # ファイルの中でヒットした場所（見つかった順・重複なし）を、見出しの右端に出す文字列にする。
    # 1 か所ならその場所、2 か所以上なら「[シート]4月 ほか 2 か所」（場所の表記は describePlace）
    param (
        [string[]]$labels
    )

    # 検索中にファイル・場所が増えるたびに呼ぶため、パイプライン（Where-Object）を使わない（1 回 1 ms を超えて検索が遅くなる）
    $first = ""
    $count = 0
    foreach ($label in $labels) {
        if ($label) {
            if ($count -eq 0) {
                $first = $label
            }
            $count++
        }
    }
    if ($count -le 1) {
        return $first
    }
    return "$first ほか $($count - 1) か所"
}

# ---- 結果の表に並べる項目（見出しと行） ----
# group は FileGroup（types.ps1）と同じ項目（Rows・ShownRows・ShownCount・IsExpanded・Order）を持つもの、
# row は HitRow と同じ項目（Order・Contains(文字列)）を持つもの。

function selectShownRows {
    # 行のうち、絞り込み（空ならすべて）に合うものを、元の順のまま返す
    param (
        $rows,
        [string]$filterText
    )

    $shown = New-Object 'System.Collections.Generic.List[object]'
    foreach ($row in $rows) {
        if ($filterText -eq "" -or $row.Contains($filterText)) {
            $shown.Add($row)
        }
    }
    return , $shown
}

function getResultItems {
    # 結果の表に並べる項目。ファイルごとに見出しを 1 つ置き、開いているファイルだけ、その下に行を並べる。
    # 絞り込みで行が 1 つも残らないファイル（ShownCount が 0）は、見出しも出さない。開いているファイルの行は作ってあること
    param (
        $groups
    )

    $items = New-Object 'System.Collections.Generic.List[object]'
    foreach ($group in $groups) {
        if ($group.ShownCount -eq 0) {
            continue
        }
        $items.Add($group)
        if ($group.IsExpanded) {
            $items.AddRange($group.ShownRows)
        }
    }
    return , $items
}

function getShownHitRows {
    # 絞り込みに合う行を、表の順（閉じているファイルの行も含む）に並べて返す（結果の出力に使う。行はすべて作ってあること）
    param (
        $groups
    )

    $rows = New-Object 'System.Collections.Generic.List[object]'
    foreach ($group in $groups) {
        $rows.AddRange($group.ShownRows)
    }
    return , $rows
}

function prepareHitRow {
    # 画面に出る行の表示用の値（強調セグメント・セル番地・「場所」の列の表記）を作る。作り済みなら何もしない
    param (
        [HitRow]$row
    )

    $row.Prepare()
    if ($row.PlaceDisplay) { return }
    $row.SetPlaceDisplay((describeHitPlace $row.PlaceText $row.IsExcel $row.IsObjectPlace $row.MatchCell $row.MatchCount $row.LineNumber))
}

function sortFileGroups {
    # 列見出しのクリックでの並べ替え。各ファイルの中の行を property の順に並べ替え、
    # ファイルの順は、並べ替えた後の先頭の行の順にする。同じ値のときは見つかった順（Order）
    param (
        $groups,
        [string]$property,
        [bool]$descending
    )

    # 鍵は項目の名前で渡す（スクリプトブロックより速い）。「場所」（Location）は同じ場所の中を行番号の順にする
    # （セル番地の文字の順だと A10 が A9 の前に来るため、番地ではなく行番号で並べる）
    $byValue = @{ Expression = $property; Descending = $descending }
    $byOrder = @{ Expression = "Order"; Descending = $false }
    $rowKeys = if ($property -eq "Location") { @($byValue, @{ Expression = "LineNumber"; Descending = $descending }, $byOrder) } else { @($byValue, $byOrder) }
    foreach ($group in $groups) {
        $sorted = @($group.Rows | Sort-Object $rowKeys)
        $group.Rows.Clear()
        $group.Rows.AddRange([object[]]$sorted)
    }
    $byFirst = @{ Expression = { if ($_.Rows.Count -gt 0) { $_.Rows[0].$property } }; Descending = $descending }
    $sortedGroups = New-Object 'System.Collections.Generic.List[object]'
    $sortedGroups.AddRange([object[]]@($groups | Sort-Object $byFirst, $byOrder))
    return , $sortedGroups
}