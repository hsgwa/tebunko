# 検索バー（検索ワード・種類のチップ・ファイル内の対象・検索ボタン）の判断（判断層）。
# 画面に触らないため、そのままテストできる（tests\tebunko\ui\search\search_bar_view.Tests.ps1）。

function describeSearchOption {
    # 既定から変えた検索条件を「大文字・小文字を区別・種類：Excel・テキスト」のように返す（無ければ空）
    param (
        [hashtable]$option
    )

    $items = @()
    if ($option.CaseSensitive) {
        $items += "大文字・小文字を区別"
    }
    $kindText = describeFileKinds $option.FileKinds
    if ($kindText -ne "") {
        $items += "種類：$kindText"
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

function getFileKindLabel {
    # 種類（excel・word・powerpoint・text）のチップに出す名前。知らない種類は空
    param (
        [string]$kind
    )

    switch ($kind) {
        "excel" { return "Excel" }
        "word" { return "Word" }
        "powerpoint" { return "PowerPoint" }
        "text" { return "テキスト" }
    }
    return ""
}

function toggleFileKind {
    # チップを押したあとの、選ばれている種類（${fileKindNames} の順）を返す。
    # 全部外してよい（空は「1 つも選んでいない」。検索を始めるときに getSearchKindBalloonText が知らせる）
    param (
        [object[]]$kinds,
        [string]$kind
    )

    $current = @(${fileKindNames} | Where-Object { @($kinds) -contains $_ })
    if (${fileKindNames} -notcontains $kind) {
        return $current
    }
    if ($current -contains $kind) {
        return @($current | Where-Object { $_ -ne $kind })
    }
    return @(${fileKindNames} | Where-Object { $current -contains $_ -or $_ -eq $kind })
}

function getSearchKindBalloonText {
    # 検索を始められない理由（種類のチップを 1 つも選んでいないとき）を知らせる吹き出しの文。
    # 始められるときは $null（吹き出しを出さない）
    param (
        [object[]]$kinds
    )

    $chosen = @(${fileKindNames} | Where-Object { @($kinds) -contains $_ })
    if ($chosen.Count -eq 0) {
        return "種類を 1 つ以上選んでください"
    }
    return $null
}

function describeFileKinds {
    # 絞り込んでいる種類を「Excel・テキスト」の形で返す。すべて選んでいる（または空）ときは空
    param (
        [object[]]$kinds
    )

    $chosen = @(${fileKindNames} | Where-Object { @($kinds) -contains $_ })
    if ($chosen.Count -eq 0 -or $chosen.Count -eq ${fileKindNames}.Count) {
        return ""
    }
    return (@($chosen | ForEach-Object { getFileKindLabel $_ }) -join "・")
}

function getNoKindMatchText {
    # 選んだ種類に合うファイルが検索の対象に無いときの文言
    param (
        [object[]]$kinds
    )

    $text = describeFileKinds $kinds
    if ($text -eq "") {
        return "検索できるファイルがありません。"
    }
    return "種類（$text）に合うファイルがありません。"
}

function getTargetCountText {
    # 左の欄の見出し。インデックス（一番上の項目）のうち、すべて選んでいる数 / インデックスの数。無いときは見出しだけ
    param (
        [int]$checked,
        [int]$total
    )

    if ($total -le 0) {
        return "検索対象"
    }
    return "検索対象 $checked / $total"
}

function getWorkspaceUnreachableText {
    # ワークスペースに接続できないときの知らせ（ステータスと、検索対象の欄の文言に使う）
    param (
        [string]$workspaceDir
    )

    return "ワークスペースに接続できません：$workspaceDir"
}

function getWorkspaceCheckingText {
    # ワークスペースが届くかを裏で確かめている間のステータス
    return "ワークスペースを確かめています…"
}

function getTreeFolderFailedText {
    # ツリーのフォルダの子を読み込めなかったときのステータス
    param (
        [string]$folder,
        [string]$reason
    )

    return "フォルダを読み込めませんでした：$folder（$reason）"
}

function getSearchTargetText {
    # 左の欄の見出しのツールチップ（検索対象の詳しい中身）。
    # 優先の順: 読み込み中、接続できない、インデックスが無い、チェックなし、一部、確認中、すべて
    param (
        [bool]$loading,            # 検索対象のツリーを裏で読んでいる最中か
        [string]$connectErrorDir,  # ワークスペースに接続できなかったときのワークスペース（できたなら空）
        [int]$total,               # インデックスの数
        [int]$targetCount,         # 検索対象に選ばれている数
        [bool]$allChecked,         # インデックスがすべてチェックされているか
        [string]$targetsText,      # 選ばれた検索対象の表示（describeSearchTargets）
        $summary                   # 本文インデックスのファイルの集計（Count・LastWrite。まだ数えていなければ $null）
    )

    if ($loading) {
        return "検索対象：読み込んでいます…"
    }
    if ($connectErrorDir) {
        return "検索対象：なし（$(getWorkspaceUnreachableText $connectErrorDir)）"
    }
    if ($total -eq 0) {
        return "［インデックス管理］で作ったインデックスの一覧です"
    }
    if ($targetCount -eq 0) {
        return "検索対象：なし"
    }
    if (!$allChecked) {
        return "検索対象：$targetsText"
    }
    if ($null -eq $summary) {
        return "検索対象：すべて（確認中…）"
    }
    return "検索対象：すべて（集約ファイル $($summary['Count'].ToString('N0')) 件 ・ 最終更新 $(formatTime $summary['LastWrite'])）"
}

function getScopeButtonText {
    # ［ファイル内の対象］ボタンの文言（Text）と、既定から変えているか（Changed）。
    # 既定（図形・コメントも検索）から外したものがあれば「・2 件変更」を付ける
    param (
        [bool]$includeShapes,
        [bool]$includeComments
    )

    $changed = @($includeShapes, $includeComments | Where-Object { !$_ }).Count
    if ($changed -eq 0) {
        return @{ Text = "ファイル内の対象"; Changed = $false }
    }
    return @{ Text = "ファイル内の対象・$changed 件変更"; Changed = $true }
}

function getWordNotice {
    # 正規表現が正しくないときの吹き出しの文（［正規表現］の下に出す。正しいとき・使わないときは空文字列）
    param (
        [string]$word,
        [bool]$useRegex
    )

    if ($useRegex -and $word -ne "" -and !(isValidRegex $word)) {
        return "正規表現が正しくありません"
    }
    return ""
}

function getFastSearchView {
    # 検索ワードの下に出す、高速検索（Windows Search で先に絞る）の使用可否。
    #   available: Windows Search が使えるか（testWindowsSearch）。$null はまだ確かめていない（使えるものとして扱う）
    # Tip は、印にマウスを置いたときのツールチップ。使えるときは印の意味、使えないときは理由と直し方（どの場合も空にしない）。
    # 使えない理由は、上から順に最初に当てはまるもの
    param (
        $available,
        [bool]$useRegex,
        [string]$word
    )

    $usable = testFastSearchUsable ($available -ne $false) $useRegex $word
    $tip = "インデックスを使って速く検索します"
    if (!$usable) {
        if ($available -eq $false) {
            $tip = "検索はできますが時間がかかります　［インデックス管理で確認］"
        } elseif ($useRegex) {
            $tip = "正規表現をオフにすると速く検索できます"
        } else {
            $tip = "空白で区切った語のどれかが 2 文字以上のときに使えます"
        }
    }
    return @{ Usable = $usable; Text = if ($usable) { "高速検索：使用可" } else { "高速検索：使用不可" }; Tip = $tip }
}

function newSearchButtonState {
    # ［検索］ボタンの文言と、押せるかどうか
    param (
        [bool]$searching,    # 検索中か
        [bool]$stopping,     # 中止を頼んだ後か
        [string]$word,       # 検索ワード
        [bool]$hasIndex,     # 検索できるインデックスがあるか
        [int]$targetCount,   # 検索対象に選ばれている数
        [bool]$wordInvalid,  # 正規表現が正しくないか（getWordNotice が空でないとき）
        [bool]$loading = $false  # 検索対象のツリーを読み込み中か（読み込み中は検索できない）
    )

    if ($searching) {
        return @{ Content = "中止"; Enabled = !$stopping }
    }
    return @{ Content = "検索"; Enabled = ($word -ne "" -and $hasIndex -and $targetCount -gt 0 -and !$wordInvalid -and !$loading) }
}

function getBalloonLeft {
    # 吹き出しの左端。基準の左端にそろえ、右端がはみ出すときは右端に収まるまで左へ寄せる（左は 0 より左へ出さない）。
    #   anchorLeft: 基準の項目の左端、width: 吹き出しの幅、available: 置ける幅、margin: 右端に残す間
    param (
        [double]$anchorLeft,
        [double]$width,
        [double]$available,
        [double]$margin = 8
    )

    $left = [Math]::Min($anchorLeft, $available - $margin - $width)
    return [Math]::Max(0.0, $left)
}

function getConditionFlow {
    # 検索条件の行（種類・チップ・ファイル内の対象・検索条件・高速検索の印）の折り返しを決める。
    # 項目は左から順に並べ、入りきらなくなったら次の行の左端へ落ちる（後ろの項目から落ちる）。
    # 「伸びる空き」は spacerIndex 番の項目の前に置き、その行の余りを全部取る（右の組を右端に寄せる）。
    # widths は各項目の幅（右の間を含む。出していない項目は 0）。available は並べられる幅（右の間を含む）。
    # 返すもの: Lines（行ごとの項目の番号）・SpacerWidth（空きの幅）・LineStarts（各項目が行の先頭か）・
    # OnSpacerLine（各項目が、空きのある行にあるか）
    param (
        [double[]]$widths,
        [int]$spacerIndex,
        [double]$available
    )

    $lines = New-Object System.Collections.Generic.List[object]
    $current = New-Object System.Collections.Generic.List[int]
    $sums = New-Object System.Collections.Generic.List[double]
    $sum = 0.0
    $starts = New-Object bool[] $widths.Count
    for ($i = 0; $i -lt $widths.Count; $i++) {
        if ($widths[$i] -le 0) { continue }
        if ($current.Count -gt 0 -and ($sum + $widths[$i]) -gt $available) {
            $lines.Add($current.ToArray())
            $sums.Add($sum)
            $current = New-Object System.Collections.Generic.List[int]
            $sum = 0.0
        }
        if ($current.Count -eq 0) { $starts[$i] = $true }
        $current.Add($i)
        $sum += $widths[$i]
    }
    if ($current.Count -gt 0) {
        $lines.Add($current.ToArray())
        $sums.Add($sum)
    }

    # 空きは、空きの前の最後の項目がある行に置く。その行の余りを取る（取りすぎて折り返さないよう 0.5 だけ残す）
    $spacer = 0.0
    $spacerLine = $null
    for ($line = 0; $line -lt $lines.Count; $line++) {
        $members = @($lines[$line] | Where-Object { $_ -lt $spacerIndex })
        if ($members.Count -gt 0) { $spacerLine = $line }
    }
    if ($null -ne $spacerLine) {
        $spacer = [Math]::Max(0.0, $available - $sums[$spacerLine] - 0.5)
    }
    $onSpacerLine = New-Object bool[] $widths.Count
    if ($null -ne $spacerLine) {
        foreach ($member in $lines[$spacerLine]) { $onSpacerLine[$member] = $true }
    }
    return @{ Lines = @($lines.ToArray()); SpacerWidth = $spacer; LineStarts = $starts; OnSpacerLine = $onSpacerLine }
}
