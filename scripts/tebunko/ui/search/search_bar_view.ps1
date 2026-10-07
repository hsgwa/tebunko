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
    # 全部外してよい（空は「1 つも選んでいない」。検索を始めるときに getSearchKindError が知らせる）
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

function getSearchKindError {
    # 検索を始められない理由（種類のチップを 1 つも選んでいないとき）の文言。始められるときは空
    param (
        [object[]]$kinds
    )

    $chosen = @(${fileKindNames} | Where-Object { @($kinds) -contains $_ })
    if ($chosen.Count -eq 0) {
        return "検索する種類を 1 つ以上選んでください。"
    }
    return ""
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

function getTargetHintText {
    # 左の欄の下に出す案内。インデックスはあるが、検索の対象にするフォルダを 1 つも選んでいないときだけ出す（出さないときは空）
    param (
        [int]$total,
        [int]$targetCount
    )

    if ($total -gt 0 -and $targetCount -le 0) {
        return "検索するフォルダを選んでください"
    }
    return ""
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

function getFastSearchView {
    # 検索ワードの下に出す、高速検索（Windows Search で先に絞る）の使用可否。
    #   available: Windows Search が使えるか（testWindowsSearch）。$null はまだ確かめていない（使えるものとして扱う）
    # Tip は、使えない理由（ツールチップに出す。分からない・理由が無いときは空）。上から順に、最初に当てはまるもの
    param (
        $available,
        [bool]$useRegex,
        [string]$word
    )

    $usable = testFastSearchUsable ($available -ne $false) $useRegex $word
    $tip = ""
    if (!$usable) {
        if ($available -eq $false) {
            $tip = "検索はできますが時間がかかります　［インデックス管理で確認］"
        } elseif ($useRegex) {
            $tip = "正規表現をオフにすると速く検索できます"
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
        [int]$targetCount    # 検索対象に選ばれている数
    )

    if ($searching) {
        return @{ Content = "中止"; Enabled = !$stopping }
    }
    return @{ Content = "検索"; Enabled = ($word -ne "" -and $hasIndex -and $targetCount -gt 0) }
}
