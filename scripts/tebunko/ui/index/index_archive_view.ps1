# ［インデックス管理］タブの判断: エクスポート・インポート。
# 画面に触らないため、そのままテストできる（tests\tebunko\ui\index\index_archive_view.Tests.ps1）。

function getIndexArchiveCancelledStatus {
    # エクスポート・インポートのダイアログを取りやめたときのステータス
    param (
        [string]$operation   # "エクスポート" / "インポート"
    )

    return "${operation}を取りやめました"
}

function getImportResultStatus {
    # インポートの結果（importIndex の戻り値）から、ステータスに出す文言を返す。
    # Warnings（同じ元のフォルダが別の名前で既に登録されている等）を添える
    param (
        $result
    )

    $text = "インデックス [$($result.Name)] をインポートしました（$($result.Files) ファイル）。"
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

function getIndexExportMissingText {
    # エクスポートする大きさを数えている間に、対象のインデックスが無くなっていたときのステータス
    return "対象のインデックスが見つかりません"
}

function getIndexExportHeader {
    # エクスポートのダイアログの上の枠（対象・大きさ・zip の名前）と注意。
    #   items: 書き出すインデックス（Name と FileCountText（"2,530"。無ければ "－" か空）を持つ行）。1 件でも 2 件以上でもよい
    # 戻り値: @{ Title（1 件は名前、2 件以上は「N 件のインデックス」）; Size（「2,530 ファイル」。大きさを数え終えていれば「2,530 ファイル・94.0 MB」。数が分からなければ空）;
    #            FileText（書き出し先の下の「<名前>_インデックス_<yyyyMMdd>.zip として保存します」）; Caution }
    param (
        [object[]]$items,
        [datetime]$now = (Get-Date),
        [long]$bytes = -1   # 数え終えた zip に入る大きさ。省略（負）なら大きさは出さない
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
    if ($known -and $bytes -ge 0) {
        $size += "・" + (formatArchiveSize $bytes)
    }
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

function getImportCollisionChoiceView {
    # インポートのダイアログの頭に出す、同じ名前のときの選び方（［別の名前にする］／［置き換える］）の表示。
    #   name     : 名前の欄の今の値
    #   zipName  : zip の中のインデックスの名前
    #   usedNames: getUsedIndexNames の結果
    #   mode     : 利用者が選んでいる方（Rename か Overwrite。既定は Rename）
    # 出すのは、名前の欄が今の一覧にある名前のとき、または zip の名前が一覧にあって名前の欄が別の名前の候補のまま（初めの状態）のとき。
    # 名前の欄を手で一覧に無い名前にしたら隠す。
    # 置き換える相手（TargetName）は、名前の欄が一覧にある名前ならその名前、そうでなければ zip の名前。
    # Mode は、置き換える相手が名前の欄で決まっている（名前の欄が一覧にある名前の）ときだけ Overwrite になれる。
    # 名前の欄が一覧に無い名前なら、選んでいても Rename（重ならないので別の名前の扱いと同じ）。
    # Danger は［インポート］を危険の型（赤）にするか（Mode が Overwrite のとき）。Notice は置き換えるときに添える一文。
    # Candidate は［別の名前にする］を選んだときに名前の欄へ入れる候補、OverwriteName は［置き換える］を選んだときの名前の欄の値
    # listedNames は今の一覧の名前（省くと、usedNames の名前はすべて一覧にあるものとして扱う）。
    # 置き換える相手が、一覧に無く作業フォルダに残っているだけの名前（設定から外した・置き場にだけある）のときは、置き換えを出さない
    # （OverwriteVisible = $false。選んでいても Rename）。見出しは「今のインデックス」と書かず、残っていることを伝える
    param (
        [string]$name,
        [string]$zipName,
        $usedNames = $null,
        [string]$mode = ${importCollisionRename},
        $listedNames = $null
    )

    $trimmed = $name.Trim()
    $candidate = getImportSuggestedName $zipName $usedNames
    $nameCollides = $trimmed -ne "" -and (testImportNameCollision $trimmed $usedNames)
    $zipCollides = testImportNameCollision $zipName $usedNames
    $visible = $nameCollides -or ($zipCollides -and $trimmed -ieq $candidate)
    $target = if ($nameCollides) { $trimmed } else { $zipName }
    $targetListed = ($null -eq $listedNames) -or (testImportNameCollision $target $listedNames)
    $effective = if ($nameCollides -and $targetListed -and $mode -eq ${importCollisionOverwrite}) { ${importCollisionOverwrite} } else { ${importCollisionRename} }
    $heading = if (!$visible) { "" } elseif ($targetListed) { "同じ名前のインデックス [${target}] があります" } else { "作業フォルダに同じ名前のインデックス [${target}] が残っています（置き換えはできません）" }
    return @{
        Visible       = $visible
        Mode          = $effective
        Danger        = ($effective -eq ${importCollisionOverwrite})
        Heading       = $heading
        RenameText    = "別の名前にする"
        OverwriteText = "置き換える"
        OverwriteVisible = $targetListed
        Notice        = if ($visible -and $effective -eq ${importCollisionOverwrite}) { "今のインデックス [${target}] は、インポートするもので置き換わります。" } else { "" }
        TargetName    = $target
        Candidate     = $candidate
        OverwriteName = $target
    }
}

function testIndexImportInput {
    # インポートのダイアログの入力を調べ、直してほしい内容を返す（問題なければ空文字列）。
    # 名前が既にあるインデックスと重なることは断らない（別の名前にするか置き換えるかの選びに回す。getImportCollisionChoiceView）。
    # 元のフォルダは、追加と詳細欄の編集と同じ決まり（getIndexFolderConflict）で調べる。同じ名前の行は上書きで置き換わる
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
    # 入力された名前が今の一覧にあり、別の名前にするか置き換えるかの選びが要るか（大文字・小文字は区別しない）
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
    return "名前と、元のフォルダの場所を変えられます。"
}

function getImportProgressView {
    # インポートの進み具合（展開した大きさ Done・全体の大きさ Total。バイト）から、帯に出す内容を返す。
    # 戻り値: @{ Text; Detail; Value（0 から 1）; Indeterminate }。
    # 全体が分からない間（目録の確かめ・空き容量の確かめ）と、展開が終わった後（入れ替え・設定の書き込み）は、割合の分からない表示にする
    param (
        [long]$done,
        [long]$total
    )

    if ($total -le 0) {
        return @{ Text = "インポートしています…"; Detail = ""; Value = 0.0; Indeterminate = $true }
    }
    if ($done -ge $total) {
        return @{ Text = "インポートを仕上げています…"; Detail = ""; Value = 1.0; Indeterminate = $true }
    }
    $ratio = [double]$done / [double]$total
    return @{
        Text = "インポートしています（$([int][Math]::Floor($ratio * 100))%）"
        Detail = "$(formatArchiveSize $done) / $(formatArchiveSize $total)"
        Value = $ratio
        Indeterminate = $false
    }
}

function getImportIndexingConfirm {
    # インポートが終わった後に、続けて更新するかを尋ねる文（importIndex の戻り値から）。戻り値: @{ Title（窓の題。インポートが終わったことを伝える）; Heading; Hint; CreateText; CancelText }。
    # 元のフォルダがこの PC に無いとき（Enabled が偽）は $null（更新を始めると、設定の enabled が真に書き換わってしまうため尋ねない）
    param (
        $result
    )

    if (!$result.Enabled) {
        return $null
    }
    return @{
        Title = "$($result.Name) をインポートしました"
        Heading = "高速化のためのインデックスを作成しますか？"
        Hint = "作成すると検索が速くなります。作成しなくても検索はできます（結果は同じ）。"
        CreateText = "作成する"
        CancelText = "あとで"
    }
}

function getImportSkippedIndexingStatus {
    # 続けて更新しなかったとき（確認を取りやめた・元のフォルダが無くて尋ねなかった）のステータス。
    # getImportResultStatus の文に、高速化のためのインデックスが、次に更新するまでは無いことを足す
    param (
        $result
    )

    return "$(getImportResultStatus $result) 次の更新までは、検索が遅くなることがあります。"
}

function getIndexExportOmittedText {
    # エクスポートで、インポートできない大きさのため zip に入れなかった本文インデックスのファイルがあるときの、ステータスに足す文言。
    # 無ければ空文字列。元のファイルは最大 5 件まで並べ、残りは「ほか M 件」にする
    param (
        $omitted,         # exportIndex の戻り値の Omitted（@{ Path; Books } の配列）
        [long]$maxBytes   # 上限（contentIndexSearchMaxBytes）。文言の「64 MB」はここから作る
    )

    $omitted = @($omitted | Where-Object { $null -ne $_ })
    if ($omitted.Count -eq 0) {
        return ""
    }
    $books = @($omitted | ForEach-Object { @($_.Books) })
    $reason = "$([long]($maxBytes / 1MB)) MB を超えるため"
    $detail = $reason
    if ($books.Count -gt 0) {
        $shown = (@($books | Select-Object -First 5)) -join ", "
        $others = $books.Count - 5
        if ($others -ge 1) {
            $shown += " ほか {0:#,0} 件" -f $others
        }
        $detail = "${reason}。元のファイル: ${shown}"
    }
    return "{0:#,0} 件のインデックスのファイルを含めませんでした（{1}）" -f $omitted.Count, $detail
}


function getIndexExportStatusText {
    # エクスポートの結果のステータス(status)に、含めなかったファイルの知らせ(getIndexExportOmittedText)を「。」でつないで足す。無ければ status のまま
    param (
        [string]$status,
        $omitted,         # exportIndex の戻り値の Omitted
        [long]$maxBytes
    )

    $text = getIndexExportOmittedText $omitted $maxBytes
    if ($text -eq "") {
        return $status
    }
    return "${status}。${text}"
}


function getIndexBulkStatusText {
    # まとめての操作(operation)の結果のステータス(status)を返す。"エクスポート"のときだけ、結果の Omitted(1 件ごと)を集めて知らせを足す(削除などは status のまま)
    param (
        [string]$operation,
        [string]$status,
        [object[]]$results,
        [long]$maxBytes
    )

    if ($operation -ne "エクスポート") {
        return $status
    }
    $omitted = @($results | Where-Object { $null -ne $_ } | ForEach-Object { @($_.Omitted) } | Where-Object { $null -ne $_ })
    return (getIndexExportStatusText $status $omitted $maxBytes)
}
