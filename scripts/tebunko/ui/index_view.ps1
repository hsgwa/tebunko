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
    # 名前が既にあるインデックスと重なることは断らない（上書き・別名・取りやめの確認に回す。getImportIndexName）
    param (
        [string]$folder,   # 入力された元のフォルダ
        [string]$name      # 入力されたインデックス名
    )

    if ((normalizeFolderPath $folder) -eq "") {
        return "元のフォルダを指定してください。"
    }
    return (testIndexName $name.Trim() @())
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
    foreach ($other in $items) {
        if ($other -eq $current) {
            continue
        }
        if (testSameFolder $other.Path $folder) {
            return "「${folder}」のインデックス [$($other.Name)] が既にあります。"
        }
        # 入れ子のフォルダは、同じファイルが2つのインデックスに入り、取り込みも検索結果も二重になるため登録しない
        if (testFolderUnder $folder $other.Path) {
            return "「${folder}」は、インデックス [$($other.Name)]（$($other.Path)）の中のフォルダです。" +
                "同じファイルが二重に取り込まれるため、登録できません。検索する範囲を絞るときは［2 検索］の検索対象で外してください。"
        }
        if (testFolderUnder $other.Path $folder) {
            return "「${folder}」の中には、インデックス [$($other.Name)]（$($other.Path)）があります。" +
                "同じファイルが二重に取り込まれるため、登録できません。まとめるときは、先に [$($other.Name)] を削除してください。"
        }
    }
    # @(getUsedIndexNames ...) と直接書くと集合が 1 要素の配列に入るだけなので、変数に受けてから配列にする
    $usedNames = getUsedIndexNames $items $current
    return (testIndexName $name.Trim() @($usedNames))
}
