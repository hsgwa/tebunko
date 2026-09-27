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
