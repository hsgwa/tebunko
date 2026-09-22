# ［1 インデックス管理］タブの判断（入力の検査・名前の重複）。
# 画面に触らないため、そのままテストできる（tests\tebunko_grep\ui\index_view.Tests.ps1）。

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

function getWorkDirView {
    # インデックスの置き場所（work）の表示: @{ Text; CanReset（既定の場所でなければ $true。［既定に戻す］を出す） }
    param (
        [string]$workDir,
        [string]$defaultDir
    )

    $isDefault = $workDir.TrimEnd("\").Equals($defaultDir.TrimEnd("\"), [System.StringComparison]::OrdinalIgnoreCase)
    $text = "インデックスの置き場所：${workDir}"
    if ($isDefault) {
        $text += "（既定）"
    }
    return @{ Text = $text; CanReset = -not $isDefault }
}

function testWorkFolderChoice {
    # 選んだフォルダを work の置き場所にしてよいかを返す: @{ Kind; Message }
    #   Kind: "same"（今と同じ。何もしない）/ "error"（使えない。Message に理由）/ "ok"
    param (
        [string]$folder,     # 選んだフォルダ
        [string]$current,    # 今の work の置き場所
        [bool]$writable,     # 選んだフォルダにファイルを作れるか（testWritableFolder）
        $drives = $null      # ドライブ文字 → 割り当て先（テストで差し替える。$null は getDriveTargets）
    )

    $folder = normalizeFolderPath $folder
    if ($folder -eq "") {
        return @{ Kind = "error"; Message = "フォルダを選んでください。" }
    }
    if (testSameFolder $folder $current $drives) {
        return @{ Kind = "same"; Message = "" }
    }
    # 今のインデックスの中に置くと、work の中身がインデックスとして検索される
    if (testFolderUnder $folder (Join-Path $current "index") $drives) {
        return @{ Kind = "error"; Message = "「${folder}」は今のインデックスのフォルダの中です。インデックスの外のフォルダを選んでください。" }
    }
    if (-not $writable) {
        return @{ Kind = "error"; Message = "「${folder}」にはファイルを作れません。書き込めるフォルダを選んでください。" }
    }
    return @{ Kind = "ok"; Message = "" }
}

function newWorkFolderConfirm {
    # 置き場所を変える前の確認ダイアログの中身: @{ Heading; Facts（@{ Kind = "next" / "kept"; Title; Detail } の配列）; Hint; ChoiceText }
    param (
        [string]$folder,     # 新しい置き場所
        [string]$current,    # 今の置き場所
        [bool]$hasIndex      # 新しい置き場所にインデックス（index フォルダ）があるか
    )

    $facts = @(@{ Kind = "next"; Title = "インデックス・取り込み一覧・ログを、このフォルダに置きます"; Detail = $folder })
    if ($hasIndex) {
        $facts += @{ Kind = "kept"; Title = "このフォルダにあるインデックスを、そのまま使います"; Detail = "" }
    } else {
        $facts += @{ Kind = "next"; Title = "このフォルダにはまだインデックスがありません"; Detail = "［インデックス作成を開始］で、一覧のフォルダを取り込み直します" }
    }
    $facts += @{ Kind = "kept"; Title = "今の場所のインデックスは、そのまま残ります（移しません）"; Detail = $current }
    return @{
        Heading    = "インデックスの置き場所を変えますか？"
        Facts      = $facts
        Hint       = "今のインデックスを移して使うときは、先にエクスプローラーで「${current}」の中身を新しいフォルダへ移してから選んでください。"
        ChoiceText = "置き場所を変えて、画面を開き直す"
    }
}