# ［8 設定］タブの判断（表示の文言・選んだ置き場所の可否・確認ダイアログの中身）。
# 画面に触らないため、そのままテストできる（tests\tebunko_grep\ui\settings_view.Tests.ps1）。

function getWorkDirView {
    # インデックスの置き場所（work）の表示: @{ Path; Note（既定かどうかの補足）; CanReset（既定の場所でなければ $true。［既定に戻す］を出す） }
    param (
        [string]$workDir,
        [string]$defaultDir
    )

    $isDefault = $workDir.TrimEnd("\").Equals($defaultDir.TrimEnd("\"), [System.StringComparison]::OrdinalIgnoreCase)
    $note = if ($isDefault) { "既定の場所（設定ファイルと同じフォルダの work）です。" } else { "既定の場所は「${defaultDir}」です。" }
    return @{ Path = $workDir; Note = $note; CanReset = -not $isDefault }
}

function getSettingsFileView {
    # 設定ファイルの場所の表示: @{ Path; Note（ツールのフォルダに置いているか） }
    param (
        [string]$settingsFile,
        [string]$rootDir
    )

    $dir = [System.IO.Path]::GetDirectoryName($settingsFile)
    if ($dir.TrimEnd("\").Equals($rootDir.TrimEnd("\"), [System.StringComparison]::OrdinalIgnoreCase)) {
        return @{ Path = $settingsFile; Note = "ツールのフォルダに置いています。" }
    }
    return @{ Path = $settingsFile; Note = "ツールのフォルダ（${rootDir}）に書き込めないため、利用者ごとの場所に置いています。" }
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