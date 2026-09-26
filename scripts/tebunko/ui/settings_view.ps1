# ［8 設定］タブの判断（表示の文言・選んだワークスペースの可否・確認ダイアログの中身）。
# ワークスペースは、インデックス・取り込み一覧・ログを置くフォルダ（$workDir。既定は %USERPROFILE%\Documents\tebunko_ws）。
# 画面に触らないため、そのままテストできる（tests\tebunko\ui\settings_view.Tests.ps1）。

${workspaceSubFolderName} = "workspace"  # 空でないフォルダを選んだとき、中に作るワークスペースのフォルダ名
${workspaceSampleCount}   = 3            # 空でないフォルダの中身の例として出す数

function getWorkspaceView {
    # ワークスペースの表示: @{ Path; Note（既定かどうかの補足）; CanReset（既定の場所でなければ $true。［既定に戻す］を出す） }
    param (
        [string]$workDir,
        [string]$defaultDir
    )

    $isDefault = $workDir.TrimEnd("\").Equals($defaultDir.TrimEnd("\"), [System.StringComparison]::OrdinalIgnoreCase)
    $note = if ($isDefault) { "既定の場所（ドキュメントの tebunko）です。" } else { "既定の場所は「${defaultDir}」です。" }
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

function testWorkspaceChoice {
    # 選んだフォルダをワークスペースにしてよいかを返す: @{ Kind; Message }
    #   Kind: "same"（今と同じ。何もしない）/ "error"（使えない。Message に理由）/ "ok"
    param (
        [string]$folder,     # 選んだフォルダ
        [string]$current,    # 今のワークスペース
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
    # 今のインデックスの中に置くと、ワークスペースの中身がインデックスとして検索される
    if (testFolderUnder $folder (Join-Path $current "index") $drives) {
        return @{ Kind = "error"; Message = "「${folder}」は今のインデックスのフォルダの中です。インデックスの外のフォルダを選んでください。" }
    }
    if (-not $writable) {
        return @{ Kind = "error"; Message = "「${folder}」にはファイルを作れません。書き込めるフォルダを選んでください。" }
    }
    return @{ Kind = "ok"; Message = "" }
}

function newWorkspaceConfirm {
    # ワークスペースを変える前の確認ダイアログの中身:
    #   @{ Heading; Facts（@{ Kind = "next" / "kept" / "warn"; Title; Detail } の配列）; Hint; Choices（@{ Text; Detail; Value; Careful } の配列） }
    # ワークスペースには空のフォルダを選んでもらう。空でなければ警告し、中にワークスペースのフォルダを作るか、そのまま使うかを選ばせる。
    # 今のワークスペースの中身（インデックス・取り込み一覧・ログ）は、いつも新しいワークスペースへ移す（moveWorkspace）。
    #   Value: "change"（選んだフォルダにする）/ "sub"（中に workspace を作ってそこにする）/ "asis"（空でないまま使う）
    param (
        [string]$folder,       # 選んだフォルダ
        [string]$current,      # 今のワークスペース
        [int]$entryCount,      # 選んだフォルダの中のファイル・フォルダの数（数えた上限で止めてよい）
        [string[]]$sampleNames = @(),  # 中身の例（先頭の数件の名前）
        [bool]$countCapped = $false,   # entryCount が数えた上限（それ以上あるかもしれない）
        [bool]$hasIndex = $false,      # 中にインデックス（index フォルダ）がある（そのまま使うと、今のワークスペースの中身を移せない）
        [bool]$canMakeSub = $true      # 中に workspace を作れる（無いか、あっても空）
    )

    $moveCurrent = @{ Kind = "next"; Title = "今のワークスペースの中身（インデックス・取り込み一覧・ログ）は、新しいワークスペースへ移します"; Detail = "移す前の場所：${current}" }
    if ($entryCount -le 0) {
        return @{
            Heading = "ワークスペースを変えますか？"
            Facts   = @(
                @{ Kind = "next"; Title = "インデックス・取り込み一覧・ログを、このフォルダに置きます"; Detail = $folder },
                $moveCurrent)
            Hint    = ""
            Choices = @(@{ Text = "ワークスペースを変える"; Detail = ""; Value = "change"; Careful = $false })
        }
    }

    $countText = if ($countCapped) { "{0:#,0} 個以上" -f $entryCount } else { "{0:#,0} 個" -f $entryCount }
    $sample = @($sampleNames | Select-Object -First ${workspaceSampleCount}) -join "、"
    if ($entryCount -gt @($sampleNames).Count -or $countCapped) {
        $sample += " など"
    }
    $facts = @(@{ Kind = "warn"; Title = "このフォルダは空ではありません（ファイル・フォルダが ${countText}）"; Detail = $sample })
    if ($hasIndex) {
        $facts += @{ Kind = "warn"; Title = "インデックス（index フォルダ）があります"; Detail = "このフォルダをそのまま使うと、今のワークスペースの中身を移せません" }
    }
    $facts += $moveCurrent

    $choices = @()
    $sub = Join-Path $folder ${workspaceSubFolderName}
    if ($canMakeSub) {
        $choices += @{ Text = "中に「${workspaceSubFolderName}」フォルダを作って、ワークスペースにする"; Detail = $sub; Value = "sub"; Careful = $false }
    }
    $choices += @{ Text = "このフォルダをそのまま使う"; Detail = "中のファイルは消しません。tebunko のファイルが同じフォルダに混ざります"; Value = "asis"; Careful = $true }
    return @{
        Heading = "選んだフォルダは空ではありません。ワークスペースには空のフォルダを選んでください。"
        Facts   = $facts
        Hint    = "空のフォルダを選び直すときは［キャンセル］を押してください。"
        Choices = $choices
    }
}
