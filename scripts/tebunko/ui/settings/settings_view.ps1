# 設定の画面の判断（表示の文言・選んだワークスペースの可否・確認ダイアログの中身）。
# ワークスペースは、インデックス・取り込み一覧・ログを置くフォルダ（$workDir。既定は %USERPROFILE%\Documents\tebunko_ws）。
# 画面に触らないため、そのままテストできる（tests\tebunko\ui\settings\settings_view.Tests.ps1）。

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
    # 今のインデックスの中に置くと、ワークスペースの中身がインデックスとして検索される（前の版の index\ が残っている場合も同じ）。
    # フォルダ名は Workspace（core\workspace.ps1）の 1 か所から取る
    $currentWorkspace = [Workspace]::new($current)
    if ((testFolderUnder $folder $currentWorkspace.IndexDir $drives) -or (testFolderUnder $folder $currentWorkspace.LegacyIndexDir $drives)) {
        return @{ Kind = "error"; Message = "「${folder}」は今のインデックスのフォルダの中です。インデックスの外のフォルダを選んでください。" }
    }
    if (-not $writable) {
        return @{ Kind = "error"; Message = "「${folder}」にはファイルを作れません。書き込めるフォルダを選んでください。" }
    }
    return @{ Kind = "ok"; Message = "" }
}

function newWorkspaceConfirm {
    # ワークスペースを変える前の確認ダイアログの中身:
    #   @{ Title; Heading; Facts（@{ Kind = "next" / "kept" / "warn"; Title; Detail } の配列）; Hint; Choices（@{ Text; Value; Danger; Careful } の配列。最後が主なボタン） }
    # ワークスペースには空のフォルダを選んでもらう。空でなければ警告し、中にワークスペースのフォルダを作るか、そのまま使うかを選ばせる。
    # 今のワークスペースの中身（インデックス・取り込み一覧・ログ）は、いつも新しいワークスペースへ移す（moveWorkspace）。
    # 選んだフォルダに tebunko のファイル（インデックスなど。ほかの人が共有したワークスペースなど）があれば、それを使うか、消して最初からやるかを選ばせる。
    #   Value: "change"（選んだフォルダにする）/ "sub"（中に workspace を作ってそこにする）/ "asis"（空でないまま使う）/
    #          "use"（中のインデックスを使う。今のワークスペースの中身は移さない）/ "reset"（中のインデックスを消し、今のワークスペースの中身を移す）
    param (
        [string]$folder,       # 選んだフォルダ
        [string]$current,      # 今のワークスペース
        [int]$entryCount,      # 選んだフォルダの中のファイル・フォルダの数（数えた上限で止めてよい）
        [string[]]$sampleNames = @(),  # 中身の例（先頭の数件の名前）
        [bool]$countCapped = $false,   # entryCount が数えた上限（それ以上あるかもしれない）
        [string[]]$workspaceNames = @(), # 中にある tebunko のファイル・フォルダの名前（getWorkspaceEntries）。あれば、使うか消すかを選ばせる
        [bool]$canMakeSub = $true,     # 中に workspace を作れる（無いか、あっても空）
        [bool]$toDefault = $false      # ［既定に戻す］から（空のときだけ、題と文言が変わる）
    )

    $moveHint = "移動中は、検索とインデックスの更新はできません。"
    if (@($workspaceNames).Count -gt 0) {
        return @{
            Title   = "保存先の変更"
            Heading = "「${folder}」には、すでにインデックスがあります。"
            Facts   = @()
            Hint    = "そのフォルダのインデックスを使うか、今のインデックスを移動するかを選んでください。移動すると、そのフォルダにあるインデックスは削除されます（元に戻せません）。"
            Choices = @(
                @{ Text = "今のインデックスを移動する"; Value = "reset"; Danger = $true; Careful = $true },
                @{ Text = "そのフォルダのインデックスを使う"; Value = "use"; Careful = $false })
        }
    }
    if ($entryCount -le 0) {
        if ($toDefault) {
            return @{
                Title   = "既定の場所に戻す"
                Heading = "インデックスとログを既定の場所「${folder}」へ移動します。"
                Facts   = @()
                Hint    = $moveHint
                Choices = @(@{ Text = "戻す"; Value = "change"; Careful = $false })
            }
        }
        return @{
            Title   = "保存先の変更"
            Heading = "インデックスとログを「${folder}」へ移動します。"
            Facts   = @()
            Hint    = $moveHint
            Choices = @(@{ Text = "移動する"; Value = "change"; Careful = $false })
        }
    }

    $countText = if ($countCapped) { "{0:#,0} 個以上" -f $entryCount } else { "{0:#,0} 個" -f $entryCount }
    $sample = @($sampleNames | Select-Object -First ${workspaceSampleCount}) -join "、"
    if ($entryCount -gt @($sampleNames).Count -or $countCapped) {
        $sample += " など"
    }
    $facts = @(@{ Kind = "warn"; Title = "このフォルダは空ではありません（ファイル・フォルダが ${countText}）"; Detail = $sample })

    $facts += @{ Kind = "next"; Title = "今のインデックス・ログは、新しい場所へ移します"; Detail = "移す前の場所：${current}" }

    # 最後の選択肢が主なボタン（青）になる。うっかり押しやすい「そのまま使う」を先に置く
    $choices = @(@{ Text = "このフォルダのまま使う"; Value = "asis"; Careful = $true })
    $sub = Join-Path $folder ${workspaceSubFolderName}
    if ($canMakeSub) {
        $choices += @{ Text = "中に「${workspaceSubFolderName}」を作って使う"; Value = "sub"; Careful = $false }
        $facts += @{ Kind = "next"; Title = "「中に作って使う」を選ぶと、ワークスペースは次の場所になります"; Detail = $sub }
    }
    return @{
        Title   = "保存先の変更"
        Heading = "選んだフォルダは空ではありません。ワークスペースには空のフォルダを選んでください。"
        Facts   = $facts
        Hint    = "空のフォルダを選び直すときは［キャンセル］を押してください。"
        Choices = $choices
    }
}
