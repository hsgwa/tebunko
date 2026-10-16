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

function getWorkspaceCautionTexts {
    # ワークスペースの説明。注意（文字がそのまま入る・更新中は変えられない）を先に、変え方をあとに出す。@{ Caution; HowTo }
    return @{
        Caution = "インデックスには文書の文字がそのまま入ります。アクセス権を絞ったフォルダや、容量のあるドライブを選べます。更新中は変えられません。"
        HowTo   = "空のフォルダを選んでください。変えるとすぐに切り替わり、今のワークスペースの中身（インデックス・ログ）は新しいフォルダへ移します。"
    }
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
    return @{ Path = $settingsFile; Note = "ツールのフォルダ（${rootDir}）に書き込めないため、既定のワークスペースに置いています。" }
}

function getCountedEntryPaths {
    # ワークスペースに選んだフォルダの中身のうち、「空でない」の数に入れるものを返す。
    # 選んだフォルダが既定のワークスペースのときは、設定ファイルと、それに付いてできるファイル（testSettingsFileName）を数えない
    # （書き込めない場所に置いた tebunko の設定ファイルが、既定のワークスペースの直下にあるため）
    param (
        [string[]]$entryPaths,
        [bool]$isDefaultWorkspace
    )

    if (-not $isDefaultWorkspace) {
        return @($entryPaths)
    }
    return @($entryPaths | Where-Object { -not (testSettingsFileName ([System.IO.Path]::GetFileName($_))) })
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
    # 今のインデックスの中に置くと、ワークスペースの中身がインデックスとして検索される（前の版の content_index\・index\ が残っている場合も同じ）。
    # フォルダ名は Workspace（core\workspace.ps1）の 1 か所から取る
    $currentWorkspace = [Workspace]::new($current)
    if ((testFolderUnder $folder $currentWorkspace.StoreDir $drives) -or (testFolderUnder $folder $currentWorkspace.OldIndexDir $drives) -or (testFolderUnder $folder $currentWorkspace.LegacyIndexDir $drives)) {
        return @{ Kind = "error"; Message = "「${folder}」は今のインデックスのフォルダの中です。インデックスの外のフォルダを選んでください。" }
    }
    if (-not $writable) {
        return @{ Kind = "error"; Message = "「${folder}」にはファイルを作れません。書き込めるフォルダを選んでください。" }
    }
    return @{ Kind = "ok"; Message = "" }
}

function getWorkspaceInputDecision {
    # 設定のワークスペースの欄に入力したパスを確定するときの判断。@{ Action; Message; Path }
    #   Action: None（変わっていない。欄を今の場所に戻す）/ Invalid（直してほしい。Message に理由）/ Blocked（今は変えられない。Message に理由）/ Apply（Path で、［変更…］で選んだときと同じ確かめ・切り替えに進む）
    # 入力した文字のまま判断し、ドライブ文字で始まるローカルのドライブのパスだけを進める（共有に触れうる整形は、そのあとで行う）。
    # 相対パス・環境変数（%…%）・ネットワークの場所・デバイス名の形（\\.\ \\?\）は受けない
    # （画面のスレッドでネットワークに触らないため。ネットワークは［変更…］で選んでもらう）。
    # フォルダの有無・書き込めるかは、［変更…］で選んだときと同じ testWorkspaceChoice（applyWorkspace）が調べる
    param (
        [string]$text,       # 入力された文字
        [string]$current,    # 今のワークスペース
        $drives = $null,     # ドライブ文字 → 割り当て先（テストで差し替える）
        [scriptblock]$driveType = $null,  # ドライブの種類を引く関数（テストで差し替える。null なら実際のドライブ）
        [string]$blockedMessage = ""      # 今は変えられない理由（更新中など。空なら変えられる）
    )

    $trimmed = $text.Trim().Trim([char]34).Trim()
    if ($trimmed -eq "") {
        return @{ Action = "Invalid"; Message = "ワークスペースにするフォルダのパスを入力してください。"; Path = $current }
    }
    $startMessage = "ドライブ文字から始まるパス（例 D:\tebunko_ws）で入力してください。"
    $networkMessage = "ネットワークの場所（\\サーバー\共有 など）や、つながっていないドライブは、入力では指定できません。［変更…］でフォルダを選んでください。"
    $typed = $trimmed.Replace("/", "\")
    if ($typed.Contains("%")) {
        return @{ Action = "Invalid"; Message = "% を含むパス（環境変数）は入力では指定できません。パスをそのまま入力するか、［変更…］でフォルダを選んでください。"; Path = $current }
    }
    if ($typed.StartsWith("\\?\") -or $typed.StartsWith("\\.\")) {
        return @{ Action = "Invalid"; Message = $startMessage; Path = $current }
    }
    if ($typed.StartsWith("\\")) {
        return @{ Action = "Invalid"; Message = $networkMessage; Path = $current }
    }
    if ($typed -notmatch '^[A-Za-z]:\\') {
        return @{ Action = "Invalid"; Message = $startMessage; Path = $current }
    }
    $isLocal = if ($null -ne $driveType) { testLocalDrivePath $typed $driveType } else { testLocalDrivePath $typed }
    if (!$isLocal) {
        return @{ Action = "Invalid"; Message = $networkMessage; Path = $current }
    }
    # ローカルのドライブと分かってから整形する（短い名前の展開などで共有に触れないため）
    $folder = normalizeFolderPath $typed
    # 予約デバイス名（con・nul など）は、整形でドライブ文字の形でなくなることがある。進めない
    if ($folder -notmatch '^[A-Za-z]:\\') {
        return @{ Action = "Invalid"; Message = "con・nul など、Windows が機器の名前として使う名前は、フォルダの名前にできません。別の名前で入力してください。"; Path = $current }
    }
    if (testSameFolder $folder $current $drives) {
        return @{ Action = "None"; Message = ""; Path = $current }
    }
    if ($blockedMessage -ne "") {
        return @{ Action = "Blocked"; Message = $blockedMessage; Path = $current }
    }
    return @{ Action = "Apply"; Message = ""; Path = $folder }
}

function newWorkspaceConfirm {
    # ワークスペースを変える前の確認ダイアログの中身:
    #   @{ Title; Heading; Facts（@{ Kind = "next" / "kept" / "warn"; Title; Detail } の配列）; Hint; Choices（@{ Text; Value; Danger; Careful } の配列。最後が主なボタン） }
    # ワークスペースには空のフォルダを選んでもらう。空でなければ警告し、中にワークスペースのフォルダを作るか、そのまま使うかを選ばせる。
    # 今のワークスペースの中身（インデックス・取り込み一覧・ログ）は、いつも新しいワークスペースへ移す（WorkspaceMover.ChangeWorkspace）。
    # 選んだフォルダに tebunko のファイル（インデックスなど。ほかの人が共有したワークスペースなど）があれば、それを使うか、消して最初からやるかを選ばせる。
    #   Value: "change"（選んだフォルダにする）/ "sub"（中に workspace を作ってそこにする）/ "asis"（空でないまま使う）/
    #          "use"（中のインデックスを使う。今のワークスペースの中身は移さない）/ "reset"（中のインデックスを消し、今のワークスペースの中身を移す）
    param (
        [string]$folder,       # 選んだフォルダ
        [string]$current,      # 今のワークスペース
        [int]$entryCount,      # 選んだフォルダの中のファイル・フォルダの数（数えた上限で止めてよい）
        [string[]]$sampleNames = @(),  # 中身の例（先頭の数件の名前）
        [bool]$countCapped = $false,   # entryCount が数えた上限（それ以上あるかもしれない）
        [string[]]$workspaceNames = @(), # 中にある tebunko のファイル・フォルダの名前（WorkspaceMover.ListEntries）。あれば、使うか消すかを選ばせる
        [bool]$canMakeSub = $true,     # 中に workspace を作れる（無いか、あっても空）
        [bool]$toDefault = $false      # ［既定に戻す］から（空のときだけ、題と文言が変わる）
    )

    $moveHint = "移動中は、検索とインデックスの更新はできません。"
    if (@($workspaceNames).Count -gt 0) {
        return @{
            Title   = "ワークスペースの変更"
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
            Title   = "ワークスペースの変更"
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
        Title   = "ワークスペースの変更"
        Heading = "選んだフォルダは空ではありません。ワークスペースには空のフォルダを選んでください。"
        Facts   = $facts
        Hint    = "空のフォルダを選び直すときは［キャンセル］を押してください。"
        Choices = $choices
    }
}

function getWorkspaceMoveChange {
    # 裏のスレッドの仕事の戻り値（output）から、WorkspaceMover.ChangeWorkspace の結果（@{ Ok; Stage; Message; Targets; Count }）を取り出す。無ければ $null
    param (
        $output
    )

    $found = $null
    foreach ($item in @($output)) {
        if ($item -is [hashtable] -and $item.ContainsKey("Ok")) {
            $found = $item
        }
    }
    return $found
}

function getWorkspaceMoveOutcome {
    # ワークスペースを移す裏の仕事が終わったときに、画面のスレッドで行うことを決める。画面に触らない。
    #   change: getWorkspaceMoveChange の結果（無い・例外のときは $null）
    #   errorText: 裏のスレッドの例外の文（無ければ空）
    #   closeRequested: 移している間に窓を閉じようとして、終わるのを待っていた（成功のときだけ、終わったら閉じる。失敗のときは閉じずに警告を出す）
    # 戻り値: @{ Ok（移せた）; Warning（失敗のとき、利用者に出す文。成功なら空）; Status（成功のとき switchWorkspace のあとに出すステータス。空なら出さない）; Close（終わったら閉じる） }
    # 成功のときだけ画面を切り替える（Ok）。失敗のときは設定も画面も変えない（WorkspaceMover は失敗したら設定を変えずに戻している）
    param (
        $change,
        [string]$errorText,
        [bool]$closeRequested,
        [string]$mode,
        [string]$folder,
        [string]$previous
    )

    # 失敗のときは閉じない（閉じるのを待っていても、失敗を利用者に伝えるため、画面を残して警告を出す）
    $failed = @{ Ok = $false; Warning = ""; Status = ""; Close = $false }
    # 移せた（設定は書き換え済み）なら、あとの例外の文があっても成功として扱う（設定と画面を食い違わせない）
    $moved = $null -ne $change -and $change.Ok
    if ($errorText -and !$moved) {
        $failed.Warning = "ワークスペースを「${folder}」へ移せませんでした（${errorText}）。今のワークスペースのまま使えます。"
        return $failed
    }
    if ($null -eq $change) {
        $failed.Warning = "ワークスペースを「${folder}」へ移せませんでした。今のワークスペースのまま使えます。"
        return $failed
    }
    if (!$change.Ok) {
        $failed.Warning = if ($change.Stage -eq "remove") {
            "「${folder}」のインデックスを削除できませんでした（$($change.Message)）。ファイルを開いているアプリを閉じてから、もう一度変えてください。"
        } else {
            [string]$change.Message
        }
        return $failed
    }
    $status = ""
    if ($mode -eq "use") {
        $status = "ワークスペースを「${folder}」に変え、そこにあるインデックスを使います（インデックス $($change.Targets) 件。前のワークスペースの中身は「${previous}」に残しています）"
    } elseif ($change.Count -gt 0) {
        $status = "ワークスペースを「${folder}」に変え、中身を移しました（移す前の場所：${previous}）"
    }
    return @{ Ok = $true; Warning = ""; Status = $status; Close = $closeRequested }
}
