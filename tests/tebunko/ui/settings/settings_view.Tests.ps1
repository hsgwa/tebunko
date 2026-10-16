# 設定の画面の判断（tebunko\ui\settings\settings_view.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\settings\settings_view.ps1"
}

Describe "getWorkspaceView" -Tag Unit {
    It "既定の場所なら既定だと伝え、［既定に戻す］は出さない" {
        $view = getWorkspaceView "C:\tool\work" "C:\Tool\work\"
        $view.Path | Should -Be "C:\tool\work"
        $view.Note | Should -Be "既定の場所（ドキュメントの tebunko）です。"
        $view.CanReset | Should -Be $false
    }

    It "既定でない場所なら、既定の場所を伝え、［既定に戻す］を出す" {
        $view = getWorkspaceView "D:\データ" "C:\tool\work"
        $view.Path | Should -Be "D:\データ"
        $view.Note | Should -Be "既定の場所は「C:\tool\work」です。"
        $view.CanReset | Should -Be $true
    }
}

Describe "getWorkspaceInputDecision" -Tag Unit {
    BeforeAll { $script:drives = @{ "C" = "\Device\HarddiskVolume3"; "D" = "\Device\HarddiskVolume4" } }

    # blocked: 今は変えられない理由（空なら変えられる）。type: ドライブの種類（Fixed なら、ふつうのローカルのドライブ）
    It "<name>" -TestCases @(
        @{ name = "空にすると入力を促す"; text = "  "; type = "Fixed"; blocked = ""; action = "Invalid"; like = "*入力してください*"; path = "C:\tool\work" }
        @{ name = "相対パスは受けない"; text = "ws\新"; type = "Fixed"; blocked = ""; action = "Invalid"; like = "*ドライブ文字*"; path = "C:\tool\work" }
        @{ name = "ドライブ文字のない絶対パス風（\ 始まり）も受けない"; text = "\ws\新"; type = "Fixed"; blocked = ""; action = "Invalid"; like = "*ドライブ文字*"; path = "C:\tool\work" }
        @{ name = "ドライブ文字だけ（\ 無し）は受けない"; text = "C:"; type = "Fixed"; blocked = ""; action = "Invalid"; like = "*ドライブ文字*"; path = "C:\tool\work" }
        @{ name = "ドライブ文字の相対（C:相対）は受けない"; text = "C:相対"; type = "Fixed"; blocked = ""; action = "Invalid"; like = "*ドライブ文字*"; path = "C:\tool\work" }
        @{ name = "環境変数で始まるパスは受けない（定義の有無を問わず）"; text = "%TEBUNKO_NO_SUCH_VARIABLE%\ws"; type = "Fixed"; blocked = ""; action = "Invalid"; like = "*% を含むパス*［変更…］*"; path = "C:\tool\work" }
        @{ name = "環境変数は展開せず断る（<SystemDrive>）"; text = "%SystemDrive%\新しい"; type = "Fixed"; blocked = ""; action = "Invalid"; like = "*% を含むパス*［変更…］*"; path = "C:\tool\work" }
        @{ name = "% で始まるだけの文字は断る"; text = "%data"; type = "Fixed"; blocked = ""; action = "Invalid"; like = "*% を含むパス*［変更…］*"; path = "C:\tool\work" }
        @{ name = "ドライブ文字の後ろに % があるパスも断る"; text = "D:\%data%"; type = "Fixed"; blocked = ""; action = "Invalid"; like = "*% を含むパス*［変更…］*"; path = "C:\tool\work" }
        @{ name = "今と同じなら何もしない（大文字小文字・末尾の \ の違い）"; text = "c:\TOOL\work\"; type = "Fixed"; blocked = ""; action = "None"; like = ""; path = "C:\tool\work" }
        @{ name = "あるフォルダは確かめに進む"; text = "D:\新しい"; type = "Fixed"; blocked = ""; action = "Apply"; like = ""; path = "D:\新しい" }
        @{ name = "取り外せるドライブも進む"; text = "D:\新しい"; type = "Removable"; blocked = ""; action = "Apply"; like = ""; path = "D:\新しい" }
        @{ name = "引用符と空白を取って進む"; text = " ""D:\新しい"" "; type = "Fixed"; blocked = ""; action = "Apply"; like = ""; path = "D:\新しい" }
        @{ name = "/ 区切りは \ にそろえて進む"; text = "D:/新しい"; type = "Fixed"; blocked = ""; action = "Apply"; like = ""; path = "D:\新しい" }
        @{ name = "末尾の \ と空白は取って進む"; text = "D:\新しい\ "; type = "Fixed"; blocked = ""; action = "Apply"; like = ""; path = "D:\新しい" }
        @{ name = "ネットワークのパスは受けず、［変更…］へ案内する"; text = "\\server\share\ws"; type = "Fixed"; blocked = ""; action = "Invalid"; like = '*\\サーバー\共有*［変更…］*'; path = "C:\tool\work" }
        @{ name = "/ 区切りのネットワークのパスも受けない"; text = "//server/share"; type = "Fixed"; blocked = ""; action = "Invalid"; like = "*［変更…］*"; path = "C:\tool\work" }
        @{ name = "デバイス名の形のネットワークのパス（\\.\UNC\）は受けない"; text = "\\.\UNC\server\share\ws"; type = "Fixed"; blocked = ""; action = "Invalid"; like = "*ドライブ文字*"; path = "C:\tool\work" }
        @{ name = "長いパスの形のネットワークのパス（\\?\UNC\）は受けない"; text = "\\?\UNC\server\share\ws"; type = "Fixed"; blocked = ""; action = "Invalid"; like = "*ドライブ文字*"; path = "C:\tool\work" }
        @{ name = "デバイス名の形のドライブ（\\.\C:\）は受けない"; text = "\\.\C:\ws"; type = "Fixed"; blocked = ""; action = "Invalid"; like = "*ドライブ文字*"; path = "C:\tool\work" }
        @{ name = "長いパスの形のドライブ（\\?\C:\）は受けない"; text = "\\?\C:\ws"; type = "Fixed"; blocked = ""; action = "Invalid"; like = "*ドライブ文字*"; path = "C:\tool\work" }
        @{ name = "割り当てたネットワークのドライブは受けない"; text = "Z:\ws"; type = "Network"; blocked = ""; action = "Invalid"; like = "*［変更…］*"; path = "C:\tool\work" }
        @{ name = "つながっていないドライブは受けない"; text = "Z:\ws"; type = "NoRootDirectory"; blocked = ""; action = "Invalid"; like = "*［変更…］*"; path = "C:\tool\work" }
        @{ name = "更新中などは、進める場所でも Blocked で理由を返す"; text = "D:\新しい"; type = "Fixed"; blocked = "更新中は変えられません。"; action = "Blocked"; like = "更新中は変えられません。"; path = "C:\tool\work" }
        @{ name = "更新中でも、今と同じなら何もしない"; text = "C:\tool\work"; type = "Fixed"; blocked = "更新中は変えられません。"; action = "None"; like = ""; path = "C:\tool\work" }
        @{ name = "更新中でも、直してほしい入力は Invalid"; text = "ws\新"; type = "Fixed"; blocked = "更新中は変えられません。"; action = "Invalid"; like = "*ドライブ文字*"; path = "C:\tool\work" }
    ) {
        $driveType = { param ($drive) [System.IO.DriveType]$type }.GetNewClosure()
        $d = getWorkspaceInputDecision $text "C:\tool\work" $script:drives $driveType $blocked
        $d.Action | Should -Be $action
        if ($like) { $d.Message | Should -BeLike $like }
        $d.Path | Should -Be $path
    }

    # 予約デバイス名は、整形の結果が版で違うことがある。文字列だけを見る（フォルダは作らない）。進めるときは、必ずドライブ文字の形
    It "予約デバイス名（<text>）は、進めるならドライブ文字の形の Path、そうでなければ Invalid" -TestCases @(
        @{ text = "D:\con" }
        @{ text = "D:\ws\nul" }
        @{ text = "D:\con.txt" }
        @{ text = "D:\COM1" }
        @{ text = "D:\con." }
    ) {
        $d = getWorkspaceInputDecision $text "C:\tool\work" $script:drives { param ($drive) [System.IO.DriveType]::Fixed }
        if ($d.Action -eq "Apply") {
            $d.Path | Should -Match '^[A-Za-z]:\\'
        } else {
            $d.Action | Should -Be "Invalid"
            $d.Message | Should -BeLike "*機器の名前*"
        }
    }

    It "整形の結果がドライブ文字の形でなくなる（予約デバイス名）なら、必ず Invalid にして今の場所を返す" {
        Mock normalizeFolderPath { '\\.\con' }
        $d = getWorkspaceInputDecision "D:\con" "C:\tool\work" $script:drives { param ($drive) [System.IO.DriveType]::Fixed }
        $d.Action | Should -Be "Invalid"
        $d.Message | Should -BeLike "*機器の名前*"
        $d.Path | Should -Be "C:\tool\work"
    }
}

Describe "getSettingsFileView" -Tag Unit {
    It "ツールのフォルダにあれば、そう伝える" {
        $view = getSettingsFileView "C:\tool\setting.config" "C:\Tool\"
        $view.Path | Should -Be "C:\tool\setting.config"
        $view.Note | Should -Be "ツールのフォルダに置いています。"
    }

    It "既定のワークスペースにあれば、ツールのフォルダに書き込めないためだと伝える" {
        $view = getSettingsFileView "C:\Users\test\Documents\tebunko_ws\setting.config" "C:\Program Files\tebunko"
        $view.Note | Should -Be "ツールのフォルダ（C:\Program Files\tebunko）に書き込めないため、既定のワークスペースに置いています。"
    }
}

Describe "getCountedEntryPaths（既定のワークスペースの設定ファイルを数えない）" -Tag Unit {
    It "<Case>" -TestCases @(
        @{ Case = "既定のワークスペースなら、設定ファイルと付いてできるファイルを数えない"; IsDefault = $true
           Paths = @("C:\ws\setting.config", "C:\ws\setting.config.tmp", "C:\ws\setting.config.broken-20261003-120000", "C:\ws\setting.config.broken-20261003-120000-2")
           Expected = @() }
        @{ Case = "既定のワークスペースでも、似た名前とほかのファイルは数える"; IsDefault = $true
           Paths = @("C:\ws\setting.config", "C:\ws\setting.config.bak", "C:\ws\README.md")
           Expected = @("C:\ws\setting.config.bak", "C:\ws\README.md") }
        @{ Case = "既定のワークスペースでないフォルダでは、設定ファイルも数える"; IsDefault = $false
           Paths = @("C:\other\setting.config", "C:\other\a.txt")
           Expected = @("C:\other\setting.config", "C:\other\a.txt") }
    ) {
        @(getCountedEntryPaths $Paths $IsDefault) | Should -Be @($Expected)
    }

    It "既定のワークスペースに設定ファイルだけなら、数えた結果は 0 個で「空でない」の警告を出さない" {
        $counted = @(getCountedEntryPaths @("C:\ws\setting.config", "C:\ws\setting.config.tmp") $true)
        $confirm = newWorkspaceConfirm "C:\ws" "C:\old" $counted.Count @() $false @() $true $false
        @($confirm.Facts).Count | Should -Be 0
        $confirm.Choices[0].Value | Should -Be "change"
    }

    It "既定のワークスペースでないフォルダでは、設定ファイルが 1 つあれば「空でない」の警告を出す" {
        $counted = @(getCountedEntryPaths @("C:\other\setting.config") $false)
        $confirm = newWorkspaceConfirm "C:\other" "C:\old" $counted.Count @("setting.config") $false @() $true $false
        @($confirm.Facts | Where-Object { $_.Kind -eq "warn" }).Count | Should -Be 1
    }
}

Describe "getWorkspaceCautionTexts" -Tag Unit {
    It "注意（文字がそのまま入る・更新中は変えられない）を先の段落に、変え方をあとの段落に出す" {
        $texts = getWorkspaceCautionTexts
        $texts.Caution | Should -BeLike "インデックスには文書の文字がそのまま入ります。*更新中は変えられません。"
        $texts.HowTo | Should -BeLike "空のフォルダを選んでください。*"
        $texts.HowTo | Should -Not -BeLike "*更新中*"
    }
}

Describe "testWorkspaceChoice" -Tag Unit {
    It "今と同じ場所（書き方の違いも）なら same" {
        (testWorkspaceChoice "C:\TOOL\work\" "C:\tool\work" $true @{}).Kind | Should -Be "same"
    }

    It "空なら選ぶよう伝える" {
        $result = testWorkspaceChoice "" "C:\tool\work" $true @{}
        $result.Kind | Should -Be "error"
        $result.Message | Should -Be "フォルダを選んでください。"
    }

    It "今のインデックスのフォルダの中は選べない（<sub>）" -TestCases @(
        @{ sub = "content_index\営業" }
        @{ sub = "index\営業" }  # 前の版の index\ の中も同じ
    ) {
        param ($sub)
        $result = testWorkspaceChoice "C:\tool\work\$sub" "C:\tool\work" $true @{}
        $result.Kind | Should -Be "error"
        $result.Message | Should -Match "今のインデックスのフォルダの中です"
    }

    It "書き込めないフォルダは選べない" {
        $result = testWorkspaceChoice "C:\Program Files\共有" "C:\tool\work" $false @{}
        $result.Kind | Should -Be "error"
        $result.Message | Should -Be "「C:\Program Files\共有」にはファイルを作れません。書き込めるフォルダを選んでください。"
    }

    It "書き込めるほかのフォルダなら ok（今の work の中の、インデックスの外も選べる）" {
        (testWorkspaceChoice "D:\データ" "C:\tool\work" $true @{}).Kind | Should -Be "ok"
        (testWorkspaceChoice "C:\tool\work\新しい場所" "C:\tool\work" $true @{}).Kind | Should -Be "ok"
    }
}

Describe "newWorkspaceConfirm" -Tag Unit {
    It "空のフォルダなら、移す先を見出しに、移動中は使えないことを補足に出し、移動するボタンを 1 つ出す" {
        $confirm = newWorkspaceConfirm "D:\データ" "C:\tool\work" 0
        $confirm.Title | Should -Be "ワークスペースの変更"
        $confirm.Heading | Should -Be "インデックスとログを「D:\データ」へ移動します。"
        @($confirm.Facts).Count | Should -Be 0
        $confirm.Hint | Should -Be "移動中は、検索とインデックスの更新はできません。"
        @($confirm.Choices | ForEach-Object { $_.Value }) -join "," | Should -Be "change"
        $confirm.Choices[0].Text | Should -Be "移動する"
    }

    It "既定に戻すときは、題と文言を戻す向けにする" {
        $confirm = newWorkspaceConfirm "C:\Tools\work" "D:\x" 0 @() $false @() $true $true
        $confirm.Title | Should -Be "既定の場所に戻す"
        $confirm.Heading | Should -Be "インデックスとログを既定の場所「C:\Tools\work」へ移動します。"
        $confirm.Choices[0].Text | Should -Be "戻す"
    }

    It "空でなければ警告し、中身の数と例を出し、中に workspace を作るか・そのまま使うかを選ばせる" {
        $confirm = newWorkspaceConfirm "D:\データ" "C:\tool\work" 5 @("見積.xlsx", "報告書", "メモ.txt")
        $confirm.Heading | Should -Be "選んだフォルダは空ではありません。ワークスペースには空のフォルダを選んでください。"
        @($confirm.Facts | ForEach-Object { $_.Kind }) -join "," | Should -Be "warn,next,next"
        $confirm.Facts[0].Title | Should -Be "このフォルダは空ではありません（ファイル・フォルダが 5 個）"
        $confirm.Facts[0].Detail | Should -Be "見積.xlsx、報告書、メモ.txt など"
        @($confirm.Choices | ForEach-Object { $_.Value }) -join "," | Should -Be "asis,sub"
        $confirm.Facts[2].Detail | Should -Be "D:\データ\workspace"
        $confirm.Choices[0].Careful | Should -Be $true
        $confirm.Hint | Should -Be "空のフォルダを選び直すときは［キャンセル］を押してください。"
    }

    It "中身がすべて例に出ていれば「など」を付けず、数え切れなければ「以上」と出す" {
        (newWorkspaceConfirm "D:\データ" "C:\tool\work" 1 @("a.txt")).Facts[0].Detail | Should -Be "a.txt"
        (newWorkspaceConfirm "D:\データ" "C:\tool\work" 1000 @("a", "b", "c") $true).Facts[0].Title | Should -Be "このフォルダは空ではありません（ファイル・フォルダが 1,000 個以上）"
    }

    It "中に workspace を作れなければ、その選択肢を出さない" {
        $confirm = newWorkspaceConfirm "D:\データ" "C:\tool\work" 3 @("a", "b", "c") $false @() $false
        @($confirm.Choices | ForEach-Object { $_.Value }) -join "," | Should -Be "asis"
    }

    It "インデックスなどがあれば、今のを移すか、そのフォルダのを使うかを選ばせる（移すほうは、そのフォルダのインデックスが消えると補足に書く）" {
        $confirm = newWorkspaceConfirm "D:\共有\tebunko_ws" "C:\tool\work" 3 @("index", "ingest_status.tsv", "memo.txt") $false @("index", "ingest_status.tsv")
        $confirm.Heading | Should -Be "「D:\共有\tebunko_ws」には、すでにインデックスがあります。"
        $confirm.Hint | Should -Match "そのフォルダにあるインデックスは削除されます"
        @($confirm.Choices | ForEach-Object { $_.Value }) -join "," | Should -Be "reset,use"
        @($confirm.Choices | ForEach-Object { $_.Text }) -join "," | Should -Be "今のインデックスを移動する,そのフォルダのインデックスを使う"
        $confirm.Choices[0].Danger | Should -Be $true
        $confirm.Choices[0].Careful | Should -Be $true
    }
}

Describe "getWorkspaceMoveChange" -Tag Unit {
    It "<label>" -TestCases @(
        @{ label = "戻り値が空なら null"; output = @(); ok = $null }
        @{ label = "Ok を持つ表だけ拾う"; output = @("x", @{ Other = 1 }, @{ Ok = $true; Count = 2 }); ok = $true }
        @{ label = "複数あれば最後のを拾う"; output = @(@{ Ok = $true }, @{ Ok = $false }); ok = $false }
        @{ label = "Ok の無い表だけなら null"; output = @(@{ Other = 1 }); ok = $null }
    ) {
        param ($label, $output, $ok)
        $change = getWorkspaceMoveChange $output
        if ($null -eq $ok) { $change | Should -BeNullOrEmpty } else { $change.Ok | Should -Be $ok }
    }
}

Describe "getWorkspaceMoveOutcome" -Tag Unit {
    It "<label>は失敗で、画面を切り替えず、今のワークスペースのまま使えると伝える" -TestCases @(
        @{ label = "裏の例外"; change = $null; err = "届きません"; expected = "ワークスペースを「D:\新」へ移せませんでした（届きません）。今のワークスペースのまま使えます。" }
        @{ label = "戻り値なし"; change = $null; err = ""; expected = "ワークスペースを「D:\新」へ移せませんでした。今のワークスペースのまま使えます。" }
        @{ label = "移せなかった結果"; change = @{ Ok = $false; Stage = "move"; Message = "足りません" }; err = ""; expected = "足りません" }
        @{ label = "削除できなかった結果"; change = @{ Ok = $false; Stage = "remove"; Message = "使用中" }; err = ""; expected = "「D:\新」のインデックスを削除できませんでした（使用中）。ファイルを開いているアプリを閉じてから、もう一度変えてください。" }
    ) {
        param ($label, $change, $err, $expected)
        $outcome = getWorkspaceMoveOutcome $change $err $false "move" "D:\新" "C:\旧"
        $outcome.Ok | Should -Be $false
        $outcome.Warning | Should -Be $expected
        $outcome.Status | Should -Be ""
        $outcome.Close | Should -Be $false
    }

    It "成功したら <label>" -TestCases @(
        @{ label = "use は使うインデックスの件数と残した場所を出す"; mode = "use"; change = @{ Ok = $true; Targets = 3; Count = 0 }; expected = "ワークスペースを「D:\新」に変え、そこにあるインデックスを使います（インデックス 3 件。前のワークスペースの中身は「C:\旧」に残しています）" }
        @{ label = "move は移した中身があれば元の場所を出す"; mode = "move"; change = @{ Ok = $true; Targets = 0; Count = 4 }; expected = "ワークスペースを「D:\新」に変え、中身を移しました（移す前の場所：C:\旧）" }
        @{ label = "move でも中身が無ければ出さない"; mode = "move"; change = @{ Ok = $true; Targets = 0; Count = 0 }; expected = "" }
    ) {
        param ($label, $mode, $change, $expected)
        $outcome = getWorkspaceMoveOutcome $change "" $false $mode "D:\新" "C:\旧"
        $outcome.Ok | Should -Be $true
        $outcome.Warning | Should -Be ""
        $outcome.Status | Should -Be $expected
    }

    It "移せたあとに例外の文が付いても、成功として扱う（設定を書き換え済みのため）" {
        $outcome = getWorkspaceMoveOutcome @{ Ok = $true; Targets = 0; Count = 4 } "あとの例外" $false "move" "D:\新" "C:\旧"
        $outcome.Ok | Should -Be $true
        $outcome.Warning | Should -Be ""
    }

    It "移している間に閉じようとしていたら、成功のときだけ Close を返す（失敗は閉じずに警告を出す）" {
        (getWorkspaceMoveOutcome @{ Ok = $true; Count = 1 } "" $true "move" "D:\新" "C:\旧").Close | Should -Be $true
        $failed = getWorkspaceMoveOutcome $null "失敗" $true "move" "D:\新" "C:\旧"
        $failed.Close | Should -Be $false
        $failed.Warning | Should -Not -BeNullOrEmpty
    }
}

Describe "検索結果の件数の上限の入力" -Tag Unit {
    It "制限あり・欄 <Text> は <Action> <Value>" -TestCases @(
        @{ Text = "20000"; Action = "Apply"; Value = 20000 }
        @{ Text = "10,000"; Action = "Apply"; Value = 10000 }
        @{ Text = "１００００"; Action = "Apply"; Value = 10000 }
        @{ Text = " 20000 "; Action = "Apply"; Value = 20000 }
        @{ Text = "1000"; Action = "Apply"; Value = 1000 }
        @{ Text = "100000"; Action = "Apply"; Value = 100000 }
        @{ Text = "5000"; Action = "None"; Value = 5000 }
        @{ Text = "5,000"; Action = "None"; Value = 5000 }
        @{ Text = "   "; Action = "None"; Value = 5000 }
        @{ Text = "1,0000"; Action = "Invalid"; Value = 5000 }
        @{ Text = "999"; Action = "Invalid"; Value = 5000 }
        @{ Text = "９９９"; Action = "Invalid"; Value = 5000 }
        @{ Text = "100001"; Action = "Invalid"; Value = 5000 }
        @{ Text = "abc"; Action = "Invalid"; Value = 5000 }
        @{ Text = "1.5"; Action = "Invalid"; Value = 5000 }
        @{ Text = "-1"; Action = "Invalid"; Value = 5000 }
        @{ Text = "99999999999999999999"; Action = "Invalid"; Value = 5000 }
    ) {
        param ($Text, $Action, $Value)
        $decision = getSearchLimitInputDecision "Limited" $Text 5000 1000 100000 10000
        $decision.Action | Should -Be $Action
        $decision.Value | Should -Be $Value
        if ($Action -eq "Invalid") {
            $decision.Message | Should -Be "1,000～100,000 の数で入力してください。"
        } else {
            $decision.Message | Should -Be ""
        }
    }

    It "<Label>" -TestCases @(
        @{ Label = "制限なしを選ぶと、今が制限ありなら 0 を保存する"; Mode = "None"; Text = "5,000"; Current = 10000; Action = "Apply"; Value = 0 }
        @{ Label = "制限なしを選んでも、今が制限なしなら何もしない"; Mode = "None"; Text = "5,000"; Current = 0; Action = "None"; Value = 0 }
        @{ Label = "制限なしから制限ありに戻して欄が空なら、既定の件数にする"; Mode = "Limited"; Text = ""; Current = 0; Action = "Apply"; Value = 10000 }
        @{ Label = "制限なしから制限ありに戻して欄に数があれば、その数にする"; Mode = "Limited"; Text = "20000"; Current = 0; Action = "Apply"; Value = 20000 }
        @{ Label = "制限ありのまま欄が空なら、何もしない"; Mode = "Limited"; Text = ""; Current = 10000; Action = "None"; Value = 10000 }
    ) {
        param ($Label, $Mode, $Text, $Current, $Action, $Value)
        $decision = getSearchLimitInputDecision $Mode $Text $Current 1000 100000 10000
        $decision.Action | Should -Be $Action
        $decision.Value | Should -Be $Value
    }

    It "testSearchLimitKeepText: 制限なしにしたとき、［制限あり］の入力として正しい文字だけ残す (<Text>)" -TestCases @(
        @{ Text = "abc"; Keep = $false }
        @{ Text = "500"; Keep = $false }
        @{ Text = "100001"; Keep = $false }
        @{ Text = "20,000"; Keep = $true }
        @{ Text = "２０００"; Keep = $true }
        @{ Text = ""; Keep = $true }
    ) {
        param ($Text, $Keep)
        testSearchLimitKeepText $Text 0 1000 100000 10000 | Should -Be $Keep
    }

    It "getSearchLimitView: 制限なしは既定の数を欄に残して見せ、制限ありはその数を出す" {
        $none = getSearchLimitView 0 10000
        $none.Mode | Should -Be "None"
        $none.Text | Should -Be "10,000"
        $limited = getSearchLimitView 20000 10000
        $limited.Mode | Should -Be "Limited"
        $limited.Text | Should -Be "20,000"
    }

    It "getSearchLimitTexts: 範囲と既定と、重さの注意を出す" {
        $texts = getSearchLimitTexts 1000 100000 10000
        $texts.Range | Should -Be "1,000～100,000 件の範囲で指定できます。既定は 10,000 件です。"
        $texts.Caution | Should -Match "メモリを多く使います"
        $texts.Error | Should -Be "1,000～100,000 の数で入力してください。"
    }
}
