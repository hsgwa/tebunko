# 画面のスレッドで動くところ（gui.ps1・shared/ui・tebunko/ui）から、ファイル・フォルダに触る呼び出しを見つけ、
# 許可の一覧に無ければ落とす。届かないネットワークのフォルダに当たると画面が「応答なし」になるため、
# 新しく足した画面のスレッドの呼び出しに気づけるようにする（docs/design/structure/threads.md「プロセス」「スレッドの一覧」）。
#
# 見るのは、組み込みのコマンド・静的メソッドの呼び出しと、状態層の関数・クラスのメソッドの呼び出し。
# 状態層のどの関数が「ファイル・フォルダに触る」かは、手で並べず、状態層のソースから求める
# （中で触る呼び出しをするもの、およびそれを呼ぶもの。getStateIoNames）。
# startJob の 1 つ目の引数（裏で動く仕事のスクリプトブロック）だけは除く。3 つ目の onDone は画面のスレッドで動くため見る。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"

    # 見つける呼び出し（組み込みのコマンド名。大文字・小文字は区別しない）
    ${uiIoCommandNames} = @(
        "Test-Path", "Resolve-Path", "Get-ChildItem", "Get-Item", "Get-Content",
        "Invoke-Item", "Start-Process",
        "New-Item", "Remove-Item", "Copy-Item", "Move-Item", "Rename-Item", "Set-Content", "Add-Content", "Out-File"
    )
    # 見つける静的メソッドの型（[System.IO.File]::GetAttributes のように、どのメソッドでも見つける）
    ${uiIoStaticTypes} = @("System.IO.File", "System.IO.Directory", "System.IO.DirectoryInfo", "System.IO.FileInfo")
    # 上のほかに、System.IO. で始まる型（[System.IO.FileStream]::new・[IO.File]・[System.IO.Compression.ZipFile]・[System.IO.DriveInfo] など）の
    # メソッド（コンストラクターの ::new を含む）の呼び出しと、New-Object の第一引数の型も見つける。
    # System. を省いた短い名前（[IO.File]）も同じに扱う。ファイルに触らない型だけは除く
    ${uiIoNonFileTypes} = @("System.IO.Path", "System.IO.MemoryStream", "System.IO.StringReader", "System.IO.StringWriter", "System.IO.TextReader", "System.IO.TextWriter")

    # 状態層（画面に触らない）のフォルダ。ここにある関数・クラスのメソッドのうち、ファイル・フォルダに触るものと、
    # それを呼ぶものの名前を自動で求め（getStateIoNames）、画面のスレッドから呼んでいれば、組み込みのコマンドと同じように見つける。
    # 名前を手で並べないので、状態層に新しくファイルに触る関数を足したり、既にある関数がファイルに触るようになったりしても、
    # それを画面のスレッドから呼んでいる所が許可の一覧に無ければ落ちる
    ${stateIoDirs} = @("tebunko\core", "tebunko\index", "tebunko\indexer", "tebunko\search", "shared\core", "shared\office")

    # 自動で求める名前から除くもの: Name（関数・メソッドの名前）・Reason（理由）。
    # 除いた名前は、それを呼ぶだけの関数も入らない。理由は「どの場所を読むか」まで書く
    ${stateIoExcluded} = @(
        @{ Name = "readSettings"; Reason = "setting.config を読む（ツールのフォルダ。書き込めないときは既定のワークスペース。どちらもローカル）" }
        @{ Name = "writeSettings"; Reason = "setting.config に書く（ツールのフォルダ。書き込めないときは既定のワークスペース。どちらもローカル）" }
        @{ Name = "updateSettings"; Reason = "setting.config の 1 項目を書き換える（ツールのフォルダ。書き込めないときは既定のワークスペース。どちらもローカル）" }
        @{ Name = "repairBrokenSettings"; Reason = "壊れた setting.config を退避する（ツールのフォルダ。ローカル）" }
        @{ Name = "getSettingsFilePath"; Reason = "setting.config の場所を決める（ツールのフォルダ。書き込めないときは既定のワークスペース。どちらもローカル）" }
        @{ Name = "readVersionFile"; Reason = "配布物の VERSION.txt を読む（ツールのフォルダの中）" }
        @{ Name = "getPartLoad"; Reason = "部品（lib.ps1 など）の場所を確かめる（ツールのフォルダの中）" }
        @{ Name = "testDefaultWorkspace"; Reason = "既定のワークスペース（ドキュメントの tebunko_ws）の中身を数える。ネットワークの場所（プロファイルが共有にある・環境変数が UNC）は testNetworkPath で先に返し、触らない" }
    )

    # 状態層から自動で求めた名前（BeforeAll の最後に求める）
    ${stateIoNames} = $null

    # 許可の一覧: File（ファイル名）・Function（囲む関数・クラスのメソッド名。トップレベルは ""）・
    # Call（見つかった呼び出しの表記）・Count（同じ関数の同じ呼び出しの数。書かなければ 1）・Reason（理由）。行番号ではなく名前で引く
    ${uiIoAllowed} = @(
        # ---- gui.ps1（トップレベル） ----
        @{ File = "gui.ps1"; Function = ""; Call = "Get-ChildItem"; Reason = "Mark-of-the-Web を消す（ツールのフォルダの中。Unblock-File）" }

        # ---- tebunko/ui/splash.ps1・startup_error.ps1（起動口から読み込む部品） ----
        @{ File = "splash.ps1"; Function = "getSplashXamlText"; Call = "[System.IO.File]"; Reason = "起動中の表示（起動時に読む splash.xaml と app_icon.xaml。ツールのフォルダの中）" }
        @{ File = "startup_error.ps1"; Function = "writeStartupErrorFile"; Call = "Add-Content"; Reason = "起動そのものに失敗したときに、ツールのフォルダの startup_error.txt へ書く（窓が無い・応答なしにならない起動の失敗時だけ）" }
        @{ File = "startup_error.ps1"; Function = "getExistingRecordFile"; Call = "Test-Path"; Reason = "起動そのものに失敗したときの reportStartupFailure が、記録が実際に書けたかを確かめる（窓が無い・応答なしにならない起動の失敗時だけ）" }

        # ---- tebunko/ui/workspace_jobs.ps1 ----
        @{ File = "workspace_jobs.ps1"; Function = "testStartupIndexExists"; Call = "testIndexExists"; Reason = "ローカルの場所だけ（ネットワークの場所は先に真を返し、呼ばない）" }
        @{ File = "workspace_jobs.ps1"; Function = "refreshLegacyIndexMessage"; Call = "getLegacyIndexState"; Reason = "ワークスペースの場所がローカルのときに限って呼ぶところ（testNetworkPath で確かめ済み。ネットワークなら裏の仕事で調べる）" }

        # ---- tebunko/ui/indexing_tab.ps1 ----
        @{ File = "indexing_tab.ps1"; Function = "startIndexing"; Call = "getLegacyIndexState"; Reason = "ワークスペースの場所がローカルのときに限って呼ぶところ（testNetworkPath で確かめ済み。ネットワークなら裏の仕事で調べる）" }
        @{ File = "gui_main.ps1"; Function = "writeCloseTrace"; Call = "[System.IO.File]"; Count = 2; Reason = "TEBUNKO_CLOSE_TRACE=1 のときだけ、閉じる途中にワークスペースの決まったファイル（close_trace.txt）へ 1 行足す。テストの診断用（ワークスペースが届かない共有にあると閉じるのが遅れ得るが、テストのときだけ）" }

        # ---- tebunko/ui/open_source.ps1 ----
        @{ File = "open_source.ps1"; Function = "continueFindSourceFile"; Call = "findSourceFileState"; Reason = "ローカルのパスに限って呼ぶところ（testNetworkPath で確かめ済み。ネットワークなら裏の仕事で呼ぶ）" }
        @{ File = "open_source.ps1"; Function = "findSourceLocationAsync"; Call = "getSourceLocation"; Reason = "ワークスペースの場所がローカルのときに限って呼ぶところ（testNetworkPath で確かめ済み。ネットワークなら裏の仕事で対応を読む）" }
        @{ File = "open_source.ps1"; Function = "promptSourceMissing"; Call = "findMovedSource"; Reason = "選んだ直後のフォルダの中から、移された元のファイルを探す（フォルダ選択で選んだ場所）" }
        @{ File = "open_source.ps1"; Function = "openWithShell"; Call = "[System.Diagnostics.Process]"; Reason = "既定のアプリで開く（元のファイルは確かめ済み。プロセスの起動は待たない）" }
        @{ File = "open_source.ps1"; Function = "openWithShell"; Call = "Invoke-Item"; Reason = "既定のアプリで開く（元のファイルは確かめ済み。プロセスの起動は待たない）" }
        @{ File = "open_source.ps1"; Function = "openSourceFolder"; Call = "Start-Process"; Reason = "エクスプローラーで選ぶ（元のファイルは確かめ済み。プロセスの起動は待たない）" }
        @{ File = "open_source.ps1"; Function = "openWithNotepad"; Call = "Start-Process"; Reason = "固定のパスのメモ帳で開く（実行・登録になる拡張子。元のファイルは確かめ済み。プロセスの起動は待たない）" }
        @{ File = "open_source.ps1"; Function = "exportResults"; Call = "[System.IO.Directory]"; Reason = "search_results.txt の出力先フォルダを作る（検索結果の出力の改善で startJob に移す。出力は利用者が押したときだけ）" }
        @{ File = "open_source.ps1"; Function = "exportResults"; Call = "New-Object"; Reason = "search_results.txt に書く（検索結果の出力の改善で startJob に移す。出力は利用者が押したときだけ）" }
        @{ File = "open_source.ps1"; Function = "exportResults"; Call = "Invoke-Item"; Reason = "search_results.txt を開く（プロセスの起動は待たない）" }

        # ---- tebunko/ui/about_dialog.ps1・shared/ui/app_host.ps1（アイコン・XAML。ツールのフォルダの中） ----
        @{ File = "app_host.ps1"; Function = "getXamlText"; Call = "[System.IO.File]"; Reason = "画面定義（XAML）の読み込み（ツールのフォルダの中）" }
        @{ File = "app_host.ps1"; Function = "newAppFontFamily"; Call = "Test-Path"; Reason = "同梱のフォント（ツールのフォルダの中）" }
        @{ File = "app_host.ps1"; Function = "writeErrorLog"; Call = "Test-Path"; Reason = "画面のエラーの記録（まれ。プロセスが落ちる前に確実に残すため、画面のスレッドで書く）" }
        @{ File = "app_host.ps1"; Function = "writeErrorLog"; Call = "New-Item"; Reason = "画面のエラーの記録（まれ。プロセスが落ちる前に確実に残すため、画面のスレッドで書く）" }
        @{ File = "app_host.ps1"; Function = "writeErrorLog"; Call = "[System.IO.File]"; Reason = "画面のエラーの記録（まれ。プロセスが落ちる前に確実に残すため、画面のスレッドで書く）" }

        # ---- tebunko/ui/settings/settings.ps1（［変更…］・［既定に戻す］） ----
        @{ File = "settings.ps1"; Function = "resetWorkspace"; Call = "[System.IO.Directory]"; Reason = "既定のワークスペース（ドキュメントの tebunko_ws。ローカル）を作る" }
        @{ File = "settings.ps1"; Function = "getFolderEntrySample"; Call = "[System.IO.Directory]"; Reason = "選んだ直後のフォルダ（［変更…］・［既定に戻す］で選んだ場所）の中身を数える" }
        @{ File = "settings.ps1"; Function = "applyWorkspace"; Call = "Test-Path"; Count = 2; Reason = "選んだ直後のフォルダ（［変更…］・［既定に戻す］で選んだ場所）" }
        @{ File = "settings.ps1"; Function = "applyWorkspace"; Call = "[System.IO.Directory]"; Reason = "選んだ直後のフォルダ（［変更…］・［既定に戻す］で選んだ場所）の下に作る" }
        @{ File = "settings.ps1"; Function = "applyWorkspace"; Call = "getWorkspaceEntries"; Reason = "選んだ直後のフォルダ（［変更…］・［既定に戻す］で選んだ場所）の中身" }
        @{ File = "settings.ps1"; Function = "applyWorkspace"; Call = "testWritableFolder"; Count = 2; Reason = "選んだ直後のフォルダ（［変更…］・［既定に戻す］で選んだ場所）に書き込めるか" }
        @{ File = "settings.ps1"; Function = "applyWorkspace"; Call = "useWorkspaceTargets"; Reason = "選んだ直後のフォルダ（［変更…］・［既定に戻す］で選んだ場所）にあるインデックスの一覧を読む" }
        @{ File = "settings.ps1"; Function = "applyWorkspace"; Call = "removeWorkspaceEntries"; Reason = "選んだ直後のフォルダ（［変更…］・［既定に戻す］で選んだ場所）のインデックスを消して最初からやり直す（利用者が選んだときだけ）" }
        @{ File = "settings.ps1"; Function = "applyWorkspace"; Call = "moveWorkspace"; Reason = "前のワークスペースの中身を移す（移す量だけかかる。裏に移すのは別の改善。利用者が変えたときだけ）" }

        # ---- tebunko/ui/index/ ----
        @{ File = "index_detail.ps1"; Function = "openFailedFileFolder"; Call = "getPathState"; Count = 2; Reason = "ローカルのパスに限って呼ぶところ（testNetworkPath で確かめ済み。ネットワークなら裏の仕事で呼ぶ）" }
        @{ File = "index_detail.ps1"; Function = "applyFailedFileState"; Call = "Start-Process"; Count = 2; Reason = "エクスプローラーで開く（プロセスの起動は待たない）" }
        @{ File = "index_list.ps1"; Function = "openIndexSourceFolder"; Call = "getPathState"; Reason = "ローカルのパスに限って呼ぶところ（testNetworkPath で確かめ済み。ネットワークなら裏の仕事で呼ぶ）" }
        @{ File = "index_list.ps1"; Function = "loadTargets"; Call = "readStatusFile"; Reason = "ワークスペースの場所がローカルのときに限って呼ぶところ（testNetworkPath で確かめ済み。ネットワークなら裏の仕事で名前を決める）" }
        @{ File = "index_tree.ps1"; Function = "loadIndexTree"; Call = "getIndexTreeData"; Reason = "ワークスペースの場所がローカルのときに限って呼ぶところ（testNetworkPath で確かめ済み。ネットワークなら裏の仕事で集める）" }
        @{ File = "index_tree.ps1"; Function = "expandIndexNode"; Call = "getIndexFolderChildren"; Reason = "ワークスペースの場所がローカルのときに限って呼ぶところ（testNetworkPath で確かめ済み。ネットワークなら裏の仕事で読む）" }
        @{ File = "index_list.ps1"; Function = "applyIndexSourceFolderState"; Call = "Start-Process"; Reason = "エクスプローラーで開く（プロセスの起動は待たない）" }

        # ---- shared/ui/shell.ps1・folder_dialog.ps1 ----
        @{ File = "shell.ps1"; Function = "showErrorDialog"; Call = "Test-Path"; Reason = "ツールのフォルダにあるエラーの記録（ローカル）" }
        @{ File = "shell.ps1"; Function = "showErrorDialog"; Call = "Start-Process"; Reason = "メモ帳で開く（プロセスの起動は待たない）" }
        @{ File = "folder_dialog.ps1"; Function = "selectFolder"; Call = "getExistingAncestorFolder"; Reason = "フォルダ選択の開始フォルダ（ネットワークのパスは調べない引数を渡す）" }
        @{ File = "folder_dialog.ps1"; Function = "getDroppedFolders"; Call = "Test-Path"; Reason = "ドロップされた直後のフォルダ" }
    )

    function getUiIoTargetFiles {
        return @(
            (Resolve-Path "${scriptsDir}\tebunko\gui.ps1").Path
        ) + @(Get-ChildItem "${scriptsDir}\tebunko\ui" -Filter "*.ps1" -Recurse | ForEach-Object { $_.FullName }) `
          + @(Get-ChildItem "${scriptsDir}\shared\ui" -Filter "*.ps1" -Recurse | ForEach-Object { $_.FullName })
    }

    function getStateTargetFiles {
        # 状態層のファイル（画面に触らない。画面のスレッドから呼ぶと、ファイル・フォルダに触る関数が入る）
        return @(${stateIoDirs} | ForEach-Object { Get-ChildItem "${scriptsDir}\$_" -Filter "*.ps1" -Recurse | ForEach-Object { $_.FullName } })
    }

    ${parsedScripts} = @{}

    function parseScriptFile {
        # 同じファイルを何度も解析しないよう、解析した結果を取っておく
        param ([string]$path)

        if (-not ${parsedScripts}.ContainsKey($path)) {
            ${parsedScripts}[$path] = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$null, [ref]$null)
        }
        return ${parsedScripts}[$path]
    }

    function getEnclosingName {
        # ast を囲む関数・クラスのメソッドの名前（FunctionDefinitionAst・FunctionMemberAst）。トップレベルは ""
        param ($ast)

        $node = $ast.Parent
        while ($null -ne $node) {
            if ($node -is [System.Management.Automation.Language.FunctionDefinitionAst]) {
                return $node.Name
            }
            if ($node -is [System.Management.Automation.Language.FunctionMemberAst]) {
                return $node.Name
            }
            $node = $node.Parent
        }
        return ""
    }

    function getStartJobExcludedRanges {
        # startJob（と、それを包む startIndexArchiveJob）の、裏で動く仕事のスクリプトブロックの範囲を返す。この中は見ない。
        #   startJob            : 1 つ目の引数（CommandElements[1]）
        #   startIndexArchiveJob: 2 つ目の引数（CommandElements[2]。1 つ目は表示用の operation の文字列）
        param ($fileAst)

        $ranges = New-Object System.Collections.Generic.List[object]
        $commands = $fileAst.FindAll({ param ($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)
        foreach ($command in $commands) {
            $name = $command.GetCommandName()
            $index = switch ($name) {
                "startJob" { 1 }
                "startIndexArchiveJob" { 2 }
                default { -1 }
            }
            if ($index -lt 0) {
                continue
            }
            if ($command.CommandElements.Count -gt $index) {
                $target = $command.CommandElements[$index]
                $ranges.Add(@{ Start = $target.Extent.StartOffset; End = $target.Extent.EndOffset })
            }
        }
        return $ranges
    }

    function findLiteralNetworkQueues {
        # startJob（と startIndexArchiveJob）の引数に、列の名前 "network" を文字列で直接書いている呼び出しを返す。
        # 列は getWorkspaceJobQueue で場所から決める（場所を渡さずに "network" と書くと、ローカルの仕事まで共有フォルダの列に並ぶ）
        param ($fileAst)

        $found = New-Object System.Collections.Generic.List[object]
        $commands = $fileAst.FindAll({ param ($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)
        foreach ($command in $commands) {
            if (@("startJob", "startIndexArchiveJob") -notcontains $command.GetCommandName()) {
                continue
            }
            foreach ($element in $command.CommandElements | Select-Object -Skip 1) {
                $target = $element
                if ($element -is [System.Management.Automation.Language.CommandParameterAst]) {
                    $target = $element.Argument
                }
                if ($target -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $target.Value -ieq "network") {
                    $found.Add($command)
                }
            }
        }
        return $found.ToArray()
    }

    function testInExcludedRange {
        param ($ast, $ranges)

        foreach ($range in $ranges) {
            if ($ast.Extent.StartOffset -ge $range.Start -and $ast.Extent.EndOffset -le $range.End) {
                return $true
            }
        }
        return $false
    }

    function testIoTypeName {
        # ファイル・フォルダに触る System.IO の型の名前か（System. を省いた短い名前も同じに扱う）
        param ([string]$typeName)

        $full = getFullIoTypeName $typeName
        if ($full -notlike "System.IO.*") {
            return $false
        }
        return -not (${uiIoNonFileTypes} -contains $full)
    }

    function getFullIoTypeName {
        param ([string]$typeName)

        if ($typeName -like "IO.*") {
            return "System.$typeName"
        }
        return $typeName
    }

    function testPlaceholderReason {
        # 置き場所の文字列（%s・{0}）や、短すぎる理由か
        param ([string]$reason)

        $text = $reason.Trim()
        return ($text.Length -lt 12) -or ($text -match '%[sd]|\{\d*\}')
    }

    function newNameSet {
        # 大文字・小文字を区別しない名前の集合。関数の戻り値にしても配列に開かれないよう、, でくるんで返す
        param ([string[]]$names = @())

        $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($name in $names) {
            [void]$set.Add($name)
        }
        return , $set
    }

    function findIoNodes {
        # ast の中の、ファイル・フォルダに触る呼び出しを @{ Node; Call } の配列で返す。
        #   excluded     : 見ない範囲（startJob の仕事のスクリプトブロック）
        #   functionNames: 状態層から求めた、ファイル・フォルダに触る関数の名前（コマンドとして呼ぶもの。Call はその名前）
        #   methodNames  : 同じく、クラスのメソッドの名前（$x.Name(...) の形で呼ぶもの。Call は ".Name()"）
        param ($ast, $excluded = @(), $functionNames = $null, $methodNames = $null)

        $found = New-Object System.Collections.Generic.List[object]

        # コマンド呼び出し（Test-Path・Get-ChildItem や、状態層の関数）
        $commands = $ast.FindAll({ param ($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)
        foreach ($command in $commands) {
            if (testInExcludedRange $command $excluded) {
                continue
            }
            $name = $command.GetCommandName()
            if ($null -eq $name) {
                continue
            }
            if ($name -eq "New-Object") {
                # New-Object <型> の形（第一引数が型名の文字列）だけを見る。-TypeName でも同じ位置に来る
                $typeArg = $null
                if ($command.CommandElements.Count -gt 1) {
                    $second = $command.CommandElements[1]
                    if ($second -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                        $typeArg = $second.Value
                    }
                }
                if ($null -ne $typeArg -and (testIoTypeName $typeArg)) {
                    $found.Add(@{ Node = $command; Call = "New-Object" })
                }
                continue
            }
            $builtin = @(${uiIoCommandNames} | Where-Object { $name -ieq $_ })
            if ($builtin.Count -gt 0) {
                $found.Add(@{ Node = $command; Call = $builtin[0] })
                continue
            }
            if ($null -ne $functionNames -and $functionNames.Contains($name)) {
                $found.Add(@{ Node = $command; Call = $name })
            }
        }

        # メソッドの呼び出し（[System.IO.File]::GetAttributes のような静的メソッドと、$x.Name(...) のクラスのメソッド）
        $members = $ast.FindAll({ param ($n) $n -is [System.Management.Automation.Language.MemberExpressionAst] }, $true)
        foreach ($member in $members) {
            if (testInExcludedRange $member $excluded) {
                continue
            }
            $methodName = $null
            if ($member.Member -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                $methodName = $member.Member.Value
            }
            if ($member.Static -and $member.Expression -is [System.Management.Automation.Language.TypeExpressionAst]) {
                $typeName = $member.Expression.TypeName.FullName
                if (${uiIoStaticTypes} -contains (getFullIoTypeName $typeName)) {
                    $found.Add(@{ Node = $member; Call = "[$(getFullIoTypeName $typeName)]" })
                    continue
                }
                if ($member -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and (testIoTypeName $typeName)) {
                    $found.Add(@{ Node = $member; Call = "[$(getFullIoTypeName $typeName)]" })
                    continue
                }
                if ($typeName -eq "System.Diagnostics.Process" -and $methodName -eq "Start") {
                    $found.Add(@{ Node = $member; Call = "[$typeName]" })
                    continue
                }
            }
            if ($member -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $null -ne $methodName -and $null -ne $methodNames -and $methodNames.Contains($methodName)) {
                $found.Add(@{ Node = $member; Call = ".${methodName}()" })
            }
        }

        # ${function:名前} で関数を取り出す所（変数に入れて & で呼ぶ形。取り出した時点で、呼ぶものとして見つける）
        $variables = $ast.FindAll({ param ($n) $n -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)
        foreach ($variable in $variables) {
            if (testInExcludedRange $variable $excluded) {
                continue
            }
            $path = $variable.VariablePath.UserPath
            if ($path -notlike "function:*") {
                continue
            }
            $functionName = $path.Substring("function:".Length)
            if ($null -ne $functionNames -and $functionNames.Contains($functionName)) {
                $found.Add(@{ Node = $variable; Call = $functionName })
            }
        }

        return $found.ToArray()
    }

    function getStateDefinitions {
        # 状態層の、関数・クラスのメソッド（コンストラクターを除く）の定義を
        # @{ Name; IsMethod; File; Direct; Commands; Members } で返す。
        #   Direct  : 中でファイル・フォルダに触る呼び出しをしているか
        #   Commands: 中で呼んでいるコマンドの名前   Members: 中で呼んでいるメソッドの名前
        $definitions = New-Object System.Collections.Generic.List[object]
        foreach ($path in (getStateTargetFiles)) {
            $ast = parseScriptFile $path
            $nodes = $ast.FindAll({
                    param ($n)
                    ($n -is [System.Management.Automation.Language.FunctionDefinitionAst]) -or
                    (($n -is [System.Management.Automation.Language.FunctionMemberAst]) -and -not $n.IsConstructor)
                }, $true)
            foreach ($node in $nodes) {
                $commandNames = New-Object System.Collections.Generic.List[string]
                foreach ($command in $node.FindAll({ param ($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)) {
                    $commandName = $command.GetCommandName()
                    if ($commandName) {
                        $commandNames.Add($commandName)
                    }
                }
                $memberNames = New-Object System.Collections.Generic.List[string]
                foreach ($member in $node.FindAll({ param ($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] }, $true)) {
                    if ($member.Member -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                        $memberNames.Add($member.Member.Value)
                    }
                }
                $definitions.Add(@{
                        Name     = $node.Name
                        IsMethod = ($node -is [System.Management.Automation.Language.FunctionMemberAst])
                        File     = (Split-Path $path -Leaf)
                        Direct   = (@(findIoNodes $node).Count -gt 0)
                        Commands = $commandNames.ToArray()
                        Members  = $memberNames.ToArray()
                    })
            }
        }
        return $definitions.ToArray()
    }

    function getStateIoNames {
        # 状態層の関数・メソッドのうち、画面のスレッドから呼ぶとファイル・フォルダに触るものの名前を求める。
        #   1. 中でファイル・フォルダに触る呼び出し（findIoNodes の見つけるもの）をするもの
        #   2. それを呼ぶもの（関数の呼び出しとメソッドの呼び出しをたどって、増えなくなるまで）
        # excludedNames の名前は入れない（入れないので、それを呼ぶだけのものも入らない）。
        # 戻り値: @{ Functions; Methods }（それぞれ名前の集合）
        param ($definitions, [string[]]$excludedNames = @())

        $excludedSet = newNameSet $excludedNames
        $functions = newNameSet
        $methods = newNameSet
        $changed = $true
        while ($changed) {
            $changed = $false
            foreach ($definition in $definitions) {
                # 空の集合は、if の式の値にすると配列に開かれて $null になるため、代入の文で選ぶ
                $set = $functions
                if ($definition.IsMethod) {
                    $set = $methods
                }
                if ($set.Contains($definition.Name) -or $excludedSet.Contains($definition.Name)) {
                    continue
                }
                $touches = $definition.Direct -or
                    (@($definition.Commands | Where-Object { $functions.Contains($_) }).Count -gt 0) -or
                    (@($definition.Members | Where-Object { $methods.Contains($_) }).Count -gt 0)
                if ($touches) {
                    [void]$set.Add($definition.Name)
                    $changed = $true
                }
            }
        }
        return @{ Functions = $functions; Methods = $methods }
    }

    function findUiIoCalls {
        # ファイルの中の、画面のスレッドで動く呼び出しを @{ File; Function; Call } の配列で返す
        param ([string]$path)

        $ast = parseScriptFile $path
        $excluded = getStartJobExcludedRanges $ast
        $fileName = Split-Path $path -Leaf
        return @(findIoNodes $ast $excluded $stateIoNames.Functions $stateIoNames.Methods | ForEach-Object {
                @{ File = $fileName; Function = (getEnclosingName $_.Node); Call = $_.Call }
            })
    }

    ${stateIoDefinitions} = getStateDefinitions
    ${stateIoNames} = getStateIoNames ${stateIoDefinitions} @(${stateIoExcluded} | ForEach-Object { $_.Name })
}

Describe "画面のスレッドのファイル・フォルダ操作" -Tag Meta {
    BeforeDiscovery {
        $files = @(
            (Resolve-Path "$PSScriptRoot\..\..\scripts\tebunko\gui.ps1").Path
        ) + @(Get-ChildItem "$PSScriptRoot\..\..\scripts\tebunko\ui" -Filter "*.ps1" -Recurse | ForEach-Object { $_.FullName }) `
          + @(Get-ChildItem "$PSScriptRoot\..\..\scripts\shared\ui" -Filter "*.ps1" -Recurse | ForEach-Object { $_.FullName })
        $cases = @($files | ForEach-Object { @{ Name = (Split-Path $_ -Leaf); Path = $_ } })
    }

    It "<Name> の呼び出しが、許可の一覧のものだけ" -TestCases $cases {
        param ($Name, $Path)

        $calls = findUiIoCalls $Path
        $unexpected = @($calls | Where-Object {
            $call = $_
            -not (${uiIoAllowed} | Where-Object {
                $_.File -eq $call.File -and $_.Function -eq $call.Function -and $_.Call -eq $call.Call
            })
        })
        $lines = @($unexpected | ForEach-Object { "$($_.File):$($_.Function):$($_.Call)" })
        ($lines -join ", ") | Should -Be ""
    }

    It "同じ関数の同じ呼び出しの数が、許可した数と同じ（Count が無い項目は 1 回）" {
        # 許可の単位は (File, Function, Call)。同じ関数に 2 つ目の呼び出しを足しても通らないよう、数も見る。
        # 減ったのに一覧の数を直し忘れると、次に増やしたときに気づけないため、少ない場合も落とす
        $files = getUiIoTargetFiles
        $allCalls = @($files | ForEach-Object { findUiIoCalls $_ })
        $over = @($allCalls | Group-Object { "$($_.File)|$($_.Function)|$($_.Call)" } | ForEach-Object {
            $group = $_
            $first = $group.Group[0]
            $allow = @(${uiIoAllowed} | Where-Object { $_.File -eq $first.File -and $_.Function -eq $first.Function -and $_.Call -eq $first.Call })
            $limit = 1
            if ($allow.Count -gt 0 -and $allow[0].ContainsKey("Count")) {
                $limit = $allow[0].Count
            }
            if ($allow.Count -gt 0 -and $group.Count -ne $limit) {
                "$($first.File):$($first.Function):$($first.Call)=$($group.Count)"
            }
        })
        ($over -join ", ") | Should -Be ""
    }

    It "<Name> が、startJob の列を文字列の `"network`" で直接書いていない（getWorkspaceJobQueue で決める）" -TestCases $cases {
        param ($Name, $Path)

        $literal = @(findLiteralNetworkQueues (parseScriptFile $Path) | ForEach-Object { $_.Extent.StartLineNumber })
        ($literal -join ", ") | Should -Be ""
    }

    It "許可の一覧に、今のコードに無い項目が残っていない" {
        # 直したのに一覧を消し忘れると、次に同じ場所を直したときに気づけない
        $files = getUiIoTargetFiles
        $allCalls = @($files | ForEach-Object { findUiIoCalls $_ })
        $stale = @(${uiIoAllowed} | Where-Object {
            $allow = $_
            -not (@($allCalls) | Where-Object {
                $_.File -eq $allow.File -and $_.Function -eq $allow.Function -and $_.Call -eq $allow.Call
            })
        } | ForEach-Object { "$($_.File):$($_.Function):$($_.Call)" })
        ($stale -join ", ") | Should -Be ""
    }

    It "許可の一覧の理由に、先送りの言葉が残っていない（直すか、直さない理由を具体的に書く）" {
        $vague = @(${uiIoAllowed} | Where-Object { $_.Reason -match "分けた PR|ワークスペースの側" } | ForEach-Object { "$($_.File):$($_.Function):$($_.Call)" })
        ($vague -join ", ") | Should -Be ""
        $noReason = @(${uiIoAllowed} | Where-Object { [string]::IsNullOrWhiteSpace($_.Reason) } | ForEach-Object { "$($_.File):$($_.Function):$($_.Call)" })
        ($noReason -join ", ") | Should -Be ""
        $excludedVague = @(${stateIoExcluded} | Where-Object { [string]::IsNullOrWhiteSpace($_.Reason) -or $_.Reason -match "分けた PR|ワークスペースの側" } | ForEach-Object { $_.Name })
        ($excludedVague -join ", ") | Should -Be ""
        # 「%s」のような置き場所の文字列や、短すぎる理由では、何が安全なのか読み取れない
        $placeholder = @(${uiIoAllowed} | Where-Object { (testPlaceholderReason $_.Reason) } | ForEach-Object { "$($_.File):$($_.Function):$($_.Call)" })
        ($placeholder -join ", ") | Should -Be ""
        $excludedPlaceholder = @(${stateIoExcluded} | Where-Object { (testPlaceholderReason $_.Reason) } | ForEach-Object { $_.Name })
        ($excludedPlaceholder -join ", ") | Should -Be ""
    }

    It "許可の一覧の項目を囲む関数が、どこかで呼ばれている（使われない関数の項目は消す）" {
        $called = newNameSet
        foreach ($file in (Get-ChildItem "${scriptsDir}" -Filter "*.ps1" -Recurse)) {
            $ast = parseScriptFile $file.FullName
            foreach ($node in $ast.FindAll({ param ($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)) {
                $name = $node.GetCommandName()
                if ($name) {
                    [void]$called.Add($name)
                }
            }
            foreach ($node in $ast.FindAll({ param ($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] }, $true)) {
                if ($node.Member -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                    [void]$called.Add($node.Member.Value)
                }
            }
            # ${function:名前} で取り出して呼ぶもの
            foreach ($node in $ast.FindAll({ param ($n) $n -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)) {
                $variable = $node.VariablePath.UserPath
                if ($variable -like "function:*") {
                    [void]$called.Add($variable.Substring("function:".Length))
                }
            }
        }
        $unused = @(${uiIoAllowed} | Where-Object { $_.Function -ne "" -and -not $called.Contains($_.Function) } | ForEach-Object { "$($_.File):$($_.Function):$($_.Call)" })
        ($unused -join ", ") | Should -Be ""
    }

    It "状態層から自動で求めた名前に <Name> が入っている" -TestCases @(
        @{ Name = "getSearchIndexes"; Kind = "Functions" }
        @{ Name = "getSourceFolderMap"; Kind = "Functions" }
        @{ Name = "getSourceLocation"; Kind = "Functions" }
        @{ Name = "renameIndex"; Kind = "Functions" }
        @{ Name = "removeIndex"; Kind = "Functions" }
        @{ Name = "readStatusFile"; Kind = "Functions" }
        @{ Name = "getIndexingState"; Kind = "Functions" }
        @{ Name = "getLegacyIndexState"; Kind = "Functions" }
        @{ Name = "getPathState"; Kind = "Functions" }
        @{ Name = "testIndexExists"; Kind = "Functions" }
        @{ Name = "findSourceFileState"; Kind = "Functions" }
        @{ Name = "getExistingAncestorFolder"; Kind = "Functions" }
        @{ Name = "getOfficeProcesses"; Kind = "Functions" }
        @{ Name = "ReadStatus"; Kind = "Methods" }
    ) {
        param ($Name, $Kind)

        $stateIoNames[$Kind].Contains($Name) | Should -Be $true
    }

    It "除く一覧の名前は、除かなければ入る名前で、実際には入っていない" {
        $without = getStateIoNames ${stateIoDefinitions} @()
        $stale = @(${stateIoExcluded} | Where-Object { -not ($without.Functions.Contains($_.Name) -or $without.Methods.Contains($_.Name)) } | ForEach-Object { $_.Name })
        ($stale -join ", ") | Should -Be ""
        $leaked = @(${stateIoExcluded} | Where-Object { $stateIoNames.Functions.Contains($_.Name) -or $stateIoNames.Methods.Contains($_.Name) } | ForEach-Object { $_.Name })
        ($leaked -join ", ") | Should -Be ""
    }

    Context "見つけ方（小さな例のスクリプト）" {
        It "コマンドの呼び出し・メソッドの呼び出し（InvokeMemberExpressionAst）・静的メソッドを見つけ、startJob の仕事の中は見ない" {
            $sample = @'
function sample {
    renameIndex "a" "b"
    $catalog.Rename("a")
    $list.Add(1)
    [System.IO.File]::Exists($p)
    New-Object System.IO.StreamWriter($p)
    New-Object System.Text.StringBuilder
    startJob { Test-Path x } @() { Test-Path y }
}
'@
            $ast = [System.Management.Automation.Language.Parser]::ParseInput($sample, [ref]$null, [ref]$null)
            $functions = newNameSet @("renameIndex")
            $methods = newNameSet @("Rename")

            $found = @(findIoNodes $ast (getStartJobExcludedRanges $ast) $functions $methods | ForEach-Object { $_.Call })

            $found | Should -Be @("renameIndex", "New-Object", "Test-Path", ".Rename()", "[System.IO.File]")
        }

        It "System.IO の型（コンストラクター・短い名前・圧縮・ドライブ）と、取り出した関数も見つけ、ファイルに触らない型は見つけない" {
            $sample = @'
function sample {
    [System.IO.FileStream]::new($p, 'Open')
    [System.IO.StreamReader]::new($p)
    [System.IO.Compression.ZipFile]::OpenRead($p)
    [System.IO.DriveInfo]::new("C")
    [IO.File]::Exists($p)
    New-Object System.IO.Compression.ZipArchive($s)
    $reader = ${function:readSomething}
    & $reader
    [System.IO.Path]::Combine($a, $b)
    New-Object System.IO.MemoryStream
    [System.IO.FileMode]::Open
}
'@
            $ast = [System.Management.Automation.Language.Parser]::ParseInput($sample, [ref]$null, [ref]$null)
            $functions = newNameSet @("readSomething")

            $found = @(findIoNodes $ast @() $functions (newNameSet) | ForEach-Object { $_.Call })

            $found | Should -Be @("New-Object", "[System.IO.FileStream]", "[System.IO.StreamReader]", "[System.IO.Compression.ZipFile]", "[System.IO.DriveInfo]", "[System.IO.File]", "readSomething")
        }

        It "startJob の引数に文字列の network を直接書いた呼び出しを見つけ、getWorkspaceJobQueue で決めた呼び出しは見つけない" {
            $sample = @'
function sample {
    startJob { 1 } @() { } "network"
    startJob { 1 } @() { } -queue 'Network'
    startJob { 1 } @() { } (getWorkspaceJobQueue $dir)
    startIndexArchiveJob "x" { 1 } "network"
    other "network"
}
'@
            $ast = [System.Management.Automation.Language.Parser]::ParseInput($sample, [ref]$null, [ref]$null)

            $found = @(findLiteralNetworkQueues $ast | ForEach-Object { $_.Extent.StartLineNumber })

            $found | Should -Be @(2, 3, 5)
        }

        It "理由が置き場所の文字列や短すぎる文字列のとき、置き場所とみなす" -TestCases @(
            @{ Reason = "%s"; Expected = $true }
            @{ Reason = "{0} を読む"; Expected = $true }
            @{ Reason = "ローカル"; Expected = $true }
            @{ Reason = "ローカルのパスに限って呼ぶところ（testNetworkPath で確かめ済み）"; Expected = $false }
        ) {
            param ($Reason, $Expected)

            (testPlaceholderReason $Reason) | Should -Be $Expected
        }

        It "状態層の関数を呼ぶ関数・メソッドも、増えなくなるまでたどって入れる（除いた名前は入れない）" {
            $definitions = @(
                @{ Name = "readSomething"; IsMethod = $false; Direct = $true; Commands = @(); Members = @() }
                @{ Name = "callsReader"; IsMethod = $false; Direct = $false; Commands = @("readSomething"); Members = @() }
                @{ Name = "Wrap"; IsMethod = $true; Direct = $false; Commands = @("callsReader"); Members = @() }
                @{ Name = "callsWrap"; IsMethod = $false; Direct = $false; Commands = @(); Members = @("Wrap") }
                @{ Name = "unrelated"; IsMethod = $false; Direct = $false; Commands = @("Get-Date"); Members = @("Add") }
                @{ Name = "viaExcluded"; IsMethod = $false; Direct = $false; Commands = @("skipped"); Members = @() }
                @{ Name = "skipped"; IsMethod = $false; Direct = $true; Commands = @(); Members = @() }
            )

            $names = getStateIoNames $definitions @("skipped")

            ($names.Functions | Sort-Object) -join "," | Should -Be "callsReader,callsWrap,readSomething"
            ($names.Methods | Sort-Object) -join "," | Should -Be "Wrap"
        }
    }
}
