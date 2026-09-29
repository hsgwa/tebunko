# 画面のスレッドで動くところ（gui.ps1・shared/ui・tebunko/ui）から、ファイル・フォルダに触る呼び出しを見つけ、
# 許可の一覧に無ければ落とす。届かないネットワークのフォルダに当たると画面が「応答なし」になるため、
# 新しく足した画面のスレッドの呼び出しに気づけるようにする（docs/design/structure/threads.md「プロセス」「スレッドの一覧」）。
#
# 見るのはこのファイルの中に書かれた呼び出しだけで、呼んだ先の関数（状態層など）の中までは見ない。
# startJob の 1 つ目の引数（裏で動く仕事のスクリプトブロック）だけは除く。3 つ目の onDone は画面のスレッドで動くため見る。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"

    # 見つける呼び出し（コマンド名。大文字・小文字は区別しない）
    ${uiIoCommandNames} = @(
        "Test-Path", "Resolve-Path", "Get-ChildItem", "Get-Item", "Get-Content",
        "Invoke-Item", "Start-Process",
        "New-Item", "Remove-Item", "Copy-Item", "Move-Item", "Rename-Item", "Set-Content", "Out-File",
        "getSearchIndexes", "testIndexExists", "getSourceFolderMap", "getExistingAncestorFolder", "getDriveTargets",
        "getPathState", "findSourceFileState"
    )
    # 見つける静的メソッドの型（[System.IO.File]::GetAttributes のように、どのメソッドでも見つける）
    ${uiIoStaticTypes} = @("System.IO.File", "System.IO.Directory", "System.IO.DirectoryInfo", "System.IO.FileInfo")
    # New-Object で見つける型（第一引数・-TypeName がこの型のときだけ）
    ${uiIoNewObjectTypes} = @("System.IO.FileStream")

    # 許可の一覧: File（ファイル名）・Function（囲む関数・クラスのメソッド名。トップレベルは ""）・
    # Call（見つかった呼び出しの表記）・Reason（理由）。行番号ではなく名前で引く
    ${uiIoAllowed} = @(
        # ---- gui.ps1（トップレベル） ----
        @{ File = "gui.ps1"; Function = ""; Call = "[System.IO.File]"; Reason = "起動中の表示（起動時に読む splash.xaml。ツールのフォルダの中）" }
        @{ File = "gui.ps1"; Function = ""; Call = "Get-ChildItem"; Reason = "Mark-of-the-Web を消す（ツールのフォルダの中。Unblock-File）" }
        @{ File = "gui.ps1"; Function = ""; Call = "testIndexExists"; Reason = "起動時のタブ選び（ワークスペースの側。分けた PR で直す）" }
        @{ File = "gui.ps1"; Function = "getExistingRecordFile"; Call = "Test-Path"; Reason = "起動そのものに失敗したときの trap が、記録が実際に書けたかを確かめる（窓が無い・応答なしにならない起動の失敗時だけ）" }
        @{ File = "gui.ps1"; Function = "writeStartupErrorFile"; Call = "Test-Path"; Reason = "起動そのものに失敗したときの記録（trap から。窓が無い・応答なしにならない起動の失敗時だけ）" }
        @{ File = "gui.ps1"; Function = "writeStartupErrorFile"; Call = "New-Item"; Reason = "起動そのものに失敗したときの記録（trap から。窓が無い・応答なしにならない起動の失敗時だけ）" }

        # ---- tebunko/ui/indexing_tab.ps1（ワークスペースの側。分けた PR） ----
        @{ File = "indexing_tab.ps1"; Function = "finishIndexing"; Call = "Test-Path"; Reason = "取り込みログの有無（ワークスペースの側。分けた PR）" }
        @{ File = "indexing_tab.ps1"; Function = ""; Call = "Test-Path"; Reason = "［ログを開く］でのログの有無（ワークスペースの側。分けた PR）" }
        @{ File = "indexing_tab.ps1"; Function = ""; Call = "Invoke-Item"; Reason = "［ログを開く］でログを開く（ワークスペースの側。分けた PR）" }

        # ---- tebunko/ui/open_source.ps1 ----
        @{ File = "open_source.ps1"; Function = "findSourceFile"; Call = "findSourceFileState"; Reason = "ローカルのパスに限って呼ぶところ（testNetworkPath で確かめ済み。ネットワークなら裏の仕事で呼ぶ）" }
        @{ File = "open_source.ps1"; Function = "openWithShell"; Call = "[System.Diagnostics.Process]"; Reason = "既定のアプリで開く（元のファイルは確かめ済み。プロセスの起動は待たない）" }
        @{ File = "open_source.ps1"; Function = "openWithShell"; Call = "Invoke-Item"; Reason = "既定のアプリで開く（元のファイルは確かめ済み。プロセスの起動は待たない）" }
        @{ File = "open_source.ps1"; Function = "openSourceFolder"; Call = "Start-Process"; Reason = "エクスプローラーで選ぶ（元のファイルは確かめ済み。プロセスの起動は待たない）" }
        @{ File = "open_source.ps1"; Function = "exportResults"; Call = "[System.IO.Directory]"; Reason = "検索結果.txt の出力先（ワークスペースの側。分けた PR）" }
        @{ File = "open_source.ps1"; Function = "exportResults"; Call = "Invoke-Item"; Reason = "検索結果.txt を開く（プロセスの起動は待たない）" }

        # ---- tebunko/ui/about_dialog.ps1・shared/ui/app_host.ps1（アイコン・XAML。ツールのフォルダの中） ----
        @{ File = "about_dialog.ps1"; Function = "showAboutDialog"; Call = "Test-Path"; Reason = "アイコン（ツールのフォルダの中）" }
        @{ File = "app_host.ps1"; Function = "loadXaml"; Call = "[System.IO.File]"; Reason = "画面定義（XAML）の読み込み（ツールのフォルダの中）" }
        @{ File = "app_host.ps1"; Function = "loadWindow"; Call = "Test-Path"; Reason = "アイコン（ツールのフォルダの中）" }
        @{ File = "app_host.ps1"; Function = "writeErrorLog"; Call = "Test-Path"; Reason = "画面のエラーの記録（ワークスペースの側。分けた PR で扱うかを決める）" }
        @{ File = "app_host.ps1"; Function = "writeErrorLog"; Call = "New-Item"; Reason = "画面のエラーの記録（ワークスペースの側。分けた PR で扱うかを決める）" }
        @{ File = "app_host.ps1"; Function = "writeErrorLog"; Call = "[System.IO.File]"; Reason = "画面のエラーの記録（ワークスペースの側。分けた PR で扱うかを決める）" }

        # ---- tebunko/ui/settings_tab.ps1・index_tree.ps1・types.ps1（ワークスペースの側。分けた PR） ----
        @{ File = "settings_tab.ps1"; Function = "resetWorkspace"; Call = "[System.IO.Directory]"; Reason = "ワークスペースの側（分けた PR）" }
        @{ File = "settings_tab.ps1"; Function = "getFolderEntrySample"; Call = "[System.IO.Directory]"; Reason = "ワークスペースの側（分けた PR）" }
        @{ File = "settings_tab.ps1"; Function = "applyWorkspace"; Call = "Test-Path"; Reason = "ワークスペースの側（分けた PR）" }
        @{ File = "settings_tab.ps1"; Function = "applyWorkspace"; Call = "[System.IO.Directory]"; Reason = "ワークスペースの側（分けた PR）" }
        @{ File = "index_tree.ps1"; Function = "loadIndexTree"; Call = "Test-Path"; Reason = "ワークスペースの側（分けた PR）" }
        @{ File = "index_tree.ps1"; Function = "loadIndexTree"; Call = "Resolve-Path"; Reason = "ワークスペースの側（分けた PR）" }
        @{ File = "index_tree.ps1"; Function = "loadIndexTree"; Call = "getSearchIndexes"; Reason = "ワークスペースの側（分けた PR）" }
        @{ File = "types.ps1"; Function = "CreateRoot"; Call = "[System.IO.Directory]"; Reason = "ワークスペースの側（分けた PR）" }
        @{ File = "types.ps1"; Function = "LoadChildren"; Call = "[System.IO.Directory]"; Reason = "ワークスペースの側（分けた PR）" }
        @{ File = "types.ps1"; Function = "IsBookDirPath"; Call = "[System.IO.Directory]"; Reason = "ワークスペースの側（分けた PR）" }
        @{ File = "types.ps1"; Function = "HasSubfolders"; Call = "[System.IO.Directory]"; Reason = "ワークスペースの側（分けた PR）" }
        @{ File = "types.ps1"; Function = "HasFiles"; Call = "[System.IO.Directory]"; Reason = "ワークスペースの側（分けた PR）" }

        # ---- tebunko/ui/index_tab.ps1 ----
        @{ File = "index_tab.ps1"; Function = "updateIndexSourceFile"; Call = "Test-Path"; Reason = "IndexDir の有無（ワークスペースの側。分けた PR）" }
        @{ File = "index_tab.ps1"; Function = "openFailedFileFolder"; Call = "getPathState"; Reason = "ローカルのパスに限って呼ぶところ（testNetworkPath で確かめ済み。ネットワークなら裏の仕事で呼ぶ）" }
        @{ File = "index_tab.ps1"; Function = "applyFailedFileState"; Call = "Start-Process"; Reason = "エクスプローラーで開く（プロセスの起動は待たない）" }

        # ---- shared/ui/shell.ps1・folder_dialog.ps1 ----
        @{ File = "shell.ps1"; Function = "readTextShared"; Call = "Test-Path"; Reason = "ローカルのパスに限って呼ぶところ（ワークスペース内のファイルを読む）" }
        @{ File = "shell.ps1"; Function = "readTextShared"; Call = "New-Object"; Reason = "ローカルのパスに限って呼ぶところ（ワークスペース内のファイルを読む）" }
        @{ File = "folder_dialog.ps1"; Function = "selectFolder"; Call = "getExistingAncestorFolder"; Reason = "フォルダ選択の開始フォルダ（ネットワークのパスは調べない引数を渡す）" }
        @{ File = "folder_dialog.ps1"; Function = "getDroppedFolders"; Call = "Test-Path"; Reason = "ドロップされた直後のフォルダ" }
    )

    function getUiIoTargetFiles {
        return @(
            (Resolve-Path "${scriptsDir}\tebunko\gui.ps1").Path
        ) + @(Get-ChildItem "${scriptsDir}\tebunko\ui" -Filter "*.ps1" | ForEach-Object { $_.FullName }) `
          + @(Get-ChildItem "${scriptsDir}\shared\ui" -Filter "*.ps1" | ForEach-Object { $_.FullName })
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

    function testInExcludedRange {
        param ($ast, $ranges)

        foreach ($range in $ranges) {
            if ($ast.Extent.StartOffset -ge $range.Start -and $ast.Extent.EndOffset -le $range.End) {
                return $true
            }
        }
        return $false
    }

    function findUiIoCalls {
        # ファイルの中の、画面のスレッドで動く呼び出しを @{ File; Function; Call } の配列で返す
        param ([string]$path)

        $errors = $null
        $tokens = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
        $excluded = getStartJobExcludedRanges $ast
        $fileName = Split-Path $path -Leaf
        $found = New-Object System.Collections.Generic.List[object]

        # コマンド呼び出し（Test-Path・Get-ChildItem・getSearchIndexes など）
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
                if ($null -ne $typeArg -and (${uiIoNewObjectTypes} -contains $typeArg)) {
                    $found.Add(@{ File = $fileName; Function = (getEnclosingName $command); Call = "New-Object" })
                }
                continue
            }
            foreach ($target in ${uiIoCommandNames}) {
                if ($name -ieq $target) {
                    # コマンド名は元の表記（getSearchIndexes など）で引けるよう、決まった名前の一覧の表記をそのまま使う
                    $call = if (@("Test-Path", "Resolve-Path", "Get-ChildItem", "Get-Item", "Get-Content", "Invoke-Item", "Start-Process", "New-Item", "Remove-Item", "Copy-Item", "Move-Item", "Rename-Item", "Set-Content", "Out-File") -contains $target) { $target } else { $target }
                    $found.Add(@{ File = $fileName; Function = (getEnclosingName $command); Call = $call })
                    break
                }
            }
        }

        # 静的メソッドの呼び出し（[System.IO.File]::GetAttributes など）
        $members = $ast.FindAll({ param ($n) $n -is [System.Management.Automation.Language.MemberExpressionAst] }, $true)
        foreach ($member in $members) {
            if (testInExcludedRange $member $excluded) {
                continue
            }
            if (-not $member.Static) {
                continue
            }
            if ($member.Expression -isnot [System.Management.Automation.Language.TypeExpressionAst]) {
                continue
            }
            $typeName = $member.Expression.TypeName.FullName
            if (${uiIoStaticTypes} -contains $typeName) {
                $found.Add(@{ File = $fileName; Function = (getEnclosingName $member); Call = "[$typeName]" })
                continue
            }
            if ($typeName -eq "System.Diagnostics.Process") {
                $methodName = $null
                if ($member.Member -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                    $methodName = $member.Member.Value
                }
                if ($methodName -eq "Start") {
                    $found.Add(@{ File = $fileName; Function = (getEnclosingName $member); Call = "[$typeName]" })
                }
            }
        }

        return $found.ToArray()
    }
}

Describe "画面のスレッドのファイル・フォルダ操作" -Tag Meta {
    BeforeDiscovery {
        $files = @(
            (Resolve-Path "$PSScriptRoot\..\..\scripts\tebunko\gui.ps1").Path
        ) + @(Get-ChildItem "$PSScriptRoot\..\..\scripts\tebunko\ui" -Filter "*.ps1" | ForEach-Object { $_.FullName }) `
          + @(Get-ChildItem "$PSScriptRoot\..\..\scripts\shared\ui" -Filter "*.ps1" | ForEach-Object { $_.FullName })
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
}
