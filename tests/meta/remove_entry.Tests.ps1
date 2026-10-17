# フォルダを消す入口を 1 つに保つ。
# フォルダを中身ごと消すのは removeDirectoryRetry（shared/core/fs.ps1。中でリンクを辿らない removeTreeNoFollow を呼び、消えたことを .NET で確かめる）だけにする。
# 別の消し方（Remove-Item -Recurse・Directory.Delete の再帰）を足すと、リンクの先まで消す・消えていないのに成功と記録する、のどちらかが起きうる。
# 見張るもの:
#   1. scripts の下に、フォルダを再帰で消す書き方（Remove-Item とその別名（モジュール名付きも）の -Recurse（省略形も）・スプラット、
#      Directory.Delete の再帰（メンバー名を引用符で囲んだ書き方も）、型が変数で読めない ::Delete(<path>, <$false でない値>)、
#      DeleteDirectory・DeleteFolder、ForEach-Object・foreach・% や .ForEach(...) にメンバー名 "Delete" を文字列で渡し、$false でない値を付ける書き方）が無い
#   2. removeTreeNoFollow を呼ぶのは shared\core\fs.ps1 の removeDirectoryRetry だけ（スクリプト直下の呼び出しも、同じ名前の別のファイルの関数も違反に数える）
#   3. removeDirectoryRetry を直接呼ぶ関数（スクリプト直下を含む）は、フォルダの有無を Test-Path で見ない（予約名に拡張子の付いた名前のフォルダは、
#      Test-Path と .NET で答えが食い違い、消さずに「削除しました」と記録する不具合になった。有無は [System.IO.Directory]::Exists で見る）
# 「消したあとに残っていれば例外にする」ことは fs.Tests.ps1 の removeDirectoryRetry のテストが守る。
# 1 つのファイルを消す書き方（File.Delete・Remove-Item -Force）は対象にしない（リンクの先を消さないため）。
# 守れない範囲（構文木で見つけられない書き方）:
#   - removeDirectoryRetry を呼ぶ関数を、さらに呼ぶ側の関数の Test-Path（3 は直接呼ぶ関数だけ見る）
#   - コマンド名を変数や文字列の組み立てで作る呼び出し（& $name ... や Invoke-Expression）
#   - Get-ChildItem -Recurse | Remove-Item のように、再帰の列挙をつないで消す書き方
#   - cmd /c rd /s・robocopy /mir など、外のプログラムで消す書き方
#   これらは、この検査では見つけられない。書かない。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"

    ${removeItemNames} = @("Remove-Item", "ri", "rm", "rmdir", "rd", "del", "erase")
    ${forEachNames} = @("ForEach-Object", "foreach", "%")
    ${otherDeleteMembers} = @("DeleteDirectory", "DeleteFolder")

    function parseCode {
        param (
            [string]$code
        )
        $errors = $null
        $tokens = $null
        return [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)
    }

    function testFalseLiteral {
        param (
            $ast
        )
        return ($ast -is [System.Management.Automation.Language.VariableExpressionAst]) -and ($ast.VariablePath.UserPath -eq "false")
    }

    function findRecursiveDeletes {
        # ast（またはその一部）の中から、フォルダを再帰で消すかもしれない書き方を、説明の文字列の配列で返す
        param (
            $ast
        )
        $found = New-Object System.Collections.Generic.List[string]
        $nodes = $ast.FindAll({
                param ($n)
                $n -is [System.Management.Automation.Language.CommandAst] -or $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst]
            }, $true)
        foreach ($node in $nodes) {
            if ($node -is [System.Management.Automation.Language.CommandAst]) {
                $commandName = "$($node.GetCommandName())".Split("\")[-1]
                if (${forEachNames} -contains $commandName) {
                    # ForEach-Object Delete $true・ForEach-Object -MemberName Delete -ArgumentList $true（メンバー名を文字列で渡す書き方）。
                    # Delete がある呼び出しで、Delete 以外に $false でない値（-ArgumentList:$true のようにパラメーターに付いた値も）があれば、前後を問わず違反に数える
                    $elements = @($node.CommandElements | Select-Object -Skip 1)
                    # -MemberName:Delete のように、パラメーターに付いた値が Delete のものも、メンバー名の指定として数える
                    $isDeleteName = {
                        param ($element)
                        ($element -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $element.Value -eq "Delete") -or
                        ($element -is [System.Management.Automation.Language.CommandParameterAst] -and $null -ne $element.Argument -and
                            "MemberName".StartsWith($element.ParameterName, [System.StringComparison]::OrdinalIgnoreCase) -and
                            $element.Argument -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $element.Argument.Value -eq "Delete")
                    }
                    $deleteElements = @($elements | Where-Object { & $isDeleteName $_ })
                    if ($deleteElements.Count -gt 0) {
                        $values = @($elements | Where-Object {
                                if (& $isDeleteName $_) {
                                    $false
                                } elseif ($_ -is [System.Management.Automation.Language.CommandParameterAst]) {
                                    $null -ne $_.Argument -and !(testFalseLiteral $_.Argument)
                                } else {
                                    !(testFalseLiteral $_)
                                }
                            })
                        if ($values.Count -gt 0) {
                            $found.Add("$commandName Delete（引数付き。再帰かもしれない）")
                        }
                    }
                    continue
                }
                if (${removeItemNames} -notcontains $commandName) {
                    continue
                }
                foreach ($element in $node.CommandElements) {
                    if ($element -is [System.Management.Automation.Language.CommandParameterAst] -and
                        "Recurse".StartsWith($element.ParameterName, [System.StringComparison]::OrdinalIgnoreCase)) {
                        $found.Add("$commandName -$($element.ParameterName)")
                    } elseif ($element -is [System.Management.Automation.Language.VariableExpressionAst] -and $element.Splatted) {
                        $found.Add("$commandName @$($element.VariablePath.UserPath)（スプラット。-Recurse が入るか分からない）")
                    }
                }
                continue
            }
            # メンバー名は、文字列の定数なら値で比べる（$info."Delete"(...)・$info.'Delete'(...) も同じ Delete）
            $memberName = if ($node.Member -is [System.Management.Automation.Language.StringConstantExpressionAst]) { $node.Member.Value } else { "" }
            if (${otherDeleteMembers} -contains $memberName) {
                # フォルダを再帰で消す別の入口（VisualBasic の FileSystem.DeleteDirectory・COM の FileSystemObject.DeleteFolder）
                $found.Add("$($node.Expression.Extent.Text).$memberName（別の消し方）")
                continue
            }
            if ($memberName -eq "ForEach" -and $node.Arguments.Count -ge 1 -and
                $node.Arguments[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $node.Arguments[0].Value -eq "Delete") {
                # $dirs.ForEach("Delete", $true)（メンバー名を文字列で渡す書き方）。Delete のあとに $false でない値があれば違反に数える
                $restArguments = @($node.Arguments | Select-Object -Skip 1 | Where-Object { !(testFalseLiteral $_) })
                if ($restArguments.Count -gt 0) {
                    $found.Add("$($node.Expression.Extent.Text).ForEach(`"Delete`", ...)（メンバー名を文字列で渡す。再帰かもしれない）")
                }
                continue
            }
            if ($memberName -ne "Delete") {
                continue
            }
            $argumentCount = if ($null -eq $node.Arguments) { 0 } else { $node.Arguments.Count }
            if ($node.Static) {
                if ($node.Expression -isnot [System.Management.Automation.Language.TypeExpressionAst]) {
                    # 型が変数などで読めない静的呼び出し。再帰の指定がありうる形（引数が 2 つ以上で 2 つ目が $false でない）は違反に倒す
                    if ($argumentCount -ge 2 -and !(testFalseLiteral $node.Arguments[1])) {
                        $found.Add("$($node.Expression.Extent.Text)::Delete（型が読めない。再帰かもしれない）")
                    }
                    continue
                }
                $typeName = $node.Expression.TypeName.Name
                if ($typeName -match "(^|\.)Directory$" -and $argumentCount -ge 2 -and !(testFalseLiteral $node.Arguments[1])) {
                    $found.Add("[$typeName]::Delete（再帰）")
                }
            } elseif ($argumentCount -ge 1 -and !(testFalseLiteral $node.Arguments[0])) {
                $found.Add(".Delete（DirectoryInfo の再帰）")
            }
        }
        return $found.ToArray()
    }

    function findCallsByOwner {
        # ast の中のコマンド呼び出しを全部拾い、呼び出しごとに、いちばん内側の関数（クラスのメソッドを含む）の名前を付けて @{ Owner; Command } で返す。
        # 関数の外（スクリプト直下）の呼び出しは Owner を "(スクリプト直下)" にする。コマンド名は、モジュール名付き（Module\Name）でも後ろの名前で返す
        param (
            $ast
        )
        $commands = $ast.FindAll({ param ($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)
        foreach ($command in $commands) {
            $name = $command.GetCommandName()
            if (!$name) {
                continue
            }
            $owner = "(スクリプト直下)"
            for ($parent = $command.Parent; $null -ne $parent; $parent = $parent.Parent) {
                if ($parent -is [System.Management.Automation.Language.FunctionDefinitionAst] -or $parent -is [System.Management.Automation.Language.FunctionMemberAst]) {
                    $owner = $parent.Name
                    break
                }
            }
            @{ Owner = $owner; Command = $name.Split("\")[-1] }
        }
    }

    function groupCommandsByOwner {
        # findCallsByOwner の結果を、Owner ごとのコマンド名の配列にまとめる（Owner → 配列）
        param (
            $calls
        )
        $groups = @{}
        foreach ($call in $calls) {
            if (!$groups.ContainsKey($call.Owner)) {
                $groups[$call.Owner] = New-Object System.Collections.Generic.List[string]
            }
            $groups[$call.Owner].Add($call.Command)
        }
        return $groups
    }

    function findFolderChangeCalls {
        # ast の中の、フォルダを消す・移す・作る呼び出し（removeDirectoryRetry・moveDirectoryRetry・Directory.Move・Directory.CreateDirectory）を、
        # 呼び出しごとに、いちばん内側の関数（クラスのメソッドを含む）の名前を付けて @{ Owner; Call } で返す。
        # 置き場の下かを確かめる入口は removeDirectoryUnder・moveDirectoryUnder・newDirectoryUnder（shared/core/fs.ps1）
        param (
            $ast
        )
        $nodes = $ast.FindAll({
                param ($n)
                $n -is [System.Management.Automation.Language.CommandAst] -or $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst]
            }, $true)
        foreach ($node in $nodes) {
            $call = $null
            if ($node -is [System.Management.Automation.Language.CommandAst]) {
                $name = "$($node.GetCommandName())".Split("\")[-1]
                if ($name -in @("removeDirectoryRetry", "moveDirectoryRetry")) {
                    $call = $name
                }
            } elseif ($node.Static -and $node.Expression -is [System.Management.Automation.Language.TypeExpressionAst] -and
                $node.Expression.TypeName.Name -match "(^|\.)Directory$" -and "$($node.Member.Value)" -in @("Move", "CreateDirectory")) {
                $call = "Directory.$($node.Member.Value)"
            }
            if (!$call) {
                continue
            }
            $owner = "(スクリプト直下)"
            for ($parent = $node.Parent; $null -ne $parent; $parent = $parent.Parent) {
                if ($parent -is [System.Management.Automation.Language.FunctionDefinitionAst] -or $parent -is [System.Management.Automation.Language.FunctionMemberAst]) {
                    $owner = $parent.Name
                    break
                }
            }
            @{ Owner = $owner; Call = $call }
        }
    }

    function getScriptFiles {
        $files = @(Get-ChildItem -LiteralPath ${scriptsDir} -Filter "*.ps1" -Recurse)
        $files.Count | Should -BeGreaterThan 0 -Because "scripts の下の .ps1 が見つからないと、何も調べずに通ってしまう"
        return $files
    }

    function getRelativeName {
        # scripts からの相対パス（違反の表示用）
        param (
            $file
        )
        return $file.FullName.Substring(${scriptsDir}.Length).TrimStart("\")
    }
}

Describe "フォルダを消す入口" -Tag Meta {
    It "scripts の下に、フォルダを再帰で消す書き方が無い（消すのは removeDirectoryRetry を通す）" {
        $hits = New-Object System.Collections.Generic.List[string]
        foreach ($file in getScriptFiles) {
            $errors = $null
            $tokens = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
            foreach ($hit in @(findRecursiveDeletes $ast)) {
                $hits.Add("$(getRelativeName $file) › $hit")
            }
        }
        @($hits) | Should -BeNullOrEmpty -Because "フォルダは removeDirectoryRetry（shared/core/fs.ps1）で消す。別の消し方を足さない"
    }

    It "removeTreeNoFollow を呼ぶのは removeDirectoryRetry だけ（消えたことを確かめる入口を通す）" {
        $callers = New-Object System.Collections.Generic.List[string]
        foreach ($file in getScriptFiles) {
            $errors = $null
            $tokens = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
            foreach ($call in @(findCallsByOwner $ast)) {
                # 呼んでよいのは shared\core\fs.ps1 の removeDirectoryRetry だけ（同じ名前の関数を別のファイルに作っても通さない）
                $isEntry = $call.Owner -eq "removeDirectoryRetry" -and (getRelativeName $file) -eq "shared\core\fs.ps1"
                if ($call.Command -eq "removeTreeNoFollow" -and !$isEntry) {
                    $callers.Add("$(getRelativeName $file) › $($call.Owner)")
                }
            }
        }
        @($callers) | Should -BeNullOrEmpty
    }

    It "removeDirectoryRetry を呼ぶ関数は、フォルダの有無を Test-Path で見ない（[System.IO.Directory]::Exists で見る）" {
        $hits = New-Object System.Collections.Generic.List[string]
        foreach ($file in getScriptFiles) {
            $errors = $null
            $tokens = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
            $groups = groupCommandsByOwner @(findCallsByOwner $ast)
            foreach ($owner in $groups.Keys) {
                if ($groups[$owner] -contains "removeDirectoryRetry" -and $groups[$owner] -contains "Test-Path") {
                    $hits.Add("$(getRelativeName $file) › $owner")
                }
            }
        }
        @($hits) | Should -BeNullOrEmpty -Because "予約名に拡張子の付いた名前のフォルダは、Test-Path が無いと答えて消さずに記録だけが進む"
    }
}

Describe "フォルダを消す・移す・作る入口" -Tag Meta {
    BeforeAll {
        # 置き場の下かを確かめる入口（removeDirectoryUnder・moveDirectoryUnder・newDirectoryUnder）を通さずに、フォルダを消す・移す・作る呼び出し。
        # 足すときは Reason に、置き場が決まっている（確かめ済み・利用者が選んだ場所・作業の根）理由を書く。
        # File = ファイル名、Owner = 関数（クラスのメソッドを含む）の名前、Call = 呼び出し
        ${folderChangeAllowed} = @(
            @{ File = "shared\core\fs.ps1"; Owner = "removeDirectoryUnder"; Call = "removeDirectoryRetry"; Reason = "入口そのもの（置き場の下かを確かめてから消す）" }
            @{ File = "shared\core\fs.ps1"; Owner = "moveDirectoryUnder"; Call = "moveDirectoryRetry"; Reason = "入口そのもの（置き場の下かを確かめてから移す）" }
            @{ File = "shared\core\fs.ps1"; Owner = "newDirectoryUnder"; Call = "Directory.CreateDirectory"; Reason = "入口そのもの（置き場の下かを確かめてから作る）" }
            @{ File = "shared\core\fs.ps1"; Owner = "moveDirectoryRetry"; Call = "Directory.Move"; Reason = "移す処理の本体（呼ぶ側が moveDirectoryUnder で確かめる）" }
            @{ File = "shared\core\fs.ps1"; Owner = "writeListFile"; Call = "Directory.CreateDirectory"; Reason = "書き込む 1 ファイルの親フォルダを作るだけ（置き場は呼ぶ側が決めたファイルのパス。消さない）" }
            @{ File = "shared\core\fs.ps1"; Owner = "writeTextLinesAtomic"; Call = "Directory.CreateDirectory"; Reason = "書き込む 1 ファイルの親フォルダを作るだけ（消さない）" }
            @{ File = "shared\office\office_process.ps1"; Owner = "addOfficeRecord"; Call = "Directory.CreateDirectory"; Reason = "Office の記録ファイルの親フォルダを作るだけ（ワークスペースの外の一時の場所。消さない）" }
            @{ File = "shared\ui\app_host.ps1"; Owner = "writeErrorLog"; Call = "Directory.CreateDirectory"; Reason = "エラーログの親フォルダを作るだけ（消さない）" }
            @{ File = "tebunko\core\workspace_mover.ps1"; Owner = "CopyTree"; Call = "Directory.CreateDirectory"; Reason = "ワークスペースを移すときの、移し先の中の複写。移し先は利用者が選んだ（検査済みの）フォルダで、リンクの先はたどらない" }
            @{ File = "tebunko\core\workspace_mover.ps1"; Owner = "MoveEntries"; Call = "Directory.CreateDirectory"; Reason = "ワークスペースを移すときの、利用者が選んだ移し先を作る" }
            @{ File = "tebunko\core\workspace_mover.ps1"; Owner = "MoveEntry"; Call = "Directory.Move"; Reason = "ワークスペースを移すときの、利用者が選んだ移し先への移動（検査済み）" }
            @{ File = "tebunko\core\workspace_mover.ps1"; Owner = "MoveEntry"; Call = "removeDirectoryRetry"; Reason = "ワークスペースを移し終えた元のフォルダと、失敗した移し先の途中を消す（利用者が選んだ移動の後始末）" }
            @{ File = "tebunko\core\workspace_mover.ps1"; Owner = "RemoveEntries"; Call = "removeDirectoryRetry"; Reason = "利用者が［消して、最初からやり直す］を選んだ回だけ、選んだフォルダの tebunko のフォルダを消す" }
            @{ File = "tebunko\index\index_archive.ps1"; Owner = "exportIndexCore"; Call = "Directory.CreateDirectory"; Reason = "書き出す zip の一時ファイルの親フォルダを作るだけ（消さない）" }
            @{ File = "tebunko\index\index_archive.ps1"; Owner = "swapInImportedIndexDir"; Call = "Directory.CreateDirectory"; Reason = "インデックスの置き場（IndexDir）そのものを作る（置き場の中ではなく、置き場自身。消さない）" }
            @{ File = "tebunko\index\index_format_store.ps1"; Owner = "updateIndexFormat"; Call = "Directory.CreateDirectory"; Reason = "目録ファイルの親フォルダを作るだけ（消さない）" }
            @{ File = "tebunko\index\index_store.ps1"; Owner = "publishIndexFiles"; Call = "Directory.Move"; Reason = "置き場の下かを testRemovablePath で、移す前に両側とも確かめてから移す（IOException のときの待ち直しのため、moveDirectoryUnder を使わない）" }
            @{ File = "tebunko\index\system_index.ps1"; Owner = "updateSystemIndexState"; Call = "Directory.CreateDirectory"; Reason = "状態ファイルの親フォルダを作るだけ（消さない）" }
            @{ File = "tebunko\indexer\indexer_run.ps1"; Owner = "invokeIndexerBody"; Call = "Directory.CreateDirectory"; Reason = "ワークスペースの根（PublishDir・IndexDir・Dir）そのものを作る（置き場自身。消さない）" }
            @{ File = "tebunko\indexer\indexer_run.ps1"; Owner = "invokeIndexer"; Call = "Directory.CreateDirectory"; Reason = "ワークスペースの根そのものを作る（置き場自身。消さない）" }
            @{ File = "tebunko\indexer\indexer_run.ps1"; Owner = "(スクリプト直下)"; Call = "Directory.CreateDirectory"; Reason = "ワークスペースの根そのものを作る（置き場自身。消さない）" }
            @{ File = "tebunko\ui\open_source.ps1"; Owner = "exportResults"; Call = "Directory.CreateDirectory"; Reason = "ワークスペースの根そのものを作る（置き場自身。消さない）" }
            @{ File = "tebunko\ui\settings\settings.ps1"; Owner = "applyWorkspace"; Call = "Directory.CreateDirectory"; Reason = "利用者が設定画面で選んだワークスペースの根を作る（置き場自身。消さない）" }
            @{ File = "tebunko\ui\settings\settings.ps1"; Owner = "resetWorkspace"; Call = "Directory.CreateDirectory"; Reason = "利用者が選んだワークスペースの根の下のフォルダを作る（消さない）" }
        )
    }

    It "フォルダを消す・移す・作る呼び出しは、入口を通すか、理由つきの一覧にある" {
        $unknown = New-Object System.Collections.Generic.List[string]
        foreach ($file in getScriptFiles) {
            $errors = $null
            $tokens = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
            $relative = getRelativeName $file
            foreach ($call in @(findFolderChangeCalls $ast)) {
                $hit = $call
                if (@(${folderChangeAllowed} | Where-Object { $_.File -eq $relative -and $_.Owner -eq $hit.Owner -and $_.Call -eq $hit.Call }).Count -eq 0) {
                    $unknown.Add("$relative › $($call.Owner) › $($call.Call)")
                }
            }
        }
        @($unknown | Sort-Object -Unique) | Should -BeNullOrEmpty -Because "消す・移す・作るは removeDirectoryUnder・moveDirectoryUnder・newDirectoryUnder（shared/core/fs.ps1）を通す。通せないなら、この一覧に理由つきで足す"
    }

    It "一覧の呼び出しは、いまも存在する（消した・名前を変えた呼び出しを一覧に残さない）" {
        $current = New-Object System.Collections.Generic.List[string]
        foreach ($file in getScriptFiles) {
            $errors = $null
            $tokens = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
            foreach ($call in @(findFolderChangeCalls $ast)) {
                $current.Add("$(getRelativeName $file)|$($call.Owner)|$($call.Call)")
            }
        }
        $stale = @(${folderChangeAllowed} | Where-Object { $current -notcontains "$($_.File)|$($_.Owner)|$($_.Call)" } |
                ForEach-Object { "$($_.File) › $($_.Owner) › $($_.Call)" })
        $stale | Should -BeNullOrEmpty
    }
}

Describe "フォルダを消す書き方の見つけ方（自己テスト）" -Tag Meta {
    It "再帰の削除を見つける: <code>" -TestCases @(
        @{ code = 'Remove-Item -LiteralPath $x -Recurse -Force' }
        @{ code = 'Remove-Item -LiteralPath $x -Recurse:$true' }
        @{ code = 'Remove-Item $x -rec' }
        @{ code = 'Remove-Item $x -r' }
        @{ code = 'rm -r $x' }
        @{ code = 'ri $x -Recurse' }
        @{ code = 'del $x -Recurse' }
        @{ code = 'rd $x -Recurse' }
        @{ code = 'rmdir $x -Recurse' }
        @{ code = 'erase $x -Recurse' }
        @{ code = 'Remove-Item @splat' }
        @{ code = '[System.IO.Directory]::Delete($x, $true)' }
        @{ code = '[IO.Directory]::Delete($x, $recursive)' }
        @{ code = '$info.Delete($true)' }
        @{ code = '([System.IO.DirectoryInfo]$x).Delete($flag)' }
        @{ code = '$info."Delete"($true)' }
        @{ code = '$info.''Delete''($true)' }
        @{ code = '[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($x, "DeleteAllContents")' }
        @{ code = '$fso.DeleteFolder($x, $true)' }
        @{ code = '$dirs | ForEach-Object Delete $true' }
        @{ code = '$dirs | foreach Delete $recursive' }
        @{ code = '$dirs | % Delete $true' }
        @{ code = '$dirs | ForEach-Object -MemberName Delete -ArgumentList $true' }
        @{ code = '$dirs | ForEach-Object -MemberName Delete -ArgumentList:$true' }
        @{ code = '$dirs | ForEach-Object -ArgumentList $true -MemberName Delete' }
        @{ code = '$dirs | ForEach-Object -MemberName:Delete -ArgumentList $true' }
        @{ code = '$dirs.ForEach("Delete", $true)' }
        @{ code = '$type::Delete($x, $true)' }
        @{ code = '$type::Delete($x, $recursive)' }
        @{ code = 'Microsoft.PowerShell.Management\Remove-Item $x -Recurse' }
        @{ code = 'Microsoft.PowerShell.Management\rm -r $x' }
    ) {
        @(findRecursiveDeletes (parseCode $code)).Count | Should -BeGreaterThan 0
    }

    It "再帰でない削除は見つけない: <code>" -TestCases @(
        @{ code = 'Remove-Item -LiteralPath $f -Force' }
        @{ code = '[System.IO.File]::Delete($f)' }
        @{ code = '[System.IO.Directory]::Delete($d, $false)' }
        @{ code = '[System.IO.Directory]::Delete($d)' }
        @{ code = '$temp.Delete()' }
        @{ code = '$info.Delete($false)' }
        @{ code = '$info."Delete"($false)' }
        @{ code = '$dirs | ForEach-Object Delete $false' }
        @{ code = '$dirs | ForEach-Object { $_.Name }' }
        @{ code = '$dirs | ForEach-Object Delete' }
        @{ code = '$dirs | ForEach-Object -MemberName Delete' }
        @{ code = '$dirs | ForEach-Object -MemberName:Delete' }
        @{ code = '$dirs.ForEach("Delete", $false)' }
        @{ code = '$dirs.ForEach({ $_.Name })' }
        @{ code = '$type::Delete($f)' }
        @{ code = '$type::Delete($d, $false)' }
        @{ code = 'Get-ChildItem $d | Remove-Item -Force' }
    ) {
        @(findRecursiveDeletes (parseCode $code)).Count | Should -Be 0
    }

    It "コマンド呼び出しを、属する関数ごとに分ける（関数の外の呼び出しも拾い、スクリプト直下として数える）" {
        $ast = parseCode 'function a { removeDirectoryRetry $x; Test-Path $x }; function b { Test-Path $y }; removeDirectoryRetry $z; Test-Path $w'
        $groups = groupCommandsByOwner @(findCallsByOwner $ast)
        $groups.Keys.Count | Should -Be 3
        $groups["a"] | Should -Be @("removeDirectoryRetry", "Test-Path")
        $groups["b"] | Should -Be @("Test-Path")
        $groups["(スクリプト直下)"] | Should -Be @("removeDirectoryRetry", "Test-Path")
    }

    It "入れ子の関数の呼び出しは、いちばん内側の関数に数え、モジュール名付きのコマンドは後ろの名前で数える" {
        $ast = parseCode 'function outer { function inner { Microsoft.PowerShell.Management\Test-Path $x }; removeTreeNoFollow $y }'
        $groups = groupCommandsByOwner @(findCallsByOwner $ast)
        $groups["inner"] | Should -Be @("Test-Path")
        $groups["outer"] | Should -Be @("removeTreeNoFollow")
    }
}
