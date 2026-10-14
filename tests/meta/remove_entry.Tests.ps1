# フォルダを消す入口を 1 つに保つ。
# フォルダを中身ごと消すのは removeDirectoryRetry（shared/core/fs.ps1。中でリンクを辿らない removeTreeNoFollow を呼び、消えたことを .NET で確かめる）だけにする。
# 別の消し方（Remove-Item -Recurse・Directory.Delete の再帰）を足すと、リンクの先まで消す・消えていないのに成功と記録する、のどちらかが起きうる。
# 見張るもの:
#   1. scripts の下に、フォルダを再帰で消す書き方（Remove-Item とその別名（モジュール名付きも）の -Recurse（省略形も）・スプラット、
#      Directory.Delete の再帰、型が変数で読めない ::Delete(<path>, <$false でない値>)）が無い
#   2. removeTreeNoFollow を呼ぶのは removeDirectoryRetry だけ（スクリプト直下の呼び出しも違反に数える）
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
                    # ForEach-Object Delete $true（メンバー名を文字列で渡す書き方）。Delete のあとに $false でない値があれば違反に数える
                    $elements = @($node.CommandElements)
                    for ($index = 0; $index -lt $elements.Count; $index++) {
                        if ($elements[$index] -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $elements[$index].Value -eq "Delete") {
                            if ($index + 1 -ge $elements.Count) {
                                continue
                            }
                            $rest = @($elements[($index + 1)..($elements.Count - 1)] | Where-Object { $_ -is [System.Management.Automation.Language.ExpressionAst] -and !(testFalseLiteral $_) })
                            if ($rest.Count -gt 0) {
                                $found.Add("$commandName Delete（引数付き。再帰かもしれない）")
                            }
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
