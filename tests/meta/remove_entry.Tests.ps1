# フォルダを消す入口を 1 つに保つ。
# フォルダを中身ごと消すのは removeDirectoryRetry（shared/core/fs.ps1。中でリンクを辿らない removeTreeNoFollow を呼び、消えたことを .NET で確かめる）だけにする。
# 別の消し方（Remove-Item -Recurse・Directory.Delete の再帰）を足すと、リンクの先まで消す・消えていないのに成功と記録する、のどちらかが起きうる。
# 見張るもの:
#   1. scripts の下に、フォルダを再帰で消す書き方（Remove-Item とその別名の -Recurse（省略形も）・スプラット、Directory.Delete の再帰）が無い
#   2. removeTreeNoFollow を呼ぶのは removeDirectoryRetry だけ
#   3. removeDirectoryRetry を呼ぶ関数は、フォルダの有無を Test-Path で見ない（予約名に拡張子の付いた名前のフォルダは、Test-Path と .NET で答えが食い違い、
#      消さずに「削除しました」と記録する不具合になった。有無は [System.IO.Directory]::Exists で見る）
# 1 つのファイルを消す書き方（File.Delete・Remove-Item -Force）は対象にしない（リンクの先を消さないため）。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"

    ${removeItemNames} = @("Remove-Item", "ri", "rm", "rmdir", "rd", "del", "erase")

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
                if (${removeItemNames} -notcontains $node.GetCommandName()) {
                    continue
                }
                foreach ($element in $node.CommandElements) {
                    if ($element -is [System.Management.Automation.Language.CommandParameterAst] -and
                        "Recurse".StartsWith($element.ParameterName, [System.StringComparison]::OrdinalIgnoreCase)) {
                        $found.Add("$($node.GetCommandName()) -$($element.ParameterName)")
                    } elseif ($element -is [System.Management.Automation.Language.VariableExpressionAst] -and $element.Splatted) {
                        $found.Add("$($node.GetCommandName()) @$($element.VariablePath.UserPath)（スプラット。-Recurse が入るか分からない）")
                    }
                }
                continue
            }
            if ($node.Member.Extent.Text -ne "Delete") {
                continue
            }
            $argumentCount = if ($null -eq $node.Arguments) { 0 } else { $node.Arguments.Count }
            if ($node.Static) {
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

    function findCallsInFunctions {
        # ast の中の関数（クラスのメソッドを含む）ごとに、呼んでいるコマンドの名前を @{ Function; Commands } で返す
        param (
            $ast
        )
        $functions = $ast.FindAll({
                param ($n)
                $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -or $n -is [System.Management.Automation.Language.FunctionMemberAst]
            }, $true)
        foreach ($function in $functions) {
            $names = @($function.FindAll({ param ($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) | ForEach-Object { $_.GetCommandName() } | Where-Object { $_ })
            @{ Function = $function.Name; Commands = $names }
        }
    }

    function getScriptFiles {
        return @(Get-ChildItem -LiteralPath ${scriptsDir} -Filter "*.ps1" -Recurse)
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
                $hits.Add("$($file.Name) › $hit")
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
            foreach ($function in @(findCallsInFunctions $ast)) {
                if ($function.Commands -contains "removeTreeNoFollow" -and $function.Function -ne "removeDirectoryRetry") {
                    $callers.Add("$($file.Name) › $($function.Function)")
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
            foreach ($function in @(findCallsInFunctions $ast)) {
                if ($function.Commands -contains "removeDirectoryRetry" -and $function.Commands -contains "Test-Path") {
                    $hits.Add("$($file.Name) › $($function.Function)")
                }
            }
        }
        @($hits) | Should -BeNullOrEmpty -Because "予約名に拡張子の付いた名前のフォルダは、Test-Path が無いと答えて消さずに記録だけが進む"
    }

    It "removeDirectoryRetry は、消したあとにフォルダが残っていれば例外にする（Directory.Exists で確かめる）" {
        $errors = $null
        $tokens = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile("${scriptsDir}\shared\core\fs.ps1", [ref]$tokens, [ref]$errors)
        $function = $ast.Find({ param ($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq "removeDirectoryRetry" }, $true)
        $function | Should -Not -BeNullOrEmpty
        $text = $function.Extent.Text
        $text.IndexOf("removeTreeNoFollow") | Should -BeGreaterThan -1
        $text.IndexOf("Directory]::Exists(`$longPath)", $text.IndexOf("removeTreeNoFollow")) | Should -BeGreaterThan $text.IndexOf("removeTreeNoFollow")
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
        @{ code = 'Get-ChildItem $d | Remove-Item -Force' }
    ) {
        @(findRecursiveDeletes (parseCode $code)).Count | Should -Be 0
    }

    It "関数ごとのコマンド呼び出しを数える（関数の外のコマンドは数えない）" {
        $ast = parseCode 'function a { removeDirectoryRetry $x; Test-Path $x }; function b { Test-Path $y }; removeDirectoryRetry $z'
        $result = @(findCallsInFunctions $ast)
        $result.Count | Should -Be 2
        ($result | Where-Object { $_.Function -eq "a" }).Commands | Should -Be @("removeDirectoryRetry", "Test-Path")
        ($result | Where-Object { $_.Function -eq "b" }).Commands | Should -Be @("Test-Path")
    }
}
