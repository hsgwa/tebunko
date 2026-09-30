# `return , <集合>` で返す関数を @(関数名 ...) で受ける書き方を見つけたら落とす。
# `return , $set` は集合を 1 つのオブジェクトのまま返す（そのまま return すると PowerShell が中身を展開し、
# 0 件なら $null・1 件なら 1 つの値になるため）。これを @(...) で包むと、集合が 1 要素の配列に入るだけになり、
# -contains・foreach・Count が中身ではなく集合そのものに当たる（インポートの名前の重なりを見落とし、常に別名になった）。
# 変数に受けてから使う（$set = 関数名 ...; @($set)）。
#
# 見つけるもの: `return , ...` を持つ関数（scripts の全体で集める）を、単独のコマンドとして @(...) の中で呼ぶ書き方。
# パイプラインに流す `@(関数名 ... | Where-Object ...)` は、パイプラインが中身を展開するため対象にしない。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"

    function getCommaReturnFunctionNames {
        # ファイル群の中の、`return , <式>` を持つ関数の名前
        param ([string[]]$paths)

        $names = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($path in $paths) {
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$null, [ref]$null)
            $functions = $ast.FindAll({ param ($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
            foreach ($function in $functions) {
                $returns = $function.Body.FindAll({
                        param ($node)
                        $node -is [System.Management.Automation.Language.ReturnStatementAst] -and
                        $node.Pipeline -is [System.Management.Automation.Language.PipelineAst] -and
                        $node.Pipeline.PipelineElements.Count -eq 1 -and
                        $node.Pipeline.PipelineElements[0] -is [System.Management.Automation.Language.CommandExpressionAst] -and
                        $node.Pipeline.PipelineElements[0].Expression -is [System.Management.Automation.Language.ArrayLiteralAst]
                    }, $false)
                if (@($returns).Count -gt 0) {
                    [void]$names.Add($function.Name)
                }
            }
        }
        return , $names
    }

    function findWrappedCommaReturnCalls {
        # ファイル path の中の、@(関数名 ...) の書き方を「<ファイル名>:<行> @(<関数名> ...)」の並びで返す
        param ([string]$path, $names)

        $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$null, [ref]$null)
        $file = [System.IO.Path]::GetFileName($path)
        $found = New-Object System.Collections.Generic.List[string]
        $arrays = $ast.FindAll({ param ($node) $node -is [System.Management.Automation.Language.ArrayExpressionAst] }, $true)
        foreach ($array in $arrays) {
            $statements = @($array.SubExpression.Statements)
            if ($statements.Count -ne 1 -or $statements[0] -isnot [System.Management.Automation.Language.PipelineAst]) {
                continue
            }
            $elements = @($statements[0].PipelineElements)
            if ($elements.Count -ne 1 -or $elements[0] -isnot [System.Management.Automation.Language.CommandAst]) {
                continue
            }
            $command = $elements[0].GetCommandName()
            if ($command -and $names.Contains($command)) {
                $found.Add("${file}:$($array.Extent.StartLineNumber) @(${command} ...)")
            }
        }
        return , $found.ToArray()
    }
}

Describe "return , で返す関数を @() で包まない" -Tag Meta {
    It "scripts の下に、@(関数名 ...) で受ける書き方が無い" {
        $paths = @(Get-ChildItem -LiteralPath $scriptsDir -Recurse -Filter *.ps1 | ForEach-Object { $_.FullName })
        $names = getCommaReturnFunctionNames $paths
        $names.Count | Should -BeGreaterThan 0
        $found = New-Object System.Collections.Generic.List[string]
        foreach ($path in $paths) {
            foreach ($item in (findWrappedCommaReturnCalls $path $names)) {
                $found.Add($item)
            }
        }
        $found | Should -BeNullOrEmpty -Because "return , で返す関数は、変数に受けてから @(`$変数) にする（@(関数名 ...) だと集合が 1 要素の配列に入る）"
    }

    It "見つける: <name>" -TestCases @(
        @{ name = "単独の呼び出し"; code = 'function f { return , $s }; $u = @(f $x)' }
        @{ name = "引数なし"; code = 'function f { return , $s }; $u = @(f)' }
        @{ name = "-contains の左"; code = 'function f { return , $s }; if (@(f $x) -contains "a") { }' }
    ) {
        param ($name, $code)
        $path = "$TestDrive\found.ps1"
        [System.IO.File]::WriteAllText($path, $code)
        $names = getCommaReturnFunctionNames @($path)

        (findWrappedCommaReturnCalls $path $names).Count | Should -Be 1
    }

    It "見つけない: <name>" -TestCases @(
        @{ name = "変数に受ける"; code = 'function f { return , $s }; $set = f $x; $u = @($set)' }
        @{ name = "パイプラインに流す"; code = 'function f { return , $s }; $u = @(f $x | Where-Object { $_ })' }
        @{ name = "return , の無い関数"; code = 'function f { return $s }; $u = @(f $x)' }
        @{ name = "別の関数"; code = 'function f { return , $s }; $u = @(g $x)' }
    ) {
        param ($name, $code)
        $path = "$TestDrive\notfound.ps1"
        [System.IO.File]::WriteAllText($path, $code)
        $names = getCommaReturnFunctionNames @($path)

        (findWrappedCommaReturnCalls $path $names).Count | Should -Be 0
    }
}
