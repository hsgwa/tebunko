# startJob の終わりの処理（3 つ目の引数のスクリプトブロック）は、あとからタイマーで呼ばれる。
# そのとき、囲んでいる関数のローカル変数（引数・代入した変数）は残っていないため、空になる。
# 終わりの処理が、囲む関数のローカル変数を使うなら、.GetNewClosure() で今の値を取り込むこと（関数は ${function:名前} で変数に取って & で呼ぶ）。
# 関数の名前に頼らず、構文の木で調べる。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"

    function findUncapturedDoneLocals {
        # 画面のファイルで、startJob の終わりの処理が取り込まずに使っている、囲む関数のローカル変数を "ファイル:行 $名前" の配列で返す
        param ([string]$directory)

        $found = New-Object System.Collections.Generic.List[string]
        foreach ($file in @(Get-ChildItem -LiteralPath $directory -Filter "*.ps1" -Recurse)) {
            $errors = $null
            $tokens = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
            $commands = $ast.FindAll({
                    param ($n)
                    $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq "startJob"
                }, $true)
            foreach ($command in $commands) {
                # 終わりの処理は 4 番目の要素（startJob・スクリプト・引数・終わりの処理）。GetNewClosure() を付けたものは式なので数えない
                $done = $command.CommandElements | Select-Object -Skip 3 -First 1
                if ($done -isnot [System.Management.Automation.Language.ScriptBlockExpressionAst]) {
                    continue
                }
                $function = $command.Parent
                while ($null -ne $function -and $function -isnot [System.Management.Automation.Language.FunctionDefinitionAst]) {
                    $function = $function.Parent
                }
                if ($null -eq $function) {
                    continue
                }
                $isInDone = {
                    param ($n)
                    for ($p = $n; $null -ne $p; $p = $p.Parent) { if ($p -eq $done) { return $true } }
                    return $false
                }
                # 関数のローカル変数: 関数の引数と、終わりの処理の外での代入
                $locals = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
                if ($null -ne $function.Body.ParamBlock) {
                    foreach ($p in $function.Body.ParamBlock.Parameters) { [void]$locals.Add($p.Name.VariablePath.UserPath) }
                }
                foreach ($p in @($function.Parameters)) { [void]$locals.Add($p.Name.VariablePath.UserPath) }
                $assigns = $function.Body.FindAll({
                        param ($n)
                        $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
                        $n.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
                        !(& $isInDone $n)
                    }, $true)
                foreach ($a in $assigns) {
                    if ($a.Left.VariablePath.IsUnscopedVariable) { [void]$locals.Add($a.Left.VariablePath.UserPath) }
                }
                # 終わりの処理の中の、自分の引数・代入は数えない
                $own = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
                foreach ($p in @($done.ScriptBlock.ParamBlock.Parameters)) { [void]$own.Add($p.Name.VariablePath.UserPath) }
                foreach ($a in $done.ScriptBlock.FindAll({ param ($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)) {
                    [void]$own.Add($a.Left.VariablePath.UserPath)
                }
                $used = $done.ScriptBlock.FindAll({ param ($n) $n -is [System.Management.Automation.Language.VariableExpressionAst] -and $n.VariablePath.IsUnscopedVariable }, $true)
                foreach ($v in $used) {
                    $name = $v.VariablePath.UserPath
                    if ($locals.Contains($name) -and !$own.Contains($name)) {
                        $found.Add("$($file.Name):$($v.Extent.StartLineNumber) `$$name")
                    }
                }
            }
        }
        return $found.ToArray()
    }
}

Describe "startJob の終わりの処理は、囲む関数のローカル変数を取り込む" -Tag Meta {
    It "画面のファイルの終わりの処理は、取り込まずに囲む関数のローカル変数を使わない" {
        $uncaptured = @(findUncapturedDoneLocals "${scriptsDir}\tebunko\ui")
        $uncaptured | Should -BeNullOrEmpty -Because "終わりの処理はあとから呼ばれ、囲む関数のローカル変数は空になる（.GetNewClosure() で取り込むこと）"
    }

    It "検査が働く（取り込まない使用を見つける・GetNewClosure・自分の引数・script スコープは見逃す）" {
        $dir = Join-Path $TestDrive "ui"
        New-Item -ItemType Directory -Path $dir | Out-Null
        $text = @'
function bad {
    param ($folder)
    $mode = "x"
    startJob { 1 } @() {
        param ($output)
        finish $mode $folder $output $script:flag $ui
    }
}
function good {
    param ($folder)
    startJob { 1 } @() {
        param ($output)
        finish $folder $output
    }.GetNewClosure()
}
'@
        [System.IO.File]::WriteAllText((Join-Path $dir "sample.ps1"), $text)
        $result = @(findUncapturedDoneLocals $dir)
        $result | Should -Be @('sample.ps1:6 $mode', 'sample.ps1:6 $folder')
    }
}
