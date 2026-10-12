# startJob の終わりの処理（3 つ目の引数のスクリプトブロック）は、あとからタイマーで呼ばれる。
# そのとき、囲んでいる関数のローカル変数（引数・代入した変数）は残っていないため、空になる。
# 終わりの処理が、囲む関数のローカル変数を使うなら、.GetNewClosure() で今の値を取り込むこと（関数は ${function:名前} で変数に取って & で呼ぶ）。
# 関数の名前に頼らず、構文の木で調べる。
# 対象: startJob の 4 番目の要素（終わりの処理）と、[scriptblock] 型で名前が on で始まる引数（$onDone・$onSuccess など）を持つ関数の、その引数に渡すスクリプトブロック。
# 見ない形（これらは検査が拾えない。書くときは自分で取り込む）:
#   - 終わりの処理を変数で渡す形（startJob ... $block）。式で渡す形は、GetNewClosure() を付けたものを含め、スクリプトブロックそのものではないので数えない
#   - 囲む関数の中の、型を付けた代入（[string]$x = ...）・foreach の変数・複数代入（$a, $b = ...）で作ったローカル変数。ローカル変数として数えない
#   - 引数を持つ関数を、名前を使わずに呼ぶ形（& $変数 ...・. ...）。on で始まる引数を、スプラットや位置のずれで渡す形
#   - イベントのハンドラーやタイマーのスクリプトブロック（対象外）
#   - 調べるフォルダ（tebunko\ui）の外（shared\ui など）で定義された、on で始まる引数を持つ関数への呼び出し（引数の表に載らないので見ない）
#   - その場で同期に呼ぶコールバック（$onDone という名前でも、あとから呼ばれない）で、囲む関数のローカル変数を使う形。動くのに、検査は落ちる
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"

    function getDoneBlockParameters {
        # on で始まる [scriptblock] 型の引数を持つ関数を、"関数名" → @{ Name; Position }[] の表で返す。Position は引数の並びの中の位置（0 から）
        param ($asts)

        $table = @{}
        foreach ($ast in $asts) {
            foreach ($function in $ast.FindAll({ param ($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
                $parameters = @()
                if ($null -ne $function.Body.ParamBlock) { $parameters = @($function.Body.ParamBlock.Parameters) }
                elseif ($null -ne $function.Parameters) { $parameters = @($function.Parameters) }
                for ($i = 0; $i -lt $parameters.Count; $i++) {
                    $parameter = $parameters[$i]
                    $isBlock = $parameter.Attributes | Where-Object { $_ -is [System.Management.Automation.Language.TypeConstraintAst] -and $_.TypeName.Name -eq "scriptblock" }
                    $name = $parameter.Name.VariablePath.UserPath
                    if ($isBlock -and $name -match '^on[A-Z]') {
                        if (!$table.ContainsKey($function.Name)) { $table[$function.Name] = @() }
                        $table[$function.Name] += @{ Name = $name; Position = $i }
                    }
                }
            }
        }
        return $table
    }

    function getDoneBlocks {
        # 呼び出し 1 つの、終わりの処理として渡されたスクリプトブロック（式）を配列で返す
        param ($command, $table)

        $blocks = @()
        $commandName = $command.GetCommandName()
        if ($commandName -eq "startJob") {
            # 終わりの処理は 4 番目の要素（startJob・スクリプト・引数・終わりの処理）
            $blocks += @($command.CommandElements | Select-Object -Skip 3 -First 1)
        } elseif ($null -ne $commandName -and $table.ContainsKey($commandName)) {
            $elements = @($command.CommandElements | Select-Object -Skip 1)
            $positional = New-Object System.Collections.Generic.List[object]
            $named = @{}
            for ($i = 0; $i -lt $elements.Count; $i++) {
                $element = $elements[$i]
                if ($element -is [System.Management.Automation.Language.CommandParameterAst]) {
                    $value = if ($null -ne $element.Argument) { $element.Argument } elseif ($i + 1 -lt $elements.Count) { $i++; $elements[$i] } else { $null }
                    $named[$element.ParameterName] = $value
                } else {
                    $positional.Add($element)
                }
            }
            foreach ($entry in $table[$commandName]) {
                if ($named.ContainsKey($entry.Name)) { $blocks += $named[$entry.Name] }
                elseif ($entry.Position -lt $positional.Count) { $blocks += $positional[$entry.Position] }
            }
        }
        return @($blocks | Where-Object { $_ -is [System.Management.Automation.Language.ScriptBlockExpressionAst] })
    }

    function findUncapturedDoneLocals {
        # 画面のファイルで、終わりの処理が取り込まずに使っている、囲む関数のローカル変数を "ファイル:行 $名前" の配列で返す
        param ([string]$directory)

        $found = New-Object System.Collections.Generic.List[string]
        $parsed = @()
        foreach ($file in @(Get-ChildItem -LiteralPath $directory -Filter "*.ps1" -Recurse)) {
            $errors = $null
            $tokens = $null
            $parsed += , @{ File = $file; Ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors) }
        }
        $table = getDoneBlockParameters @($parsed | ForEach-Object { $_.Ast })
        foreach ($item in $parsed) {
            $file = $item.File
            $commands = $item.Ast.FindAll({ param ($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)
            foreach ($command in $commands) {
                foreach ($done in @(getDoneBlocks $command $table)) {
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
                    # 関数のローカル変数: 関数の引数と、終わりの処理の外での代入（$null への代入は変数ではない）
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
                        if ($a.Left.VariablePath.IsUnscopedVariable -and $a.Left.VariablePath.UserPath -ne "null") { [void]$locals.Add($a.Left.VariablePath.UserPath) }
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
        }
        return $found.ToArray()
    }
}

Describe "startJob の終わりの処理は、囲む関数のローカル変数を取り込む" -Tag Meta {
    It "画面のファイルの終わりの処理は、取り込まずに囲む関数のローカル変数を使わない" {
        $uncaptured = @(findUncapturedDoneLocals "${scriptsDir}\tebunko\ui")
        $uncaptured | Should -BeNullOrEmpty -Because "終わりの処理はあとから呼ばれ、囲む関数のローカル変数は空になる（.GetNewClosure() で取り込むこと）"
    }

    It "検査が働く（取り込まない使用を見つける・GetNewClosure・自分の引数・script スコープ・null への代入は見逃す。on で始まる引数を持つ関数の呼び出しも、位置でも名前でも見る）" {
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
function nullAssign {
    param ($folder)
    $null = Get-Date
    startJob { 1 } @() {
        param ($output)
        $null = finish $output
    }
}
function startSample {
    param (
        [string]$kind,
        [scriptblock]$onDone,
        [scriptblock]$notCallback
    )
}
function callerPositional {
    param ($folder)
    startSample "a" {
        finish $folder
    } { $folder }
}
function callerNamed {
    param ($folder)
    startSample -kind "a" -onDone { finish $folder }
}
function callerCaptured {
    param ($folder)
    startSample "a" { finish $folder }.GetNewClosure()
}
'@
        [System.IO.File]::WriteAllText((Join-Path $dir "sample.ps1"), $text)
        $result = @(findUncapturedDoneLocals $dir)
        $result | Should -Be @('sample.ps1:6 $mode', 'sample.ps1:6 $folder', 'sample.ps1:34 $folder', 'sample.ps1:39 $folder')
    }
}
