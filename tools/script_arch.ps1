# スクリプトの層の決まり（依存の向き）を、AST で関数の呼び出し・変数・型の使い方まで読んで確かめる部品。
# tests/meta/layers.Tests.ps1 が読む。tebunko のフォルダ構成は知らない（どのパスがどのパスに依存してよいかは、呼び出し側が表で渡す）。
#
# 全部のファイルが同じスコープに読み込まれるので、下の層が上の層の関数・変数を名前で使っても動いてしまう。
# 文字列を探す検査では、関数や変数の使われ方まで見えない。ここでは、ファイルごとに構文木を作り、次を見る。
#
# - 関数の呼び出し（コマンドの名前）が、別のパスのファイルで定義した関数を指していること
# - 別のパスのファイルが最上位（関数の外）・$script:・$global: で代入した変数を、読んでいること
# - 決めた変数（$ui・$window など）と型（System.Windows.* など）に触っていること
# - Set-Variable・New-Variable の -Name が定数のものは、代入として数えること（関数の外、または -Scope が Script・Global・1 以上）
# - 別のファイルで定義したクラスを、型（[型]::new()・引数・戻り値・キャスト）として使っていること（名前は [型名]）
# - ${function:名前} を、その名前の関数の呼び出しとして数えること
# - Get-Variable（gv）の名前・${variable:名前}・$variable:名前・Get-Item などの variable:名前 が定数のものは、その変数の読みとして数えること
# - New-Object の型（-TypeName は位置によらない）の名前が定数のものは、[ ] , で区切った葉のクラスを型として使う依存として数えること（名前は [型名]）
# - $variable:名前 = … の代入は、その名前の変数の代入として数えること
# - モジュール名の修飾（Microsoft.PowerShell.Utility\Get-Command など）は外して、コマンドの名前を比べること
#
# 構文木から呼ぶ先が決まらない書き方は、findStringNameCalls が「形」ごとに返す（呼び出し側の閉じた一覧に載せて確かめる）。
# 数える形: & (式)・& "展開文字列"、Get-Command、Invoke-Expression、[scriptblock]::Create、$ExecutionContext.InvokeCommand、
# .AddCommand・.AddScript、Set-Alias・New-Alias、
# function: ドライブを引く Get-Item・Get-Content・Test-Path・Set-Item・New-Item など（読む・書くのどちらも）、
# 名前が定数でない Set-Variable・New-Variable・Get-Variable、名前が定数でない variable: ドライブの読み、
# variable: ドライブへの書き込み（Set-Item・New-Item・Remove-Item・Copy-Item）、型が定数でない New-Object。
# & $name・. $name は、呼ぶ値（スクリプトブロック）を作った所の呼び出し・${function:名前} で数えるので、ここでは数えない。
#
# 見つからないもの（限界）: コールバックで渡したスクリプトブロックの中から呼ぶ形（値を作った所で数える）、
# $ui を別の名前の引数で受け取る形、クラスのメソッドの呼び出し。
# 同じ名前の関数・変数・クラスが、見る側のパスにもあるときは、そちらを使っているとみなして返さない。
#
# 使い方:
#   $arch = newScriptArch $scriptsDir                                   # 全ファイルの解析（1 回だけ作って使い回す）
#   findFileDependencies $arch -From "shared/*" -To "tebunko/*"         # "<from のパス> -> <to のパス>（<名前>・<名前>）" の一覧
#   findTypeUses $arch -In "*_view.ps1" -Variable ui, window -TypeLike "System.Windows.*"
#
# PowerShell 5.1 で動く書き方にする（7 だけの API・?.・?? は使わない）。

# パスが、ワイルドカードのどれかに当たるか（パスは scripts からの相対で、区切りは /）
function testPathLike {
    param (
        [string]$path,
        [string[]]$patterns
    )

    foreach ($pattern in @($patterns)) {
        if ($pattern -and $path -like $pattern) { return $true }
    }
    return $false
}

# 代入の左辺にある変数のパス（VariablePath）を返す。$a = ・[int]$a = ・[string][int]$a = ・$a, $b = の形
function getAssignedVariablePaths {
    param ($left)

    while ($left -is [System.Management.Automation.Language.ConvertExpressionAst]) { $left = $left.Child }
    if ($left -is [System.Management.Automation.Language.VariableExpressionAst]) { return @($left.VariablePath) }
    $result = @()
    if ($left -is [System.Management.Automation.Language.ArrayLiteralAst]) {
        foreach ($element in $left.Elements) {
            $result += @(getAssignedVariablePaths $element)
        }
    }
    return $result
}

# ノードが、関数（クラスのメソッドを含む）の中にあるか
function testInsideFunction {
    param ($node)

    $parent = $node.Parent
    while ($parent) {
        if ($parent -is [System.Management.Automation.Language.FunctionDefinitionAst] -or
            $parent -is [System.Management.Automation.Language.FunctionMemberAst]) {
            return $true
        }
        $parent = $parent.Parent
    }
    return $false
}

# 変数のスコープの印（script: など）を外した名前
function getVariableBaseName {
    param ($variablePath)

    return (removeVariableScope $variablePath.UserPath)
}

# 名前の先頭のスコープの印（script: など）を、印と残りの名前に分ける。印が無ければ Scope は空
function splitVariableScope {
    param ([string]$name)

    if ($name -match '^(script|global|local|private):(.*)$') {
        return [pscustomobject]@{ Scope = $Matches[1]; Name = $Matches[2] }
    }
    return [pscustomobject]@{ Scope = ''; Name = $name }
}

# スコープの印（script: など）を外した名前
function removeVariableScope {
    param ([string]$name)

    return (splitVariableScope $name).Name
}

# 関数・スクリプトブロックの中にある引数・代入・foreach の変数の名前（その中では自分のものとみなす名前）の集まりを返す。
# $cache（ノード → 名前の集まり）に覚えておき、同じものを何度も数えない
function getOwnedVariableNames {
    param (
        $scope,
        $cache
    )

    # 集まり（HashSet）をそのまま返すと、受け取る側で 1 つずつの要素にほどけるため、, で包んで返す
    if ($cache.ContainsKey($scope)) { return , $cache[$scope] }
    $names = New-Object 'System.Collections.Generic.HashSet[string]' -ArgumentList ([System.StringComparer]::OrdinalIgnoreCase)
    $owned = $scope.FindAll({
            param ($n)
            ($n -is [System.Management.Automation.Language.ParameterAst]) -or
            ($n -is [System.Management.Automation.Language.AssignmentStatementAst]) -or
            ($n -is [System.Management.Automation.Language.ForEachStatementAst])
        }, $true)
    foreach ($item in $owned) {
        if ($item -is [System.Management.Automation.Language.ParameterAst]) {
            [void]$names.Add((getVariableBaseName $item.Name.VariablePath))
        } elseif ($item -is [System.Management.Automation.Language.AssignmentStatementAst]) {
            foreach ($path in @(getAssignedVariablePaths $item.Left)) {
                if (-not $path.IsScript -and -not $path.IsGlobal) { [void]$names.Add((getVariableBaseName $path)) }
            }
        } elseif ($item.Variable) {
            [void]$names.Add((getVariableBaseName $item.Variable.VariablePath))
        }
    }
    $cache[$scope] = $names
    return , $names
}

# 変数の使い方が、それを囲む関数・スクリプトブロックの引数・代入・foreach の変数のどれかとして、自分のものか
function testLocalVariable {
    param (
        $node,
        [string]$name,
        $cache
    )

    $scope = $node.Parent
    while ($scope) {
        if ($scope -is [System.Management.Automation.Language.FunctionDefinitionAst] -or
            $scope -is [System.Management.Automation.Language.FunctionMemberAst] -or
            $scope -is [System.Management.Automation.Language.ScriptBlockExpressionAst]) {
            if ((getOwnedVariableNames $scope $cache).Contains($name)) { return $true }
        }
        $scope = $scope.Parent
    }
    return $false
}

# 定数の文字列・数の式の値を文字列で返す。定数でなければ $null
function getConstantText {
    param ($ast)

    if ($ast -is [System.Management.Automation.Language.ConstantExpressionAst]) { return [string]$ast.Value }
    return $null
}

# モジュール名の修飾（Microsoft.PowerShell.Utility\Get-Command など）を外したコマンドの名前
function getBareCommandName {
    param ([string]$name)

    return ($name -replace '^.*\\', '')
}

# 名前にワイルドカードの文字が入っているか（入っていれば、1 つの名前に決まらない）
function testWildcardName {
    param ([string]$name)

    return [bool]($name -match '[*?\[\]]')
}

# Set-Variable・New-Variable（$kind が Set）・Get-Variable（$kind が Get）の呼び出しの -Name（位置引数を含む）と -Scope の式を返す。
# 引数の名前は、前の部分だけでも（-N など）一つに決まれば読む。値を取る共通の引数（-ErrorAction など）の値は、位置引数に数えない
function getVariableCommandInfo {
    param (
        $commandAst,
        [string]$kind = 'Set'
    )

    $valueParams = @('Name', 'Include', 'Exclude', 'Scope')
    $switchParams = @('Verbose', 'Debug', 'vb', 'db')
    if ($kind -eq 'Set') {
        $valueParams += @('Value', 'Description', 'Option', 'Visibility')
        $switchParams += @('Force', 'PassThru', 'Confirm', 'WhatIf')
    } else {
        $switchParams += @('ValueOnly')
    }
    $valueParams += @('ErrorAction', 'WarningAction', 'InformationAction', 'ProgressAction', 'ErrorVariable', 'WarningVariable', 'InformationVariable', 'OutVariable', 'OutBuffer', 'PipelineVariable', 'EA', 'WA', 'IA', 'EV', 'WV', 'IV', 'OV', 'OB', 'PV')
    $all = $valueParams + $switchParams
    $nameArg = $null
    $scopeArg = $null
    $positional = New-Object System.Collections.Generic.List[object]
    $elements = $commandAst.CommandElements
    for ($i = 1; $i -lt $elements.Count; $i++) {
        $element = $elements[$i]
        if ($element -isnot [System.Management.Automation.Language.CommandParameterAst]) {
            $positional.Add($element)
            continue
        }
        $given = $element.ParameterName
        $match = @($all | Where-Object { $_ -eq $given })
        if ($match.Count -eq 0) { $match = @($all | Where-Object { $_.StartsWith($given, [System.StringComparison]::OrdinalIgnoreCase) }) }
        if ($match.Count -ne 1 -or $switchParams -contains $match[0]) { continue }
        $arg = $element.Argument
        if (-not $arg -and $i + 1 -lt $elements.Count) {
            $i++
            $arg = $elements[$i]
        }
        if ($match[0] -eq 'Name') { $nameArg = $arg } elseif ($match[0] -eq 'Scope') { $scopeArg = $arg }
    }
    if (-not $nameArg -and $positional.Count -gt 0) { $nameArg = $positional[0] }
    return [pscustomobject]@{
        Name       = (getConstantText $nameArg)
        Scope      = (getConstantText $scopeArg)
        ScopeGiven = [bool]$scopeArg
    }
}

# Set-Variable・New-Variable が、そのファイルの最上位の変数を作るか。= の代入と同じ条件
# （関数の外なら作る。関数の中なら、-Scope が Script・Global・1 以上のときだけ作る）。
# 結果は [pscustomobject]@{ Kind; Name }。Kind は、名前や -Scope が定数でなく決まらないとき Unknown、
# 最上位の変数を作らないとき None、作るとき Top（Name が変数の名前）
function getCommandTopVariable {
    param ($commandAst)

    $info = getVariableCommandInfo $commandAst 'Set'
    if ($null -eq $info.Name -or ($info.ScopeGiven -and $null -eq $info.Scope)) { return [pscustomobject]@{ Kind = 'Unknown'; Name = $null } }
    $name = $info.Name
    $scope = $info.Scope
    $split = splitVariableScope $name
    if ($split.Scope) {
        $scope = $split.Scope
        $name = $split.Name
    }
    if (-not (testInsideFunction $commandAst)) { return [pscustomobject]@{ Kind = 'Top'; Name = $name } }
    $number = 0
    $isTop = ($scope -in @('script', 'global')) -or ([int]::TryParse([string]$scope, [ref]$number) -and $number -ge 1)
    if ($isTop) { return [pscustomobject]@{ Kind = 'Top'; Name = $name } }
    return [pscustomobject]@{ Kind = 'None'; Name = $null }
}

# 型の名前に出る名前（配列の要素・ジェネリックの引数を含む）を返す。名前空間付きの名前は、そのまま 1 つの名前
function getTypeNameLeaves {
    param ($typeName)

    if ($typeName -is [System.Management.Automation.Language.ArrayTypeName]) { return @(getTypeNameLeaves $typeName.ElementType) }
    if ($typeName -is [System.Management.Automation.Language.GenericTypeName]) {
        $result = @(getTypeNameLeaves $typeName.TypeName)
        foreach ($argument in $typeName.GenericArguments) { $result += @(getTypeNameLeaves $argument) }
        return $result
    }
    return @($typeName.Name)
}

# 構文木から呼ぶ先が決まらない書き方の「形」を返す（findStringNameCalls が数える形）。当てはまらなければ $null
function getStringCallForm {
    param ($node)

    if ($node -is [System.Management.Automation.Language.InvokeMemberExpressionAst]) {
        $member = if ($node.Member -is [System.Management.Automation.Language.StringConstantExpressionAst]) { [string]$node.Member.Value } else { '' }
        if ($node.Static -and $node.Expression -is [System.Management.Automation.Language.TypeExpressionAst] -and
            $node.Expression.TypeName.Name -match '(^|\.)scriptblock$' -and $member -eq 'Create') { return '[scriptblock]::Create' }
        if ($node.Expression.Extent.Text -match '^\$ExecutionContext\.InvokeCommand$') { return '$ExecutionContext.InvokeCommand' }
        if ($member -eq 'AddCommand') { return '.AddCommand' }
        if ($member -eq 'AddScript') { return '.AddScript' }
        return $null
    }
    $name = $node.GetCommandName()
    if ($name) {
        $name = getBareCommandName $name
        if ($name -match '^(Get-Command|gcm)$') { return 'Get-Command' }
        if ($name -match '^(Invoke-Expression|iex)$') { return 'Invoke-Expression' }
        if ($name -match '^(Set-Alias|New-Alias|sal|nal)$') { return 'Set-Alias・New-Alias' }
        return $null
    }
    if ($node.InvocationOperator -eq [System.Management.Automation.Language.TokenKind]::Ampersand) {
        $first = $node.CommandElements[0]
        if ($first -is [System.Management.Automation.Language.ExpandableStringExpressionAst] -or
            $first -is [System.Management.Automation.Language.ParenExpressionAst]) { return '& (式・展開文字列)' }
    }
    return $null
}

# $variable:名前・${variable:\名前} などの variable: ドライブの変数のパス（UserPath）から、名前（先頭の区切りは外す）を返す。
# variable: でなければ $null
function getDriveVariableName {
    param ([string]$userPath)

    if ($userPath -match '^variable:[\\/]?(.+)$') { return $Matches[1] }
    return $null
}

# コマンドの 1 つの要素の、パスとして読む文字列を返す。-Path:variable:名前 のように : でつないだ引数は値の部分。
# 値の無い引数の名前（次の要素が値）は $null
function getPathElementText {
    param ($element)

    if ($element -is [System.Management.Automation.Language.CommandParameterAst]) {
        if (-not $element.Argument) { return $null }
        return $element.Argument.Extent.Text
    }
    return $element.Extent.Text
}

# ドライブ（function:・variable:）を引くコマンドの種類。読む（Read）・書く（Write）・どちらでもない（$null）
function getProviderCommandKind {
    param ([string]$name)

    if ($name -match '^(Get-Item|gi|Get-Content|gc|cat|type|Test-Path|Get-ChildItem|gci|ls|dir)$') { return 'Read' }
    if ($name -match '^(Set-Item|si|New-Item|ni|Remove-Item|ri|rm|del|erase|Copy-Item|cpi|cp|copy)$') { return 'Write' }
    return $null
}

# Set-Variable・New-Variable の -Name が定数なら、最上位の変数の作成として数える。定数でなければ形として数える
function addSetVariableFacts {
    param ($node, $facts, [int]$line)

    $top = getCommandTopVariable $node
    if ($top.Kind -eq 'Unknown') {
        $facts.StringCalls.Add([pscustomobject]@{ Form = 'Set-Variable・New-Variable の -Name が定数でない'; Line = $line })
    } elseif ($top.Kind -eq 'Top') {
        $facts.TopVars.Add($top.Name)
    }
}

# Get-Variable の名前が定数なら、その変数の読みとして数える。定数でなければ形として数える
function addGetVariableFacts {
    param ($node, $facts, [int]$line)

    $info = getVariableCommandInfo $node 'Get'
    if ($null -eq $info.Name -or (testWildcardName $info.Name)) {
        $facts.StringCalls.Add([pscustomobject]@{ Form = 'Get-Variable の名前が定数でない'; Line = $line })
    } else {
        $facts.VarUses.Add([pscustomobject]@{ Node = $node; Name = (removeVariableScope $info.Name); Explicit = $true; Line = $line })
    }
}

# 型の文字列から、クラスの名前になりうる葉を取り出す。[ ] , で区切る（'System.Collections.Generic.List[Workspace]'・'Workspace[]'）
function getTypeTextLeaves {
    param ([string]$typeText)

    return @(($typeText -split '[\[\],]') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

# New-Object の型の引数（-TypeName は位置によらず、無ければ最初の位置引数）を返す。-ComObject など型でない呼び出しは $null
function getNewObjectTypeArgument {
    param ($commandAst)

    $elements = $commandAst.CommandElements
    $positional = $null
    for ($i = 1; $i -lt $elements.Count; $i++) {
        $element = $elements[$i]
        if ($element -isnot [System.Management.Automation.Language.CommandParameterAst]) {
            if (-not $positional) { $positional = $element }
            continue
        }
        $given = $element.ParameterName
        $arg = $element.Argument
        if ($given -and 'TypeName'.StartsWith($given, [System.StringComparison]::OrdinalIgnoreCase)) {
            if (-not $arg -and $i + 1 -lt $elements.Count) { $arg = $elements[$i + 1] }
            return $arg
        }
        # 値を取る引数（-ArgumentList など）の値は、型の位置引数に数えない。-Strict などの値を取らない引数は、次の要素を読み飛ばさない
        $isSwitch = $given -and (@('Strict', 'Verbose', 'Debug') | Where-Object { $_.StartsWith($given, [System.StringComparison]::OrdinalIgnoreCase) })
        if (-not $arg -and -not $isSwitch) { $i++ }
    }
    return $positional
}

# New-Object の型が定数なら、葉（クラスの名前）を型として使う依存に数える（System.* などクラスの定義が無い名前は findFileDependencyItems が数えない）。
# 型を変数・式で渡す形は、形として数える
function addNewObjectFacts {
    param ($node, $facts, [int]$line)

    $typeArg = getNewObjectTypeArgument $node
    if (-not $typeArg) { return }
    $typeText = getConstantText $typeArg
    if ($null -eq $typeText) {
        $facts.StringCalls.Add([pscustomobject]@{ Form = 'New-Object の型が定数でない'; Line = $line })
        return
    }
    foreach ($leaf in @(getTypeTextLeaves $typeText)) {
        $facts.TypeRefs.Add([pscustomobject]@{ Name = $leaf; Line = $line })
    }
}

# ドライブ（function:・variable:）を引くコマンドから、依存の材料を足す。
# function: は、読む・書くのどちらも形として数える（1 回の呼び出しに 1 つ）。
# variable: は、読むときは名前が定数ならその変数の読み、定数でなければ形。書くときは形として数える
function addProviderDriveFacts {
    param ($node, $facts, [string]$name, [int]$line)

    $kind = getProviderCommandKind $name
    if (-not $kind) { return }
    $hasFunctionDrive = $false
    foreach ($element in $node.CommandElements) {
        $text = getPathElementText $element
        if ($null -eq $text) { continue }
        if ($text -match '^[''"]?function:') {
            $hasFunctionDrive = $true
        } elseif ($text -match '^[''"]?variable:') {
            if ($kind -eq 'Write') {
                $facts.StringCalls.Add([pscustomobject]@{ Form = 'variable: ドライブへの書き込み'; Line = $line })
            } elseif ($text -match '^[''"]?variable:[\\/]?([^''"`$*?\[\]\s\\/]+)[''"]?$') {
                # variable:\名前・variable:/名前 のように、ドライブの印のあとに区切りが付く書き方も、同じ名前として読む
                $facts.VarUses.Add([pscustomobject]@{ Node = $node; Name = (removeVariableScope $Matches[1]); Explicit = $true; Line = $line })
            } else {
                $facts.StringCalls.Add([pscustomobject]@{ Form = 'variable: ドライブの名前が定数でない'; Line = $line })
            }
        }
    }
    if ($hasFunctionDrive) { $facts.StringCalls.Add([pscustomobject]@{ Form = 'function: ドライブ'; Line = $line }) }
}

# コマンドの呼び出し 1 つから、依存の材料（関数の呼び出し・変数・型・名前を渡す呼び出し）を $facts に足す
function addCommandFacts {
    param ($node, $facts)

    $name = $node.GetCommandName()
    if ($name) { $name = getBareCommandName $name }
    $line = $node.Extent.StartLineNumber
    if ($name) { $facts.Calls.Add([pscustomobject]@{ Name = $name; Line = $line }) }
    if ($name -match '^(Set-Variable|New-Variable|sv|nv|set)$') {
        addSetVariableFacts $node $facts $line
    } elseif ($name -match '^(Get-Variable|gv)$') {
        addGetVariableFacts $node $facts $line
    } elseif ($name -match '^New-Object$') {
        addNewObjectFacts $node $facts $line
    } else {
        addProviderDriveFacts $node $facts $name $line
    }
}

# 1 つの構文木のノードから、依存の材料（関数の定義・呼び出し・変数・型・クラス・名前を渡す呼び出し）を $facts に足す
function addNodeFacts {
    param ($node, $facts, $defs, [string]$rel)

    if ($node -is [System.Management.Automation.Language.FunctionDefinitionAst]) {
        # クラスのメソッドは、関数の呼び出しでは呼べないので、関数の定義に数えない
        if ($node.Parent -is [System.Management.Automation.Language.FunctionMemberAst]) { return }
        if (-not $defs.ContainsKey($node.Name)) { $defs[$node.Name] = New-Object System.Collections.Generic.List[string] }
        if (-not $defs[$node.Name].Contains($rel)) { $defs[$node.Name].Add($rel) }
    } elseif ($node -is [System.Management.Automation.Language.CommandAst]) {
        addCommandFacts $node $facts
        $form = getStringCallForm $node
        if ($form) { $facts.StringCalls.Add([pscustomobject]@{ Form = $form; Line = $node.Extent.StartLineNumber }) }
    } elseif ($node -is [System.Management.Automation.Language.InvokeMemberExpressionAst]) {
        $form = getStringCallForm $node
        if ($form) { $facts.StringCalls.Add([pscustomobject]@{ Form = $form; Line = $node.Extent.StartLineNumber }) }
    } elseif ($node -is [System.Management.Automation.Language.VariableExpressionAst]) {
        # variable: は、構文木では IsDriveQualified にならず、名前に variable: が付いたまま出る
        if ($node.VariablePath.IsDriveQualified -or $node.VariablePath.UserPath -match '^variable:') {
            # ${function:名前} は関数の呼び出し、${variable:名前}・$variable:名前 は変数の読みとして数える。
            # $env:Path のようなほかのドライブの変数は、スクリプトの変数ではないので数えない
            if ($node.VariablePath.UserPath -match '^function:(.+)$') {
                $facts.Calls.Add([pscustomobject]@{ Name = $Matches[1]; Line = $node.Extent.StartLineNumber })
            } elseif ($null -ne (getDriveVariableName $node.VariablePath.UserPath)) {
                $facts.VarUses.Add([pscustomobject]@{ Node = $node; Name = (removeVariableScope (getDriveVariableName $node.VariablePath.UserPath)); Explicit = $true; Line = $node.Extent.StartLineNumber })
            }
            return
        }
        $facts.VarUses.Add([pscustomobject]@{
                Node     = $node
                Name     = (getVariableBaseName $node.VariablePath)
                Explicit = ($node.VariablePath.IsScript -or $node.VariablePath.IsGlobal)
                Line     = $node.Extent.StartLineNumber
            })
    } elseif ($node -is [System.Management.Automation.Language.AssignmentStatementAst]) {
        $inFunction = testInsideFunction $node
        foreach ($path in @(getAssignedVariablePaths $node.Left)) {
            $driveName = getDriveVariableName $path.UserPath
            if ($null -ne $driveName) {
                # $variable:名前 = … は、その名前の変数への代入として数える（名前の先頭に script: などが付けば、そのスコープ）
                $split = splitVariableScope $driveName
                if (-not $inFunction -or $split.Scope -in @('script', 'global')) { $facts.TopVars.Add($split.Name) }
            } elseif (-not $inFunction -or $path.IsScript -or $path.IsGlobal) {
                $facts.TopVars.Add((getVariableBaseName $path))
            }
        }
    } elseif ($node -is [System.Management.Automation.Language.TypeDefinitionAst]) {
        if ($node.IsClass) { $facts.Classes.Add($node.Name) }
    } else {
        $facts.Types.Add($node)
        foreach ($leaf in @(getTypeNameLeaves $node.TypeName)) {
            $facts.TypeRefs.Add([pscustomobject]@{ Name = $leaf; Line = $node.Extent.StartLineNumber })
        }
    }
}

# root の下の .ps1 を全部解析して、依存を調べるための材料を作る。解析は重いので、1 回だけ作って使い回す。
# Files の各要素: Path（root からの相対。区切りは /）・Calls（コマンドの名前と行。${function:名前} を含む）・VarUses（変数の式）・
# Types（型の式）・TypeRefs（型として使った名前と行）・Classes（定義したクラスの名前）・
# StringCalls（呼ぶ先が決まらない書き方の形と行）・TopVars（最上位・$script:・$global: で代入した変数の名前。Set-Variable・New-Variable を含む）。
# Defs: 関数の名前 → 定義したファイルの Path。ClassDefs: クラスの名前 → 定義したファイルの Path
function newScriptArch {
    param (
        [string]$root
    )

    $rootFull = (Resolve-Path -LiteralPath $root).Path.TrimEnd('\', '/')
    $files = New-Object System.Collections.Generic.List[object]
    $defs = @{}
    $classDefs = @{}
    $interesting = {
        param ($n)
        ($n -is [System.Management.Automation.Language.FunctionDefinitionAst]) -or
        ($n -is [System.Management.Automation.Language.CommandAst]) -or
        ($n -is [System.Management.Automation.Language.VariableExpressionAst]) -or
        ($n -is [System.Management.Automation.Language.AssignmentStatementAst]) -or
        ($n -is [System.Management.Automation.Language.InvokeMemberExpressionAst]) -or
        ($n -is [System.Management.Automation.Language.TypeDefinitionAst]) -or
        ($n -is [System.Management.Automation.Language.TypeExpressionAst]) -or
        ($n -is [System.Management.Automation.Language.TypeConstraintAst])
    }

    foreach ($file in (Get-ChildItem -LiteralPath $rootFull -Recurse -Filter "*.ps1" | Sort-Object FullName)) {
        $rel = $file.FullName.Substring($rootFull.Length + 1) -replace '\\', '/'
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$errors)
        # 別のファイルで定義した型を使うクラス（TypeNotFound）は、1 ファイルだけでは解決できないので、構文エラーに数えない
        $syntaxErrors = @($errors | Where-Object { $_.ErrorId -ne 'TypeNotFound' })
        if ($syntaxErrors.Count -gt 0) { throw "構文エラーがあり、依存を読めない: ${rel}:$($syntaxErrors[0].Extent.StartLineNumber) $($syntaxErrors[0].Message)" }

        $facts = [pscustomobject]@{
            Calls       = New-Object System.Collections.Generic.List[object]
            VarUses     = New-Object System.Collections.Generic.List[object]
            Types       = New-Object System.Collections.Generic.List[object]
            TypeRefs    = New-Object System.Collections.Generic.List[object]
            Classes     = New-Object System.Collections.Generic.List[string]
            StringCalls = New-Object System.Collections.Generic.List[object]
            TopVars     = New-Object System.Collections.Generic.List[string]
        }
        foreach ($node in $ast.FindAll($interesting, $true)) { addNodeFacts $node $facts $defs $rel }
        foreach ($class in $facts.Classes) {
            if (-not $classDefs.ContainsKey($class)) { $classDefs[$class] = New-Object System.Collections.Generic.List[string] }
            $classDefs[$class].Add($rel)
        }
        $files.Add([pscustomobject]@{
                Path        = $rel
                Calls       = $facts.Calls.ToArray()
                VarUses     = $facts.VarUses.ToArray()
                Types       = $facts.Types.ToArray()
                TypeRefs    = $facts.TypeRefs.ToArray()
                Classes     = $facts.Classes.ToArray()
                StringCalls = $facts.StringCalls.ToArray()
                TopVars     = @($facts.TopVars | Sort-Object -Unique)
            })
    }
    # OwnedCache は、変数の判定（getOwnedVariableNames）が行をまたいで使い回す入れ物
    return [pscustomobject]@{ Root = $rootFull; Files = $files.ToArray(); Defs = $defs; ClassDefs = $classDefs; OwnedCache = @{} }
}

# -From のパスのファイルが、-To のパスのファイル（-ToExcept を除く）に依存している所を、
# [pscustomobject]@{ From; To; Name; Line } の一覧で返す。Name は関数の名前か、$ を付けた変数の名前か、[ ] で囲んだクラスの名前。
# $arch には newScriptArch の結果か、root のパスを渡す（パスなら、ここで解析する）
function findFileDependencyItems {
    param (
        $arch,
        [string[]]$From,
        [string[]]$To,
        [string[]]$ToExcept = @()
    )

    if ($arch -is [string]) { $arch = newScriptArch $arch }
    $fromFiles = @($arch.Files | Where-Object { testPathLike $_.Path $From })
    $fromSet = @{}
    foreach ($file in $fromFiles) { $fromSet[$file.Path] = $true }
    $toSet = @{}
    $toVars = @{}
    foreach ($file in $arch.Files) {
        if (-not (testPathLike $file.Path $To) -or (testPathLike $file.Path $ToExcept)) { continue }
        $toSet[$file.Path] = $true
        foreach ($name in $file.TopVars) {
            if (-not $toVars.ContainsKey($name)) { $toVars[$name] = New-Object System.Collections.Generic.List[string] }
            $toVars[$name].Add($file.Path)
        }
    }
    $fromVars = @{}
    foreach ($file in $fromFiles) {
        foreach ($name in $file.TopVars) { $fromVars[$name] = $true }
    }

    $items = New-Object System.Collections.Generic.List[object]
    $ownedCache = $arch.OwnedCache
    foreach ($file in $fromFiles) {
        # 関数の呼び出し。自分のファイルや、見る側のパスにも同じ名前の関数があるなら、そちらを呼んだとみなす
        foreach ($call in $file.Calls) {
            if (-not $arch.Defs.ContainsKey($call.Name)) { continue }
            $where = @($arch.Defs[$call.Name])
            if ($where -contains $file.Path) { continue }
            if (@($where | Where-Object { $fromSet.ContainsKey($_) -and -not $toSet.ContainsKey($_) }).Count -gt 0) { continue }
            foreach ($def in $where) {
                if ($toSet.ContainsKey($def)) {
                    $items.Add([pscustomobject]@{ From = $file.Path; To = $def; Name = $call.Name; Line = $call.Line })
                }
            }
        }
        # 変数の読み。見る側にも最上位の同じ名前があるか、関数・スクリプトブロックの引数・代入・foreach の変数なら、自分のものとみなす
        foreach ($use in $file.VarUses) {
            $name = $use.Name
            if (-not $toVars.ContainsKey($name) -or $fromVars.ContainsKey($name)) { continue }
            if (-not $use.Explicit -and (testLocalVariable $use.Node $name $ownedCache)) { continue }
            foreach ($def in $toVars[$name]) {
                if ($def -ne $file.Path) {
                    $items.Add([pscustomobject]@{ From = $file.Path; To = $def; Name = '$' + $name; Line = $use.Line })
                }
            }
        }
        # 別のファイルで定義したクラスを型として使う所（名前は [型名]）。関数と同じく、自分のファイル・見る側のパスにも同じ名前のクラスがあれば、そちらを使ったとみなす
        foreach ($ref in $file.TypeRefs) {
            if (-not $arch.ClassDefs.ContainsKey($ref.Name)) { continue }
            $where = @($arch.ClassDefs[$ref.Name])
            if ($where -contains $file.Path) { continue }
            if (@($where | Where-Object { $fromSet.ContainsKey($_) -and -not $toSet.ContainsKey($_) }).Count -gt 0) { continue }
            foreach ($def in $where) {
                if ($toSet.ContainsKey($def)) {
                    $items.Add([pscustomobject]@{ From = $file.Path; To = $def; Name = "[$($ref.Name)]"; Line = $ref.Line })
                }
            }
        }
    }
    return $items.ToArray()
}

# 呼ぶ先が構文木から決まらない書き方（getStringCallForm の形）を、-In のパスのファイルから探し、
# [pscustomobject]@{ From; To; Name; Line } の一覧で返す。To は決まった文字列、Name は形。
# removeAllowedDependencies・findUnusedAllowances に、同じ形の例外の一覧を渡して使える
function findStringNameCalls {
    param (
        $arch,
        [string[]]$In = @('*')
    )

    if ($arch -is [string]) { $arch = newScriptArch $arch }
    $items = New-Object System.Collections.Generic.List[object]
    foreach ($file in ($arch.Files | Where-Object { testPathLike $_.Path $In })) {
        foreach ($call in $file.StringCalls) {
            $items.Add([pscustomobject]@{ From = $file.Path; To = '(名前の文字列)'; Name = $call.Form; Line = $call.Line })
        }
    }
    return $items.ToArray()
}

# findFileDependencyItems の結果を、"<from のパス> -> <to のパス>（<名前>・<名前>）" の一覧（ファイルの組ごと）にする
function formatFileDependencies {
    param ($items)

    $pairs = @{}
    foreach ($item in @($items)) {
        $key = "$($item.From) -> $($item.To)"
        if (-not $pairs.ContainsKey($key)) { $pairs[$key] = New-Object System.Collections.Generic.List[string] }
        if (-not $pairs[$key].Contains($item.Name)) { $pairs[$key].Add($item.Name) }
    }
    return @($pairs.Keys | Sort-Object | ForEach-Object { "$_（$((@($pairs[$_] | Sort-Object)) -join '・')）" })
}

# ファイルの依存を、"<from のパス> -> <to のパス>（<名前>・<名前>）" の一覧で返す（findFileDependencyItems の結果を整える）
function findFileDependencies {
    param (
        $arch,
        [string[]]$From,
        [string[]]$To,
        [string[]]$ToExcept = @()
    )

    return @(formatFileDependencies @(findFileDependencyItems $arch -From $From -To $To -ToExcept $ToExcept))
}

# 例外（From・To・Name の組）として許した依存を、items から外して返す
function removeAllowedDependencies {
    param (
        $items,
        $allowances
    )

    $allowed = @{}
    foreach ($allowance in @($allowances)) { $allowed["$($allowance.From)|$($allowance.To)|$($allowance.Name)"] = $true }
    return @(@($items) | Where-Object { -not $allowed.ContainsKey("$($_.From)|$($_.To)|$($_.Name)") })
}

# 例外のうち、items に出てこないもの（使われなくなった例外）を返す。
# items には、全部の行の結果をまとめたものを渡す（行ごとに突き合わせると、別の行で出る依存の例外を使われないとしてしまう）
function findUnusedAllowances {
    param (
        $allowances,
        $items
    )

    $found = @{}
    foreach ($item in @($items)) { $found["$($item.From)|$($item.To)|$($item.Name)"] = $true }
    return @(@($allowances) | Where-Object { -not $found.ContainsKey("$($_.From)|$($_.To)|$($_.Name)") })
}

# -In のパスのファイルが、決めた変数（-Variable。$ なしの名前）と型（-TypeLike。-like のワイルドカード）に触っている所を、
# "<パス>:<行> $<変数>" か "<パス>:<行> [<型>]" の一覧で返す。コメント・文字列の中の名前は構文木に変数・型としては出ないので、数えない
function findTypeUses {
    param (
        $arch,
        [string[]]$In,
        [string[]]$Variable = @(),
        [string[]]$TypeLike = @()
    )

    if ($arch -is [string]) { $arch = newScriptArch $arch }
    $found = New-Object System.Collections.Generic.List[string]
    foreach ($file in ($arch.Files | Where-Object { testPathLike $_.Path $In })) {
        foreach ($use in $file.VarUses) {
            if ($Variable -contains $use.Name) { $found.Add("$($file.Path):$($use.Line) `$$($use.Name)") }
        }
        foreach ($node in $file.Types) {
            $type = $node.TypeName.FullName
            if (testPathLike $type $TypeLike) { $found.Add("$($file.Path):$($node.Extent.StartLineNumber) [$type]") }
        }
    }
    return $found.ToArray()
}
