# スクリプトの層の決まり（依存の向き）を、AST で関数の呼び出し・変数・型の使い方まで読んで確かめる部品。
# tests/meta/layers.Tests.ps1 が読む。tebunko のフォルダ構成は知らない（どのパスがどのパスに依存してよいかは、呼び出し側が表で渡す）。
#
# 全部のファイルが同じスコープに読み込まれるので、下の層が上の層の関数・変数を名前で使っても動いてしまう。
# 文字列を探す検査では、関数や変数の使われ方まで見えない。ここでは、ファイルごとに構文木を作り、次を見る。
#
# - 関数の呼び出し（コマンドの名前）が、別のパスのファイルで定義した関数を指していること
# - 別のパスのファイルが最上位（関数の外）・$script:・$global: で代入した変数を、読んでいること
# - 決めた変数（$ui・$window など）と型（System.Windows.* など）に触っていること
#
# 見つからないもの（限界）: & $name の呼び出し、コールバックで渡したスクリプトブロックの中から呼ぶ形、
# $ui を別の名前の引数で受け取る形、クラスのメソッドの呼び出し、Set-Variable での代入。
# 同じ名前の関数・変数が、見る側のパスにもあるときは、そちらを使っているとみなして返さない。
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

    return ($variablePath.UserPath -replace '^(script|global|local|private):', '')
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

# root の下の .ps1 を全部解析して、依存を調べるための材料を作る。解析は重いので、1 回だけ作って使い回す。
# Files の各要素: Path（root からの相対。区切りは /）・Calls（コマンドの名前と行）・VarUses（変数の式）・
# Types（型の式）・TopVars（最上位・$script:・$global: で代入した変数の名前）。Defs: 関数の名前 → 定義したファイルの Path
function newScriptArch {
    param (
        [string]$root
    )

    $rootFull = (Resolve-Path -LiteralPath $root).Path.TrimEnd('\', '/')
    $files = New-Object System.Collections.Generic.List[object]
    $defs = @{}
    $interesting = {
        param ($n)
        ($n -is [System.Management.Automation.Language.FunctionDefinitionAst]) -or
        ($n -is [System.Management.Automation.Language.CommandAst]) -or
        ($n -is [System.Management.Automation.Language.VariableExpressionAst]) -or
        ($n -is [System.Management.Automation.Language.AssignmentStatementAst]) -or
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

        $calls = New-Object System.Collections.Generic.List[object]
        $varUses = New-Object System.Collections.Generic.List[object]
        $types = New-Object System.Collections.Generic.List[object]
        $topVars = New-Object System.Collections.Generic.List[string]
        foreach ($node in $ast.FindAll($interesting, $true)) {
            if ($node -is [System.Management.Automation.Language.FunctionDefinitionAst]) {
                # クラスのメソッドは、関数の呼び出しでは呼べないので、関数の定義に数えない
                if ($node.Parent -is [System.Management.Automation.Language.FunctionMemberAst]) { continue }
                if (-not $defs.ContainsKey($node.Name)) { $defs[$node.Name] = New-Object System.Collections.Generic.List[string] }
                if (-not $defs[$node.Name].Contains($rel)) { $defs[$node.Name].Add($rel) }
            } elseif ($node -is [System.Management.Automation.Language.CommandAst]) {
                $name = $node.GetCommandName()
                if ($name) { $calls.Add([pscustomobject]@{ Name = $name; Line = $node.Extent.StartLineNumber }) }
            } elseif ($node -is [System.Management.Automation.Language.VariableExpressionAst]) {
                # $env:Path のようなドライブの変数は、スクリプトの変数ではないので数えない
                if ($node.VariablePath.IsDriveQualified) { continue }
                $varUses.Add([pscustomobject]@{
                        Node     = $node
                        Name     = (getVariableBaseName $node.VariablePath)
                        Explicit = ($node.VariablePath.IsScript -or $node.VariablePath.IsGlobal)
                        Line     = $node.Extent.StartLineNumber
                    })
            } elseif ($node -is [System.Management.Automation.Language.AssignmentStatementAst]) {
                $inFunction = testInsideFunction $node
                foreach ($path in @(getAssignedVariablePaths $node.Left)) {
                    if (-not $inFunction -or $path.IsScript -or $path.IsGlobal) { $topVars.Add((getVariableBaseName $path)) }
                }
            } else {
                $types.Add($node)
            }
        }
        $files.Add([pscustomobject]@{
                Path    = $rel
                Calls   = $calls.ToArray()
                VarUses = $varUses.ToArray()
                Types   = $types.ToArray()
                TopVars = @($topVars | Sort-Object -Unique)
            })
    }
    # OwnedCache は、変数の判定（getOwnedVariableNames）が行をまたいで使い回す入れ物
    return [pscustomobject]@{ Root = $rootFull; Files = $files.ToArray(); Defs = $defs; OwnedCache = @{} }
}

# -From のパスのファイルが、-To のパスのファイル（-ToExcept を除く）に依存している所を、
# [pscustomobject]@{ From; To; Name; Line } の一覧で返す。Name は関数の名前か、$ を付けた変数の名前。
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
