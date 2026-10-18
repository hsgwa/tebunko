# インデックスの名前を決める入口が、設定から外れたまま残してあるインデックス（$script:leftoverIndexes）を渡していることのテスト。
# 外れたインデックスと同じ名前を別のフォルダに付けると、次の更新で前の中身が確認なしに消える。名前の集合を作る getUsedIndexNames と、
# 名前を調べる testIndexEditInput・getIndexDetailEditDecision を呼ぶ画面層の所が、残っている組を渡し忘れると落ちる（入口を足したときの渡し忘れを止める）。
# 画面を開かずに、呼び出しを構文木で確かめる。判断層の中の呼び出し（*_view.ps1。一覧の名前だけを見る内側の呼び出しを含む）は対象外
BeforeAll {
    $script:uiDir = (Resolve-Path "$PSScriptRoot\..\..\scripts\tebunko\ui").Path
    $script:entries = @("getUsedIndexNames", "testIndexEditInput", "getIndexDetailEditDecision")

    # 名前の集合を受け取る関数と、その集合が並ぶ位置（コマンド名を 0 とした引数の位置）
    $script:setReceivers = @{ newIndexName = 2; getImportSuggestedName = 2; getImportCollisionChoiceView = 3 }

    $script:calls = New-Object System.Collections.Generic.List[object]
    $script:setCalls = New-Object System.Collections.Generic.List[object]
    foreach ($file in Get-ChildItem $script:uiDir -Recurse -Filter "*.ps1" | Where-Object { $_.Name -notlike "*_view.ps1" }) {
        $parseErrors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$parseErrors)
        $found = $ast.FindAll({ param ($a) $a -is [System.Management.Automation.Language.CommandAst] -and $a.GetCommandName() -in $script:entries }, $true)
        foreach ($call in $found) {
            $script:calls.Add(@{ Where = "$($file.Name):$($call.Extent.StartLineNumber)"; Name = $call.GetCommandName(); Text = $call.Extent.Text })
        }
        # getUsedIndexNames の結果を受けた変数の名前
        $fromUsed = @($ast.FindAll({ param ($a) $a -is [System.Management.Automation.Language.AssignmentStatementAst] -and $a.Right.Extent.Text -match '^getUsedIndexNames\b' }, $true) |
            ForEach-Object { $_.Left.VariablePath.UserPath })
        $receivers = $ast.FindAll({ param ($a) $a -is [System.Management.Automation.Language.CommandAst] -and $script:setReceivers.ContainsKey([string]$a.GetCommandName()) }, $true)
        foreach ($call in $receivers) {
            $index = $script:setReceivers[[string]$call.GetCommandName()]
            $arg = if ($call.CommandElements.Count -gt $index) { $call.CommandElements[$index] } else { $null }
            $ok = $false
            # getImportCollisionChoiceView は、今の一覧の名前（5 つ目の引数）も要る（外れた名前に［置き換える］を出さないため）
            if ($call.GetCommandName() -eq "getImportCollisionChoiceView" -and $call.CommandElements.Count -lt 6) {
                $script:setCalls.Add(@{ Where = "$($file.Name):$($call.Extent.StartLineNumber)"; Name = "getImportCollisionChoiceView（一覧の名前の引数が無い）"; Ok = $false })
            }
            if ($arg -is [System.Management.Automation.Language.ParenExpressionAst]) {
                $ok = $arg.Extent.Text -match '^\(\s*getUsedIndexNames\b'
            } elseif ($arg -is [System.Management.Automation.Language.VariableExpressionAst]) {
                $ok = $arg.VariablePath.UserPath -in $fromUsed
            }
            $script:setCalls.Add(@{ Where = "$($file.Name):$($call.Extent.StartLineNumber)"; Name = $call.GetCommandName(); Ok = $ok })
        }
    }
}

Describe "インデックスの名前を決める入口" -Tag Meta {
    It "名前を決める入口の呼び出しを集められている" {
        $script:calls.Count | Should -BeGreaterThan 5
    }

    It "画面層の呼び出しは、すべて残っているインデックス（leftoverIndexes）を渡している" {
        $missing = @($script:calls | Where-Object { $_.Text -notmatch 'leftoverIndexes' } | ForEach-Object { "$($_.Where) $($_.Name)" })
        $missing | Should -Be @()
    }

    It "名前の集合を受け取る関数（newIndexName・getImportSuggestedName・getImportCollisionChoiceView）には、getUsedIndexNames の結果を渡している" {
        $script:setCalls.Count | Should -BeGreaterThan 3
        $bad = @($script:setCalls | Where-Object { !$_.Ok } | ForEach-Object { "$($_.Where) $($_.Name)" })
        $bad | Should -Be @()
    }
}