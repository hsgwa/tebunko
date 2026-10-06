# インデックス管理の詳細が、選んだ行の値の変わるたびに書き直されることのテスト。
# 詳細（updateIndexDetailPanel）は、選んだ行（FolderItem）の値を読んで出す。行の値を変える所が詳細を
# 書き直さないと、一覧と詳細で値が食い違う（行のチェックを切り替えても詳細が古いままだった不具合）。
# 画面を開かずに、呼び出しの有無を構文木で確かめる（実際に書き直るかは tests/gui/index.Tests.ps1）。
# コメントや文字列には一致させない（コマンドの呼び出しだけを数える）。行の値を変える方法は、
# ui/types.ps1 の FolderItem の Set* メソッドと、プロパティへの直接の代入から集める
BeforeAll {
    $script:uiDir = (Resolve-Path "$PSScriptRoot\..\..\scripts\tebunko\ui").Path
    $script:refreshers = @("updateIndexDetailPanel", "updateIndexListView", "refreshIndexViews", "updateFastSearchRows", "applyIndexStats")

    function getAst {
        param ([string]$path)
        $parseErrors = $null
        return [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$null, [ref]$parseErrors)
    }

    # FolderItem の Set* メソッド名とプロパティ名
    $typesAst = getAst "$script:uiDir\types.ps1"
    $folderItem = $typesAst.Find({ param ($a) $a -is [System.Management.Automation.Language.TypeDefinitionAst] -and $a.Name -eq "FolderItem" }, $true)
    $script:setters = @($folderItem.Members | Where-Object { $_ -is [System.Management.Automation.Language.FunctionMemberAst] -and $_.Name -like "Set*" -and !$_.IsHidden } | ForEach-Object { $_.Name })
    $script:properties = @($folderItem.Members | Where-Object { $_ -is [System.Management.Automation.Language.PropertyMemberAst] } | ForEach-Object { $_.Name })

    # 「ファイル名:名前」→ @{ Commands; Writes }（関数と、関数の外の $名前 = { ... }）
    $script:units = @{}
    foreach ($file in Get-ChildItem $script:uiDir -Recurse -Filter "*.ps1") {
        $ast = getAst $file.FullName
        $inIndexDir = $file.DirectoryName -like "*\ui\index"
        $nodes = @()
        foreach ($fn in $ast.FindAll({ param ($a) $a -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
            $nodes += @{ Key = "$($file.Name):$($fn.Name)"; Ast = $fn.Body }
        }
        foreach ($as in $ast.FindAll({ param ($a) $a -is [System.Management.Automation.Language.AssignmentStatementAst] -and $a.Right.Find({ param ($b) $b -is [System.Management.Automation.Language.ScriptBlockExpressionAst] }, $false) }, $true)) {
            $outer = $false
            for ($p = $as.Parent; $null -ne $p; $p = $p.Parent) {
                if ($p -is [System.Management.Automation.Language.FunctionDefinitionAst]) { $outer = $true; break }
            }
            if (!$outer) {
                $nodes += @{ Key = "$($file.Name):" + ($as.Left.Extent.Text -replace '^\$(script:)?', ''); Ast = $as.Right }
            }
        }
        foreach ($node in $nodes) {
            $commands = @($node.Ast.FindAll({ param ($a) $a -is [System.Management.Automation.Language.CommandAst] }, $true) | ForEach-Object { $_.GetCommandName() } | Where-Object { $_ })
            $writes = @($node.Ast.FindAll({ param ($a) $a -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $a.Member.Value -in $script:setters }, $true) | ForEach-Object { $_.Member.Value })
            if ($inIndexDir) {
                $writes += @($node.Ast.FindAll({ param ($a) $a -is [System.Management.Automation.Language.AssignmentStatementAst] -and $a.Left -is [System.Management.Automation.Language.MemberExpressionAst] -and $a.Left.Member.Value -in $script:properties }, $true) | ForEach-Object { $_.Left.Member.Value })
            }
            # 同じファイルに同じ名前が 2 つあるときは、後ろに番号を付けて両方を見る（上書きして片方を見落とさない）
            $key = $node.Key
            for ($i = 2; $script:units.ContainsKey($key); $i++) {
                $key = "$($node.Key)#$i"
            }
            $script:units[$key] = @{ Commands = $commands; Writes = $writes }
        }
    }
}

Describe "インデックス管理の詳細を書き直す入口" -Tag Meta {
    It "行の値を変える方法を集められている（FolderItem の Set* が 1 つ以上ある）" {
        $script:setters | Should -Contain "SetStatus"
        $script:setters | Should -Contain "SetEnabled"
    }

    It "<key> は updateIndexDetailPanel を呼ぶ（<reason>）" -TestCases @(
        @{ key = "index_list.ps1:onIndexGridToggled"; reason = "行のチェックの切り替え" }
        @{ key = "index_list.ps1:applyFolderStatus"; reason = "フォルダの有無の結果が遅れて届く" }
        @{ key = "index_list.ps1:updateFolderItemStatus"; reason = "場所の変更でフォルダの状態が「確認しています」に戻る" }
        @{ key = "index_list.ps1:updateIndexListView"; reason = "選び直し・追加・削除・読み直し" }
        @{ key = "index_list.ps1:updateFastSearchRows"; reason = "取り込みの集計と高速検索の確かめが届く" }
        @{ key = "index_list.ps1:refreshIndexViews"; reason = "追加・編集・削除のあと" }
    ) {
        param ($key, $reason)
        $script:units.ContainsKey($key) | Should -BeTrue
        $script:units[$key].Commands | Should -Contain "updateIndexDetailPanel"
    }

    It "一覧の行の値を書き換える関数は、書き直す入口のどれかを呼ぶ" {
        # newFolderItem は一覧に加える前の行を作るだけなので、加えたあとに呼ぶ側（updateIndexListView）が書き直す。
        # 書き直す入口そのものも、ここでは調べない
        $missing = @($script:units.Keys | Where-Object {
            $name = ($_ -replace '^[^:]*:', '') -replace '#\d+$', ''
            $unit = $script:units[$_]
            $name -ne "newFolderItem" -and $name -notin $script:refreshers -and $unit.Writes.Count -gt 0 -and
            @($unit.Commands | Where-Object { $_ -in $script:refreshers }).Count -eq 0
        } | Sort-Object)
        ($missing -join ", ") | Should -Be ""
    }
}
