# インデックス管理の詳細が、選んだ行の値の変わるたびに書き直されることのテスト。
# 詳細（updateIndexDetailPanel）は、選んだ行（FolderItem）の値を読んで出す。行の値を変える所が詳細を
# 書き直さないと、一覧と詳細で値が食い違う（［作成］のチェックを切り替えても詳細が古いままだった不具合）。
# 画面を開かずに、呼び出しの有無を構文木で確かめる（実際に書き直るかは tests/gui/index.Tests.ps1）
BeforeAll {
    $script:indexUiDir = (Resolve-Path "$PSScriptRoot\..\..\scripts\tebunko\ui\index").Path
    $script:asts = @{}
    foreach ($file in Get-ChildItem $script:indexUiDir -Filter "*.ps1") {
        $parseErrors = $null
        $script:asts[$file.Name] = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$parseErrors)
    }
    # 関数の本文・名前付きのスクリプトブロック（$name = { ... }）の本文を、名前 → 本文の文字列で集める
    $script:bodies = @{}
    foreach ($name in $script:asts.Keys) {
        $ast = $script:asts[$name]
        foreach ($fn in $ast.FindAll({ param ($a) $a -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
            $script:bodies[$fn.Name] = $fn.Body.Extent.Text
        }
        foreach ($as in $ast.FindAll({ param ($a) $a -is [System.Management.Automation.Language.AssignmentStatementAst] -and $a.Right.Extent.Text.TrimStart().StartsWith("{") }, $true)) {
            $script:bodies[$as.Left.Extent.Text.TrimStart('$')] = $as.Right.Extent.Text
        }
    }
}

Describe "インデックス管理の詳細を書き直す入口" -Tag Meta {
    It "<name> は updateIndexDetailPanel を呼ぶ（<reason>）" -TestCases @(
        @{ name = "onIndexGridToggled"; reason = "［作成］のチェックの切り替え" }
        @{ name = "applyFolderStatus"; reason = "フォルダの有無の結果が遅れて届く" }
        @{ name = "updateFolderItemStatus"; reason = "場所の変更でフォルダの状態が「確認しています」に戻る" }
        @{ name = "updateIndexListView"; reason = "選び直し・追加・削除・読み直し" }
        @{ name = "updateFastSearchRows"; reason = "取り込みの集計と高速検索の確かめが届く" }
        @{ name = "refreshIndexViews"; reason = "追加・編集・削除のあと" }
    ) {
        param ($name, $reason)
        $script:bodies.ContainsKey($name) | Should -BeTrue
        $script:bodies[$name] | Should -Match 'updateIndexDetailPanel'
    }

    It "一覧の行の値を変える関数は、書き直す入口のどれかを呼ぶ" {
        # 行の値（状態・件数・高速検索・名前・場所・［作成］）を書き換える関数。newFolderItem は一覧に加える前の行を作るだけなので、
        # 加えたあとに呼ぶ側（updateIndexListView）が書き直す
        $setters = '\.(SetStatus|SetStats|SetFast|SetIndexState|SetName|SetPath)\(|\.Enabled\s*='
        $refresh = 'updateIndexDetailPanel|updateIndexListView|refreshIndexViews|updateFastSearchRows|applyIndexStats'
        $missing = @($script:bodies.Keys | Where-Object {
            $_ -ne "newFolderItem" -and $script:bodies[$_] -match $setters -and $script:bodies[$_] -notmatch $refresh
        } | Sort-Object)
        ($missing -join ", ") | Should -Be ""
    }
}
