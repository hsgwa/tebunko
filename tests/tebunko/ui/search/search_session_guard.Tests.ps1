# 検索の開始（tebunko\ui\search\search_session.ps1 の startSearch）の、移動中の守りのテスト。
# 関数だけをファイルから取り出し、画面の部品は呼ばれた操作を残す偽物にする（ファイルは画面を作る処理を含むため、そのままは読み込まない）。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\index\index_edit_view.ps1"

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile("${scriptsDir}\tebunko\ui\search\search_session.ps1", [ref]$tokens, [ref]$errors)
    $definition = $ast.Find({ param ($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq "startSearch" }, $true)
    . ([scriptblock]::Create($definition.Extent.Text))

    function getWordText { $global:searchLog.Add("getWordText"); return "" }
    function setStatus { param ($text) $global:searchLog.Add("setStatus") }
    function cancelSearch { $global:searchLog.Add("cancelSearch") }
}

Describe "startSearch" -Tag Unit {
    BeforeEach {
        $global:searchLog = New-Object System.Collections.Generic.List[string]
        $script:search = $null
        $script:workspaceMoving = $false
    }
    AfterAll {
        Remove-Variable -Name searchLog -Scope Global -ErrorAction SilentlyContinue
    }

    It "ワークスペースを移している間は、何も始めず、検索中の中止もしない" {
        $script:workspaceMoving = $true
        startSearch
        @($global:searchLog).Count | Should -Be 0
        $script:search = @{ Shared = $null }
        startSearch
        @($global:searchLog).Count | Should -Be 0
    }

    It "移していなければ、ワードを読む（守りが常に効いているのではない）" {
        startSearch
        @($global:searchLog) | Should -Be @("getWordText", "setStatus")
    }

    It "移していなくて検索中なら、中止する" {
        $script:search = @{ Shared = $null }
        startSearch
        @($global:searchLog) | Should -Be @("cancelSearch")
    }
}