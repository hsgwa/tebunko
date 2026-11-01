# GetNewClosure() したスクリプトブロックの中では、$script: の変数は、ツールの変数ではなく、閉じ込めたブロックの側の空の変数を指す
# （値はいつも空になる。関数を名前で探せないのと同じ理由。docs/design/gui/responsiveness.md）。
# タイマーやイベントのブロックが $script:<名前> を読むと、更新の流れが黙って止まる不具合になる（#268 の始め直し）。
# 閉じ込めの外のブロック（$isNaming = { $script:x } のように、GetNewClosure() を付けないもの）で読み、ブロックからはそれを & で呼ぶ。
# 構文の木で、.GetNewClosure() を呼ぶスクリプトブロックの中の $script: の変数を探す。許可の一覧（今ある使い方。足さない）に無ければ落とす。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"

    function getClosureScriptScopeUses {
        # scripts\ の .ps1 から、GetNewClosure() を付けたスクリプトブロックの中の $script: の変数の使い方を、"<ファイル名>:<変数名>" の並びで返す
        $uses = New-Object System.Collections.Generic.List[string]
        foreach ($file in Get-ChildItem -LiteralPath ${scriptsDir} -Recurse -Filter *.ps1) {
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$null)
            $calls = $ast.FindAll({ param ($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and "$($n.Member.Value)" -eq "GetNewClosure" -and $n.Expression -is [System.Management.Automation.Language.ScriptBlockExpressionAst] }, $true)
            foreach ($call in $calls) {
                $variables = $call.Expression.FindAll({ param ($n) $n -is [System.Management.Automation.Language.VariableExpressionAst] -and $n.VariablePath.IsScript }, $true)
                foreach ($variable in $variables) {
                    $uses.Add("$($file.Name):$($variable.VariablePath.UserPath)")
                }
            }
        }
        return @($uses | Sort-Object -Unique)
    }
}

Describe "GetNewClosure() したブロックの中の script スコープの変数" -Tag Meta {
    It "許可の一覧に無い使い方が無い（閉じ込めの外のブロックで読み、& で呼ぶ）" {
        ${allowed} = @(
        )
        $unknown = @(getClosureScriptScopeUses | Where-Object { ${allowed} -notcontains $_ })
        $unknown | Should -BeNullOrEmpty -Because "GetNewClosure() したブロックの中の script スコープの変数は空になる。閉じ込めの外のブロックで読み、& で呼ぶ"
    }
}
