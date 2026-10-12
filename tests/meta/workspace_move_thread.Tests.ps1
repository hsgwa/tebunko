# ワークスペースの中身を移す・消す処理は、画面のスレッドで動かさない（応答なしにしない）。
# 画面のファイル（scripts\tebunko\ui\ の下）で WorkspaceMover の ChangeWorkspace・MoveEntries・RemoveEntries を呼ぶ所は、
# すべて startJob に渡すスクリプトブロックの中にあること。関数の名前に頼らず、構文の木で調べる。
# 1 件でもスクリプトブロックの外にあれば、画面のスレッドで数十秒止まる（大きなワークスペースの移動・別のドライブへの写し）。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"

    ${moverMethods} = @("ChangeWorkspace", "MoveEntries", "RemoveEntries")

    function findUnthreadedMoverCalls {
        # 画面のファイルで、startJob のスクリプトブロックの外から WorkspaceMover の重い呼び出しをしている所を "ファイル:行 メソッド" の配列で返す
        param (
            [string]$directory,
            [string[]]$methods
        )

        $found = New-Object System.Collections.Generic.List[string]
        foreach ($file in @(Get-ChildItem -LiteralPath $directory -Filter "*.ps1" -Recurse)) {
            $errors = $null
            $tokens = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
            $calls = $ast.FindAll({
                    param ($n)
                    $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and
                    $n.Member -is [System.Management.Automation.Language.StringConstantExpressionAst] -and
                    $methods -contains $n.Member.Value
                }, $true)
            foreach ($call in $calls) {
                $inJob = $false
                $parent = $call.Parent
                while ($null -ne $parent) {
                    if ($parent -is [System.Management.Automation.Language.ScriptBlockExpressionAst] -and
                        $parent.Parent -is [System.Management.Automation.Language.CommandAst] -and
                        $parent.Parent.GetCommandName() -eq "startJob" -and
                        $parent.Parent.CommandElements.IndexOf($parent) -eq 1) {
                        # startJob の最初の引数（裏のスレッドで動くスクリプト）。onDone（画面のスレッド）の中は数えない
                        $inJob = $true
                        break
                    }
                    $parent = $parent.Parent
                }
                if (!$inJob) {
                    $found.Add("$($file.Name):$($call.Extent.StartLineNumber) $($call.Member.Value)")
                }
            }
        }
        return $found.ToArray()
    }
}

Describe "ワークスペースの移動を画面のスレッドで動かさない" -Tag Meta {
    It "画面のファイルは、WorkspaceMover の重い呼び出しを startJob のスクリプトの中でだけ行う" {
        $outside = @(findUnthreadedMoverCalls "${scriptsDir}\tebunko\ui" ${moverMethods})
        $outside | Should -BeNullOrEmpty -Because "ワークスペースを移す・消す処理は裏のスレッド（startJob）で動かすこと（待つ間のダイアログの棒が止まり、応答なしになるため）"
    }

    It "検査が働く（startJob の外の呼び出しを見つける・startJob のスクリプトの中は見逃す・onDone の中は見つける）" {
        $dir = Join-Path $TestDrive "ui"
        New-Item -ItemType Directory -Path $dir | Out-Null
        $text = @'
function direct { $mover.ChangeWorkspace("move", "a", "b") }
function threaded {
    startJob {
        param ($m)
        $m.MoveEntries("a", "b")
    } @() {
        param ($output)
        $mover.RemoveEntries("c")
    }
}
'@
        [System.IO.File]::WriteAllText((Join-Path $dir "sample.ps1"), $text)
        $result = @(findUnthreadedMoverCalls $dir ${moverMethods})
        $result | Should -Be @("sample.ps1:1 ChangeWorkspace", "sample.ps1:8 RemoveEntries")
    }
}
