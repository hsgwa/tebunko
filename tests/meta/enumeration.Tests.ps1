# フォルダの列挙子（[System.IO.Directory]::EnumerateFiles・EnumerateDirectories・EnumerateFileSystemEntries、
# DirectoryInfo の EnumerateFiles など）を途中で抜ける書き方を見つけたら落とす。
# 途中で抜けると列挙子が Dispose されず、GC されるまで調べていたフォルダを掴んだまま残る。その間、そのフォルダと
# 上のフォルダを移動・削除できない（インポートの上書きで content_index\<名前> の Directory.Move がアクセス拒否になった）。
# 途中でやめる列挙は shared/core/fs.ps1 の testAnyEntry・selectFirstEntries を通すか、配列で受ける Get* にする。
#
# 見つけるもの:
#   - Enumerate*(...) を foreach で回し、その本体で break・return・throw する（本体の中の別のループ・switch の break、
#     スクリプトブロック・関数の中は除く）
#   - Enumerate*(...).GetEnumerator() を直接呼ぶ（Dispose を忘れやすいため）
#   - Enumerate*(...) を Select-Object -First に流す
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"

    ${enumerateMethodPattern} = '(\.|::)Enumerate(Files|Directories|FileSystemEntries|FileSystemInfos)\s*\('

    function findEarlyExitEnumerations {
        # ファイル path の中の、列挙子を途中で抜ける書き方を「<ファイル名>:<行> <説明>」の並びで返す
        param ([string]$path)

        $tokens = $null
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
        $name = [System.IO.Path]::GetFileName($path)
        $found = New-Object System.Collections.Generic.List[string]

        $loops = $ast.FindAll({ param ($node) $node -is [System.Management.Automation.Language.ForEachStatementAst] }, $true)
        foreach ($loop in $loops) {
            if ($loop.Condition.Extent.Text -notmatch ${enumerateMethodPattern}) {
                continue
            }
            $exits = $loop.Body.FindAll({
                    param ($node)
                    $node -is [System.Management.Automation.Language.BreakStatementAst] -or
                    $node -is [System.Management.Automation.Language.ReturnStatementAst] -or
                    $node -is [System.Management.Automation.Language.ThrowStatementAst]
                }, $true)
            foreach ($exit in $exits) {
                $isBreak = $exit -is [System.Management.Automation.Language.BreakStatementAst]
                $inner = $false
                for ($parent = $exit.Parent; $parent -ne $loop.Body; $parent = $parent.Parent) {
                    if ($parent -is [System.Management.Automation.Language.ScriptBlockExpressionAst] -or
                        $parent -is [System.Management.Automation.Language.FunctionDefinitionAst]) {
                        $inner = $true
                        break
                    }
                    if ($isBreak -and ($parent -is [System.Management.Automation.Language.LoopStatementAst] -or
                            $parent -is [System.Management.Automation.Language.SwitchStatementAst])) {
                        $inner = $true
                        break
                    }
                }
                if (!$inner) {
                    $found.Add("${name}:$($exit.Extent.StartLineNumber) foreach の中の $($exit.Extent.Text.Split("`n")[0].Trim())")
                }
            }
        }

        $members = $ast.FindAll({
                param ($node)
                $node -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and
                $node.Member.Extent.Text -eq "GetEnumerator" -and
                $node.Expression.Extent.Text -match ${enumerateMethodPattern}
            }, $true)
        foreach ($member in $members) {
            $found.Add("${name}:$($member.Extent.StartLineNumber) Enumerate*(...).GetEnumerator()")
        }

        $pipelines = $ast.FindAll({ param ($node) $node -is [System.Management.Automation.Language.PipelineAst] }, $true)
        foreach ($pipeline in $pipelines) {
            $elements = $pipeline.PipelineElements
            if ($elements.Count -lt 2 -or $elements[0].Extent.Text -notmatch ${enumerateMethodPattern}) {
                continue
            }
            foreach ($element in $elements) {
                if ($element -is [System.Management.Automation.Language.CommandAst] -and
                    $element.GetCommandName() -in @("Select-Object", "select") -and $element.Extent.Text -match '-First\b') {
                    $found.Add("${name}:$($pipeline.Extent.StartLineNumber) Enumerate*(...) | Select-Object -First")
                }
            }
        }
        return , $found.ToArray()
    }
}

Describe "列挙子を途中で抜けない" -Tag Meta {
    It "scripts の下に、Enumerate* を途中で抜ける書き方が無い" {
        $found = New-Object System.Collections.Generic.List[string]
        foreach ($file in Get-ChildItem -LiteralPath $scriptsDir -Recurse -Filter *.ps1) {
            foreach ($item in (findEarlyExitEnumerations $file.FullName)) {
                $found.Add($item)
            }
        }
        $found | Should -BeNullOrEmpty -Because "途中でやめる列挙は testAnyEntry・selectFirstEntries（shared/core/fs.ps1）を通すか、Get* で配列に受ける"
    }

    It "見つける: <name>" -TestCases @(
        @{ name = "foreach の中の break"; code = 'foreach ($f in [System.IO.Directory]::EnumerateFiles($d)) { break }' }
        @{ name = "foreach の中の return"; code = 'function f { foreach ($f in [System.IO.Directory]::EnumerateDirectories($d)) { if ($f) { return $true } } }' }
        @{ name = "foreach の中の throw"; code = 'foreach ($f in [System.IO.Directory]::EnumerateFiles($d, "*", [System.IO.SearchOption]::AllDirectories)) { throw "x" }' }
        @{ name = "DirectoryInfo の列挙"; code = 'foreach ($f in (New-Object System.IO.DirectoryInfo($d)).EnumerateFiles("*.tsv")) { return }' }
        @{ name = "入れ子のループの外へ return"; code = 'function f { foreach ($f in [System.IO.Directory]::EnumerateFiles($d)) { foreach ($x in 1..2) { return $x } } }' }
        @{ name = "GetEnumerator を直接呼ぶ"; code = '$e = [System.IO.Directory]::EnumerateFiles($d).GetEnumerator()' }
        @{ name = "Select-Object -First に流す"; code = '$n = [System.IO.Directory]::EnumerateFileSystemEntries($d) | Select-Object -First 10' }
    ) {
        param ($name, $code)
        $path = "$TestDrive\found.ps1"
        [System.IO.File]::WriteAllText($path, $code)

        (findEarlyExitEnumerations $path).Count | Should -Be 1
    }

    It "見つけない: <name>" -TestCases @(
        @{ name = "最後まで回す"; code = 'foreach ($f in [System.IO.Directory]::EnumerateFiles($d)) { $n++ }' }
        @{ name = "continue"; code = 'foreach ($f in [System.IO.Directory]::EnumerateFiles($d)) { if ($f) { continue } }' }
        @{ name = "入れ子のループの break"; code = 'foreach ($f in [System.IO.Directory]::EnumerateFiles($d)) { foreach ($x in 1..2) { break } }' }
        @{ name = "スクリプトブロックの中の return"; code = 'foreach ($f in [System.IO.Directory]::EnumerateFiles($d)) { $s = { return 1 } }' }
        @{ name = "配列で受ける Get*"; code = 'foreach ($f in [System.IO.Directory]::GetFiles($d)) { return $true }' }
        @{ name = "testAnyEntry を通す"; code = 'testAnyEntry ([System.IO.Directory]::EnumerateFiles($d))' }
    ) {
        param ($name, $code)
        $path = "$TestDrive\notfound.ps1"
        [System.IO.File]::WriteAllText($path, $code)

        (findEarlyExitEnumerations $path).Count | Should -Be 0
    }
}
