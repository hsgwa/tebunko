# スクリプトのクラスのテスト。
# Windows PowerShell 5.1 では、メソッドを持ち、コンストラクタも初期値付きのフィールドも無いクラスを
# Set-PSDebug -Trace の最中に作ると ArgumentOutOfRangeException になる。テストのカバレッジの計測（Pester の Profiler）が
# トレースを使うため、そうしたクラスがあると、そのクラスを作るテストがカバレッジを取るときだけ失敗する。
# 既定のコンストラクタを明示すれば避けられる（scripts\shared\ui\types.ps1 の NotifyBase）。
# 子プロセスでトレースを動かすため 20 秒ほどかかる。コミット前のフック（Unit・Meta）で流さないよう、タグは Io にする
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
}

Describe "クラス" -Tag Io {
    It "引数なしで作れるクラスは、トレースの最中でも作れる" {
        $probe = Join-Path $TestDrive "probe.ps1"
        $result = Join-Path $TestDrive "probe_result.txt"
        # gui.ps1 と同じ順で型を読み込み、scripts の全クラスのうち引数なしで作れるものを、トレースの最中に作る。
        # トレースの出力は多いため、子プロセスで動かして結果だけをファイルに書く
        Set-Content -LiteralPath $probe -Encoding UTF8 -Value @(
            'param([string]$scripts, [string]$result)'
            'Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase'
            '. "$scripts\tebunko\lib.ps1"'
            '. "$scripts\shared\ui\types.ps1"'
            '. "$scripts\tebunko\ui\types.ps1"'
            '$failed = New-Object System.Collections.Generic.List[string]'
            'foreach ($file in Get-ChildItem $scripts -Recurse -Filter *.ps1) {'
            '    $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$null)'
            '    foreach ($type in $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.TypeDefinitionAst] }, $true)) {'
            '        if (!$type.IsClass) { continue }'
            '        $ctors = @($type.Members | Where-Object { $_ -is [System.Management.Automation.Language.FunctionMemberAst] -and $_.IsConstructor })'
            '        if ($ctors.Count -gt 0 -and !@($ctors | Where-Object { $_.Parameters.Count -eq 0 })) { continue }'
            '        if (!($type.Name -as [type])) { continue }'
            '        try {'
            '            Set-PSDebug -Trace 1'
            '            $null = Invoke-Expression "[$($type.Name)]::new()"'
            '        } catch {'
            '            $failed.Add("$($type.Name): $($_.Exception.Message)")'
            '        } finally {'
            '            Set-PSDebug -Off'
            '        }'
            '    }'
            '}'
            '[System.IO.File]::WriteAllLines($result, [string[]]@("done") + $failed.ToArray())'
        )
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $probe -scripts $scriptsDir -result $result *> $null
        $lines = @(Get-Content -LiteralPath $result -Encoding UTF8)
        $lines[0] | Should -Be "done"
        ($lines | Select-Object -Skip 1) -join "`n" | Should -BeNullOrEmpty
    }
}
