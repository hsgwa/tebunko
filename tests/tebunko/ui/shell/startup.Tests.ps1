# 起動時の読み込み（tebunko\ui\shell\startup.ps1 の loadStartupData・startOldSystemIndexRemoval）の、前の版のシステムインデックスの片付けのテスト。
# 関数だけをファイルから取り出し、画面の部品と裏の仕事は偽物にする。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile("${scriptsDir}\tebunko\ui\shell\startup.ps1", [ref]$tokens, [ref]$errors)
    foreach ($name in "startOldSystemIndexRemoval", "loadStartupData") {
        $definition = $ast.Find({ param ($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
        . ([scriptblock]::Create($definition.Extent.Text))
    }

    # 裏の仕事は、その場で（同じスレッドで）流し、終わりの処理を呼ぶ
    function startJob {
        param ([scriptblock]$scriptBlock, [object[]]$arguments, [scriptblock]$onDone, [string]$queue = "default")
        $output = $null
        $errorText = $null
        try { $output = @(& $scriptBlock @arguments) } catch { $errorText = $_.Exception.Message }
        & $onDone $output $errorText
    }
    function writeErrorLog { param ($context, $record) $global:errorLogs.Add("$context : $($record.Exception.Message)") }
    function loadWorkspaceViews { }
    function refreshLegacyIndexMessage { }
    function ensureNetworkDriveCache { }
    function setStatus { param ($text) }
    function startLeftoverCheck { }
    function getOpenCopyDir { param ($ws) return "$($ws.Dir)\open_copy" }
    function clearOldOpenCopies { param ($copyDir, $tmpRoot) return 0 }
}

Describe "起動時の前の版のシステムインデックスの片付け" -Tag Io {
    BeforeEach {
        $global:errorLogs = New-Object System.Collections.Generic.List[string]
        $script:workspace = [Workspace]::new("$TestDrive\ws")
        New-Item -ItemType Directory -Path (Join-Path $script:workspace.OldSystemIndexDir "旧") -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $script:workspace.OldSystemIndexDir "旧\システムインデックス.txt") -Value "x00000000" -Encoding UTF8
        Set-Content -LiteralPath $script:workspace.OldSystemIndexStateFile -Value "旧" -Encoding UTF8
    }
    AfterAll {
        Remove-Variable -Name errorLogs -Scope Global -ErrorAction SilentlyContinue
    }

    It "起動時の読み込み（loadStartupData）が、裏の仕事で system_index と状態ファイルを消す" {
        loadStartupData
        Test-Path -LiteralPath $script:workspace.OldSystemIndexDir | Should -Be $false
        Test-Path -LiteralPath $script:workspace.OldSystemIndexStateFile | Should -Be $false
        @($global:errorLogs).Count | Should -Be 0
    }

    It "startOldSystemIndexRemoval は、消す仕事を画面のスレッドでなく startJob に（workspace の Dir を渡して）渡す" {
        Mock removeOldSystemIndex { return @{ Ok = $true; Reason = "" } }
        Mock startJob { }
        startOldSystemIndexRemoval
        Should -Invoke startJob -Times 1 -Exactly -ParameterFilter { $scriptBlock.ToString() -match 'removeOldSystemIndex' -and @($arguments)[0] -eq $script:workspace.Dir }
        Should -Invoke removeOldSystemIndex -Times 0 -Exactly
    }

    It "消す仕事が <name> ときも例外にせず、理由を画面エラーの記録に残す" -TestCases @(
        @{ name = "失敗を返した"; behavior = { return @{ Ok = $false; Reason = "使用中です" } }; expected = "*使用中です*" }
        @{ name = "例外で終わった"; behavior = { throw "仕事が落ちました" }; expected = "*仕事が落ちました*" }
    ) {
        Mock removeOldSystemIndex $behavior
        { startOldSystemIndexRemoval } | Should -Not -Throw
        @($global:errorLogs).Count | Should -Be 1
        $global:errorLogs[0] | Should -BeLike $expected
    }

    It "消せなくても例外にせず、画面エラーの記録に 1 行残す" {
        $stream = [System.IO.File]::Open((Join-Path $script:workspace.OldSystemIndexDir "旧\システムインデックス.txt"), "Open", "Read", "None")
        try {
            { startOldSystemIndexRemoval } | Should -Not -Throw
        } finally {
            $stream.Dispose()
        }
        @($global:errorLogs).Count | Should -Be 1
        $global:errorLogs[0] | Should -BeLike "前の版が作った検索用のファイルを消せませんでした*"
    }

    It "無ければ何もせず、記録も残さない" {
        Remove-Item -LiteralPath $script:workspace.OldSystemIndexDir -Recurse -Force
        Remove-Item -LiteralPath $script:workspace.OldSystemIndexStateFile -Force
        startOldSystemIndexRemoval
        @($global:errorLogs).Count | Should -Be 0
    }
}
