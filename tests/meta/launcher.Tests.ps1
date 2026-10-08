# tebunko.bat の起動失敗の知らせ（docs/safety/disclosure.md「起動に失敗したときの知らせ（tebunko.bat）」）のテスト。
# tebunko.bat の -Command はファイルに出来ない（ファイルにすると実行ポリシーの対象になり、
# 止まる場面で動かせない）ため、bat から中身を取り出し、子プロセス（powershell.exe）で動かして確かめる。
#   ・$opener（notepad.exe を開く指定）は、開いた対象をファイルに書くだけのものに差し替える
#   ・$root（%~dp0）は、TestDrive に作ったツールの場所に差し替える
#   ・子プロセスには -EncodedCommand で渡す（引用符が崩れないため。テストのコードだけで使い、bat では使わない）
#   ・子プロセスの LOCALAPPDATA・TEMP を TestDrive の下にする（tebunko がここに何も作らないことを確かめる）
BeforeAll {
    $here = (Resolve-Path "$PSScriptRoot\..").Path
    $rootDir = (Resolve-Path "$here\..").Path
    $launcherPath = "$rootDir\tebunko.bat"
    $startupSourceDir = "$rootDir\scripts\tebunko\startup"
    $realScriptsDir = "$rootDir\scripts"

    # -Command の中身は 1 行に書かず、読みやすいよう set "PSCMD=..." と set "PSCMD=%PSCMD%..." で
    # 意味のまとまりごとに組み立てる（bat の中のコメント参照）。各行の中身を順につないで取り出す
    # （1 行目は PSCMD=<中身>、2 行目以降は PSCMD=%PSCMD%<足す中身> の形）
    $script:batText = [System.IO.File]::ReadAllText($launcherPath)
    $setLines = @([regex]::Matches($script:batText, '(?m)^set "PSCMD=(?:%PSCMD%)?(.*)"\r?$'))
    if ($setLines.Count -eq 0) {
        throw "tebunko.bat から PSCMD を組み立てる set 行が見つからない"
    }
    $script:launcherCommand = ($setLines | ForEach-Object { $_.Groups[1].Value }) -join ""

    $script:toolCount = 0
    function newLauncherTool {
        # テストごとに別の場所（ツールの形・サンドボックスの LOCALAPPDATA・TEMP）を作る
        $script:toolCount++
        $root = Join-Path $TestDrive "tool$($script:toolCount)"
        [System.IO.Directory]::CreateDirectory("$root\scripts\tebunko\startup") | Out-Null
        Copy-Item -Path "$startupSourceDir\*" -Destination "$root\scripts\tebunko\startup"
        $sandbox = Join-Path $TestDrive "sandbox$($script:toolCount)"
        [System.IO.Directory]::CreateDirectory("$sandbox\LOCALAPPDATA") | Out-Null
        [System.IO.Directory]::CreateDirectory("$sandbox\TEMP") | Out-Null
        return [pscustomobject]@{
            Root         = $root
            Gui          = "$root\scripts\tebunko\gui.ps1"
            Startup      = "$root\scripts\tebunko\startup"
            LocalAppData = "$sandbox\LOCALAPPDATA"
            Temp         = "$sandbox\TEMP"
        }
    }

    function useRealScripts {
        # 本物の scripts 一式に差し替える（制限言語モード・実行ポリシーの確かめは、本物の gui.ps1 でないと再現しない）
        param ($Tool)
        Remove-Item -LiteralPath "$($Tool.Root)\scripts" -Recurse -Force
        Copy-Item -Path $realScriptsDir -Destination "$($Tool.Root)\scripts" -Recurse
    }

    function setFakeGui {
        # 偽物の gui.ps1（印を書いて終わる・exit・例外を投げる）を置く
        param ($Tool, [string]$Content)
        Set-Content -LiteralPath $Tool.Gui -Value $Content -Encoding UTF8
    }

    function buildLauncherScript {
        # $opener・$root を差し替えた -Command の中身を返す（LanguageMode を渡すと先頭に付ける）
        param ($Tool, [string]$LanguageMode = "")

        $marker = Join-Path $Tool.Temp "opened_marker.txt"
        $openerStub = "`$opener = { param(`$target) Set-Content -LiteralPath '$marker' -Value `$target -Encoding UTF8 }"
        $script = $script:launcherCommand.Replace("`$opener='notepad.exe'", $openerStub)
        if ($script -eq $script:launcherCommand) {
            # bat の書き方が変わって置き換わらないと、本物の notepad.exe が開いたまま残ってしまう
            throw "tebunko.bat に `$opener='notepad.exe' が見つからない（書き方が変わった？）"
        }
        $beforeRoot = $script
        $script = $script.Replace("`$root='%~dp0'", "`$root='$($Tool.Root)\'")
        if ($script -eq $beforeRoot) {
            throw "tebunko.bat に `$root='%~dp0' が見つからない（書き方が変わった？）"
        }
        if ($LanguageMode) {
            $script = "`$ExecutionContext.SessionState.LanguageMode = '$LanguageMode'`r`n" + $script
        }
        return $script
    }

    function invokeLauncher {
        # 子プロセスで動かす。LOCALAPPDATA・TEMP は Tool のサンドボックスに差し替える
        param ($Tool, [string]$Script, [string]$ExecutionPolicy = "Bypass", [int]$TimeoutMs = 20000)

        $bytes = [System.Text.Encoding]::Unicode.GetBytes($Script)
        $encoded = [Convert]::ToBase64String($bytes)

        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "powershell.exe"
        $psi.Arguments = "-NoProfile -ExecutionPolicy $ExecutionPolicy -EncodedCommand $encoded"
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.EnvironmentVariables["LOCALAPPDATA"] = $Tool.LocalAppData
        $psi.EnvironmentVariables["TEMP"] = $Tool.Temp
        $psi.EnvironmentVariables["TMP"] = $Tool.Temp
        $process = [System.Diagnostics.Process]::Start($psi)
        try {
            if (-not $process.WaitForExit($TimeoutMs)) {
                throw "起動が $TimeoutMs ミリ秒以内に終わらなかった"
            }
        } finally {
            if (-not $process.HasExited) {
                Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
            }
        }
        return $process
    }

    function getConstrainedLanguageAddTypeMessage {
        # 制限言語モードで gui.ps1 の最初の読み込み（paths.ps1。New-Object が使えず失敗する）が失敗したときの、このマシン・ロケールでの
        # 実際のメッセージ（英語・日本語などで文言が変わるため、決め打ちにせずその場で再現して得る）。
        # 記録を作る側と同じ起動のしかた（invokeLauncher）で動かし、結果はファイルに書かせて読む
        # （コンソールを引き継ぐ起動では、UI の言語が変わってメッセージの文言が食い違うことがある）
        param ($Tool)
        $out = Join-Path $Tool.Temp "addtype_message.txt"
        $script = "`$ExecutionContext.SessionState.LanguageMode = 'ConstrainedLanguage'; try { . '$($Tool.Root)\scripts\shared\core\paths.ps1' } catch { Set-Content -LiteralPath '$out' -Value `$_.Exception.Message -Encoding UTF8 }"
        invokeLauncher $Tool $script | Out-Null
        if (-not (Test-Path -LiteralPath $out)) { return "" }
        return ((Get-Content -LiteralPath $out -Encoding UTF8) -join "`n").Trim()
    }

    function getLauncherRecord {
        # ツールのフォルダの記録（無ければ $null）
        param ($Tool)
        $path = Join-Path $Tool.Root "startup_error.txt"
        if (Test-Path -LiteralPath $path -PathType Leaf) { return Get-Content -LiteralPath $path -Encoding UTF8 }
        return $null
    }

    function getLauncherOpenedTarget {
        # $opener に渡された対象（メモ帳で開こうとしたファイル）。開いていなければ $null
        param ($Tool)
        $marker = Join-Path $Tool.Temp "opened_marker.txt"
        if (Test-Path -LiteralPath $marker) { return (Get-Content -LiteralPath $marker -Encoding UTF8 -Raw).TrimEnd("`r", "`n") }
        return $null
    }
}

Describe "tebunko.bat の起動失敗の知らせ（scripts\tebunko\startup\）" -Tag Io {
    It "制限言語モードでは、窓を出さずに終わり、記録に制限言語モードの文言・言語モード・実行ポリシーの一覧と、最初の読み込みの失敗が1件だけ書かれ、メモ帳を開く指定に進む" {
        $tool = newLauncherTool
        useRealScripts $tool
        $script = buildLauncherScript $tool -LanguageMode "ConstrainedLanguage"
        $process = invokeLauncher $tool $script

        $process.ExitCode | Should -Be 0
        $record = getLauncherRecord $tool
        $record | Should -Not -BeNullOrEmpty
        ($record -join "`n") | Should -Match "制限言語モード"
        ($record -join "`n") | Should -Match "ConstrainedLanguage"
        ($record -join "`n") | Should -Match "Scope\s+ExecutionPolicy"
        (@($record | Where-Object { $_ -match "^====" })).Count | Should -Be 1
        # 言語モードだけでなく、gui.ps1 の trap から投げ直された元の例外（最初の読み込みの失敗）が
        # 実際に記録されていることも確かめる（別の理由で落ちても言語モードだけでは区別できないため）
        $expectedMessage = getConstrainedLanguageAddTypeMessage $tool
        $expectedMessage | Should -Not -BeNullOrEmpty
        ($record -join "`n") | Should -Match ([regex]::Escape($expectedMessage))
        (getLauncherOpenedTarget $tool) | Should -Be (Join-Path $tool.Root "startup_error.txt")
    }

    It "実行ポリシー（AllSigned）では、記録に実行ポリシーの文言と Get-ExecutionPolicy -List の結果がある" {
        $tool = newLauncherTool
        useRealScripts $tool
        $script = buildLauncherScript $tool
        $process = invokeLauncher $tool $script -ExecutionPolicy "AllSigned"

        $process.ExitCode | Should -Be 0
        $record = getLauncherRecord $tool
        ($record -join "`n") | Should -Match "スクリプトの実行が制限されている"
        ($record -join "`n") | Should -Match "Scope\s+ExecutionPolicy"
        (getLauncherOpenedTarget $tool) | Should -Be (Join-Path $tool.Root "startup_error.txt")
    }

    It "<name>" -TestCases @(
        @{ name = "gui.ps1 が無いと「ファイルが足りない」の文言"; hasGui = $false; guiContent = ""; languageMode = ""; expect = "ファイルが足りない" }
        @{ name = "制限言語モードでも gui.ps1 が無ければ、許可の相談ではなく「ファイルが足りない」の文言"; hasGui = $false; guiContent = ""; languageMode = "ConstrainedLanguage"; expect = "ファイルが足りない" }
        @{ name = "gui.ps1 が例外を投げるとそのメッセージ"; hasGui = $true; guiContent = "throw 'テスト用の例外'"; languageMode = ""; expect = "テスト用の例外" }
    ) {
        param ($name, $hasGui, $guiContent, $languageMode, $expect)
        $tool = newLauncherTool
        if ($hasGui) { setFakeGui $tool $guiContent }
        $script = buildLauncherScript $tool -LanguageMode $languageMode
        invokeLauncher $tool $script | Out-Null

        ((getLauncherRecord $tool) -join "`n") | Should -Match ([regex]::Escape($expect))
    }

    It "<name>" -TestCases @(
        @{ name = "正常終了では、記録を作らずメモ帳も開かない"; guiContent = "exit 0" }
        @{ name = "trap が知らせて exit 1 で終わったときも、記録を作らずメモ帳も開かない（二重に知らせない）"; guiContent = "exit 1" }
    ) {
        param ($name, $guiContent)
        $tool = newLauncherTool
        setFakeGui $tool $guiContent
        $script = buildLauncherScript $tool
        invokeLauncher $tool $script | Out-Null

        (getLauncherRecord $tool) | Should -BeNullOrEmpty
        (getLauncherOpenedTarget $tool) | Should -BeNullOrEmpty
    }

    It "文言のファイルが無くても、記録に英語の1行と詳しい情報が書かれる" {
        $tool = newLauncherTool
        setFakeGui $tool "throw 'テスト用の例外2'"
        Remove-Item -LiteralPath $tool.Startup -Recurse -Force
        $script = buildLauncherScript $tool
        invokeLauncher $tool $script | Out-Null

        $record = getLauncherRecord $tool
        $record[0] | Should -Be "tebunko could not start."
        ($record -join "`n") | Should -Match "テスト用の例外2"
    }

    It "ツールのフォルダに記録を書けないときは、記録を作らず、メモ帳は理由のファイルを開き、LOCALAPPDATA・TEMP には何も作らない" {
        $tool = newLauncherTool
        setFakeGui $tool "throw 'テスト用の例外3'"
        # 記録の名前のフォルダを作って、書き込みを失敗させる
        [System.IO.Directory]::CreateDirectory((Join-Path $tool.Root "startup_error.txt")) | Out-Null
        $script = buildLauncherScript $tool
        invokeLauncher $tool $script | Out-Null

        (getLauncherRecord $tool) | Should -BeNullOrEmpty
        (getLauncherOpenedTarget $tool) | Should -Be (Join-Path $tool.Startup "failed.txt")
        @(Get-ChildItem -LiteralPath $tool.LocalAppData -Force).Count | Should -Be 0
        @(Get-ChildItem -LiteralPath $tool.Temp -Force | Where-Object { $_.Name -ne "opened_marker.txt" }).Count | Should -Be 0
    }
}

Describe "tebunko.bat の書式（ASCII・CRLF・PowerShell の場所・外部プログラム）" -Tag Meta {
    It "ASCII の文字だけで書かれている" {
        $bytes = [System.IO.File]::ReadAllBytes($launcherPath)
        (@($bytes | Where-Object { $_ -gt 0x7F }).Count) | Should -Be 0
    }

    It "改行がすべて CRLF" {
        ([regex]::Matches($script:batText, "(?<!`r)`n")).Count | Should -Be 0
    }

    It "PowerShell を PATH からではなく %SystemRoot% からの絶対パスで指す" {
        $script:batText | Should -Match ([regex]::Escape('%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe'))
    }

    It "PowerShell が無いときは、実行せずにメモ帳で理由を示して終わる" {
        $script:batText | Should -Match "if not exist ""%PS1%"""
        $script:batText | Should -Match ([regex]::Escape('notepad.exe "%~dp0scripts\tebunko\startup\no_powershell.txt"'))
    }

    It "start 行が、組み立てた PSCMD を -Command に渡して conhost.exe 経由で起動する" {
        # set "PSCMD=..." を積み重ねるだけで、実際に渡す先（start 行）が壊れていないことを確かめる
        # （start 行だけを壊しても、上の一致のテストは通ってしまうため）
        $script:batText | Should -Match ([regex]::Escape('start "" conhost.exe "%PS1%" -NoProfile -STA -ExecutionPolicy RemoteSigned -WindowStyle Hidden -Command "%PSCMD%"'))
    }

    It "docs/safety/disclosure.md のコードブロックに、bat の set ""PS1=（PowerShell の場所）から start 行までがそのまま入っている" {
        # 安全性の開示に載せた bat の全文の写しが、bat とずれていないことを確かめる
        $batBody = [regex]::Match($script:batText, '(?s)set "PS1=.*?\r?\nstart "" conhost\.exe[^\r\n]*').Value
        $batBody | Should -Not -BeNullOrEmpty
        $disclosure = [System.IO.File]::ReadAllText("$rootDir\docs\safety\disclosure.md")
        $normalize = { param ($text) ($text -replace "`r`n", "`n") }
        (& $normalize $disclosure).Contains((& $normalize $batBody)) | Should -Be $true
    }

    It "set ""PSCMD=...""（意味のまとまりごと）を順につなぐと、想定した -Command の中身とちょうど一致する" {
        # bat の 6 つの rem（準備・印の解除・try で gui.ps1・catch で理由の選び方・記録の中身・
        # 記録の書き込みと表示）と同じ区切りで書く。ここが変わったら、この期待値も同じ PR で直す
        $step1Prepare = "`$opener='notepad.exe';`$root='%~dp0';`$startup=Join-Path `$root 'scripts\tebunko\startup';`$gui=Join-Path `$root 'scripts\tebunko\gui.ps1';"
        $step2Unblock = "Get-ChildItem -LiteralPath (Join-Path `$root 'scripts') -Recurse -File -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue; "
        $step3TryGui = "try { & `$gui } catch { `$err = `$_; "
        $step4PickReason = "if (-not (Test-Path -LiteralPath `$gui)) { `$reason = Join-Path `$startup 'missing_files.txt' } elseif (`$ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') { `$reason = Join-Path `$startup 'constrained_language.txt' } elseif (`$err.FullyQualifiedErrorId -like 'UnauthorizedAccess*') { `$reason = Join-Path `$startup 'execution_policy.txt' } else { `$reason = Join-Path `$startup 'failed.txt' }; `$reasonLines = @(if (Test-Path -LiteralPath `$reason) { Get-Content -LiteralPath `$reason -Encoding UTF8 } else { @('tebunko could not start.') }); "
        $step5Detail = "`$detailLines = @(('==== {0} startup ====' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')), ('Tool: {0}' -f `$root), ('LanguageMode: {0}' -f `$ExecutionContext.SessionState.LanguageMode), ('PSVersion: {0}' -f `$PSVersionTable.PSVersion), (Get-ExecutionPolicy -List | Out-String), ('{0}' -f `$err.Exception.Message)); `$allLines = `$reasonLines + '' + `$detailLines; "
        $step6WriteShow = "`$openTarget = `$reason; `$candidate = Join-Path `$root 'startup_error.txt'; try { Set-Content -LiteralPath `$candidate -Value `$allLines -Encoding UTF8 -ErrorAction Stop; `$openTarget = `$candidate } catch { }; & `$opener `$openTarget }"

        $expected = $step1Prepare + $step2Unblock + $step3TryGui + $step4PickReason + $step5Detail + $step6WriteShow
        $script:launcherCommand | Should -Be $expected
    }
}

Describe "文言のファイル（scripts\tebunko\startup\）の書式" -Tag Meta {
    BeforeAll {
        $files = Get-ChildItem -LiteralPath $startupSourceDir -Filter "*.txt"
    }

    It "場面ごとに 4 つ以上あり、内容が空でない" {
        ($files.Count -ge 4) | Should -Be $true
        foreach ($file in $files) {
            ([System.IO.File]::ReadAllText($file.FullName).Trim().Length -gt 0) | Should -Be $true -Because $file.Name
        }
    }

    It "すべて BOM 付き UTF-8" {
        $bad = @($files | Where-Object {
            $bytes = [System.IO.File]::ReadAllBytes($_.FullName)
            !($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
        } | ForEach-Object { $_.Name })
        ($bad -join ", ") | Should -Be ""
    }

    It "すべて改行が CRLF" {
        $bad = @($files | Where-Object {
            $text = [System.IO.File]::ReadAllText($_.FullName)
            ([regex]::Matches($text, "(?<!`r)`n")).Count -gt 0
        } | ForEach-Object { $_.Name })
        ($bad -join ", ") | Should -Be ""
    }
}
