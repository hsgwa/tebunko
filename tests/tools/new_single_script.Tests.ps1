# tools\new_single_script.ps1（展開せずに動く単一 .ps1 版を組み立てる道具）のテスト。
# 作った単一 .ps1 そのものはコミットしない（tests/meta/structure.Tests.ps1 の M5）。設計は docs/design/structure/single-script.md
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"

    $tool = "${scriptsDir}\..\tools\new_single_script.ps1"
    . "${scriptsDir}\..\tools\script_rules.ps1"

    $headSha = (& git -C "${scriptsDir}\.." rev-parse HEAD).Trim()
    $docxSource = "${testDataDir}\office\Word\形式\大文字拡張子.DOCX"
    $pptxSource = "${testDataDir}\office\PowerPoint\基本.pptx"

    # 結合した単一 .ps1 の中で、pattern に一致する最後の行の番号を返す（indexer.Tests.ps1 の findLine と同じ考え方）。
    # 結合した .ps1 は元のソースをそのまま埋め込むため、元のファイルと同じ行の文字列で探せる。
    # ${bundledParts} の中身（lib・indexerLib の文字列）にも同じ行の文字列が入るため、最後の行（実際に動く本体側）を探す
    # 結合した単一 .ps1 の部品（lib・indexerLib）を、製品と同じ経路（getPartLoad）で読んだ新しい runspace を返す。
    # 部品の埋め込み（目印の行より上）だけを実行して ${bundledParts} などを取り出し、getPartLoad に渡す。
    # 戻り値の Prelude を実行した runspace の中で、部品の関数が使える
    function newBundledPartRunspace {
        param ([string]$builtPath, [string]$name)
        $text = [System.IO.File]::ReadAllText($builtPath)
        $end = $text.IndexOf("# ---- BODY-START (above: parts embedded for other threads) ----")
        if ($end -lt 0) { throw "${builtPath} に、部品の埋め込みの終わりの目印がありません" }
        $headPath = Join-Path (Split-Path $builtPath -Parent) "head-only.ps1"
        [System.IO.File]::WriteAllText($headPath, $text.Substring(0, $end), (New-Object System.Text.UTF8Encoding($true)))
        $bundled = & {
            param ($path)
            . $path
            @{ Parts = ${bundledParts}; ScriptPath = ${bundledScriptPath}; Version = ${bundledVersion} }
        } $headPath
        $global:bundledParts = $bundled.Parts
        $global:bundledScriptPath = $bundled.ScriptPath
        $global:bundledVersion = $bundled.Version
        try {
            $load = getPartLoad $name
        } finally {
            Remove-Variable -Name bundledParts, bundledScriptPath, bundledVersion -Scope Global -ErrorAction SilentlyContinue
        }
        $runspace = [runspacefactory]::CreateRunspace($load.State)
        $runspace.Open()
        return @{ Runspace = $runspace; Prelude = $load.Prelude }
    }

    # 単一 .ps1 の中身（byte[]）を別プロセスの Windows PowerShell 5.1 で `-File` 起動する（`-Part indexer`。画面も取り込みも始めない）。
    # 壊れた setting.config を置くので、日本語の記録（indexing_log.txt）が正しく出るのは、ファイルを正しい文字コードで読めたときだけ。
    # 書き先（ワークスペース・USERPROFILE・LOCALAPPDATA・APPDATA）は、すべて testRoot の下に向ける
    function invokeEncodingProbe {
        param ([string]$testRoot, [string]$name, [byte[]]$content, [string]$exe = (Join-Path $PSHOME "powershell.exe"))
        $caseDir = Join-Path $testRoot $name
        New-Item -ItemType Directory -Force -Path "$caseDir\profile\Local", "$caseDir\profile\Roaming" | Out-Null
        $script = Join-Path $caseDir "tebunko.ps1"
        [System.IO.File]::WriteAllBytes($script, $content)
        [System.IO.File]::WriteAllText("$caseDir\setting.config", '{ "targetFolders": [', ${utf8Bom})
        $workspace = Join-Path $caseDir "ws"

        $psi = New-Object System.Diagnostics.ProcessStartInfo $exe
        $psi.Arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$script`" -Part indexer"
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.EnvironmentVariables["TEBUNKO_DEFAULT_WORKSPACE"] = $workspace
        $psi.EnvironmentVariables["USERPROFILE"] = "$caseDir\profile"
        $psi.EnvironmentVariables["LOCALAPPDATA"] = "$caseDir\profile\Local"
        $psi.EnvironmentVariables["APPDATA"] = "$caseDir\profile\Roaming"
        $process = [System.Diagnostics.Process]::Start($psi)
        try {
            $stdout = $process.StandardOutput.ReadToEndAsync()
            $stderr = $process.StandardError.ReadToEndAsync()
            if (-not $process.WaitForExit(120000)) { throw "${name}: 起動したプロセスが 120 秒で終わりませんでした" }
            $errorText = $stderr.Result
            $null = $stdout.Result
            $exitCode = $process.ExitCode
        } finally {
            if (-not $process.HasExited) { Stop-Process -Id $process.Id -Force }
            $process.Dispose()
        }
        $logPath = Join-Path $workspace "indexing_log.txt"
        $log = if (Test-Path -LiteralPath $logPath) { [System.IO.File]::ReadAllText($logPath) } else { "" }
        $broken = @(Get-ChildItem -LiteralPath $caseDir -Filter "setting.config.broken-*").Count
        return @{ ExitCode = $exitCode; Log = $log; ErrorText = $errorText; Broken = $broken }
    }

    function findLine {
        param ([string]$path, [string]$pattern)
        $lines = [System.IO.File]::ReadAllLines($path)
        for ($i = $lines.Count - 1; $i -ge 0; $i--) {
            if ($lines[$i] -match $pattern) { return $i + 1 }
        }
        throw "${path} に ${pattern} がありません"
    }
}

Describe "new_single_script.ps1 の道具の検査" -Tag Io {
    BeforeAll {
        $outFile = Join-Path $TestDrive "release\tebunko-v9.9.9-test.ps1"
        & $tool -Version "v9.9.9" -OutFile $outFile | Out-Null
        $bytes = [System.IO.File]::ReadAllBytes($outFile)
        $text = [System.Text.Encoding]::UTF8.GetString($bytes)
    }

    It "作れる（構文エラー・ファイルの漏れや重複・読み込み行の残り・起動口の param の不足・禁止の語・版の文字列の不一致のどれも起こさない）" {
        Test-Path -LiteralPath $outFile | Should -Be $true
    }

    It "ASCII だけで書き出す（BOM なし・CRLF・どの文字コードで読み直しても同じ）" {
        [regex]::IsMatch($text, '[^\x00-\x7F]') | Should -Be $false
        ($bytes[0], $bytes[1], $bytes[2]) | Should -Not -Be @(0xEF, 0xBB, 0xBF)
        ([regex]::IsMatch($text, "(?<!`r)`n")) | Should -Be $false
        # 日本語の Windows の既定（Shift_JIS）で読んでも、UTF-8 で読んだ文字列と同じ
        ([System.Text.Encoding]::GetEncoding(932).GetString($bytes) -ceq $text) | Should -Be $true
    }

    It "日本語の文字列は [regex]::Unescape で組み立てた形で残る" {
        $text.Contains("([regex]::Unescape('") | Should -Be $true
        $text.Contains('$([regex]::Unescape(') | Should -Be $true
    }

    It "構文として読める（0 個のエラー）" {
        $tokens = $null
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($text, [ref]$tokens, [ref]$errors)
        $errors.Count | Should -Be 0
        $ast | Should -Not -BeNullOrEmpty
    }

    It "読み込み口の dot-source 行（`$PSScriptRoot・`$TebunkoDir）が残っていない" {
        ($text -match '(?m)^[ \t]*\.\s+(?:\(\[WildcardPattern\]::Escape\()?"\$(PSScriptRoot|TebunkoDir)\\[^"]+"(?:\.Replace\([^)]*\)\)\))?[ \t]*$') | Should -Be $false
    }

    It "起動口（gui.ps1・indexer.ps1）の param の名前が、結合した頭の param にすべてある" {
        $tokens = $null
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($text, [ref]$tokens, [ref]$errors)
        $headerNames = @($ast.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
        $headerNames | Should -Contain "Part"
        $headerNames | Should -Contain "RetryFailed"
        $headerNames | Should -Contain "Channel"
    }

    It "本体は、起動の失敗の文言（startup_error_view）、知らせ方（startup_error）、lib の順に並ぶ（gui.ps1 の先頭と同じ順）" {
        $body = $text.Substring($text.IndexOf('# ---- BODY-START'))
        $view = $body.IndexOf('function getStartupErrorMessage')
        $report = $body.IndexOf('function reportStartupFailure')
        $view | Should -BeGreaterThan -1
        $report | Should -BeGreaterThan $view
    }

    It "scripts\ の XAML が、行頭の `'@ を含まない（単一 .ps1 のヒアストリングを閉じてしまわないか）" {
        $xamlFiles = @(Get-ChildItem -Path "${scriptsDir}\tebunko\xaml", "${scriptsDir}\shared\xaml" -Filter *.xaml -Recurse)

        $xamlFiles.Count | Should -BeGreaterThan 0
        foreach ($file in $xamlFiles) {
            $content = [System.IO.File]::ReadAllText($file.FullName)
            ($content -match "(?m)^'@") | Should -Be $false -Because $file.FullName
        }
    }

    It "xaml\ の下のフォルダの XAML（shell\nav.xaml）も埋め込む" {
        $text.Contains('${bundledXaml}[[System.IO.Path]::GetFullPath("$PSScriptRoot\xaml\shell\nav.xaml")]') | Should -Be $true
    }

    It "版の文字列が引数（タグ・今のコミットの SHA）と同じ" {
        $text.Contains("Tag = 'v9.9.9'; Sha = '$headSha'") | Should -Be $true
    }

    It "別プロセスで `-File` 起動すると、BOM の有無にかかわらず、構文エラーも文字化けも無く本体が動き、日本語の記録が正しく出る（<name>）" -TestCases @(
        @{ name = "Windows PowerShell 5.1・作ったまま（BOM なし）"; shell = "powershell"; withBom = $false }
        @{ name = "Windows PowerShell 5.1・BOM を付けたもの"; shell = "powershell"; withBom = $true }
        @{ name = "PowerShell 7・作ったまま（BOM なし）"; shell = "pwsh"; withBom = $false }
        @{ name = "PowerShell 7・BOM を付けたもの"; shell = "pwsh"; withBom = $true }
    ) {
        param ($name, $shell, $withBom)
        $exe = if ($shell -eq "pwsh") {
            $cmd = Get-Command pwsh.exe -ErrorAction SilentlyContinue
            if (-not $cmd) {
                Set-ItResult -Skipped -Because "この機械に PowerShell 7 が無いため"
                return
            }
            $cmd.Source
        } else {
            Join-Path $PSHOME "powershell.exe"
        }
        $content = $bytes
        if ($withBom) {
            $content = New-Object byte[] ($bytes.Length + 3)
            [System.Array]::Copy([byte[]](0xEF, 0xBB, 0xBF), 0, $content, 0, 3)
            [System.Array]::Copy($bytes, 0, $content, 3, $bytes.Length)
        }
        $result = invokeEncodingProbe $TestDrive "$shell-$withBom" $content $exe
        $result.ErrorText | Should -Not -Match "ParserError"
        $result.ExitCode | Should -Be 1
        $result.Broken | Should -Be 1
        $result.Log | Should -Match "設定ファイルが壊れていたため、既定の設定で起動しました"
        $result.Log | Should -Match "クロール対象フォルダがありません"
    }
}

Describe "script_rules.ps1 の関数（new_single_script.ps1 の道具の検査が使う判定）" -Tag Unit {
    It "getDotSourceTargets: `$PSScriptRoot・`$TebunkoDir の dot-source 行から読み込み先の絶対パスを拾う" {
        $dir = Join-Path $TestDrive "rules-a"
        New-Item -ItemType Directory -Force -Path "$dir\tebunko\core" | Out-Null
        $entry = Join-Path $dir "entry.ps1"
        $content = @'
. "$PSScriptRoot\child.ps1"
. "$TebunkoDir\core\x.ps1"
'@
        [System.IO.File]::WriteAllText($entry, $content, ${utf8Bom})

        $targets = @(getDotSourceTargets $entry "$dir\tebunko")

        $targets | Should -Be @("$dir\child.ps1", "$dir\tebunko\core\x.ps1")
    }

    It "getReachableFiles: 連鎖する dot-source をすべてたどり、重複なく返す" {
        $dir = Join-Path $TestDrive "rules-b"
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        [System.IO.File]::WriteAllText("$dir\a.ps1", (@'
. "$PSScriptRoot\b.ps1"
. "$PSScriptRoot\c.ps1"
'@), ${utf8Bom})
        [System.IO.File]::WriteAllText("$dir\b.ps1", (@'
. "$PSScriptRoot\c.ps1"
'@), ${utf8Bom})
        [System.IO.File]::WriteAllText("$dir\c.ps1", "function f {}`r`n", ${utf8Bom})

        $reachable = @(getReachableFiles -entries @("$dir\a.ps1") | Sort-Object)

        $reachable | Should -Be (@("$dir\a.ps1", "$dir\b.ps1", "$dir\c.ps1") | Sort-Object)
    }

    It "findBannedCode: コードの中の禁止の語だけを見つけ、行コメント・行内コメントは見ない" {
        $code = @'
Invoke-WebRequest http://example.com
# Invoke-WebRequest は使わない（説明のコメント）
$x = 1 # Invoke-WebRequest も行内コメントなら見ない
'@

        $hits = findBannedCode $code

        $hits | Should -Match "Network"
        (@([regex]::Matches($hits, "Network")).Count) | Should -Be 1
    }
}

Describe "新しい runspace での部品（lib）の読み込み（結合した単一 .ps1。getPartLoad の経路）" -Tag Slow {
    BeforeAll {
        $builtScript = Join-Path $TestDrive "slow-lib\tebunko-v9.9.9-slow.ps1"
        & $tool -Version "v9.9.9" -OutFile $builtScript | Out-Null
    }

    It "新しい runspace で lib の部品を読み込むと、30 秒以内に戻り、検索の関数が使える" {
        $part = newBundledPartRunspace $builtScript "lib"
        $runspace = $part.Runspace
        $ps = [powershell]::Create()
        $ps.Runspace = $runspace
        try {
            [void]$ps.AddScript($part.Prelude + "`r`n" + 'return @(Get-Command searchContentIndex, newSearchRequest -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })')
            $handle = $ps.BeginInvoke()
            $completed = $handle.AsyncWaitHandle.WaitOne(30000)
            if (!$completed) {
                $ps.Stop()
                throw "30 秒以内に戻りませんでした（lib の部品の読み込みが応答しません）。"
            }
            $names = @($ps.EndInvoke($handle))
            $ps.HadErrors | Should -Be $false
            $names | Should -Contain "searchContentIndex"
            $names | Should -Contain "newSearchRequest"
        } finally {
            $ps.Dispose()
            $runspace.Dispose()
        }
    }

    It "lib の部品を読み込んだ runspace で例外を投げても、その runspace の中のエラーになるだけで、プロセスは残る（MessageBox・exit に入らない）" {
        $part = newBundledPartRunspace $builtScript "lib"
        $runspace = $part.Runspace
        $ps = [powershell]::Create()
        $ps.Runspace = $runspace
        try {
            [void]$ps.AddScript($part.Prelude + "`r`n" + 'throw "テスト用の例外"')
            $handle = $ps.BeginInvoke()
            $completed = $handle.AsyncWaitHandle.WaitOne(30000)
            $completed | Should -Be $true -Because "応答が無ければ、MessageBox などで止まっている疑いがある"
            { $ps.EndInvoke($handle) } | Should -Throw
            $ps.HadErrors | Should -Be $true
            # ここまで来ていること自体が、例外で自分（テストランナー）のプロセスが終わっていない証拠
            (Get-Process -Id $PID) | Should -Not -BeNullOrEmpty
        } finally {
            $ps.Dispose()
            $runspace.Dispose()
        }
    }
}

Describe "`-Part indexer`（結合した単一 .ps1）" -Tag Slow {
    It "tests/testdata の一部を取り込むと、件数が入力のファイル数と同じで失敗が 0 になり、setting.config・work がその場にでき、索引を検索で見つけられる" {
        $dir = Join-Path $TestDrive "ingest"
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        $script = Join-Path $dir "tebunko-v9.9.9-ingest.ps1"
        & $tool -Version "v9.9.9" -OutFile $script | Out-Null

        $source = Join-Path $dir "source"
        New-Item -ItemType Directory -Force -Path $source | Out-Null
        Copy-Item -LiteralPath $docxSource -Destination "$source\議事録.docx"
        Copy-Item -LiteralPath $pptxSource -Destination "$source\資料.pptx"

        $settings = newSettings
        $settings.targetFolders = @(@{ name = "資料"; path = $source; enabled = $true })
        $settings.workspaceFolder = "$dir\work"
        writeSettings $settings "$dir\setting.config"

        # workspaceFolder を明示しているため既定のワークスペースに触れない（ブレークポイントは不要）
        & $script -Part indexer *> $null
        $exitCode = $LASTEXITCODE

        $exitCode | Should -Be 0
        Test-Path -LiteralPath "$dir\setting.config" | Should -Be $true
        Test-Path -LiteralPath "$dir\work\ingest_status.tsv" | Should -Be $true

        $state = getIndexingState -path "$dir\work\ingest_status.tsv"
        $state.Total | Should -Be 2
        $state.Failed | Should -Be 0
        $state.Done | Should -Be 2

        # 索引の検索（lib の部品を、同じプロセスの新しい runspace に getPartLoad の経路で読んで確かめる）
        $part = newBundledPartRunspace $script "lib"
        $runspace = $part.Runspace
        $ps = [powershell]::Create()
        $ps.Runspace = $runspace
        try {
            [void]$ps.AddScript($part.Prelude + "`r`n" + @'
                initWorkspace
                $contentIndexFiles = findContentIndexFiles $workspace.IndexDir
                $search = newSearchRegex "TC21" $true $false
                $filter = newFileFilter ""
                $excludePlace = newPlaceExclude $true $true
                $hits = searchContentIndexFiles $contentIndexFiles 0 $contentIndexFiles.Count $search.Regex -1 $search.TextRegex $search.ScanMode $null $filter.Include $filter.Exclude $excludePlace
                return $hits.Count
'@)
            $handle = $ps.BeginInvoke()
            $completed = $handle.AsyncWaitHandle.WaitOne(30000)
            if (!$completed) {
                $ps.Stop()
                throw "30 秒以内に戻りませんでした（lib の部品での検索が応答しません）。"
            }
            $hitCount = @($ps.EndInvoke($handle))[0]
            $ps.HadErrors | Should -Be $false
            $hitCount | Should -BeGreaterThan 0
        } finally {
            $ps.Dispose()
            $runspace.Dispose()
        }
    }
}
