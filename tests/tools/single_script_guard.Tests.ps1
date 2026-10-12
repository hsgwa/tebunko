# tools\single_script_guard.ps1（単一 .ps1 の本体を、BOM が落ちても動く形に包む道具）のテスト。
# 設計は docs/design/structure/single-script.md の「文字コード」
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "${scriptsDir}\..\tools\single_script_guard.ps1"

    # 包んだ結果から、本体の文字列（退避した行は戻す）を取り出す
    function extractBody {
        param ([string]$text)
        $match = [regex]::Match($text, '(?s)\$\{tebunkoBody\} = @''\r?\n(.*?)\r?\n''@')
        if (!$match.Success) { throw "本体の文字列がありません" }
        return [regex]::Replace($match.Groups[1].Value, '(?m)^@@''@', '''@')
    }

    # 包んだ .ps1 を別プロセスで `-File` 起動する。本体が書く出力ファイルの中身・終了コード・標準エラーを返す
    function invokeGuarded {
        param ([string]$dir, [byte[]]$content, [string]$exe = (Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"))
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        $script = Join-Path $dir "guarded.ps1"
        $outPath = Join-Path $dir "out.txt"
        [System.IO.File]::WriteAllBytes($script, $content)
        $psi = New-Object System.Diagnostics.ProcessStartInfo $exe
        $psi.Arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$script`" -Part guard"
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.EnvironmentVariables["TEBUNKO_GUARD_OUT"] = $outPath
        $process = [System.Diagnostics.Process]::Start($psi)
        try {
            $stdout = $process.StandardOutput.ReadToEndAsync()
            $stderr = $process.StandardError.ReadToEndAsync()
            if (-not $process.WaitForExit(60000)) { throw "起動したプロセスが 60 秒で終わりませんでした" }
            $null = $stdout.Result
            $errorText = $stderr.Result
            $exitCode = $process.ExitCode
        } finally {
            if (-not $process.HasExited) { Stop-Process -Id $process.Id -Force }
            $process.Dispose()
        }
        $out = if (Test-Path -LiteralPath $outPath) { [System.IO.File]::ReadAllText($outPath) } else { "" }
        return @{ Script = $script; Out = $out; ExitCode = $exitCode; ErrorText = $errorText }
    }

    # BOM を付けた / 落とした UTF-8 のバイト列にする
    function toBytes {
        param ([string]$text, [bool]$bom)
        return (New-Object System.Text.UTF8Encoding($bom)).GetPreamble() + (New-Object System.Text.UTF8Encoding($false)).GetBytes($text)
    }

    # CP932 で読み違えると、日本語の文字の最後のバイトが次の ASCII（バッククォート・英字）を飲み込む並びを入れた本体
    $sampleBody = @'
param (
    [string]$Part = "gui",
    [switch]$RetryFailed,
    $Channel = $null
)
${bundledScriptPath} = $PSCommandPath
function getRoot { return $PSScriptRoot }
$text = "終わり。`n次の行 ぁe ぁf ぁg ぁh {x} 設定ファイルが壊れていたため"
$hereLike = @'
行の先頭が引用符と @ の行を含む（本体の中のヒアストリング）
{{CLOSE}}
[System.IO.File]::WriteAllText($env:TEBUNKO_GUARD_OUT, ($text + "|" + ${bundledScriptPath} + "|" + (getRoot) + "|" + $Part + "|" + $hereLike.Length), (New-Object System.Text.UTF8Encoding($false)))
exit 7
'@.Replace('{{CLOSE}}', "'@")
    $expectedText = "終わり。`n次の行 ぁe ぁf ぁg ぁh {x} 設定ファイルが壊れていたため"
}

Describe "convertToGuardedScript（本体を包む）" -Tag Unit {
    It "`$PSScriptRoot・`$PSCommandPath を、文字列の中のものも含めて頭が決める名前に書き換え、元の名前を残さない" {
        $body = @'
param ($Part = "gui")
${bundledScriptPath} = $PSCommandPath
$a = "$PSScriptRoot\xaml\x.xaml"
$b = ${PSScriptRoot}
function f { return "$PSCommandPath" }
'@
        $result = extractBody (convertToGuardedScript $body)

        $result | Should -Match ([regex]::Escape('${bundledScriptPath} = ${tebunkoPath}'))
        $result | Should -Match ([regex]::Escape('"${tebunkoRoot}\xaml\x.xaml"'))
        $result | Should -Match ([regex]::Escape('$b = ${tebunkoRoot}'))
        $result | Should -Match ([regex]::Escape('return "${tebunkoPath}"'))
        $result | Should -Not -Match 'PSScriptRoot|PSCommandPath'
    }

    It "本体の中の、行の先頭が '@ の行（ヒアストリングの終わり）を退避し、取り出すと元の本体と同じになる" {
        $body = "param (`$Part = `"gui`")`r`n`${x} = @'`r`n日本語の行`r`n'@`r`n`${y} = 1`r`n"
        $wrapped = convertToGuardedScript $body

        @([regex]::Matches($wrapped, "(?m)^'@\r?$")).Count | Should -Be 2
        $wrapped.Contains("@@'@") | Should -Be $true
        (extractBody $wrapped) | Should -Be $body.TrimEnd("`r", "`n")
        assertGuardedScript $wrapped "テスト"
    }

    It "<name> のとき止まる" -TestCases @(
        @{ name = "param ブロックが無い"; body = "`$a = 1`r`n"; message = "param ブロックがありません" }
        @{ name = "param に ASCII 以外の文字がある"; body = "param (`$日本 = 1)`r`n"; message = "ASCII 以外" }
        @{ name = "行の先頭が @@'@ の行がある"; body = "param (`$a = 1)`r`n`$x = @'`r`n@@'@`r`n'@`r`n"; message = "@@'@" }
        @{ name = "構文エラーがある"; body = "param (`$a = 1)`r`nif (`r`n"; message = "構文エラー" }
    ) {
        param ($name, $body, $message)
        { convertToGuardedScript $body } | Should -Throw "*$message*"
    }

    It "頭の行数だけ空行を足す数が、本体の 1 行目のファイルの行番号から 1 引いた数と同じ" {
        $wrapped = convertToGuardedScript "param (`$a = 1)`r`n`$b = 2`r`n"
        $lines = $wrapped -split "\r?\n"
        $bodyFirst = [array]::IndexOf($lines, '${tebunkoBody} = @''') + 2
        $pad = [int][regex]::Match($wrapped, '\("`r`n" \* (\d+)\)').Groups[1].Value
        $lines[$bodyFirst - 1] | Should -Be 'param ($a = 1)'
        $pad | Should -Be ($bodyFirst - 1)
    }
}

Describe "assertGuardedScript（包んだ結果の確かめ）" -Tag Unit {
    BeforeAll {
        $wrapped = convertToGuardedScript "param (`$a = 1)`r`n`$b = `"日本語。`"`r`n"
    }

    It "正しく包んだものは通る" {
        { assertGuardedScript $wrapped "テスト" } | Should -Not -Throw
    }

    It "<name> は止める" -TestCases @(
        @{ name = "頭に日本語がある"; mutate = { param ($t) "# 日本語`r`n" + $t }; message = "頭に" }
        @{ name = "尾に日本語がある"; mutate = { param ($t) $t + "# 日本語`r`n" }; message = "尾に" }
        @{ name = "判定用の文字が無い"; mutate = { param ($t) $t.Replace([string][char]0x30A2, "A") }; message = "判定用" }
    ) {
        param ($name, $mutate, $message)
        { assertGuardedScript (& $mutate $wrapped) "テスト" } | Should -Throw "*$message*"
    }
}

Describe "包んだ .ps1 を別プロセスで起動する" -Tag Io {
    BeforeAll {
        $wrapped = convertToGuardedScript $sampleBody
    }

    It "<name>: 日本語の文字列がそのまま使え、`$PSScriptRoot・`$PSCommandPath・引数・終了コードが働く" -TestCases @(
        @{ name = "Windows PowerShell 5.1・BOM 付き"; shell = "powershell"; bom = $true }
        @{ name = "Windows PowerShell 5.1・BOM なし（BOM が落ちたもの）"; shell = "powershell"; bom = $false }
        @{ name = "PowerShell 7・BOM 付き"; shell = "pwsh"; bom = $true }
        @{ name = "PowerShell 7・BOM なし（BOM が落ちたもの）"; shell = "pwsh"; bom = $false }
    ) {
        param ($name, $shell, $bom)
        $exe = if ($shell -eq "pwsh") {
            $cmd = Get-Command pwsh.exe -ErrorAction SilentlyContinue
            if (-not $cmd) {
                Set-ItResult -Skipped -Because "この機械に PowerShell 7 が無いため"
                return
            }
            $cmd.Source
        } else {
            Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
        }
        $dir = Join-Path $TestDrive "run-$shell-$bom"
        $result = invokeGuarded $dir (toBytes $wrapped $bom) $exe

        $result.ErrorText | Should -BeNullOrEmpty
        $result.ExitCode | Should -Be 7
        $result.Out | Should -Be ($expectedText + "|" + $result.Script + "|" + $dir + "|guard|" + "行の先頭が引用符と @ の行を含む（本体の中のヒアストリング）".Length)
    }

    It "BOM が落ちても付いていても、エラーの `At line:N` が、ファイルの該当の行と同じ番号になる（5.1）" -TestCases @(
        @{ bom = $true }
        @{ bom = $false }
    ) {
        param ($bom)
        $body = "param (`$Part = `"gui`")`r`n# 日本語のコメント。`r`n`$a = `"終わり。`"`r`nthrow 'boom-marker'`r`n"
        $text = convertToGuardedScript $body
        $dir = Join-Path $TestDrive "line-$bom"
        $result = invokeGuarded $dir (toBytes $text $bom)
        $fileLine = [array]::IndexOf(([System.IO.File]::ReadAllLines($result.Script)), "throw 'boom-marker'") + 1

        $fileLine | Should -BeGreaterThan 1
        $result.ExitCode | Should -Be 1
        $result.ErrorText | Should -Match "boom-marker"
        # 英語は `At line:19 char:1`、日本語の Windows は `発生場所 行:19 文字:1`（標準エラーは文字化けするため行の番号だけで見る）
        $result.ErrorText | Should -Match ":$fileLine\s"
    }

    It "Shift_JIS で保存し直したものは範囲の外: 5.1 では本体の日本語以外が ? になるだけで動き、7 では英語の案内を出して止まる（5.1 のみ確かめる）" {
        $dir = Join-Path $TestDrive "sjis"
        $sjis = [System.Text.Encoding]::GetEncoding(932).GetBytes($wrapped)
        $result = invokeGuarded $dir $sjis

        $result.ExitCode | Should -Be 7
        $result.Out | Should -BeLike "終わり。*"
    }
}
