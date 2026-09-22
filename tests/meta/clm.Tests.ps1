# 制限言語モード（CLM）で動かすファイル（start.ps1 と、そこから読み込むもの）のテスト。
#
# 本物の CLM（WDAC）での確認は CI の clm.yml（tools\clm\check.ps1）が受け持つ。ここでは手元でも動くよう、
# 子プロセスで $ExecutionContext.SessionState.LanguageMode を書き換えた模擬の CLM を使う。
# 模擬は本物より厳しい（Microsoft が署名したモジュールの Expand-Archive なども止まる）ため、模擬で動けば本物でも動く。
. "$PSScriptRoot\..\helpers\load.ps1"

$startScript = "${scriptsDir}\tebunko_grep\start.ps1"

# start.ps1 から dot-source でたどれるファイル（. "$PSScriptRoot\..." の形だけを見る）
function getClmFiles {
    $seen = @{}
    $queue = New-Object System.Collections.Queue
    $queue.Enqueue((Resolve-Path -LiteralPath $startScript).Path)
    while ($queue.Count -gt 0) {
        $path = $queue.Dequeue()
        if ($seen.ContainsKey($path)) { continue }
        $seen[$path] = $true
        $dir = Split-Path $path -Parent
        $text = [System.IO.File]::ReadAllText($path)
        foreach ($match in [regex]::Matches($text, '(?m)^\s*\.\s+"\$PSScriptRoot\\([^"]+)"')) {
            $full = Join-Path $dir $match.Groups[1].Value
            if (Test-Path -LiteralPath $full) { $queue.Enqueue((Resolve-Path -LiteralPath $full).Path) }
        }
    }
    return @($seen.Keys | Sort-Object)
}

# 子プロセスの powershell.exe で start.ps1 を動かし、終了コードと出力を返す。-Clm なら模擬の CLM にしてから動かす
function invokeStart {
    param ([switch]$Clm)

    $quoted = $startScript.Replace("'", "''")
    $prefix = if ($Clm) { "`$ExecutionContext.SessionState.LanguageMode = 'ConstrainedLanguage'; " } else { "" }
    $output = & powershell -NoProfile -ExecutionPolicy Bypass -Command "$prefix& '$quoted' -CheckOnly; exit `$LASTEXITCODE" 2>&1
    return @{ ExitCode = $LASTEXITCODE; Output = (@($output | ForEach-Object { "$_" }) -join "`n") }
}

Describe "制限言語モードで動かすファイル" -Tag Meta {
    $files = getClmFiles

    It "start.ps1 から読み込むファイルがある" {
        $names = @($files | ForEach-Object { Split-Path $_ -Leaf })
        $names -contains "startup_view.ps1" | Should Be $true
        $names -contains "startup_check.ps1" | Should Be $true
        $names -contains "lib_restricted.ps1" | Should Be $true
        $names -contains "index_name.ps1" | Should Be $true
    }

    # 模擬の CLM で読み込むだけでは、実行されない行の書き方は分からないため、CLM で使えない書き方を字面でも調べる
    foreach ($file in $files) {
        $name = Split-Path $file -Leaf
        $text = [System.IO.File]::ReadAllText($file)

        It "$name に PowerShell class が無い" {
            ([regex]::Matches($text, '(?m)^\s*class\s+\w+')).Count | Should Be 0
        }

        # いつもの画面と共用のファイル（shared\・tebunko_grep の core\ など）には、いつもの画面だけが使う関数もある。
        # そこでは [pscustomobject] を許し、制限モードが使う関数は下の「制限モードが使う関数の結果」で確かめる
        if ($file -like "*\tebunko_grep\restricted\*" -or $name -eq "start.ps1") {
            It "$name で [pscustomobject] を作らない" {
                ([regex]::Matches($text, '\[pscustomobject\]\s*@\{')).Count | Should Be 0
            }
        }

        # CLM ではエラーにならずに、ブロックの文字列そのものに置き換わる
        It "$name で [regex]::Replace にスクリプトブロックを渡さない" {
            ([regex]::Matches($text, '\[regex\]::Replace\([^\r\n]*\{')).Count | Should Be 0
        }

        # Add-Type は、FullLanguage のときに画面の部品を読み込めるかを確かめる startup_check.ps1 だけに許す
        if ($name -ne "startup_check.ps1") {
            It "$name で Add-Type を使わない" {
                ([regex]::Matches($text, '(?m)^[^#\r\n]*\bAdd-Type\b')).Count | Should Be 0
            }
        }
    }

    It "模擬の CLM で start.ps1 が、制限モードで起動する旨を出す（終了コード 11）" {
        $result = invokeStart -Clm
        $result.Output | Should Match "ConstrainedLanguage"
        $result.ExitCode | Should Be 11
    }

    It "FullLanguage では、いつもの画面で起動できると返す（終了コード 10）" {
        $result = invokeStart
        $result.ExitCode | Should Be 10
    }
}

Describe "制限モードが使う関数の結果" -Tag Meta {
    # tests\meta\clm_cases.ps1 の呼び出しを、FullLanguage と模擬の制限言語モードの子プロセスでそれぞれ動かし、結果を比べる
    $lib = "${scriptsDir}\tebunko_grep\restricted\lib_restricted.ps1"
    $cases = "$PSScriptRoot\clm_cases.ps1"
    $runner = "$PSScriptRoot\clm_run.ps1"

    function invokeCases {
        param ([string]$label, [switch]$Clm)

        $dir = Join-Path $TestDrive "cases_$label"
        New-Item -ItemType Directory -Path $dir | Out-Null
        $out = Join-Path $TestDrive "result_$label.txt"
        $prefix = if ($Clm) { "`$ExecutionContext.SessionState.LanguageMode = 'ConstrainedLanguage'; " } else { "" }
        $quote = { param($s) "'" + $s.Replace("'", "''") + "'" }
        $command = "$prefix& $(& $quote $runner) -Lib $(& $quote $lib) -Cases $(& $quote $cases) -CaseDir $(& $quote $dir) -Out $(& $quote $out)"
        $output = & powershell -NoProfile -ExecutionPolicy Bypass -Command $command 2>&1
        $results = [ordered]@{}
        if (Test-Path -LiteralPath $out) {
            foreach ($line in (Get-Content -LiteralPath $out -Encoding UTF8)) {
                $i = $line.IndexOf("`t")
                $results[$line.Substring(0, $i)] = $line.Substring($i + 1)
            }
        }
        $results["（出力）"] = (@($output | ForEach-Object { "$_" }) -join "`n")
        return $results
    }

    $full = invokeCases "full"
    $clm = invokeCases "clm" -Clm
    . $cases

    It "呼び出し例がそろって動いた" {
        $full["（出力）"] | Should Be ""
        $clm["（出力）"] | Should Be ""
        @($full.Keys).Count | Should Be (@($clmCases.Keys).Count + 1)
    }

    foreach ($name in @($clmCases.Keys)) {
        It "$name の結果が同じ" {
            $full[$name] | Should Not Match "^エラー: "
            $clm[$name] | Should Be $full[$name]
        }
    }
}
