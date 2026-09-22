# 起動時の確認に使う事実を集める（言語モード・画面の部品・Excel・work フォルダ）。
#
# 制限言語モード（CLM）でも読み込めて動く書き方だけで書く（start.ps1 が、画面を開く前に読み込むため）。
# 集めた値の読み方（どちらのモードで起動するか・何を表示するか）は startup_view.ps1 が決める。

# Excel が入っているか。Office の決まった置き場所に EXCEL.EXE があるかで見る。
# レジストリは読まない（docs/04_安全性.md 2.1）。CLM では COM を作れないため、起動して確かめることもしない。
# クイック実行版（root\Office16）と MSI 版（Office16 など）の 64 bit・32 bit を見る
function testExcelInstalled {
    $roots = @($env:ProgramFiles, ${env:ProgramFiles(x86)}) | Where-Object { $_ }
    foreach ($root in $roots) {
        foreach ($pattern in @("Microsoft Office\root\Office*\EXCEL.EXE", "Microsoft Office\Office*\EXCEL.EXE")) {
            if (@(Get-ChildItem -Path (Join-Path $root $pattern) -ErrorAction SilentlyContinue).Count -gt 0) {
                return $true
            }
        }
    }
    return $false
}

# フォルダに書き込めるか。無ければ作り、確認用のファイルを書いて消す
function testFolderWritable {
    param ([string]$dir)

    try {
        if (!(Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null
        }
        $probe = Join-Path $dir ("書き込み確認_" + $PID + ".tmp")
        Set-Content -LiteralPath $probe -Value "" -ErrorAction Stop
        Remove-Item -LiteralPath $probe -Force -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

# 起動時の確認に使う事実を集める（キーの意味は startup_view.ps1 の先頭）。
# WPF は FullLanguage のときだけ読み込んでみる（CLM では Add-Type が使えないことが分かっているため）
function getStartupFacts {
    param ([string]$workDir)

    $mode = [string]$ExecutionContext.SessionState.LanguageMode
    $wpfError = ""
    if ($mode -eq "FullLanguage") {
        try {
            Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
        } catch {
            $wpfError = $_.Exception.Message
        }
    }
    return @{
        LanguageMode = $mode
        WpfError     = $wpfError
        HasExcel     = (testExcelInstalled)
        WorkDir      = $workDir
        CanWriteWork = (testFolderWritable $workDir)
    }
}
