# フォルダ構成の基本（どのツールからも使う）。

# このファイルは scripts\shared\core\ に置く。3 つ上がリポジトリ直下になる
${rootDir} = (Resolve-Path "$PSScriptRoot\..\..\..").Path
${workDir} = "${rootDir}\work"

# 制限言語モード（ConstrainedLanguage。AppLocker・WDAC の PC）では .NET の多くの型を使えない。
# 制限モード（制限言語モードの PC で使うツールの版）が読み込む関数は、これを見て書き方を分ける（fs.ps1 など）
${fullLanguage} = ([string]$ExecutionContext.SessionState.LanguageMode) -eq "FullLanguage"

# BOM 付き UTF-8（制限言語モードでは作れない。そこでは Set-Content -Encoding UTF8 で書く）
${utf8Bom} = $null
if (${fullLanguage}) {
    ${utf8Bom} = New-Object System.Text.UTF8Encoding($true)
}
