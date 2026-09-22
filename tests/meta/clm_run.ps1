# tests/meta/clm_cases.ps1 の呼び出しを動かし、1 行 1 件（名前<TAB>結果の JSON）でファイルに書く（tests/meta/clm.Tests.ps1 が使う）。
# 模擬の制限言語モードの子プロセスでも動かすため、制限言語モードで使える書き方だけで書く。
param (
    [string]$Lib,      # 制限モードの読み込み口（scripts\tebunko_grep\restricted\lib_restricted.ps1）
    [string]$Cases,    # 呼び出し例（clm_cases.ps1）
    [string]$CaseDir,  # 呼び出し例がファイルを書いてよいフォルダ（空）
    [string]$Out       # 結果を書くファイル
)

. $Lib
$caseDir = $CaseDir
. $Cases

$lines = @(foreach ($name in $clmCases.Keys) {
        try {
            $value = & $clmCases[$name]
            $json = ConvertTo-Json -InputObject @($value) -Depth 5 -Compress
        } catch {
            $json = "エラー: " + $_.Exception.Message
        }
        $name + "`t" + $json
    })
Set-Content -LiteralPath $Out -Value $lines -Encoding UTF8
