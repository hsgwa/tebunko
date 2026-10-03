# コミットメッセージ・PR のタイトル・Issue のタイトルが Conventional Commits の形かを確かめる（AGENTS.md「GitHub の運用」）。
#
#   .\tools\check_commit_message.ps1 -Path .git\COMMIT_EDITMSG   コミットメッセージのファイルの 1 行目（commit-msg フック）
#   .\tools\check_commit_message.ps1 -Title "feat: ..."           PR・Issue のタイトル（CI。.github\workflows\title.yml）
#
# 形は「<型>(<範囲>)!: <説明>」。範囲と ! は省略できる（! は前の版と互換が無くなる変更）。説明は日本語で書く。
#   例: feat: Excel の図形の文字を検索できるようにする
#       fix(gui): 検索結果の件数が 0 のままになるのを直す
# git が自動で作るメッセージ（Merge ... / Revert "..." / fixup! ... など）は調べない。
# 形が違えば理由と書き方を出して終了コード 1 で終わる。Windows PowerShell 5.1 と PowerShell 7（CI の ubuntu）の両方で動かす。
param (
    [Parameter(ParameterSetName = "Path", Mandatory = $true)]
    [string]$Path,
    [Parameter(ParameterSetName = "Title", Mandatory = $true)]
    [AllowEmptyString()]
    [string]$Title
)

$ErrorActionPreference = "Stop"

# 型・判定の正規表現・findProblem は tools/commit_message_rules.ps1 にある
# （tools/check_compat_golden.ps1 の ! の判定と同じものを使うため、2 か所に書かない）
. (Join-Path $PSScriptRoot "commit_message_rules.ps1")

if ($PSCmdlet.ParameterSetName -eq "Path") {
    # git は # で始まる行をコメントとして捨てる。空行も飛ばし、最初の行を 1 行目とする
    $lines = [System.IO.File]::ReadAllLines((Resolve-Path -LiteralPath $Path).ProviderPath, (New-Object System.Text.UTF8Encoding($false)))
    $first = @($lines | Where-Object { $_.Trim() -and !$_.StartsWith("#") } | Select-Object -First 1)
    # 空のメッセージは git がコミットを中止する
    if (!$first) { exit 0 }
    $line = $first[0]
    $what = "コミットメッセージの 1 行目"
} else {
    $line = $Title
    $what = "タイトル"
}

$problem = findProblem $line
if (!$problem) { exit 0 }

Write-Host "$($what)が Conventional Commits の形ではありません: $problem" -ForegroundColor Red
Write-Host "  $line"
Write-Host ""
Write-Host "書き方: <型>(<範囲>)!: <説明>   （範囲と ! は省略できる。! は前の版と互換が無くなる変更）"
Write-Host "  例: feat: Excel の図形の文字を検索できるようにする"
Write-Host "      fix(gui): 検索結果の件数が 0 のままになるのを直す"
Write-Host "型:"
foreach ($name in $types.Keys) {
    Write-Host ("  {0,-9}{1}" -f $name, $types[$name])
}
exit 1
