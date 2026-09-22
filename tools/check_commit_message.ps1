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

# 使える型と使い分け。PR に付けるラベル（.github\release.yml がリリースノートの分類に使う）の目安も並べる
$types = [ordered]@{
    feat     = "機能の追加・変更（ラベル enhancement）"
    fix      = "不具合の修正（ラベル bug）"
    docs     = "文書だけの変更（ラベル documentation）"
    refactor = "動きを変えない書き直し"
    perf     = "速さの改善"
    test     = "テストだけの追加・修正"
    style    = "書式だけの変更（空白・改行など）"
    build    = "配布物の作り方・依存の更新（ラベル dependencies）"
    ci       = "CI・git のフック・開発用の道具"
    chore    = "上のどれにも当たらないもの"
    revert   = "前の変更の取り消し"
}

# git が自動で作るメッセージ
$generated = '^(Merge |Revert "|fixup! |squash! |amend! )'

# 1 行目を調べ、問題があれば理由を返す（問題が無ければ $null）
function findProblem([string]$line) {
    if ($line -match $generated) { return $null }
    if ($line -cnotmatch '^(?<type>[^(!:\s]+)(\((?<scope>[^()]*)\))?!?: (?<subject>.*)$') {
        return "「<型>: <説明>」の形になっていません（型の後ろに半角のコロンと空白が 1 つ要ります）"
    }
    # 後の -match で $Matches が上書きされるため、先に取り出す
    $type = $Matches.type
    $scope = $Matches.scope
    $subject = $Matches.subject
    # [ordered] のキーは大文字・小文字を区別しないため、小文字であることは別に確かめる
    if (!$types.Contains($type) -or $type -cne $type.ToLowerInvariant()) {
        return "型「$type」は使えません"
    }
    if ($null -ne $scope -and $scope -notmatch '^\S+$') {
        return "範囲「($scope)」は空白を含まない 1 語にしてください"
    }
    if ($subject -notmatch '^\S') {
        return "説明がありません（コロンの後ろの空白は 1 つにしてください）"
    }
    return $null
}

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
