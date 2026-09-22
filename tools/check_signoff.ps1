# コミットメッセージに Signed-off-by（DCO の署名）が付いているかを確かめる（AGENTS.md「GitHub の運用」）。
#
#   .\tools\check_signoff.ps1 -Path .git\COMMIT_EDITMSG    これから作るコミットのメッセージ（commit-msg フック）
#   .\tools\check_signoff.ps1 -Range HEAD^1..HEAD^2         範囲に含まれるコミット全部（CI。PR のマージ用のコミットの 2 つの親の間）
#
# 「Signed-off-by: <名前> <メールアドレス>」の行があり、そのメールアドレスがコミットの作者のものと同じなら通す（大文字・小文字は区別しない）。
# git commit -s（--signoff）で付く。作者は -Path では git var GIT_AUTHOR_IDENT（テストでは -AuthorEmail で渡す）、-Range では各コミットの作者。
# 調べないもの: マージコミット（git が作る）、bot（Dependabot など）が作ったコミット、空のメッセージ（git がコミットを中止する）。
# 付いていなければ理由と直し方を出して終了コード 1 で終わる。
param (
    [Parameter(ParameterSetName = "Path", Mandatory = $true)]
    [string]$Path,
    [Parameter(ParameterSetName = "Path")]
    [string]$AuthorEmail,
    [Parameter(ParameterSetName = "Range", Mandatory = $true)]
    [string]$Range
)

$ErrorActionPreference = "Stop"

# bot の作者（GitHub の App のアカウントは <ID>+<名前>[bot]@users.noreply.github.com）
$botMail = '(?i)\[bot\]@users\.noreply\.github\.com$'

# メッセージの本文に、作者のメールアドレスの Signed-off-by があるか
function hasSignOff([string[]]$lines, [string]$mail) {
    foreach ($line in $lines) {
        if ($line -match '^Signed-off-by:\s*.*<(?<mail>[^<>]+)>\s*$' -and $Matches.mail -eq $mail) {
            return $true
        }
    }
    return $false
}

if ($PSCmdlet.ParameterSetName -eq "Path") {
    $lines = [System.IO.File]::ReadAllLines((Resolve-Path -LiteralPath $Path).ProviderPath, (New-Object System.Text.UTF8Encoding($false)))
    # git commit -v の差分は切り取り線より後ろに付く。差分の中の Signed-off-by を拾わないよう、そこで切る
    $scissors = [array]::FindIndex($lines, [Predicate[string]] { param($l) $l -match '^# -+ >8 -+$' })
    if ($scissors -ge 0) { $lines = @($lines | Select-Object -First $scissors) }
    $lines = @($lines | Where-Object { !$_.StartsWith("#") })
    $first = @($lines | Where-Object { $_.Trim() } | Select-Object -First 1)
    if (!$first) { exit 0 }
    if ($first[0] -match '^Merge ') { exit 0 }

    if (!$AuthorEmail) {
        # 「名前 <メールアドレス> 日時 時差」の形
        $ident = git var GIT_AUTHOR_IDENT
        if ($LASTEXITCODE -ne 0 -or $ident -notmatch '<(?<mail>[^<>]*)>') {
            Write-Host "コミットの作者を取得できませんでした（git config user.email を設定してください）。" -ForegroundColor Red
            exit 1
        }
        $AuthorEmail = $Matches.mail
    }
    if ($AuthorEmail -match $botMail -or (hasSignOff $lines $AuthorEmail)) { exit 0 }

    Write-Host "コミットメッセージに Signed-off-by がありません（作者 $AuthorEmail のもの）。" -ForegroundColor Red
    Write-Host "git commit -s（--signoff）でコミットしてください。付けることで、DCO（https://developercertificate.org/）に同意したことを表します。"
    exit 1
}

# 範囲のコミットを「SHA・作者のメールアドレス・メッセージ」で受け取る。区切りは NUL、コミットの間は RS（0x1E）
# 1 行目を一覧に出すため、git の出力を UTF-8 として受け取る（Windows PowerShell 5.1 の既定はコンソールのコードページ）
$encoding = [Console]::OutputEncoding
try {
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
    $output = git log --no-merges "--format=%H%x00%ae%x00%B%x1e" $Range
} finally {
    [Console]::OutputEncoding = $encoding
}
if ($LASTEXITCODE -ne 0) {
    Write-Host "git log $Range が失敗しました。" -ForegroundColor Red
    exit 1
}
$missing = New-Object System.Collections.Generic.List[string]
foreach ($record in (($output -join "`n") -split [char]0x1E)) {
    $fields = $record.TrimStart("`n").Split([char]0)
    if ($fields.Length -lt 3) { continue }
    $sha, $mail, $body = $fields
    if ($mail -match $botMail) { continue }
    if (!(hasSignOff ($body -split "`r?`n") $mail)) {
        $missing.Add("$($sha.Substring(0, 12)) $(($body -split "`r?`n")[0])")
    }
}
if ($missing.Count -eq 0) { exit 0 }

Write-Host "Signed-off-by（作者のメールアドレスのもの）が無いコミットがあります:" -ForegroundColor Red
$missing | ForEach-Object { Write-Host "  $_" }
Write-Host ""
Write-Host "直し方: git rebase --signoff origin/main で付け直し、git push --force-with-lease する。"
Write-Host "以後は git commit -s でコミットする（tools\install_hooks.ps1 を実行してあれば、commit-msg フックが付け忘れを止める）。"
exit 1
