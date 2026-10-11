# コミットメッセージに Signed-off-by（DCO の署名）が付いているかを確かめる（AGENTS.md「GitHub の運用」）。
#
#   .\tools\check_signoff.ps1 -Path .git\COMMIT_EDITMSG    これから作るコミットのメッセージ（commit-msg フック）
#   .\tools\check_signoff.ps1 -Range HEAD^1..HEAD^2         範囲に含まれるコミット全部（CI。PR のマージ用のコミットの 2 つの親の間）
#
# 「Signed-off-by: <名前> <メールアドレス>」の行があり、そのメールアドレスがコミットの作者のものと同じで、作者と違うメールアドレスの行が 1 つも無ければ通す（大文字・小文字は区別しない）。
# git commit -s（--signoff）で付く。作者は -Path では git var GIT_AUTHOR_IDENT（テストでは -AuthorEmail で渡す）、-Range では各コミットの作者。
# 調べないもの: マージコミット（git が作る）、bot（Dependabot など）が作ったコミット、空のメッセージ（git がコミットを中止する）。
# 付いていない・作者と違う行がある場合は理由と直し方を出して終了コード 1 で終わる。作者と違う行のメールアドレスは、CI のログが公開のため出さない（-Range は SHA だけを出す）。
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

# メッセージの本文の Signed-off-by を調べる。none = 作者の行が無い、extra = 作者と違う行がある、ok = 作者の行だけ
function signOffState([string[]]$lines, [string]$mail) {
    $own = $false
    $other = $false
    foreach ($line in $lines) {
        if ($line -notmatch '^Signed-off-by:') { continue }
        # 「Signed-off-by: <名前> <メールアドレス>」の形で、メールアドレスが作者と同じものだけを作者の行にする。それ以外の形は全部、違う行
        if ($line -match '^Signed-off-by:\s*[^<>\s][^<>]*<(?<mail>[^<>]+)>\s*$' -and $Matches.mail -eq $mail) { $own = $true } else { $other = $true }
    }
    if ($other) { return "extra" }
    if ($own) { return "ok" }
    return "none"
}

# 作者と違う行を消して作者の行だけにする直し方（対話なし。git rebase -i と git rebase --signoff は使わない）
function showFix([bool]$rebase) {
    if (!$rebase) {
        Write-Host "コミットメッセージから、作者と違う Signed-off-by の行を消してください。作者の行は git commit -s で付きます。"
        return
    }
    Write-Host "直し方（対話なし）: Signed-off-by を全部消してから、git config user.email を直したうえで git commit -s で作者の 1 行だけを付け直す。"
    Write-Host "  git fetch origin"
    Write-Host "  git rebase origin/main -x ""git log -1 --format=%B | grep -vi '^Signed-off-by:' | git commit --amend -s -F -"""
    Write-Host "  （-s が付けるのはコミッターの行なので、コミットの作者が今の user.email と違うときは --amend -s の後ろに --reset-author を付ける。"
    Write-Host "  push 済みなら、続けて git push --force-with-lease。-x の式は git の sh で動くので PowerShell でも同じに書ける）"
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
    if ($AuthorEmail -match $botMail) { exit 0 }
    $state = signOffState $lines $AuthorEmail
    if ($state -eq "ok") { exit 0 }

    if ($state -eq "extra") {
        Write-Host "コミットメッセージに、作者と違うメールアドレスの Signed-off-by の行があります。" -ForegroundColor Red
        showFix $false
        exit 1
    }
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
$extra = New-Object System.Collections.Generic.List[string]
foreach ($record in (($output -join "`n") -split [char]0x1E)) {
    $fields = $record.TrimStart("`n").Split([char]0)
    if ($fields.Length -lt 3) { continue }
    $sha, $mail, $body = $fields
    if ($mail -match $botMail) { continue }
    $state = signOffState ($body -split "`r?`n") $mail
    if ($state -eq "none") {
        $missing.Add("$($sha.Substring(0, 12)) $(($body -split "`r?`n")[0])")
    } elseif ($state -eq "extra") {
        $extra.Add($sha.Substring(0, 12))
    }
}
if ($missing.Count -eq 0 -and $extra.Count -eq 0) { exit 0 }

if ($missing.Count -gt 0) {
    Write-Host "Signed-off-by（作者のメールアドレスのもの）が無いコミットがあります:" -ForegroundColor Red
    $missing | ForEach-Object { Write-Host "  $_" }
    Write-Host ""
    Write-Host "直し方: git rebase --signoff origin/main で付け直し、git push --force-with-lease する。"
}
if ($extra.Count -gt 0) {
    Write-Host "作者と違うメールアドレスの Signed-off-by の行があるコミットがあります:" -ForegroundColor Red
    $extra | ForEach-Object { Write-Host "  $_" }
    Write-Host ""
    showFix $true
}
Write-Host "以後は git commit -s でコミットする（tools\install_hooks.ps1 を実行してあれば、commit-msg フックが付け忘れを止める）。"
exit 1
