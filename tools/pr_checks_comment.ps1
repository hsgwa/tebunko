# PR で動いたワークフロー（test・title・docs・codeql・gui・perf-check）の結果を、PR のコメント 1 件にまとめて書く
# （.github\workflows\pr-comment.yml の workflow_run から呼ぶ）。
#
#   pwsh -File tools\pr_checks_comment.ps1 -Repo <持ち主/名前> -HeadSha <SHA>           実際にコメントを書く・書き換える
#   pwsh -File tools\pr_checks_comment.ps1 -Repo <持ち主/名前> -HeadSha <SHA> -DryRun    書かずに、見つけた PR の番号と表だけを標準出力に出す（GET だけ）
#
# PR は常に head の SHA から探す（workflow_run.pull_requests はフォークの PR で空になるため使わない）。
# 一致する開いた PR が無ければ、何もせずに終わる（-DryRun はその旨を出す）。
# 書き換えるのは、作者が github-actions[bot] で 1 行目が目印 <!-- pr-checks --> のコメントだけ。無ければ新しく書く。
# perf-check の数字（Summary の表）はここに書き写さず、実行へのリンクだけを出す（artifact は読まない）。
#
# 対象のワークフローの一覧は $targetWorkflows（ファイル名 → workflow の name:）。
# .github\workflows\pr-comment.yml の workflow_run.workflows: と、各ファイルの name: が、この一覧と
# 食い違っていないかは tests\tools\pr_checks_comment.Tests.ps1 の「一覧のずれ」が確かめる
# （どちらかを改名したときに黙って外れるのを止める）。
#
# 関数はこのファイルを dot-source して Unit テストから呼べる（param はどれも必須にしていないため、
# -Repo・-HeadSha を渡さずに dot-source しても、下の実行部分（$MyInvocation.InvocationName が "." でないときだけ動く）は動かない）。
param (
    [string]$Repo = "",
    [string]$HeadSha = "",
    [switch]$DryRun
)

# 対象のワークフロー（ファイル名 → workflow の name:。並びがコメントの行の並びになる）
$targetWorkflows = [ordered]@{
    "test.yml"       = "test"
    "title.yml"      = "title"
    "docs.yml"       = "docs"
    "codeql.yml"     = "codeql"
    "gui.yml"        = "gui"
    "perf-check.yml" = "perf-check"
}

# 表の行の名前を読み替えるもの（無ければ $targetWorkflows の名前のまま）
$displayNameOverride = @{
    "perf-check.yml" = "速さの回帰テスト（perf-check）"
}

$commentMarker = "<!-- pr-checks -->"
$botLogin = "github-actions[bot]"

# 結果の値 → 表示文字。conclusion が空のときは status（queued・in_progress なら実行中）を見る。知らない値は「その他」
function getRunResultLabel {
    param ($status, $conclusion)

    if ([string]::IsNullOrEmpty($conclusion)) {
        if ($status -eq "queued" -or $status -eq "in_progress") { return "実行中" }
        return "その他"
    }
    switch ($conclusion) {
        "success" { return "成功" }
        "failure" { return "失敗" }
        "cancelled" { return "取り消し" }
        "skipped" { return "スキップ" }
        "timed_out" { return "時間切れ" }
        "action_required" { return "承認待ち" }
        default { return "その他" }
    }
}

# 実行へのリンクは、この形のときだけ出す（フォークが名前を書き換えた実行も、形が違えば出さない）
function getRunLink {
    param ([string]$repo, [string]$htmlUrl)

    if ([string]::IsNullOrEmpty($htmlUrl)) { return $null }
    $pattern = "^https://github\.com/" + [regex]::Escape($repo) + "/actions/runs/\d+$"
    if ($htmlUrl -match $pattern) { return $htmlUrl }
    return $null
}

# 同じワークフロー・同じ head の実行が複数あるとき、スキップでない最新を選ぶ。スキップしか無ければスキップの最新を選ぶ
function selectLatestRun {
    param ([array]$runs)

    if (!$runs -or $runs.Count -eq 0) { return $null }
    $sorted = @($runs | Sort-Object -Property @{ Expression = { [datetime]$_.created_at } } -Descending)
    $notSkipped = @($sorted | Where-Object { $_.conclusion -ne "skipped" })
    if ($notSkipped.Count -gt 0) { return $notSkipped[0] }
    return $sorted[0]
}

# 開いた PR の一覧から、head の SHA が一致するものを選ぶ（無ければ $null）
function selectPullRequestBySha {
    param ([array]$pulls, [string]$headSha)

    foreach ($pr in $pulls) {
        if ($pr.state -eq "open" -and $pr.head.sha -eq $headSha) { return $pr }
    }
    return $null
}

# 書き換える対象のコメントを選ぶ。作者が github-actions[bot] で、1 行目が目印のものだけ（無ければ $null）
function selectExistingComment {
    param ([array]$comments)

    foreach ($comment in $comments) {
        if ($comment.user.login -ne $botLogin) { continue }
        $firstLine = ($comment.body -split "`r?`n", 2)[0]
        if ($firstLine -eq $commentMarker) { return $comment }
    }
    return $null
}

# head の SHA に対する実行の一覧から、対象の 6 つのワークフローごとの行を作る（実行が無ければその旨の文）
function buildChecksRows {
    param ([array]$runs, [string]$repo)

    $rows = @()
    foreach ($path in $targetWorkflows.Keys) {
        $name = $targetWorkflows[$path]
        $displayName = if ($displayNameOverride.Contains($path)) { $displayNameOverride[$path] } else { $name }
        $matching = @($runs | Where-Object { $_.path -eq ".github/workflows/$path" })
        $run = selectLatestRun $matching
        if ($null -eq $run) {
            $rows += [pscustomobject]@{ Path = $path; Name = $displayName; Result = "まだ実行がありません"; Link = $null }
            continue
        }
        $rows += [pscustomobject]@{
            Path   = $path
            Name   = $displayName
            Result = getRunResultLabel $run.status $run.conclusion
            Link   = getRunLink $repo $run.html_url
        }
    }
    return $rows
}

# コメント本文を作る（1 行目が目印 $commentMarker）
function formatChecksComment {
    param ([array]$rows, [string]$headSha)

    $short = $headSha.Substring(0, [Math]::Min(7, $headSha.Length))
    $lines = @($commentMarker, "最新のコミット ``$short`` の結果:", "", "| ワークフロー | 結果 | 実行 |", "|---|---|---|")
    foreach ($row in $rows) {
        $linkCell = if ($row.Link) { "[実行]($($row.Link))" } else { "-" }
        $lines += "| $($row.Name) | $($row.Result) | $linkCell |"
    }
    $lines += ""
    $lines += "perf-check の数字（検索・取り込みの速さ）はここには書き写していません。「速さの回帰テスト（perf-check）」の実行へのリンク先（ジョブの Summary）で見てください。"
    return ($lines -join "`n")
}

# gh api で GET したページを 1 件ずつの JSON にして読む。
# 応答が配列そのものなら -ArrayField を渡さない。{ <ArrayField>: [...] } の形（actions/runs など）なら -ArrayField を渡す
function getApiItems {
    param ([string]$apiPath, [string]$arrayField = "")

    $jqExpr = if ($arrayField) { ".$arrayField[]?" } else { ".[]?" }
    $lines = & gh api --paginate $apiPath --jq $jqExpr
    if ($LASTEXITCODE -ne 0) { throw "gh api $apiPath に失敗しました（終了コード $LASTEXITCODE）。" }
    if (!$lines) { return @() }
    return @($lines | ForEach-Object { $_ | ConvertFrom-Json })
}

# -Repo・-HeadSha を渡して実行したときだけ動く（dot-source のときは $MyInvocation.InvocationName が "."）
if ($MyInvocation.InvocationName -ne ".") {
    $ErrorActionPreference = "Stop"

    if ([string]::IsNullOrEmpty($Repo) -or [string]::IsNullOrEmpty($HeadSha)) {
        throw "-Repo と -HeadSha を指定してください。"
    }

    $encodedSha = [uri]::EscapeDataString($HeadSha)
    $pulls = getApiItems "repos/$Repo/pulls?state=open&per_page=100"
    $pr = selectPullRequestBySha $pulls $HeadSha

    if ($null -eq $pr) {
        Write-Host "PR が見つかりません（head の SHA: $HeadSha）。コメントは書きません。"
        exit 0
    }

    $runs = getApiItems "repos/$Repo/actions/runs?head_sha=$encodedSha&per_page=100" "workflow_runs"
    $rows = buildChecksRows $runs $Repo
    $body = formatChecksComment $rows $HeadSha

    if ($DryRun) {
        Write-Host "PR #$($pr.number)"
        Write-Host $body
        exit 0
    }

    $comments = getApiItems "repos/$Repo/issues/$($pr.number)/comments?per_page=100"
    $existing = selectExistingComment $comments

    $tmpFile = [System.IO.Path]::GetTempFileName()
    try {
        [System.IO.File]::WriteAllText($tmpFile, $body, (New-Object System.Text.UTF8Encoding($false)))
        if ($existing) {
            & gh api "repos/$Repo/issues/comments/$($existing.id)" -X PATCH -f "body=@$tmpFile" | Out-Null
        } else {
            & gh api "repos/$Repo/issues/$($pr.number)/comments" -f "body=@$tmpFile" | Out-Null
        }
        if ($LASTEXITCODE -ne 0) { throw "PR へのコメントの書き込みに失敗しました（終了コード $LASTEXITCODE）。" }
    } finally {
        Remove-Item -LiteralPath $tmpFile -ErrorAction SilentlyContinue
    }
}
