# 動いたワークフロー自身の結果を、PR のコメント（ワークフローごとに 1 件）に書く・書き換える
# （.github\actions\pr-comment\action.yml から呼ぶ。各ワークフロー（test・title・docs・codeql・gui・perf-check）が、
# 自分の結果を自分で書く。結果を 1 つのコメントにまとめる集約ワークフローは無い。個別のワークフローをやり直しても、
# そのワークフロー自身のコメントだけが書き換わる）。
#
#   pwsh -File tools\pr_checks_comment.ps1 -Repo <持ち主/名前> -PrNumber <番号> -WorkflowKey <id> -Result <結果>
#     -RunUrl <実行への URL> [-DisplayName <表示名>] [-Note <補足の 1 行>] [-HeadSha <PR の head のコミット>]
#         実際にコメントを書く・書き換える
#   同じ引数に -DryRun を足す
#         書かずに、組み立てたコメントの本文だけを標準出力に出す（gh を呼ばない）
#
# -Result は needs.<job>.result の値（success・failure・cancelled・skipped）。1 つのワークフローに
# 複数のジョブがあるとき（perf-check の search・ingest など）は、カンマ区切りで並べて渡す（悪いほうを採用する）。
#
# 書き換えるのは、作者が github-actions[bot] で、1 行目が目印 <!-- pr-check:<WorkflowKey> --> のコメントだけ。
# 無ければ新しく書く。目印はワークフローごとに違うため、ほかのワークフローが書いたコメントは書き換えない。
#
# 書き込みに失敗しても（フォークの PR など、権限が無く 403 になるとき）、このスクリプトは例外を外に投げない。
# 警告を出すだけで、呼び出し元のジョブを失敗にはしない（結果は Actions の Summary に別途書く）。
#
# 関数はこのファイルを dot-source して Unit テストから呼べる（param はどれも必須にしていないため、
# -Repo などを渡さずに dot-source しても、下の実行部分（$MyInvocation.InvocationName が "." でないときだけ動く）は動かない）。
param (
    [string]$Repo = "",
    [int]$PrNumber = 0,
    [string]$WorkflowKey = "",
    [string]$DisplayName = "",
    [string]$Result = "",
    [string]$RunUrl = "",
    [string]$Note = "",
    [string]$HeadSha = "",
    [switch]$DryRun
)

$botLogin = "github-actions[bot]"

# 結果の値 → 表示文字。知らない値は「その他」
function getResultLabel {
    param ([string]$result)

    switch ($result) {
        "success" { return "成功" }
        "failure" { return "失敗" }
        "cancelled" { return "取り消し" }
        "skipped" { return "スキップ" }
        default { return "その他" }
    }
}

# 1 つのワークフローに複数のジョブがあるとき、悪いほうの結果を採用する。
# 優先度: failure（・知らない値）> cancelled > skipped > success
function combineJobResults {
    param ([string[]]$results)

    $priority = @{ failure = 0; cancelled = 1; skipped = 2; success = 3 }
    $worstResult = "success"
    $worstPriority = 3
    foreach ($r in $results) {
        $p = if ($priority.ContainsKey($r)) { $priority[$r] } else { 0 }
        if ($p -lt $worstPriority) {
            $worstPriority = $p
            $worstResult = $r
        }
    }
    return $worstResult
}

# 実行へのリンクは、この形のときだけ出す（形が違えば出さない）
function getRunLink {
    param ([string]$repo, [string]$url)

    if ([string]::IsNullOrEmpty($url)) { return $null }
    $pattern = "^https://github\.com/" + [regex]::Escape($repo) + "/actions/runs/\d+$"
    if ($url -match $pattern) { return $url }
    return $null
}

# ワークフローごとの目印（1 行目）
function getCommentMarker {
    param ([string]$workflowKey)

    return "<!-- pr-check:$workflowKey -->"
}

# 書き換える対象のコメントを選ぶ。作者が github-actions[bot] で、1 行目がこのワークフローの目印のものだけ（無ければ $null）
function selectExistingComment {
    param ([array]$comments, [string]$marker)

    foreach ($comment in $comments) {
        if ($comment.user.login -ne $botLogin) { continue }
        $firstLine = ($comment.body -split "`r?`n", 2)[0]
        if ($firstLine -eq $marker) { return $comment }
    }
    return $null
}

# コメント本文を作る（1 行目が目印）。result はカンマ区切りで複数渡ってもよい。
# headSha（PR の head のコミット）があれば短い形で添える（遅れて終わった古い実行が上書きしても見分けるため）
function formatCheckComment {
    param ([string]$workflowKey, [string]$displayName, [string]$result, [string]$runUrl, [string]$repo, [string]$note, [string]$headSha = "")

    $combined = combineJobResults ($result -split ",")
    $label = getResultLabel $combined
    $link = getRunLink $repo $runUrl
    $linkText = if ($link) { "[実行]($link)" } else { "実行へのリンクなし" }
    $lines = @((getCommentMarker $workflowKey), "**$displayName** の結果: $label（$linkText）")
    if ($headSha -match '^[0-9a-fA-F]{7,40}$') {
        $lines += ('対象のコミット: `' + $headSha.Substring(0, 7).ToLowerInvariant() + '`')
    }
    if (![string]::IsNullOrEmpty($note)) {
        $lines += ""
        $lines += $note
    }
    return ($lines -join "`n")
}

# gh api で GET したページを 1 件ずつの JSON にして読む
function getApiItems {
    param ([string]$apiPath)

    $lines = & gh api --paginate $apiPath --jq ".[]?"
    if ($LASTEXITCODE -ne 0) { throw "gh api $apiPath に失敗しました（終了コード $LASTEXITCODE）。" }
    if (!$lines) { return @() }
    return @($lines | ForEach-Object { $_ | ConvertFrom-Json })
}

# gh api に渡す引数の配列そのものを作る（post なら新しいコメントを POST、patch なら既存のコメントを PATCH）。
# 本文はファイルの中身を -F（--field）の @<ファイル> で渡す（-f/--raw-field は @ をファイル読み込みと
# 解釈しないため、ここを -f にすると、コメント本文が一時ファイルのパスそのものの文字列になる）。
function getCommentWriteArgs {
    param ([string]$repo, [string]$action, [int]$prNumber, $commentId, [string]$tmpFile)

    $bodyArg = "body=@$tmpFile"
    if ($action -eq "patch") {
        return @("api", "repos/$repo/issues/comments/$commentId", "-X", "PATCH", "-F", $bodyArg)
    }
    return @("api", "repos/$repo/issues/$prNumber/comments", "-F", $bodyArg)
}

# コメントを書く・書き換えるを決めて実行する。
# gh を呼ぶ部分（$getComments・$writeComment）は引数で渡す関数に閉じ込め、Unit テストでは偽物に差し替えて、
# 実際に gh を呼ばずに分岐（PATCH・POST）を確かめる。
function invokePrCheckComment {
    param (
        [string]$repo,
        [int]$prNumber,
        [string]$workflowKey,
        [string]$displayName,
        [string]$result,
        [string]$runUrl,
        [string]$note,
        [string]$headSha = "",
        [scriptblock]$getComments,
        [scriptblock]$writeComment
    )

    $body = formatCheckComment $workflowKey $displayName $result $runUrl $repo $note $headSha
    $marker = getCommentMarker $workflowKey
    $comments = & $getComments $prNumber
    $existing = selectExistingComment $comments $marker

    if ($existing) {
        & $writeComment "patch" $prNumber $existing.id $body
        return [pscustomobject]@{ Action = "patch"; Body = $body }
    }

    & $writeComment "post" $prNumber $null $body
    return [pscustomobject]@{ Action = "post"; Body = $body }
}

# invokePrCheckComment を呼び、書き込みで例外が出ても（権限の無いフォークの PR で 403 になるときなど）
# 外に投げず、Warning に理由を入れて返す（呼び出し元のジョブを失敗にしないため）。書けたときは Warning が $null
function tryInvokePrCheckComment {
    param (
        [string]$repo,
        [int]$prNumber,
        [string]$workflowKey,
        [string]$displayName,
        [string]$result,
        [string]$runUrl,
        [string]$note,
        [string]$headSha = "",
        [scriptblock]$getComments,
        [scriptblock]$writeComment
    )

    try {
        $written = invokePrCheckComment -repo $repo -prNumber $prNumber -workflowKey $workflowKey `
            -displayName $displayName -result $result -runUrl $runUrl -note $note -headSha $headSha `
            -getComments $getComments -writeComment $writeComment
        return [pscustomobject]@{ Action = $written.Action; Body = $written.Body; Warning = $null }
    } catch {
        return [pscustomobject]@{ Action = $null; Body = $null; Warning = $_.Exception.Message }
    }
}

# -Repo・-PrNumber などを渡して実行したときだけ動く（dot-source のときは $MyInvocation.InvocationName が "."）
if ($MyInvocation.InvocationName -ne ".") {
    if ([string]::IsNullOrEmpty($Repo) -or $PrNumber -le 0 -or [string]::IsNullOrEmpty($WorkflowKey) -or [string]::IsNullOrEmpty($Result)) {
        throw "-Repo・-PrNumber（1 以上）・-WorkflowKey・-Result を指定してください。"
    }
    $effectiveDisplayName = if ([string]::IsNullOrEmpty($DisplayName)) { $WorkflowKey } else { $DisplayName }

    $body = formatCheckComment $WorkflowKey $effectiveDisplayName $Result $RunUrl $Repo $Note $HeadSha

    # Actions の Summary には、コメントが書けたかどうかに関わらず結果を残す
    if ($env:GITHUB_STEP_SUMMARY) {
        [System.IO.File]::AppendAllText($env:GITHUB_STEP_SUMMARY, ($body + "`n"), (New-Object System.Text.UTF8Encoding($false)))
    }

    if ($DryRun) {
        Write-Host "PR #$PrNumber"
        Write-Host $body
    } else {
        $getComments = {
            param ($prNumber)
            getApiItems "repos/$Repo/issues/$prNumber/comments?per_page=100"
        }
        $writeComment = {
            param ($action, $prNumber, $commentId, $body)

            $tmpFile = [System.IO.Path]::GetTempFileName()
            try {
                [System.IO.File]::WriteAllText($tmpFile, $body, (New-Object System.Text.UTF8Encoding($false)))
                $ghArgs = getCommentWriteArgs $Repo $action $prNumber $commentId $tmpFile
                & gh @ghArgs | Out-Null
                if ($LASTEXITCODE -ne 0) { throw "gh api に失敗しました（終了コード $LASTEXITCODE）。" }
            } finally {
                Remove-Item -LiteralPath $tmpFile -ErrorAction SilentlyContinue
            }
        }

        $written = tryInvokePrCheckComment -repo $Repo -prNumber $PrNumber -workflowKey $WorkflowKey `
            -displayName $effectiveDisplayName -result $Result -runUrl $RunUrl -note $Note -headSha $HeadSha `
            -getComments $getComments -writeComment $writeComment
        if ($written.Warning) {
            # フォークの PR など、書き込みの権限が無いときもジョブは失敗にしない（結果は上で Summary に書いた）
            Write-Warning "PR へのコメントの書き込みに失敗しました（権限の無いフォークの PR など）: $($written.Warning)"
            # Actions の pwsh は末尾で $LASTEXITCODE を終了コードにする。失敗した gh の 1 が残るとジョブが落ちるため戻す
            $global:LASTEXITCODE = 0
        } else {
            Write-Host "PR #$PrNumber へ $($written.Action) しました。"
        }
    }
}
