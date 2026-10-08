# 動いたワークフロー自身の結果を PR のコメントに書くスクリプト（tools\pr_checks_comment.ps1）のテスト
# gh api は呼ばない（純粋な関数だけを dot-source して確かめる）。
BeforeAll {
    . "$PSScriptRoot\..\..\tools\pr_checks_comment.ps1"
}

Describe "getResultLabel" -Tag Unit {
    It "<result> なら <expected>" -TestCases @(
        @{ result = "success"; expected = "成功" }
        @{ result = "failure"; expected = "失敗" }
        @{ result = "cancelled"; expected = "取り消し" }
        @{ result = "skipped"; expected = "スキップ" }
        @{ result = "neutral"; expected = "その他" }
        @{ result = ""; expected = "その他" }
    ) {
        param ($result, $expected)
        getResultLabel $result | Should -Be $expected
    }
}

Describe "combineJobResults" -Tag Unit {
    It "<reason> なら <expected>" -TestCases @(
        @{ reason = "1 つだけ"; results = @("success"); expected = "success" }
        @{ reason = "failure がどれかにあれば failure"; results = @("success", "failure"); expected = "failure" }
        @{ reason = "cancelled は skipped より悪い"; results = @("skipped", "cancelled"); expected = "cancelled" }
        @{ reason = "skipped は success より悪い"; results = @("success", "skipped"); expected = "skipped" }
        @{ reason = "全部 success なら success"; results = @("success", "success"); expected = "success" }
        @{ reason = "知らない値は failure と同じ扱い"; results = @("success", "neutral"); expected = "neutral" }
    ) {
        param ($reason, $results, $expected)
        combineJobResults $results | Should -Be $expected
    }
}

Describe "getRunLink" -Tag Unit {
    It "自分のリポジトリの実行への URL ならそのまま返す" {
        getRunLink "hsgwa/tebunko" "https://github.com/hsgwa/tebunko/actions/runs/12345" | Should -Be "https://github.com/hsgwa/tebunko/actions/runs/12345"
    }

    It "<reason> なら返さない: <url>" -TestCases @(
        @{ reason = "リポジトリが違う"; url = "https://github.com/other/repo/actions/runs/12345" }
        @{ reason = "実行の番号でない"; url = "https://github.com/hsgwa/tebunko/actions/runs/abc" }
        @{ reason = "ドメインが違う"; url = "https://evil.example.com/hsgwa/tebunko/actions/runs/12345" }
        @{ reason = "末尾に余計な文字がある"; url = "https://github.com/hsgwa/tebunko/actions/runs/12345/jobs/1" }
        @{ reason = "空"; url = "" }
        @{ reason = "null"; url = $null }
    ) {
        param ($reason, $url)
        getRunLink "hsgwa/tebunko" $url | Should -Be $null
    }
}

Describe "getCommentMarker" -Tag Unit {
    It "ワークフローの id ごとに違う目印になる" {
        getCommentMarker "test" | Should -Be "<!-- pr-check:test -->"
        getCommentMarker "perf-check" | Should -Be "<!-- pr-check:perf-check -->"
    }
}

Describe "selectExistingComment" -Tag Unit {
    It "github-actions[bot] の、1 行目がその目印のコメントを選ぶ" {
        $marker = getCommentMarker "test"
        $comments = @(
            [pscustomobject]@{ id = 1; user = [pscustomobject]@{ login = "github-actions[bot]" }; body = "関係ない自動コメント" }
            [pscustomobject]@{ id = 2; user = [pscustomobject]@{ login = "someone" }; body = "$marker`n表" }
            [pscustomobject]@{ id = 3; user = [pscustomobject]@{ login = "github-actions[bot]" }; body = "$marker`n前の結果" }
            [pscustomobject]@{ id = 4; user = [pscustomobject]@{ login = "github-actions[bot]" }; body = "$(getCommentMarker 'gui')`n別のワークフローの結果" }
        )
        (selectExistingComment $comments $marker).id | Should -Be 3
    }

    It "目印の行が無い・作者が違う・ほかのワークフローの目印のときは null" {
        $marker = getCommentMarker "test"
        $comments = @(
            [pscustomobject]@{ id = 1; user = [pscustomobject]@{ login = "github-actions[bot]" }; body = "目印なし" }
            [pscustomobject]@{ id = 2; user = [pscustomobject]@{ login = "someone" }; body = "$marker`n表" }
            [pscustomobject]@{ id = 3; user = [pscustomobject]@{ login = "github-actions[bot]" }; body = "$(getCommentMarker 'gui')`n結果" }
        )
        selectExistingComment $comments $marker | Should -Be $null
        selectExistingComment @() $marker | Should -Be $null
    }
}

Describe "formatCheckComment" -Tag Unit {
    It "1 行目が目印で、結果とリンクが入る" {
        $body = formatCheckComment "test" "test" "success" "https://github.com/hsgwa/tebunko/actions/runs/1" "hsgwa/tebunko" ""
        ($body -split "`n")[0] | Should -Be "<!-- pr-check:test -->"
        $body | Should -Match "成功"
        $body | Should -Match "\[実行\]\(https://github.com/hsgwa/tebunko/actions/runs/1\)"
    }

    It "複数のジョブの結果（カンマ区切り）は悪いほうを採用する" {
        $body = formatCheckComment "perf-check" "速さの回帰テスト（perf-check）" "success,failure" "" "hsgwa/tebunko" ""
        $body | Should -Match "失敗"
    }

    It "補足（note）があれば添える" {
        $body = formatCheckComment "perf-check" "perf-check" "skipped" "" "hsgwa/tebunko" "数字はジョブの Summary で見る。"
        $body | Should -Match "数字はジョブの Summary で見る。"
    }

    It "headSha があれば、短い形で添える（形が違えば添えない）" -TestCases @(
        @{ sha = "0123456789abcdef0123456789abcdef01234567"; expected = '対象のコミット: `0123456`' }
        @{ sha = ""; expected = $null }
        @{ sha = "zzzz"; expected = $null }
    ) {
        param ($sha, $expected)
        $body = formatCheckComment "test" "test" "success" "" "hsgwa/tebunko" "" $sha
        if ($expected) { $body | Should -Match ([regex]::Escape($expected)) }
        else { $body | Should -Not -Match "対象のコミット" }
    }

    It "実行へのリンクが無ければ、その旨を書く" {
        $body = formatCheckComment "test" "test" "success" "" "hsgwa/tebunko" ""
        $body | Should -Match "実行へのリンクなし"
    }
}

Describe "getCommentWriteArgs" -Tag Unit {
    It "<action> なら <note>" -TestCases @(
        @{ action = "post"; prNumber = 42; commentId = $null; note = "issues/<PR番号>/comments に POST する" }
        @{ action = "patch"; prNumber = 42; commentId = 999; note = "issues/comments/<ID> に PATCH する" }
    ) {
        param ($action, $prNumber, $commentId, $note)
        $ghArgs = getCommentWriteArgs "hsgwa/tebunko" $action $prNumber $commentId "C:\temp\abc.txt"

        # 本文は -F（--field）で渡す。-f（--raw-field）は @ をファイル読み込みと解釈しないため、
        # ここが -f に戻ると PR のコメント本文が一時ファイルのパスの文字列になってしまう
        # （-Contain は大小を区別しないため、大文字小文字を区別する -ceq で確かめる）
        @($ghArgs -ceq "-F").Count | Should -Be 1
        @($ghArgs -ceq "-f").Count | Should -Be 0
        $ghArgs | Should -Contain "body=@C:\temp\abc.txt"

        if ($action -eq "patch") {
            $ghArgs | Should -Contain "-X"
            $ghArgs | Should -Contain "PATCH"
            $ghArgs | Should -Contain "repos/hsgwa/tebunko/issues/comments/999"
        } else {
            $ghArgs | Should -Not -Contain "-X"
            $ghArgs | Should -Contain "repos/hsgwa/tebunko/issues/42/comments"
        }
    }
}

Describe "invokePrCheckComment" -Tag Unit {
    BeforeEach {
        $script:writeCalls = @()
        $script:fakeWriteComment = {
            param ($action, $prNumber, $commentId, $body)
            $script:writeCalls += [pscustomobject]@{ Action = $action; PrNumber = $prNumber; CommentId = $commentId; Body = $body }
        }
    }

    It "書き換える対象のコメントが無ければ post で書く" {
        $getComments = { param ($n) @() }

        $result = invokePrCheckComment -repo "hsgwa/tebunko" -prNumber 42 -workflowKey "test" -displayName "test" `
            -result "success" -runUrl "" -note "" -getComments $getComments -writeComment $fakeWriteComment

        $result.Action | Should -Be "post"
        $writeCalls.Count | Should -Be 1
        $writeCalls[0].Action | Should -Be "post"
        $writeCalls[0].PrNumber | Should -Be 42
        $writeCalls[0].CommentId | Should -Be $null
    }

    It "書き換える対象のコメントがあれば patch で書く" {
        $marker = getCommentMarker "test"
        $existing = [pscustomobject]@{ id = 999; user = [pscustomobject]@{ login = "github-actions[bot]" }; body = "$marker`n前の結果" }
        $getComments = { param ($n) @($existing) }

        $result = invokePrCheckComment -repo "hsgwa/tebunko" -prNumber 42 -workflowKey "test" -displayName "test" `
            -result "failure" -runUrl "" -note "" -getComments $getComments -writeComment $fakeWriteComment

        $result.Action | Should -Be "patch"
        $writeCalls.Count | Should -Be 1
        $writeCalls[0].Action | Should -Be "patch"
        $writeCalls[0].CommentId | Should -Be 999
    }

    It "ほかのワークフローのコメントがあっても、そちらは書き換えない" {
        $otherMarker = getCommentMarker "gui"
        $other = [pscustomobject]@{ id = 111; user = [pscustomobject]@{ login = "github-actions[bot]" }; body = "$otherMarker`n別の結果" }
        $getComments = { param ($n) @($other) }

        $result = invokePrCheckComment -repo "hsgwa/tebunko" -prNumber 42 -workflowKey "test" -displayName "test" `
            -result "success" -runUrl "" -note "" -getComments $getComments -writeComment $fakeWriteComment

        $result.Action | Should -Be "post"
        $writeCalls[0].CommentId | Should -Be $null
    }
}

Describe "tryInvokePrCheckComment" -Tag Unit {
    It "書き込みで例外が出ても投げず、Warning に理由を入れて返す（フォークの PR で 403 になるとき）" {
        $getComments = { param ($n) @() }
        $throwingWrite = { param ($action, $prNumber, $commentId, $body) throw "gh api に失敗しました（終了コード 1）。" }

        $r = tryInvokePrCheckComment -repo "hsgwa/tebunko" -prNumber 42 -workflowKey "test" -displayName "test" `
            -result "success" -runUrl "" -note "" -getComments $getComments -writeComment $throwingWrite

        $r.Warning | Should -Match "gh api に失敗しました"
        $r.Action | Should -Be $null
    }

    It "コメントの一覧を読む段階で例外が出ても、投げずに Warning を返す" {
        $throwingGet = { param ($n) throw "403" }
        $noWrite = { param ($action, $prNumber, $commentId, $body) }

        $r = tryInvokePrCheckComment -repo "hsgwa/tebunko" -prNumber 42 -workflowKey "test" -displayName "test" `
            -result "success" -runUrl "" -note "" -getComments $throwingGet -writeComment $noWrite

        $r.Warning | Should -Be "403"
    }

    It "書けたときは Warning が無く、Action を返す" {
        $getComments = { param ($n) @() }
        $okWrite = { param ($action, $prNumber, $commentId, $body) }

        $r = tryInvokePrCheckComment -repo "hsgwa/tebunko" -prNumber 42 -workflowKey "test" -displayName "test" `
            -result "success" -runUrl "" -note "" -headSha "0123456789abcdef" -getComments $getComments -writeComment $okWrite

        $r.Warning | Should -Be $null
        $r.Action | Should -Be "post"
        $r.Body | Should -Match "対象のコミット"
    }
}

Describe "各ワークフローが、自分の結果を pr-comment の複合アクションに渡している" -Tag Meta {
    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\.."
        $actionPath = Join-Path $root ".github\actions\pr-comment\action.yml"
    }

    It "複合アクション（.github/actions/pr-comment）がある" {
        Test-Path -LiteralPath $actionPath | Should -BeTrue
    }

    # 6 か所に同じ形で写したジョブの肝（権限を pr-comment ジョブだけに絞る・既定のブランチの内容で動かす）が食い違わないことを確かめる
    It "<path> の pr-comment ジョブは workflow-key: <key> で複合アクションを呼び、権限と checkout の肝を守っている" -TestCases @(
        @{ path = "test.yml"; key = "test" }
        @{ path = "title.yml"; key = "title" }
        @{ path = "docs.yml"; key = "docs" }
        @{ path = "codeql.yml"; key = "codeql" }
        @{ path = "gui.yml"; key = "gui" }
        @{ path = "perf-check.yml"; key = "perf-check" }
    ) {
        param ($path, $key)
        $yml = [System.IO.File]::ReadAllText((Join-Path $root ".github\workflows\$path")) -replace "`r`n", "`n"

        # ワークフロー全体の permissions（jobs: より前）は pull-requests を持たない
        $header = ($yml -split "(?m)^jobs:", 2)[0]
        $header | Should -Not -Match "pull-requests"

        # pr-comment ジョブの本文（次のジョブか末尾まで）
        $m = [regex]::Match($yml, "(?ms)^  pr-comment:\n(.*?)(?=^  [A-Za-z0-9_-]+:[ ]*\n|\z)")
        $m.Success | Should -BeTrue
        $job = $m.Groups[1].Value

        # 書き込みの権限は、このジョブだけが持つ
        $job | Should -Match "(?m)^      pull-requests: write$"

        # 複合アクションを呼ぶ前に、既定のブランチの内容を checkout している（PR のコードを書き込み権限で動かさない）
        $usesIndex = $job.IndexOf("uses: ./.github/actions/pr-comment")
        $usesIndex | Should -BeGreaterThan 0
        $before = $job.Substring(0, $usesIndex)
        $before | Should -Match ([regex]::Escape('ref: ${{ github.event.repository.default_branch }}'))
        $before | Should -Match "persist-credentials: false"
        $job | Should -Match ([regex]::Escape("workflow-key: $key"))
    }

    It "release.yml は test.yml を呼ぶジョブに pull-requests: write を許している（足りないと release が起動しない）" {
        $yml = [System.IO.File]::ReadAllText((Join-Path $root ".github\workflows\release.yml")) -replace "`r`n", "`n"
        $yml | Should -Match "(?ms)^  test:\n.*?uses: \./\.github/workflows/test\.yml\n.*?^      pull-requests: write$"
    }

    It "perf-check.yml の pr-comment は、perf-check 以外のラベルを付けたときには書かない" {
        $yml = [System.IO.File]::ReadAllText((Join-Path $root ".github\workflows\perf-check.yml")) -replace "`r`n", "`n"
        $yml | Should -Match ([regex]::Escape("(github.event.action != 'labeled' || github.event.label.name == 'perf-check')"))
    }
}

# GitHub Actions の pwsh の run は、末尾で $LASTEXITCODE を終了コードにする。
# gh が失敗で終わったあとに $LASTEXITCODE が残ると、警告だけのはずのジョブが失敗になる
Describe "書き込みに失敗してもジョブの終了コードは 0（フォークの PR など）" -Tag Unit {
    BeforeAll {
        $script:scriptPath = (Resolve-Path "$PSScriptRoot\..\..\tools\pr_checks_comment.ps1").Path
        $script:fakeBin = Join-Path $TestDrive "fakebin"
        New-Item -ItemType Directory -Path $script:fakeBin | Out-Null
        # 本物の gh は呼ばない。常に失敗で終わる偽の gh
        [System.IO.File]::WriteAllText((Join-Path $script:fakeBin "gh.cmd"), "@echo off`r`nexit /b 1`r`n")
        $script:savedPath = $env:PATH
    }

    AfterAll {
        $env:PATH = $script:savedPath
    }

    It "gh が失敗しても、ランナーと同じ終わり方で終了コードが 0 になり、警告を出す" {
        $env:PATH = "$($script:fakeBin);$($script:savedPath)"
        try {
            $command = "& '$($script:scriptPath)' -Repo 'example/repo' -PrNumber 1 -WorkflowKey 'test' -Result 'success'; " +
                "if ((Test-Path -LiteralPath variable:\LASTEXITCODE)) { exit `$LASTEXITCODE }"
            $out = & powershell.exe -NoProfile -NonInteractive -Command $command 3>&1 2>&1 | Out-String
            $code = $LASTEXITCODE
        } finally {
            $env:PATH = $script:savedPath
        }
        $out | Should -Match "書き込みに失敗しました"
        $code | Should -Be 0
    }
}
