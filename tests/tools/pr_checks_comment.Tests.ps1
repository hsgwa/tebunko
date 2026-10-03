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

Describe "各ワークフローが、自分の結果を pr-comment の複合アクションに渡している" -Tag Meta {
    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\.."
        $actionPath = Join-Path $root ".github\actions\pr-comment\action.yml"
    }

    It "複合アクション（.github/actions/pr-comment）がある" {
        Test-Path -LiteralPath $actionPath | Should -BeTrue
    }

    It "集約ワークフロー（pr-comment.yml）は無い（各ワークフローが自分で書く方式のため）" {
        Test-Path -LiteralPath (Join-Path $root ".github\workflows\pr-comment.yml") | Should -BeFalse
    }

    It "<path> が pr-comment ジョブで workflow-key: <key> を渡している" -TestCases @(
        @{ path = "test.yml"; key = "test" }
        @{ path = "title.yml"; key = "title" }
        @{ path = "docs.yml"; key = "docs" }
        @{ path = "codeql.yml"; key = "codeql" }
        @{ path = "gui.yml"; key = "gui" }
        @{ path = "perf-check.yml"; key = "perf-check" }
    ) {
        param ($path, $key)
        $yml = [System.IO.File]::ReadAllText((Join-Path $root ".github\workflows\$path"))
        $yml | Should -Match ([regex]::Escape("uses: ./.github/actions/pr-comment"))
        $yml | Should -Match ([regex]::Escape("workflow-key: $key"))
    }
}
