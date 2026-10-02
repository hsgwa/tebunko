# PR のコメントに結果をまとめるスクリプト（tools\pr_checks_comment.ps1）のテスト
# gh api は呼ばない（純粋な関数だけを dot-source して確かめる）。
# $targetWorkflows を -TestCases（発見の段階で評価する）で使うため、ここでも dot-source する
# （実行の段階ではトップレベルの文は動かないため、BeforeAll でも dot-source し直す）。
. "$PSScriptRoot\..\..\tools\pr_checks_comment.ps1"

BeforeAll {
    . "$PSScriptRoot\..\..\tools\pr_checks_comment.ps1"
}

Describe "getRunResultLabel" -Tag Unit {
    It "<conclusion>/<status> なら <expected>" -TestCases @(
        @{ status = "completed"; conclusion = "success"; expected = "成功" }
        @{ status = "completed"; conclusion = "failure"; expected = "失敗" }
        @{ status = "completed"; conclusion = "cancelled"; expected = "取り消し" }
        @{ status = "completed"; conclusion = "skipped"; expected = "スキップ" }
        @{ status = "completed"; conclusion = "timed_out"; expected = "時間切れ" }
        @{ status = "completed"; conclusion = "action_required"; expected = "承認待ち" }
        @{ status = "completed"; conclusion = "neutral"; expected = "その他" }
        @{ status = "completed"; conclusion = "stale"; expected = "その他" }
        @{ status = "completed"; conclusion = "startup_failure"; expected = "その他" }
        @{ status = "queued"; conclusion = $null; expected = "実行中" }
        @{ status = "in_progress"; conclusion = $null; expected = "実行中" }
        @{ status = "waiting"; conclusion = $null; expected = "その他" }
        @{ status = "completed"; conclusion = ""; expected = "その他" }
    ) {
        param ($status, $conclusion, $expected)
        getRunResultLabel $status $conclusion | Should -Be $expected
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
        @{ reason = "空" ; url = "" }
        @{ reason = "null"; url = $null }
    ) {
        param ($reason, $url)
        getRunLink "hsgwa/tebunko" $url | Should -Be $null
    }
}

Describe "selectLatestRun" -Tag Unit {
    It "実行が無ければ null" {
        selectLatestRun @() | Should -Be $null
        selectLatestRun $null | Should -Be $null
    }

    It "最新の実行を選ぶ（スキップは避ける）" {
        $runs = @(
            [pscustomobject]@{ created_at = "2026-01-01T00:00:00Z"; conclusion = "failure" }
            [pscustomobject]@{ created_at = "2026-01-02T00:00:00Z"; conclusion = "skipped" }
            [pscustomobject]@{ created_at = "2026-01-01T12:00:00Z"; conclusion = "success" }
        )
        (selectLatestRun $runs).conclusion | Should -Be "success"
    }

    It "全部スキップなら、その中の最新を選ぶ" {
        $runs = @(
            [pscustomobject]@{ created_at = "2026-01-01T00:00:00Z"; conclusion = "skipped" }
            [pscustomobject]@{ created_at = "2026-01-02T00:00:00Z"; conclusion = "skipped" }
        )
        (selectLatestRun $runs).created_at | Should -Be "2026-01-02T00:00:00Z"
    }
}

Describe "selectPullRequestBySha" -Tag Unit {
    BeforeAll {
        $script:pulls = @(
            [pscustomobject]@{ number = 1; state = "open"; head = [pscustomobject]@{ sha = "aaa" } }
            [pscustomobject]@{ number = 2; state = "open"; head = [pscustomobject]@{ sha = "bbb" } }
            [pscustomobject]@{ number = 3; state = "closed"; head = [pscustomobject]@{ sha = "ccc" } }
        )
    }

    It "head の SHA が一致する開いた PR を選ぶ" {
        (selectPullRequestBySha $pulls "bbb").number | Should -Be 2
    }

    It "閉じた PR の SHA には一致しない" {
        selectPullRequestBySha $pulls "ccc" | Should -Be $null
    }

    It "どれにも一致しなければ null（古い・進んだ SHA）" {
        selectPullRequestBySha $pulls "zzz" | Should -Be $null
    }
}

Describe "selectExistingComment" -Tag Unit {
    It "github-actions[bot] の、1 行目が目印のコメントを選ぶ" {
        $comments = @(
            [pscustomobject]@{ id = 1; user = [pscustomobject]@{ login = "github-actions[bot]" }; body = "関係ない自動コメント" }
            [pscustomobject]@{ id = 2; user = [pscustomobject]@{ login = "someone" }; body = "<!-- pr-checks -->`n表" }
            [pscustomobject]@{ id = 3; user = [pscustomobject]@{ login = "github-actions[bot]" }; body = "<!-- pr-checks -->`n前の表" }
        )
        (selectExistingComment $comments).id | Should -Be 3
    }

    It "目印の行が無い・作者が違うときは null" {
        $comments = @(
            [pscustomobject]@{ id = 1; user = [pscustomobject]@{ login = "github-actions[bot]" }; body = "目印なし" }
            [pscustomobject]@{ id = 2; user = [pscustomobject]@{ login = "someone" }; body = "<!-- pr-checks -->`n表" }
        )
        selectExistingComment $comments | Should -Be $null
        selectExistingComment @() | Should -Be $null
    }
}

Describe "buildChecksRows" -Tag Unit {
    It "対象の 6 つのワークフローだけの行を作り、関係ない実行は無視する" {
        $runs = @(
            [pscustomobject]@{ path = ".github/workflows/test.yml"; status = "completed"; conclusion = "success"; html_url = "https://github.com/hsgwa/tebunko/actions/runs/1"; created_at = "2026-01-01T00:00:00Z" }
            [pscustomobject]@{ path = ".github/workflows/scorecard.yml"; status = "completed"; conclusion = "success"; html_url = "https://github.com/hsgwa/tebunko/actions/runs/2"; created_at = "2026-01-01T00:00:00Z" }
        )
        $rows = buildChecksRows $runs "hsgwa/tebunko"
        $rows.Count | Should -Be 6
        ($rows | Where-Object { $_.Path -eq "test.yml" }).Result | Should -Be "成功"
        ($rows | Where-Object { $_.Path -eq "title.yml" }).Result | Should -Be "まだ実行がありません"
        $rows.Path | Should -Not -Contain "scorecard.yml"
    }

    It "perf-check.yml の表示名は読み替える" {
        $rows = buildChecksRows @() "hsgwa/tebunko"
        ($rows | Where-Object { $_.Path -eq "perf-check.yml" }).Name | Should -Be "速さの回帰テスト（perf-check）"
    }
}

Describe "formatChecksComment" -Tag Unit {
    BeforeAll {
        $script:rows = @(
            [pscustomobject]@{ Path = "test.yml"; Name = "test"; Result = "成功"; Link = "https://github.com/hsgwa/tebunko/actions/runs/1" }
            [pscustomobject]@{ Path = "title.yml"; Name = "title"; Result = "まだ実行がありません"; Link = $null }
        )
    }

    It "1 行目が目印で、表にリンクが入る" {
        $body = formatChecksComment $rows "0123456789abcdef"
        ($body -split "`n")[0] | Should -Be "<!-- pr-checks -->"
        $body | Should -Match "\[実行\]\(https://github.com/hsgwa/tebunko/actions/runs/1\)"
        $body | Should -Match "0123456"
    }

    It "perf-check の数字を書き写さない（リンクだけ）" {
        $body = formatChecksComment $rows "0123456789abcdef"
        $body | Should -Not -Match "ms|秒"
        $body | Should -Match "ここには書き写していません"
    }

    It "リンクが無い行は「-」にする" {
        $body = formatChecksComment $rows "0123456789abcdef"
        $body | Should -Match "\| まだ実行がありません \| - \|"
    }
}

Describe "pr-comment.yml・各ワークフローの name: との一覧のずれ" -Tag Meta {
    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\.."
    }

    It "pr-comment.yml の workflow_run.workflows: が対象の一覧と同じ" {
        $yml = [System.IO.File]::ReadAllText((Join-Path $root ".github\workflows\pr-comment.yml"))
        $match = [regex]::Match($yml, "workflows:\s*\[([^\]]*)\]")
        $match.Success | Should -BeTrue
        $listed = @($match.Groups[1].Value -split "," | ForEach-Object { $_.Trim() })
        $listed | Should -Be @($targetWorkflows.Values)
    }

    It "対象の各ファイルの name: が一覧のとおり: <path>" -TestCases @(
        $targetWorkflows.Keys | ForEach-Object { @{ path = $_; name = $targetWorkflows[$_] } }
    ) {
        param ($path, $name)
        $yml = [System.IO.File]::ReadAllText((Join-Path $root ".github\workflows\$path"))
        $m = [regex]::Match($yml, "(?m)^name:\s*(\S+)")
        $m.Success | Should -BeTrue
        $m.Groups[1].Value | Should -Be $name
    }
}
