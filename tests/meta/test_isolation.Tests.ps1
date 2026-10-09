# テストと実機の確かめが、利用者の本物の既定のワークスペース（Documents\tebunko_ws）に書かない仕組みが崩れていないかを調べる。
# 事故の再発防止: 既定のワークスペースを差し替えずに initWorkspace を呼ぶ・既定のワークスペースを使う画面を開く、を機械的に止める。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    $script:repo = (Resolve-Path "$here\..").Path

    function getBareWorkspaceUses {
        # 既定のワークスペースを差し替えずに使う書き方の行を "ファイル:行" の一覧で返す。
        # $files は @{ Name; Lines } の配列（Name はリポジトリからの相対パス。/ 区切り）、$allowed は書いてよい相対パス。
        # 限界: この検査は書き方の見かけを見るだけ。workspaceFolder = '' （単一引用符）・JSON の "workspaceFolder": ""・
        # 変数を経由した空の指定は見つけられない。見つからなかった分は、run.ps1 の前後の比べ（本物の既定のワークスペース）が最後に止める
        param ([object[]]$files, [string[]]$allowed)

        $result = New-Object System.Collections.Generic.List[string]
        foreach ($file in $files) {
            if ($allowed -contains $file.Name) { continue }
            $number = 0
            foreach ($text in $file.Lines) {
                $number++
                if ($text -match '^\s*initWorkspace\b' -or $text -match 'workspaceFolder\s*=\s*""') {
                    $result.Add("$($file.Name):$number")
                }
            }
        }
        return @($result)
    }
}

Describe "テストの既定のワークスペース" -Tag Meta {
    It "テストの間は、環境変数で使い捨ての場所に差し替わっている" {
        $env:TEBUNKO_DEFAULT_WORKSPACE | Should -Not -BeNullOrEmpty
        getDefaultWorkDir | Should -Be $env:TEBUNKO_DEFAULT_WORKSPACE.TrimEnd("\", "/")
        { assertNotRealWorkspace (getDefaultWorkDir) } | Should -Not -Throw
        { assertNotRealWorkspace $workspace.Dir } | Should -Not -Throw
    }

    It "load.ps1 は、lib.ps1 を読む前に差し替え、initWorkspace のあとに本物でないことを確かめる" {
        $text = [System.IO.File]::ReadAllText("$here\helpers\load.ps1")
        $set = $text.IndexOf('$env:TEBUNKO_DEFAULT_WORKSPACE =')
        $lib = $text.IndexOf('lib.ps1"')
        $init = $text.IndexOf("`ninitWorkspace")
        $assert = $text.IndexOf("assertNotRealWorkspace")
        $set | Should -BeGreaterThan -1
        $set | Should -BeLessThan $lib
        $lib | Should -BeLessThan $init
        $init | Should -BeLessThan $assert
    }

    It "run.ps1 は、Pester の前に使い捨てを入れ、後に前後を比べ、その終了コードで終わる（判定は finishWorkspaceGuard）" {
        $text = [System.IO.File]::ReadAllText("$here\run.ps1")
        $start = $text.IndexOf("startWorkspaceGuard")
        $pester = $text.IndexOf("Invoke-Pester -Configuration")
        $finish = $text.IndexOf("finishWorkspaceGuard")
        $start | Should -BeGreaterThan -1
        $start | Should -BeLessThan $pester
        $finish | Should -BeGreaterThan $pester
        $text | Should -Match 'exit \$guardReport\.ExitCode'
        # 違いの判定の終了コードを、run.ps1 の中に別に書かない（カバレッジの下限の exit 1 と取り違えない）
        $text.Substring($finish) | Should -Match '(?s)^.*exit \$guardReport\.ExitCode'
    }

    It "capture_screens.ps1 は、使い捨てを入れてから撮り、終わりに前後を比べて、違いがあれば失敗にする" {
        $text = [System.IO.File]::ReadAllText("$script:repo\tools\capture_screens.ps1")
        $start = $text.IndexOf("startWorkspaceGuard")
        $finish = $text.IndexOf("finishWorkspaceGuard")
        $start | Should -BeGreaterThan -1
        $finish | Should -BeGreaterThan $start
        $text | Should -Match 'if \(\$guardReport\.ExitCode -ne 0\) \{\s*throw'
        # 作業用のドライブを作るのは、使い捨てを入れたあと
        $text.IndexOf("useCaptureDrive `$tempBase") | Should -BeGreaterThan $start
    }

    It "画面のテストの起動は、起動の前に本物でないことを確かめ、環境変数を渡す" {
        $text = [System.IO.File]::ReadAllText("$here\gui\gui_helpers.ps1")
        $text | Should -Match 'assertNotRealWorkspace \$Tool\.DefaultWorkspace'
        $text | Should -Match '\$env:TEBUNKO_DEFAULT_WORKSPACE\s*='
        $code = @($text -split "\r?\n" | Where-Object { $_ -notmatch '^\s*#' })
        ($code -match 'getRealDefaultWorkspace').Count | Should -BeGreaterThan 0   # コメントではなく、コードで本物を控えている
    }

    It "取り込みの起動の前に、書き込み先が使い捨ての中であることを確かめる（runIndexer）" {
        $text = [System.IO.File]::ReadAllText("$here\helpers\indexer.ps1")
        $text | Should -Match 'assertIndexerWorkspaceIsolated'
    }

    It "既定のワークスペースを差し替えずに使う書き方が、決まったファイルの外に無い" {
        $files = @(Get-ChildItem -LiteralPath $here, "$script:repo\tools" -Recurse -Filter *.ps1 -File |
            Where-Object { $_.Name -ne "test_isolation.Tests.ps1" } |
            ForEach-Object { @{ Name = $_.FullName.Substring($script:repo.Length + 1).Replace("\", "/"); Lines = [System.IO.File]::ReadAllLines($_.FullName) } })
        # load.ps1 は差し替えのあとで呼ぶ。new_single_script・parts のテストは文字列として扱うか、環境変数を入れた子で呼ぶ。
        # settings.Tests.ps1（画面）の S6 は gui_helpers が環境変数を渡す。
        # indexer_workspace_isolation.Tests.ps1 は、空の指定をわざと書いて止まることを確かめる
        $allowed = @("tests/helpers/load.ps1", "tests/tools/new_single_script.Tests.ps1", "tests/tebunko/core/parts.Tests.ps1", "tests/gui/settings.Tests.ps1",
            "tests/tebunko/indexer/indexer_workspace_isolation.Tests.ps1")
        (getBareWorkspaceUses $files $allowed) -join ", " | Should -Be ""
    }

    It "差し替えずに使う書き方の検査は、違反を入れると落ちる: <Case>" -TestCases @(
        @{ Case = "initWorkspace"; Text = "initWorkspace"; Count = 1 }
        @{ Case = "字下げした initWorkspace"; Text = "        initWorkspace"; Count = 1 }
        @{ Case = "workspaceFolder を空にする"; Text = '$s = @{ workspaceFolder = "" }'; Count = 1 }
        @{ Case = "関係ない行"; Text = '$x = 1'; Count = 0 }
        @{ Case = "別の名前の関数"; Text = "initWorkspaceFoo"; Count = 0 }
    ) {
        @(getBareWorkspaceUses @(@{ Name = "x.Tests.ps1"; Lines = @($Text) }) @()).Count | Should -Be $Count
    }

    It "許したのは相対パスだけ。同じ名前の別の場所のファイルは通さない" {
        @(getBareWorkspaceUses @(@{ Name = "tests/other/load.ps1"; Lines = @("initWorkspace") }) @("tests/helpers/load.ps1")).Count | Should -Be 1
    }

    It "検査は、許したファイルなら通す" {
        @(getBareWorkspaceUses @(@{ Name = "tests/helpers/load.ps1"; Lines = @("initWorkspace") }) @("tests/helpers/load.ps1")).Count | Should -Be 0
    }
}
