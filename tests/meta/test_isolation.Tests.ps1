# テストと実機の確かめが、利用者の既定のワークスペース（Documents\tebunko_ws）に書かない仕組みが崩れていないかを調べる。
# 事故の再発防止: 既定のワークスペースを差し替えずに initWorkspace を呼ぶ・既定のワークスペースを使う画面を開く、を機械的に止める。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    $script:repo = (Resolve-Path "$here\..").Path

    function getBareWorkspaceUses {
        # 既定のワークスペースを差し替えずに使う書き方の行を "ファイル:行" の一覧で返す。
        # $files は @{ Name; Lines } の配列（Name はリポジトリからの相対パス。/ 区切り）、$allowed は書いてよい相対パス。
        # 限界: この検査は書き方の見かけを見るだけ。workspaceFolder = '' （単一引用符）・JSON の "workspaceFolder": ""・
        # 変数を経由した空の指定は見つけられない。見つからなかった分は、run.ps1 の前後の比べ（既定のワークスペース）が最後に止める
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

    function getToolCodeLines {
        # 本文の、コメントでない行の配列
        param ([string]$Text)
        return @($Text -split "\r?\n" | Where-Object { $_ -notmatch '^\s*#' })
    }

    function getToolGuardOrderViolations {
        # 性能・変換の道具の守りの順（読み込む前の差し替え・終わりの比べ・違いがあれば失敗）の違反の一覧（空なら通る）
        param ([string]$Text)

        $found = New-Object System.Collections.Generic.List[string]
        $lines = getToolCodeLines $Text
        $start = -1; $finish = -1; $load = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($start -lt 0 -and $lines[$i] -match 'startWorkspaceGuard') { $start = $i }
            if ($finish -lt 0 -and $lines[$i] -match 'finishWorkspaceGuard') { $finish = $i }
            if ($load -lt 0 -and $lines[$i] -match '^\s*\.\s+\(.*(lib\.ps1|\$lib\b)') { $load = $i }
        }
        if ($start -lt 0) {
            $found.Add("startWorkspaceGuard が無い")
        } elseif ($load -ge 0 -and $load -lt $start) {
            $found.Add("lib.ps1 を読む行が startWorkspaceGuard より前にある")
        }
        if ($finish -lt 0) {
            $found.Add("finishWorkspaceGuard が無い")
        } elseif ($start -ge 0 -and $finish -lt $start) {
            $found.Add("finishWorkspaceGuard が startWorkspaceGuard より前にある")
        }
        if ($Text -notmatch 'if \(\$guardReport\.ExitCode -ne 0\) \{\s*throw') {
            $found.Add('$guardReport.ExitCode -ne 0 の直後に throw が無い')
        }
        return @($found)
    }

    function getToolPathCheckViolations {
        # assertToolPathsNotRealWorkspace を呼ぶ行が、最初の書き込み（Copy-Item・WriteAll・CreateDirectory）より前にあるかの違反の一覧（空なら通る）
        param ([string]$Text)

        $found = New-Object System.Collections.Generic.List[string]
        $lines = getToolCodeLines $Text
        $call = -1; $write = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($call -lt 0 -and $lines[$i] -match '^\s*assertToolPathsNotRealWorkspace\b') { $call = $i }
            if ($write -lt 0 -and $lines[$i] -match 'Copy-Item|WriteAll|CreateDirectory') { $write = $i }
        }
        if ($call -lt 0) {
            $found.Add("assertToolPathsNotRealWorkspace を呼ぶ行が無い")
        } elseif ($write -ge 0 -and $write -lt $call) {
            $found.Add("最初の書き込みが assertToolPathsNotRealWorkspace より前にある")
        }
        return @($found)
    }

    function getUnlistedLibReaders {
        # lib.ps1 を読む（resolveTebunkoLib を呼ぶ）道具で、守りを持つ一覧にも除外の一覧にも載っていないものの名前の配列。
        # $files は @{ Name; Text }（Name はリポジトリからの相対パス。/ 区切り）、$guarded は守りを持つ相対パス、$excluded は 相対パス → 理由
        param ([object[]]$files, [string[]]$guarded, $excluded)

        $result = New-Object System.Collections.Generic.List[string]
        foreach ($file in $files) {
            if ($guarded -contains $file.Name -or $excluded.Contains($file.Name)) { continue }
            if (@(getToolCodeLines $file.Text) -match 'lib\.ps1|resolveTebunkoLib') { $result.Add($file.Name) }
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

    It "load.ps1 は、lib.ps1 を読む前に差し替え、initWorkspace のあとに利用者の既定のワークスペースでないことを確かめる" {
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

    It "画面のテストの起動は、起動の前に利用者の既定のワークスペースでないことを確かめ、環境変数を渡す" {
        $text = [System.IO.File]::ReadAllText("$here\gui\gui_helpers.ps1")
        $text | Should -Match 'assertNotRealWorkspace \$Tool\.DefaultWorkspace'
        $text | Should -Match '\$env:TEBUNKO_DEFAULT_WORKSPACE\s*=\s*\$Tool\.DefaultWorkspace'
        $code = @($text -split "\r?\n" | Where-Object { $_ -notmatch '^\s*#' })
        ($code -match 'getRealDefaultWorkspace').Count | Should -BeGreaterThan 0   # コメントではなく、コードで利用者の既定のワークスペースを控えている
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

    It "性能・変換の道具は、lib.ps1 を読む前に使い捨てを入れ、終わりに前後を比べ、違いがあれば失敗にする: <File>" -TestCases @(
        @{ File = "tools/measure_perf.ps1" }
        @{ File = "tools/perf/measure_ingest.ps1" }
        @{ File = "tools/perf/measure_search.ps1" }
        @{ File = "tools/convert_to_content_index.ps1" }
    ) {
        $text = [System.IO.File]::ReadAllText("$script:repo\$($File.Replace('/', '\'))")
        (getToolGuardOrderViolations $text) -join ", " | Should -Be ""
    }

    It "見本作りの道具は、書く前に書き込み先を確かめる" {
        $text = [System.IO.File]::ReadAllText("$script:repo\tools\make_index_golden.ps1")
        (getToolPathCheckViolations $text) -join ", " | Should -Be ""
    }

    It "5 本の道具は assertToolPathsNotRealWorkspace を呼ぶ: <File>" -TestCases @(
        @{ File = "tools/measure_perf.ps1" }
        @{ File = "tools/perf/measure_ingest.ps1" }
        @{ File = "tools/perf/measure_search.ps1" }
        @{ File = "tools/convert_to_content_index.ps1" }
        @{ File = "tools/make_index_golden.ps1" }
    ) {
        $text = [System.IO.File]::ReadAllText("$script:repo\$($File.Replace('/', '\'))")
        @($text -split "\r?\n" | Where-Object { $_ -match '^\s*assertToolPathsNotRealWorkspace\b' }).Count | Should -BeGreaterThan 0
        (getToolPathCheckViolations $text) -join ", " | Should -Be ""
    }

    It "守りの順の検査は、違反を入れると落ちる: <Case>" -TestCases @(
        @{ Case = "差し替えより前に lib.ps1 を読む"; Text = ". (`$lib)`n`$g = startWorkspaceGuard`n`$r = finishWorkspaceGuard `$g`nif (`$guardReport.ExitCode -ne 0) { throw 'x' }"; Expect = "lib.ps1" }
        @{ Case = "finishWorkspaceGuard が無い"; Text = "`$g = startWorkspaceGuard`nif (`$guardReport.ExitCode -ne 0) { throw 'x' }"; Expect = "finishWorkspaceGuard" }
        @{ Case = "守りの順どおり"; Text = "`$g = startWorkspaceGuard`n. (`$lib)`n`$r = finishWorkspaceGuard `$g`nif (`$guardReport.ExitCode -ne 0) { throw 'x' }"; Expect = "" }
    ) {
        $v = (getToolGuardOrderViolations $Text) -join ", "
        if ($Expect) { $v | Should -BeLike "*$Expect*" } else { $v | Should -Be "" }
    }

    It "書き込み先の確かめの検査は、違反を入れると落ちる: <Case>" -TestCases @(
        @{ Case = "書いた後で確かめる"; Text = "[void][System.IO.Directory]::CreateDirectory(`$a)`nassertToolPathsNotRealWorkspace @{ A = `$a }"; Expect = "最初の書き込み" }
        @{ Case = "呼ばない"; Text = "Copy-Item a b"; Expect = "呼ぶ行が無い" }
        @{ Case = "書く前に確かめる"; Text = "assertToolPathsNotRealWorkspace @{ A = `$a }`nCopy-Item a b"; Expect = "" }
    ) {
        $v = (getToolPathCheckViolations $Text) -join ", "
        if ($Expect) { $v | Should -BeLike "*$Expect*" } else { $v | Should -Be "" }
    }

    It "lib.ps1 を読む tools/ のスクリプトは、守りを持つ一覧か除外の一覧（理由つき）のどちらかに載っている" {
        $files = @(Get-ChildItem -LiteralPath "$script:repo\tools" -Recurse -Filter *.ps1 -File |
            ForEach-Object { @{ Name = $_.FullName.Substring($script:repo.Length + 1).Replace("\", "/"); Text = [System.IO.File]::ReadAllText($_.FullName) } })
        $guarded = @("tools/measure_perf.ps1", "tools/perf/measure_ingest.ps1", "tools/perf/measure_search.ps1", "tools/convert_to_content_index.ps1")
        $childReason = "measure_perf.ps1 の子として動く（親が差し替えた環境変数を引き継ぐ）。単独で動かしたときの守りは、必要なら別に入れる"
        $excluded = [ordered]@{
            "tools/new_single_script.ps1"    = "単一 .ps1 版を組み立てる。lib.ps1 は文字列として読むだけで、実行しない"
            "tools/run_isolated.ps1"         = "自分で使い捨てのワークスペースを作って差し替える"
            "tools/perf/perf_common.ps1"     = "関数だけのファイル。lib.ps1 の場所を探すだけで、実行しない"
            "tools/perf/measure_xlsx.ps1"    = $childReason
            "tools/perf/measure_preview.ps1" = $childReason
            "tools/perf/measure_typing.ps1"  = $childReason
            "tools/perf/measure_index.ps1"   = $childReason
            "tools/perf/measure_export.ps1"  = $childReason
        }
        (getUnlistedLibReaders $files $guarded $excluded) -join ", " | Should -Be ""
    }

    It "一覧の検査は、一覧に無い lib.ps1 を読む道具を、ファイル名を出して落とす: <Case>" -TestCases @(
        @{ Case = "一覧に無い"; Text = 'resolveTebunkoLib $Tool'; Guarded = @(); Excluded = @{}; Expect = "tools/x.ps1" }
        @{ Case = "守りの一覧にある"; Text = 'resolveTebunkoLib $Tool'; Guarded = @("tools/x.ps1"); Excluded = @{}; Expect = "" }
        @{ Case = "除外の一覧にある"; Text = 'resolveTebunkoLib $Tool'; Guarded = @(); Excluded = @{ "tools/x.ps1" = "理由" }; Expect = "" }
        @{ Case = "lib.ps1 を読まない"; Text = '$a = 1'; Guarded = @(); Excluded = @{}; Expect = "" }
        @{ Case = "コメントだけ"; Text = '# lib.ps1 の説明'; Guarded = @(); Excluded = @{}; Expect = "" }
    ) {
        (getUnlistedLibReaders @(@{ Name = "tools/x.ps1"; Text = $Text }) $Guarded $Excluded) -join ", " | Should -Be $Expect
    }

    It "検査は、許したファイルなら通す" {
        @(getBareWorkspaceUses @(@{ Name = "tests/helpers/load.ps1"; Lines = @("initWorkspace") }) @("tests/helpers/load.ps1")).Count | Should -Be 0
    }
}
