# テストと実機の確かめが、利用者の本物の既定のワークスペース（Documents\tebunko_ws）に書かない仕組みが崩れていないかを調べる。
# 事故の再発防止: 既定のワークスペースを差し替えずに initWorkspace を呼ぶ・既定のワークスペースを使う画面を開く、を機械的に止める。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    $script:repo = (Resolve-Path "$here\..").Path

    function getBareWorkspaceUses {
        # 既定のワークスペースを差し替えずに使う書き方の行を "ファイル:行" の一覧で返す。
        # $files は @{ Name; Lines } の配列、$allowed は書いてよいファイル名
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

    It "run.ps1 は、Pester の前に使い捨てを入れ、後に本物の前後を比べて、違いがあれば失敗にする" {
        $text = [System.IO.File]::ReadAllText("$here\run.ps1")
        $start = $text.IndexOf("startWorkspaceGuard")
        $pester = $text.IndexOf("Invoke-Pester -Configuration")
        $stop = $text.IndexOf("stopWorkspaceGuard")
        $start | Should -BeGreaterThan -1
        $start | Should -BeLessThan $pester
        $stop | Should -BeGreaterThan $pester
        $text | Should -Match '(?s)stopWorkspaceGuard.*exit 1'
    }

    It "capture_screens.ps1 は、使い捨てを入れてから撮り、終わりに本物の前後を比べる" {
        $text = [System.IO.File]::ReadAllText("$script:repo\tools\capture_screens.ps1")
        $text | Should -Match 'startWorkspaceGuard'
        $text | Should -Match 'stopWorkspaceGuard'
    }

    It "画面のテストの起動は、起動の前に本物でないことを確かめ、環境変数を渡す" {
        $text = [System.IO.File]::ReadAllText("$here\gui\gui_helpers.ps1")
        $text | Should -Match 'assertNotRealWorkspace \$Tool\.DefaultWorkspace'
        $text | Should -Match '\$env:TEBUNKO_DEFAULT_WORKSPACE\s*='
        $text | Should -Match '既定のワークスペース'
    }

    It "取り込みの起動の前に、書き込み先が使い捨ての中であることを確かめる（runIndexer）" {
        $text = [System.IO.File]::ReadAllText("$here\helpers\indexer.ps1")
        $text | Should -Match 'assertIndexerWorkspaceIsolated'
    }

    It "既定のワークスペースを差し替えずに使う書き方が、決まったファイルの外に無い" {
        $files = @(Get-ChildItem -LiteralPath $here, "$script:repo\tools" -Recurse -Filter *.ps1 -File |
            Where-Object { $_.Name -ne "test_isolation.Tests.ps1" } |
            ForEach-Object { @{ Name = $_.Name; Lines = [System.IO.File]::ReadAllLines($_.FullName) } })
        # load.ps1 は差し替えのあとで呼ぶ。new_single_script・parts のテストは文字列として扱うか、環境変数を入れた子で呼ぶ。
        # settings.Tests.ps1（画面）の S6 は gui_helpers が環境変数を渡す
        $allowed = @("load.ps1", "new_single_script.Tests.ps1", "parts.Tests.ps1", "settings.Tests.ps1")
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

    It "検査は、許したファイルなら通す" {
        @(getBareWorkspaceUses @(@{ Name = "load.ps1"; Lines = @("initWorkspace") }) @("load.ps1")).Count | Should -Be 0
    }
}

Describe "assertIndexerWorkspaceIsolated（indexer.ps1 を動かす前の確かめ）" -Tag Io {
    BeforeAll {
        . "$here\helpers\indexer.ps1"
    }

    It "<Case>" -TestCases @(
        @{ Case = "テスト用の置き場所の中の work は通る"; Folder = "work"; Throws = $false }
        @{ Case = "空（既定のワークスペース）は、差し替えた場所なら通る"; Folder = ""; Throws = $false }
        @{ Case = "置き場所の外は止める"; Folder = "..\other"; Throws = $true }
        @{ Case = "本物の場所は止める"; Folder = "REAL"; Throws = $true }
    ) {
        $root = Join-Path $TestDrive "iso-$([guid]::NewGuid().ToString('N').Substring(0, 6))"
        New-Item -ItemType Directory -Path $root | Out-Null
        $real = Join-Path $TestDrive "real-ws"
        $default = Join-Path $TestDrive "default-ws"
        $value = if ($Folder -eq "REAL") { $real } else { $Folder }
        $settings = newSettings
        $settings.workspaceFolder = $value
        writeSettings $settings "$root\setting.config"
        if ($Throws) {
            { assertIndexerWorkspaceIsolated $root $real $default } | Should -Throw
        } else {
            { assertIndexerWorkspaceIsolated $root $real $default } | Should -Not -Throw
        }
    }

    It "空で、既定のワークスペースが本物を指しているときは止める" {
        $root = Join-Path $TestDrive "iso-real"
        New-Item -ItemType Directory -Path $root | Out-Null
        $real = Join-Path $TestDrive "real-ws2"
        $settings = newSettings
        $settings.workspaceFolder = ""
        writeSettings $settings "$root\setting.config"
        { assertIndexerWorkspaceIsolated $root $real $real } | Should -Throw
    }
}
