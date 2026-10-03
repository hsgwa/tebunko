# Windows Search への問い合わせ（tebunko\search\windows_search.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "invokeWindowsSearch" -Tag Unit {
    It "SELECT 以外は送らない" {
        { invokeWindowsSearch $null "DELETE FROM SystemIndex" } | Should -Throw -ExpectedMessage "*SELECT だけ*"
        { invokeWindowsSearch $null "  update x" } | Should -Throw -ExpectedMessage "*SELECT だけ*"
    }
}

Describe "testWindowsSearch" -Tag Io {
    It "system_index が無ければ使えない" {
        testWindowsSearch "$TestDrive\無い\system_index" | Should -Be $false
    }

    It "Windows Search を開けないときは使えない" {
        Mock openWindowsSearch { $null }
        [System.IO.Directory]::CreateDirectory("$TestDrive\ws\system_index") | Out-Null
        testWindowsSearch "$TestDrive\ws\system_index" | Should -Be $false
        testTsvIndexedByWindowsSearch "$TestDrive\ws\index" | Should -Be $false
    }

    It "問い合わせに失敗したら使えない（例外にしない）" {
        Mock openWindowsSearch { [pscustomobject]@{} | Add-Member -MemberType ScriptMethod -Name Dispose -Value {} -PassThru }
        Mock invokeWindowsSearch { throw "失敗" }
        [System.IO.Directory]::CreateDirectory("$TestDrive\ws2\system_index") | Out-Null
        testWindowsSearch "$TestDrive\ws2\system_index" | Should -Be $false
        testTsvIndexedByWindowsSearch "$TestDrive\ws2\index" | Should -Be $false
    }

    It "行が返れば使える" {
        Mock openWindowsSearch { [pscustomobject]@{} | Add-Member -MemberType ScriptMethod -Name Dispose -Value {} -PassThru }
        Mock invokeWindowsSearch { , (New-Object 'System.Collections.Generic.List[object[]]' (, [object[][]]@(, @("file:C:/x")))) }
        [System.IO.Directory]::CreateDirectory("$TestDrive\ws3\system_index") | Out-Null
        testWindowsSearch "$TestDrive\ws3\system_index" | Should -Be $true
        testTsvIndexedByWindowsSearch "$TestDrive\ws3\index" | Should -Be $true
    }
}

Describe "getWindowsSearchState" -Tag Io {
    BeforeAll {
        # system_index への問い合わせ（SCOPE が system_index で終わる）とワークスペースへの問い合わせで、返す行の有無を変える
        function script:mockWindowsSearch {
            param ([bool]$systemHits, [bool]$workspaceHits)
            $script:systemHits = $systemHits
            $script:workspaceHits = $workspaceHits
            Mock openWindowsSearch { [pscustomobject]@{} | Add-Member -MemberType ScriptMethod -Name Dispose -Value {} -PassThru }
            Mock invokeWindowsSearch {
                $rows = New-Object 'System.Collections.Generic.List[object[]]'
                if ($script:systemHits) { $rows.Add(@("file:C:/x")) }
                , $rows
            } -ParameterFilter { $sql -like "*system_index'" }
            Mock invokeWindowsSearch {
                $rows = New-Object 'System.Collections.Generic.List[object[]]'
                if ($script:workspaceHits) { $rows.Add(@("file:C:/x")) }
                , $rows
            } -ParameterFilter { $sql -notlike "*system_index'" }
        }
    }

    It "<name>" -TestCases @(
        @{ name = "system_index が索引されていれば Ok"; system = $true; workspace = $true; expected = "Ok" }
        @{ name = "system_index が索引されていれば、ワークスペースの様子に関わらず Ok"; system = $true; workspace = $false; expected = "Ok" }
        @{ name = "system_index が 0 件でも、ワークスペースのほかのものが索引されていれば NotYet（対象だが、まだ索引していない）"; system = $false; workspace = $true; expected = "NotYet" }
        @{ name = "system_index もワークスペースも 0 件なら NotInScope（索引の対象外）"; system = $false; workspace = $false; expected = "NotInScope" }
    ) {
        param ($name, $system, $workspace, $expected)
        mockWindowsSearch $system $workspace
        $root = "$TestDrive\state_$expected$system$workspace\ws"
        [System.IO.Directory]::CreateDirectory("$root\system_index") | Out-Null
        getWindowsSearchState "$root\system_index" | Should -Be $expected
        getWindowsSearchState "$root\system_index" $root | Should -Be $expected
        # ワークスペースは systemRoot の 1 つ上（省略時）
        if ($expected -ne "Ok") {
            Should -Invoke invokeWindowsSearch -Times 2 -Exactly -ParameterFilter { $sql -like "*state_*/ws'" }
        }
    }

    It "system_index のフォルダが無ければ NoFolder（Windows Search には問い合わせない）" {
        mockWindowsSearch $true $true
        getWindowsSearchState "$TestDrive\無い\system_index" | Should -Be "NoFolder"
        Should -Invoke openWindowsSearch -Times 0 -Exactly
    }

    It "Windows Search を開けない・問い合わせに失敗したら NoConnection（例外にしない）" {
        [System.IO.Directory]::CreateDirectory("$TestDrive\conn\system_index") | Out-Null
        Mock openWindowsSearch { $null }
        getWindowsSearchState "$TestDrive\conn\system_index" | Should -Be "NoConnection"
        Mock openWindowsSearch { [pscustomobject]@{} | Add-Member -MemberType ScriptMethod -Name Dispose -Value {} -PassThru }
        Mock invokeWindowsSearch { throw "時間切れ" }
        getWindowsSearchState "$TestDrive\conn\system_index" | Should -Be "NoConnection"
    }

    It "開いた接続を渡せば、それを使い、閉じない" {
        [System.IO.Directory]::CreateDirectory("$TestDrive\shared\system_index") | Out-Null
        mockWindowsSearch $true $true
        $script:disposed = 0
        $connection = [pscustomobject]@{} | Add-Member -MemberType ScriptMethod -Name Dispose -Value { $script:disposed++ } -PassThru
        getWindowsSearchState "$TestDrive\shared\system_index" "" $connection | Should -Be "Ok"
        testTsvIndexedByWindowsSearch "$TestDrive\shared\index" $connection | Should -Be $true
        $script:disposed | Should -Be 0
        Should -Invoke openWindowsSearch -Times 0 -Exactly
    }
}

Describe "Windows Search（本物）" -Tag Io {
    BeforeAll {
        $connection = openWindowsSearch
    }

    It "開けるなら、SELECT の結果を行の一覧で返す（Windows Search が無い環境では確かめない）" {
        if ($null -eq $connection) {
            Set-ItResult -Inconclusive -Because "Windows Search を開けない環境"
            return
        }
        try {
            try {
                $rows = invokeWindowsSearch $connection "SELECT TOP 1 System.ItemUrl FROM SystemIndex"
            } catch {
                # CI のランナーなど、開けても問い合わせに答えない環境（E_FAIL）がある
                Set-ItResult -Inconclusive -Because "Windows Search が問い合わせに答えない環境: $($_.Exception.Message)"
                return
            }
            $rows.GetType().Name | Should -Be 'List`1'
        } finally {
            $connection.Dispose()
        }
    }
}
