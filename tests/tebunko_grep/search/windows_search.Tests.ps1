# Windows Search への問い合わせ（tebunko_grep\search\windows_search.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"

Describe "invokeWindowsSearch" -Tag Unit {
    It "SELECT 以外は送らない" {
        { invokeWindowsSearch $null "DELETE FROM SystemIndex" } | Should Throw "SELECT だけ"
        { invokeWindowsSearch $null "  update x" } | Should Throw "SELECT だけ"
    }
}

Describe "testWindowsSearch" -Tag Io {
    It "system_index が無ければ使えない" {
        testWindowsSearch "$TestDrive\無い\system_index" | Should Be $false
    }

    It "Windows Search を開けないときは使えない" {
        Mock openWindowsSearch { $null }
        [System.IO.Directory]::CreateDirectory("$TestDrive\ws\system_index") | Out-Null
        testWindowsSearch "$TestDrive\ws\system_index" | Should Be $false
        testTsvIndexedByWindowsSearch "$TestDrive\ws\index" | Should Be $false
    }

    It "問い合わせに失敗したら使えない（例外にしない）" {
        Mock openWindowsSearch { [pscustomobject]@{} | Add-Member -MemberType ScriptMethod -Name Dispose -Value {} -PassThru }
        Mock invokeWindowsSearch { throw "失敗" }
        [System.IO.Directory]::CreateDirectory("$TestDrive\ws2\system_index") | Out-Null
        testWindowsSearch "$TestDrive\ws2\system_index" | Should Be $false
        testTsvIndexedByWindowsSearch "$TestDrive\ws2\index" | Should Be $false
    }

    It "行が返れば使える" {
        Mock openWindowsSearch { [pscustomobject]@{} | Add-Member -MemberType ScriptMethod -Name Dispose -Value {} -PassThru }
        Mock invokeWindowsSearch { , (New-Object 'System.Collections.Generic.List[object[]]' (, [object[][]]@(, @("file:C:/x")))) }
        [System.IO.Directory]::CreateDirectory("$TestDrive\ws3\system_index") | Out-Null
        testWindowsSearch "$TestDrive\ws3\system_index" | Should Be $true
        testTsvIndexedByWindowsSearch "$TestDrive\ws3\index" | Should Be $true
    }
}

Describe "Windows Search（本物）" -Tag Io {
    $connection = openWindowsSearch
    It "開けるなら、SELECT の結果を行の一覧で返す（Windows Search が無い環境では確かめない）" {
        if ($null -eq $connection) {
            Set-TestInconclusive "Windows Search を開けない環境"
            return
        }
        try {
            try {
                $rows = invokeWindowsSearch $connection "SELECT TOP 1 System.ItemUrl FROM SystemIndex"
            } catch {
                # CI のランナーなど、開けても問い合わせに答えない環境（E_FAIL）がある
                Set-TestInconclusive "Windows Search が問い合わせに答えない環境: $($_.Exception.Message)"
                return
            }
            $rows.GetType().Name | Should Be 'List`1'
        } finally {
            $connection.Dispose()
        }
    }
}
