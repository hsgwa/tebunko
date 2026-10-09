# 取り込み（indexer.ps1）を動かす前に、書き込み先が使い捨ての中であることを確かめる関数（tests\helpers\indexer.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "assertIndexerWorkspaceIsolated（indexer.ps1 を動かす前の確かめ）" -Tag Io {
    BeforeAll {
        . "$PSScriptRoot\..\..\helpers\indexer.ps1"
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
