# ファイルの変更の見張り（shared\core\file_watch.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    function waitSignaled {
        param ($Watch, [int]$Seconds = 5)
        $deadline = (Get-Date).AddSeconds($Seconds)
        while ((Get-Date) -lt $deadline) {
            if (testFileWatchSignaled $Watch) { return $true }
            Start-Sleep -Milliseconds 100
        }
        return $false
    }
}

Describe "newFileWatch・testFileWatchSignaled・closeFileWatch" -Tag Io {
    BeforeEach {
        $script:dir = Join-Path $TestDrive ("watch-" + [guid]::NewGuid().ToString("N"))
        New-Item -ItemType Directory -Path $script:dir | Out-Null
        $script:file = Join-Path $script:dir "setting.config"
        Set-Content -LiteralPath $script:file -Value "a"
        $script:watch = newFileWatch $script:file
    }
    AfterEach {
        closeFileWatch $script:watch
    }

    It "書き込みで合図が来て、読むと消える" {
        $script:watch | Should -Not -BeNullOrEmpty
        Set-Content -LiteralPath $script:file -Value "bb"
        (waitSignaled $script:watch) | Should -BeTrue
        Start-Sleep -Milliseconds 300
        [void](testFileWatchSignaled $script:watch)
        (testFileWatchSignaled $script:watch) | Should -BeFalse
    }

    It "一時ファイルからの置き換えでも合図が来る" {
        $tmp = "$script:file.tmp"
        Set-Content -LiteralPath $tmp -Value "ccc"
        Move-Item -LiteralPath $tmp -Destination $script:file -Force
        (waitSignaled $script:watch) | Should -BeTrue
    }

    It "別のファイルの変更では合図が来ない" {
        Set-Content -LiteralPath (Join-Path $script:dir "other.txt") -Value "x"
        Start-Sleep -Milliseconds 500
        (testFileWatchSignaled $script:watch) | Should -BeFalse
    }

    It "閉じると購読が残らず、2 回閉じても失敗しない" {
        closeFileWatch $script:watch
        { closeFileWatch $script:watch } | Should -Not -Throw
        @(Get-EventSubscriber | Where-Object { $_.SourceIdentifier -like "$($script:watch.SourceId)*" }).Count | Should -Be 0
    }
}

Describe "newFileWatch が見張らない場合" -Tag Io {
    It "<name>は `$null を返す" -TestCases @(
        @{ name = "ネットワークのパス"; path = "\\server\share\setting.config" }
        @{ name = "無いフォルダ"; path = "Z:\no_such_folder_for_watch\a\setting.config" }
        @{ name = "空のパス"; path = "" }
    ) {
        newFileWatch $path | Should -BeNullOrEmpty
    }

    It "`$null の見張りは、合図なし・閉じても失敗しない" {
        (testFileWatchSignaled $null) | Should -BeFalse
        { closeFileWatch $null } | Should -Not -Throw
    }
}

Describe "getFileStamp" -Tag Io {
    It "書き込みで変わり、同じ内容のままなら変わらない" {
        $file = Join-Path $TestDrive "stamp.txt"
        Set-Content -LiteralPath $file -Value "a"
        $before = getFileStamp $file
        getFileStamp $file | Should -Be $before
        Set-Content -LiteralPath $file -Value "abcdef"
        getFileStamp $file | Should -Not -Be $before
    }

    It "無いファイルは空" {
        getFileStamp (Join-Path $TestDrive "none.txt") | Should -Be ""
    }
}
