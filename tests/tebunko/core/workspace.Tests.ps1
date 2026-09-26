# ワークスペースの中の場所（tebunko\core\workspace.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "Workspace" -Tag Unit {
    It "中の場所をワークスペースのフォルダから組み立てる（出力用のフォルダはプロセスごと）" {
        $target = [Workspace]::new("C:\Users\test\Documents\tebunko_ws")

        $target.Dir | Should -Be "C:\Users\test\Documents\tebunko_ws"
        $target.IndexDir | Should -Be "C:\Users\test\Documents\tebunko_ws\index"
        $target.SystemIndexDir | Should -Be "C:\Users\test\Documents\tebunko_ws\system_index"
        $target.SystemIndexStateFile | Should -Be "C:\Users\test\Documents\tebunko_ws\システムインデックスの状態.tsv"
        $target.PublishDir | Should -Be "C:\Users\test\Documents\tebunko_ws\取り込み出力\$PID"
        $target.StatusFile | Should -Be "C:\Users\test\Documents\tebunko_ws\取り込み一覧.tsv"
        $target.IngestingFile | Should -Be "C:\Users\test\Documents\tebunko_ws\取り込み中.txt"
        $target.ResultFile | Should -Be "C:\Users\test\Documents\tebunko_ws\検索結果.txt"
        $target.IndexingLogFile | Should -Be "C:\Users\test\Documents\tebunko_ws\インデックス作成ログ.txt"
        $target.GuiErrorLogFile | Should -Be "C:\Users\test\Documents\tebunko_ws\画面エラー.txt"
    }

    It "関数は呼んだときの `$workspace の場所を使う（差し替えれば、読み込み直さずに別のワークスペースを使う）" {
        $workspace = newTestWorkspace @{} "$TestDrive\別"

        @((getIndexSummary).Missing) | Should -Be @("$TestDrive\別\index")
    }
}

Describe "moveWorkspace" -Tag Io {
    BeforeAll {
        function newWorkspaceFiles([string]$dir) {
            # tebunko のファイル・フォルダと、利用者のファイル（移さない）を置く
            newTsv "$dir\index\営業\見積\content.xlsx.001.tsv" @("a")
            newTsv "$dir\system_index\営業\見積\システムインデックス.txt" @("b")
            newTsv "$dir\取り込み一覧.tsv" @("c")
            newTsv "$dir\インデックス作成ログ.txt" @("d")
            newTsv "$dir\利用者のメモ.txt" @("e")
        }
    }

    It "tebunko のファイル・フォルダだけを移し、移した数を返す（利用者のファイルは残す）" {
        newWorkspaceFiles "$TestDrive\move\from"

        moveWorkspace "$TestDrive\move\from" "$TestDrive\move\to" | Should -Be 4

        Test-Path -LiteralPath "$TestDrive\move\to\index\営業\見積\content.xlsx.001.tsv" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\move\to\system_index\営業\見積\システムインデックス.txt" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\move\to\取り込み一覧.tsv" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\move\to\インデックス作成ログ.txt" | Should -Be $true
        @(getWorkspaceEntries "$TestDrive\move\from").Count | Should -Be 0
        Test-Path -LiteralPath "$TestDrive\move\from\利用者のメモ.txt" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\move\to\利用者のメモ.txt" | Should -Be $false
    }

    It "移し先に同じ名前があれば、何も移さずに例外にする" {
        newWorkspaceFiles "$TestDrive\conflict\from"
        newTsv "$TestDrive\conflict\to\取り込み一覧.tsv" @("別のワークスペース")

        { moveWorkspace "$TestDrive\conflict\from" "$TestDrive\conflict\to" } | Should -Throw "*すでに 取り込み一覧.tsv があります*"
        @(getWorkspaceEntries "$TestDrive\conflict\from").Count | Should -Be 4
        Test-Path -LiteralPath "$TestDrive\conflict\to\index" | Should -Be $false
    }

    It "途中で移せなければ、移した分を戻して例外にする（中身が 2 つに分かれない）" {
        newWorkspaceFiles "$TestDrive\rollback\from"
        # system_index の中のファイルを開いておき、フォルダを移せなくする（index は先に移る）
        $stream = [System.IO.File]::Open("$TestDrive\rollback\from\system_index\営業\見積\システムインデックス.txt", "Open", "Read", "None")
        try {
            { moveWorkspace "$TestDrive\rollback\from" "$TestDrive\rollback\to" } | Should -Throw "*中身は「$TestDrive\rollback\from」に残しています*"
        } finally {
            $stream.Dispose()
        }
        @(getWorkspaceEntries "$TestDrive\rollback\from").Count | Should -Be 4
        @(getWorkspaceEntries "$TestDrive\rollback\to").Count | Should -Be 0
    }

    It "中身が無ければ何も移さない" {
        moveWorkspace "$TestDrive\empty\from" "$TestDrive\empty\to" | Should -Be 0
    }
}

Describe "copyDirectoryTree" -Tag Io {
    It "フォルダを中身ごと写す（別のドライブへ移すとき）" {
        newTsv "$TestDrive\copy\from\a\b.tsv" @("b")
        newTsv "$TestDrive\copy\from\c.txt" @("c")

        copyDirectoryTree (toLongPath "$TestDrive\copy\from") (toLongPath "$TestDrive\copy\to")

        Test-Path -LiteralPath "$TestDrive\copy\to\a\b.tsv" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\copy\to\c.txt" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\copy\from\c.txt" | Should -Be $true
    }
}

Describe "moveSearchExcludes" -Tag Io {
    It "前のワークスペースのインデックスの下を、移した先のインデックスの下に付け替える（ほかは変えない）" {
        $settings = "$TestDrive\excludes\setting.config"
        writeSearchExcludes @(
            [pscustomobject]@{ Path = "C:\old\index"; Subfolders = $true },
            [pscustomobject]@{ Path = "C:\OLD\index\営業\見積"; Subfolders = $false },
            [pscustomobject]@{ Path = "C:\old\index2\技術"; Subfolders = $true }
        ) $settings

        moveSearchExcludes "C:\old" "D:\新しい場所" $settings | Should -Be 2

        $excludes = @(readSearchExcludes $settings)
        @($excludes | ForEach-Object { $_.Path }) -join "|" | Should -Be "D:\新しい場所\index|D:\新しい場所\index\営業\見積|C:\old\index2\技術"
        @($excludes | ForEach-Object { $_.Subfolders }) -join "," | Should -Be "True,False,True"
    }
}

Describe "removeWorkspaceEntries" -Tag Io {
    It "tebunko のファイル・フォルダだけを削除し、削除した数を返す（利用者のファイルは残す）" {
        newTsv "$TestDrive\remove_ws\index\営業\見積\content.xlsx.001.tsv" @("a")
        newTsv "$TestDrive\remove_ws\取り込み一覧.tsv" @("b")
        newTsv "$TestDrive\remove_ws\利用者のメモ.txt" @("c")

        removeWorkspaceEntries "$TestDrive\remove_ws" | Should -Be 2

        @(getWorkspaceEntries "$TestDrive\remove_ws").Count | Should -Be 0
        Test-Path -LiteralPath "$TestDrive\remove_ws\利用者のメモ.txt" | Should -Be $true
    }
}

Describe "useWorkspaceTargets" -Tag Io {
    It "ワークスペースの取り込み一覧にあるクロール対象フォルダを、インデックスの一覧にする" {
        $settings = "$TestDrive\use\setting.config"
        writeTargetFolders @([pscustomobject]@{ Name = "自分"; Path = "C:\自分のフォルダ"; Enabled = $true }) $settings
        writeStatusFile @(
            [pscustomobject]@{ Path = "\\server\共有\営業部"; Name = "営業" },
            [pscustomobject]@{ Path = "\\server\共有\技術部"; Name = "技術" }
        ) @() "$TestDrive\use\ws\取り込み一覧.tsv"

        useWorkspaceTargets "$TestDrive\use\ws" $settings | Should -Be 2

        $targets = @(getTargetFolders $settings)
        @($targets | ForEach-Object { "$($_.Name)=$($_.Path)=$($_.Enabled)" }) -join "|" | Should -Be "営業=\\server\共有\営業部=True|技術=\\server\共有\技術部=True"
    }

    It "取り込み一覧が無ければ、一覧は変えない" {
        $settings = "$TestDrive\use_none\setting.config"
        writeTargetFolders @([pscustomobject]@{ Name = "自分"; Path = "C:\自分のフォルダ"; Enabled = $true }) $settings

        useWorkspaceTargets "$TestDrive\use_none\ws" $settings | Should -Be 0

        @(getTargetFolders $settings)[0].Name | Should -Be "自分"
    }
}
