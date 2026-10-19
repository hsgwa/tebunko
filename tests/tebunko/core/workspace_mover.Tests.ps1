# ワークスペースの切り替え（tebunko\core\workspace_mover.ps1 の WorkspaceMover）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    function script:newWorkspaceFiles([string]$dir) {
        # tebunko のファイル・フォルダと、利用者のファイル（移さない）を置く。
        # content_index（今の版）と、前の版の残骸の index の両方を置いて、どちらも移す・消すことを確かめる
        newTsv "$dir\content_index\営業\見積\content_index.xlsx.001.tsv" @("a")
        newTsv "$dir\index\営業\見積\content.xlsx.001.tsv" @("a-前の版")
        newTsv "$dir\ingest_status.tsv" @("c")
        newTsv "$dir\indexing_log.txt" @("d")
        newTsv "$dir\利用者のメモ.txt" @("e")
    }
}

Describe "WorkspaceMover.MoveEntries" -Tag Io {
    BeforeAll {
        $script:mover = [WorkspaceMover]::new("$TestDrive\setting.config")
    }

    It "tebunko のファイル・フォルダだけを移し、移した数を返す（利用者のファイルは残す）" {
        newWorkspaceFiles "$TestDrive\move\from"

        $result = $script:mover.MoveEntries("$TestDrive\move\from", "$TestDrive\move\to")
        $result.Count | Should -Be 4
        @($result.Leftovers).Count | Should -Be 0

        Test-Path -LiteralPath "$TestDrive\move\to\content_index\営業\見積\content_index.xlsx.001.tsv" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\move\to\index\営業\見積\content.xlsx.001.tsv" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\move\to\ingest_status.tsv" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\move\to\indexing_log.txt" | Should -Be $true
        @($script:mover.ListEntries("$TestDrive\move\from")).Count | Should -Be 0
        Test-Path -LiteralPath "$TestDrive\move\from\利用者のメモ.txt" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\move\to\利用者のメモ.txt" | Should -Be $false
    }

    It "移し先に同じ名前があれば、何も移さずに例外にする" {
        newWorkspaceFiles "$TestDrive\conflict\from"
        newTsv "$TestDrive\conflict\to\ingest_status.tsv" @("別のワークスペース")

        { $script:mover.MoveEntries("$TestDrive\conflict\from", "$TestDrive\conflict\to") } | Should -Throw "*すでに ingest_status.tsv があります*"
        @($script:mover.ListEntries("$TestDrive\conflict\from")).Count | Should -Be 4
        Test-Path -LiteralPath "$TestDrive\conflict\to\content_index" | Should -Be $false
    }

    It "一時的に掴まれていても、放されれば移せる（待ちながら試し直す）" {
        newWorkspaceFiles "$TestDrive\grab\from"
        # content_index の中のファイルを開いておき、1 回目の待ちで放す
        $script:grabbed = [System.IO.File]::Open("$TestDrive\grab\from\content_index\営業\見積\content_index.xlsx.001.tsv", "Open", "Read", "None")
        Mock Start-Sleep { if ($script:grabbed) { $script:grabbed.Dispose(); $script:grabbed = $null } }
        try {
            $result = $script:mover.MoveEntries("$TestDrive\grab\from", "$TestDrive\grab\to", 8, 100)
        } finally {
            if ($script:grabbed) { $script:grabbed.Dispose(); $script:grabbed = $null }
        }
        Should -Invoke Start-Sleep -Times 1 -Exactly -Scope It

        $result.Count | Should -Be 4
        Test-Path -LiteralPath "$TestDrive\grab\to\content_index\営業\見積\content_index.xlsx.001.tsv" | Should -Be $true
        @($script:mover.ListEntries("$TestDrive\grab\from")).Count | Should -Be 0
    }

    It "掴まれたままなら、決めた回数だけ試したあとに移した分を戻して例外にする（中身が 2 つに分かれない）" {
        newWorkspaceFiles "$TestDrive\rollback\from"
        # content_index の中のファイルを開いておき、フォルダを移せなくする（先に並ぶ index・ingest_status.tsv は先に移る）
        $stream = [System.IO.File]::Open("$TestDrive\rollback\from\content_index\営業\見積\content_index.xlsx.001.tsv", "Open", "Read", "None")
        try {
            { $script:mover.MoveEntries("$TestDrive\rollback\from", "$TestDrive\rollback\to", 3, 1) } | Should -Throw "*移せませんでした*中身は「$TestDrive\rollback\from」に残しています*"
        } finally {
            $stream.Dispose()
        }
        @($script:mover.ListEntries("$TestDrive\rollback\from")).Count | Should -Be 4
        @($script:mover.ListEntries("$TestDrive\rollback\to")).Count | Should -Be 0
    }

    It "戻せなかった項目があるときは、「移せませんでした」に続けてその項目と場所を書き、「残しています」と書かない" {
        newWorkspaceFiles "$TestDrive\norestore\from"
        # 先に移した content_index は、戻すときに掴まれて戻せない。あとに並ぶ index は移せず、ここで失敗にする
        Mock moveDirectoryRetry {
            if ($destination -like "*\norestore\from\*" -or $source -like "*\index") { throw "掴まれています" }
            [System.IO.Directory]::Move((toLongPath $source), (toLongPath $destination))
        }
        $caught = $null
        try { $script:mover.MoveEntries("$TestDrive\norestore\from", "$TestDrive\norestore\to", 2, 1) } catch { $caught = $_.Exception.Message }

        $caught | Should -BeLike "*移せませんでした*元へ戻せなかった項目があります（content_index）。それは「$TestDrive\norestore\to」にあります*"
        $caught | Should -Not -BeLike "*残しています*"
        Test-Path -LiteralPath "$TestDrive\norestore\to\content_index" | Should -Be $true
    }

    It "別のドライブで元を消せないときは、移したものとして扱い、残った元の場所を返す" {
        newWorkspaceFiles "$TestDrive\leftover\from"
        Mock testSameDriveRoot { $false }
        Mock removeDirectoryRetry { throw "消せません" } -ParameterFilter { $path -like "*\leftover\from\*" }

        $result = $script:mover.MoveEntries("$TestDrive\leftover\from", "$TestDrive\leftover\to", 2, 1)

        $result.Count | Should -Be 4
        @($result.Leftovers | ForEach-Object { [System.IO.Path]::GetFileName($_) }) | Should -Contain "content_index"
        Test-Path -LiteralPath "$TestDrive\leftover\to\content_index\営業\見積\content_index.xlsx.001.tsv" | Should -Be $true
    }

    It "中身が無ければ何も移さない" {
        $result = $script:mover.MoveEntries("$TestDrive\empty\from", "$TestDrive\empty\to")
        $result.Count | Should -Be 0
        @($result.Leftovers).Count | Should -Be 0
    }
}

Describe "WorkspaceMover の設定ファイルは、移しても消しても動かない" -Tag Io {
    BeforeEach {
        # 書き込めない場所に置いた tebunko の設定は、既定のワークスペースの直下にある（getSettingsFilePath）
        $script:ws = "$TestDrive\settings_ws-" + [guid]::NewGuid().ToString("N")
        $script:mover = [WorkspaceMover]::new("$TestDrive\setting.config")
        newTsv "$script:ws\content_index\営業\見積\content_index.xlsx.001.tsv" @("a")
        newTsv "$script:ws\ingest_status.tsv" @("b")
        newTsv "$script:ws\setting.config" @("{}")
        newTsv "$script:ws\setting.config.broken-20261003-120000" @("{")
    }

    It "setting.config とそれに付いてできるファイルは、ワークスペースの中身（Entries）に入らない" {
        $names = @($script:mover.ListEntries($script:ws) | ForEach-Object { [System.IO.Path]::GetFileName($_) })
        $names | Where-Object { testSettingsFileName $_ } | Should -BeNullOrEmpty
        $names | Should -Not -Contain "setting.config"
    }

    It "別のフォルダへ移しても、元の場所に残り、移した先には行かない" {
        $script:mover.MoveEntries($script:ws, "$TestDrive\settings_to").Count | Should -Be 2

        Test-Path -LiteralPath "$script:ws\setting.config" | Should -Be $true
        Test-Path -LiteralPath "$script:ws\setting.config.broken-20261003-120000" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\settings_to\setting.config" | Should -Be $false
    }

    It "「消して最初から」（RemoveEntries）でも消えない" {
        $script:mover.RemoveEntries($script:ws) | Should -Be 2

        Test-Path -LiteralPath "$script:ws\setting.config" | Should -Be $true
        Test-Path -LiteralPath "$script:ws\setting.config.broken-20261003-120000" | Should -Be $true
    }
}

Describe "WorkspaceMover の別のドライブへ移すときの写し（CopyTree）" -Tag Io {
    It "別のドライブのとき（ルートが違うとき）は、写してから元を消す" {
        # 別のドライブは用意できないため、写す処理そのもの（hidden の CopyTree）を呼んで確かめる
        newTsv "$TestDrive\copy\from\a\b.tsv" @("b")
        newTsv "$TestDrive\copy\from\c.txt" @("c")
        $mover = [WorkspaceMover]::new("$TestDrive\setting.config")

        $mover.CopyTree((toLongPath "$TestDrive\copy\from"), (toLongPath "$TestDrive\copy\to"))

        Test-Path -LiteralPath "$TestDrive\copy\to\a\b.tsv" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\copy\to\c.txt" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\copy\from\c.txt" | Should -Be $true
    }
}

Describe "WorkspaceMover.RemoveEntries" -Tag Io {
    It "tebunko のファイル・フォルダだけを削除し、削除した数を返す（content_index・前の版の index も含めて消し、利用者のファイルは残す）" {
        newTsv "$TestDrive\remove_ws\content_index\営業\見積\content_index.xlsx.001.tsv" @("a")
        newTsv "$TestDrive\remove_ws\index\営業\見積\content.xlsx.001.tsv" @("a-前の版")
        newTsv "$TestDrive\remove_ws\ingest_status.tsv" @("b")
        newTsv "$TestDrive\remove_ws\利用者のメモ.txt" @("c")
        $mover = [WorkspaceMover]::new("$TestDrive\setting.config")

        $mover.RemoveEntries("$TestDrive\remove_ws") | Should -Be 3

        @($mover.ListEntries("$TestDrive\remove_ws")).Count | Should -Be 0
        Test-Path -LiteralPath "$TestDrive\remove_ws\content_index" | Should -Be $false
        Test-Path -LiteralPath "$TestDrive\remove_ws\index" | Should -Be $false
        Test-Path -LiteralPath "$TestDrive\remove_ws\利用者のメモ.txt" | Should -Be $true
    }
}

Describe "WorkspaceMover.UseTargets" -Tag Io {
    It "ワークスペースの取り込み一覧にあるクロール対象フォルダを、インデックスの一覧にする" {
        $settings = "$TestDrive\use\setting.config"
        writeTargetFolders @([pscustomobject]@{ Name = "自分"; Path = "C:\自分のフォルダ"; Enabled = $true }) $settings
        writeStatusFile @(
            [pscustomobject]@{ Path = "\\server\共有\営業部"; Name = "営業" },
            [pscustomobject]@{ Path = "\\server\共有\技術部"; Name = "技術" }
        ) @() "$TestDrive\use\ws\store\ingest_status.tsv"

        ([WorkspaceMover]::new($settings)).UseTargets("$TestDrive\use\ws") | Should -Be 2

        $targets = @(getTargetFolders $settings)
        @($targets | ForEach-Object { "$($_.Name)=$($_.Path)=$($_.Enabled)" }) -join "|" | Should -Be "営業=\\server\共有\営業部=True|技術=\\server\共有\技術部=True"
    }

    It "前の版のワークスペース（取り込み一覧がワークスペース直下）でも、クロール対象フォルダを一覧にする" {
        $settings = "$TestDrive\use_old\setting.config"
        writeTargetFolders @([pscustomobject]@{ Name = "自分"; Path = "C:\自分のフォルダ"; Enabled = $true }) $settings
        writeStatusFile @(
            [pscustomobject]@{ Path = "\\server\共有\営業部"; Name = "営業" }
        ) @() "$TestDrive\use_old\ws\ingest_status.tsv"

        ([WorkspaceMover]::new($settings)).UseTargets("$TestDrive\use_old\ws") | Should -Be 1

        @(getTargetFolders $settings)[0].Name | Should -Be "営業"
    }

    It "移行の途中（新旧の取り込み一覧の両方があり、名前が一部だけ重なる）は、新しい方を先に、無い名前を前の版から足す" {
        $settings = "$TestDrive\use_both\setting.config"
        writeTargetFolders @([pscustomobject]@{ Name = "自分"; Path = "C:\自分のフォルダ"; Enabled = $true }) $settings
        writeStatusFile @(
            [pscustomobject]@{ Path = "\\server\共有\営業部"; Name = "営業" }
        ) @() "$TestDrive\use_both\ws\store\ingest_status.tsv"
        writeStatusFile @(
            [pscustomobject]@{ Path = "\\old\営業部"; Name = "営業" }
            [pscustomobject]@{ Path = "\\server\共有\技術部"; Name = "技術" }
        ) @() "$TestDrive\use_both\ws\ingest_status.tsv"

        ([WorkspaceMover]::new($settings)).UseTargets("$TestDrive\use_both\ws") | Should -Be 2

        $targets = @(getTargetFolders $settings)
        @($targets | ForEach-Object { "$($_.Name)=$($_.Path)" }) -join "|" | Should -Be "営業=\\server\共有\営業部|技術=\\server\共有\技術部"
    }

    It "取り込み一覧が無ければ、一覧は変えない" {
        $settings = "$TestDrive\use_none\setting.config"
        writeTargetFolders @([pscustomobject]@{ Name = "自分"; Path = "C:\自分のフォルダ"; Enabled = $true }) $settings

        ([WorkspaceMover]::new($settings)).UseTargets("$TestDrive\use_none\ws") | Should -Be 0

        @(getTargetFolders $settings)[0].Name | Should -Be "自分"
    }
}

Describe "WorkspaceMover.ChangeWorkspace（画面が呼ぶ入口）" -Tag Io {
    BeforeEach {
        $script:settings = "$TestDrive\change\setting.config"
        Remove-Item -LiteralPath "$TestDrive\change" -Recurse -Force -ErrorAction SilentlyContinue
        $script:mover = [WorkspaceMover]::new($script:settings)
    }

    It "move: 中身を移し、searchExcludes を付け替え、ワークスペースを設定に保存する" {
        newWorkspaceFiles "$TestDrive\change\from"
        writeSearchExcludes @([pscustomobject]@{ Path = "$TestDrive\change\from\store\content\営業"; Subfolders = $true }) $script:settings

        $result = $script:mover.ChangeWorkspace("move", "$TestDrive\change\from", "$TestDrive\change\to")

        $result.Ok | Should -Be $true
        $result.Count | Should -Be 4
        Test-Path -LiteralPath "$TestDrive\change\to\ingest_status.tsv" | Should -Be $true
        @(readSearchExcludes $script:settings)[0].Path | Should -Be "$TestDrive\change\to\store\content\営業"
        (getWorkDir $script:settings) | Should -Be "$TestDrive\change\to"
    }

    It "move: 移せなければ Ok=false・Stage=move で、設定は変えない（移した分は戻してある）" {
        newWorkspaceFiles "$TestDrive\change\from"
        newTsv "$TestDrive\change\to\ingest_status.tsv" @("別のワークスペース")

        $result = $script:mover.ChangeWorkspace("move", "$TestDrive\change\from", "$TestDrive\change\to")

        $result.Ok | Should -Be $false
        $result.Stage | Should -Be "move"
        $result.Message | Should -Match "すでに ingest_status.tsv があります"
        @($script:mover.ListEntries("$TestDrive\change\from")).Count | Should -Be 4
        Test-Path -LiteralPath $script:settings | Should -Be $false
    }

    It "reset: 選んだ先の tebunko のファイルを消してから移す" {
        newWorkspaceFiles "$TestDrive\change\from"
        newTsv "$TestDrive\change\to\ingest_status.tsv" @("消す")
        newTsv "$TestDrive\change\to\利用者のメモ.txt" @("残す")

        $result = $script:mover.ChangeWorkspace("reset", "$TestDrive\change\from", "$TestDrive\change\to")

        $result.Ok | Should -Be $true
        Get-Content -LiteralPath "$TestDrive\change\to\ingest_status.tsv" | Should -Be "c"
        Test-Path -LiteralPath "$TestDrive\change\to\利用者のメモ.txt" | Should -Be $true
    }

    It "reset: 消せなければ Ok=false・Stage=remove で、何も移さない" {
        newWorkspaceFiles "$TestDrive\change\from"
        newTsv "$TestDrive\change\to\content_index\営業\見積\a.tsv" @("消す")
        $stream = [System.IO.File]::Open("$TestDrive\change\to\content_index\営業\見積\a.tsv", "Open", "Read", "None")
        try {
            $result = $script:mover.ChangeWorkspace("reset", "$TestDrive\change\from", "$TestDrive\change\to")
        } finally {
            $stream.Dispose()
        }

        $result.Ok | Should -Be $false
        $result.Stage | Should -Be "remove"
        $result.Message | Should -Not -BeNullOrEmpty
        @($script:mover.ListEntries("$TestDrive\change\from")).Count | Should -Be 4
        Test-Path -LiteralPath $script:settings | Should -Be $false
    }

    It "use: 中身は移さず、選んだ先の取り込み一覧をインデックスの一覧にして、ワークスペースを保存する" {
        newWorkspaceFiles "$TestDrive\change\from"
        writeStatusFile @([pscustomobject]@{ Path = "\\server\共有\営業部"; Name = "営業" }) @() "$TestDrive\change\to\store\ingest_status.tsv"

        $result = $script:mover.ChangeWorkspace("use", "$TestDrive\change\from", "$TestDrive\change\to")

        $result.Ok | Should -Be $true
        $result.Targets | Should -Be 1
        @($script:mover.ListEntries("$TestDrive\change\from")).Count | Should -Be 4
        @(getTargetFolders $script:settings)[0].Name | Should -Be "営業"
        (getWorkDir $script:settings) | Should -Be "$TestDrive\change\to"
    }
}
