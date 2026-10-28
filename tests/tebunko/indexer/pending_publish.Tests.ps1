# 本文インデックスへの書き出し待ち（tebunko\indexer\pending_publish.ps1 の PendingPublish）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\indexer\indexer_plan.ps1"  # getBookDir
    . "${scriptsDir}\tebunko\indexer\pending_publish.ps1"
}

Describe "PendingPublish" -Tag Unit {
    It "取り込んだファイルのフォルダは、無くなったファイル名を記録せずに書き出し待ちにする" {
        $pending = [PendingPublish]::new()
        $pending.Add("営業\見積.xlsx", $false)
        $flush = $pending.TakeAll()
        $flush.Count | Should -Be 1
        $folder = [System.IO.Path]::GetDirectoryName((getBookDir "営業\見積.xlsx"))
        $flush[$folder].Count | Should -Be 0
    }

    It "無くなったファイルは、そのファイル名も書き出し待ちに記録する" {
        $pending = [PendingPublish]::new()
        $pending.Add("営業\無い.xlsx", $true)
        $folder = [System.IO.Path]::GetDirectoryName((getBookDir "営業\無い.xlsx"))
        $flush = $pending.TakeAll()
        $flush[$folder].Contains("無い.xlsx") | Should -Be $true
    }

    It "一度取り出したフォルダは、書き出し待ちから消える（二重に書き出さない）" {
        $pending = [PendingPublish]::new()
        $pending.Add("営業\b.xlsx", $false)
        [void]$pending.TakeAll()
        $pending.TakeAll().Count | Should -Be 0
    }

    It "前回のインデックス作成で本文インデックスに入れていないフォルダ（MarkFolder）も書き出し待ちにする" {
        $pending = [PendingPublish]::new()
        $pending.MarkFolder("C:\work\content_index\営業")
        $flush = $pending.TakeAll()
        $flush.ContainsKey("C:\work\content_index\営業") | Should -Be $true
        $flush["C:\work\content_index\営業"].Count | Should -Be 0
    }

    It "戻り値は大文字・小文字を区別しない Dictionary のまま（Hashtable に変わって比較が変わらない）" {
        $pending = [PendingPublish]::new()
        $pending.Add("Zz\a.xlsx", $false)
        $pending.Add("aa\b.xlsx", $false)
        $flush = $pending.TakeAll()
        $folderB = [System.IO.Path]::GetDirectoryName((getBookDir "aa\b.xlsx"))
        $flush.ContainsKey($folderB.ToUpper()) | Should -Be $true
    }

    It "取り込み中（Dispatch）のフォルダは、TakeFlushable では取り出さない。終わる（Complete）と取り出せる" {
        $pending = [PendingPublish]::new()
        $pending.Add("営業\a.xlsx", $false)
        $folder = [System.IO.Path]::GetDirectoryName((getBookDir "営業\a.xlsx"))
        $pending.AddPending($folder)
        $pending.Dispatch($folder)

        $pending.TakeFlushable().Count | Should -Be 0

        $pending.Complete($folder)
        $flush = $pending.TakeFlushable()
        $flush.ContainsKey($folder) | Should -Be $true
    }

    It "まだ渡していない（AddPending だけ）フォルダは、TakeFlushable では取り出さない" {
        $pending = [PendingPublish]::new()
        $pending.Add("営業\a.xlsx", $false)
        $pending.Add("営業\b.xlsx", $false)
        $folder = [System.IO.Path]::GetDirectoryName((getBookDir "営業\a.xlsx"))
        $pending.AddPending($folder)  # a.xlsx: まだ渡していない
        $pending.AddPending($folder)  # b.xlsx: まだ渡していない
        $pending.Dispatch($folder)    # a.xlsx を渡す（渡していない数が 1 減り、取り込み中が 1 増える）
        $pending.Complete($folder)    # a.xlsx が終わる（取り込み中が 0 に戻るが、b.xlsx はまだ渡していない）

        $pending.TakeFlushable().Count | Should -Be 0

        $pending.Skip($folder)        # b.xlsx を渡さずに済ませる（元のファイルが無くなった等）
        $flush = $pending.TakeFlushable()
        $flush.ContainsKey($folder) | Should -Be $true
    }

    It "取り込み中・まだ渡していないファイルが 1 つでもあるフォルダだけを残し、無いフォルダだけ取り出す" {
        $pending = [PendingPublish]::new()
        $pending.Add("busy\a.xlsx", $false)
        $pending.Add("done\b.xlsx", $false)
        $busyFolder = [System.IO.Path]::GetDirectoryName((getBookDir "busy\a.xlsx"))
        $doneFolder = [System.IO.Path]::GetDirectoryName((getBookDir "done\b.xlsx"))
        $pending.AddPending($busyFolder)
        $pending.Dispatch($busyFolder)  # busy はまだ取り込み中のまま

        $flush = $pending.TakeFlushable()
        $flush.ContainsKey($doneFolder) | Should -Be $true
        $flush.ContainsKey($busyFolder) | Should -Be $false
    }
}

Describe "PendingPublish（外し終えていない削除 Unpublished）" -Tag Unit {
    BeforeAll {
        $script:folder = [System.IO.Path]::GetDirectoryName((getBookDir "営業\無い.xlsx"))
    }

    It "無くなったファイル（Add の removed）は Unpublished に入り、取り込んだだけのファイルは入らない" {
        $pending = [PendingPublish]::new()
        $pending.Add("営業\無い.xlsx", $true)
        $pending.Add("経理\ある.xlsx", $false)

        $pending.Unpublished.Count | Should -Be 1
        $pending.Unpublished[$script:folder].Contains("無い.xlsx") | Should -Be $true
    }

    It "TakeAll・TakeFlushable で取り出しても、MarkPublished されるまで Unpublished に残る" {
        $pending = [PendingPublish]::new()
        $pending.Add("営業\無い.xlsx", $true)
        [void]$pending.TakeFlushable()
        $pending.Unpublished.ContainsKey($script:folder) | Should -Be $true
        $pending.Add("営業\無い.xlsx", $true)
        [void]$pending.TakeAll()
        $pending.Unpublished[$script:folder].Contains("無い.xlsx") | Should -Be $true
    }

    It "MarkPublished で、書き出し終えたフォルダだけ Unpublished から外れる" {
        $pending = [PendingPublish]::new()
        $pending.Add("営業\無い.xlsx", $true)
        $pending.Add("経理\無い.xlsx", $true)
        $other = [System.IO.Path]::GetDirectoryName((getBookDir "経理\無い.xlsx"))

        $taken = $pending.TakeAll()
        $pending.MarkPublished([string[]]@($script:folder), $taken)

        $pending.Unpublished.ContainsKey($script:folder) | Should -Be $false
        $pending.Unpublished.ContainsKey($other) | Should -Be $true
    }

    It "書き出しを飛ばしたフォルダの名前は、同じ回の後の書き出しの MarkPublished で外れない" {
        $pending = [PendingPublish]::new()
        $pending.Restore($script:folder, [string[]]@("前回の残り.xlsx"))
        [void]$pending.TakeAll()  # 1 回目: 書き出しに失敗して飛ばした（MarkPublished しない）
        $pending.Add("営業\無い.xlsx", $true)
        $taken = $pending.TakeAll()  # 2 回目: 成功した
        $pending.MarkPublished([string[]]@($script:folder), $taken)

        $pending.Unpublished.ContainsKey($script:folder) | Should -Be $true
        $pending.Unpublished[$script:folder].Contains("前回の残り.xlsx") | Should -Be $true
        $pending.Unpublished[$script:folder].Contains("無い.xlsx") | Should -Be $false
    }

    It "Restore は、書き出し待ちと Unpublished の両方に名前を戻す" {
        $pending = [PendingPublish]::new()
        $pending.Restore($script:folder, [string[]]@("a.xlsx", "b.xlsx"))

        $pending.Unpublished[$script:folder].Count | Should -Be 2
        $flush = $pending.TakeAll()
        $flush[$script:folder].Contains("a.xlsx") | Should -Be $true
        $flush[$script:folder].Contains("b.xlsx") | Should -Be $true
    }
}

Describe "PendingPublish.SnapshotRemovals / RestoreRemovals" -Tag Unit {
    It "写しへ戻すと、写しのあとに足した無くなったファイルの記録は消え、写しの中身は戻る。写しは何度戻しても変わらない" {
        $pending = [PendingPublish]::new()
        $pending.Add("営業\前.docx", $true)
        $snapshot = $pending.SnapshotRemovals()
        $pending.Add("営業\後.docx", $true)
        $pending.Add("経理\別.docx", $true)

        $pending.RestoreRemovals($snapshot)

        $names = @($pending.Unpublished.Values | ForEach-Object { $_ } | ForEach-Object { $_ })
        $names | Should -Be @("前.docx")
        @($pending.Folders.Keys).Count | Should -Be 1
        # 戻したあとに足しても、写しは書き換わらない（もう一度戻せる）
        $pending.Add("営業\また.docx", $true)
        $pending.RestoreRemovals($snapshot)
        @($pending.Unpublished.Values | ForEach-Object { $_ } | ForEach-Object { $_ }) | Should -Be @("前.docx")
    }
}
