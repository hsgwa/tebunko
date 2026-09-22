# フォルダ同士の比較（tebunko_diff\diff\diff_folder.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\tebunko_diff\lib.ps1"

function newFile {
    param ([string]$relPath, [long]$size = 100)

    return @{ RelPath = $relPath; Path = "C:\x\$relPath"; Size = $size; Time = [datetime]"2025-04-02" }
}

function newSet {
    return New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
}

Describe "getFolderEntries" -Tag Unit {
    $entries = getFolderEntries @((newFile "見積\A.xlsx" 10), (newFile "見積\C.xls"), (newFile "契約\N.docx" 5)) @((newFile "見積\a.xlsx" 11), (newFile "見積\D.xlsx"), (newFile "契約\N.docx" 5))

    It "相対パス（大文字・小文字を区別しない）で対応づけ、片方だけは追加・削除にする" {
        (@($entries | ForEach-Object { "$($_.RelPath)=$($_.Status)" }) -join ",") | Should Be "契約\N.docx=pending,見積\A.xlsx=pending,見積\C.xls=delete,見積\D.xlsx=insert"
    }

    It "サイズが同じ組だけ、ハッシュを比べる" {
        ($entries | Where-Object { $_.RelPath -eq "契約\N.docx" }).HashNeeded | Should Be $true
        ($entries | Where-Object { $_.RelPath -eq "見積\A.xlsx" }).HashNeeded | Should Be $false
    }
}

Describe "buildTreeRows" -Tag Unit {
    function newEntries {
        $entries = getFolderEntries @((newFile "見積\A.xlsx"), (newFile "見積\C.xls"), (newFile "契約\N.docx"), (newFile "顧客\一覧.xlsx")) @((newFile "見積\A.xlsx"), (newFile "見積\D.xlsx"), (newFile "契約\N.docx"), (newFile "顧客\一覧.xlsx"), (newFile "計画\案.pptx"))
        foreach ($entry in $entries) {
            switch ($entry.RelPath) {
                "見積\A.xlsx"   { $entry.Status = "change"; $entry.Changes = 3 }
                "契約\N.docx"   { $entry.Status = "same" }
                "顧客\一覧.xlsx" { $entry.Status = "same" }
            }
        }
        return $entries
    }

    It "左右に並べたツリーの行（フォルダが先、違いのあるフォルダを開く）" {
        $rows = buildTreeRows (newEntries) (newSet) (newSet) $false
        (@($rows | ForEach-Object { "$($_.Depth)$($_.Mark)$($_.Name)" }) -join ",") |
            Should Be "0=契約,0+計画,1+案.pptx,0≠見積,1≠A.xlsx,1−C.xls,1+D.xlsx,0=顧客"
    }

    It "片方にしか無いファイル・フォルダは、もう片方を空きにする" {
        $rows = buildTreeRows (newEntries) (newSet) (newSet) $false
        ($rows | Where-Object { $_.Name -eq "計画" }).LeftEmpty | Should Be $true
        ($rows | Where-Object { $_.Name -eq "C.xls" }).RightEmpty | Should Be $true
        ($rows | Where-Object { $_.Name -eq "計画" }).RightMeta | Should Be "フォルダ · 1 ファイル"
    }

    It "同じファイルと、中がすべて同じフォルダを隠す" {
        $rows = buildTreeRows (newEntries) (newSet) (newSet) $true
        (@($rows | ForEach-Object { $_.Name }) -join ",") | Should Be "計画,案.pptx,見積,A.xlsx,C.xls,D.xlsx"
    }

    It "開いた・閉じたフォルダを覚える" {
        $opened = newSet
        [void]$opened.Add("契約")
        $closed = newSet
        [void]$closed.Add("見積")
        $rows = buildTreeRows (newEntries) $opened $closed $false
        (@($rows | ForEach-Object { $_.Name }) -join ",") | Should Be "契約,N.docx,計画,案.pptx,見積,顧客"
        ($rows | Where-Object { $_.Name -eq "見積" }).Glyph | Should Be "▸"
    }

    It "比べている途中のフォルダは、進み具合を出す" {
        $entries = getFolderEntries @((newFile "会議\1.pptx"), (newFile "会議\2.pptx")) @((newFile "会議\1.pptx"), (newFile "会議\2.pptx"))
        $entries[0].Status = "change"
        $rows = buildTreeRows $entries (newSet) (newSet) $false
        $rows[0].Status | Should Be "running"
        $rows[0].RightMeta | Should Be "比較中 1 / 2"
    }

    It "ファイルの行の右側に、結果とサイズ・更新日時を出す" {
        $rows = buildTreeRows (newEntries) (newSet) (newSet) $false
        ($rows | Where-Object { $_.Name -eq "A.xlsx" }).RightMeta | Should Be "変更 3 · 100 B · 2025/04/02"
    }
}

Describe "setEntryResult" -Tag Unit {
    It "中身に違いが無ければ「中身は同じ」、あれば変更の数" {
        $entry = @{ Status = "running" }
        $same = compareOfficeUnits "Word" ([ordered]@{ "ページ001" = [string[]]@("a") }) ([ordered]@{ "ページ001" = [string[]]@("a") })
        setEntryResult $entry $same
        $entry.Status | Should Be "similar"
        $changed = compareOfficeUnits "Word" ([ordered]@{ "ページ001" = [string[]]@("a") }) ([ordered]@{ "ページ001" = [string[]]@("a", "b") })
        setEntryResult $entry $changed
        $entry.Status | Should Be "change"
        $entry.Inserts | Should Be 1
    }
}

Describe "次の違うファイル" -Tag Unit {
    $entries = getFolderEntries @((newFile "a\1.xlsx"), (newFile "a\2.xlsx"), (newFile "b\3.xlsx")) @((newFile "a\1.xlsx"), (newFile "a\2.xlsx"), (newFile "b\3.xlsx"), (newFile "c.xlsx"))
    $entries | Where-Object { $_.RelPath -eq "a\1.xlsx" } | ForEach-Object { $_.Status = "same" }
    $entries | Where-Object { $_.RelPath -eq "a\2.xlsx" } | ForEach-Object { $_.Status = "change" }
    $entries | Where-Object { $_.RelPath -eq "b\3.xlsx" } | ForEach-Object { $_.Status = "same" }

    It "ツリーの順（フォルダが先）で、次・前の違うファイルを返す" {
        findNextDiffEntry $entries "" 1 | Should Be "a\2.xlsx"
        findNextDiffEntry $entries "a\2.xlsx" 1 | Should Be "c.xlsx"
        findNextDiffEntry $entries "c.xlsx" 1 | Should Be ""
        findNextDiffEntry $entries "c.xlsx" -1 | Should Be "a\2.xlsx"
    }

    It "閉じたフォルダの中へ移るときに開くフォルダ" {
        (getFolderPathsToOpen "営業\2024\見積\A.xlsx") -join "|" | Should Be "営業|営業\2024|営業\2024\見積"
    }

    It "表示中の行で、次の違うファイルの位置" {
        $rows = buildTreeRows $entries (newSet) (newSet) $false
        $index = findNextDiffRow $rows -1 1
        $rows[$index].Name | Should Be "2.xlsx"
    }
}

Describe "formatFileSize" -Tag Unit {
    It "B・KB・MB" {
        formatFileSize 500 | Should Be "500 B"
        formatFileSize 1536 | Should Be "2 KB"
        formatFileSize (1.5MB) | Should Be "1.5 MB"
    }
}
