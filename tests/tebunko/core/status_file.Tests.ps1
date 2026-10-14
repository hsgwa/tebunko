# 取り込み一覧の読み書き（tebunko\core\status_file.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "readStatusFile / writeStatusFile" -Tag Io {
    It "まとめて書き出した行と、1件ずつ追記した行の整形が同じ（速さのために別々に書いているため）" {
        # writeStatusFile は数万行を速く書くため、toStatusLine と同じ整形をその場に展開している
        $row = newStatusRow "営業\見積.xlsx" "2025/01/10 12:34:56" "10420" $stateFailed "" "2026/09/20 10:00:00" "エラー`tの`r`n説明"
        $path = "$TestDrive\status_same.tsv"
        writeStatusFile @() @($row) $path
        $written = @(readStatusLines $path)[-1]
        $written | Should -Be (toStatusLine $row)
        $written | Should -Be "営業\見積.xlsx`t2025/01/10 12:34:56`t10420`t${stateFailed}`t`t2026/09/20 10:00:00`tエラー の 説明`t"
        $done = newStatusRow "営業\見積.xlsx" "2025/01/10 12:34:56" "10420" $stateDone "3" "2026/09/20 10:00:00" "" "2"
        writeStatusFile @() @($done) $path
        @(readStatusLines $path)[-1] | Should -Be (toStatusLine $done)
    }

    It "書き込んだクロール対象フォルダと行をそのまま読み込める（[ ] や先頭の空白を含むパス）" {
        $path = "$TestDrive\status[1].tsv"
        $rows = @(
            (newStatusRow "a\[確定]見積.xlsx" "2025/01/10 12:34:56" "10420" $stateDone "3" "2026/09/19 10:00:00"),
            (newStatusRow " b.xls" "2025/02/01 08:00:00" "0" $stateNew)
        )
        $folders = @(
            [pscustomobject]@{ Path = "C:\data [1]"; Name = "data [1]" },
            [pscustomobject]@{ Path = "D:\"; Name = "D" }
        )
        writeStatusFile $folders $rows $path

        $status = readStatusFile $path
        $status.Folders.Count | Should -Be 2
        $status.Folders[0].Path | Should -Be "C:\data [1]"
        $status.Folders[0].Name | Should -Be "data [1]"
        $status.Folders[1].Path | Should -Be "D:\"
        $status.Folders[1].Name | Should -Be "D"
        $status.Rows.Count | Should -Be 2
        $row = $status.Rows["a\[確定]見積.xlsx"]
        $row.更新日時 | Should -Be "2025/01/10 12:34:56"
        $row.サイズ | Should -Be "10420"
        $row.状態 | Should -Be $stateDone
        $row.TSV数 | Should -Be "3"
        $status.Rows[" b.xls"].状態 | Should -Be $stateNew
    }

    It "先頭にクロール対象フォルダ（パス・インデックス名）、次に見出しのTSVになる" {
        $path = "$TestDrive\status_lines.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "a.xlsx" "2025/01/10 12:34:56" "1" $stateNew) $path
        $lines = [System.IO.File]::ReadAllLines($path)
        $lines[0] | Should -Be "クロール対象フォルダ`tC:\data`tdata"
        $lines[1] | Should -Be "相対パス`t更新日時`tサイズ`t状態`tTSV数`t取り込み日時`tエラー`t抽出版"
        $lines[2] | Should -Be "a.xlsx`t2025/01/10 12:34:56`t1`t未取り込み`t`t`t`t"
    }

    It "抽出版を書き込んで読み込める" {
        $path = "$TestDrive\status_version.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @((newStatusRow "a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "2" "" "" "2"), (newStatusRow "b.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "" "" "3")) $path
        $status = readStatusFile $path
        $status.Rows["a.xlsx"].抽出版 | Should -Be "2"
        $status.Rows["b.xlsx"].抽出版 | Should -Be "3"
    }

    It "エラーメッセージのタブ・改行はスペースにして1行に収める" {
        $path = "$TestDrive\status_error.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "a.xlsx" "" "" $stateFailed "" "" "1行目`r`n2行目`tタブ") $path
        (readStatusFile $path).Rows["a.xlsx"].エラー | Should -Be "1行目 2行目 タブ"
    }

    It "列数の合わない行（書き込み途中で中断した行）は無視する" {
        $path = "$TestDrive\status_broken.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1") $path
        [System.IO.File]::AppendAllText($path, "a.xlsx`t2025/01/10", $utf8Bom)

        $status = readStatusFile $path
        $status.Rows["a.xlsx"].状態 | Should -Be $stateDone
    }

    It "書き直すと既存のファイルを置き換え、一時ファイルは残らない" {
        $path = "$TestDrive\status_replace.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\old"; Name = "old" }) @(newStatusRow "old\a.xlsx") $path
        writeStatusFile @([pscustomobject]@{ Path = "C:\new"; Name = "new" }) @(newStatusRow "new\b.xlsx") $path

        $status = readStatusFile $path
        @($status.Folders | ForEach-Object { $_.Path }) | Should -Be @("C:\new")
        @($status.Rows.Keys) | Should -Be @("new\b.xlsx")
        Test-Path -LiteralPath "${path}.tmp" | Should -Be $false
    }

    It "ファイルが無ければ空の一覧を返す" {
        $status = readStatusFile "$TestDrive\none_status.tsv"
        $status.Folders.Count | Should -Be 0
        $status.Rows.Count | Should -Be 0
    }
}

Describe "getIndexNameMap（見出し行まで読む）" -Tag Io {
    It "クロール対象フォルダの行だけを読み、見出し行の後は読まない" {
        $path = "$TestDrive\name_map.tsv"
        writeListFile $path @(
            "${statusFolderKey}`tC:\data\見積`t見積",
            "${statusFolderKey}`tC:\old",
            ($statusColumns -join "`t"),
            "${statusFolderKey}`tC:\x`tx")
        $map = getIndexNameMap $path
        $map.Count | Should -Be 1
        $map["見積"] | Should -Be "C:\data\見積"
        (getIndexNameMap "$TestDrive\none_status.tsv").Count | Should -Be 0
    }
}
