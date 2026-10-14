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
            (newStatusRow "data [1]\ b.xls" "2025/02/01 08:00:00" "0" $stateNew)
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
        $status.Rows["data [1]\ b.xls"].状態 | Should -Be $stateNew
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
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @((newStatusRow "data\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "2" "" "" "2"), (newStatusRow "data\b.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "" "" "3")) $path
        $status = readStatusFile $path
        $status.Rows["data\a.xlsx"].抽出版 | Should -Be "2"
        $status.Rows["data\b.xlsx"].抽出版 | Should -Be "3"
    }

    It "エラーメッセージのタブ・改行はスペースにして1行に収める" {
        $path = "$TestDrive\status_error.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "data\a.xlsx" "" "" $stateFailed "" "" "1行目`r`n2行目`tタブ") $path
        (readStatusFile $path).Rows["data\a.xlsx"].エラー | Should -Be "1行目 2行目 タブ"
    }

    It "列数の合わない行（書き込み途中で中断した行）は無視する" {
        $path = "$TestDrive\status_broken.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "data" }) @(newStatusRow "data\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1") $path
        [System.IO.File]::AppendAllText($path, "data\a.xlsx`t2025/01/10", $utf8Bom)

        $status = readStatusFile $path
        $status.Rows["data\a.xlsx"].状態 | Should -Be $stateDone
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

    It "インデックス名・相対パスが使えない行は読み飛ばし、読み飛ばした行数を Rejected で返す" {
        $path = "$TestDrive\status_unsafe.tsv"
        $folders = @(
            [pscustomobject]@{ Path = "C:\data"; Name = "営業" },
            [pscustomobject]@{ Path = "C:\bad1"; Name = ".." },
            [pscustomobject]@{ Path = "C:\bad2"; Name = "..\外" },
            [pscustomobject]@{ Path = "C:\old1"; Name = "aux.backup" },  # 予約名でも、既にある名前はそのまま使う
            [pscustomobject]@{ Path = "C:\old2"; Name = "旧営業." }      # 今の決まりに合わないが、外へは出ない名前（前の版で作られた）
        )
        $rows = @(
            (newStatusRow "営業\見積.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1"),
            (newStatusRow "営業\サブ\見積2.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1"),
            (newStatusRow "..\外\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1"),
            (newStatusRow "営業\..\..\外\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1"),
            (newStatusRow "..\..\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1"),
            (newStatusRow "営業\.\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1"),
            (newStatusRow "営業\\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1"),
            (newStatusRow "a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1")
        )
        writeStatusFile $folders $rows $path

        $status = readStatusFile $path
        @($status.Folders | ForEach-Object { $_.Name }) | Should -Be @("営業", "aux.backup")
        @($status.Rows.Keys | Sort-Object) | Should -Be @(@("営業\サブ\見積2.xlsx", "営業\見積.xlsx") | Sort-Object)
        @($status.LegacyFolders | ForEach-Object { $_.Name }) | Should -Be @("旧営業.")
        $status.Rejected | Should -Be 8
    }

    It "予約名（aux.backup・CON.txt・CON）のインデックスは、フォルダも行も読み飛ばさず、そのまま読む" -TestCases @(
        @{ name = "aux.backup" }
        @{ name = "CON.txt" }
        @{ name = "CON" }
    ) {
        param ($name)
        $path = "$TestDrive\status_reserved_$([guid]::NewGuid().ToString('N')).tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = $name }) @(newStatusRow "${name}\見積.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1") $path

        $status = readStatusFile $path
        @($status.Folders | ForEach-Object { $_.Name }) | Should -Be @($name)
        @($status.Rows.Keys) | Should -Be @("${name}\見積.xlsx")
        $status.LegacyFolders.Count | Should -Be 0
        $status.Rejected | Should -Be 0
    }

    It "最後が . や空白の名前は、予約名と違い使えない名前として扱う（行は読み飛ばし、フォルダは LegacyFolders に分ける）" -TestCases @(
        @{ name = "営業." }
        @{ name = "営業 " }
    ) {
        param ($name)
        $path = "$TestDrive\status_trailing_$([guid]::NewGuid().ToString('N')).tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = $name }) @(newStatusRow "${name}\見積.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1") $path

        $status = readStatusFile $path
        $status.Folders.Count | Should -Be 0
        @($status.LegacyFolders | ForEach-Object { $_.Name }) | Should -Be @($name)
        $status.Rows.Count | Should -Be 0
        $status.Rejected | Should -Be 1
    }

    It "行の確かめは testIndexRelPath（名前だけの行は通さない）と同じ結果になる" -TestCases @(
        @{ rel = "営業\見積.xlsx" }
        @{ rel = "営業\サブ\見積2.xlsx" }
        @{ rel = "営業\a [1]\ b .xlsx" }
        @{ rel = "営業\レポート.\a.xlsx" }
        @{ rel = "営業\CON\a.xlsx" }
        @{ rel = "a.xlsx" }
        @{ rel = "営業" }
        @{ rel = "\a.xlsx" }
        @{ rel = "..\a.xlsx" }
        @{ rel = ".\a.xlsx" }
        @{ rel = "営業.\a.xlsx" }
        @{ rel = " 営業\a.xlsx" }
        @{ rel = "CON\a.xlsx" }
        @{ rel = "aux.backup\a.xlsx" }
        @{ rel = "営業\..\a.xlsx" }
        @{ rel = "営業\.\a.xlsx" }
        @{ rel = "営業\\a.xlsx" }
        @{ rel = "営業\a.xlsx\" }
        @{ rel = "営業/a.xlsx" }
        @{ rel = "営業\a:b.xlsx" }
        @{ rel = "営業\a*b.xlsx" }
        @{ rel = "営業\" + ("あ" * 256) }
        @{ rel = ("あ" * 256) + "\a.xlsx" }
        @{ rel = "営業\" + ("あ" * 255) + "\a.xlsx" }
    ) {
        param ($rel)
        $path = "$TestDrive\status_same_$([guid]::NewGuid().ToString('N')).tsv"
        writeStatusFile @() @(newStatusRow $rel "2025/01/10 12:34:56" "1" $stateDone "1") $path
        $accepted = @((readStatusFile $path).Rows.Keys).Count -eq 1
        $accepted | Should -Be ((testIndexRelPath $rel $false) -eq "")
    }

    It "同じ名前の 2 行目以降も、1 行目と同じ結果になる（通る名前・通らない名前）" {
        $path = "$TestDrive\status_second.tsv"
        $rows = @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1"),
            (newStatusRow "営業\b.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1"),
            (newStatusRow "CON\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1"),
            (newStatusRow "CON\b.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1"),
            (newStatusRow "営業.\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1"),
            (newStatusRow "営業.\b.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1"),
            (newStatusRow "営業\c.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1")
        )
        writeStatusFile @() $rows $path

        $status = readStatusFile $path
        @($status.Rows.Keys | Sort-Object) | Should -Be @(@("CON\a.xlsx", "CON\b.xlsx", "営業\a.xlsx", "営業\b.xlsx", "営業\c.xlsx") | Sort-Object)
        $status.Rejected | Should -Be 2
    }

    It "使えない行が無ければ Rejected は 0" {
        $path = "$TestDrive\status_safe.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "営業" }) @(newStatusRow "営業\見積.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1") $path
        (readStatusFile $path).Rejected | Should -Be 0
        (readStatusFile "$TestDrive\none_status2.tsv").Rejected | Should -Be 0
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
