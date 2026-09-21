# インデックスの作成・集計・改名・削除（windox_grep\index\index_store.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"

Describe "getIndexStats" -Tag Io {
    It "インデックス名ごとに件数と最終変換日時を集計する" {
        $path = "$TestDrive\stats.tsv"
        writeStatusFile @(
            [pscustomobject]@{ Path = "C:\data"; Name = "営業" },
            [pscustomobject]@{ Path = "D:\tech"; Name = "技術" }
        ) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00"),
            (newStatusRow "営業\b.xlsx" "2025/01/10 12:34:56" "1" $stateFailed "" "2026/09/19 11:00:00" "原因"),
            (newStatusRow "営業\c.docx" "2025/01/10 12:34:56" "1" $stateNew),
            (newStatusRow "技術\d.pptx" "2025/01/10 12:34:56" "1" $stateDone "2" "2026/09/17 09:00:00")
        ) $path

        $stats = getIndexStats (readStatusFile $path).Rows
        $stats["営業"].Total | Should Be 3
        $stats["営業"].Done | Should Be 1
        $stats["営業"].Failed | Should Be 1
        $stats["営業"].Pending | Should Be 1
        $stats["営業"].LastConverted | Should Be "2026/09/19 11:00:00"
        $stats["技術"].Total | Should Be 1
        $stats["技術"].LastConverted | Should Be "2026/09/17 09:00:00"
    }

    It "行が無ければ空を返す" {
        (getIndexStats $null).Count | Should Be 0
    }

    It "getConversionState からも集計を取れる" {
        $path = "$TestDrive\stats_state.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "営業" }) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00")
        ) $path

        (getConversionState -path $path).IndexStats["営業"].Total | Should Be 1
    }
}

Describe "renameIndex" -Tag Io {
    It "インデックスのフォルダと変換一覧の記録の名前を変え、中身はそのまま残す" {
        $dir = "$TestDrive\rename\index"
        $path = "$TestDrive\rename\変換一覧.tsv"
        New-Item -ItemType Directory -Path "$dir\営業\a.xlsx" -Force | Out-Null
        Set-Content -LiteralPath "$dir\営業\a.xlsx\Sheet1.tsv" -Value "本文" -Encoding UTF8
        New-Item -ItemType Directory -Path "$dir\技術" -Force | Out-Null
        writeStatusFile @(
            [pscustomobject]@{ Path = "C:\data"; Name = "営業" },
            [pscustomobject]@{ Path = "D:\tech"; Name = "技術" }
        ) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00"),
            (newStatusRow "技術\d.pptx" "2025/01/10 12:34:56" "1" $stateNew)
        ) $path

        renameIndex "営業" "営業部" $dir $path

        Test-Path "$dir\営業" | Should Be $false
        Get-Content -LiteralPath "$dir\営業部\a.xlsx\Sheet1.tsv" | Should Be "本文"

        $status = readStatusFile $path
        $status.Folders[0].Name | Should Be "営業部"
        $status.Folders[0].Path | Should Be "C:\data"
        $status.Rows.ContainsKey("営業部\a.xlsx") | Should Be $true
        $status.Rows["営業部\a.xlsx"].状態 | Should Be $stateDone
        # ほかのインデックスはそのまま
        $status.Folders[1].Name | Should Be "技術"
        $status.Rows.ContainsKey("技術\d.pptx") | Should Be $true
    }

    It "インデックスのフォルダがまだ無くても、変換一覧の記録は変える" {
        $dir = "$TestDrive\rename2\index"
        $path = "$TestDrive\rename2\変換一覧.tsv"
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "営業" }) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateNew)
        ) $path

        renameIndex "営業" "営業部" $dir $path

        (readStatusFile $path).Rows.ContainsKey("営業部\a.xlsx") | Should Be $true
    }

    It "同じ名前のフォルダが既にあれば例外にする" {
        $dir = "$TestDrive\rename3\index"
        $path = "$TestDrive\rename3\変換一覧.tsv"
        New-Item -ItemType Directory -Path "$dir\営業" -Force | Out-Null
        New-Item -ItemType Directory -Path "$dir\技術" -Force | Out-Null

        { renameIndex "営業" "技術" $dir $path } | Should Throw
    }
}

Describe "removeIndex" -Tag Io {
    It "インデックスのフォルダと変換一覧の記録を削除し、ほかのインデックスは残す" {
        $dir = "$TestDrive\remove\index"
        $path = "$TestDrive\remove\変換一覧.tsv"
        New-Item -ItemType Directory -Path "$dir\営業\a.xlsx" -Force | Out-Null
        Set-Content -LiteralPath "$dir\営業\a.xlsx\Sheet1.tsv" -Value "本文" -Encoding UTF8
        New-Item -ItemType Directory -Path "$dir\技術" -Force | Out-Null
        writeStatusFile @(
            [pscustomobject]@{ Path = "C:\data"; Name = "営業" },
            [pscustomobject]@{ Path = "D:\tech"; Name = "技術" }
        ) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00"),
            (newStatusRow "技術\d.pptx" "2025/01/10 12:34:56" "1" $stateNew)
        ) $path

        removeIndex "営業" $dir $path

        Test-Path "$dir\営業" | Should Be $false
        Test-Path "$dir\技術" | Should Be $true

        $status = readStatusFile $path
        $status.Folders.Count | Should Be 1
        $status.Folders[0].Name | Should Be "技術"
        $status.Rows.Count | Should Be 1
        $status.Rows.ContainsKey("技術\d.pptx") | Should Be $true
    }

    It "名前が空なら何もしない" {
        $path = "$TestDrive\remove2\変換一覧.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "営業" }) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateNew)
        ) $path

        removeIndex "" "$TestDrive\remove2\index" $path

        (readStatusFile $path).Rows.Count | Should Be 1
    }
}

Describe "getSearchIndexes" -Tag Io {
    It "インデックスのフォルダが無ければ空" {
        @(getSearchIndexes "$TestDrive\無いフォルダ\index").Count | Should Be 0
    }

    It "work\index 直下のフォルダをインデックス 1 件として返し、元のフォルダも返す" {
        $dir = "$TestDrive\一覧\index"
        $settings = "$TestDrive\一覧\setting.config"
        foreach ($name in @("見積", "営業")) {
            [void](New-Item -ItemType Directory -Path "$dir\$name" -Force)
        }
        writeTargetFolders @(
            [pscustomobject]@{ Name = "見積"; Path = "C:\data\見積"; Enabled = $true },
            [pscustomobject]@{ Name = "営業"; Path = "C:\data\営業"; Enabled = $true }) $settings

        $indexes = @(getSearchIndexes $dir "$TestDrive\変換一覧なし.tsv" $settings)
        $indexes.Count | Should Be 2
        # ［1 インデックス作成］の一覧と同じ並び
        $indexes[0].Name | Should Be "見積"
        $indexes[0].SourcePath | Should Be "C:\data\見積"
        $indexes[0].Path | Should Be (Resolve-Path -LiteralPath "$dir\見積").ProviderPath
        $indexes[1].Name | Should Be "営業"
        $indexes[1].SourcePath | Should Be "C:\data\営業"
    }

    It "一覧に無いインデックス（コピーしたものなど）も名前順で後ろに並べる" {
        $dir = "$TestDrive\コピー\index"
        $settings = "$TestDrive\コピー\setting.config"
        foreach ($name in @("報告書", "あとから", "見積")) {
            [void](New-Item -ItemType Directory -Path "$dir\$name" -Force)
        }
        writeTargetFolders @([pscustomobject]@{ Name = "見積"; Path = "C:\data\見積"; Enabled = $true }) $settings

        $indexes = @(getSearchIndexes $dir "$TestDrive\変換一覧なし.tsv" $settings)
        @($indexes | ForEach-Object { $_.Name }) -join "," | Should Be "見積,あとから,報告書"
        # 元のフォルダが分からないものは空
        $indexes[1].SourcePath | Should Be ""
    }

    It "一覧にも変換一覧にも無いインデックスは、そのフォルダの 元のフォルダ.txt から元のフォルダを読む" {
        $dir = "$TestDrive\コピー2\index"
        $settings = "$TestDrive\コピー2\setting.config"
        [void](New-Item -ItemType Directory -Path "$dir\営業" -Force)
        # ほかの PC で作ったインデックスをフォルダごとコピーした状態
        writeSourceFolderFile @([pscustomobject]@{ Path = "\\server\営業"; Name = "営業" }) $dir

        $indexes = @(getSearchIndexes $dir "$TestDrive\変換一覧なし.tsv" $settings)
        $indexes.Count | Should Be 1
        $indexes[0].Name | Should Be "営業"
        $indexes[0].SourcePath | Should Be "\\server\営業"
    }
}

Describe "getIndexNameMap（見出し行まで読む）" -Tag Io {
    It "変換対象フォルダの行だけを読み、見出し行の後は読まない" {
        $path = "$TestDrive\name_map.tsv"
        writeListFile $path @(
            "${statusFolderKey}`tC:\data\見積`t見積",
            "${statusFolderKey}`tC:\old",
            ($statusColumns -join "`t"),
            "${statusFolderKey}`tC:\x`tx")
        $map = getIndexNameMap $path
        $map.Count | Should Be 1
        $map["見積"] | Should Be "C:\data\見積"
        (getIndexNameMap "$TestDrive\none_status.tsv").Count | Should Be 0
    }
}

Describe "getIndexTsvCounts / testIndexComplete" -Tag Io {
    function newTestIndex {
        # テスト用のインデックス（work\index 相当）を作る
        param ([string]$dir)

        [System.IO.Directory]::CreateDirectory("$dir\営業\2024\A社.xlsx") | Out-Null
        writeListFile "$dir\営業\2024\A社.xlsx\明細.tsv" @("a")
        writeListFile "$dir\営業\2024\A社.xlsx\表紙.tsv" @("b")
        [System.IO.Directory]::CreateDirectory("$dir\営業\空.xlsx") | Out-Null
        [System.IO.Directory]::CreateDirectory("$dir\営業\資料.docx") | Out-Null
        writeListFile "$dir\営業\資料.docx\ページ001.tsv" @("c")
    }

    It "元のファイル1つ分のフォルダごとにTSVの数を数える" {
        $dir = "$TestDrive\index1"
        newTestIndex $dir
        $counts = getIndexTsvCounts $dir
        $counts["営業\2024\A社.xlsx"] | Should Be 2
        $counts["営業\資料.docx"] | Should Be 1
    }

    It "TSVの無いフォルダ（内容が空のファイル）は0件として数える" {
        $dir = "$TestDrive\index2"
        newTestIndex $dir
        $counts = getIndexTsvCounts $dir
        $counts.ContainsKey("営業\空.xlsx") | Should Be $true
        $counts["営業\空.xlsx"] | Should Be 0
    }

    It "大文字・小文字を区別しない" {
        $dir = "$TestDrive\index3"
        newTestIndex $dir
        (getIndexTsvCounts $dir)["営業\2024\a社.XLSX"] | Should Be 2
    }

    It "インデックスのフォルダの直下のTSV（以前の形式）は数えない" {
        $dir = "$TestDrive\index4"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        writeListFile "$dir\ブック.xlsx_シート.tsv" @("a")
        (getIndexTsvCounts $dir).Count | Should Be 0
    }

    It "フォルダが無ければ空を返す" {
        (getIndexTsvCounts "$TestDrive\none_index").Count | Should Be 0
    }

    It "TSVがそろっていれば「済」のままにする" {
        $dir = "$TestDrive\index5"
        newTestIndex $dir
        $counts = getIndexTsvCounts $dir
        $row = newStatusRow "営業\2024\A社.xlsx" "2025/01/10 12:34:56" "100" ${stateDone} "2"
        testIndexComplete $row $row.相対パス $counts | Should Be $true
    }

    It "インデックスのフォルダを直接削除した場合は、そろっていないとする" {
        $dir = "$TestDrive\index6"
        newTestIndex $dir
        $counts = getIndexTsvCounts $dir
        $row = newStatusRow "営業\2024\消えたブック.xlsx" "2025/01/10 12:34:56" "100" ${stateDone} "3"
        testIndexComplete $row $row.相対パス $counts | Should Be $false
    }

    It "TSVが足りない場合も、そろっていないとする" {
        $dir = "$TestDrive\index7"
        newTestIndex $dir
        $counts = getIndexTsvCounts $dir
        $row = newStatusRow "営業\2024\A社.xlsx" "2025/01/10 12:34:56" "100" ${stateDone} "5"
        testIndexComplete $row $row.相対パス $counts | Should Be $false
    }

    It "内容が空のファイル（TSV 0 件）は、フォルダがあればそろっているとする" {
        $dir = "$TestDrive\index8"
        newTestIndex $dir
        $counts = getIndexTsvCounts $dir
        $row = newStatusRow "営業\空.xlsx" "2025/01/10 12:34:56" "100" ${stateDone} "0"
        testIndexComplete $row $row.相対パス $counts | Should Be $true
    }

    It "0 バイトのTSVがあるフォルダは、壊れているとして作り直す" {
        $dir = "$TestDrive\index9"
        newTestIndex $dir
        # 書き込みの途中で電源が落ちた場合など（空のシート・ページは保存しないため、0 バイトのTSVは異常）
        [System.IO.File]::WriteAllBytes("$dir\営業\2024\A社.xlsx\途中.tsv", (New-Object byte[] 0))
        $counts = getIndexTsvCounts $dir
        $counts["営業\2024\A社.xlsx"] | Should Be ${indexBrokenCount}
        $counts["営業\資料.docx"] | Should Be 1   # ほかのファイルは巻き込まない

        $row = newStatusRow "営業\2024\A社.xlsx" "2025/01/10 12:34:56" "100" ${stateDone} "2"
        testIndexComplete $row $row.相対パス $counts | Should Be $false
        $other = newStatusRow "営業\資料.docx" "2025/01/10 12:34:56" "100" ${stateDone} "1"
        testIndexComplete $other $other.相対パス $counts | Should Be $true
    }

    It "0 バイトのTSVが先に見つかっても、後のTSVで数え直さない" {
        $dir = "$TestDrive\index10"
        [System.IO.Directory]::CreateDirectory("$dir\営業\B社.xlsx") | Out-Null
        [System.IO.File]::WriteAllBytes("$dir\営業\B社.xlsx\001_途中.tsv", (New-Object byte[] 0))
        writeListFile "$dir\営業\B社.xlsx\002_あと.tsv" @("a")
        (getIndexTsvCounts $dir)["営業\B社.xlsx"] | Should Be ${indexBrokenCount}
    }

    It "TSVの数を記録していない行・数えられなかった場合は確認しない" {
        $row = newStatusRow "営業\2024\A社.xlsx" "2025/01/10 12:34:56" "100" ${stateDone} ""
        testIndexComplete $row $row.相対パス (getIndexTsvCounts "$TestDrive\none_index2") | Should Be $true
        $done = newStatusRow "営業\2024\A社.xlsx" "2025/01/10 12:34:56" "100" ${stateDone} "2"
        testIndexComplete $done $done.相対パス $null | Should Be $true
    }
}

Describe "publishIndexFiles" -Tag Io {
    It "作業フォルダのTSVを、元のファイルのフォルダに入れる" {
        $from = "$TestDrive\pub1\tmp"
        [System.IO.Directory]::CreateDirectory($from) | Out-Null
        writeListFile "$from\明細.tsv" @("a")
        writeListFile "$from\表紙.tsv" @("b")
        $bookDir = "$TestDrive\pub1\index\営業\A社.xlsx"

        publishIndexFiles $from $bookDir "$TestDrive\pub1\出力\A社.xlsx"

        @(Get-ChildItem -LiteralPath $bookDir -Filter "*.tsv").Count | Should Be 2
        @(Get-ChildItem -LiteralPath $from -Filter "*.tsv").Count | Should Be 0
        Test-Path -LiteralPath "$TestDrive\pub1\出力\A社.xlsx" | Should Be $false
    }

    It "以前の変換結果は残さず入れ替える（シートの削除・名前変更に追従する）" {
        $from = "$TestDrive\pub2\tmp"
        [System.IO.Directory]::CreateDirectory($from) | Out-Null
        $bookDir = "$TestDrive\pub2\index\営業\A社.xlsx"
        [System.IO.Directory]::CreateDirectory($bookDir) | Out-Null
        writeListFile "$bookDir\前のシート.tsv" @("old")
        writeListFile "$from\新しいシート.tsv" @("new")

        publishIndexFiles $from $bookDir "$TestDrive\pub2\出力\A社.xlsx"

        Test-Path -LiteralPath "$bookDir\前のシート.tsv" | Should Be $false
        Test-Path -LiteralPath "$bookDir\新しいシート.tsv" | Should Be $true
    }

    It "TSVが1件も無ければ、空のフォルダにする（内容が空のファイル）" {
        $from = "$TestDrive\pub3\tmp"
        [System.IO.Directory]::CreateDirectory($from) | Out-Null
        $bookDir = "$TestDrive\pub3\index\営業\空.xlsx"

        publishIndexFiles $from $bookDir "$TestDrive\pub3\出力\空.xlsx"

        Test-Path -LiteralPath $bookDir -PathType Container | Should Be $true
        @(Get-ChildItem -LiteralPath $bookDir -Filter "*.tsv").Count | Should Be 0
    }

    It "前回の出力用フォルダが残っていても入れ替えられる" {
        $from = "$TestDrive\pub4\tmp"
        [System.IO.Directory]::CreateDirectory($from) | Out-Null
        writeListFile "$from\明細.tsv" @("a")
        $staging = "$TestDrive\pub4\出力\A社.xlsx"
        [System.IO.Directory]::CreateDirectory($staging) | Out-Null
        writeListFile "$staging\前回の残り.tsv" @("old")
        $bookDir = "$TestDrive\pub4\index\営業\A社.xlsx"

        publishIndexFiles $from $bookDir $staging

        @(Get-ChildItem -LiteralPath $bookDir -Filter "*.tsv" | ForEach-Object { $_.Name }) | Should Be "明細.tsv"
    }
}
