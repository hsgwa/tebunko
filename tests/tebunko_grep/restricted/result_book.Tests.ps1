# 制限モードの検索結果のブック（tebunko_grep\restricted\result_book.ps1）のテスト。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\tebunko_grep\restricted\restricted_search.ps1"
. "${scriptsDir}\tebunko_grep\restricted\result_book_view.ps1"
. "${scriptsDir}\tebunko_grep\restricted\result_book.ps1"

# インデックス（work\index にあたるフォルダ）と、元のフォルダの対応を作る
function newBookFixture {
    $root = Join-Path $TestDrive "index"
    $files = [ordered]@{
        "見積\2024\a b.xlsx\売上 '1'.tsv"      = "品名`t金額`r`nりんご`t=SUM(A1)`r`nみかん`t`"青森$([char]0x2028)りんご`"`r`nx`r`ny`r`n"
        "見積\2024\a b.xlsx\売上 '1'[図形].tsv" = "C5`tりんごの図形`r`n"
        "見積\c.docx\ページ001.tsv"             = "前`r`nりんご と `"q`"`tセル2`r`n後`r`n"
    }
    foreach ($rel in $files.Keys) {
        $path = Join-Path $root $rel
        New-Item -ItemType Directory -Path (Split-Path $path -Parent) -Force | Out-Null
        [System.IO.File]::WriteAllText($path, $files[$rel], (New-Object System.Text.UTF8Encoding($true)))
    }
    writeSourceFolderFile @((New-Object PSObject -Property @{ Name = "見積"; Path = "D:\元 データ" })) $root
    return (Resolve-Path -LiteralPath $root).ProviderPath
}

Describe "newResultRows" -Tag Io {
    $root = newBookFixture
    $statusFile = Join-Path $TestDrive "none.tsv"
    $settingsFile = Join-Path $TestDrive "none.config"
    $files = getRestrictedTsvFiles @(@{ Name = "見積"; Path = "$root\見積" })
    $result = searchRestricted "りんご" $files $true
    $regex = (newSearchRegex "りんご" $true $false).Regex
    $built = newResultRows $result.Hits $regex @{}
    $rows = $built.Rows

    It "見出し・ファイルごとのまとまり・ヒット・前後の行を並べる" {
        @($rows | ForEach-Object { "$($_.Level)$(if ($_.Hidden) { 'h' })" }) -join "," | Should Be "0,0,1,2h,2h,2h,1,2h,2h,2h,2h,1,0,1,2h,2h"
        ($rows[0].Texts -join "|") | Should Be "ファイル|場所|種別|行|A|B"
        $built.ColumnCount | Should Be 6
    }

    It "まとまりの行: ファイル名とフォルダのリンク・件数" {
        ($rows[1].Texts -join "|") | Should Be "見積\2024\a b.xlsx|フォルダを開く|3 件"
        $rows[1].Links[0].Target | Should Be "D:\元 データ\2024\a b.xlsx"
        $rows[1].Links[1].Target | Should Be "D:\元 データ\2024"
        ($rows[12].Texts -join "|") | Should Be "見積\c.docx|フォルダを開く|1 件"
    }

    It "ヒットの行: 場所・種別・行番号・セル（Excel の囲みの `" を外し、セル内改行を戻す）" {
        ($rows[2].Texts -join "|") | Should Be "a b.xlsx|[シート] 売上 '1'|セル|2|りんご|=SUM(A1)"
        ($rows[6].Texts -join "|") | Should Be "a b.xlsx|[シート] 売上 '1'|セル|3|みかん|青森`nりんご"
        ($rows[13].Texts -join "|") | Should Be "c.docx|[ページ] 1（目安）|本文|2|りんご と `"q`"|セル2"
        $rows[2].HighlightFrom | Should Be 4
        $rows[2].Collapsed | Should Be $true
    }

    It "Excel のヒットは、リンクで該当のシートと、一致した最初のセルを開く（図形は 1 列目のセル番地）" {
        $rows[2].Links[0].Location | Should Be "'売上 ''1'''!A2"
        $rows[6].Links[0].Location | Should Be "'売上 ''1'''!B3"
        $rows[11].Links[0].Location | Should Be "'売上 ''1'''!C5"
        $rows[13].Links[0].Location | Should Be ""
    }

    It "前後の行: 前後 2 行まで（TSV の端では少ない）・灰色・隠す" {
        ($rows[3].Texts -join "|") | Should Be "||前の行|1|品名|金額"
        ($rows[5].Texts -join "|") | Should Be "||後の行|4|x"
        $rows[3].Style | Should Be ${resultStyleContext}
        ($rows[14].Texts -join "|") | Should Be "||前の行|1|前"
    }

    It "元のフォルダが分からなければ、リンクを付けない" {
        $unknown = @(@{ Root = $root; RelPath = "不明\x.docx\ページ001.tsv"; RelDir = "不明"; FileName = "ページ001.tsv"; Book = "x.docx"; Location = "ページ001"; LineNumber = 1; Line = "りんご" })
        $built = newResultRows $unknown $regex @{} 0
        ($built.Rows[1].Texts -join "|") | Should Be "不明\x.docx|元のフォルダが分かりません|1 件"
        $built.Rows[1].Links | Should BeNullOrEmpty
        $built.Rows[2].Links | Should BeNullOrEmpty
    }
}

Describe "saveResultBook / writeResultBook" -Tag Io {
    $root = newBookFixture
    $statusFile = Join-Path $TestDrive "none.tsv"
    $settingsFile = Join-Path $TestDrive "none.config"
    $files = getRestrictedTsvFiles @(@{ Name = "見積"; Path = "$root\見積" })
    $result = searchRestricted "りんご" $files $true
    $regex = (newSearchRegex "りんご" $true $false).Regex
    $dir = Join-Path $TestDrive "検索結果 𠮷"

    It "xlsx（ZIP）を時刻入りの名前で作り、中の部品がそろう（日本語・CP932 に無い文字のフォルダでも作れる）" {
        $book = saveResultBook $result $regex @(, @("検索ワード", "りんご")) $dir
        Test-Path -LiteralPath $book | Should Be $true
        (Split-Path $book -Leaf) | Should Match '^検索結果_\d{8}_\d{6}(_\d+)?\.xlsx$'
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::OpenRead($book)
        try {
            $names = @($zip.Entries | ForEach-Object { $_.FullName })
            foreach ($part in @("[Content_Types].xml", "_rels/.rels", "xl/workbook.xml", "xl/styles.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml", "xl/worksheets/_rels/sheet1.xml.rels")) {
                $names -contains $part | Should Be $true
            }
            $reader = New-Object System.IO.StreamReader($zip.GetEntry("xl/worksheets/sheet1.xml").Open())
            try { [xml]$sheet = $reader.ReadToEnd() } finally { $reader.Dispose() }
            @($sheet.worksheet.sheetData.row).Count | Should Be 16
        } finally {
            $zip.Dispose()
        }
        # 作業フォルダは残さない
        @(Get-ChildItem -LiteralPath $dir -Directory).Count | Should Be 0
    }

    It "同じ秒に作っても上書きしない" {
        $now = Get-Date
        $first = newResultBookPath $dir $now
        Set-Content -LiteralPath $first -Value ""
        newResultBookPath $dir $now | Should Not Be $first
    }

    It "tar.exe が失敗したら、分かるメッセージで例外にする" {
        $tarExe = Join-Path $TestDrive "無い\tar.exe"
        { writeResultBook ([ordered]@{ "a.xml" = "<a/>" }) (Join-Path $dir "失敗.xlsx") } | Should Throw "検索結果のブックを作れませんでした"
    }
}

Describe "removeOldResultBooks" -Tag Io {
    It "以前の結果を消し、開いたまま（消せない）のものは飛ばす" {
        $dir = Join-Path $TestDrive "old"
        New-Item -ItemType Directory -Path "$dir\作成中_1" -Force | Out-Null
        Set-Content -LiteralPath "$dir\検索結果_1.xlsx" -Value "a"
        Set-Content -LiteralPath "$dir\検索結果_2.xlsx" -Value "b"
        $stream = [System.IO.File]::Open("$dir\検索結果_2.xlsx", "Open", "Read", "None")
        try {
            removeOldResultBooks $dir
        } finally {
            $stream.Dispose()
        }
        @(Get-ChildItem -LiteralPath $dir | ForEach-Object { $_.Name }) -join "," | Should Be "検索結果_2.xlsx"
        removeOldResultBooks (Join-Path $TestDrive "無い")
    }
}

Describe "newResultRows（端の場合）" -Tag Io {
    $root = Join-Path $TestDrive "edge"
    New-Item -ItemType Directory -Path "$root\見積\b.xlsx" -Force | Out-Null
    [System.IO.File]::WriteAllText("$root\見積\b.xlsx\売上.tsv", "りんご`t1`r`n", (New-Object System.Text.UTF8Encoding($true)))
    $statusFile = Join-Path $TestDrive "none.tsv"
    $settingsFile = Join-Path $TestDrive "none.config"
    $regex = (newSearchRegex "りんご" $true $false).Regex
    $maps = @{ $root = @{ "見積" = "D:\元" }; "$root\見積" = @{ "見積" = "D:\元" } }

    It "相対フォルダが無い・場所が無い・図形の 1 列目がセル番地でない・TSV が読めない・TSV の行より後ろのヒット" {
        $hits = @(
            @{ Root = $root; RelPath = "見積\b.xlsx\売上.tsv"; RelDir = "見積"; FileName = "売上.tsv"; Book = "b.xlsx"; Location = "売上"; LineNumber = 5; Line = "りんご`t青森$([char]0x2028)ふじ" },
            @{ Root = $root; RelPath = "見積\c.xlsx\売上[図形].tsv"; RelDir = "見積"; FileName = "売上[図形].tsv"; Book = "c.xlsx"; Location = "売上[図形]"; LineNumber = 1; Line = "図形1`tりんご" },
            @{ Root = $root; RelPath = "見積\d.xlsx.tsv"; RelDir = "見積"; FileName = "d.xlsx.tsv"; Book = "d.xlsx"; Location = ""; LineNumber = 1; Line = "りんご" }
        )
        $built = newResultRows $hits $regex $maps
        $rows = $built.Rows
        # 読めない TSV・TSV の行より後ろのヒットは、前後の行を付けない
        ($rows[2].Texts -join "|") | Should Be "b.xlsx|[シート] 売上|セル|5|りんご|青森`nふじ"
        $rows[2].Links[0].Location | Should Be "'売上'!A5"
        @($rows | ForEach-Object { $_.Level }) -join "," | Should Be "0,0,1,0,1,0,1"
        $rows[4].Links[0].Location | Should Be "'売上'!A1"
        $rows[6].Links[0].Location | Should Be ""
    }

    It "インデックスの直下の TSV（相対フォルダが無い）と、ネットワークのパスのインデックス" {
        $hit = @{ Root = "\\server\share\index"; RelPath = "x.docx\ページ001.tsv"; RelDir = ""; FileName = "ページ001.tsv"; Book = "x.docx"; Location = "ページ001"; LineNumber = 1; Line = "りんご" }
        $built = newResultRows @($hit) $regex @{} 1
        $built.Rows[1].Texts[0] | Should Be "x.docx"
    }
}

Describe "writeResultBook（残っていた作業フォルダ）" -Tag Io {
    It "前に残った作業フォルダがあっても作り直す" {
        $dir = Join-Path $TestDrive "leftover"
        New-Item -ItemType Directory -Path (Join-Path $dir "作成中_$PID") -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $dir "作成中_$PID\古い.txt") -Value "x"
        $parts = getResultBookParts (getResultSheetXml @(@{ Style = 1; Level = 0; Texts = @("ファイル") }) 1) (getResultInfoSheetXml @(, @("a", "b"))) 1 1
        writeResultBook $parts (Join-Path $dir "a.xlsx")
        Test-Path -LiteralPath (Join-Path $dir "a.xlsx") | Should Be $true
        Test-Path -LiteralPath (Join-Path $dir "作成中_$PID") | Should Be $false
    }
}
