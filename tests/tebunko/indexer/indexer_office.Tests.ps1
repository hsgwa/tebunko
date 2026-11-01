# indexer.ps1 を本物の Office で端から端まで動かすテストと、抽出結果の期待値（snapshot）のテスト。
# 結合テスト（手動）の「基本の取り込み・2 回目のスキップ・Office のプロセスが残らない」を自動にする（docs/design/testing/index.md）。
#   Office の Describe … Excel・Word・PowerPoint の入った機械で `.\tests\run.ps1 -Tag Office -Path tests\tebunko\indexer\indexer_office.Tests.ps1`。
#     CI のランナーには Office が無く、既定では外れる。
#   Io の Describe … Office の要らない形式（.docx・.pptx・.pdf・.txt）を取り出しの関数と本文インデックスの組み立てで通し、CI で毎回流す。
# 期待値は tests\testdata\snapshot\（作り直し方は docs/design/testing/run.md の「snapshot の作り直し」）。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "$PSScriptRoot\..\..\helpers\indexer.ps1"
    . "$PSScriptRoot\..\..\helpers\snapshot.ps1"
    # ingestFile（Office の要らない形式の取り出し）を呼ぶため、取り出しの部品だけを読み込む
    . "${scriptsDir}\shared\office\office_reader.ps1"
    . "${scriptsDir}\shared\office\office_embedded.ps1"
    . "${scriptsDir}\shared\office\office_app.ps1"
    . "${scriptsDir}\tebunko\indexer\extract_office.ps1"
    . "${scriptsDir}\tebunko\indexer\extract_text.ps1"
    . "${scriptsDir}\tebunko\indexer\extract_pdf.ps1"

    $snapshotDir = "${testDataDir}\snapshot"
    $reporterPath = "${scriptsDir}\tebunko\indexer\indexing_reporter.ps1"

    function getOfficeProcessIds {
        # 走らせる前から無かった Excel・Word・PowerPoint のうち、$since より後に起きたものの PID。
        # COM で起きる Office の親は走らせた powershell ではないため、親では絞れない。時刻で絞れば、先に起きていたほかのテストの Office を数えない
        param ([int[]]$before, [datetime]$since)
        $ids = foreach ($process in @(Get-Process -Name EXCEL, WINWORD, POWERPNT -ErrorAction SilentlyContinue)) {
            if ($before -contains $process.Id) { continue }
            try { $startedAt = $process.StartTime } catch { continue }
            if ($startedAt -gt $since) { $process.Id }
        }
        return @($ids)
    }

    function getRunningOfficeProcessIds {
        # 今動いている Excel・Word・PowerPoint の PID（走らせる前の控え）
        return @(Get-Process -Name EXCEL, WINWORD, POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    }

    function replaceSnapshotRoots {
        # 比べる前の置き換え。本文インデックスのファイルに作業フォルダの絶対パスが入るなら、ここで <ROOT> にする
        param ([string]$text, [string[]]$roots)
        foreach ($root in $roots) {
            $text = $text.Replace($root, "<ROOT>")
        }
        return $text
    }
}

Describe "indexer.ps1 を本物の Office で通しで動かす" -Tag Office {
    BeforeAll {
        $source = Join-Path $TestDrive "営業"
        foreach ($sub in @("", "\資料")) { [System.IO.Directory]::CreateDirectory("$source$sub") | Out-Null }
        $files = @{
            "基本.xlsx"      = "Excel\基本.xlsx"
            "基本.docx"      = "Word\基本.docx"
            "基本.pptx"      = "PowerPoint\基本.pptx"
            "資料\旧形式.xls" = "Excel\旧形式.xls"
            "資料\旧形式.doc" = "Word\形式\旧形式.doc"
            "資料\旧形式.ppt" = "PowerPoint\形式\旧形式.ppt"
        }
        foreach ($name in $files.Keys) {
            Copy-Item -LiteralPath "${testDataDir}\office\$($files[$name])" -Destination "$source\$name"
        }
        $other = Join-Path $TestDrive "補助"
        [System.IO.Directory]::CreateDirectory($other) | Out-Null
        [System.IO.File]::WriteAllText("$other\メモ.txt", "山田 補助のメモ`r`n", (New-Object System.Text.UTF8Encoding($false)))

        $root = newRoot
        writeTestSettings $root @(@{ name = "営業"; path = $source; enabled = $true }, @{ name = "補助"; path = $other; enabled = $true })
        $officeBefore = getRunningOfficeProcessIds
        $startedAt = Get-Date
        $firstExit = runIndexer $root @{ Workers = 1 }
        # Office が終わるのを少し待つ（後始末は非同期）
        $deadline = (Get-Date).AddSeconds(15)
        do {
            $leftover = @(getOfficeProcessIds $officeBefore $startedAt)
            if ($leftover.Count -eq 0) { break }
            Start-Sleep -Milliseconds 500
        } while ((Get-Date) -lt $deadline)
        $contentDir = "$root\work\store\content\営業"
    }

    It "終了コード 0 で、取り込み一覧の全行が取り込み済みでエラーの列が空" {
        $firstExit | Should -Be 0
        $status = readTestStatus $root
        $rows = @($status.Rows.Values | Where-Object { $_.相対パス -like "営業\*" })
        $rows.Count | Should -Be 6
        @($rows | Where-Object { $_.状態 -ne ${stateDone} }).Count | Should -Be 0
        @($rows | Where-Object { $_.エラー -ne "" }).Count | Should -Be 0
        $status.Rows["補助\メモ.txt"].状態 | Should -Be ${stateDone}
    }

    It "本文インデックスのファイルが、フォルダごと・拡張子ごとに 1 つだけあり、本文インデックスに入れる前の TSV が残らず、gram 索引がある" {
        function getIndexFileNames ($dir) { return (@(Get-ChildItem -LiteralPath $dir -File -Filter "content_index.*.tsv" | ForEach-Object { $_.Name } | Sort-Object) -join ",") }
        getIndexFileNames $contentDir | Should -Be "content_index.docx.001.tsv,content_index.pptx.001.tsv,content_index.xlsx.001.tsv"
        getIndexFileNames "$contentDir\資料" | Should -Be "content_index.doc.001.tsv,content_index.ppt.001.tsv,content_index.xls.001.tsv"
        # サブフォルダは、元のフォルダの階層の「資料」だけ（元のファイルごとの TSV のフォルダは残らない）
        (@(Get-ChildItem -LiteralPath $contentDir -Directory | ForEach-Object { $_.Name }) -join ",") | Should -Be "資料"
        @(Get-ChildItem -LiteralPath "$contentDir\資料" -Directory).Count | Should -Be 0
        [System.IO.File]::Exists("$root\work\store\gram\営業\toc.tsv") | Should -BeTrue
    }

    It "Office のプロセスが残らない" {
        $leftover.Count | Should -Be 0
    }

    It "本文インデックスのファイルの全文が期待値と同じ（<ext>）" -TestCases @(
        # xlsx の期待値は、Excel の表示形が OS の地域設定に従うため、ja-JP でない機械では比べない
        @{ ext = "xlsx"; folder = "";       expected = "office\xlsx.txt"; skip = ((Get-Culture).Name -ne "ja-JP") }
        @{ ext = "xls";  folder = "\資料"; expected = "office\xls.txt"; skip = $false }
        @{ ext = "doc";  folder = "\資料"; expected = "office\doc.txt"; skip = $false }
        @{ ext = "ppt";  folder = "\資料"; expected = "office\ppt.txt"; skip = $false }
        @{ ext = "docx"; folder = "";       expected = "io\docx.txt"; skip = $false }
        @{ ext = "pptx"; folder = "";       expected = "io\pptx.txt"; skip = $false }
    ) {
        param ($ext, $folder, $expected, $skip)
        if ($skip) {
            Set-ItResult -Skipped -Because "地域設定が ja-JP でないため（Excel の表示形が揺れる）"
            return
        }
        $text = readContentIndexText "$contentDir$folder\content_index.$ext.001.tsv"
        $text | Should -Match "^$([regex]::Escape([string]${contentIndexMark})) 版=2\n"
        assertSnapshot (replaceSnapshotRoots $text @($TestDrive)) "$snapshotDir\$expected"
    }

    It "2 回目は取り込み済みのファイルを飛ばし、Office を起動しない" {
        $before = (readTestStatus $root).Rows["営業\基本.xlsx"].取り込み日時
        $officeBefore = getRunningOfficeProcessIds
        $startedAt = Get-Date
        runIndexer $root @{ Workers = 1 } | Should -Be 0
        (readTestProgress).Detail | Should -Be "更新が必要なファイルはありませんでした"
        (readTestStatus $root).Rows["営業\基本.xlsx"].取り込み日時 | Should -Be $before
        @(getOfficeProcessIds $officeBefore $startedAt).Count | Should -Be 0
    }

    It "設定からフォルダを外して確認で進めると、そのインデックスが消える（補助は残る）" {
        writeTestSettings $root @(@{ name = "補助"; path = (Join-Path $TestDrive "補助"); enabled = $true })
        $answer = @{ Script = $reporterPath; Pattern = '^\s+if \(!\$ch\.Answered\.WaitOne'; Action = { answerIndexingPlan $channel @{ RetryFailed = $false; IncludeCloud = $false } } }
        runIndexer $root @{ Workers = 1; ConfirmTargets = $true } @($answer) | Should -Be 0
        [System.IO.Directory]::Exists("$root\work\store\content\営業") | Should -BeFalse
        [System.IO.Directory]::Exists("$root\work\store\gram\営業") | Should -BeFalse
        $status = readTestStatus $root
        @($status.Rows.Keys | Where-Object { $_ -like "営業\*" }).Count | Should -Be 0
        [System.IO.File]::Exists("$root\work\store\content\補助\content_index.txt.001.tsv") | Should -BeTrue
    }
}

Describe "Office の要らない形式の抽出結果が期待値と同じ" -Tag Io {
    It "<name>" -TestCases @(
        @{ name = "docx"; source = "office\Word\基本.docx" }
        @{ name = "pptx"; source = "office\PowerPoint\基本.pptx" }
        @{ name = "pdf"; source = "office\Word\形式\PDF.pdf" }
        @{ name = "txt"; source = "text\TXT01.txt" }
    ) {
        param ($name, $source)
        $work = Join-Path $TestDrive "抽出_$name"
        # ingestFile は書き出し先を引数に持たず、$tmpDir に書く。利用者の作業フォルダに書かないよう、テストの下に向ける
        $tmpDir = Join-Path $work "tmp"
        $indexFolder = Join-Path $work "営業"
        $book = [System.IO.Path]::GetFileName($source)
        [System.IO.Directory]::CreateDirectory($tmpDir) | Out-Null
        [System.IO.Directory]::CreateDirectory("$indexFolder\$book") | Out-Null
        (ingestFile "${testDataDir}\$source") | Should -BeGreaterThan 0
        Get-ChildItem -LiteralPath $tmpDir -File | Move-Item -Destination "$indexFolder\$book"

        $dest = Join-Path $work "content"
        $result = convertFolderToContentIndex $indexFolder $dest
        $result.Books | Should -Be 1
        $text = readContentIndexText "$dest\content_index.$name.001.tsv"
        $text | Should -Match "^$([regex]::Escape([string]${contentIndexMark})) 版=2\n"
        assertSnapshot (replaceSnapshotRoots $text @($TestDrive)) "$snapshotDir\io\$name.txt"
    }
}
