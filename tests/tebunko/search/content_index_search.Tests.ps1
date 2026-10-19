# 検索用の本文インデックスのファイルの検索（tebunko\search\content_index_search.ps1）と読み書き（index\content_index_store.ps1）のテスト。
# 元の TSV を 1 行ずつ読んで照合した結果（referenceKeys）と、本文インデックスのファイルの検索の結果が同じになることを確かめる
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    function script:toKeys {
        # ヒットを比べられる文字列にする（RelPath・FileName は本文インデックスのファイルのもののため比べない）
        param ($hits)
        return @($hits | ForEach-Object { "{0}|{1}|{2}|{3}|{4}" -f $_.RelDir, $_.Book, $_.Location, $_.LineNumber, $_.Line })
    }

    function script:sortedKeys {
        param ($hits)
        return ((toKeys $hits) | Sort-Object) -join "`n"
    }

    function script:referenceKeys {
        # 元の TSV（<相対フォルダ>\<ファイル名.xlsx>\<場所>.tsv）を StreamReader.ReadLine で 1 行ずつ読み、
        # 検索語の正規表現で照合した結果（本文インデックスのファイルの検索が同じ結果になるべきもの）
        param (
            [string]$tsvRoot,
            [string]$word,
            [bool]$simple,
            [bool]$caseSensitive = $false,
            [string]$fileFilter = "",
            [bool]$includeShapes = $true,
            [bool]$includeComments = $true
        )
        $search = newSearchRegex $word $simple $caseSensitive
        $filter = newFileFilter $fileFilter
        $exclude = newPlaceExclude $includeShapes $includeComments
        $keys = New-Object System.Collections.Generic.List[string]
        foreach ($path in [System.IO.Directory]::GetFiles($tsvRoot, "*.tsv", "AllDirectories")) {
            $bookDir = [System.IO.Path]::GetDirectoryName($path)
            $book = [System.IO.Path]::GetFileName($bookDir)
            $place = convertIndexFileNameToPlace ([System.IO.Path]::GetFileNameWithoutExtension($path))
            if ($filter.Include -and !$filter.Include.IsMatch($book)) { continue }
            if ($filter.Exclude -and $filter.Exclude.IsMatch($book)) { continue }
            if ($exclude -and $exclude.IsMatch($place)) { continue }
            $cellPrefixed = $book -match '\.xls[a-z]?$' -and $place -match '\[(図形|コメント)\]$'
            $relDir = [System.IO.Path]::GetDirectoryName($bookDir).Substring($tsvRoot.Length).Trim("\")
            $reader = [System.IO.StreamReader]::new($path, [System.Text.Encoding]::UTF8, $true)
            try {
                $number = 0
                while ($null -ne ($line = $reader.ReadLine())) {
                    $number++
                    # Excel の図形・コメントの行は「セル番地 + タブ + 文字」のため、最初のタブより後から始まる一致だけをヒットにする（元のファイル名と場所の名前だけで判定する）
                    if ($cellPrefixed) {
                        $tab = $line.IndexOf("`t")
                        $hit = $tab -ge 0 -and $search.Regex.IsMatch($line, $tab + 1)
                    } else {
                        $hit = $search.Regex.IsMatch($line)
                    }
                    if ($hit) { $keys.Add("$relDir|$book|$place|$number|$line") }
                }
            } finally {
                $reader.Dispose()
            }
        }
        return ($keys | Sort-Object) -join "`n"
    }
}

Describe "本文インデックスのファイルの作成と検索" -Tag Io {
    BeforeDiscovery {
        $cases = @(
            @{ Word = "単価"; Simple = $true },
            @{ Word = "存在しない語"; Simple = $true },
            @{ Word = "abc"; Simple = $true },
            @{ Word = "^単価"; Simple = $false },
            @{ Word = "税抜$"; Simple = $false },
            @{ Word = "品名\s数量"; Simple = $false },
            @{ Word = "単[^x]"; Simple = $false },
            @{ Word = "(?!予備)単価"; Simple = $false },
            @{ Word = "^$"; Simple = $false },
            @{ Word = "営業"; Simple = $true },
            @{ Word = "シート"; Simple = $true }
        )
    }

    BeforeAll {
        $tsvRoot = Join-Path $TestDrive "tsv"
        $contentIndexRoot = Join-Path $TestDrive "content_index"
        $idx = "$tsvRoot\営業"
        newTsv "$idx\見積.xlsx\$(toIndexFileName "見積")" @("品名`t数量`t単価", "", "りんご`t10`t100", "ABC`tabc")
        newTsv "$idx\見積.xlsx\$(toIndexFileName "見積[図形]")" @("E2`t承認済み 単価")
        newTsv "$idx\見積.xlsx\$(toIndexFileName "見積[ヘッダー・フッター]")" @("単価表", "社外秘")
        newTsv "$idx\見積.xlsx\$(toIndexFileName "50%引き")" @("単価は税抜")
        newTsv "$idx\議事録.docx\$(toIndexFileName "ページ001")" @("見積の方針", "単価は据え置き")
        newTsv "$idx\議事録.docx\$(toIndexFileName "ページ001[コメント]")" @("単価の確認")
        newTsv "$idx\2025\提案.pptx\$(toIndexFileName "スライド001_ノート")" @("単価の背景")
        newTsv "$idx\2025\提案.pptx\$(toIndexFileName "スライド002（非表示）")" @("予備", "単価")
        $contentIndexFiles = newContentIndexFiles $tsvRoot $contentIndexRoot
    }

    It "フォルダごと・拡張子ごとに本文インデックスのファイルを作り、フォルダの順・名前の順に並べる" {
        $contentIndexFiles.Count | Should -Be 3
        $contentIndexFiles[0].RelPath | Should -Be "営業\content_index.docx.001.tsv"
        $contentIndexFiles[1].RelPath | Should -Be "営業\content_index.xlsx.001.tsv"
        $contentIndexFiles[2].RelPath | Should -Be "営業\2025\content_index.pptx.001.tsv"
    }

    It "ブックの中の場所の並びは、前の名前の付け方（encodeIndexPlace）で並べたときと同じになる（検索結果の順を変えないため）" {
        $books = getIndexFolderBooks $idx
        $book = $books | Where-Object { $_.Name -eq "見積.xlsx" }
        $places = @("見積", "見積[図形]", "見積[ヘッダー・フッター]", "50%引き")
        $keys = @($places | ForEach-Object { "{0}.tsv" -f (encodeIndexPlace $_) })
        [System.Array]::Sort($keys, $places, [System.StringComparer]::CurrentCultureIgnoreCase)
        @($book.Places.Place) -join "|" | Should -Be ($places -join "|")
    }

    It "<Name> は、実際のファイル名（英語の固定名）の順ではなく、前の名前の付け方（encodeIndexPlace）の順で並ぶ" -TestCases @(
        @{
            Name = "Word の固定の場所（新しい名前では doc_footnotes が先頭になる組み合わせ）"
            Book = "議事録2.docx"
            Places = @("ページ001", "ヘッダー・フッター", "脚注", "文書[コメント]")
        }
        @{
            Name = "PowerPoint の非表示・ノート（_hidden と _notes の順が逆になる組み合わせ）"
            Book = "提案2.pptx"
            Places = @("スライド001", "スライド001（非表示）", "スライド001_ノート")
        }
    ) {
        param ($Name, $Book, $Places)

        $root = Join-Path $TestDrive "order-$($Book -replace '\.', '-')"
        foreach ($place in $Places) {
            newTsv "$root\$Book\$(toIndexFileName $place)" @("ダミー")
        }

        $legacyOrder = [string[]]$Places.Clone()
        $legacyKeys = [string[]]@($Places | ForEach-Object { "{0}.tsv" -f (encodeIndexPlace $_) })
        [System.Array]::Sort($legacyKeys, $legacyOrder, [System.StringComparer]::CurrentCultureIgnoreCase)

        $realOrder = [string[]]$Places.Clone()
        $realKeys = [string[]]@($Places | ForEach-Object { toIndexFileName $_ })
        [System.Array]::Sort($realKeys, $realOrder, [System.StringComparer]::CurrentCultureIgnoreCase)

        # この組み合わせでは、実際のファイル名の順と前の名前の付け方の順が異なる（そうでないと並べ替えの鍵の取り違えを見逃す）
        ($realOrder -join "|") | Should -Not -Be ($legacyOrder -join "|")

        $book = (getIndexFolderBooks $root) | Where-Object { $_.Name -eq $Book }
        @($book.Places.Place) -join "|" | Should -Be ($legacyOrder -join "|")
    }

    It "元のファイルが無くなった拡張子の本文インデックスのファイルは、変換し直すときに消す" {
        $dest = Join-Path $TestDrive "reconvert"
        [void][System.IO.Directory]::CreateDirectory($dest)
        writeContentIndexFile "$dest\content_index.pptx.001.tsv" (convertToContentIndexText @(@{ Name = "古い.pptx"; Places = @(@{ Place = "スライド001"; Text = "古い" }) }))
        $result = convertFolderToContentIndex $idx $dest @("古い.pptx")
        $result.Files | Should -Be 2
        @([System.IO.Directory]::GetFiles($dest) | ForEach-Object { [System.IO.Path]::GetFileName($_) } | Sort-Object) -join "," | Should -Be "content_index.docx.001.tsv,content_index.xlsx.001.tsv"
    }

    It "本文インデックスのファイルは UTF-16LE（BOM 付き）で、一時ファイルを残さない" {
        $bytes = [System.IO.File]::ReadAllBytes($contentIndexFiles[0].Path)
        "{0:X2}{1:X2}" -f $bytes[0], $bytes[1] | Should -Be "FFFE"
        @(Get-ChildItem $contentIndexRoot -Recurse -Filter "*.tmp").Count | Should -Be 0
    }

    It "「<Word>」の結果が、TSV を 1 行ずつ照合したときと同じ" -ForEach $cases {
        sortedKeys (searchContentIndex $Word $contentIndexFiles $Simple).Hits | Should -BeExactly (referenceKeys $tsvRoot $Word $Simple)
    }

    It "大文字・小文字の区別、図形・コメントの除外、対象ファイルの条件が、TSV を 1 行ずつ照合したときと同じ" {
        sortedKeys (searchContentIndex "ABC" $contentIndexFiles $true -caseSensitive $true).Hits | Should -BeExactly (referenceKeys $tsvRoot "ABC" $true $true)
        sortedKeys (searchContentIndex "単価" $contentIndexFiles $true -includeShapes $false -includeComments $false).Hits |
            Should -BeExactly (referenceKeys $tsvRoot "単価" $true $false "" $false $false)
        sortedKeys (searchContentIndex "単価" $contentIndexFiles $true -fileFilter "*.docx").Hits | Should -BeExactly (referenceKeys $tsvRoot "単価" $true $false "*.docx")
    }

    It "Excel のヘッダー・フッターは、図形・コメントを除いた検索でも、場所の名前つきでヒットする" {
        $hits = (searchContentIndex "社外秘" $contentIndexFiles $true -includeShapes $false -includeComments $false).Hits
        toKeys $hits | Should -Be @("営業|見積.xlsx|見積[ヘッダー・フッター]|2|社外秘")
    }

    It "フォルダの一部・直下だけ・無いフォルダを列挙できる。元のファイルが無いフォルダは作らない" {
        (findContentIndexFiles $contentIndexRoot "営業" $false).Count | Should -Be 2
        (findContentIndexFiles $contentIndexRoot "営業\2025" $true).Count | Should -Be 1
        (findContentIndexFiles $contentIndexRoot "無いフォルダ").Count | Should -Be 0
        $empty = Join-Path $TestDrive "empty"
        [void][System.IO.Directory]::CreateDirectory($empty)
        (convertFolderToContentIndex $empty (Join-Path $TestDrive "empty_content_index")).Books | Should -Be 0
        [System.IO.Directory]::Exists((Join-Path $TestDrive "empty_content_index")) | Should -Be $false
    }

    It "更新日時が変わった本文インデックスのファイルは読み直し、キャッシュの古い内容を置き換える" {
        $cache = newTsvTextCache
        [void](searchContentIndex "単価" $contentIndexFiles $true -cache $cache)
        $chars = $cache.Chars[0]
        $changed = @($contentIndexFiles | ForEach-Object { $copy = $_.Clone(); $copy.Ticks = $_.Ticks + 1; $copy })
        (toKeys (searchContentIndex "単価" $changed $true -cache $cache).Hits).Count | Should -Be 8
        $cache.Texts.Count | Should -Be 3
        $cache.Chars[0] | Should -Be $chars
        $cache.Texts[$contentIndexFiles[0].Path][0] | Should -Be ($contentIndexFiles[0].Ticks + 1)
    }

    It "中止を求められたら止める" {
        $result = searchContentIndex "単価" $contentIndexFiles $true -stopSignal @{ Stop = $true }
        $result.Cancelled | Should -Be $true
        $result.Hits.Count | Should -Be 0
    }

    It "照合は本文インデックスのファイル 1 件ごとに Stop を見て、立ったら次のファイルを読まずにここまでのヒットを返す" {
        $search = newSearchRegex "単価" $true $false
        $firstOnly = searchContentIndexFiles $contentIndexFiles 0 1 $search.Regex -1 $search.TextRegex $search.ScanMode
        $firstOnly.Hits.Count | Should -BeGreaterThan 0
        # 2 回目に読まれたときから Stop が立つ入れ物（1 件目を照合し、2 件目の前で止まる）
        $signal = [pscustomobject]@{ Reads = 0 }
        $signal | Add-Member -MemberType ScriptProperty -Name Stop -Value { $this.Reads++; $this.Reads -ge 2 }
        $found = searchContentIndexFiles $contentIndexFiles 0 $contentIndexFiles.Count $search.Regex -1 $search.TextRegex $search.ScanMode $null $null $null $null $signal
        $found.Hits.Count | Should -Be $firstOnly.Hits.Count
        $signal.Reads | Should -Be 2
        (searchContentIndexFiles $contentIndexFiles 0 $contentIndexFiles.Count $search.Regex -1 $search.TextRegex $search.ScanMode $null $null $null $null @{ Stop = $true }).Hits.Count | Should -Be 0
    }

    It "照合のスレッドのスクリプトも、渡された Stop を見て止まる（このスレッドで直接動かす）" {
        $search = newSearchRegex "単価" $true $false
        $firstOnly = searchContentIndexFiles $contentIndexFiles 0 1 $search.Regex -1 $search.TextRegex $search.ScanMode
        $signal = [pscustomobject]@{ Reads = 0 }
        $signal | Add-Member -MemberType ScriptProperty -Name Stop -Value { $this.Reads++; $this.Reads -ge 2 }
        $output = & ${contentIndexWorkerScript} $contentIndexFiles 0 $contentIndexFiles.Count $search.Regex -1 $search.TextRegex $search.ScanMode $null $null $null $null $signal
        $output.Timeout | Should -Be $false
        @($output.Hits).Count | Should -Be $firstOnly.Hits.Count
        $output = & ${contentIndexWorkerScript} $contentIndexFiles 0 $contentIndexFiles.Count $search.Regex -1 $search.TextRegex $search.ScanMode $null $null $null $null @{ Stop = $true }
        @($output.Hits).Count | Should -Be 0
    }

    It "1 つのスレッドで作業が 1 つ（ファイル全部）でも、作業の途中で止まる。止めた作業の分は進み具合として知らせない" {
        # 止めない入れ物で、Stop を見る回数を数える（最後まで照合したときの回数）
        $counter = [pscustomobject]@{ Reads = 0 }
        $counter | Add-Member -MemberType ScriptProperty -Name Stop -Value { $this.Reads++; $false }
        [void](searchContentIndex "単価" $contentIndexFiles $true -workerCount 1 -stopSignal $counter)
        # 3 回目に読まれたときから Stop が立つ。作業の途中で止まったことは、読んだ回数が最後まで照合したときより少ないことで確かめる
        $signal = [pscustomobject]@{ Reads = 0 }
        $signal | Add-Member -MemberType ScriptProperty -Name Stop -Value { $this.Reads++; $this.Reads -ge 3 }
        $script:progress = 0
        $result = searchContentIndex "単価" $contentIndexFiles $true -workerCount 1 -stopSignal $signal -onProgress { param($done, $total, $newHits) $script:progress++ }
        $result.Cancelled | Should -Be $true
        $signal.Reads | Should -BeLessThan $counter.Reads
        $result.Hits.Count | Should -Be 0   # 止めた作業のヒットは、結果にも進み具合にも入れない
        $script:progress | Should -Be 0
    }

    It "列挙は検索対象のフォルダ 1 つごとに Stop を見て、立ったら残りを列挙しない" {
        $targets = @(@{ Root = $contentIndexRoot; RelPath = "営業"; Recurse = $false }, @{ Root = $contentIndexRoot; RelPath = "営業\2025"; Recurse = $false })
        $stop = @{ Stop = $false }
        $listed = getContentIndexFiles $targets { param ($count) $stop.Stop = $true } $stop
        $listed.Folders.Count | Should -Be 1
        $listed.ContentIndexFiles.Count | Should -Be 2
        $none = getContentIndexFiles $targets $null @{ Stop = $true }
        $none.Folders.Count | Should -Be 0
        (getContentIndexFiles $targets).Folders.Count | Should -Be 2
    }

    It "1 行ずつ照合する検索語でも上限で打ち切る" {
        $result = searchContentIndex "(?!予備)単価" $contentIndexFiles $false 2
        $result.Hits.Count | Should -Be 2
        $result.Truncated | Should -Be $true
    }

    It "全文への照合が時間切れになったファイルは、1 行ずつ照合し直す" {
        # 1 行は短いため 1 行ずつなら速いが、全文には時間がかかる正規表現（全文の方だけ時間切れを短くする）
        $slowRoot = Join-Path $TestDrive "slow"
        newTsv "$slowRoot\idx\遅い.xlsx\$(toIndexFileName "S")" (@(1..3000 | ForEach-Object { "aaaaaaaaaaaaaaaaaaaaaa!" }) + @("aaab"))
        convertFolderToContentIndex "$slowRoot\idx" "$slowRoot\content_index\idx" | Out-Null
        $slowContentIndexFiles = findContentIndexFiles "$slowRoot\content_index"
        $pattern = "^(a|aa)+b"
        $options = [System.Text.RegularExpressions.RegexOptions]::None
        $lineRegex = [regex]::new($pattern, $options, [timespan]::FromSeconds(5))
        $textRegex = [regex]::new($pattern, $options -bor [System.Text.RegularExpressions.RegexOptions]::Multiline, [timespan]::FromMilliseconds(1))
        $hits = (searchContentIndexFiles $slowContentIndexFiles 0 1 $lineRegex -1 $textRegex "lines").Hits
        $hits.Count | Should -Be 1
        $hits[0].LineNumber | Should -Be 3001
    }

    It "置かれた TSV を本文インデックスのファイルに入れて TSV を消し、変わらない元のファイルは前の本文インデックスのファイルから写す" {
        $dir = Join-Path $TestDrive "inplace\営業"
        newTsv "$dir\A.xlsx\$(toIndexFileName "S")" @("A の単価")
        newTsv "$dir\B.xlsx\$(toIndexFileName "S")" @("B の古い単価")
        (updateFolderContentIndex $dir).Books | Should -Be 2
        @([System.IO.Directory]::GetDirectories($dir)).Count | Should -Be 0
        # B を更新し、C（Word）を足す。A の TSV はもう無い
        newTsv "$dir\B.xlsx\$(toIndexFileName "S")" @("B の新しい単価")
        newTsv "$dir\C.docx\$(toIndexFileName "ページ001")" @("C の単価")
        $result = updateFolderContentIndex $dir
        $result.Books | Should -Be 3
        $result.Tsv | Should -Be 2
        @([System.IO.Directory]::GetDirectories($dir)).Count | Should -Be 0
        $inContentIndexFiles = findContentIndexFiles (Join-Path $TestDrive "inplace")
        @($inContentIndexFiles | ForEach-Object { [System.IO.Path]::GetFileName($_.RelPath) }) -join "," | Should -Be "content_index.docx.001.tsv,content_index.xlsx.001.tsv"
        (toKeys (searchContentIndex "単価" $inContentIndexFiles $true).Hits) -join "`n" | Should -BeExactly ((
            "営業|C.docx|ページ001|1|C の単価", "営業|A.xlsx|S|1|A の単価", "営業|B.xlsx|S|1|B の新しい単価") -join "`n")
        # C が無くなったら外し、Word の本文インデックスのファイルを消す
        (updateFolderContentIndex $dir @("C.docx")).Books | Should -Be 2
        @([System.IO.Directory]::GetFiles($dir) | ForEach-Object { [System.IO.Path]::GetFileName($_) }) -join "," | Should -Be "content_index.xlsx.001.tsv"
    }

    It "TSV の残ったフォルダを見つけ、本文インデックスのファイルに書き出して TSV を消す" {
        $work = Join-Path $TestDrive "publish"
        $index = "$work\index"
        newTsv "$index\人事\A.xlsx\$(toIndexFileName "S")" @("採用の計画")
        newTsv "$index\人事\2025\B.docx\$(toIndexFileName "ページ001")" @("評価の方針")
        # 本文インデックスのファイルは、元のファイルごとのフォルダではない
        writeContentIndexFile "$index\人事\2025\content_index.xlsx.001.tsv" (convertToContentIndexText @(@{ Name = "C.xlsx"; Places = @(@{ Place = "S"; Text = "既に入っている" }) }))
        $found = findIndexFoldersWithBooks $index
        @($found | ForEach-Object { $_.Substring($index.Length) }) -join "," | Should -Be "\人事,\人事\2025"
        (findIndexFoldersWithBooks "$work\無い").Count | Should -Be 0

        $pending = New-Object 'System.Collections.Generic.Dictionary[string,object]'
        foreach ($folder in $found) { $pending[$folder] = @() }
        $published = publishIndexFolders $pending $index
        $published.Count | Should -Be 2
        @($published.Skipped).Count | Should -Be 0
        (findIndexFoldersWithBooks $index).Count | Should -Be 0
        [System.IO.File]::Exists("$index\人事\content_index.xlsx.001.tsv") | Should -Be $true
        [System.IO.File]::Exists("$index\人事\2025\content_index.docx.001.tsv") | Should -Be $true
    }

    It "途中にジャンクションがあるフォルダは、リンクの先の本文インデックスのファイルを消さず・書かず、飛ばしてログに出す" {
        Mock writeIndexerLog {}
        $work = Join-Path $TestDrive "publink"
        $index = "$work\index"
        $outside = "$work\外"
        # 外のフォルダに、前の本文インデックスのファイルだけがある（元のファイルが無くなって、中身が空になる場面）
        [System.IO.Directory]::CreateDirectory("$outside\deep") | Out-Null
        writeContentIndexFile "$outside\deep\content_index.xlsx.001.tsv" (convertToContentIndexText @(@{ Name = "C.xlsx"; Places = @(@{ Place = "S"; Text = "外の中身" }) }))
        [System.IO.Directory]::CreateDirectory($index) | Out-Null
        if (!(newTestJunction "$index\営業" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        $pending = New-Object 'System.Collections.Generic.Dictionary[string,object]'
        $pending["$index\営業\deep"] = @("C.xlsx")

        (publishIndexFolders $pending $index).Count | Should -Be 0

        [System.IO.File]::Exists("$outside\deep\content_index.xlsx.001.tsv") | Should -Be $true
        Should -Invoke writeIndexerLog -Times 1 -Exactly -Scope It -ParameterFilter { $text -like "*書き出しませんでした*" }
    }

    It "名前が .xlsx などで終わる本物のフォルダ・中身が空のファイルのフォルダは、本文インデックスにする前の TSV と取り違えない" {
        $index = Join-Path $TestDrive "bookdir\index"
        $folder = "$index\営業"
        # 元のフォルダに「資料.xlsx」という名前のフォルダがあり、その中の本文インデックスのファイルがある
        [void][System.IO.Directory]::CreateDirectory("$folder\資料.xlsx")
        writeContentIndexFile "$folder\資料.xlsx\content_index.docx.001.tsv" (convertToContentIndexText @(@{ Name = "中の文書.docx"; Places = @(@{ Place = "ページ001"; Text = "中の文書" }) }))
        # 取り込んだが中身が空のファイル（フォルダだけ残る）と、本文インデックスにする前の TSV
        [void][System.IO.Directory]::CreateDirectory("$folder\空.xlsx")
        newTsv "$folder\B.xlsx\$(toIndexFileName "S")" @("B の中身")
        # 名前が x.log で終わる本物のフォルダ（テキストの拡張子を足したことで当たりやすくなった名前。中身（サブフォルダ）で見分ける）
        [void][System.IO.Directory]::CreateDirectory("$folder\x.log\サブ")

        testIndexBookDir "$folder\資料.xlsx" | Should -Be $false
        testIndexBookDir "$folder\空.xlsx" | Should -Be $false
        testIndexBookDir "$folder\空.xlsx" $false | Should -Be $true
        testIndexBookDir "$folder\B.xlsx" | Should -Be $true
        testIndexBookDir "$folder\無い.xlsx" | Should -Be $false
        testIndexBookDir "$folder\x.log" | Should -Be $false
        @((findIndexFoldersWithBooks $index) | ForEach-Object { $_.Substring($index.Length) }) -join "," | Should -Be "\営業"

        $result = updateFolderContentIndex $folder
        $result.Books | Should -Be 1
        [System.IO.File]::Exists("$folder\資料.xlsx\content_index.docx.001.tsv") | Should -Be $true
        [System.IO.Directory]::Exists("$folder\空.xlsx") | Should -Be $true
        [System.IO.Directory]::Exists("$folder\B.xlsx") | Should -Be $false
        (findIndexFoldersWithBooks $index).Count | Should -Be 0
    }

    It "大きさの上限を超えたら次の番号の本文インデックスのファイルに分け、変わった本文インデックスのファイルだけを書き直す。検索の結果は変わらない" {
        $dir = Join-Path $TestDrive "split\営業"
        foreach ($i in 1..5) { newTsv ("$dir\資料{0}.xlsx\S.tsv" -f $i) @(("行 $i " + ("あ" * 600)), "単価 $i") }
        # 1 冊は約 1.3KB。上限 2KB なら 2 冊ずつの本文インデックスのファイルに分かれる
        $result = convertFolderToContentIndex $dir $dir @() $true 2048
        $result.Files | Should -Be 3
        $result.Texts.Count | Should -Be 3
        @([System.IO.Directory]::GetFiles($dir) | ForEach-Object { [System.IO.Path]::GetFileName($_) } | Sort-Object) -join "," | Should -Be "content_index.xlsx.001.tsv,content_index.xlsx.002.tsv,content_index.xlsx.003.tsv"
        $splitContentIndexFiles = findContentIndexFiles (Join-Path $TestDrive "split")
        (toKeys (searchContentIndex "単価" $splitContentIndexFiles $true).Hits | ForEach-Object { ($_ -split "\|")[1] }) -join "," | Should -Be "資料1.xlsx,資料2.xlsx,資料3.xlsx,資料4.xlsx,資料5.xlsx"

        # 資料1 を更新し、資料6 を足す。資料1 の本文インデックスのファイル（001）と最後の本文インデックスのファイル（003）だけを書き直す
        $before = @{}
        foreach ($p in $splitContentIndexFiles) { $before[[System.IO.Path]::GetFileName($p.RelPath)] = $p.Ticks }
        Start-Sleep -Milliseconds 20
        newTsv "$dir\資料1.xlsx\S.tsv" @("更新した単価")
        newTsv "$dir\資料6.xlsx\S.tsv" @("単価 6")
        $result = convertFolderToContentIndex $dir $dir @() $true 2048
        $result.Written | Should -Be 2
        $result.Texts.Count | Should -Be 3
        $after = @{}
        foreach ($p in (findContentIndexFiles (Join-Path $TestDrive "split"))) { $after[[System.IO.Path]::GetFileName($p.RelPath)] = $p.Ticks }
        ($after["content_index.xlsx.002.tsv"] -eq $before["content_index.xlsx.002.tsv"]) | Should -Be $true
        ($after["content_index.xlsx.001.tsv"] -ne $before["content_index.xlsx.001.tsv"]) | Should -Be $true
        (searchContentIndex "単価" (findContentIndexFiles (Join-Path $TestDrive "split")) $true).Hits.Count | Should -Be 6

        # 2 冊とも無くなった本文インデックスのファイル（002）は消す
        $result = convertFolderToContentIndex $dir $dir @("資料3.xlsx", "資料4.xlsx") $true 2048
        [System.IO.File]::Exists("$dir\content_index.xlsx.002.tsv") | Should -Be $false
        $result.Files | Should -Be 2
    }

    It "書き直すと中身が置き換わる" {
        $path = Join-Path $TestDrive "rewrite\$(getContentIndexFileName "xlsx")"
        [void][System.IO.Directory]::CreateDirectory((Split-Path $path))
        writeContentIndexFile $path "a"
        writeContentIndexFile $path "b"
        readContentIndexText $path | Should -BeExactly "b"
        [System.IO.File]::Exists("$path.tmp") | Should -Be $false
    }
}

Describe "searchContentIndex（埋め込みの場所。[図形も検索] を切ると除かれる）" -Tag Io {
    BeforeAll {
        $embedRoot = Join-Path $TestDrive "embed_tsv"
        $embedDir = "$embedRoot\idx"
        newTsv "$embedDir\報告.docx\$(toIndexFileName "ページ001")" @("本文の合計")
        newTsv "$embedDir\報告.docx\$(toIndexFileName "ページ001[埋め込み1]")" @("品名`t合計", "りんご`t合計 100")
        newTsv "$embedDir\資料.pptx\$(toIndexFileName "スライド002[埋め込み12]")" @("埋め込みの合計")
        $embedContentIndexFiles = newContentIndexFiles $embedRoot (Join-Path $TestDrive "embed_content_index")
    }

    It "既定では、埋め込みの中の文字もヒットになり、場所の名前に [埋め込みN] が付く" {
        sortedKeys (searchContentIndex "合計" $embedContentIndexFiles $true).Hits | Should -BeExactly ((@(
            "idx|報告.docx|ページ001|1|本文の合計"
            "idx|報告.docx|ページ001[埋め込み1]|1|品名`t合計"
            "idx|報告.docx|ページ001[埋め込み1]|2|りんご`t合計 100"
            "idx|資料.pptx|スライド002[埋め込み12]|1|埋め込みの合計"
        ) | Sort-Object) -join "`n")
    }

    It "図形を除いた検索では、埋め込みの中の文字は出ない。TSV を 1 行ずつ照合したときと同じ" {
        $hits = (searchContentIndex "合計" $embedContentIndexFiles $true -includeShapes $false).Hits
        toKeys $hits | Should -Be @("idx|報告.docx|ページ001|1|本文の合計")
        sortedKeys $hits | Should -BeExactly (referenceKeys $embedRoot "合計" $true $false "" $false $true)
    }
}

Describe "searchContentIndex（テキストの長い行を切る。search_query.ps1 の truncateHitLine）" -Tag Io {
    BeforeAll {
        $tsvRoot = Join-Path $TestDrive "tsv-long"
        $contentIndexRoot = Join-Path $TestDrive "content_index-long"
        $idx = "$tsvRoot\長い行"

        # 20万文字ほどの1行の中ほどに検索語を1つ置く（テキストは切る対象、Excel は切らない対象）
        $filler = "あ" * 100000
        $longLine = $filler + "TARGET" + $filler
        newTsv "$idx\資料.log\$(toIndexFileName "本文")" @($longLine)
        newTsv "$idx\資料.xlsx\$(toIndexFileName "シート1")" @($longLine)
        $contentIndexFiles = newContentIndexFiles $tsvRoot $contentIndexRoot
    }

    It "テキスト（.log）のヒットの行は1,000文字＋…以下に切られ、一致した語は残る" {
        $hits = @((searchContentIndex "TARGET" $contentIndexFiles $true).Hits | Where-Object { $_.Book -eq "資料.log" })
        $hits.Count | Should -Be 1
        $hits[0].Line.Length | Should -BeLessOrEqual 1002
        $hits[0].Line | Should -Match "TARGET"
        $hits[0].Line.StartsWith("…") | Should -Be $true
        $hits[0].Line.EndsWith("…") | Should -Be $true
    }

    It "Excel のヒットの行は切られない（列の位置が変わるため）" {
        $hits = @((searchContentIndex "TARGET" $contentIndexFiles $true).Hits | Where-Object { $_.Book -eq "資料.xlsx" })
        $hits.Count | Should -Be 1
        $hits[0].Line.Length | Should -Be $longLine.Length
    }
}

Describe "searchContentIndex（改行の種類・照合のしかたによらず、TSV を 1 行ずつ照合したときと同じ）" -Tag Io {
    BeforeDiscovery {
        $cases = @(
            @("abc", $true), @("見積", $true), @("(株)", $true), @("存在しない", $true), @("`t", $true),
            @("^abc", $false), @("abc$", $false), @("^$", $false), @("x*", $false), @("見積.*確定", $false),
            @("見積\s確定", $false), @("\s", $false), @("[^a]", $false), @("abc(?=\s)", $false),
            @("abc(?!\s)", $false), @("\Aabc", $false), @("abc\z", $false), @("(?i)ABC", $false),
            @("A1", $true), @("\d", $false), @("(?i)a1", $false), @("^A1", $false), @("A1\tabc", $false), @("(?<=\t)abc", $false)
        ) | ForEach-Object { @{ Word = $_[0]; Simple = $_[1] } }
    }

    BeforeAll {
        # 改行の種類（CRLF・LF・CR）、末尾の改行の有無、空行、空のファイル、BOM の無いファイルを混ぜる
        $tsvRoot = Join-Path $TestDrive "modes_tsv"
        $dir = "$tsvRoot\idx"
        foreach ($book in "a.xlsx", "b.docx", "c.pptx", "d.xlsx") { [void][System.IO.Directory]::CreateDirectory("$dir\$book") }
        [System.IO.File]::WriteAllText("$dir\a.xlsx\S.tsv", "見積`t(株)山田商事`r`n`r`nabc 見積`r`n確定`tABC`r`n", ${utf8Bom})
        [System.IO.File]::WriteAllText("$dir\b.docx\ページ001.tsv", "りんご`nみかん abc`n`n見積 確定`nabc", ${utf8Bom})
        [System.IO.File]::WriteAllText("$dir\c.pptx\スライド001.tsv", "abc`rりんご`r`r見積`r", (New-Object System.Text.UTF8Encoding $false))
        [System.IO.File]::WriteAllText("$dir\d.xlsx\空.tsv", "", ${utf8Bom})
        # Excel の図形・コメント（セル番地 + タブ + 文字の行）
        newTsv "$dir\e.xlsx\$(toIndexFileName "S[図形]")" @("A1`tabc 見積", "A1`t確定")
        newTsv "$dir\e.xlsx\$(toIndexFileName "S[コメント]")" @("B3`tA1 の件")
        $contentIndexFiles = newContentIndexFiles $tsvRoot (Join-Path $TestDrive "modes_content_index")
    }

    It "「<Word>」（文字どおり=<Simple>）" -ForEach $cases {
        sortedKeys (searchContentIndex $Word $contentIndexFiles $Simple).Hits | Should -BeExactly (referenceKeys $tsvRoot $Word $Simple)
    }

    It "行番号は 1 行ずつ読んだときと同じ（CR だけの改行・空行も数える）" {
        sortedKeys (searchContentIndex "見積" $contentIndexFiles $true).Hits |
            Should -BeExactly ((@("idx|a.xlsx|S|1|見積`t(株)山田商事", "idx|a.xlsx|S|3|abc 見積", "idx|b.docx|ページ001|4|見積 確定", "idx|e.xlsx|S[図形]|1|A1`tabc 見積", "idx|c.pptx|スライド001|4|見積") | Sort-Object) -join "`n")
    }
}

Describe "searchContentIndex（Excel の図形・コメントは、セル番地だけに一致した行をヒットにしない）" -Tag Io {
    BeforeAll {
        $tsvRoot = Join-Path $TestDrive "cell_tsv"
        $dir = "$tsvRoot\idx"
        newTsv "$dir\見積.xlsx\$(toIndexFileName "見積")" @("品名", "C2 番の部品")
        # 4 行目は文字の中にタブがある（本物の書き出しでは " で囲まれる。最初のタブより後は、文字の中のタブも検索の対象）
        newTsv "$dir\見積.xlsx\$(toIndexFileName "見積[図形]")" @("C2`t納期", "C2`tC2 の部品", "B10`t2個", "C3`t`"品名`t納期`"")
        newTsv "$dir\見積.xlsx\$(toIndexFileName "見積[コメント]")" @("C2`t確認済み")
        newTsv "$dir\報告.docx\$(toIndexFileName "ページ001[図形]")" @("C2 と書いた図形")
        newTsv "$dir\見積2.xlsx\$(toIndexFileName "見積2[図形]")" @("D4`tあ", "D4`tい", "E5`tD4 の件")
        $contentIndexFiles = newContentIndexFiles $tsvRoot (Join-Path $TestDrive "cell_content_index")
    }

    It "「<Word>」は、文字に一致する行だけがヒットになる（文字の中のタブも対象）" -ForEach @(
        @{ Word = "C2"; Simple = $true; Expected = @("idx|見積.xlsx|見積|2|C2 番の部品", "idx|見積.xlsx|見積[図形]|2|C2`tC2 の部品", "idx|報告.docx|ページ001[図形]|1|C2 と書いた図形") }
        @{ Word = "2"; Simple = $true; Expected = @("idx|見積.xlsx|見積|2|C2 番の部品", "idx|見積.xlsx|見積[図形]|2|C2`tC2 の部品", "idx|見積.xlsx|見積[図形]|3|B10`t2個", "idx|報告.docx|ページ001[図形]|1|C2 と書いた図形") }
        @{ Word = "C"; Simple = $true; Expected = @("idx|見積.xlsx|見積|2|C2 番の部品", "idx|見積.xlsx|見積[図形]|2|C2`tC2 の部品", "idx|報告.docx|ページ001[図形]|1|C2 と書いた図形") }
        @{ Word = "納期"; Simple = $true; Expected = @("idx|見積.xlsx|見積[図形]|1|C2`t納期", "idx|見積.xlsx|見積[図形]|4|C3`t`"品名`t納期`"") }
        @{ Word = "品名\t納期"; Simple = $false; Expected = @("idx|見積.xlsx|見積[図形]|4|C3`t`"品名`t納期`"") }
        @{ Word = "\t"; Simple = $false; Expected = @("idx|見積.xlsx|見積[図形]|4|C3`t`"品名`t納期`"") }
        @{ Word = "(?<=\t)納期"; Simple = $false; Expected = @("idx|見積.xlsx|見積[図形]|1|C2`t納期", "idx|見積.xlsx|見積[図形]|4|C3`t`"品名`t納期`"") }
        @{ Word = "C2\t納期"; Simple = $false; Expected = @() }
        @{ Word = "^C2"; Simple = $false; Expected = @("idx|見積.xlsx|見積|2|C2 番の部品", "idx|報告.docx|ページ001[図形]|1|C2 と書いた図形") }
        @{ Word = "確認済み"; Simple = $true; Expected = @("idx|見積.xlsx|見積[コメント]|1|C2`t確認済み") }
    ) {
        sortedKeys (searchContentIndex $Word $contentIndexFiles $Simple).Hits | Should -BeExactly (($Expected | Sort-Object) -join "`n")
    }

    It "除いた行があっても、後の行の行番号がずれない（<Word>）" -ForEach @(
        @{ Word = "D4"; Simple = $true }
        @{ Word = "(?i)D4"; Simple = $false }
    ) {
        toKeys (searchContentIndex $Word $contentIndexFiles $Simple).Hits | Should -Be @("idx|見積2.xlsx|見積2[図形]|3|E5`tD4 の件")
    }

    Context "複数行・囲みのある文字" {
        BeforeAll {
            # 本物の書き出し（toObjectCellText）と同じ形: 改行はセル内改行（U+2028）、改行・"・タブを含む文字は " で囲み、中の " は "" にする
            $lsep = [string][char]0x2028
            $multiRoot = Join-Path $TestDrive "multi_tsv"
            newTsv "$multiRoot\idx\複数.xlsx\$(toIndexFileName "S[図形]")" @("C2`t`"納期${lsep}C2 の部品`"", "D5`t`"彼は`"`"はい`"`"と言った`"")
            $multiContentIndexFiles = newContentIndexFiles $multiRoot (Join-Path $TestDrive "multi_content_index")
            $line1 = "idx|複数.xlsx|S[図形]|1|C2`t`"納期${lsep}C2 の部品`""
            $line2 = "idx|複数.xlsx|S[図形]|2|D5`t`"彼は`"`"はい`"`"と言った`""
        }

        It "「<Word>」の結果が、手で書いた期待値と同じ" -ForEach @(
            @{ Word = "C2"; Simple = $true; Which = @(1) }
            @{ Word = "納期"; Simple = $true; Which = @(1) }
            @{ Word = "の部品"; Simple = $true; Which = @(1) }
            @{ Word = "D5"; Simple = $true; Which = @() }
            @{ Word = "彼は"; Simple = $true; Which = @(2) }
            @{ Word = "`"`""; Simple = $true; Which = @(2) }
            @{ Word = "`""; Simple = $true; Which = @(1, 2) }
            @{ Word = "`"納期"; Simple = $true; Which = @(1) }
        ) {
            $all = @{ 1 = $line1; 2 = $line2 }
            $expected = @($Which | ForEach-Object { $all[$_] })
            sortedKeys (searchContentIndex $Word $multiContentIndexFiles $Simple).Hits | Should -BeExactly (($expected | Sort-Object) -join "`n")
        }
    }

    Context "件数と上限" {
        BeforeAll {
            # 各フォルダに、セル番地だけに一致する行が 3 行、その後に文字に一致する行が 2 行
            $limitRoot = Join-Path $TestDrive "limit_tsv"
            foreach ($folder in "a", "b", "c") {
                newTsv "$limitRoot\$folder\上限.xlsx\$(toIndexFileName "S[図形]")" @("F1`tx", "F1`ty", "F1`tz", "G1`tF1 あ", "G1`tF1 い")
            }
            $limitContentIndexFiles = newContentIndexFiles $limitRoot (Join-Path $TestDrive "limit_content_index")
        }

        It "上限 <Limit>・スレッド <Workers> で、件数 <Count>・Truncated=<Truncated>" -ForEach @(
            @{ Limit = 1; Workers = 1; Count = 1; Truncated = $true }
            @{ Limit = 1; Workers = 3; Count = 1; Truncated = $true }
            @{ Limit = 6; Workers = 1; Count = 6; Truncated = $false }
            @{ Limit = 6; Workers = 3; Count = 6; Truncated = $false }
            @{ Limit = 0; Workers = 1; Count = 6; Truncated = $false }
            @{ Limit = 0; Workers = 3; Count = 6; Truncated = $false }
        ) {
            $result = searchContentIndex "F1" $limitContentIndexFiles $true $Limit -workerCount $Workers -taskBytes 1
            $result.Hits.Count | Should -Be $Count
            $result.Truncated | Should -Be $Truncated
            # 返った行は、文字に一致した行（セル番地だけの F1 x・y・z ではない）
            foreach ($hit in $result.Hits) { $hit.Line | Should -Match "^G1`tF1 " }
        }
    }
}

Describe "searchContentIndex（検索語・条件・上限・中止）" -Tag Io {
    BeforeAll {
        $tsvRoot = Join-Path $TestDrive "search_tsv"
        newTsv "$tsvRoot\A社.xlsx\Sheet1.tsv" @("見積先：`t(株)山田商事", "", "株式会社`t1.5", "ABC`t105")
        newTsv "$tsvRoot\sub\文書.docx\ページ001.tsv" @("りんご (株) abc")
        $contentIndexRoot = Join-Path $TestDrive "search_content_index"
        $contentIndexFiles = newContentIndexFiles $tsvRoot $contentIndexRoot
    }

    It "文字どおりに検索すると (株) は記号のまま探す" {
        $result = searchContentIndex "(株)" $contentIndexFiles $true
        $result.Hits.Count | Should -Be 2
        $result.SimpleMatch | Should -Be $true
    }

    It "正規表現として検索すると (株) は「株」に一致する" {
        $result = searchContentIndex "(株)" $contentIndexFiles $false
        $result.Hits.Count | Should -Be 3
        $result.SimpleMatch | Should -Be $false
    }

    It "正規表現として不正なワードは文字どおりに検索する" {
        $result = searchContentIndex "(" $contentIndexFiles $false
        $result.SimpleMatch | Should -Be $true
        $result.Hits.Count | Should -Be 2
    }

    It "ファイル名・場所・行番号・相対フォルダ・インデックスのフォルダを返す（空行も行番号に数える）" {
        $hit = @((searchContentIndex "1.5" $contentIndexFiles $true).Hits)[0]
        $hit.Book | Should -Be "A社.xlsx"
        $hit.Location | Should -Be "Sheet1"
        $hit.LineNumber | Should -Be 3
        $hit.RelDir | Should -Be ""
        $hit.Root | Should -Be $contentIndexRoot
        @((searchContentIndex "りんご" $contentIndexFiles $true).Hits)[0].RelDir | Should -Be "sub"
    }

    It "上限を超えたら打ち切り、上限ちょうどなら打ち切りにしない。上限 0 は上限なし" {
        $result = searchContentIndex "株" $contentIndexFiles $true 2
        $result.Hits.Count | Should -Be 2
        $result.Truncated | Should -Be $true
        $result = searchContentIndex "株" $contentIndexFiles $true 3
        $result.Hits.Count | Should -Be 3
        $result.Truncated | Should -Be $false
        (searchContentIndex "株" $contentIndexFiles $true 0).Hits.Count | Should -Be 3
    }

    It "進み具合（照合した本文インデックスのファイルの数）を知らせ、中止できる" {
        $script:progress = @()
        $stop = @{ Stop = $false }
        $result = searchContentIndex "株" $contentIndexFiles $true -taskBytes 1 -workerCount 1 -onProgress { param($done, $total, $newHits) $script:progress += "$done/$total"; $stop.Stop = $true } -stopSignal $stop
        $result.Cancelled | Should -Be $true
        ($script:progress -join ",") | Should -Be "1/2"
    }

    It "対象ファイルで、元のファイル名が一致するものだけを検索する" {
        @((searchContentIndex "株" $contentIndexFiles $true -fileFilter "*.docx").Hits).Count | Should -Be 1
        @((searchContentIndex "株" $contentIndexFiles $true -fileFilter "!*.docx").Hits).Count | Should -Be 2
        @((searchContentIndex "株" $contentIndexFiles $true -fileFilter "*.pptx").Hits).Count | Should -Be 0
    }

    It "検索の途中で読めなくなった本文インデックスのファイル（インデックス作成中に削除された等）は飛ばす。対象が無ければ 0 件" {
        $copyRoot = Join-Path $TestDrive "deleted_content_index"
        $copies = newContentIndexFiles $tsvRoot $copyRoot
        Remove-Item -LiteralPath (fromLongPath $copies[0].Path)
        $hits = @((searchContentIndex "株" $copies $true).Hits)
        $hits.Count | Should -Be 1
        $none = searchContentIndex "株" @() $true
        $none.Total | Should -Be 0
        $none.Hits.Count | Should -Be 0
        $none.Truncated | Should -Be $false
        $none.Cancelled | Should -Be $false
    }
}

Describe "searchContentIndex（正規表現の照合の時間切れ）" -Tag Io {
    BeforeAll {
        # 入れ子の繰り返し（(a+)+$）は、一致しない文字列で照合の時間が指数的に増える
        $tsvRoot = Join-Path $TestDrive "timeout_tsv"
        for ($i = 0; $i -lt 3; $i++) {
            newTsv "$tsvRoot\s$i\book$i.xlsx\S.tsv" @(("a" * 40) + "!")
        }
        $contentIndexFiles = newContentIndexFiles $tsvRoot (Join-Path $TestDrive "timeout_content_index")
        $message = "正規表現の照合に時間がかかりすぎるため、検索を中止しました。正規表現を見直してください。"
    }

    It "時間切れなら、分かるメッセージで検索を中止する" {
        $regexTimeout = [timespan]::FromMilliseconds(1)
        { searchContentIndex "(a+)+$" $contentIndexFiles $false -workerCount 1 } | Should -Throw -ExpectedMessage "*$message*"
    }

    It "並列に検索していても、時間切れなら同じメッセージで中止する" {
        $regexTimeout = [timespan]::FromMilliseconds(1)
        { searchContentIndex "(a+)+$" $contentIndexFiles $false -workerCount 2 -taskBytes 1 } | Should -Throw -ExpectedMessage "*$message*"
    }

    It "照合のスレッドのスクリプトは、時間切れを例外にせず Timeout で返す（このスレッドで直接動かす）" {
        $regexTimeout = [timespan]::FromMilliseconds(1)
        $search = newSearchRegex "(a+)+$" $false $false
        $output = & ${contentIndexWorkerScript} $contentIndexFiles 0 $contentIndexFiles.Count $search.Regex -1 $search.TextRegex $search.ScanMode $null $null $null $null
        $output.Timeout | Should -Be $true
        $output.Hits | Should -BeNullOrEmpty
    }

    It "照合のスレッドのスクリプトは、ヒットを Hits で返す" {
        $search = newSearchRegex "a!" $true $false
        $output = & ${contentIndexWorkerScript} $contentIndexFiles 0 $contentIndexFiles.Count $search.Regex -1 $search.TextRegex $search.ScanMode $null $null $null $null
        $output.Timeout | Should -Be $false
        @($output.Hits).Count | Should -Be 3
    }
}

Describe "searchContentIndex（並列検索・読んだ内容の使い回し）" -Tag Io {
    BeforeAll {
        $tsvRoot = Join-Path $TestDrive "parallel_tsv"
        for ($i = 0; $i -lt 30; $i++) {
            newTsv ("$tsvRoot\sub{0}\book{1:D2}.xlsx\S.tsv" -f ($i % 3), $i) @("見積 $i", "x", "見積 確定 $i")
        }
        $contentIndexRoot = Join-Path $TestDrive "parallel_content_index"
        $contentIndexFiles = newContentIndexFiles $tsvRoot $contentIndexRoot

        function formatBookLines {
            param ($hits)
            return (@($hits | ForEach-Object { "$($_.Book):$($_.LineNumber)" }) -join ",")
        }
    }

    It "並列に検索しても、本文インデックスのファイルの順に同じ結果を返す" {
        $expected = formatBookLines (searchContentIndex "見積" $contentIndexFiles $true -workerCount 1).Hits
        $result = searchContentIndex "見積" $contentIndexFiles $true -workerCount 3 -taskBytes 1
        formatBookLines $result.Hits | Should -Be $expected
        $result.Hits.Count | Should -Be 60
    }

    It "並列でも上限で打ち切り、進み具合を本文インデックスのファイルの順に知らせる" {
        $script:done = New-Object System.Collections.Generic.List[int]
        $result = searchContentIndex "見積" $contentIndexFiles $true 5 -workerCount 3 -taskBytes 1 -onProgress { param($done, $total, $newHits) $script:done.Add($done) }
        $result.Truncated | Should -Be $true
        formatBookLines $result.Hits | Should -Be "book00.xlsx:1,book00.xlsx:3,book03.xlsx:1,book03.xlsx:3,book06.xlsx:1"
        ($script:done -join ",") | Should -Be ((@($script:done | Sort-Object)) -join ",")
    }

    It "並列でも中止できる" {
        $script:calls = 0
        $stop = [hashtable]::Synchronized(@{ Stop = $false })
        $result = searchContentIndex "見積" $contentIndexFiles $true -workerCount 3 -taskBytes 1 -onProgress { param($done, $total, $newHits) $script:calls++; if ($script:calls -ge 2) { $stop.Stop = $true } } -stopSignal $stop
        $result.Cancelled | Should -Be $true
        $script:calls | Should -Be 2
        $result.Hits.Count | Should -Be 40   # 2 つの作業（本文インデックスのファイル 20 件ずつ）の分だけ。止めた後の作業のヒットは入れない
    }

    It "並列で中止したあとも、同じプールで次の検索がいつもの結果になる（残った作業を受け取って捨てる）" {
        $pool = newContentIndexWorkerPool 3
        try {
            $stop = [hashtable]::Synchronized(@{ Stop = $false })
            $result = searchContentIndex "見積" $contentIndexFiles $true -taskBytes 1 -pool $pool -onProgress { param($done, $total, $newHits) $stop.Stop = $true } -stopSignal $stop
            $result.Cancelled | Should -Be $true
            $expected = formatBookLines (searchContentIndex "見積" $contentIndexFiles $true -workerCount 1).Hits
            $again = searchContentIndex "見積" $contentIndexFiles $true -taskBytes 1 -pool $pool
            formatBookLines $again.Hits | Should -Be $expected
            $again.Cancelled | Should -Be $false
        } finally {
            $pool.Close()
        }
    }

    It "並列でも、始めから Stop が立っていれば何も照合しない" {
        $result = searchContentIndex "見積" $contentIndexFiles $true -workerCount 3 -taskBytes 1 -stopSignal @{ Stop = $true }
        $result.Cancelled | Should -Be $true
        $result.Hits.Count | Should -Be 0
    }

    It "内容を使い回し、書き直された本文インデックスのファイルは読み直す" {
        $cache = newTsvTextCache
        (searchContentIndex "更新後" $contentIndexFiles $true -cache $cache).Hits.Count | Should -Be 0
        newTsv "$contentIndexRoot\sub0\book00.xlsx\S.tsv" @("更新後の内容")
        [void](updateFolderContentIndex "$contentIndexRoot\sub0")
        (Get-Item -LiteralPath "$contentIndexRoot\sub0\content_index.xlsx.001.tsv").LastWriteTime = (Get-Date).AddMinutes(1)
        (searchContentIndex "更新後" (findContentIndexFiles $contentIndexRoot) $true -cache $cache).Hits.Count | Should -Be 1
    }

    It "上限を超える内容は残さない" {
        $cache = newTsvTextCache 20
        [void](searchContentIndex "見積" $contentIndexFiles $true -cache $cache)
        $cache.Chars[0] | Should -BeLessThan 21
    }
}

Describe "getContentIndexFiles" -Tag Io {
    BeforeAll {
        $index = Join-Path $TestDrive "list[1]"
        $other = Join-Path $TestDrive "list_other"
        foreach ($item in @(
                @{ Folder = "$index"; Name = "b.xlsx" }, @{ Folder = "$index\sub"; Name = "a.xlsx" }, @{ Folder = "$other"; Name = "c.docx" })) {
            [void][System.IO.Directory]::CreateDirectory($item.Folder)
            writeContentIndexFile "$($item.Folder)\$(getContentIndexFileName (getContentIndexExtension $item.Name))" (convertToContentIndexText @(@{ Name = $item.Name; Places = @(@{ Place = "S"; Text = "x" }) }))
        }
        $missing = Join-Path $TestDrive "list_missing"
    }

    It "フォルダごとの件数と、インデックスのフォルダからの相対パスを返す" {
        $result = getContentIndexFiles @($index, $missing, $other)
        $result.Folders.Count | Should -Be 3
        $result.Folders[0].Exists | Should -Be $true
        $result.Folders[0].Count | Should -Be 2
        $result.Folders[1].Exists | Should -Be $false
        $result.ContentIndexFiles.Count | Should -Be 3
        @($result.ContentIndexFiles | Where-Object { $_.RelPath -eq "sub\content_index.xlsx.001.tsv" }).Count | Should -Be 1
    }

    It "入れ子のフォルダを指定しても同じ本文インデックスのファイルを重複させない" {
        (getContentIndexFiles @($index, "$index\sub")).ContentIndexFiles.Count | Should -Be 2
    }

    It "インデックスの中のフォルダだけ・直下だけを指定できる" {
        $result = getContentIndexFiles @(@{ Root = $index; RelPath = "sub"; Recurse = $true })
        $result.ContentIndexFiles.Count | Should -Be 1
        $result.ContentIndexFiles[0].Root | Should -Be $index
        $result.ContentIndexFiles[0].RelDir | Should -Be "sub"
        $result.Folders[0].Path | Should -Be "$index\sub"
        $direct = getContentIndexFiles @(@{ Root = $index; RelPath = ""; Recurse = $false })
        $direct.ContentIndexFiles.Count | Should -Be 1
        $direct.ContentIndexFiles[0].RelPath | Should -Be "content_index.xlsx.001.tsv"
    }

    It "存在しないフォルダは Exists が false。数えた件数を知らせる" {
        $result = getContentIndexFiles @(@{ Root = $index; RelPath = "なし"; Recurse = $true }, @{ Root = $missing; RelPath = "sub"; Recurse = $true })
        ($result.Folders | ForEach-Object { $_.Exists }) -join "," | Should -Be "False,False"
        $result.ContentIndexFiles.Count | Should -Be 0
        $counts = New-Object System.Collections.Generic.List[int]
        [void](getContentIndexFiles @($index, $other) { param ($count) $counts.Add($count) })
        ($counts -join ",") | Should -Be "2,3"
    }
}

Describe "readContentIndexLines（ヒットの前後の行を、前に読んだファイルの控えを続けて渡して読む）" -Tag Io {
    BeforeAll {
        # 画面の読み方に合わせた、テスト用の呼び出し口。lineNumber の前 before 行・後 after 行を読み、
        # 最後に読んだファイルの控えを次の読みに渡す（キャッシュから読んだときは控えを変えない）
        $script:ctxMarks = $null
        function readCtx {
            param ([string]$path, [string]$book, [string]$location, [int]$lineNumber, [int]$before = 3, [int]$after = 3, $cache = $null)
            $first = [Math]::Max(1, $lineNumber - $before)
            $last = $lineNumber + $after
            if ($last -lt $first) { return @() }
            $result = readContentIndexLines $path $book $location $first ($last - $first + 1) $cache $script:ctxMarks
            if ($null -ne $result.Marks -and $null -ne $result.Marks.Places) { $script:ctxMarks = $result.Marks }
            return $result.Rows
        }

        $folder = Join-Path $TestDrive "context"
        [void][System.IO.Directory]::CreateDirectory($folder)
        $path = "$folder\content_index.xlsx.001.tsv"
        writeContentIndexFile $path (convertToContentIndexText @(
                @{ Name = "A.xlsx"; Places = @(@{ Place = "S"; Text = "1`r`n2`r`n3`r`n4`r`n5`r`n6`r`n" }, @{ Place = "T"; Text = "t1`r`n" }) },
                @{ Name = "B.xlsx"; Places = @(@{ Place = "S"; Text = "b1`r`nb2`r`n" }) }))

        function formatContext {
            param ($rows)
            return (@($rows | ForEach-Object { "$($_.LineNumber):$($_.Line)" }) -join ",")
        }
    }

    It "元のファイル・場所の中で、その行と前後の行を返す（ほかの場所の行は入れない）" {
        formatContext (readCtx $path "A.xlsx" "S" 3 1 2) | Should -Be "2:2,3:3,4:4,5:5"
        formatContext (readCtx $path "A.xlsx" "S" 6 2 3) | Should -Be "4:4,5:5,6:6"
        formatContext (readCtx $path "B.xlsx" "S" 1 3 3) | Should -Be "1:b1,2:b2"
    }

    It "見つからない・読めないときは空。キャッシュに同じ内容があればそれを使う" {
        @(readCtx $path "無い.xlsx" "S" 1).Count | Should -Be 0
        @(readCtx "$folder\無い.tsv" "A.xlsx" "S" 1).Count | Should -Be 0
        $contentIndexFiles = findContentIndexFiles $folder
        $cache = newTsvTextCache
        [void](searchContentIndex "t1" $contentIndexFiles $true -cache $cache)
        # キャッシュの内容を書き換えると、ファイルではなくキャッシュから読んだことが分かる
        $entry = $cache.Texts[$contentIndexFiles[0].Path]
        $entry[2] = $entry[2].Replace("t1", "c1")
        formatContext (readCtx $path "A.xlsx" "T" 1 0 0 $cache) | Should -Be "1:c1"
    }

    It "キャッシュに無くても、キャッシュに在るときと同じ行を返す（1 回目はファイルの全体、2 回目からは場所の一覧を使い回して範囲だけを読む）" {
        $bigPath = "$folder\content_index.docx.001.tsv"
        $bigText = (1..30000 | ForEach-Object { "行$_" }) -join "`r`n"
        writeContentIndexFile $bigPath (convertToContentIndexText @(
                @{ Name = "前.docx"; Places = @(@{ Place = "P1"; Text = "a1`r`na2`r`n" }) },
                @{ Name = "長い.docx"; Places = @(@{ Place = "P1"; Text = "x1`r`nx2`r`n" }, @{ Place = "P2"; Text = "$bigText`r`n" }, @{ Place = "P3"; Text = "z1`r`nz2`r`n" }) }))
        $files = findContentIndexFiles $folder
        $cache = newTsvTextCache
        [void](searchContentIndex "行" $files $true -cache $cache)
        # 期待値は、検索のキャッシュを通る経路（ファイルを読まない）で作る
        $cache.Texts.ContainsKey((toLongPath $bigPath)) | Should -Be $true
        $originalReadText = Get-Command readContentIndexText -CommandType Function
        Mock readContentIndexText { & $originalReadText @args }
        $cases = @(
            @("長い.docx", "P2", 1, 3, 3), @("長い.docx", "P2", 15000, 3, 3), @("長い.docx", "P2", 30000, 3, 5),
            @("長い.docx", "P1", 2, 3, 3), @("長い.docx", "P3", 2, 1, 9), @("前.docx", "P1", 1, 0, 0), @("無い.docx", "P1", 1, 3, 3)
        )
        foreach ($round in 1..2) {
            foreach ($case in $cases) {
                $expected = formatContext (readCtx $bigPath $case[0] $case[1] $case[2] $case[3] $case[4] $cache)
                $actual = formatContext (readCtx $bigPath $case[0] $case[1] $case[2] $case[3] $case[4])
                $actual | Should -Be $expected -Because "$($case -join ',')（$round 回目）"
            }
        }
        formatContext (readCtx $bigPath "長い.docx" "P2" 29999 1 5) | Should -Be "29998:行29998,29999:行29999,30000:行30000"
        # ファイルの全体を読むのは、場所の一覧が無い 1 回目だけ（2 回目以降と、無い場所の問い合わせでは読まない）
        Should -Invoke readContentIndexText -Times 1 -Exactly -Scope It
    }

    Context "範囲だけを読む経路（場所の一覧を覚えたあと）" {
        BeforeAll {
            # 1 つの It ごとに別のファイルを作り、場所の一覧を覚えさせる（本の名前は E.csv）
            function newContextFile {
                param ([string]$tag)
                $dir = "$folder\range_$tag"
                [void][System.IO.Directory]::CreateDirectory($dir)
                $file = "$dir\content_index.csv.001.tsv"
                writeContentIndexFile $file (convertToContentIndexText @(@{ Name = "E.csv"; Places = @(@{ Place = "本文"; Text = "1`r`n2`r`n3`r`n4`r`n5`r`n6`r`n" }) }))
                [void](readCtx $file "E.csv" "本文" 3 1 1)
                return $file
            }
        }

        It "<Name>" -ForEach @(
            @{ Name = "BOM が無く範囲を読めないときは、全体を読んで同じ行を返す"; Expected = "2:2,3:3,4:4"; Scenario = "nobom" }
            @{ Name = "開けない（IOException）ときは空"; Expected = ""; Scenario = "io" }
            @{ Name = "アクセスできない（UnauthorizedAccessException）ときは空"; Expected = ""; Scenario = "denied" }
            @{ Name = "覚えたあとにファイルが無くなったら空"; Expected = ""; Scenario = "moved" }
            @{ Name = "読んでいる間にファイルが置き換わったら、古い位置の行を返さずに新しい内容を読み直す"; Expected = "2:n2,3:n3,4:n4"; Scenario = "replaced" }
        ) {
            $file = newContextFile $Scenario
            $book = "E.csv"
            $originalRange = Get-Command readContentIndexRange -CommandType Function
            switch ($Scenario) {
                "nobom" { Mock readContentIndexRange { $null } }
                "io" { Mock readContentIndexRange { throw [System.IO.IOException]::new("使用中") } }
                "denied" { Mock readContentIndexRange { throw [System.UnauthorizedAccessException]::new("拒否") } }
                "moved" { [System.IO.File]::Move($file, "$file.moved") }
                "replaced" {
                    $replacedFile = $file
                    Mock readContentIndexRange {
                        writeContentIndexFile $replacedFile (convertToContentIndexText @(@{ Name = $book; Places = @(@{ Place = "本文"; Text = "n1`r`nn2`r`nn3`r`nn4`r`nn5`r`n" }) }))
                        [System.IO.File]::SetLastWriteTimeUtc($replacedFile, [DateTime]::UtcNow.AddDays(1))
                        & $originalRange @args
                    }
                }
            }
            formatContext @(readCtx $file $book "本文" 3 1 1) | Should -Be $Expected
        }

        It "BOM の無いファイルは readContentIndexRange が `$null を返す" {
            $noBom = "$TestDrive\nobom.tsv"
            [System.IO.File]::WriteAllBytes($noBom, [System.Text.Encoding]::Unicode.GetBytes("1`r`n2`r`n"))
            readContentIndexRange $noBom 0 4 1 2 | Should -BeNullOrEmpty
        }

        It "範囲だけ読んでも、読み飛ばした行の数え方がずれない（1 行目・途中・最後・場所の外）" -ForEach @(
            @{ First = 1; Last = 2; Expected = "1:1,2:2" }
            @{ First = 3; Last = 5; Expected = "3:3,4:4,5:5" }
            @{ First = 6; Last = 9; Expected = "6:6" }
            @{ First = 7; Last = 9; Expected = "" }
        ) {
            $file = newContextFile "n$First$Last"
            $rows = readCtx $file "E.csv" "本文" $First 0 ($Last - $First)
            formatContext @($rows) | Should -Be $Expected
        }

        It "テキストの前後の行は、範囲だけ読んだときも先頭から 1,000 文字に切る" {
            $textPath = "$folder\content_index.log.002.tsv"
            $longLine = "y" * 2000
            writeContentIndexFile $textPath (convertToContentIndexText @(@{ Name = "長い2.log"; Places = @(@{ Place = "本文"; Text = "短い`r`n${longLine}`r`n短い2`r`n" }) }))
            [void](readCtx $textPath "長い2.log" "本文" 1 0 0)
            $rows = readCtx $textPath "長い2.log" "本文" 2 1 1
            $rows[0].Line | Should -Be "短い"
            $rows[1].Line.Length | Should -Be 1001
            $rows[1].Line.EndsWith("…") | Should -Be $true
            $rows[2].Line | Should -Be "短い2"
        }
    }

    It "ファイルが書き換わったら、使い回した場所の一覧を使わない" {
        $changedPath = "$folder\content_index.csv.001.tsv"
        writeContentIndexFile $changedPath (convertToContentIndexText @(@{ Name = "C.csv"; Places = @(@{ Place = "本文"; Text = "1`r`n2`r`n3`r`n" }) }))
        formatContext (readCtx $changedPath "C.csv" "本文" 2 0 0) | Should -Be "2:2"
        formatContext (readCtx $changedPath "C.csv" "本文" 2 0 0) | Should -Be "2:2"
        writeContentIndexFile $changedPath (convertToContentIndexText @(@{ Name = "D.csv"; Places = @(@{ Place = "別"; Text = "d1`r`nd2`r`nd3`r`nd4`r`n" }) }))
        @(readCtx $changedPath "C.csv" "本文" 2 0 0).Count | Should -Be 0
        formatContext (readCtx $changedPath "D.csv" "別" 4 1 0) | Should -Be "3:d3,4:d4"
    }

    It "テキストの前後の行は、一致の位置が分からないため先頭から1,000文字に切る（Excel は切らない）" {
        $textPath = "$folder\content_index.log.001.tsv"
        $longLine = "x" * 2000
        writeContentIndexFile $textPath (convertToContentIndexText @(@{ Name = "長い.log"; Places = @(@{ Place = "本文"; Text = "短い`r`n${longLine}`r`n短い2`r`n" }) }))
        $rows = readCtx $textPath "長い.log" "本文" 2 1 1
        $rows[0].Line | Should -Be "短い"
        $rows[1].Line.Length | Should -Be 1001
        $rows[1].Line.Substring(0, 1000) | Should -Be $longLine.Substring(0, 1000)
        $rows[1].Line.EndsWith("…") | Should -Be $true
        $rows[2].Line | Should -Be "短い2"
    }
}

Describe "readContentIndexLines（続きから読む。控えを引数で渡し、戻り値で受ける）" -Tag Io {
    BeforeAll {
        $folder = Join-Path $TestDrive "lines"
        [void][System.IO.Directory]::CreateDirectory($folder)
        $path = "$folder\content_index.docx.001.tsv"
        $bigText = (1..1000 | ForEach-Object { "行$_" }) -join "`r`n"
        writeContentIndexFile $path (convertToContentIndexText @(
                @{ Name = "前.docx"; Places = @(@{ Place = "P1"; Text = "a1`r`na2`r`n" }) },
                @{ Name = "長い.docx"; Places = @(@{ Place = "P1"; Text = "$bigText`r`n" }, @{ Place = "P2"; Text = "z1`r`nz2`r`n" }) }))

        function formatLines {
            param ($result)
            return (@($result.Rows | ForEach-Object { "$($_.LineNumber):$($_.Line)" }) -join ",")
        }
    }

    It "1 回目はファイルの全体、続きは控えを使って範囲だけを読む。どちらも同じ行を返し、終わりを知らせる" {
        $originalReadText = Get-Command readContentIndexText -CommandType Function
        Mock readContentIndexText { & $originalReadText @args }
        $first = readContentIndexLines $path "長い.docx" "P1" 1 3
        formatLines $first | Should -Be "1:行1,2:行2,3:行3"
        $first.End | Should -Be $false
        $next = readContentIndexLines $path "長い.docx" "P1" 4 3 $null $first.Marks
        formatLines $next | Should -Be "4:行4,5:行5,6:行6"
        $far = readContentIndexLines $path "長い.docx" "P1" 998 5 $null $next.Marks
        formatLines $far | Should -Be "998:行998,999:行999,1000:行1000"
        $far.End | Should -Be $true
        $up = readContentIndexLines $path "長い.docx" "P1" 1 2 $null $far.Marks
        formatLines $up | Should -Be "1:行1,2:行2"
        # ファイルの全体を読むのは 1 回目だけ
        Should -Invoke readContentIndexText -Times 1 -Exactly -Scope It
    }

    It "渡した控えは書き換えない。200 行ごとの位置を控え、遠い行はその手前から読む" {
        $first = readContentIndexLines $path "長い.docx" "P1" 1 1000
        $key = "長い.docx`tP1"
        $keys = @($first.Marks.Checkpoints[$key].Keys | Sort-Object)
        ($keys -join ",") | Should -Be "201,401,601,801"
        $before = $first.Marks.Checkpoints[$key].Count
        $result = readContentIndexLines $path "長い.docx" "P1" 650 2 $null $first.Marks
        formatLines $result | Should -Be "650:行650,651:行651"
        $first.Marks.Checkpoints[$key].Count | Should -Be $before
        # 範囲を読むと、通った 200 行ごとの位置が増える（読み進めた分だけ）
        $fresh = readContentIndexLines $path "長い.docx" "P1" 1 1000
        $fresh.Marks.Checkpoints[$key].Remove(801)
        $fresh.Marks.Checkpoints[$key].Remove(601)
        $fresh.Marks.Checkpoints[$key].Remove(401)
        $more = readContentIndexLines $path "長い.docx" "P1" 790 20 $null $fresh.Marks
        formatLines $more | Should -Be ((790..809 | ForEach-Object { "${_}:行$_" }) -join ",")
        $more.Marks.Checkpoints[$key].ContainsKey(801) | Should -Be $true
    }

    It "同じ控えを別の場所の読みに使っても、別の場所の位置から読まない" {
        # 同じファイルの P1（1000 行）で 200 行ごとの位置を控えたあと、P2（2 行）を同じ控えで読む
        $first = readContentIndexLines $path "長い.docx" "P1" 1 1000
        $other = readContentIndexLines $path "長い.docx" "P2" 1 5 $null $first.Marks
        formatLines $other | Should -Be "1:z1,2:z2"
        $other.End | Should -Be $true
        # P1 の位置は P2 の読みで壊れず、P1 に戻っても同じ行を返す
        formatLines (readContentIndexLines $path "長い.docx" "P1" 650 2 $null $other.Marks) | Should -Be "650:行650,651:行651"
        # 控えが無い別の場所から先に読んでも、あとの場所の読みが先の場所の位置を使わない
        $p2 = readContentIndexLines $path "長い.docx" "P2" 1 2
        formatLines (readContentIndexLines $path "長い.docx" "P1" 401 2 $null $p2.Marks) | Should -Be "401:行401,402:行402"
    }

    It "場所の外・見つからない・読めないときは空で End。頼んだ数が 0 なら空で End にしない" {
        (readContentIndexLines $path "長い.docx" "P2" 3 5).End | Should -Be $true
        @((readContentIndexLines $path "長い.docx" "P2" 3 5).Rows).Count | Should -Be 0
        @((readContentIndexLines $path "無い.docx" "P1" 1 5).Rows).Count | Should -Be 0
        $missing = readContentIndexLines "$folder\無い.tsv" "長い.docx" "P1" 1 5
        @($missing.Rows).Count | Should -Be 0
        $missing.End | Should -Be $true
        $none = readContentIndexLines $path "長い.docx" "P1" 1 0
        $none.End | Should -Be $false
        @($none.Rows).Count | Should -Be 0
    }

    It "ファイルが書き換わったら、古い控えを使わない" {
        $file = "$folder\content_index.csv.001.tsv"
        writeContentIndexFile $file (convertToContentIndexText @(@{ Name = "C.csv"; Places = @(@{ Place = "本文"; Text = "1`r`n2`r`n3`r`n" }) }))
        $old = readContentIndexLines $file "C.csv" "本文" 1 2
        formatLines $old | Should -Be "1:1,2:2"
        writeContentIndexFile $file (convertToContentIndexText @(@{ Name = "C.csv"; Places = @(@{ Place = "本文"; Text = "n1`r`nn2`r`nn3`r`nn4`r`n" }) }))
        [System.IO.File]::SetLastWriteTimeUtc($file, [DateTime]::UtcNow.AddDays(1))
        formatLines (readContentIndexLines $file "C.csv" "本文" 2 3 $null $old.Marks) | Should -Be "2:n2,3:n3,4:n4"
    }

    # 控えが古い（ファイルが変わった）ことは、パス・更新日時・大きさで見分ける。更新日時も大きさも同じまま中身だけが変わると見分けられない（この表には入れない）
    It "<Name>、古い控えの行は返さず、今のファイルの行を返す" -ForEach @(
        @{ Name = "大きさが変わり、更新日時は同じ"; Scenario = "size" }
        @{ Name = "更新日時が変わり、大きさは同じ"; Scenario = "ticks" }
        @{ Name = "インデックスを作り直して場所の並びが変わった"; Scenario = "rebuilt" }
        @{ Name = "別のファイルの控え（本の名前・場所の名前が同じ）"; Scenario = "otherfile" }
    ) {
        $dir = "$folder\stale_$Scenario"
        [void][System.IO.Directory]::CreateDirectory($dir)
        $file = "$dir\content_index.csv.001.tsv"
        $oldText = (1..300 | ForEach-Object { "旧$_" }) -join "`r`n"
        writeContentIndexFile $file (convertToContentIndexText @(@{ Name = "C.csv"; Places = @(@{ Place = "本文"; Text = "$oldText`r`n" }) }))
        $old = readContentIndexLines $file "C.csv" "本文" 1 250
        $oldTicks = (Get-Item -LiteralPath $file).LastWriteTimeUtc.Ticks
        $target = $file
        switch ($Scenario) {
            "size" {
                $newText = (1..300 | ForEach-Object { "新$_" }) -join "`r`n"
                writeContentIndexFile $file (convertToContentIndexText @(@{ Name = "C.csv"; Places = @(@{ Place = "本文"; Text = "$newText`r`n新301`r`n" }) }))
                [System.IO.File]::SetLastWriteTimeUtc($file, [DateTime]::new($oldTicks, [DateTimeKind]::Utc))
            }
            "ticks" {
                $newText = (1..300 | ForEach-Object { "新$_" }) -join "`r`n"
                writeContentIndexFile $file (convertToContentIndexText @(@{ Name = "C.csv"; Places = @(@{ Place = "本文"; Text = "$newText`r`n" }) }))
                [System.IO.File]::SetLastWriteTimeUtc($file, [DateTime]::UtcNow.AddDays(2))
            }
            "rebuilt" {
                $newText = (1..300 | ForEach-Object { "新$_" }) -join "`r`n"
                writeContentIndexFile $file (convertToContentIndexText @(@{ Name = "X.csv"; Places = @(@{ Place = "別"; Text = "x`r`n" }) }, @{ Name = "C.csv"; Places = @(@{ Place = "本文"; Text = "$newText`r`n" }) }))
                [System.IO.File]::SetLastWriteTimeUtc($file, [DateTime]::UtcNow.AddDays(3))
            }
            "otherfile" {
                $target = "$dir\content_index.csv.002.tsv"
                $newText = (1..300 | ForEach-Object { "別$_" }) -join "`r`n"
                writeContentIndexFile $target (convertToContentIndexText @(@{ Name = "C.csv"; Places = @(@{ Place = "本文"; Text = "$newText`r`n" }) }))
            }
        }
        $prefix = if ($Scenario -eq "otherfile") { "別" } else { "新" }
        $result = readContentIndexLines $target "C.csv" "本文" 250 3 $null $old.Marks
        formatLines $result | Should -Be "250:${prefix}250,251:${prefix}251,252:${prefix}252"
        # 返した控えは今のファイルのもの（次の読みでも、古い行は返らない）
        $result.Marks.Path | Should -Be (toLongPath $target)
        formatLines (readContentIndexLines $target "C.csv" "本文" 1 2 $null $result.Marks) | Should -Be "1:${prefix}1,2:${prefix}2"
    }

    It "キャッシュに同じ内容があれば、ファイルを読まずにキャッシュから返す" {
        $files = findContentIndexFiles $folder
        $cache = newTsvTextCache
        [void](searchContentIndex "行" $files $true -cache $cache)
        $entry = $cache.Texts[(toLongPath $path)]
        $entry[2] = $entry[2].Replace("行5`n", "改5`n")
        formatLines (readContentIndexLines $path "長い.docx" "P1" 4 3 $cache) | Should -Be "4:行4,5:改5,6:行6"
    }
}

Describe "testCellPrefixedHit（セル番地 + 区切りのタブ + 文字の行で、最初のタブより後の一致だけを数える）" -Tag Unit {
    It "<Name>" -ForEach @(
        @{ Name = "タブが無い行は一致なし"; Pattern = "C2"; Line = "C2 納期"; Expected = $false }
        @{ Name = "文字が空（タブの後ろに何も無い）なら、番地に一致する語は一致なし"; Pattern = "C2"; Line = "C2`t"; Expected = $false }
        @{ Name = "文字が空でも、タブに当たる語は区切りのタブなので一致なし"; Pattern = "`t"; Line = "C2`t"; Expected = $false }
        @{ Name = "番地だけに一致する語は一致なし"; Pattern = "C2"; Line = "C2`t納期"; Expected = $false }
        @{ Name = "文字に一致する語は一致"; Pattern = "納期"; Line = "C2`t納期"; Expected = $true }
        @{ Name = "文字の中にも同じ語があれば一致"; Pattern = "C2"; Line = "C2`tC2 の部品"; Expected = $true }
        @{ Name = "行頭の ^ は、文字の先頭に一致しない（既知の制限）"; Pattern = "^C2"; Line = "C2`tC2 の部品"; Expected = $false }
        @{ Name = "区切りのタブを見る後読みは、文字の先頭に一致する"; Pattern = "(?<=	)納期"; Line = "C2`t納期"; Expected = $true }
        @{ Name = "文字の中のタブも検索の対象"; Pattern = "品名`t納期"; Line = "C3`t`"品名`t納期`""; Expected = $true }
        @{ Name = "文字の中のタブだけに当たる語も一致"; Pattern = "`t"; Line = "C3`t`"品名`t納期`""; Expected = $true }
    ) {
        testCellPrefixedHit ([regex]::new($Pattern)) $Line | Should -Be $Expected
    }
}

Describe "壊れた・大きすぎる本文インデックスのファイルは、そのファイルだけ飛ばす" -Tag Io {
    BeforeAll {
        function script:newMixedContentIndexFiles {
            # 正しい本文インデックスのファイル（営業）に、中身 text の本文インデックスのファイル（壊れ）を混ぜた並びを返す
            param ([string]$name, [string]$text)
            $tsvRoot = Join-Path $TestDrive "mixed_tsv_$name"
            newTsv "$tsvRoot\営業\見積.xlsx\Sheet1.tsv" @("単価`t105")
            $root = Join-Path $TestDrive "mixed_content_index_$name"
            [void](newContentIndexFiles $tsvRoot $root)
            [void][System.IO.Directory]::CreateDirectory("$root\壊れ")
            writeContentIndexFile "$root\壊れ\content_index.xlsx.001.tsv" $text
            return , (getContentIndexFiles @($root)).ContentIndexFiles
        }
    }

    It "<Name> のファイルが 1 つ混ざっても、ほかのファイルのヒットを返し、そのファイルだけを Skipped に入れる（直列・並列とも）" -TestCases @(
        @{ Name = "版が違う"; Text = "{RS} 版=999{LF}{RS} ファイル名=a.xlsx{LF}{RS} シート=S{LF}単価{LF}" }
        @{ Name = "版が無い"; Text = "{RS} ファイル名=a.xlsx{LF}{RS} シート=S{LF}単価{LF}" }
        @{ Name = "ファイル名の行だけ"; Text = "{RS} ファイル名=単価.xlsx" }
        @{ Name = "版の値が数でない"; Text = "{RS} 版=あ{LF}{RS} ファイル名=a.xlsx{LF}単価{LF}" }
    ) {
        $text = $Text.Replace("{RS}", [string][char]0x1E).Replace("{LF}", "`n")
        $files = newMixedContentIndexFiles $Name $text
        $files.Count | Should -Be 2
        $pool = newContentIndexWorkerPool 2
        try {
            foreach ($usePool in @($false, $true)) {
                $result = if ($usePool) { searchContentIndex "単価" $files $true -pool $pool -taskBytes 1 } else { searchContentIndex "単価" $files $true }
                @($result.Hits).Count | Should -Be 1
                $result.Hits[0].Book | Should -Be "見積.xlsx"
                @($result.Skipped).Count | Should -Be 1
                $result.Skipped[0] | Should -BeLike "*\壊れ\content_index.xlsx.001.tsv"
            }
        } finally {
            $pool.Close()
        }
    }

    It "大きさが上限を超える本文インデックスのファイルは、読まずに飛ばす（直列・並列とも）" {
        $files = newMixedContentIndexFiles "large" "$([char]0x1E) 版=2`n"
        $large = @($files | ForEach-Object { $copy = $_.Clone(); if ($copy.RelPath -like "壊れ*") { $copy.Size = ${contentIndexSearchMaxBytes} + 1 }; $copy })
        $pool = newContentIndexWorkerPool 2
        try {
            foreach ($usePool in @($false, $true)) {
                $result = if ($usePool) { searchContentIndex "単価" $large $true -pool $pool -taskBytes 1 } else { searchContentIndex "単価" $large $true }
                @($result.Hits).Count | Should -Be 1
                @($result.Skipped).Count | Should -Be 1
            }
        } finally {
            $pool.Close()
        }
    }

    It "場所の番号が数でなくても例外にならず、飛ばさない（場所の名前はその値のまま）" {
        $text = "{RS} 版=2{LF}{RS} ファイル名=a.pdf{LF}{RS} ページ=あ{LF}単価{LF}".Replace("{RS}", [string][char]0x1E).Replace("{LF}", "`n")
        $files = newMixedContentIndexFiles "page" $text
        $result = searchContentIndex "単価" $files $true
        @($result.Skipped).Count | Should -Be 0
        @($result.Hits | Where-Object { $_.Book -eq "a.pdf" }).Count | Should -Be 1
        @($result.Hits | Where-Object { $_.Book -eq "a.pdf" })[0].Location | Should -Be "あ"
    }

    It "照合の時間切れは飛ばさず、そのまま伝える（検索の全体を中止する）" {
        $slowRoot = Join-Path $TestDrive "skip_slow"
        newTsv "$slowRoot\idx\遅い.xlsx\Sheet1.tsv" @(("a" * 40) + "!")
        convertFolderToContentIndex "$slowRoot\idx" "$slowRoot\content_index\idx" | Out-Null
        $slowFiles = findContentIndexFiles "$slowRoot\content_index"
        $lineRegex = [regex]::new("^(a|aa)+b", [System.Text.RegularExpressions.RegexOptions]::None, [timespan]::FromMilliseconds(1))
        $thrown = $null
        try { [void](searchContentIndexFiles $slowFiles 0 1 $lineRegex -1 $null) } catch { $thrown = $_.Exception }
        $thrown | Should -Not -BeNullOrEmpty
        ($thrown -is [System.Text.RegularExpressions.RegexMatchTimeoutException] -or $thrown.InnerException -is [System.Text.RegularExpressions.RegexMatchTimeoutException]) | Should -Be $true
    }
}
