# 検索用の集約ファイルの検索（tebunko\search\pack_search.ps1）と読み書き（index\pack_store.ps1）のテスト。
# 元の TSV を 1 行ずつ読んで照合した結果（referenceKeys）と、集約ファイルの検索の結果が同じになることを確かめる
. "$PSScriptRoot\..\..\helpers\load.ps1"

function script:toKeys {
    # ヒットを比べられる文字列にする（RelPath・FileName は集約ファイルのもののため比べない）
    param ($hits)
    return @($hits | ForEach-Object { "{0}|{1}|{2}|{3}|{4}" -f $_.RelDir, $_.Book, $_.Location, $_.LineNumber, $_.Line })
}

function script:sortedKeys {
    param ($hits)
    return ((toKeys $hits) | Sort-Object) -join "`n"
}

function script:referenceKeys {
    # 元の TSV（<相対フォルダ>\<ファイル名.xlsx>\<場所>.tsv）を StreamReader.ReadLine で 1 行ずつ読み、
    # 検索語の正規表現で照合した結果（集約ファイルの検索が同じ結果になるべきもの）
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
        $place = decodeIndexPlace ([System.IO.Path]::GetFileNameWithoutExtension($path))
        if ($filter.Include -and !$filter.Include.IsMatch($book)) { continue }
        if ($filter.Exclude -and $filter.Exclude.IsMatch($book)) { continue }
        if ($exclude -and $exclude.IsMatch($place)) { continue }
        $relDir = [System.IO.Path]::GetDirectoryName($bookDir).Substring($tsvRoot.Length).Trim("\")
        $reader = [System.IO.StreamReader]::new($path, [System.Text.Encoding]::UTF8, $true)
        try {
            $number = 0
            while ($null -ne ($line = $reader.ReadLine())) {
                $number++
                if ($search.Regex.IsMatch($line)) { $keys.Add("$relDir|$book|$place|$number|$line") }
            }
        } finally {
            $reader.Dispose()
        }
    }
    return ($keys | Sort-Object) -join "`n"
}

Describe "集約ファイルの作成と検索" -Tag Io {
    $tsvRoot = Join-Path $TestDrive "tsv"
    $packRoot = Join-Path $TestDrive "pack"
    $idx = "$tsvRoot\営業"
    newTsv "$idx\見積.xlsx\$(toIndexFileName "見積")" @("品名`t数量`t単価", "", "りんご`t10`t100", "ABC`tabc")
    newTsv "$idx\見積.xlsx\$(toIndexFileName "見積[図形]")" @("E2`t承認済み 単価")
    newTsv "$idx\見積.xlsx\$(toIndexFileName "50%引き")" @("単価は税抜")
    newTsv "$idx\議事録.docx\$(toIndexFileName "ページ001")" @("見積の方針", "単価は据え置き")
    newTsv "$idx\議事録.docx\$(toIndexFileName "ページ001[コメント]")" @("単価の確認")
    newTsv "$idx\2025\提案.pptx\$(toIndexFileName "スライド001_ノート")" @("単価の背景")
    newTsv "$idx\2025\提案.pptx\$(toIndexFileName "スライド002（非表示）")" @("予備", "単価")
    $packs = newPackIndex $tsvRoot $packRoot

    It "フォルダごと・拡張子ごとに集約ファイルを作り、フォルダの順・名前の順に並べる" {
        $packs.Count | Should Be 3
        $packs[0].RelPath | Should Be "営業\content.docx.001.tsv"
        $packs[1].RelPath | Should Be "営業\content.xlsx.001.tsv"
        $packs[2].RelPath | Should Be "営業\2025\content.pptx.001.tsv"
    }

    It "元のファイルが無くなった拡張子の集約ファイルは、変換し直すときに消す" {
        $dest = Join-Path $TestDrive "reconvert"
        [void][System.IO.Directory]::CreateDirectory($dest)
        writePackFile "$dest\content.pptx.001.tsv" (convertToPackText @(@{ Name = "古い.pptx"; Places = @(@{ Place = "スライド001"; Text = "古い" }) }))
        $result = convertIndexFolderToPack $idx $dest @("古い.pptx")
        $result.Files | Should Be 2
        @([System.IO.Directory]::GetFiles($dest) | ForEach-Object { [System.IO.Path]::GetFileName($_) } | Sort-Object) -join "," | Should Be "content.docx.001.tsv,content.xlsx.001.tsv"
    }

    It "集約ファイルは UTF-16LE（BOM 付き）で、一時ファイルを残さない" {
        $bytes = [System.IO.File]::ReadAllBytes($packs[0].Path)
        "{0:X2}{1:X2}" -f $bytes[0], $bytes[1] | Should Be "FFFE"
        @(Get-ChildItem $packRoot -Recurse -Filter "*.tmp").Count | Should Be 0
    }

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
    foreach ($case in $cases) {
        It "「$($case.Word)」の結果が、TSV を 1 行ずつ照合したときと同じ" {
            sortedKeys (searchPackIndex $case.Word $packs $case.Simple).Hits | Should BeExactly (referenceKeys $tsvRoot $case.Word $case.Simple)
        }
    }

    It "大文字・小文字の区別、図形・コメントの除外、対象ファイルの条件が、TSV を 1 行ずつ照合したときと同じ" {
        sortedKeys (searchPackIndex "ABC" $packs $true -caseSensitive $true).Hits | Should BeExactly (referenceKeys $tsvRoot "ABC" $true $true)
        sortedKeys (searchPackIndex "単価" $packs $true -includeShapes $false -includeComments $false).Hits |
            Should BeExactly (referenceKeys $tsvRoot "単価" $true $false "" $false $false)
        sortedKeys (searchPackIndex "単価" $packs $true -fileFilter "*.docx").Hits | Should BeExactly (referenceKeys $tsvRoot "単価" $true $false "*.docx")
    }

    It "上限で打ち切る" {
        $result = searchPackIndex "単価" $packs $true 2
        $result.Hits.Count | Should Be 2
        $result.Truncated | Should Be $true
    }

    It "キャッシュを使っても結果が同じ" {
        $cache = newTsvTextCache
        $first = toKeys (searchPackIndex "単価" $packs $true -cache $cache).Hits
        $cache.Texts.Count | Should Be 3
        $second = toKeys (searchPackIndex "単価" $packs $true -cache $cache).Hits
        $second -join "`n" | Should BeExactly ($first -join "`n")
    }

    It "並列でも 1 スレッドと同じ結果を同じ順で返す" {
        $single = toKeys (searchPackIndex "単価" $packs $true -workerCount 1).Hits
        $tasks = splitPackTasks $packs 1
        $tasks.Count | Should Be 3
        # 1 バイトごとに分けて、2 つのスレッドで探す
        $parallel = toKeys (searchPackIndex "単価" $packs $true -workerCount 2 -taskBytes 1).Hits
        $parallel -join "`n" | Should BeExactly ($single -join "`n")
    }

    It "フォルダの一部・直下だけ・無いフォルダを列挙できる。元のファイルが無いフォルダは作らない" {
        (getPackFiles $packRoot "営業" $false).Count | Should Be 2
        (getPackFiles $packRoot "営業\2025" $true).Count | Should Be 1
        (getPackFiles $packRoot "無いフォルダ").Count | Should Be 0
        $empty = Join-Path $TestDrive "empty"
        [void][System.IO.Directory]::CreateDirectory($empty)
        (convertIndexFolderToPack $empty (Join-Path $TestDrive "empty_pack")).Books | Should Be 0
        [System.IO.Directory]::Exists((Join-Path $TestDrive "empty_pack")) | Should Be $false
    }

    It "更新日時が変わった集約ファイルは読み直し、キャッシュの古い内容を置き換える" {
        $cache = newTsvTextCache
        [void](searchPackIndex "単価" $packs $true -cache $cache)
        $chars = $cache.Chars[0]
        $changed = @($packs | ForEach-Object { $copy = $_.Clone(); $copy.Ticks = $_.Ticks + 1; $copy })
        (toKeys (searchPackIndex "単価" $changed $true -cache $cache).Hits).Count | Should Be 7
        $cache.Texts.Count | Should Be 3
        $cache.Chars[0] | Should Be $chars
        $cache.Texts[$packs[0].Path][0] | Should Be ($packs[0].Ticks + 1)
    }

    It "中止を求められたら止める" {
        $result = searchPackIndex "単価" $packs $true -shouldStop { $true }
        $result.Cancelled | Should Be $true
        $result.Hits.Count | Should Be 0
    }

    It "並列でも上限で打ち切り、残りの検索を止める" {
        $result = searchPackIndex "単価" $packs $true 1 -workerCount 2 -taskBytes 1
        $result.Hits.Count | Should Be 1
        $result.Truncated | Should Be $true
    }

    It "1 行ずつ照合する検索語でも上限で打ち切る" {
        $result = searchPackIndex "(?!予備)単価" $packs $false 2
        $result.Hits.Count | Should Be 2
        $result.Truncated | Should Be $true
    }

    It "照合のしかたが無い（全文の正規表現を渡さない）ときは 1 行ずつ照合する" {
        $hits = searchPackFiles $packs 0 $packs.Count ([regex]"単価") -1 $null "lines"
        sortedKeys $hits | Should BeExactly (sortedKeys (searchPackIndex "単価" $packs $true).Hits)
    }

    It "全文への照合が時間切れになったファイルは、1 行ずつ照合し直す" {
        # 1 行は短いため 1 行ずつなら速いが、全文には時間がかかる正規表現（全文の方だけ時間切れを短くする）
        $slowRoot = Join-Path $TestDrive "slow"
        newTsv "$slowRoot\idx\遅い.xlsx\$(toIndexFileName "S")" (@(1..3000 | ForEach-Object { "aaaaaaaaaaaaaaaaaaaaaa!" }) + @("aaab"))
        convertIndexFolderToPack "$slowRoot\idx" "$slowRoot\pack\idx" | Out-Null
        $slowPacks = getPackFiles "$slowRoot\pack"
        $pattern = "^(a|aa)+b"
        $options = [System.Text.RegularExpressions.RegexOptions]::None
        $lineRegex = [regex]::new($pattern, $options, [timespan]::FromSeconds(5))
        $textRegex = [regex]::new($pattern, $options -bor [System.Text.RegularExpressions.RegexOptions]::Multiline, [timespan]::FromMilliseconds(1))
        $hits = searchPackFiles $slowPacks 0 1 $lineRegex -1 $textRegex "lines"
        $hits.Count | Should Be 1
        $hits[0].LineNumber | Should Be 3001
    }

    It "置かれた TSV を集約ファイルに入れて TSV を消し、変わらない元のファイルは前の集約ファイルから写す" {
        $dir = Join-Path $TestDrive "inplace\営業"
        newTsv "$dir\A.xlsx\$(toIndexFileName "S")" @("A の単価")
        newTsv "$dir\B.xlsx\$(toIndexFileName "S")" @("B の古い単価")
        (updateIndexFolderPack $dir).Books | Should Be 2
        @([System.IO.Directory]::GetDirectories($dir)).Count | Should Be 0
        # B を更新し、C（Word）を足す。A の TSV はもう無い
        newTsv "$dir\B.xlsx\$(toIndexFileName "S")" @("B の新しい単価")
        newTsv "$dir\C.docx\$(toIndexFileName "ページ001")" @("C の単価")
        $result = updateIndexFolderPack $dir
        $result.Books | Should Be 3
        $result.Tsv | Should Be 2
        @([System.IO.Directory]::GetDirectories($dir)).Count | Should Be 0
        $inPacks = getPackFiles (Join-Path $TestDrive "inplace")
        @($inPacks | ForEach-Object { [System.IO.Path]::GetFileName($_.RelPath) }) -join "," | Should Be "content.docx.001.tsv,content.xlsx.001.tsv"
        (toKeys (searchPackIndex "単価" $inPacks $true).Hits) -join "`n" | Should BeExactly ((
            "営業|C.docx|ページ001|1|C の単価", "営業|A.xlsx|S|1|A の単価", "営業|B.xlsx|S|1|B の新しい単価") -join "`n")
        # C が無くなったら外し、Word の集約ファイルを消す
        (updateIndexFolderPack $dir @("C.docx")).Books | Should Be 2
        @([System.IO.Directory]::GetFiles($dir) | ForEach-Object { [System.IO.Path]::GetFileName($_) }) -join "," | Should Be "content.xlsx.001.tsv"
    }

    It "TSV の残ったフォルダを見つけ、集約ファイルとシステムインデックスに書き出して TSV を消す" {
        $work = Join-Path $TestDrive "publish"
        $index = "$work\index"
        newTsv "$index\人事\A.xlsx\$(toIndexFileName "S")" @("採用の計画")
        newTsv "$index\人事\2025\B.docx\$(toIndexFileName "ページ001")" @("評価の方針")
        # 集約ファイルは、元のファイルごとのフォルダではない
        writePackFile "$index\人事\2025\content.xlsx.001.tsv" (convertToPackText @(@{ Name = "C.xlsx"; Places = @(@{ Place = "S"; Text = "既に入っている" }) }))
        $found = findIndexFoldersWithBooks $index
        @($found | ForEach-Object { $_.Substring($index.Length) }) -join "," | Should Be "\人事,\人事\2025"
        (findIndexFoldersWithBooks "$work\無い").Count | Should Be 0

        $pending = New-Object 'System.Collections.Generic.Dictionary[string,object]'
        foreach ($folder in $found) { $pending[$folder] = @() }
        $state = "$work\システムインデックスの状態.tsv"
        publishIndexFolders $pending $index "$work\system_index" $state | Should Be 2
        (findIndexFoldersWithBooks $index).Count | Should Be 0
        [System.IO.File]::Exists("$index\人事\content.xlsx.001.tsv") | Should Be $true
        [System.IO.File]::Exists("$index\人事\2025\content.docx.001.tsv") | Should Be $true
        $txt = "$work\system_index\人事\2025\${systemIndexFileName}"
        [System.IO.File]::Exists($txt) | Should Be $true
        (readSystemIndexState $state).Pending["人事\2025\${systemIndexFileName}"] | Should Be ([System.IO.File]::GetLastWriteTimeUtc($txt).Ticks)
    }

    It "名前が .xlsx などで終わる本物のフォルダ・中身が空のファイルのフォルダは、集約する前の TSV と取り違えない" {
        $index = Join-Path $TestDrive "bookdir\index"
        $folder = "$index\営業"
        # 元のフォルダに「資料.xlsx」という名前のフォルダがあり、その中の集約ファイルがある
        [void][System.IO.Directory]::CreateDirectory("$folder\資料.xlsx")
        writePackFile "$folder\資料.xlsx\content.docx.001.tsv" (convertToPackText @(@{ Name = "中の文書.docx"; Places = @(@{ Place = "ページ001"; Text = "中の文書" }) }))
        # 取り込んだが中身が空のファイル（フォルダだけ残る）と、集約する前の TSV
        [void][System.IO.Directory]::CreateDirectory("$folder\空.xlsx")
        newTsv "$folder\B.xlsx\$(toIndexFileName "S")" @("B の中身")

        testIndexBookDir "$folder\資料.xlsx" | Should Be $false
        testIndexBookDir "$folder\空.xlsx" | Should Be $false
        testIndexBookDir "$folder\空.xlsx" $false | Should Be $true
        testIndexBookDir "$folder\B.xlsx" | Should Be $true
        testIndexBookDir "$folder\無い.xlsx" | Should Be $false
        @((findIndexFoldersWithBooks $index) | ForEach-Object { $_.Substring($index.Length) }) -join "," | Should Be "\営業"

        $result = updateIndexFolderPack $folder
        $result.Books | Should Be 1
        [System.IO.File]::Exists("$folder\資料.xlsx\content.docx.001.tsv") | Should Be $true
        [System.IO.Directory]::Exists("$folder\空.xlsx") | Should Be $true
        [System.IO.Directory]::Exists("$folder\B.xlsx") | Should Be $false
        (findIndexFoldersWithBooks $index).Count | Should Be 0
        # 本物のフォルダは、システムインデックスでも自分の txt を持つ（親の txt に入れない）
        (getSystemIndexFolderTsvPaths $folder | ForEach-Object { [System.IO.Path]::GetFileName($_) }) -join "," | Should Be "content.xlsx.001.tsv"
    }

    It "大きさの上限を超えたら次の番号の集約ファイルに分け、変わった集約ファイルだけを書き直す。検索の結果は変わらない" {
        $dir = Join-Path $TestDrive "split\営業"
        foreach ($i in 1..5) { newTsv ("$dir\資料{0}.xlsx\S.tsv" -f $i) @(("行 $i " + ("あ" * 600)), "単価 $i") }
        # 1 冊は約 1.3KB。上限 2KB なら 2 冊ずつの集約ファイルに分かれる
        $result = convertIndexFolderToPack $dir $dir @() $true 2048
        $result.Files | Should Be 3
        $result.Texts.Count | Should Be 3
        @([System.IO.Directory]::GetFiles($dir) | ForEach-Object { [System.IO.Path]::GetFileName($_) } | Sort-Object) -join "," | Should Be "content.xlsx.001.tsv,content.xlsx.002.tsv,content.xlsx.003.tsv"
        $splitPacks = getPackFiles (Join-Path $TestDrive "split")
        (toKeys (searchPackIndex "単価" $splitPacks $true).Hits | ForEach-Object { ($_ -split "\|")[1] }) -join "," | Should Be "資料1.xlsx,資料2.xlsx,資料3.xlsx,資料4.xlsx,資料5.xlsx"

        # 資料1 を更新し、資料6 を足す。資料1 の集約ファイル（001）と最後の集約ファイル（003）だけを書き直す
        $before = @{}
        foreach ($p in $splitPacks) { $before[[System.IO.Path]::GetFileName($p.RelPath)] = $p.Ticks }
        Start-Sleep -Milliseconds 20
        newTsv "$dir\資料1.xlsx\S.tsv" @("更新した単価")
        newTsv "$dir\資料6.xlsx\S.tsv" @("単価 6")
        $result = convertIndexFolderToPack $dir $dir @() $true 2048
        $result.Written | Should Be 2
        $result.Texts.Count | Should Be 3
        $after = @{}
        foreach ($p in (getPackFiles (Join-Path $TestDrive "split"))) { $after[[System.IO.Path]::GetFileName($p.RelPath)] = $p.Ticks }
        ($after["content.xlsx.002.tsv"] -eq $before["content.xlsx.002.tsv"]) | Should Be $true
        ($after["content.xlsx.001.tsv"] -ne $before["content.xlsx.001.tsv"]) | Should Be $true
        (searchPackIndex "単価" (getPackFiles (Join-Path $TestDrive "split")) $true).Hits.Count | Should Be 6

        # 2 冊とも無くなった集約ファイル（002）は消す
        $result = convertIndexFolderToPack $dir $dir @("資料3.xlsx", "資料4.xlsx") $true 2048
        [System.IO.File]::Exists("$dir\content.xlsx.002.tsv") | Should Be $false
        $result.Files | Should Be 2
    }

    It "書き直すと中身が置き換わる" {
        $path = Join-Path $TestDrive "rewrite\$(getPackFileName "xlsx")"
        [void][System.IO.Directory]::CreateDirectory((Split-Path $path))
        writePackFile $path "a"
        writePackFile $path "b"
        readPackText $path | Should BeExactly "b"
        [System.IO.File]::Exists("$path.tmp") | Should Be $false
    }
}

Describe "searchPackIndex（改行の種類・照合のしかたによらず、TSV を 1 行ずつ照合したときと同じ）" -Tag Io {
    # 改行の種類（CRLF・LF・CR）、末尾の改行の有無、空行、空のファイル、BOM の無いファイルを混ぜる
    $tsvRoot = Join-Path $TestDrive "modes_tsv"
    $dir = "$tsvRoot\idx"
    foreach ($book in "a.xlsx", "b.docx", "c.pptx", "d.xlsx") { [void][System.IO.Directory]::CreateDirectory("$dir\$book") }
    [System.IO.File]::WriteAllText("$dir\a.xlsx\S.tsv", "見積`t(株)山田商事`r`n`r`nabc 見積`r`n確定`tABC`r`n", ${utf8Bom})
    [System.IO.File]::WriteAllText("$dir\b.docx\ページ001.tsv", "りんご`nみかん abc`n`n見積 確定`nabc", ${utf8Bom})
    [System.IO.File]::WriteAllText("$dir\c.pptx\スライド001.tsv", "abc`rりんご`r`r見積`r", (New-Object System.Text.UTF8Encoding $false))
    [System.IO.File]::WriteAllText("$dir\d.xlsx\空.tsv", "", ${utf8Bom})
    $packs = newPackIndex $tsvRoot (Join-Path $TestDrive "modes_pack")

    $cases = @(
        @("abc", $true), @("見積", $true), @("(株)", $true), @("存在しない", $true), @("`t", $true),
        @("^abc", $false), @("abc$", $false), @("^$", $false), @("x*", $false), @("見積.*確定", $false),
        @("見積\s確定", $false), @("\s", $false), @("[^a]", $false), @("abc(?=\s)", $false),
        @("abc(?!\s)", $false), @("\Aabc", $false), @("abc\z", $false), @("(?i)ABC", $false)
    )
    foreach ($case in $cases) {
        It "「$($case[0])」（文字どおり=$($case[1])）" {
            sortedKeys (searchPackIndex $case[0] $packs $case[1]).Hits | Should BeExactly (referenceKeys $tsvRoot $case[0] $case[1])
        }
    }

    It "行番号は 1 行ずつ読んだときと同じ（CR だけの改行・空行も数える）" {
        sortedKeys (searchPackIndex "見積" $packs $true).Hits |
            Should BeExactly ((@("idx|a.xlsx|S|1|見積`t(株)山田商事", "idx|a.xlsx|S|3|abc 見積", "idx|b.docx|ページ001|4|見積 確定", "idx|c.pptx|スライド001|4|見積") | Sort-Object) -join "`n")
    }
}

Describe "searchPackIndex（検索語・条件・上限・中止）" -Tag Io {
    $tsvRoot = Join-Path $TestDrive "search_tsv"
    newTsv "$tsvRoot\A社.xlsx\Sheet1.tsv" @("見積先：`t(株)山田商事", "", "株式会社`t1.5", "ABC`t105")
    newTsv "$tsvRoot\sub\文書.docx\ページ001.tsv" @("りんご (株) abc")
    $packRoot = Join-Path $TestDrive "search_pack"
    $packs = newPackIndex $tsvRoot $packRoot

    It "文字どおりに検索すると (株) は記号のまま探す" {
        $result = searchPackIndex "(株)" $packs $true
        $result.Hits.Count | Should Be 2
        $result.SimpleMatch | Should Be $true
    }

    It "正規表現として検索すると (株) は「株」に一致する" {
        $result = searchPackIndex "(株)" $packs $false
        $result.Hits.Count | Should Be 3
        $result.SimpleMatch | Should Be $false
    }

    It "正規表現として不正なワードは文字どおりに検索する" {
        $result = searchPackIndex "(" $packs $false
        $result.SimpleMatch | Should Be $true
        $result.Hits.Count | Should Be 2
    }

    It "ファイル名・場所・行番号・相対フォルダ・インデックスのフォルダを返す（空行も行番号に数える）" {
        $hit = @((searchPackIndex "1.5" $packs $true).Hits)[0]
        $hit.Book | Should Be "A社.xlsx"
        $hit.Location | Should Be "Sheet1"
        $hit.LineNumber | Should Be 3
        $hit.RelDir | Should Be ""
        $hit.Root | Should Be $packRoot
        @((searchPackIndex "りんご" $packs $true).Hits)[0].RelDir | Should Be "sub"
    }

    It "上限を超えたら打ち切り、上限ちょうどなら打ち切りにしない。上限 0 は上限なし" {
        $result = searchPackIndex "株" $packs $true 2
        $result.Hits.Count | Should Be 2
        $result.Truncated | Should Be $true
        $result = searchPackIndex "株" $packs $true 3
        $result.Hits.Count | Should Be 3
        $result.Truncated | Should Be $false
        (searchPackIndex "株" $packs $true 0).Hits.Count | Should Be 3
    }

    It "進み具合（照合した集約ファイルの数）を知らせ、中止できる" {
        $script:progress = @()
        $result = searchPackIndex "株" $packs $true -taskBytes 1 -workerCount 1 -onProgress { param($done, $total, $newHits) $script:progress += "$done/$total" } -shouldStop { $script:progress.Count -ge 1 }
        $result.Cancelled | Should Be $true
        ($script:progress -join ",") | Should Be "1/2"
    }

    It "対象ファイルで、元のファイル名が一致するものだけを検索する" {
        @((searchPackIndex "株" $packs $true -fileFilter "*.docx").Hits).Count | Should Be 1
        @((searchPackIndex "株" $packs $true -fileFilter "!*.docx").Hits).Count | Should Be 2
        @((searchPackIndex "株" $packs $true -fileFilter "*.pptx").Hits).Count | Should Be 0
    }

    It "検索の途中で読めなくなった集約ファイル（インデックス作成中に削除された等）は飛ばす。対象が無ければ 0 件" {
        $copyRoot = Join-Path $TestDrive "deleted_pack"
        $copies = newPackIndex $tsvRoot $copyRoot
        Remove-Item -LiteralPath (fromLongPath $copies[0].Path)
        $hits = @((searchPackIndex "株" $copies $true).Hits)
        $hits.Count | Should Be 1
        $none = searchPackIndex "株" @() $true
        $none.Total | Should Be 0
        $none.Hits.Count | Should Be 0
        $none.Truncated | Should Be $false
        $none.Cancelled | Should Be $false
    }
}

Describe "searchPackIndex（正規表現の照合の時間切れ）" -Tag Io {
    # 入れ子の繰り返し（(a+)+$）は、一致しない文字列で照合の時間が指数的に増える
    $tsvRoot = Join-Path $TestDrive "timeout_tsv"
    for ($i = 0; $i -lt 3; $i++) {
        newTsv "$tsvRoot\s$i\book$i.xlsx\S.tsv" @(("a" * 40) + "!")
    }
    $packs = newPackIndex $tsvRoot (Join-Path $TestDrive "timeout_pack")
    $message = "正規表現の照合に時間がかかりすぎるため、検索を中止しました。正規表現を見直してください。"

    It "時間切れなら、分かるメッセージで検索を中止する" {
        $regexTimeout = [timespan]::FromMilliseconds(1)
        { searchPackIndex "(a+)+$" $packs $false -workerCount 1 } | Should Throw $message
    }

    It "並列に検索していても、時間切れなら同じメッセージで中止する" {
        $regexTimeout = [timespan]::FromMilliseconds(1)
        { searchPackIndex "(a+)+$" $packs $false -workerCount 2 -taskBytes 1 } | Should Throw $message
    }

    It "照合のスレッドのスクリプトは、時間切れを例外にせず Timeout で返す（このスレッドで直接動かす）" {
        $regexTimeout = [timespan]::FromMilliseconds(1)
        $search = newSearchRegex "(a+)+$" $false $false
        $output = & ${packWorkerScript} $packs 0 $packs.Count $search.Regex -1 $search.TextRegex $search.ScanMode $null $null $null $null
        $output.Timeout | Should Be $true
        $output.Hits | Should BeNullOrEmpty
    }

    It "照合のスレッドのスクリプトは、ヒットを Hits で返す" {
        $search = newSearchRegex "a!" $true $false
        $output = & ${packWorkerScript} $packs 0 $packs.Count $search.Regex -1 $search.TextRegex $search.ScanMode $null $null $null $null
        $output.Timeout | Should Be $false
        @($output.Hits).Count | Should Be 3
    }
}

Describe "searchPackIndex（並列検索・読んだ内容の使い回し）" -Tag Io {
    $tsvRoot = Join-Path $TestDrive "parallel_tsv"
    for ($i = 0; $i -lt 30; $i++) {
        newTsv ("$tsvRoot\sub{0}\book{1:D2}.xlsx\S.tsv" -f ($i % 3), $i) @("見積 $i", "x", "見積 確定 $i")
    }
    $packRoot = Join-Path $TestDrive "parallel_pack"
    $packs = newPackIndex $tsvRoot $packRoot

    function formatBookLines {
        param ($hits)
        return (@($hits | ForEach-Object { "$($_.Book):$($_.LineNumber)" }) -join ",")
    }

    It "並列に検索しても、集約ファイルの順に同じ結果を返す" {
        $expected = formatBookLines (searchPackIndex "見積" $packs $true -workerCount 1).Hits
        $result = searchPackIndex "見積" $packs $true -workerCount 3 -taskBytes 1
        formatBookLines $result.Hits | Should Be $expected
        $result.Hits.Count | Should Be 60
    }

    It "並列でも上限で打ち切り、進み具合を集約ファイルの順に知らせる" {
        $script:done = New-Object System.Collections.Generic.List[int]
        $result = searchPackIndex "見積" $packs $true 5 -workerCount 3 -taskBytes 1 -onProgress { param($done, $total, $newHits) $script:done.Add($done) }
        $result.Truncated | Should Be $true
        formatBookLines $result.Hits | Should Be "book00.xlsx:1,book00.xlsx:3,book03.xlsx:1,book03.xlsx:3,book06.xlsx:1"
        ($script:done -join ",") | Should Be ((@($script:done | Sort-Object)) -join ",")
    }

    It "並列でも中止できる" {
        $script:calls = 0
        $result = searchPackIndex "見積" $packs $true -workerCount 3 -taskBytes 1 -onProgress { param($done, $total, $newHits) $script:calls++ } -shouldStop { $script:calls -ge 2 }
        $result.Cancelled | Should Be $true
        $script:calls | Should Be 2
    }

    It "内容を使い回し、書き直された集約ファイルは読み直す" {
        $cache = newTsvTextCache
        (searchPackIndex "更新後" $packs $true -cache $cache).Hits.Count | Should Be 0
        newTsv "$packRoot\sub0\book00.xlsx\S.tsv" @("更新後の内容")
        [void](updateIndexFolderPack "$packRoot\sub0")
        (Get-Item -LiteralPath "$packRoot\sub0\content.xlsx.001.tsv").LastWriteTime = (Get-Date).AddMinutes(1)
        (searchPackIndex "更新後" (getPackFiles $packRoot) $true -cache $cache).Hits.Count | Should Be 1
    }

    It "上限を超える内容は残さない" {
        $cache = newTsvTextCache 20
        [void](searchPackIndex "見積" $packs $true -cache $cache)
        $cache.Chars[0] | Should BeLessThan 21
    }
}

Describe "getIndexPackFiles" -Tag Io {
    $index = Join-Path $TestDrive "list[1]"
    $other = Join-Path $TestDrive "list_other"
    foreach ($item in @(
            @{ Folder = "$index"; Name = "b.xlsx" }, @{ Folder = "$index\sub"; Name = "a.xlsx" }, @{ Folder = "$other"; Name = "c.docx" })) {
        [void][System.IO.Directory]::CreateDirectory($item.Folder)
        writePackFile "$($item.Folder)\$(getPackFileName (getPackExtension $item.Name))" (convertToPackText @(@{ Name = $item.Name; Places = @(@{ Place = "S"; Text = "x" }) }))
    }
    $missing = Join-Path $TestDrive "list_missing"

    It "フォルダごとの件数と、インデックスのフォルダからの相対パスを返す" {
        $result = getIndexPackFiles @($index, $missing, $other)
        $result.Folders.Count | Should Be 3
        $result.Folders[0].Exists | Should Be $true
        $result.Folders[0].Count | Should Be 2
        $result.Folders[1].Exists | Should Be $false
        $result.Packs.Count | Should Be 3
        @($result.Packs | Where-Object { $_.RelPath -eq "sub\content.xlsx.001.tsv" }).Count | Should Be 1
    }

    It "入れ子のフォルダを指定しても同じ集約ファイルを重複させない" {
        (getIndexPackFiles @($index, "$index\sub")).Packs.Count | Should Be 2
    }

    It "インデックスの中のフォルダだけ・直下だけを指定できる" {
        $result = getIndexPackFiles @(@{ Root = $index; RelPath = "sub"; Recurse = $true })
        $result.Packs.Count | Should Be 1
        $result.Packs[0].Root | Should Be $index
        $result.Packs[0].RelDir | Should Be "sub"
        $result.Folders[0].Path | Should Be "$index\sub"
        $direct = getIndexPackFiles @(@{ Root = $index; RelPath = ""; Recurse = $false })
        $direct.Packs.Count | Should Be 1
        $direct.Packs[0].RelPath | Should Be "content.xlsx.001.tsv"
    }

    It "存在しないフォルダは Exists が false。数えた件数を知らせる" {
        $result = getIndexPackFiles @(@{ Root = $index; RelPath = "なし"; Recurse = $true }, @{ Root = $missing; RelPath = "sub"; Recurse = $true })
        ($result.Folders | ForEach-Object { $_.Exists }) -join "," | Should Be "False,False"
        $result.Packs.Count | Should Be 0
        $counts = New-Object System.Collections.Generic.List[int]
        [void](getIndexPackFiles @($index, $other) { param ($count) $counts.Add($count) })
        ($counts -join ",") | Should Be "2,3"
    }
}

Describe "readPackContext" -Tag Io {
    $folder = Join-Path $TestDrive "context"
    [void][System.IO.Directory]::CreateDirectory($folder)
    $path = "$folder\content.xlsx.001.tsv"
    writePackFile $path (convertToPackText @(
            @{ Name = "A.xlsx"; Places = @(@{ Place = "S"; Text = "1`r`n2`r`n3`r`n4`r`n5`r`n6`r`n" }, @{ Place = "T"; Text = "t1`r`n" }) },
            @{ Name = "B.xlsx"; Places = @(@{ Place = "S"; Text = "b1`r`nb2`r`n" }) }))

    function formatContext {
        param ($rows)
        return (@($rows | ForEach-Object { "$($_.LineNumber):$($_.Line)" }) -join ",")
    }

    It "元のファイル・場所の中で、その行と前後の行を返す（ほかの場所の行は入れない）" {
        formatContext (readPackContext $path "A.xlsx" "S" 3 1 2) | Should Be "2:2,3:3,4:4,5:5"
        formatContext (readPackContext $path "A.xlsx" "S" 6 2 3) | Should Be "4:4,5:5,6:6"
        formatContext (readPackContext $path "B.xlsx" "S" 1 3 3) | Should Be "1:b1,2:b2"
    }

    It "見つからない・読めないときは空。キャッシュに同じ内容があればそれを使う" {
        @(readPackContext $path "無い.xlsx" "S" 1).Count | Should Be 0
        @(readPackContext "$folder\無い.tsv" "A.xlsx" "S" 1).Count | Should Be 0
        $packs = getPackFiles $folder
        $cache = newTsvTextCache
        [void](searchPackIndex "t1" $packs $true -cache $cache)
        # キャッシュの内容を書き換えると、ファイルではなくキャッシュから読んだことが分かる
        $entry = $cache.Texts[$packs[0].Path]
        $entry[2] = $entry[2].Replace("t1", "c1")
        formatContext (readPackContext $path "A.xlsx" "T" 1 0 0 $cache) | Should Be "1:c1"
    }
}

