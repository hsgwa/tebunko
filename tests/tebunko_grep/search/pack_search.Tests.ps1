# 検索用のまとめファイルの検索（tebunko_grep\search\pack_search.ps1）と読み書き（index\pack_store.ps1）のテスト。
# 同じインデックスを、今の TSV の検索とまとめファイルの検索で探し、結果が同じになることを確かめる
. "$PSScriptRoot\..\..\helpers\load.ps1"

function script:toKeys {
    # ヒットを比べられる文字列にする（まとめファイルと TSV で違う RelPath・FileName は比べない）
    param ($hits)
    return @($hits | ForEach-Object { "{0}|{1}|{2}|{3}|{4}" -f $_.RelDir, $_.Book, $_.Location, $_.LineNumber, $_.Line })
}

Describe "まとめファイルの検索" -Tag Io {
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
    foreach ($folder in "$idx", "$idx\2025") {
        convertIndexFolderToPack $folder ($folder.Replace($tsvRoot, $packRoot)) | Out-Null
    }
    $tsvFiles = (getIndexTsvFiles @($tsvRoot)).Files
    $packs = getPackFiles $packRoot

    It "フォルダごと・拡張子ごとにまとめファイルを作り、フォルダの順・名前の順に並べる" {
        $packs.Count | Should Be 3
        $packs[0].RelPath | Should Be "営業\本文.docx.tsv"
        $packs[1].RelPath | Should Be "営業\本文.xlsx.tsv"
        $packs[2].RelPath | Should Be "営業\2025\本文.pptx.tsv"
    }

    It "元のファイルが無くなった拡張子のまとめファイルは、変換し直すときに消す" {
        $dest = Join-Path $TestDrive "reconvert"
        [void][System.IO.Directory]::CreateDirectory($dest)
        writePackFile "$dest\本文.pptx.tsv" (convertToPackText @(@{ Name = "古い.pptx"; Places = @(@{ Place = "スライド001"; Text = "古い" }) }))
        $result = convertIndexFolderToPack $idx $dest @("古い.pptx")
        $result.Files | Should Be 2
        @([System.IO.Directory]::GetFiles($dest) | ForEach-Object { [System.IO.Path]::GetFileName($_) } | Sort-Object) -join "," | Should Be "本文.docx.tsv,本文.xlsx.tsv"
    }

    It "まとめファイルは UTF-16LE（BOM 付き）で、一時ファイルを残さない" {
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
        It "「$($case.Word)」の結果が TSV の検索と同じ" {
            $expected = toKeys (searchIndex $case.Word $tsvFiles $case.Simple).Hits
            $actual = toKeys (searchPackIndex $case.Word $packs $case.Simple).Hits
            ($actual | Sort-Object) -join "`n" | Should BeExactly (($expected | Sort-Object) -join "`n")
        }
    }

    It "大文字・小文字の区別、図形・コメントの除外、対象ファイルの条件が TSV の検索と同じ" {
        $expected = toKeys (searchIndex "ABC" $tsvFiles $true -caseSensitive $true).Hits
        (toKeys (searchPackIndex "ABC" $packs $true -caseSensitive $true).Hits) -join "`n" | Should BeExactly ($expected -join "`n")
        $expected = toKeys (searchIndex "単価" $tsvFiles $true -includeShapes $false -includeComments $false).Hits
        ((toKeys (searchPackIndex "単価" $packs $true -includeShapes $false -includeComments $false).Hits) | Sort-Object) -join "`n" | Should BeExactly (($expected | Sort-Object) -join "`n")
        $expected = toKeys (searchIndex "単価" $tsvFiles $true -fileFilter "*.docx").Hits
        (toKeys (searchPackIndex "単価" $packs $true -fileFilter "*.docx").Hits) -join "`n" | Should BeExactly ($expected -join "`n")
    }

    It "上限で打ち切る" {
        $result = searchPackIndex "単価" $packs $true 2
        $result.Hits.Count | Should Be 2
        $result.Truncated | Should Be $true
    }

    It "キャッシュを使っても結果が同じで、2 回目はファイルを読まない" {
        $cache = newTsvTextCache
        $first = toKeys (searchPackIndex "単価" $packs $true -cache $cache).Hits
        $cache.Texts.Count | Should Be 3
        # ファイルを消しても、更新日時・大きさが同じ（列挙したときの値）ならキャッシュから探せる
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

    It "更新日時が変わったまとめファイルは読み直し、キャッシュの古い内容を置き換える" {
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
        $expected = ((toKeys (searchPackIndex "単価" $packs $true).Hits) | Sort-Object) -join "`n"
        ((toKeys $hits) | Sort-Object) -join "`n" | Should BeExactly $expected
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

    It "置かれた TSV をまとめファイルに入れて TSV を消し、変わらない元のファイルは前のまとめファイルから写す" {
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
        @($inPacks | ForEach-Object { [System.IO.Path]::GetFileName($_.RelPath) }) -join "," | Should Be "本文.docx.tsv,本文.xlsx.tsv"
        (toKeys (searchPackIndex "単価" $inPacks $true).Hits) -join "`n" | Should BeExactly ((
            "営業|C.docx|ページ001|1|C の単価", "営業|A.xlsx|S|1|A の単価", "営業|B.xlsx|S|1|B の新しい単価") -join "`n")
        # C が無くなったら外し、Word のまとめファイルを消す
        (updateIndexFolderPack $dir @("C.docx")).Books | Should Be 2
        @([System.IO.Directory]::GetFiles($dir) | ForEach-Object { [System.IO.Path]::GetFileName($_) }) -join "," | Should Be "本文.xlsx.tsv"
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
