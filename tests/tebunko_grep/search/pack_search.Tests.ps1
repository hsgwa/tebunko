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

    It "フォルダごとに 1 つのまとめファイルを作る" {
        $packs.Count | Should Be 2
        $packs[0].RelDir | Should Be "営業"
        $packs[1].RelDir | Should Be "営業\2025"
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
        $cache.Texts.Count | Should Be 2
        # ファイルを消しても、更新日時・大きさが同じ（列挙したときの値）ならキャッシュから探せる
        $second = toKeys (searchPackIndex "単価" $packs $true -cache $cache).Hits
        $second -join "`n" | Should BeExactly ($first -join "`n")
    }

    It "並列でも 1 スレッドと同じ結果を同じ順で返す" {
        $single = toKeys (searchPackIndex "単価" $packs $true -workerCount 1).Hits
        $tasks = splitPackTasks $packs 1
        $tasks.Count | Should Be 2
        # 1 バイトごとに分けて、2 つのスレッドで探す
        $parallel = toKeys (searchPackIndex "単価" $packs $true -workerCount 2 -taskBytes 1).Hits
        $parallel -join "`n" | Should BeExactly ($single -join "`n")
    }

    It "書き直すと中身が置き換わる" {
        $path = Join-Path $TestDrive "rewrite\${packFileName}"
        [void][System.IO.Directory]::CreateDirectory((Split-Path $path))
        writePackFile $path "a"
        writePackFile $path "b"
        readPackText $path | Should BeExactly "b"
        [System.IO.File]::Exists("$path.tmp") | Should Be $false
    }
}
