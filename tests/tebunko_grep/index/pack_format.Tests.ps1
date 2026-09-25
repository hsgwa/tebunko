# 検索用のまとめファイルの形式（tebunko_grep\index\pack_format.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"

Describe "convertPlaceToPackMeta / convertPackMetaToPlace" -Tag Unit {
    It "今の場所の名前をメタ情報に分け、元の名前に戻せる" {
        $cases = @(
            @("見積.xlsx", "見積", "シート=見積|対象=本文"),
            @("見積.xlsx", "見積[図形]", "シート=見積|対象=図形"),
            @("見積.xlsx", "見積[コメント]", "シート=見積|対象=コメント"),
            @("見積.xlsx", "ページ001", "シート=ページ001|対象=本文"),
            @("議事録.docx", "ページ001", "ページ=1|対象=本文"),
            @("議事録.docx", "ページ012[図形]", "ページ=12|対象=図形"),
            @("議事録.docx", "ヘッダー・フッター", "部分=ヘッダー・フッター|対象=本文"),
            @("議事録.docx", "文書[コメント]", "部分=文書|対象=コメント"),
            @("議事録.doc", "脚注", "部分=脚注|対象=本文"),
            @("提案.pptx", "スライド001", "スライド=1|対象=本文"),
            @("提案.pptx", "スライド001_ノート", "スライド=1|対象=ノート"),
            @("提案.pptx", "スライド002（非表示）", "スライド=2|対象=本文|非表示=はい"),
            @("提案.pptx", "スライド002（非表示）[図形]", "スライド=2|対象=図形|非表示=はい"),
            @("提案.pptx", "ヘッダー・フッター", "部分=ヘッダー・フッター|対象=本文")
        )
        foreach ($c in $cases) {
            $meta = convertPlaceToPackMeta $c[0] $c[1]
            (($meta.Keys | ForEach-Object { "$_=$($meta[$_])" }) -join "|") | Should Be $c[2]
            convertPackMetaToPlace $meta | Should BeExactly $c[1]
        }
    }

    It "種類の分からないファイルは、部分に場所の名前を持つ" {
        getPackFileKind "メモ.txt" | Should Be ""
        $meta = convertPlaceToPackMeta "メモ.txt" "本文[図形]"
        (($meta.Keys | ForEach-Object { "$_=$($meta[$_])" }) -join "|") | Should Be "部分=本文|対象=図形"
        convertPackMetaToPlace $meta | Should BeExactly "本文[図形]"
    }

    It "組み立て直すと同じにならない名前は、部分にそのまま持つ" {
        $meta = convertPlaceToPackMeta "議事録.docx" "ページ1"
        $meta["部分"] | Should Be "ページ1"
        convertPackMetaToPlace $meta | Should BeExactly "ページ1"
    }
}

Describe "getPackFileName / splitPackBooksByExtension" -Tag Unit {
    It "まとめファイルの名前に元のファイルの拡張子（小文字）を入れる" {
        getPackExtension "見積.XLSX" | Should Be "xlsx"
        getPackFileName (getPackExtension "議事録.docx") | Should Be "content.docx.tsv"
    }

    It "元のファイルを拡張子ごとに分け、それぞれの中の順は変えない" {
        $books = @(@{ Name = "a.xlsx" }, @{ Name = "b.docx" }, @{ Name = "c.XLSX" }, @{ Name = "d.xlsm" })
        $groups = splitPackBooksByExtension $books
        @($groups.Keys) -join "," | Should Be "xlsx,docx,xlsm"
        @($groups["xlsx"] | ForEach-Object { $_.Name }) -join "," | Should Be "a.xlsx,c.XLSX"
    }
}

Describe "encodePackValue / decodePackValue" -Tag Unit {
    It "改行・制御文字・% を符号化し、タブはそのまま残して戻せる" {
        $value = "50%引き`n2行目`t" + [char]0x1E
        $encoded = encodePackValue $value
        $encoded | Should Be "50%25引き%0A2行目`t%1E"
        decodePackValue $encoded | Should BeExactly $value
    }
}

Describe "convertToPackBody" -Tag Unit {
    It "改行を LF にそろえ、ReadLine と同じ行に分かれる形にする" {
        convertToPackBody "a`r`n`r`nb`r`n" | Should BeExactly "a`n`nb`n"
        convertToPackBody "a`rb" | Should BeExactly "a`nb`n"
        convertToPackBody "" | Should BeExactly ""
        convertToPackBody "`r`n" | Should BeExactly "`n"
    }

    It "区切りに使う U+001C〜U+001F を取り除く" {
        convertToPackBody ("a" + [char]0x1E + "b") | Should BeExactly "ab`n"
    }
}

Describe "convertToPackText / readPackPlaces" -Tag Unit {
    $books = @(
        @{ Name = "見積.xlsx"; Places = @(@{ Place = "見積"; Text = "品名`t単価`r`n`r`nりんご`t100`r`n" }, @{ Place = "見積[図形]"; Text = "E2`t承認済み`r`n" }) },
        @{ Name = "議事録.docx"; Places = @(@{ Place = "ページ001"; Text = "方針`r`n" }, @{ Place = "空"; Text = "" }) }
    )
    $text = convertToPackText $books

    It "メタ情報の行と中身を並べる" {
        $mark = [string][char]0x1E
        $lines = $text.Split("`n")
        $lines[0] | Should Be "$mark 版=1"
        $lines[1] | Should Be "$mark ファイル名=見積.xlsx"
        $lines[2] | Should Be "$mark 種類=Excel"
        $lines[3] | Should Be "$mark シート=見積"
        $lines[4] | Should Be "$mark 対象=本文"
        $lines[5] | Should Be "品名`t単価"
        $lines[6] | Should Be ""
    }

    It "場所ごとに、元のファイル名・場所の名前・中身の範囲を読み取る" {
        $places = readPackPlaces $text
        $places.Count | Should Be 4
        $places[0].Book | Should Be "見積.xlsx"
        $places[0].Location | Should Be "見積"
        $text.Substring($places[0].Start, $places[0].End - $places[0].Start) | Should BeExactly "品名`t単価`n`nりんご`t100`n"
        $places[1].Location | Should Be "見積[図形]"
        $places[2].Book | Should Be "議事録.docx"
        $places[2].Location | Should Be "ページ001"
        $places[3].Location | Should Be "空"
        $places[3].End - $places[3].Start | Should Be 0
    }

    It "元のファイルごとのまとまりに分け、そのまま並べ直すと元に戻る" {
        $blocks = splitPackTextByBook $text
        @($blocks | ForEach-Object { $_.Name }) -join "," | Should Be "見積.xlsx,議事録.docx"
        $blocks[0].Block.StartsWith([string][char]0x1E + " ファイル名=見積.xlsx`n") | Should Be $true
        convertToPackText $blocks | Should BeExactly $text
        (splitPackTextByBook "").Count | Should Be 0
        { splitPackTextByBook ([string][char]0x1E + " 版=2`n") } | Should Throw
    }

    It "版が違う・無いときは例外にする" {
        { readPackPlaces ([string][char]0x1E + " 版=2`n") } | Should Throw
        { readPackPlaces ([string][char]0x1E + " ファイル名=a.xlsx`n") } | Should Throw
    }

    It "末尾に改行が無くても、最後のメタ情報の行まで読む" {
        $mark = [string][char]0x1E
        $places = readPackPlaces "$mark 版=1`n$mark ファイル名=a.xlsx`n$mark シート=S`n$mark 対象=図形"
        $places.Count | Should Be 1
        $places[0].Location | Should Be "S[図形]"
    }
}
