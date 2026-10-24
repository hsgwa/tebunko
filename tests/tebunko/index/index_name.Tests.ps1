# インデックス名と TSV のファイル名（tebunko\index\index_name.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "encodeIndexPlace / decodeIndexPlace" -Tag Unit {
    # PowerShell は ” を " と同じに扱うため、' で囲む
    It "<name>" -TestCases @(
        @{ name = "ファイル名に使えない文字・_・% を %XX にし、元に戻せる"; place = 'a<b>c\d*e:f?g|h/i"j_k%l'; encoded = 'a%3Cb%3Ec%5Cd%2Ae%3Af%3Fg%7Ch%2Fi%22j%5Fk%25l' }
        @{ name = "制御文字も %XX にする"; place = "a`tb"; encoded = 'a%09b' }
        @{ name = "半角の記号と全角の記号は別の名前になる（衝突しない）"; place = '衝突"'; encoded = '衝突%22' }
        @{ name = "使える文字（全角記号・空白・&'#() 等）はそのまま返す"; place = 'シート1 (2)&''#＜” 全角　'; encoded = 'シート1 (2)&''#＜” 全角　' }
    ) {
        param ($name, $place, $encoded)
        encodeIndexPlace $place | Should -BeExactly $encoded
        decodeIndexPlace $encoded | Should -BeExactly $place
    }

    It "符号化で作らない %XX はそのまま返す（シート名 100% 等）" {
        decodeIndexPlace '100%' | Should -Be '100%'
        decodeIndexPlace '%41%2a' | Should -Be '%41%2a'
    }
}

Describe "splitObjectPlace" -Tag Unit {
    It "図形・コメントの場所を、元の場所と種類に分ける（Excel のシート・Word のページ・PowerPoint のスライド）" {
        $shape = splitObjectPlace "売上[図形]"
        $shape.Base | Should -Be "売上"
        $shape.Kind | Should -Be "図形"
        (splitObjectPlace "売上 (2)[コメント]").Base | Should -Be "売上 (2)"
        (splitObjectPlace "ページ003[コメント]").Kind | Should -Be "コメント"
        (splitObjectPlace "スライド002[図形]").Base | Should -Be "スライド002"
    }

    It "ふつうの場所・知らない種類はそのまま（種類は空）" {
        $plain = splitObjectPlace "ページ001"
        $plain.Base | Should -Be "ページ001"
        $plain.Kind | Should -Be ""
        (splitObjectPlace "売上[メモ]").Kind | Should -Be ""
    }

    It "種類の名前が、書き出す側（office_reader.ps1）でも同じ" {
        # shared はツールの変数を使えないため、書き出す側は同じ名前を別に書いている
        $reader = [System.IO.File]::ReadAllText("${scriptsDir}\shared\office\office_reader.ps1")
        foreach ($kind in @(${placeKindShape}, ${placeKindComment}, ${placeKindHeaderFooter})) {
            $reader.Contains("[$kind]") | Should -Be $true
        }
        $reader.Contains("[${placeKindEmbed}") | Should -Be $true
    }
}

Describe "newObjectPlaceKindPattern" -Tag Unit {
    It "省略すると、表のすべての種類を当てる" {
        $pattern = newObjectPlaceKindPattern
        foreach ($kind in ${placeKindFileNames}.Keys) {
            $place = if ($kind -eq ${placeKindEmbed}) { "ページ001[${kind}1]" } else { "売上[$kind]" }
            $place | Should -Match $pattern
        }
    }

    It "場所の末尾の種類だけを、種類を絞って当てる（-kinds）" -TestCases @(
        @{ Kinds = @("図形");            Place = "売上[図形]";            Expected = $true }
        @{ Kinds = @("図形");            Place = "売上[コメント]";        Expected = $false }
        @{ Kinds = @("図形", "埋め込み"); Place = "ページ001[埋め込み12]"; Expected = $true }
        @{ Kinds = @("図形", "埋め込み"); Place = "ページ001[埋め込み]";   Expected = $false }
        @{ Kinds = @("図形", "埋め込み"); Place = "ページ001[埋め込み01]"; Expected = $false }
        @{ Kinds = @("コメント");        Place = "売上[コメント]";        Expected = $true }
        @{ Kinds = @("コメント");        Place = "売上[コメント]の続き";  Expected = $false }
    ) {
        param ($Kinds, $Place, $Expected)
        [regex]::new((newObjectPlaceKindPattern -kinds $Kinds)).IsMatch($Place) | Should -Be $Expected
    }

    It "全部を当てるとき、[埋め込み]・[埋め込み01]・知らない種類は当たらない" -TestCases @(
        @{ Place = "ページ001[埋め込み]";   Expected = $false }
        @{ Place = "ページ001[埋め込み01]"; Expected = $false }
        @{ Place = "ページ001[埋め込み1]";  Expected = $true }
        @{ Place = "売上[ヘッダー・フッター]"; Expected = $true }
        @{ Place = "売上[メモ]";            Expected = $false }
        @{ Place = "売上";                  Expected = $false }
    ) {
        param ($Place, $Expected)
        [regex]::new((newObjectPlaceKindPattern)).IsMatch($Place) | Should -Be $Expected
    }
}

Describe "describePlace" -Tag Unit {
    BeforeAll {
        function described([string]$book, [string]$place) {
            $d = describePlace $book $place
            return "$($d.Place)|$($d.Kind)"
        }
    }

    It "Excel はシート名と、セル・図形・コメントの種別にする（検索条件のチェックと同じ言葉）" {
        described "見積.xlsx" "売上" | Should -Be "[シート]売上|セル"
        described "見積.xlsx" "売上[図形]" | Should -Be "[シート]売上|図形"
        described "見積.xlsx" "売上[コメント]" | Should -Be "[シート]売上|コメント"
        described "見積.xlsx" "売上[ヘッダー・フッター]" | Should -Be "[シート]売上|ヘッダー・フッター"
        # シート名が「ページ001」でも、Excel ならシートとして出す
        described "旧.XLS" "ページ001" | Should -Be "[シート]ページ001|セル"
        # シート名が固定名のファイル名（page_001）・Word の固定の場所（ヘッダー・フッター）と同じ文字列でも、Excel ならシートとして出す
        described "旧.XLS" "page_001" | Should -Be "[シート]page_001|セル"
        described "旧.XLS" "ヘッダー・フッター" | Should -Be "[シート]ヘッダー・フッター|セル"
    }

    It "Word のページは番号にし、目安であることを付ける。番号の無い場所は名前のまま" {
        described "報告.docx" "ページ003" | Should -Be "3 ページ（目安）|本文"
        described "報告.docx" "ページ120" | Should -Be "120 ページ（目安）|本文"
        described "報告.docx" "ヘッダー・フッター" | Should -Be "ヘッダー・フッター|本文"
        described "報告.docx" "脚注" | Should -Be "脚注|本文"
        # Word のコメント・図形も同じ決まりで出す
        described "報告.docx" "ページ003[コメント]" | Should -Be "3 ページ（目安）|コメント"
    }

    It "PDF のページは番号にする（Word と違い目安ではない）" {
        described "仕様書.pdf" "ページ003" | Should -Be "3 ページ|本文"
        described "仕様書.PDF" "ページ120" | Should -Be "120 ページ|本文"
    }

    It "PowerPoint はスライド番号にし、非表示はそのまま付け、ノートは種別で分ける" {
        described "提案.pptx" "スライド001" | Should -Be "スライド 1|本文"
        described "提案.pptx" "スライド002（非表示）" | Should -Be "スライド 2（非表示）|本文"
        described "提案.pptx" "スライド002_ノート" | Should -Be "スライド 2|ノート"
        described "提案.pptx" "スライド002[図形]" | Should -Be "スライド 2|図形"
    }

    It "PowerPoint のスライドマスターは本文、ノートマスターはノート（図形の種類が付けば、その種別）" {
        described "提案.pptx" "スライドマスター" | Should -Be "スライドマスター|本文"
        described "提案.pptx" "ノートマスター" | Should -Be "ノートマスター|ノート"
        described "提案.pptx" "ノートマスター[図形]" | Should -Be "ノートマスター|図形"
    }

    It "場所が空なら空" {
        described "a.docx" "" | Should -Be "|本文"
    }

    It "テキストは場所が「本文」だけなので、表記は空・種別は本文" {
        described "議事メモ.txt" "本文" | Should -Be "|本文"
    }
}

Describe "describeHitPlace" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "Excel のセルは、場所ごとの表記にセル番地を足す"; place = "[シート]売上"; excel = $true; object = $false; cell = "B12"; count = 1; line = 12; expected = "[シート]売上!B12" }
        @{ name = "Excel の 1 行に複数のセルが一致したら、ほかの数を足す"; place = "[シート]売上"; excel = $true; object = $false; cell = "B12"; count = 3; line = 12; expected = "[シート]売上!B12 ほか 2" }
        @{ name = "Excel のセル番地が求まらないときは行番号を足す"; place = "[シート]売上"; excel = $true; object = $false; cell = ""; count = 0; line = 12; expected = "[シート]売上 12 行目" }
        @{ name = "Excel の図形・コメントは、左上・コメントのセル番地を足す"; place = "[シート]売上"; excel = $true; object = $true; cell = "D5"; count = 1; line = 1; expected = "[シート]売上!D5" }
        @{ name = "Excel の図形・コメントでセル番地が求まらないときは、通し番号のため行番号を出さない"; place = "[シート]売上"; excel = $true; object = $true; cell = ""; count = 0; line = 2; expected = "[シート]売上" }
        @{ name = "Excel のヘッダー・フッターは、セル番地が無いのでシート名だけ"; place = "[シート]売上"; excel = $true; object = $true; cell = ""; count = 1; line = 3; expected = "[シート]売上" }
        @{ name = "Word は場所ごとの表記のまま"; place = "3 ページ（目安）"; excel = $false; object = $false; cell = ""; count = 0; line = 5; expected = "3 ページ（目安）" }
        @{ name = "PowerPoint は場所ごとの表記のまま"; place = "スライド 9（非表示）"; excel = $false; object = $true; cell = ""; count = 0; line = 2; expected = "スライド 9（非表示）" }
        @{ name = "テキストは行番号だけ（場所は 1 つしかないため）"; place = ""; excel = $false; object = $false; cell = ""; count = 0; line = 12; expected = "12 行目"; text = $true }
    ) {
        param ($name, $place, $excel, $object, $cell, $count, $line, $expected, [bool]$text = $false)
        describeHitPlace $place $excel $object $cell $count $line $text | Should -Be $expected
    }
}

Describe "toIndexFileName" -Tag Unit {
    It "<場所>.tsv にし、場所は符号化する（元のファイル名はフォルダ名にするため入れない）" {
        toIndexFileName "記号<>_1" | Should -Be "記号%3C%3E%5F1.tsv"
    }

    It "<name>" -TestCases @(
        @{ name = "場所が長くても255文字を超えなければ返す"; length = 251; message = "" }
        @{ name = "255文字を超えると、文字数の分かるメッセージで例外にする"; length = 252; message = "256 文字。上限 255 文字" }
    ) {
        param ($name, $length, $message)
        $place = "あ" * $length
        if ($message) {
            { toIndexFileName $place } | Should -Throw -ExpectedMessage "*$message*"
        } else {
            (toIndexFileName $place).Length | Should -Be 255
        }
    }

    It "<name>" -TestCases @(
        @{ name = "ページは page_<番号>"; place = "ページ001"; fileName = "page_001.tsv"; ascii = $true }
        @{ name = "桁数の多いページ番号もそのまま写す"; place = "ページ12345"; fileName = "page_12345.tsv"; ascii = $true }
        @{ name = "非表示のスライドは slide_<番号>_hidden"; place = "スライド002（非表示）"; fileName = "slide_002_hidden.tsv"; ascii = $true }
        @{ name = "スライドのノートは slide_<番号>_notes"; place = "スライド002_ノート"; fileName = "slide_002_notes.tsv"; ascii = $true }
        @{ name = "ふつうのスライドは slide_<番号>"; place = "スライド002"; fileName = "slide_002.tsv"; ascii = $true }
        @{ name = "スライドマスターは slide_master"; place = "スライドマスター"; fileName = "slide_master.tsv"; ascii = $true }
        @{ name = "ノートマスターは notes_master"; place = "ノートマスター"; fileName = "notes_master.tsv"; ascii = $true }
        @{ name = "ノートマスターの図形は末尾に [shape]"; place = "ノートマスター[図形]"; fileName = "notes_master[shape].tsv"; ascii = $true }
        @{ name = "ヘッダー・フッターは header_footer"; place = "ヘッダー・フッター"; fileName = "header_footer.tsv"; ascii = $true }
        @{ name = "脚注は doc_footnotes"; place = "脚注"; fileName = "doc_footnotes.tsv"; ascii = $true }
        @{ name = "文書は doc_whole"; place = "文書"; fileName = "doc_whole.tsv"; ascii = $true }
        @{ name = "本文は doc_body"; place = "本文"; fileName = "doc_body.tsv"; ascii = $true }
        @{ name = "固定名の場所の図形は末尾に [shape]"; place = "スライド002[図形]"; fileName = "slide_002[shape].tsv"; ascii = $true }
        @{ name = "固定名の場所のコメントは末尾に [comment]"; place = "ページ001[コメント]"; fileName = "page_001[comment].tsv"; ascii = $true }
        @{ name = "非表示のスライドの図形も末尾に [shape]"; place = "スライド002（非表示）[図形]"; fileName = "slide_002_hidden[shape].tsv"; ascii = $true }
        @{ name = "文書のコメントは doc_whole[comment]"; place = "文書[コメント]"; fileName = "doc_whole[comment].tsv"; ascii = $true }
        @{ name = "Excel の任意のシート名は符号化する（固定名に当てはまらない）"; place = "売上"; fileName = "売上.tsv"; ascii = $false }
        @{ name = "シート名の図形は末尾に [shape]"; place = "売上[図形]"; fileName = "売上[shape].tsv"; ascii = $false }
        @{ name = "シート名の _ は %5F にする（固定名と区別するため）"; place = "2024_上期"; fileName = "2024%5F上期.tsv"; ascii = $false }
        @{ name = "シート名が固定名のファイル名（page_001）と同じでも、_ を %5F にするため区別できる"; place = "page_001"; fileName = "page%5F001.tsv"; ascii = $true }
        @{ name = "シート名の % は %25 にする"; place = "50%引き"; fileName = "50%25引き.tsv"; ascii = $false }
        @{ name = "シート名のファイル名禁止文字は符号化する"; place = "記号<>"; fileName = "記号%3C%3E.tsv"; ascii = $false }
        @{ name = "Excel のシート名のヘッダー・フッターは末尾に [header_footer]"; place = "売上[ヘッダー・フッター]"; fileName = "売上[header_footer].tsv"; ascii = $false }
        @{ name = "固定名のヘッダー・フッター（Word・PowerPoint）に、同名のシートの種類の [header_footer] が付いても戻せる"; place = "ヘッダー・フッター[ヘッダー・フッター]"; fileName = "header_footer[header_footer].tsv"; ascii = $true }
    ) {
        param ($name, $place, $fileName, $ascii)
        toIndexFileName $place | Should -Be $fileName
        # convertIndexFileNameToPlace で元の場所へ一意に戻せる（往復できる）
        convertIndexFileNameToPlace ([System.IO.Path]::GetFileNameWithoutExtension($fileName)) | Should -Be $place
        if ($ascii) {
            $fileName | Should -Match "^[\x00-\x7F]+$"
        }
    }
}

Describe "placeKindFileNames" -Tag Unit {
    It "objectPlacePattern の種類の選択肢が、表（placeKindFileNames）のキーと集合として同じ（足し忘れを防ぐ）" {
        $marker = "(?<kind>"
        $start = ${objectPlacePattern}.IndexOf($marker) + $marker.Length
        $end = ${objectPlacePattern}.IndexOf(")|(?<numberedKind>")
        $alt = ${objectPlacePattern}.Substring($start, $end - $start)
        $patternKinds = $alt -split '\|' | ForEach-Object { [regex]::Unescape($_) }
        # 番号の付く種類（埋め込み）は選択肢の外に、番号つきで別に書く
        $numbered = [regex]::Match(${objectPlacePattern}, '\(\?<numberedKind>([^)]*)\)').Groups[1].Value
        (@($patternKinds) + [regex]::Unescape($numbered) | Sort-Object) | Should -Be (${placeKindFileNames}.Keys | Sort-Object)
    }

    It "表のキーだけが objectPlacePattern に一致し、知らない種類は一致しない" {
        foreach ($kind in ${placeKindFileNames}.Keys) {
            if ($kind -eq ${placeKindEmbed}) { continue }
            "売上[$kind]" | Should -Match ${objectPlacePattern}
        }
        "売上[未知の種類]" | Should -Not -Match ${objectPlacePattern}
    }

    It "埋め込みは番号つきだけが一致し、番号なし・0 始まりはふつうの場所" {
        "ページ001[埋め込み1]" | Should -Match ${objectPlacePattern}
        "ページ001[埋め込み12]" | Should -Match ${objectPlacePattern}
        "ページ001[埋め込み]" | Should -Not -Match ${objectPlacePattern}
        "ページ001[埋め込み01]" | Should -Not -Match ${objectPlacePattern}
        "ページ001[埋め込み0]" | Should -Not -Match ${objectPlacePattern}
    }

    It "埋め込みの場所は分割・ファイル名・表示で番号が往復する" {
        $s = splitObjectPlace "ページ003[埋め込み2]"
        $s.Base | Should -Be "ページ003"
        $s.Kind | Should -Be "埋め込み"
        $s.Number | Should -Be 2
        toIndexFileName "ページ003[埋め込み2]" | Should -Be "page_003[embed2].tsv"
        convertIndexFileNameToPlace "page_003[embed2]" | Should -Be "ページ003[埋め込み2]"
        (describePlace "a.docx" "ページ003[埋め込み2]").Kind | Should -Be "埋め込み 2"
        # 番号の付き方が違うファイル名は、種類として戻さない
        convertIndexFileNameToPlace "ページ003[embed]" | Should -Not -Match "埋め込み"
        convertIndexFileNameToPlace "売上[shape1]" | Should -Not -Match "図形"
    }
}

Describe "newIndexName / assignIndexNames" -Tag Unit {
    It "フォルダ名（ドライブ直下はドライブ名、UNC は共有名）を使い、重複すれば (2) を付ける" {
        $used = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        newIndexName "C:\data\見積" $used | Should -Be "見積"
        newIndexName "D:\" $used | Should -Be "D"
        newIndexName "\\server\share" $used | Should -Be "share"
        [void]$used.Add("見積")
        [void]$used.Add("見積(2)")
        newIndexName "E:\見積" $used | Should -Be "見積(3)"
    }

    It '使用中の名前が $null・配列・空でも落ちずに名前を作る' {
        # インデックスが 1 つも無いとき、呼び出し側から $null が渡ることがある（画面の追加ダイアログ）
        newIndexName "C:\data\sample" $null | Should -Be "sample"
        newIndexName "C:\data\見積" | Should -Be "見積"
        newIndexName "C:\data\見積" @() | Should -Be "見積"
        newIndexName "C:\data\見積" @("見積") | Should -Be "見積(2)"
        # インデックスが 1 件のときは、集合が展開されて文字列 1 個で渡ることがある
        newIndexName "C:\data\見積" "見積" | Should -Be "見積(2)"
        newIndexName "C:\data\見" "見積" | Should -Be "見"
        # 前方一致・大文字小文字違いで取り違えない
        newIndexName "C:\data\見積" @("見積書") | Should -Be "見積"
        newIndexName "C:\data\sample" @("SAMPLE") | Should -Be "sample(2)"
    }

    It "フォルダ名が取れなければ「フォルダ」とする" {
        newIndexName "" | Should -Be "フォルダ"
        newIndexName "" @("フォルダ") | Should -Be "フォルダ(2)"
    }

    It "重複して (2) を付けても、ファイル名の上限（255 文字）を超えない" {
        $long = "あ" * 255
        $name = newIndexName "C:\$long" @($long)
        $name.Length | Should -BeLessThan 256
        $name | Should -Match "\(2\)$"
        testIndexName $name @($long) | Should -Be ""
        # 切り詰めた名前どうしも重複させない
        $third = newIndexName "C:\$long" @($long, $name)
        $third | Should -Match "\(3\)$"
        $third.Length | Should -BeLessThan 256
    }

    It "設定に名前があればそれを使い、フォルダの場所が変わっても同じ名前のままにする" {
        $targets = @(
            [pscustomobject]@{ Name = "見積"; Path = "\server\新しい場所\見積書"; Enabled = $true },
            [pscustomobject]@{ Name = ""; Path = "F:\売上"; Enabled = $true }
        )
        # 前回は別の場所だったが、名前が同じなので同じインデックスとして扱う
        $previous = @([pscustomobject]@{ Path = "C:\data\見積"; Name = "見積" })
        $folders = @(assignIndexNames $targets $previous)
        $folders.Count | Should -Be 2
        $folders[0].Name | Should -Be "見積"
        $folders[0].Path | Should -Be "\server\新しい場所\見積書"
        $folders[1].Name | Should -Be "売上"
    }

    It "名前が無ければ、前回の取り込み一覧の同じフォルダの名前を使い、無ければフォルダ名から作る" {
        $targets = @(
            [pscustomobject]@{ Name = ""; Path = "C:\data\見積"; Enabled = $false },
            [pscustomobject]@{ Name = ""; Path = "E:\new\見積"; Enabled = $true },
            [pscustomobject]@{ Name = ""; Path = "F:\売上"; Enabled = $true }
        )
        $previous = @(
            [pscustomobject]@{ Path = "c:\data\見積"; Name = "見積" },
            [pscustomobject]@{ Path = "G:\削除した\報告"; Name = "報告" }
        )
        $folders = @(assignIndexNames $targets $previous)
        $folders[0].Name | Should -Be "見積"
        $folders[0].Enabled | Should -Be $false
        # 前回の名前（削除したフォルダの名前を含む）と重複しない名前を付ける
        $folders[1].Name | Should -Be "見積(2)"
        $folders[2].Name | Should -Be "売上"
    }

    It "設定にある名前は、ほかのフォルダの名前には使わない" {
        $targets = @(
            [pscustomobject]@{ Name = ""; Path = "E:\新\見積"; Enabled = $true },
            [pscustomobject]@{ Name = "見積"; Path = "C:\data\見積"; Enabled = $true }
        )
        $folders = @(assignIndexNames $targets @())
        $folders[0].Name | Should -Be "見積(2)"
        $folders[1].Name | Should -Be "見積"
    }

    It "パスにつなげない名前（.. ・末尾の .）は名前が無いものとして割り当て直す" {
        $targets = @(
            [pscustomobject]@{ Name = ".."; Path = "C:\data\見積"; Enabled = $true },
            [pscustomobject]@{ Name = "営業."; Path = "C:\data\営業"; Enabled = $true },
            [pscustomobject]@{ Name = "a\b"; Path = "C:\data\売上"; Enabled = $true },
            [pscustomobject]@{ Name = "経理"; Path = "C:\data\経理"; Enabled = $true }
        )
        $folders = @(assignIndexNames $targets @())
        $folders[0].Name | Should -Be "見積"
        $folders[1].Name | Should -Be "営業"
        $folders[2].Name | Should -Be "売上"
        $folders[3].Name | Should -Be "経理"
    }

    It "既にある予約名（<name>）は、そのまま使い続け、ほかのフォルダの名前には使わない" -TestCases @(
        @{ name = "aux.backup" }
        @{ name = "CON.txt" }
        @{ name = "CON" }
    ) {
        param ($name)
        $targets = @(
            [pscustomobject]@{ Name = $name; Path = "C:\data\売上"; Enabled = $true },
            [pscustomobject]@{ Name = ""; Path = "C:\data\$name"; Enabled = $true }
        )
        $folders = @(assignIndexNames $targets @())
        $folders[0].Name | Should -Be $name
        # 新しく付ける名前は予約名を避ける（newIndexName）ので、同じ名前にはならない
        $folders[1].Name | Should -Not -Be $name
        (testIndexName $folders[1].Name @()) | Should -Be ""
    }
}

Describe "testStoredIndexName" -Tag Unit {
    It "保存してある名前として使える: <name>" -TestCases @(
        @{ name = "営業部 2025" }
        @{ name = "aux.backup" }
        @{ name = "CON.txt" }
        @{ name = "CON" }
        @{ name = "LPT1" }
        @{ name = ("あ" * 255) }
    ) {
        param ($name)
        testStoredIndexName $name | Should -Be ""
    }

    It "置き場の外へ出る・作れない名前は拒む: <name>" -TestCases @(
        @{ name = "" }
        @{ name = ".." }
        @{ name = "." }
        @{ name = "営業." }
        @{ name = " 営業" }
        @{ name = "a\b" }
        @{ name = "C:\x" }
        @{ name = "a/b" }
        @{ name = ("あ" * 256) }
    ) {
        param ($name)
        testStoredIndexName $name | Should -Not -BeNullOrEmpty
    }

    It "新しく付ける名前（testIndexName・newIndexName）は、予約名を拒み続ける: <name>" -TestCases @(
        @{ name = "aux.backup" }
        @{ name = "CON.txt" }
        @{ name = "CON" }
    ) {
        param ($name)
        testIndexName $name @() | Should -Not -BeNullOrEmpty
        $made = newIndexName "C:\data\$name" @()
        testIndexName $made @() | Should -Be ""
        $made | Should -Not -Be $name
    }
}

Describe "splitIndexRelPath" -Tag Unit {
    It "先頭のインデックス名と残りに分ける（残りの \ や 2 はそのまま）" {
        $parts = splitIndexRelPath "excel\2024\見積\A社2.xlsx"
        $parts.Name | Should -Be "excel"
        $parts.Rest | Should -Be "2024\見積\A社2.xlsx"
        $parts = splitIndexRelPath "excel"
        $parts.Name | Should -Be "excel"
        $parts.Rest | Should -Be ""
    }
}

Describe "testIndexName" -Tag Unit {
    # 使える名前なら空文字列を返す。used はほかのインデックスの名前
    It "<name>" -TestCases @(
        @{ name = "使える名前なら空文字列を返す"; value = "営業部 2025" }
        @{ name = "ちょうど上限の長さなら使える"; value = ("あ" * 255) }
        @{ name = "予約語を含むだけの名前は使える（CONSOLE）"; value = "CONSOLE" }
        @{ name = "予約語を含むだけの名前は使える（営業NUL）"; value = "営業NUL" }
        @{ name = "ほかのインデックスの名前が無い（`$null）場合も使える"; value = "営業"; used = $null }
    ) {
        param ($name, $value, $used)
        testIndexName $value $used | Should -Be ""
    }

    # 使えない名前なら、理由の分かる文言を返す
    It "<name>" -TestCases @(
        @{ name = "空なら入力を促す"; value = ""; pattern = "入力してください" }
        @{ name = "前後に空白があれば使えない（前）"; value = " 営業"; pattern = "前後に空白" }
        @{ name = "前後に空白があれば使えない（後）"; value = "営業 "; pattern = "前後に空白" }
        @{ name = "ファイル名に使えない文字があれば使えない（\）"; value = "営業\部"; pattern = "使えない文字" }
        @{ name = "ファイル名に使えない文字があれば使えない（:）"; value = "営業:部"; pattern = "使えない文字" }
        @{ name = "末尾が . なら使えない"; value = "営業."; pattern = "最後に \." }
        @{ name = "Windows の予約語は使えない（大文字・小文字を区別しない。con）"; value = "con"; pattern = "使えない名前" }
        @{ name = "Windows の予約語は使えない（大文字・小文字を区別しない。LPT1）"; value = "LPT1"; pattern = "使えない名前" }
        @{ name = "ほかのインデックスと同じ名前は使えない（大文字・小文字を区別しない）"; value = "Sales"; used = @("sales", "tech"); pattern = "ほかのインデックスが使っています" }
        @{ name = "長すぎる名前は使えない"; value = ("あ" * 256); pattern = "長すぎます" }
        @{ name = "拡張子の付いた予約語・. だけの名前・制御文字も使えない（CON.txt）"; value = "CON.txt"; pattern = "使えない名前" }
        @{ name = "拡張子の付いた予約語・. だけの名前・制御文字も使えない（com9.backup）"; value = "com9.backup"; pattern = "使えない名前" }
        @{ name = "拡張子の付いた予約語・. だけの名前・制御文字も使えない（.）"; value = "."; pattern = "最後に \." }
        @{ name = "拡張子の付いた予約語・. だけの名前・制御文字も使えない（タブ）"; value = "営業`t部"; pattern = "使えない文字" }
    ) {
        param ($name, $value, $used, $pattern)
        testIndexName $value $used | Should -Match $pattern
    }

    It "画面に出す理由は「インデックス名」でなく「名前」で言う（一覧の見出し・編集の欄と同じ言葉）" -TestCases @(
        @{ value = "" }, @{ value = " 営業" }, @{ value = "営業\部" }, @{ value = "営業." }, @{ value = ("あ" * 256) }
    ) {
        param ($value)
        $reason = testIndexName $value @()
        $reason | Should -Match "^名前"
        $reason | Should -Not -Match "インデックス名"
    }
}

Describe "testIndexRelPath" -Tag Unit {
    It "使える相対パスなら空文字列を返す: <path>" -TestCases @(
        @{ path = "営業\a.xlsx" }
        @{ path = "営業\2024\a.xlsx" }
        @{ path = "営業\a b.xlsx" }
        @{ path = "営業\末尾が点." }
        @{ path = "営業\CON" }
        @{ path = "CON\x" }
        @{ path = "aux.backup\a.xlsx" }
    ) {
        param ($path)
        testIndexRelPath $path | Should -Be ""
    }

    It "外へ出る・使えない相対パスは理由を返す: <path>" -TestCases @(
        @{ path = "..\x" }
        @{ path = "..\..\..\x" }
        @{ path = "営業\..\..\x" }
        @{ path = "営業\.\x" }
        @{ path = "営業\\x" }
        @{ path = "営業\" }
        @{ path = "C:\x" }
        @{ path = "\\srv\x" }
        @{ path = "営業\a:b" }
        @{ path = "営業/a" }
        @{ path = "営業" }
        @{ path = "" }
    ) {
        param ($path)
        testIndexRelPath $path | Should -Not -BeNullOrEmpty
    }

    It "allowNameOnly なら、残りの無い名前だけも通し、外へ出る名前は通さない" {
        testIndexRelPath "営業" $true | Should -Be ""
        testIndexRelPath ".." $true | Should -Not -BeNullOrEmpty
        testIndexRelPath "営業\..\x" $true | Should -Not -BeNullOrEmpty
    }

    It "checkedNames に通った名前を足し、あるものは調べ直さない" {
        $checked = New-Object 'System.Collections.Generic.HashSet[string]'
        testIndexRelPath "営業\a.xlsx" $false $checked | Should -Be ""
        $checked.Contains("営業") | Should -BeTrue
        testIndexRelPath "..\a.xlsx" $false $checked | Should -Not -BeNullOrEmpty
        $checked.Contains("..") | Should -BeFalse
    }
}

Describe "testPathSegments" -Tag Unit {
    It "厳しい確かめは、最後の . や空白・予約語も拒む: <segments>" -TestCases @(
        @{ segments = @("a", "b.") }
        @{ segments = @("a ", "b") }
        @{ segments = @("a", "NUL.txt") }
        @{ segments = @("a", "..") }
        @{ segments = @("a", "") }
    ) {
        param ($segments)
        testPathSegments $segments | Should -Not -BeNullOrEmpty
    }

    It "ゆるい確かめ（lenient）は、最後の . や空白・予約語は通し、外へ出る区切りは拒む" {
        testPathSegments @("a", "b.", "c ", "NUL") $true | Should -Be ""
        testPathSegments @("a", "..") $true | Should -Not -BeNullOrEmpty
        testPathSegments @("a", "b:c") $true | Should -Not -BeNullOrEmpty
        testPathSegments @("あ" * 256) $true | Should -Not -BeNullOrEmpty
    }
}

Describe "newIndexName が作る名前は testIndexName を通る" -Tag Unit {
    It "<folder> の名前は <expected>" -TestCases @(
        @{ folder = "C:\data\　営業　"; used = @(); expected = "営業" }
        @{ folder = "C:\data\ 営業 "; used = @(); expected = "営業" }
        @{ folder = "C:\data\営業."; used = @(); expected = "営業" }
        @{ folder = "C:\data\営業 . "; used = @(); expected = "営業" }
        @{ folder = "C:\data\　"; used = @(); expected = "フォルダ" }
        @{ folder = "C:\data\..."; used = @(); expected = "フォルダ" }
        @{ folder = "C:\data\CON"; used = @(); expected = "フォルダ" }
        @{ folder = "C:\data\nul.txt"; used = @(); expected = "フォルダ" }
        @{ folder = "C:\data\　営業　"; used = @("営業"); expected = "営業(2)" }
        @{ folder = "C:\data\　"; used = @("フォルダ"); expected = "フォルダ(2)" }
        @{ folder = ("C:\data\" + ("あ" * 300)); used = @(); expected = ("あ" * 255) }
        @{ folder = ("C:\data\" + ("あ" * 300)); used = @(("あ" * 255)); expected = (("あ" * 252) + "(2)") }
        # 255 文字に切ったあとの末尾が . や空白になる（削って 254 文字になる）
        @{ folder = ("C:\data\" + ("あ" * 254) + ".いう"); used = @(); expected = ("あ" * 254) }
        @{ folder = ("C:\data\" + ("あ" * 254) + " いう"); used = @(); expected = ("あ" * 254) }
        @{ folder = ("C:\data\" + ("あ" * 254) + "　いう"); used = @(); expected = ("あ" * 254) }
        @{ folder = ("C:\data\" + ("あ" * 253) + " .いう"); used = @(); expected = ("あ" * 253) }
        # 切ったあとに (2) を付けて空きを作る（元の名前の途中が . や空白でも使える名前になる）
        @{ folder = ("C:\data\" + ("あ" * 254) + ".いう"); used = @(("あ" * 254)); expected = (("あ" * 252) + "(2)") }
        @{ folder = ("C:\data\" + ("あ" * 251) + ".いうえ"); used = @(("あ" * 251) + ".いうえ"); expected = (("あ" * 251) + ".(2)") }
        @{ folder = ("C:\data\" + ("あ" * 251) + " いうえ"); used = @(("あ" * 251) + " いうえ"); expected = (("あ" * 251) + " (2)") }
    ) {
        param ($folder, $used, $expected)
        $name = newIndexName $folder $used
        $name | Should -Be $expected
        testIndexName $name @() | Should -Be ""
    }
}

Describe "newIndexName（入れ子の集合）" -Tag Unit {
    It "集合が 1 要素の配列に入って渡されても、使用済みの名前を見落とさない" {
        $set = New-Object 'System.Collections.Generic.HashSet[string]'
        [void]$set.Add("営業"); [void]$set.Add("経理")
        newIndexName "C:\data\営業" @(, $set) | Should -Be "営業(2)"
    }
}

Describe "testSourceBookName" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "Office の拡張子"; book = "a.xlsx"; expected = $true }
        @{ name = "大文字の拡張子"; book = "a.TXT"; expected = $true }
        @{ name = "PDF"; book = "a.pdf"; expected = $true }
        @{ name = "末尾のピリオドは Windows が無視するので、拡張子は手前で見る"; book = "a.xlsx."; expected = $true }
        @{ name = "末尾の空白も同じ"; book = "a.txt "; expected = $true }
        @{ name = "日本語の名前"; book = "見積（最終）.docx"; expected = $true }
        @{ name = "取り込み対象でない拡張子（.hta）"; book = "a.hta"; expected = $false }
        @{ name = "ショートカット（.lnk）"; book = "a.lnk"; expected = $false }
        @{ name = "インターネットショートカット（.url）"; book = "a.url"; expected = $false }
        @{ name = "実行ファイル（.exe）"; book = "a.exe"; expected = $false }
        @{ name = "マクロ入りテンプレート（.dotm）は対象外"; book = "a.dotm"; expected = $false }
        @{ name = "上の階層を指す名前"; book = '..\a.xlsx'; expected = $false }
        @{ name = "区切り（\）を含む名前"; book = 'sub\a.xlsx'; expected = $false }
        @{ name = "区切り（/）を含む名前"; book = "a/b.txt"; expected = $false }
        @{ name = "代替データストリームの指定"; book = "a.xlsm::`$DATA"; expected = $false }
        @{ name = "表示の向きを変える文字（RLO）を含む名前"; book = "a" + [string][char]0x202E + "txt.exe"; expected = $false }
        @{ name = "制御文字（U+0001）を含む名前"; book = "a" + [string][char]1 + ".txt"; expected = $false }
        @{ name = "C1 制御文字（U+0085）を含む名前"; book = "a" + [string][char]0x85 + ".txt"; expected = $false }
        @{ name = "ピリオドだけ"; book = "."; expected = $false }
        @{ name = "ピリオド 2 つ"; book = ".."; expected = $false }
        @{ name = "拡張子が無い"; book = "readme"; expected = $false }
        @{ name = "空"; book = ""; expected = $false }
    ) {
        param ($name, $book, $expected)
        testSourceBookName $book | Should -Be $expected
    }
}

Describe "assignIndexNames（Subfolders）" -Tag Unit {
    It "Subfolders（false・true・無し）を、名前を割り当てたあとの項目にも引き継ぐ" {
        $targets = @(
            [pscustomobject]@{ Name = ""; Path = "C:\data\見積"; Enabled = $true; Subfolders = $false },
            [pscustomobject]@{ Name = "営業"; Path = "C:\data\営業"; Enabled = $true; Subfolders = $true },
            [pscustomobject]@{ Name = ""; Path = "C:\data\経理"; Enabled = $true }
        )
        $folders = @(assignIndexNames $targets @())
        $folders[0].Subfolders | Should -Be $false
        $folders[1].Subfolders | Should -Be $true
        $folders[2].Subfolders | Should -BeNullOrEmpty
    }
}
