# PDF ファイルの読み取り（shared\core\pdf_file.ps1）のテスト。PDF はテストの中で組み立てる（tests\helpers\pdf.ps1）
BeforeDiscovery {
    $surrogate = [string][char]0xD840 + [string][char]0xDC0B

    # 文字の命令だけが違う表（フォントは Helvetica・WinAnsi）。圧縮なし・Flate の両方で確かめる
    $textCases = @(
        @{ name = "Tj"; pages = @("BT /F1 12 Tf 72 700 Td (abc) Tj ET"); expected = @("abc") }
        @{ name = "Td で行が変わる"; pages = @("BT /F1 12 Tf (a) Tj 0 -14 Td (b) Tj ET"); expected = @("a`nb") }
        @{ name = "T* と ' で行が変わる"; pages = @("BT /F1 12 Tf (a) Tj T* (b) Tj (c) ' ET"); expected = @("a`nb`nc") }
        @{ name = "`" は 3 つの引数のあとの文字を出す"; pages = @("BT /F1 12 Tf (a) Tj 1 2 (b) `" ET"); expected = @("a`nb") }
        @{ name = "同じ行の Tj は続けて出す"; pages = @("BT /F1 12 Tf (ab) Tj (cd) Tj ET"); expected = @("abcd") }
        @{ name = "TJ の大きな詰めは半角の空白"; pages = @("BT /F1 12 Tf [(Foo) -300 (Bar)] TJ ET"); expected = @("Foo Bar") }
        @{ name = "TJ の閾値ちょうど（-250）も空白"; pages = @("BT /F1 12 Tf [(Foo) -250 (Bar)] TJ ET"); expected = @("Foo Bar") }
        @{ name = "TJ の小さな詰めは空白なし"; pages = @("BT /F1 12 Tf [(Fo) -100 (o)] TJ ET"); expected = @("Foo") }
        @{ name = "TJ の正の数（左への詰め）は空白なし"; pages = @("BT /F1 12 Tf [(Fo) 300 (o)] TJ ET"); expected = @("Foo") }
        @{ name = "エスケープと 8 進"; pages = @("BT /F1 12 Tf (a\(b\)\\c\101) Tj ET"); expected = @("a(b)\cA") }
        @{ name = "16 進の文字列"; pages = @("BT /F1 12 Tf <414243> Tj ET"); expected = @("ABC") }
        @{ name = "WinAnsi の é"; pages = @("BT /F1 12 Tf (caf\351) Tj ET"); expected = @("café") }
        @{ name = "コメントは読まない"; pages = @("BT /F1 12 Tf % (no) Tj`n(yes) Tj ET"); expected = @("yes") }
        @{ name = "埋め込み画像の中身は読まない"; pages = @("BI /W 1 /H 1 /BPC 8 ID (x) Tj EI BT /F1 12 Tf (ok) Tj ET"); expected = @("ok") }
        @{ name = "文字の無いページは空"; pages = @("0 0 m 10 10 l S"); expected = @("") }
        @{ name = "複数ページ"; pages = @("BT /F1 12 Tf (one) Tj ET", "BT /F1 12 Tf (two) Tj ET", ""); expected = @("one", "two", "") }
        # 行の分け方。ET では分けず、縦位置が違うときだけ分ける（Word・Chromium は 1 行を複数の BT〜ET に分ける）
        @{ name = "BT を分けても同じ縦位置なら同じ行"; pages = @("BT /F1 12 Tf 1 0 0 1 72 700 Tm (Hello) Tj ET BT /F1 12 Tf 1 0 0 1 120 700 Tm (world) Tj ET"); expected = @("Hello world") }
        @{ name = "BT を分けて縦位置が違えば別の行"; pages = @("BT /F1 12 Tf 1 0 0 1 72 700 Tm (a) Tj ET BT /F1 12 Tf 1 0 0 1 72 686 Tm (b) Tj ET"); expected = @("a`nb") }
        @{ name = "位置を決めない BT〜ET の続きは同じ行"; pages = @("BT /F1 12 Tf (a) Tj ET BT /F1 12 Tf (b) Tj ET"); expected = @("ab") }
        @{ name = "Tm で縦位置が変われば別の行"; pages = @("BT /F1 12 Tf 1 0 0 1 72 700 Tm (a) Tj 1 0 0 1 72 686 Tm (b) Tj ET"); expected = @("a`nb") }
        @{ name = "Tm が同じ縦位置なら同じ行"; pages = @("BT /F1 12 Tf 1 0 0 1 72 700 Tm (a) Tj 1 0 0 1 100 700 Tm (b) Tj ET"); expected = @("a b") }
        @{ name = "TD で行が変わり、行送りを覚えて T* で送る"; pages = @("BT /F1 12 Tf 0 -14 TD (a) Tj 0 -14 TD (b) Tj T* (c) Tj ET"); expected = @("a`nb`nc") }
        @{ name = "TL の行送りで T* と ' が送る"; pages = @("BT /F1 12 Tf 14 TL (a) Tj T* (b) Tj (c) ' ET"); expected = @("a`nb`nc") }
        @{ name = "横だけの Td は同じ行（半角どうしなら空白）"; pages = @("BT /F1 12 Tf (a) Tj 50 0 Td (b) Tj ET"); expected = @("a b") }
        @{ name = "1 em より小さい Td は詰めにしない（1 文字ずつ並べる PDF）"; pages = @("BT /F1 16 Tf 8 94 Td (E) Tj 9 0 Td (d) Tj 9.9 0 Td (g) Tj ET"); expected = @("Edg") }
        @{ name = "1 em 以上の Td は空白でつなぐ"; pages = @("BT /F1 16 Tf (a) Tj 16 0 Td (b) Tj ET"); expected = @("a b") }
        @{ name = "後ろ向きの Td は空白でつなぐ"; pages = @("BT /F1 16 Tf (a) Tj -3 0 Td (b) Tj ET"); expected = @("a b") }
        @{ name = "Tf の大きさが 0 のときは Td を空白でつなぐ"; pages = @("BT /F1 0 Tf (a) Tj 5 0 Td (b) Tj ET"); expected = @("a b") }
        @{ name = "2 文字以上の語の後ろは、1 em より小さい Td でも空白でつなぐ"; pages = @("BT /F1 16 Tf (ab) Tj 5 0 Td (cd) Tj ET"); expected = @("ab cd") }
        @{ name = "Tm で拡大した中でも、Td は Tf の大きさと比べる（小さい）"; pages = @("BT 12 0 0 12 0 0 Tm /F1 1 Tf (a) Tj 0.5 0 Td (b) Tj ET"); expected = @("ab") }
        @{ name = "Tm で拡大した中でも、Td は Tf の大きさと比べる（1 em 以上）"; pages = @("BT 12 0 0 12 0 0 Tm /F1 1 Tf (a) Tj 1.5 0 Td (b) Tj ET"); expected = @("a b") }
        @{ name = "途中で Tf を挟むと、新しい大きさで比べる"; pages = @("BT /F1 16 Tf (a) Tj /F1 4 Tf 5 0 Td (b) Tj ET"); expected = @("a b") }
        @{ name = "Tm で拡大した後の TD の縦の動きは拡大を掛ける"; pages = @("BT 12 0 0 12 72 700 Tm /F1 1 Tf (a) Tj 0 -1 TD (b) Tj ET"); expected = @("a`nb") }
        @{ name = "縦のずれが小さい Td は同じ行（上付き文字など）"; pages = @("BT /F1 12 Tf (a) Tj 20 0.5 Td (b) Tj ET"); expected = @("a b") }
        @{ name = "回した Tm（d が 0）の後でも、Td の縦の動きで行が分かれる"; pages = @("BT 0 1 -1 0 72 700 Tm /F1 12 Tf (a) Tj 0 -14 Td (b) Tj ET"); expected = @("a`nb") }
        @{ name = "上下を反転した Tm（d が負）の後の Td は、縦の動きも反転する"; pages = @("BT 1 0 0 -1 0 700 Tm /F1 12 Tf (a) Tj 0 -14 Td (b) Tj 1 0 0 1 0 714 Tm (c) Tj ET"); expected = @("a`nb c") }
        @{ name = "BT で縦位置が 0 に戻る"; pages = @("BT /F1 12 Tf 0 -14 Td (a) Tj ET BT /F1 12 Tf 0 -14 Td (b) Tj ET"); expected = @("ab") }
        @{ name = "16 進の空白と奇数桁"; pages = @("BT /F1 12 Tf <41 42 4> Tj ET"); expected = @("AB@") }
    )

    # Type0 フォント（Identity-H）＋ ToUnicode の表
    $cidCases = @(
        @{ name = "bfchar と bfrange"; content = "BT /F1 12 Tf <0001001000110012> Tj ET"; expected = "あいぅう" }
        @{ name = "サロゲートペア"; content = "BT /F1 12 Tf <0002> Tj ET"; expected = $surrogate }
        @{ name = "TJ の詰めで前後が半角なら空白"; content = "BT /F1 12 Tf [<0020> -300 <0021>] TJ ET"; expected = "A B" }
        @{ name = "TJ の詰めで前後が全角なら空白なし"; content = "BT /F1 12 Tf [<0001> -300 <0001>] TJ ET"; expected = "ああ" }
        @{ name = "TJ の詰めで片方が全角なら空白なし"; content = "BT /F1 12 Tf [<0020> -300 <0001>] TJ ET"; expected = "Aあ" }
        @{ name = "配列の bfrange"; content = "BT /F1 12 Tf <0030> Tj <0031> Tj ET"; expected = "XY" }
        @{ name = "1 コードが 2 文字になる合字は 1 字形と数え、小さい Td でも語をつなぐ"; content = "BT /F1 12 Tf <0040> Tj 3 0 Td <0021> Tj ET"; expected = "fiB" }
    )

    # 読めない PDF。kind は BeforeAll の newBrokenPdf が組み立てる
    $brokenCases = @(
        @{ kind = "empty"; message = "PDFではないか、壊れているため読めません。" }
        @{ kind = "notPdf"; message = "PDFではないか、壊れているため読めません。" }
        @{ kind = "truncated"; message = "PDFの相互参照が見つからないため読めません。" }
        @{ kind = "encrypted"; message = "PDFの暗号化の情報が壊れているため読めません。" }
        @{ kind = "startxrefOutside"; message = "*ファイルの外*" }
        @{ kind = "entryOutside"; message = "PDFの相互参照がファイルの外を指しているため読めません。" }
        @{ kind = "kidsCycle"; message = "*循環*" }
        @{ kind = "refCycle"; message = "*循環*" }
        @{ kind = "garbageFlate"; message = "PDFが壊れているため読めません。" }
        @{ kind = "lzw"; message = "*未対応の圧縮*" }
        @{ kind = "cidNoToUnicode"; message = "*文字の対応表が無い*" }
        @{ kind = "inflatedOver"; message = "*解凍後*" }
        @{ kind = "nestedValue"; message = "*入れ子*" }
        @{ kind = "nestedPages"; message = "*入れ子*" }
        @{ kind = "tooManyObjects"; message = "*オブジェクトの数*" }
        @{ kind = "tooManyPages"; message = "*ページ数*" }
        @{ kind = "uncompressedOver"; message = "*解凍後*" }
        @{ kind = "uncompressedBig"; message = "*解凍後*" }
        @{ kind = "mapOverOne"; message = "*文字の対応表が大きすぎる*" }
        @{ kind = "mapOverTotal"; message = "*文字の対応表が大きすぎる*" }
        @{ kind = "lengthChain"; message = "*参照が深すぎる*" }
        @{ kind = "prevCycle"; message = "*相互参照が循環*" }
        @{ kind = "prevTooMany"; message = "*相互参照の数が上限*" }
        @{ kind = "objStreamCount"; message = "*件数が中身と合わない*" }
        @{ kind = "arrayTooMany"; message = "*配列の要素の数*" }
        @{ kind = "dictTooMany"; message = "*辞書の項目の数*" }
        @{ kind = "xrefRowsRepeat"; message = "*オブジェクトの数が上限*" }
        @{ kind = "compressedOver"; message = "*解凍後*" }
        @{ kind = "searchOver"; message = "*終わりを探す量*" }
        @{ kind = "objStreamHeader"; message = "*見出しが読めない*" }
    )
}

BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "$PSScriptRoot\..\..\helpers\pdf.ps1"

    function newCidFontObjects {
        # フォント F1（obj 3）を Type0 にする。toUnicode が $false なら ToUnicode（obj 5）を付けない。cmapText を渡すと、その表にする
        param ([bool]$toUnicode = $true, [string]$cmapText = "")

        $cmap = @"
/CIDInit /ProcSet findresource begin 12 dict begin begincmap
/CMapName /Adobe-Identity-UCS def /CMapType 2 def
1 begincodespacerange <0000> <FFFF> endcodespacerange
3 beginbfchar <0001> <3042> <0002> <D840DC0B> <0040> <00660069> endbfchar
2 beginbfrange <0010> <0012> <3044> <0020> <0021> <0041> endbfrange
1 beginbfrange <0030> <0031> [<0058> <0059>] endbfrange
endcmap end end
"@
        if ($cmapText.Length -gt 0) { $cmap = $cmapText }
        $to = if ($toUnicode) { " /ToUnicode 5 0 R" } else { "" }
        $objs = @{}
        $objs[4] = '<< /Type /Font /Subtype /CIDFontType2 /BaseFont /Test /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> >>'
        $objs[5] = newPdfStreamObject '' (toPdfBytes $cmap)
        return @{
            Font  = "<< /Type /Font /Subtype /Type0 /BaseFont /Test /Encoding /Identity-H /DescendantFonts [4 0 R]$to >>"
            Extra = $objs
        }
    }

    function newPdfString([byte[]]$bytes) {
        return [System.Text.Encoding]::GetEncoding(28591).GetString($bytes)
    }

    function getStartXref([string]$pdf) {
        return [int][regex]::Match($pdf, 'startxref\r?\n(\d+)\r?\n%%EOF\s*$').Groups[1].Value
    }

    function addXrefUpdates {
        # 内容を変えない増分更新（空の相互参照）を count 回足す。それぞれの Prev は 1 つ前
        param ([byte[]]$pdf, [int]$count)

        $s = newPdfString $pdf
        for ($i = 0; $i -lt $count; $i++) {
            $prevPos = getStartXref $s
            $pos = $s.Length
            $s += "xref`n0 1`n0000000000 65535 f `ntrailer`n<< /Size 12 /Root 1 0 R /Prev $prevPos >>`nstartxref`n$pos`n%%EOF`n"
        }
        return , (toPdfBytes $s)
    }

    function newMapPdf {
        # 2 つのフォント F1・F2（Type0。ToUnicode は obj 8・9）を持つ 1 ページの PDF。ToUnicode の表はどれも entries 個の項目
        param ([string]$content, [int]$entries, [bool]$sameMap = $false, [bool]$directFont = $false)

        $cmap = "1 begincodespacerange <0000> <FFFF> endcodespacerange 1 beginbfrange <0000> <{0:X4}> <0041> endbfrange" -f ($entries - 1)
        $fontTail = '/Subtype /Type0 /BaseFont /Test /Encoding /Identity-H /DescendantFonts [4 0 R] /ToUnicode'
        $second = if ($sameMap) { 8 } else { 9 }
        $objs = @{}
        $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
        $objs[2] = '<< /Type /Pages /Kids [3 0 R] /Count 1 >>'
        if ($directFont) {
            $objs[3] = "<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 << /Type /Font $fontTail 8 0 R >> >> >> /Contents 6 0 R >>"
        } else {
            $objs[3] = '<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 7 0 R /F2 10 0 R >> >> /Contents 6 0 R >>'
        }
        $objs[6] = newPdfStreamObject '' (toPdfBytes $content)
        $objs[7] = "<< /Type /Font $fontTail 8 0 R >>"
        $objs[10] = "<< /Type /Font $fontTail $second 0 R >>"
        $objs[8] = newPdfStreamObject '' (toPdfBytes $cmap)
        $objs[9] = newPdfStreamObject '' (toPdfBytes $cmap)
        return , (newPdfFromObjects $objs)
    }

    function newBrokenPdf {
        param ([string]$kind)

        $simple = "BT /F1 12 Tf (hello) Tj ET"
        $latin1 = [System.Text.Encoding]::GetEncoding(28591)
        switch ($kind) {
            "empty" { return , ([byte[]]@()) }
            "notPdf" { return , (toPdfBytes "this is not a pdf file") }
            "truncated" {
                $b = newSimplePdf -pages @($simple)
                return , ([byte[]]$b[0..([int]($b.Length * 0.6))])
            }
            "encrypted" {
                $objs = @{}
                $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
                $objs[2] = '<< /Type /Pages /Kids [] /Count 0 >>'
                $objs[3] = '<< /Filter /Standard /V 5 /R 6 /CF << /StdCF << /CFM /AESV3 >> >> /StmF /StdCF >>'
                return , (newPdfFromObjects $objs '/Root 1 0 R /Encrypt 3 0 R')
            }
            "startxrefOutside" {
                $s = $latin1.GetString((newSimplePdf -pages @($simple)))
                $i = $s.LastIndexOf("startxref")
                return , (toPdfBytes ($s.Substring(0, $i) + "startxref`n99999999`n%%EOF`n"))
            }
            "entryOutside" {
                $s = $latin1.GetString((newSimplePdf -pages @($simple)))
                return , (toPdfBytes $s.Replace("0000000009 00000 n", "0099999999 00000 n"))
            }
            "kidsCycle" {
                $objs = @{}
                $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
                $objs[2] = '<< /Type /Pages /Kids [2 0 R] /Count 1 >>'
                return , (newPdfFromObjects $objs)
            }
            "refCycle" {
                $objs = @{}
                $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
                $objs[2] = '<< /Type /Pages /Kids [3 0 R] /Count 1 >>'
                $objs[3] = '<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>'
                $objs[4] = '5 0 R'
                $objs[5] = '4 0 R'
                return , (newPdfFromObjects $objs)
            }
            "garbageFlate" {
                $objs = @{}
                $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
                $objs[2] = '<< /Type /Pages /Kids [3 0 R] /Count 1 >>'
                $objs[3] = '<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>'
                $objs[4] = newPdfStreamObject '/Filter /FlateDecode' ([byte[]](@(255) * 40))
                return , (newPdfFromObjects $objs)
            }
            "lzw" {
                $objs = @{}
                $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
                $objs[2] = '<< /Type /Pages /Kids [3 0 R] /Count 1 >>'
                $objs[3] = '<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>'
                $objs[4] = newPdfStreamObject '/Filter /LZWDecode' ([byte[]](1..40))
                return , (newPdfFromObjects $objs)
            }
            "cidNoToUnicode" {
                $cid = newCidFontObjects $false
                return , (newSimplePdf -pages @("BT /F1 12 Tf <0001> Tj ET") -fontObject $cid.Font -extraObjects $cid.Extra)
            }
            "inflatedOver" { return , (newSimplePdf -pages @("BT /F1 12 Tf (" + ("a" * 500) + ") Tj ET") -compress) }
            "nestedValue" {
                $objs = @{}
                $objs[1] = '<< /Type /Catalog /Pages 2 0 R /X [[[[[[[[[[1]]]]]]]]]] >>'
                $objs[2] = '<< /Type /Pages /Kids [] /Count 0 >>'
                return , (newPdfFromObjects $objs)
            }
            "nestedPages" {
                $objs = @{}
                $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
                $objs[2] = '<< /Type /Pages /Kids [3 0 R] /Count 1 >>'
                $objs[3] = '<< /Type /Pages /Parent 2 0 R /Kids [4 0 R] /Count 1 >>'
                $objs[4] = '<< /Type /Pages /Parent 3 0 R /Kids [5 0 R] /Count 1 >>'
                $objs[5] = '<< /Type /Pages /Parent 4 0 R /Kids [6 0 R] /Count 1 >>'
                $objs[6] = '<< /Type /Page /Parent 5 0 R >>'
                return , (newPdfFromObjects $objs)
            }
            "uncompressedOver" {
                # 圧縮なしの同じストリームを何度も参照しても、取り出した合計を数える
                $objs = @{}
                $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
                $objs[2] = '<< /Type /Pages /Kids [3 0 R] /Count 1 >>'
                $objs[3] = '<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /Contents [4 0 R 4 0 R 4 0 R] >>'
                $objs[4] = newPdfStreamObject '' (toPdfBytes ("BT /F1 12 Tf (a) Tj ET" + (" " * 20)))
                $objs[5] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
                return , (newPdfFromObjects $objs)
            }
            "uncompressedBig" { return , (newSimplePdf -pages @("BT /F1 12 Tf (" + ("a" * 200) + ") Tj ET")) }
            "mapOverOne" { return , (newMapPdf "BT /F1 12 Tf <0001> Tj ET" 5000) }
            "mapOverTotal" { return , (newMapPdf "BT /F1 12 Tf <0001> Tj /F2 12 Tf <0001> Tj ET" 600) }
            "lengthChain" {
                # /Length が別のストリームを指し、その /Length がまた別のストリームを指す…
                $objs = @{}
                $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
                $objs[2] = '<< /Type /Pages /Kids [3 0 R] /Count 1 >>'
                $objs[3] = '<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>'
                for ($i = 4; $i -le 14; $i++) {
                    $objs[$i] = @{ Dict = "<< /Length $($i + 1) 0 R >>"; Stream = (toPdfBytes "x") }
                }
                $objs[15] = '1'
                return , (newPdfFromObjects $objs)
            }
            "prevCycle" {
                $s = newPdfString (newSimplePdf -pages @($simple))
                $pos = getStartXref $s
                return , (toPdfBytes ($s -replace '/Root 1 0 R >>', "/Root 1 0 R /Prev $pos >>"))
            }
            "prevTooMany" { return , (addXrefUpdates (newSimplePdf -pages @($simple)) 4) }
            "objStreamCount" {
                # オブジェクトのストリームの /N だけを、中身に見合わない大きな値（同じ桁数）に書き換える
                $p = @{}
                $p[1] = '<< /Type /Catalog /Pages 2 0 R >>'
                $p[2] = '<< /Type /Pages /Kids [] /Count 0 >>'
                foreach ($i in 3..10) { $p[$i] = '1' }
                $s = newPdfString (newPdfWithXrefStream -normal @{} -packed $p)
                if (-not $s.Contains("/N 10 ")) { throw "test setup: /N not found" }
                return , (toPdfBytes ($s.Replace("/N 10 ", "/N 99 ")))
            }
            "arrayTooMany" {
                $objs = @{}
                $objs[1] = '<< /Type /Catalog /Pages 2 0 R /X [1 2 3 4 5 6 7 8] >>'
                $objs[2] = '<< /Type /Pages /Kids [] /Count 0 >>'
                return , (newPdfFromObjects $objs)
            }
            "dictTooMany" {
                $objs = @{}
                $objs[1] = '<< /Type /Catalog /Pages 2 0 R /A 1 /B 2 /C 3 /D 4 /E 5 /F 6 >>'
                $objs[2] = '<< /Type /Pages /Kids [] /Count 0 >>'
                return , (newPdfFromObjects $objs)
            }
            "xrefRowsRepeat" {
                # 同じ番号の行を繰り返す相互参照の表（番号の数は増えないが、行の総数は数える）
                $objs = @{}
                $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
                $objs[2] = '<< /Type /Pages /Kids [] /Count 0 >>'
                $s = newPdfString (newPdfFromObjects $objs)
                $m = [regex]::Match($s, "xref\n(0 3\n(?:\d{10} \d{5} [nf] \n){3})")
                if (-not $m.Success) { throw "test setup: xref not found" }
                $section = $m.Groups[1].Value
                return , (toPdfBytes ($s.Replace($section, $section * 3)))
            }
            "compressedOver" {
                # 解凍するとごく小さい圧縮のストリームも、取り出す前の大きさを数える（解凍後だけなら 66、圧縮の分を足すと上限を超える）
                $objs = @{}
                $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
                $objs[2] = '<< /Type /Pages /Kids [3 0 R] /Count 1 >>'
                $objs[3] = '<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /Contents [4 0 R 4 0 R 4 0 R] >>'
                $objs[4] = newPdfStreamObject '/Filter /FlateDecode' (compressPdfBytes (toPdfBytes "BT /F1 12 Tf (a) Tj ET"))
                $objs[5] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
                return , (newPdfFromObjects $objs)
            }
            "searchOver" {
                # /Length が合わないストリームを /Kids に並べる（ページにならないので、取り出す合計には入らない。終わりを探す量だけで止める）
                $objs = @{}
                $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
                $objs[2] = '<< /Type /Pages /Kids [3 0 R 4 0 R 5 0 R 6 0 R] /Count 4 >>'
                foreach ($i in 3..6) {
                    $objs[$i] = @{ Dict = '<< /Length 1 >>'; Stream = (toPdfBytes ("x" * 200)) }
                }
                return , (newPdfFromObjects $objs)
            }
            "objStreamHeader" {
                $p = @{}
                $p[1] = '<< /Type /Catalog /Pages 2 0 R >>'
                $p[2] = '<< /Type /Pages /Kids [] /Count 0 >>'
                return , (newPdfWithXrefStream -normal @{} -packed $p -badHeader)
            }
            "tooManyObjects" { return , (newSimplePdf -pages @($simple, $simple, $simple, $simple)) }
            "tooManyPages" { return , (newSimplePdf -pages @($simple, $simple, $simple)) }
        }
        throw "unknown kind: $kind"
    }

    function getBrokenLimits([string]$kind) {
        switch ($kind) {
            "inflatedOver" { return (newPdfLimits -maxInflated 100) }
            "nestedValue" { return (newPdfLimits -maxDepth 4) }
            "nestedPages" { return (newPdfLimits -maxDepth 3) }
            "tooManyObjects" { return (newPdfLimits -maxObjects 5) }
            "tooManyPages" { return (newPdfLimits -maxPages 2) }
            "uncompressedOver" { return (newPdfLimits -maxInflated 100) }
            "uncompressedBig" { return (newPdfLimits -maxInflated 100) }
            "mapOverOne" { return (newPdfLimits -maxMapEntries 1000) }
            "mapOverTotal" { return (newPdfLimits -maxMapEntries 1000) }
            "lengthChain" { return (newPdfLimits -maxDepth 4) }
            "prevTooMany" { return (newPdfLimits -maxXref 3) }
            "arrayTooMany" { return (newPdfLimits -maxObjects 5) }
            "dictTooMany" { return (newPdfLimits -maxObjects 5) }
            "xrefRowsRepeat" { return (newPdfLimits -maxObjects 8) }
            "compressedOver" { return (newPdfLimits -maxInflated 110) }
            "searchOver" { return (newPdfLimits -maxSearchRatio 0.5) }
        }
        return $null
    }
}

Describe "testPdfExtension" -Tag Unit {
    It "<path> は <expected>" -TestCases @(
        @{ path = "a.pdf"; expected = $true }
        @{ path = "C:\共有\A.PDF"; expected = $true }
        @{ path = "a.pdf.txt"; expected = $false }
        @{ path = "a.docx"; expected = $false }
        @{ path = "pdf"; expected = $false }
    ) {
        param ($path, $expected)
        testPdfExtension $path | Should -Be $expected
    }
}

Describe "newPdfLimits" -Tag Unit {
    It "既定の上限" {
        $l = newPdfLimits
        $l.MaxInflated | Should -Be ${pdfInflatedMaxBytes}
        $l.MaxObjects | Should -Be ${pdfObjectMaxCount}
        $l.MaxDepth | Should -Be ${pdfNestMaxDepth}
        $l.MaxPages | Should -Be ${pdfPageMaxCount}
        $l.MaxMapEntries | Should -Be ${pdfMapMaxEntries}
        $l.MaxXref | Should -Be ${pdfXrefMaxCount}
    }
}

Describe "readPdfPages" -Tag Unit {
    It "<name>（圧縮なし・Flate）" -TestCases $textCases {
        param ($name, $pages, $expected)
        foreach ($compress in $false, $true) {
            $r = readPdfPages (newSimplePdf -pages $pages -compress:$compress)
            $r.Count | Should -Be $expected.Count
            for ($i = 0; $i -lt $expected.Count; $i++) { $r[$i] | Should -BeExactly $expected[$i] }
        }
    }

    It "xref ストリーム + オブジェクトのストリーム（<name>）" -TestCases @(
        @{ name = "predictor なし"; predictor = $false; widths = @(1, 4, 2); pngType = 2; useIndex = $false }
        @{ name = "PNG Up"; predictor = $true; widths = @(1, 4, 2); pngType = 2; useIndex = $false }
        @{ name = "PNG なし（行の種類 0）"; predictor = $true; widths = @(1, 4, 2); pngType = 0; useIndex = $false }
        @{ name = "PNG Sub・幅 1 3 1"; predictor = $true; widths = @(1, 3, 1); pngType = 1; useIndex = $false }
        @{ name = "PNG Average・幅 1 2 2"; predictor = $true; widths = @(1, 2, 2); pngType = 3; useIndex = $false }
        @{ name = "PNG Paeth・Index あり"; predictor = $true; widths = @(1, 4, 2); pngType = 4; useIndex = $true }
        @{ name = "predictor なし・Index あり"; predictor = $false; widths = @(1, 3, 1); pngType = 2; useIndex = $true }
    ) {
        param ($name, $predictor, $widths, $pngType, $useIndex)
        $n = @{}
        $n[1] = '<< /Type /Catalog /Pages 2 0 R >>'
        $n[2] = '<< /Type /Pages /Kids [10 0 R] /Count 1 >>'
        $n[11] = newPdfStreamObject '/Filter /FlateDecode' (compressPdfBytes (toPdfBytes "BT /F1 12 Tf (Packed Text) Tj ET"))
        $p = @{}
        $p[3] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
        $p[10] = '<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 3 0 R >> >> /Contents 11 0 R >>'
        $r = readPdfPages (newPdfWithXrefStream -normal $n -packed $p -predictor:$predictor -widths $widths -pngType $pngType -useIndex:$useIndex)
        $r.Count | Should -Be 1
        $r[0] | Should -BeExactly "Packed Text"
    }

    It "Prev をたどり、新しい更新の内容が勝つ" {
        $s = newPdfString (newSimplePdf -pages @("BT /F1 12 Tf (old) Tj ET"))
        $prevPos = getStartXref $s
        $new = "BT /F1 12 Tf (new) Tj ET"
        $objPos = $s.Length
        $s += "11 0 obj`n<< /Length $($new.Length) >>`nstream`n$new`nendstream`nendobj`n"
        $xrefPos = $s.Length
        $s += ("xref`n11 1`n{0:D10} 00000 n `ntrailer`n<< /Size 12 /Root 1 0 R /Prev $prevPos >>`nstartxref`n$xrefPos`n%%EOF`n" -f $objPos)
        (readPdfPages (toPdfBytes $s))[0] | Should -BeExactly "new"
    }

    It "Prev が何本続いても（上限の中なら）読める" {
        $r = readPdfPages (addXrefUpdates (newSimplePdf -pages @("BT /F1 12 Tf (ok) Tj ET")) 20)
        $r[0] | Should -BeExactly "ok"
    }

    It "XRefStm（表で解放済みとされた番号を相互参照のストリームで補う）" {
        $s = newPdfString (newSimplePdf -pages @("BT /F1 12 Tf (hybrid) Tj ET"))
        $table = $s.LastIndexOf("xref`n0 ")
        $lineStart = $s.IndexOf("`n", $table + 5) + 1 + 20 * 11
        $objOffset = [int]$s.Substring($lineStart, 10)
        $s = $s.Substring(0, $lineStart) + "0000000000 00000 f `n" + $s.Substring($lineStart + 20)
        $s = $s.Replace("/Root 1 0 R >>", "/Root 1 0 R /XRefStm 0000000000 >>")
        $stmPos = $s.Length
        $s = $s.Replace("0000000000 >>", ("{0:D10} >>" -f $stmPos))
        $row = [string][char]1 + [string][char]0 + [string][char]0 + [string][char]($objOffset -shr 8) + [string][char]($objOffset -band 255) + [string][char]0 + [string][char]0
        $s += "20 0 obj`n<< /Type /XRef /W [1 4 2] /Index [11 1] /Size 21 /Length 7 >>`nstream`n$row`nendstream`nendobj`n"
        (readPdfPages (toPdfBytes $s))[0] | Should -BeExactly "hybrid"
    }

    It "内容の /Length が間接参照でも読む" {
        $body = toPdfBytes "BT /F1 12 Tf (indirect) Tj ET"
        $objs = @{}
        $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
        $objs[2] = '<< /Type /Pages /Kids [3 0 R] /Count 1 >>'
        $objs[3] = '<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>'
        $objs[4] = @{ Dict = '<< /Length 20 0 R >>'; Stream = $body }
        $objs[5] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
        $objs[20] = "$($body.Length)"
        (readPdfPages (newPdfFromObjects $objs))[0] | Should -BeExactly "indirect"
    }

    It "/Length が合わないストリームを、ページの並び（Kids）が番号の降順でも読む" {
        $objs = @{}
        $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
        $objs[2] = '<< /Type /Pages /Kids [20 0 R 18 0 R 16 0 R 14 0 R 12 0 R] /Count 5 >>'
        $objs[3] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
        foreach ($i in 1..5) {
            $page = 10 + $i * 2
            $objs[$page] = "<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 3 0 R >> >> /Contents $($page + 1) 0 R >>"
            $objs[$page + 1] = @{ Dict = '<< /Length 3 >>'; Stream = (toPdfBytes "BT /F1 12 Tf (p$i) Tj ET") }
        }
        $r = readPdfPages (newPdfFromObjects $objs)
        ($r -join ",") | Should -BeExactly "p5,p4,p3,p2,p1"
    }

    It "終わりを探す量は、/Length が合わないストリームが降順に並んでも、実際に探した分だけ数える（上限の内で読める）" {
        $objs = @{}
        $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
        $objs[2] = '<< /Type /Pages /Kids [10 0 R 9 0 R 8 0 R 7 0 R 6 0 R 5 0 R 4 0 R 3 0 R] /Count 8 >>'
        foreach ($i in 3..10) {
            $objs[$i] = @{ Dict = '<< /Length 1 >>'; Stream = (toPdfBytes ("x" * 200)) }
        }
        $bytes = newPdfFromObjects $objs
        { readPdfPages $bytes (newPdfLimits -maxSearchRatio 0.75) } | Should -Not -Throw
    }

    It "相互参照の示す位置が合わないオブジェクトのストリームでも、番号から探して読む" {
        $n = @{}
        $n[1] = '<< /Type /Catalog /Pages 2 0 R >>'
        $n[2] = '<< /Type /Pages /Kids [10 0 R] /Count 1 >>'
        $n[11] = newPdfStreamObject '/Filter /FlateDecode' (compressPdfBytes (toPdfBytes "BT /F1 12 Tf (Packed Text) Tj ET"))
        $p = @{}
        $p[3] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
        $p[10] = '<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 3 0 R >> >> /Contents 11 0 R >>'
        $r = readPdfPages (newPdfWithXrefStream -normal $n -packed $p -reverseIndex)
        ($r -join ",") | Should -BeExactly "Packed Text"
    }

    It "内容の /Length が合わなくても endstream までを読む（<declared>）" -TestCases @(
        @{ declared = "5" }
        @{ declared = "99999" }
        @{ declared = "(x)" }
    ) {
        param ($declared)
        $objs = @{}
        $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
        $objs[2] = '<< /Type /Pages /Kids [3 0 R 6 0 R] /Count 2 >>'
        $objs[3] = '<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>'
        $objs[4] = @{ Dict = "<< /Length $declared >>"; Stream = (toPdfBytes "BT /F1 12 Tf (first) Tj ET") }
        $objs[5] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
        $objs[6] = '<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /Contents 7 0 R >>'
        $objs[7] = @{ Dict = "<< /Length $declared >>"; Stream = (toPdfBytes "BT /F1 12 Tf (second) Tj ET") }
        $r = readPdfPages (newPdfFromObjects $objs)
        $r[0] | Should -BeExactly "first"
        $r[1] | Should -BeExactly "second"
    }

    It "Resources は親のページツリーから受け継ぐ" {
        $objs = @{}
        $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
        $objs[2] = '<< /Type /Pages /Kids [3 0 R] /Count 1 /Resources << /Font << /F1 5 0 R >> >> >>'
        $objs[3] = '<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>'
        $objs[4] = newPdfStreamObject '' (toPdfBytes "BT /F1 12 Tf (inherited) Tj ET")
        $objs[5] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
        (readPdfPages (newPdfFromObjects $objs))[0] | Should -BeExactly "inherited"
    }

    It "フォント名の #xx と直接書いたフォントの辞書" {
        $objs = @{}
        $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
        $objs[2] = '<< /Type /Pages /Kids [3 0 R] /Count 1 >>'
        $objs[3] = '<< /Type /Page /Parent 2 0 R /Resources << /Font << /F#31 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> /Contents 4 0 R >>'
        $objs[4] = newPdfStreamObject '' (toPdfBytes "BT /F1 12 Tf (caf\351) Tj ET")
        (readPdfPages (newPdfFromObjects $objs))[0] | Should -BeExactly "café"
    }

    It "単純フォントにも ToUnicode を使う（圧縮した表。<name>）" -TestCases @(
        @{ name = "圧縮なし"; compress = $false }
        @{ name = "Flate"; compress = $true }
    ) {
        param ($name, $compress)
        $cmap = toPdfBytes "1 begincodespacerange <00> <FF> endcodespacerange 1 beginbfchar <41> <0058> endbfchar"
        $objs = @{}
        $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
        $objs[2] = '<< /Type /Pages /Kids [3 0 R] /Count 1 >>'
        $objs[3] = '<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>'
        $objs[4] = newPdfStreamObject '' (toPdfBytes "BT /F1 12 Tf (AB) Tj ET")
        $objs[5] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /ToUnicode 6 0 R >>'
        $objs[6] = if ($compress) { newPdfStreamObject '/Filter /FlateDecode' (compressPdfBytes $cmap) } else { newPdfStreamObject '' $cmap }
        (readPdfPages (newPdfFromObjects $objs))[0] | Should -BeExactly "XB"
    }

    It "同じ ToUnicode を使うフォントは、表の項目を 1 度だけ数える" {
        $r = readPdfPages (newMapPdf "BT /F1 12 Tf <0001> Tj /F2 12 Tf <0001> Tj ET" 600 $true) (newPdfLimits -maxMapEntries 1000)
        $r[0] | Should -BeExactly "BB"
    }

    It "直接書いたフォントの辞書を何度選んでも、表は 1 度だけ読む" {
        $content = "BT /F1 12 Tf <0001> Tj /F1 12 Tf <0001> Tj /F1 12 Tf <0001> Tj ET"
        $r = readPdfPages (newMapPdf $content 600 $true $true) (newPdfLimits -maxMapEntries 1000)
        $r[0] | Should -BeExactly "BBB"
    }

    It "内容ストリームが配列でも順に読む" {
        $objs = @{}
        $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
        $objs[2] = '<< /Type /Pages /Kids [3 0 R] /Count 1 >>'
        $objs[3] = '<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 7 0 R >> >> /Contents [4 0 R 5 0 R] >>'
        $objs[4] = newPdfStreamObject '' (toPdfBytes "BT /F1 12 Tf (first) Tj")
        $objs[5] = newPdfStreamObject '' (toPdfBytes "0 -14 Td (second) Tj ET")
        $objs[7] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
        (readPdfPages (newPdfFromObjects $objs))[0] | Should -BeExactly "first`nsecond"
    }

    It "ToUnicode: <name>" -TestCases $cidCases {
        param ($name, $content, $expected)
        $cid = newCidFontObjects $true
        $r = readPdfPages (newSimplePdf -pages @($content) -fontObject $cid.Font -extraObjects $cid.Extra)
        $r[0] | Should -BeExactly $expected
    }

    It "読めない PDF（<kind>）は「PDF」で始まる例外" -TestCases $brokenCases {
        param ($kind, $message)
        $bytes = newBrokenPdf $kind
        $limits = getBrokenLimits $kind
        $thrown = $null
        try { [void](readPdfPages $bytes $limits) } catch { $thrown = $_.Exception.Message }
        $thrown | Should -Not -BeNullOrEmpty
        $thrown | Should -BeLike "PDF*"
        $thrown | Should -BeLike $message
    }

    It "暗号化された PDF（<method>）を、空の利用者パスワードで復号して読む（圧縮なし・Flate）" -TestCases @(
        @{ method = "aes256r5" }
        @{ method = "aes256r6" }
    ) {
        param ($method)
        $pages = @("BT /F1 12 Tf 72 700 Td (secret abc) Tj ET", "BT /F1 12 Tf (page two) Tj ET")
        foreach ($compress in $false, $true) {
            $r = readPdfPages (newEncryptedPdf -pages $pages -method $method -compress:$compress)
            $r.Count | Should -Be 2
            $r[0] | Should -BeExactly "secret abc"
            $r[1] | Should -BeExactly "page two"
        }
    }

    It "暗号化された PDF（AES-256）でも、相互参照がストリームなら読む（相互参照は暗号化されない）" {
        $r = readPdfPages (newEncryptedPdf -pages @("BT /F1 12 Tf (abc) Tj ET") -method aes256r6 -compress -xrefStream)
        $r[0] | Should -BeExactly "abc"
    }

    It "暗号化された PDF（<method>）で、カタログ・ページの木・ページがオブジェクトのストリームの中にあっても読む（ストリームは 1 回だけ復号し、中は復号しない）" -TestCases @(
        @{ method = "aes256r5" }
        @{ method = "aes256r6" }
    ) {
        param ($method)
        $pages = @("BT /F1 12 Tf (first) Tj ET", "BT /F1 12 Tf (second) Tj ET")
        $r = readPdfPages (newEncryptedPdf -pages $pages -method $method -compress -objectStream)
        $r.Count | Should -Be 2
        $r[0] | Should -BeExactly "first"
        $r[1] | Should -BeExactly "second"
    }

    It "開くのにパスワードが要る PDF（<method>）は、そのファイルだけ「パスワードが必要」で失敗する" -TestCases @(
        @{ method = "aes256r5" }
        @{ method = "aes256r6" }
    ) {
        param ($method)
        $bytes = newEncryptedPdf -pages @("BT /F1 12 Tf (abc) Tj ET") -method $method -userPassword "pass"
        { readPdfPages $bytes } | Should -Throw "PDFにパスワードが必要なため読めません。"
    }

    It "AES-256 以外の暗号化（<method>）は、「方式に対応していない」でそのファイルだけ失敗する（パスワードが要るかは判断しない）" -TestCases @(
        @{ method = "rc40" }
        @{ method = "rc128" }
        @{ method = "aes128" }
    ) {
        param ($method)
        $bytes = newEncryptedPdf -pages @("BT /F1 12 Tf (abc) Tj ET") -method $method
        { readPdfPages $bytes } | Should -Throw "PDFの暗号化の方式に対応していないため読めません。"
    }

    It "暗号化の辞書が対応しない・壊れているもの（<name>）は、そのファイルだけ失敗する" -TestCases @(
        @{ name = "独自のフィルター"; dict = "/Filter /Adobe.PubSec /V 5 /R 6"; message = "PDFの暗号化の方式に対応していないため読めません。" }
        @{ name = "未知の V"; dict = "/Filter /Standard /V 3 /R 3"; message = "PDFの暗号化の方式に対応していないため読めません。" }
        @{ name = "V が 5 でない小数（4.6）"; dict = "/Filter /Standard /V 4.6 /R 6"; message = "PDFの暗号化の方式に対応していないため読めません。" }
        @{ name = "V が 5 でない小数（5.4）"; dict = "/Filter /Standard /V 5.4 /R 6"; message = "PDFの暗号化の方式に対応していないため読めません。" }
        @{ name = "R が 5・6 でない小数（6.4）"; dict = "/Filter /Standard /V 5 /R 6.4"; message = "PDFの暗号化の方式に対応していないため読めません。" }
        @{ name = "V と R の組み合わせが違う"; dict = "/Filter /Standard /V 5 /R 4"; message = "PDFの暗号化の方式に対応していないため読めません。" }
        @{ name = "V が数でない"; dict = "/Filter /Standard /V /x /R 6"; message = "PDFの暗号化の方式に対応していないため読めません。" }
        @{ name = "未知の暗号方式"; dict = "/Filter /Standard /V 5 /R 6 /CF << /StdCF << /CFM /Other >> >> /StmF /StdCF /U <00> /UE <00>"; message = "PDFの暗号化の方式に対応していないため読めません。" }
        @{ name = "StmF が名前でない"; dict = "/Filter /Standard /V 5 /R 6 /CF << /StdCF << /CFM /AESV3 >> >> /StmF 5 /U <00> /UE <00>"; message = "PDFの暗号化の方式に対応していないため読めません。" }
        @{ name = "StmF があるのに CF が辞書でない"; dict = "/Filter /Standard /V 5 /R 6 /CF 5 /StmF /StdCF /U <00> /UE <00>"; message = "PDFの暗号化の方式に対応していないため読めません。" }
        @{ name = "短い U と UE（R6）"; dict = "/Filter /Standard /V 5 /R 6 /CF << /StdCF << /CFM /AESV3 >> >> /StmF /StdCF /O <00> /U <00> /UE <00>"; message = "PDFの暗号化の情報が壊れているため読めません。" }
    ) {
        param ($name, $dict, $message)
        $bytes = newEncryptedPdf -pages @("BT /F1 12 Tf (abc) Tj ET") -method aes256r6 -encryptDict $dict
        { readPdfPages $bytes } | Should -Throw $message
    }

    It "組み立てた PDF に作成者・絶対パスが入っていない" {
        $b = newSimplePdf -pages @("BT /F1 12 Tf (x) Tj ET") -compress
        $s = [System.Text.Encoding]::GetEncoding(28591).GetString($b)
        $s | Should -Not -Match "/Author|/Creator|/Producer|C:\\|Users"
    }
}

Describe "readPdfFile" -Tag Io {
    It "ファイルから読む" {
        $path = Join-Path $TestDrive "a.pdf"
        [System.IO.File]::WriteAllBytes($path, (newSimplePdf -pages @("BT /F1 12 Tf (file) Tj ET")))
        (readPdfFile $path)[0] | Should -BeExactly "file"
    }

    It "大きさの上限を超えたら読まずに失敗する" {
        $path = Join-Path $TestDrive "big.pdf"
        [System.IO.File]::WriteAllBytes($path, (newSimplePdf -pages @("BT /F1 12 Tf (file) Tj ET")))
        { readPdfFile $path 10 } | Should -Throw "*ファイルサイズが大きすぎる*"
    }
}

Describe "readPdfPageText（命令の読み方）" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "文字列の中の BT ET は命令にしない"; content = 'BT (BT ET) Tj ET'; expected = "BT ET" }
        @{ name = "入れ子の括弧とエスケープ"; content = 'BT (a(b)c\)d) Tj ET'; expected = "a(b)c)d" }
        @{ name = "TJ の字間で空白を入れる"; content = 'BT [(A) -300 (B) 20 (C)] TJ ET'; expected = "A BC" }
        @{ name = "区切りのない TJ"; content = 'BT [(A)-300(B)] TJ ET'; expected = "A B" }
        @{ name = "コメントの中の命令は読まない"; content = "BT % (x) Tj`n(y) Tj ET"; expected = "y" }
        @{ name = "図形・色の命令は読み飛ばす"; content = 'q 1 0 0 RG 0 0 m 5 5 l S Q BT (z) Tj ET'; expected = "z" }
        @{ name = "Tm と Td で行を分ける"; content = 'BT /F1 12 Tf 1 0 0 1 10 700 Tm (a) Tj 0 -14 Td (b) Tj ET'; expected = "a`nb" }
        @{ name = "インライン画像の中は読まない"; content = 'BI /W 1 /H 1 ID (x) Tj EI BT (p) Tj ET'; expected = "p" }
        @{ name = "インライン画像の中の偽の EI は終わりにしない"; content = 'BI /W 1 /H 1 ID xEI BT (q) EI BT (p) Tj ET'; expected = "p" }
        @{ name = "辞書の中の命令は読まない"; content = '/Span << /MCID 0 /A Tj >> BDC BT (m) Tj ET EMC'; expected = "m" }
        @{ name = "16 進と 8 進の文字列"; content = 'BT <41 42> Tj (\101\102) Tj ET'; expected = "ABAB" }
        @{ name = "TL・T*・引用符で行を分ける"; content = "BT 12 TL (a) Tj T* (b) Tj (c) ' ET"; expected = "a`nb`nc" }
        @{ name = "BT の外の文字列も出す"; content = '0 g (n) Tj BT (m) Tj ET'; expected = "nm" }
        @{ name = "TJ の中の指数の数は字間"; content = 'BT [(a) 1e3 (b)] TJ ET'; expected = "ab" }
        @{ name = "波括弧のあとも読む"; content = '{ (a) Tj } BT (b) Tj ET'; expected = "ab" }
        @{ name = "文字に効かない命令の前後の引数"; content = 'BT 5 w 0 -14 Td (a) Tj 3 7 Td /GS gs (b) Tj ET'; expected = "a`nb" }
        @{ name = "読み飛ばす並びの文字列・名前は残らない"; content = '(skip) /N 1 2 re (a) Tj'; expected = "a" }
        @{ name = "BT の外の Td と色"; content = '5 7 Td 1 0 0 rg (a) Tj'; expected = "a" }
        @{ name = "TL のあとの線幅の命令"; content = 'BT 14 TL (a) Tj 9 w T* (b) Tj ET'; expected = "a`nb" }
        @{ name = "前方一致: Tj の続き"; content = '(a) Tjx (b) Tj'; expected = "b" }
        @{ name = "前方一致: BI の続き"; content = 'BIx (a) Tj'; expected = "a" }
        @{ name = "前方一致: T* の続き"; content = '(a) Tj T*x (b) Tj'; expected = "ab" }
        @{ name = "大文字・小文字を区別する"; content = '(a) tj (b) Tj'; expected = "b" }
        @{ name = "空白なしの w と文字列"; content = '1 w(a)Tj'; expected = "a" }
        @{ name = "空白なしの BT と名前"; content = 'BT/F1 12 Tf(a)Tj 0 -14 Td(b)Tj ET'; expected = "a`nb" }
        @{ name = "空白なしの BI"; content = 'BI/W 1 ID x EI (p) Tj'; expected = "p" }
        @{ name = "二重引用符で行を分ける"; content = '(a) Tj 0 0 (b) "'; expected = "a`nb" }
        @{ name = "引用符で始まる未知の語"; content = "(a) 'x (b) Tj"; expected = "b" }
        @{ name = "英字で始まらない未知の語（*）"; content = '(a) *x Tj'; expected = "" }
        @{ name = "英字で始まらない未知の語（数字）"; content = '(a) 1x Tj'; expected = "" }
        @{ name = "コメントの中の命令の名前"; content = "1 0 0 % rg (x) Tj`n rg (a) Tj"; expected = "a" }
        @{ name = "配列の中のコメント"; content = "[(a) % (x)`n (b)] TJ"; expected = "ab" }
        @{ name = "読み飛ばしの直後のエスケープのある文字列"; content = '1 w (a\)b) Tj'; expected = "a)b" }
        @{ name = "読み飛ばしの直後の 16 進の文字列"; content = '0 g <41> Tj'; expected = "A" }
        @{ name = "配列の中のエスケープ・入れ子・16 進の文字列"; content = '[(a\)) (b(c)) <41>] TJ'; expected = "a)b(c)A" }
        @{ name = "辞書の入れ子"; content = '/P << /A << /B (x) Tj >> >> BDC (a) Tj'; expected = "a" }
        @{ name = "閉じない文字列"; content = '(a) Tj (bc'; expected = "a" }
        @{ name = "閉じない配列"; content = '(a) Tj [(b) (c)'; expected = "a" }
        @{ name = "ID のあとが閉じないインライン画像"; content = '(a) Tj BI /W 1 ID xx'; expected = "a" }
        @{ name = "ID の無いインライン画像"; content = '(a) Tj BI /W 1'; expected = "a" }
        @{ name = "符号つきの数"; content = 'BT (a) Tj +0 -14. Td (b) Tj ET'; expected = "a`nb" }
        @{ name = "TJ の中の小数の字間"; content = '[(A) -.3e3 (B)] TJ'; expected = "A B" }
        @{ name = "TJ の中の数でないものは字間にしない"; content = '[(A) --300 (B)] TJ'; expected = "AB" }
        @{ name = "空白の種類（タブ・改ページ・復帰）"; content = "(a)`t Tj`f(b)`r Tj"; expected = "ab" }
        @{ name = "空白の種類（コード 0）"; content = "(a)$([char]0)Tj$([char]0)(b)$([char]0)Tj"; expected = "ab" }
        @{ name = "空の文字列"; content = '() Tj (a) Tj'; expected = "a" }
    ) {
        readPdfPageText @{ Fonts = @{}; CMaps = @{} } $content $null | Should -BeExactly $expected
    }
}

Describe "readPdfPageText（引数が延々と続く内容でも終わる）" -Tag Unit {
    # 読み飛ばしの試みが字句の数の 2 乗の時間になっていないことを、結果と、余裕のある上限の時間で確かめる
    It "<name>" -TestCases @(
        @{ name = "数が続いてから Tj"; make = { ('0 ' * 200000) + '(a) Tj' }; expected = "a" }
        @{ name = "文字列が続いてから Tj"; make = { ('(x) ' * 100000) + 'Tj' }; expected = "x" }
        @{ name = "名前が続いて命令なしで終わる"; make = { '/N ' * 200000 }; expected = "" }
        @{ name = "数が続いてから TJ"; make = { ('0 ' * 200000) + '[(a)] TJ' }; expected = "a" }
    ) {
        $content = & $make
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        $text = readPdfPageText @{ Fonts = @{}; CMaps = @{} } $content $null
        $watch.Stop()
        $text | Should -BeExactly $expected
        $watch.Elapsed.TotalSeconds | Should -BeLessThan 120
    }
}

Describe "PDF の字句の正規表現（部品と分岐のつながり）" -Tag Unit {
    It "語の字の表は、区切りと空白の表（pdfDelimiters）と同じ" {
        $delimiters = [System.Collections.Generic.HashSet[char]]::new([char[]]${pdfDelimiters})
        $rest = [regex]::new('^' + ${pdfRegexRest} + '\z')
        $start = [regex]::new('^[^' + ${pdfRegexNotStart} + ']\z')
        $restDiff = @(0..255 | Where-Object { $rest.IsMatch([string][char]$_) -eq $delimiters.Contains([char]$_) })
        $startDiff = @(0..255 | Where-Object { $start.IsMatch([string][char]$_) -eq ($delimiters.Contains([char]$_) -or $_ -le 32) })
        $restDiff | Should -BeNullOrEmpty
        $startDiff | Should -BeNullOrEmpty
    }

    It "文字に効く命令の一覧は、readPdfPageText の分岐と同じ" {
        $source = (Get-Command readPdfPageText).ScriptBlock.ToString()
        $branches = @([regex]::Matches($source, '\$t -ceq (?:"([^"]+)"|''([^'']+)'')') | ForEach-Object {
            if ($_.Groups[1].Success) { $_.Groups[1].Value } else { $_.Groups[2].Value }
        } | Sort-Object -Unique -CaseSensitive)
        $listed = @(@(${pdfTextOpNames}) + "'" + '"' | Sort-Object -Unique -CaseSensitive)
        ($branches -join " ") | Should -BeExactly ($listed -join " ")
    }

    It "実行時のコンパイル（Compiled）を使わない" {
        ([int](${pdfTokenRegex}.Options -band [System.Text.RegularExpressions.RegexOptions]::Compiled)) | Should -Be 0
        ([int](${pdfPageRegex}.Options -band [System.Text.RegularExpressions.RegexOptions]::Compiled)) | Should -Be 0
    }
}
