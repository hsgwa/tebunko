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
    )

    # Type0 フォント（Identity-H）＋ ToUnicode の表
    $cidCases = @(
        @{ name = "bfchar と bfrange"; content = "BT /F1 12 Tf <0001001000110012> Tj ET"; expected = "あいぅう" }
        @{ name = "サロゲートペア"; content = "BT /F1 12 Tf <0002> Tj ET"; expected = $surrogate }
        @{ name = "TJ の詰めで前後が半角なら空白"; content = "BT /F1 12 Tf [<0020> -300 <0021>] TJ ET"; expected = "A B" }
        @{ name = "TJ の詰めで前後が全角なら空白なし"; content = "BT /F1 12 Tf [<0001> -300 <0001>] TJ ET"; expected = "ああ" }
        @{ name = "TJ の詰めで片方が全角なら空白なし"; content = "BT /F1 12 Tf [<0020> -300 <0001>] TJ ET"; expected = "Aあ" }
        @{ name = "配列の bfrange"; content = "BT /F1 12 Tf <0030> Tj <0031> Tj ET"; expected = "XY" }
    )

    # 読めない PDF。kind は BeforeAll の newBrokenPdf が組み立てる
    $brokenCases = @(
        @{ kind = "empty"; message = "*PDF*" }
        @{ kind = "notPdf"; message = "*PDF*" }
        @{ kind = "truncated"; message = "*PDF*" }
        @{ kind = "encrypted"; message = "*暗号化*" }
        @{ kind = "startxrefOutside"; message = "*ファイルの外*" }
        @{ kind = "entryOutside"; message = "*PDF*" }
        @{ kind = "kidsCycle"; message = "*循環*" }
        @{ kind = "refCycle"; message = "*循環*" }
        @{ kind = "garbageFlate"; message = "*PDF*" }
        @{ kind = "lzw"; message = "*未対応の圧縮*" }
        @{ kind = "cidNoToUnicode"; message = "*文字の対応表が無い*" }
        @{ kind = "inflatedOver"; message = "*解凍後*" }
        @{ kind = "nestedValue"; message = "*入れ子*" }
        @{ kind = "nestedPages"; message = "*入れ子*" }
        @{ kind = "tooManyObjects"; message = "*オブジェクトの数*" }
        @{ kind = "tooManyPages"; message = "*ページ数*" }
    )
}

BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "$PSScriptRoot\..\..\helpers\pdf.ps1"

    function newCidFontObjects {
        # フォント F1（obj 3）を Type0 にする。toUnicode が $false なら ToUnicode（obj 5）を付けない
        param ([bool]$toUnicode = $true)

        $cmap = @"
/CIDInit /ProcSet findresource begin 12 dict begin begincmap
/CMapName /Adobe-Identity-UCS def /CMapType 2 def
1 begincodespacerange <0000> <FFFF> endcodespacerange
2 beginbfchar <0001> <3042> <0002> <D840DC0B> endbfchar
2 beginbfrange <0010> <0012> <3044> <0020> <0021> <0041> endbfrange
1 beginbfrange <0030> <0031> [<0058> <0059>] endbfrange
endcmap end end
"@
        $to = if ($toUnicode) { " /ToUnicode 5 0 R" } else { "" }
        $objs = @{}
        $objs[4] = '<< /Type /Font /Subtype /CIDFontType2 /BaseFont /Test /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> >>'
        $objs[5] = newPdfStreamObject '' (toPdfBytes $cmap)
        return @{
            Font  = "<< /Type /Font /Subtype /Type0 /BaseFont /Test /Encoding /Identity-H /DescendantFonts [4 0 R]$to >>"
            Extra = $objs
        }
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
                $objs[3] = '<< /Filter /Standard /V 1 /R 2 >>'
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

    It "xref ストリーム + オブジェクトのストリーム（predictor <predictor>）" -TestCases @(
        @{ predictor = $false }
        @{ predictor = $true }
    ) {
        param ($predictor)
        $n = @{}
        $n[1] = '<< /Type /Catalog /Pages 2 0 R >>'
        $n[2] = '<< /Type /Pages /Kids [10 0 R] /Count 1 >>'
        $n[11] = newPdfStreamObject '/Filter /FlateDecode' (compressPdfBytes (toPdfBytes "BT /F1 12 Tf (Packed Text) Tj ET"))
        $p = @{}
        $p[3] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
        $p[10] = '<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 3 0 R >> >> /Contents 11 0 R >>'
        $r = readPdfPages (newPdfWithXrefStream -normal $n -packed $p -predictor:$predictor)
        $r.Count | Should -Be 1
        $r[0] | Should -BeExactly "Packed Text"
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
