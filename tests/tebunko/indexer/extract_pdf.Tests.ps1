# 1 ファイル（PDF）の抽出（tebunko\indexer\extract_pdf.ps1）のテスト。PDF はテストの中で組み立てる（tests\helpers\pdf.ps1）。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "$PSScriptRoot\..\..\helpers\pdf.ps1"
    . "${scriptsDir}\shared\office\office_reader.ps1"
    . "${scriptsDir}\tebunko\indexer\extract_pdf.ps1"

    function newSourceFile([string]$name, [byte[]]$bytes) {
        $path = Join-Path $TestDrive $name
        [System.IO.File]::WriteAllBytes($path, $bytes)
        return $path
    }
}

Describe "extractPdfFile" -Tag Io {
    BeforeEach {
        $outDir = Join-Path $TestDrive "out-$(New-Guid)"
        [System.IO.Directory]::CreateDirectory($outDir) | Out-Null
    }

    It "ページごとの TSV を書き出し、コピーが残らない" {
        $source = newSourceFile "a.pdf" (newSimplePdf -pages @("BT /F1 12 Tf (one) Tj 0 -14 Td (line2) Tj ET", "BT /F1 12 Tf (two) Tj ET") -compress)
        (extractPdfFile $source $outDir) | Should -Be 2
        (@(Get-ChildItem -LiteralPath $outDir -File | ForEach-Object { $_.Name } | Sort-Object) -join ",") | Should -Be "page_001.tsv,page_002.tsv"
        ([System.IO.File]::ReadAllLines((Join-Path $outDir "page_001.tsv")) -join "|") | Should -Be "one|line2"
        ([System.IO.File]::ReadAllLines((Join-Path $outDir "page_002.tsv")) -join "|") | Should -Be "two"
    }

    It "文字の無いページは出さない" {
        $source = newSourceFile "gap.pdf" (newSimplePdf -pages @("BT /F1 12 Tf (one) Tj ET", "0 0 m 1 1 l S", "BT /F1 12 Tf (three) Tj ET"))
        (extractPdfFile $source $outDir) | Should -Be 2
        (@(Get-ChildItem -LiteralPath $outDir -File | ForEach-Object { $_.Name } | Sort-Object) -join ",") | Should -Be "page_001.tsv,page_003.tsv"
    }

    It "<name> は、コピーが残らず（出力フォルダが空）、結果は <expected>" -TestCases @(
        @{ name = "文字の無い PDF（画像だけ）"; kind = "noText"; message = $null; expected = "0" }
        @{ name = "暗号化された PDF"; kind = "encrypted"; message = "PDFが暗号化されているため読めません。"; expected = "失敗" }
        @{ name = "壊れた PDF"; kind = "broken"; message = "PDFの相互参照が見つからないため読めません。"; expected = "失敗" }
        @{ name = "大きさの上限を超える PDF"; kind = "tooBig"; message = "ファイルサイズが大きすぎるため更新できません。"; expected = "失敗" }
    ) {
        param ($name, $kind, $message, $expected)
        $max = 100MB
        switch ($kind) {
            "noText" { $bytes = newSimplePdf -pages @("0 0 m 1 1 l S") }
            "encrypted" {
                $objs = @{}
                $objs[1] = '<< /Type /Catalog /Pages 2 0 R >>'
                $objs[2] = '<< /Type /Pages /Kids [] /Count 0 >>'
                $objs[3] = '<< /Filter /Standard /V 1 /R 2 >>'
                $bytes = newPdfFromObjects $objs '/Root 1 0 R /Encrypt 3 0 R'
            }
            "broken" { $bytes = toPdfBytes "%PDF-1.4`nbroken" }
            "tooBig" { $bytes = newSimplePdf -pages @("BT /F1 12 Tf (x) Tj ET"); $max = 10 }
        }
        $source = newSourceFile "case.pdf" $bytes
        if ($message) {
            { extractPdfFile $source $outDir $max } | Should -Throw $message
        } else {
            (extractPdfFile $source $outDir $max) | Should -Be 0
        }
        (Get-ChildItem -LiteralPath $outDir -File).Count | Should -Be 0
    }

    It "解凍後の大きさの上限を超える PDF は失敗し、コピーが残らない" {
        $source = newSourceFile "inflate.pdf" (newSimplePdf -pages @("BT /F1 12 Tf (" + ("a" * 500) + ") Tj ET") -compress)
        { extractPdfFile $source $outDir 100MB (newPdfLimits -maxInflated 100) } | Should -Throw "PDFの解凍後の大きさが上限を超えるため読めません。"
        (Get-ChildItem -LiteralPath $outDir -File).Count | Should -Be 0
    }

    It "大きさの上限を超えるときは、コピーを作らない" {
        Mock copyFileShared { }
        $source = newSourceFile "big2.pdf" (newSimplePdf -pages @("BT /F1 12 Tf (x) Tj ET"))
        { extractPdfFile $source $outDir 10 } | Should -Throw "ファイルサイズが大きすぎるため更新できません。"
        Should -Invoke copyFileShared -Times 0 -Exactly
    }
}
