# Word・PowerPoint に埋め込んだ Office のファイル（Office Open XML。Word 文書・Excel ブック・PowerPoint）の中の文字を読み出す。
# 埋め込んだファイルは、親のファイルの ZIP の中の 1 つの部品（word/embeddings/*.xlsx など。リレーションシップの種類が package）で、
# それ自体が ZIP になっている。部品をバイト列で読み（readZipEntryBytes。部品ごと・1ファイルの合計の上限は親と同じに数える）、
# メモリ上で開いて読む。ファイルには書き出さない。埋め込んだブックのシートも同じ readZipEntryBytes で読むため、
# 申告の大きさを偽ったシートも、親と同じく部品ごとの上限（zipPartMaxBytes）で打ち切り、シートの大きさは合計に数える
# office_reader.ps1 を先に読み込んでおくこと（readDocxUnits・readPptxUnits が readEmbeddedObjectLines を呼ぶ）。
#
# ・形式は、拡張子や名前ではなく中身で決める（xl/workbook.xml があれば Excel、word/document.xml があれば Word、
#   ppt/presentation.xml があれば PowerPoint）。ZIP でない・どれにも当たらないものは読まない（失敗にもしない）
# ・Excel のブックは、表示のシートの文字のセル（共有文字列・インライン文字列・数式の文字の結果）を、行ごとに
#   タブ区切りの 1 行にして返す。表示の範囲に限らず、ほかの表示のシートも読む（非表示・完全に非表示のシートは読まない）
# ・Word・PowerPoint は、中の全ユニットの行を順に並べて返す（埋め込みの中の埋め込みは読まない。深さは 1 段まで）
# ・同じ部品を指す参照（代替表示と本体など）は 1 回だけ読む。N は 1 ファイルの中の、参照の順の通し番号
#   （番号は読む前に振るため、読めなかったものは欠番になる）
# ・出す量にも上限がある。共有文字列は 1 つを多数のセルが参照でき、数 KB の埋め込みから巨大な行ができるため、
#   読んだ入力の大きさとは別に、出す行の文字数（共有文字列を解いた後・参照のたびに数える）を 1 ファイル（親）の中の
#   全埋め込みで合計し、embeddedOutputMaxChars を超えたらその埋め込みを読まない（ZipSizeLimitException の LimitKind が Output）

# 1 ファイル（親の Word・PowerPoint）の全埋め込みから出す、行の文字数の合計の上限（文字）。
# $script: の変数は読み取りのスレッド（runspace）ごとに別になる（テストから一時的に値を変えるときは、この変数を直接上書きする）
$script:embeddedOutputMaxChars = 32MB

function newEmbeddedState {
    # 1 ファイル（親の Word・PowerPoint）の埋め込みを数える状態を返す（通し番号・読んだ部品の名前・出した文字数の合計）
    return @{ Next = 1; Seen = (New-Object System.Collections.Generic.HashSet[string]); OutputChars = 0L }
}

function addEmbeddedOutputChars {
    # 埋め込みから出す行の文字数を合計に足し、上限（embeddedOutputMaxChars）を超えたら ZipSizeLimitException（Output）にする。
    # $state が $null のときは数えない
    param (
        [hashtable]$state,
        [long]$chars
    )

    if ($null -eq $state) {
        return
    }
    $state.OutputChars += $chars
    if ($state.OutputChars -gt $script:embeddedOutputMaxChars) {
        throw [ZipSizeLimitException]::new($script:zipTooLargeMessage, "(埋め込みから出す文字)", $state.OutputChars, "Output")
    }
}

function readXlsxSharedStrings {
    # 共有文字列（xl/sharedStrings.xml）を、番号の順の配列で返す。読み仮名（rPh）は含めない。無ければ空
    param (
        [System.IO.Compression.ZipArchive]$zip
    )

    $xml = readZipEntry $zip "xl/sharedStrings.xml"
    $strings = New-Object System.Collections.Generic.List[string]
    if (-not $xml) {
        return $strings.ToArray()
    }

    $settings = New-Object System.Xml.XmlReaderSettings
    $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $reader = [System.Xml.XmlReader]::Create((New-Object System.IO.StringReader($xml)), $settings)
    try {
        $sb = New-Object System.Text.StringBuilder
        $inT = $false
        $inPhonetic = $false
        while ($reader.Read()) {
            $type = $reader.NodeType
            if ($type -eq [System.Xml.XmlNodeType]::Element) {
                $name = $reader.LocalName
                if ($name -eq "si") {
                    [void]$sb.Clear()
                    if ($reader.IsEmptyElement) { $strings.Add("") }
                } elseif ($name -eq "rPh") {
                    $inPhonetic = -not $reader.IsEmptyElement
                } elseif ($name -eq "t") {
                    $inT = -not $reader.IsEmptyElement
                }
            } elseif ($type -eq [System.Xml.XmlNodeType]::EndElement) {
                $name = $reader.LocalName
                if ($name -eq "si") {
                    $strings.Add($sb.ToString())
                } elseif ($name -eq "rPh") {
                    $inPhonetic = $false
                } elseif ($name -eq "t") {
                    $inT = $false
                }
            } elseif ($inT -and -not $inPhonetic) {
                if ($type -eq [System.Xml.XmlNodeType]::Text -or $type -eq [System.Xml.XmlNodeType]::SignificantWhitespace -or
                    $type -eq [System.Xml.XmlNodeType]::Whitespace -or $type -eq [System.Xml.XmlNodeType]::CDATA) {
                    [void]$sb.Append($reader.Value)
                }
            }
        }
    } finally {
        $reader.Dispose()
    }
    return $strings.ToArray()
}

function readXlsxSheetCellLines {
    # シートの部品（xl/worksheets/sheetN.xml）を流れで読み、文字のセルを行ごとにタブ区切りにして返す。
    # 文字のセルの文字数は、共有文字列を解いたあとで、参照のたびに $state の合計に足す（上限を超えたら例外。addEmbeddedOutputChars）
    param (
        [System.IO.Stream]$stream,
        [string[]]$sharedStrings,
        [hashtable]$state = $null
    )

    $settings = New-Object System.Xml.XmlReaderSettings
    $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $settings.MaxCharactersInDocument = $script:zipPartMaxBytes
    $reader = [System.Xml.XmlReader]::Create($stream, $settings)
    $lines = New-Object System.Collections.Generic.List[string]
    try {
        $rowCells = New-Object System.Collections.Generic.List[string]
        $cellType = $null
        $valueText = New-Object System.Text.StringBuilder
        $inlineText = New-Object System.Text.StringBuilder
        $inValue = $false
        $inInline = $false
        $inT = $false
        $inPhonetic = $false
        while ($reader.Read()) {
            $type = $reader.NodeType
            if ($type -eq [System.Xml.XmlNodeType]::Element) {
                $name = $reader.LocalName
                $isEmpty = $reader.IsEmptyElement
                if ($name -eq "row") {
                    $rowCells.Clear()
                } elseif ($name -eq "c") {
                    $cellType = $reader.GetAttribute("t")
                    [void]$valueText.Clear()
                    [void]$inlineText.Clear()
                } elseif ($name -eq "v") {
                    $inValue = -not $isEmpty
                } elseif ($name -eq "is") {
                    $inInline = -not $isEmpty
                } elseif ($name -eq "rPh") {
                    $inPhonetic = -not $isEmpty
                } elseif ($name -eq "t") {
                    $inT = -not $isEmpty
                }
            } elseif ($type -eq [System.Xml.XmlNodeType]::EndElement) {
                $name = $reader.LocalName
                if ($name -eq "v") {
                    $inValue = $false
                } elseif ($name -eq "is") {
                    $inInline = $false
                } elseif ($name -eq "rPh") {
                    $inPhonetic = $false
                } elseif ($name -eq "t") {
                    $inT = $false
                } elseif ($name -eq "c") {
                    $text = $null
                    if ($cellType -eq "s") {
                        $index = 0
                        if ([int]::TryParse($valueText.ToString().Trim(), [ref]$index) -and $index -ge 0 -and $index -lt $sharedStrings.Count) {
                            $text = $sharedStrings[$index]
                        }
                    } elseif ($cellType -eq "inlineStr") {
                        $text = $inlineText.ToString()
                    } elseif ($cellType -eq "str") {
                        $text = $valueText.ToString()
                    }
                    if ($null -ne $text) {
                        addEmbeddedOutputChars $state ($text.Length + 1)
                        $text = ($text -replace "[\r\n\t]+", " ").Trim()
                        if ($text -ne "") {
                            $rowCells.Add($text)
                        }
                    }
                } elseif ($name -eq "row") {
                    if ($rowCells.Count -gt 0) {
                        $lines.Add(($rowCells -join "`t"))
                    }
                    $rowCells.Clear()
                }
            } elseif ($type -eq [System.Xml.XmlNodeType]::Text -or $type -eq [System.Xml.XmlNodeType]::SignificantWhitespace -or
                $type -eq [System.Xml.XmlNodeType]::Whitespace -or $type -eq [System.Xml.XmlNodeType]::CDATA) {
                if ($inValue) {
                    [void]$valueText.Append($reader.Value)
                } elseif ($inInline -and $inT -and -not $inPhonetic) {
                    [void]$inlineText.Append($reader.Value)
                }
            }
        }
    } finally {
        $reader.Dispose()
    }
    return $lines.ToArray()
}

function readXlsxCellLines {
    # 開いたブック（ZipArchive）の、表示のワークシートの文字のセルを、シートの順・行の順に並べた行の配列で返す
    # （数値・日付・空のセルは読まない）。シートの部品は readZipEntryBytes でバイト列に読む（部品ごとの上限・
    # 1 ファイルの合計の上限は親と同じ。申告を偽ったシートは、部品ごとの上限の超過として例外になる）。
    # 出す文字数は $state の合計に数え、上限を超えたら例外にする
    param (
        [System.IO.Compression.ZipArchive]$zip,
        [hashtable]$state = $null
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $workbookXml = readZipEntry $zip "xl/workbook.xml"
    if ($null -eq $workbookXml) {
        return $lines.ToArray()
    }
    $workbook = newXmlDocument $workbookXml
    $workbookRels = readRelationships $zip "xl/workbook.xml"
    $nsSheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    $sharedStrings = $null
    foreach ($sheet in $workbook.GetElementsByTagName("sheet", $nsSheet)) {
        if ($sheet.GetAttribute("state") -in @("hidden", "veryHidden")) {
            continue
        }
        $rel = $workbookRels[$sheet.GetAttribute("id", ${nsRel})]
        if ($null -eq $rel -or $rel.Type -notlike "*/worksheet") {
            continue
        }
        $bytes = readZipEntryBytes $zip $rel.Target
        if ($null -eq $bytes) {
            continue
        }
        if ($null -eq $sharedStrings) {
            $sharedStrings = @(readXlsxSharedStrings $zip)
        }
        $stream = New-Object System.IO.MemoryStream (, $bytes)
        try {
            $lines.AddRange([string[]]@(readXlsxSheetCellLines $stream $sharedStrings $state))
        } finally {
            $stream.Dispose()
        }
    }
    return $lines.ToArray()
}

function readEmbeddedPackageLines {
    # 埋め込んだファイルのバイト列から、文字の行を返す（形式は中身で決める。読めない形式は空）
    param (
        [byte[]]$bytes,
        [hashtable]$state = $null,
        [System.Collections.Generic.List[string]]$failures = $null,
        [System.Collections.Generic.List[object]]$sizeFailures = $null
    )

    $lines = New-Object System.Collections.Generic.List[string]
    # ZIP の先頭の印（PK 03 04）。ZIP でなければ（古い形式の OLE など）読まない
    if ($bytes.Length -lt 4 -or $bytes[0] -ne 0x50 -or $bytes[1] -ne 0x4B -or $bytes[2] -ne 3 -or $bytes[3] -ne 4) {
        return $lines.ToArray()
    }

    $memory = New-Object System.IO.MemoryStream (, $bytes)
    $inner = $null
    try {
        $inner = New-Object System.IO.Compression.ZipArchive ($memory, [System.IO.Compression.ZipArchiveMode]::Read, $false)
        if ($null -ne $inner.GetEntry("xl/workbook.xml")) {
            $lines.AddRange([string[]]@(readXlsxCellLines $inner $state))
        } else {
            $units = $null
            if ($null -ne $inner.GetEntry("word/document.xml")) {
                $units = readDocxUnitsFromZip $inner $false $failures $sizeFailures
            } elseif ($null -ne $inner.GetEntry("ppt/presentation.xml")) {
                $units = readPptxUnitsFromZip $inner $false $failures $sizeFailures
            }
            if ($null -ne $units) {
                foreach ($unitName in $units.Keys) {
                    foreach ($line in $units[$unitName]) {
                        addEmbeddedOutputChars $state ($line.Length + 1)
                        $lines.Add($line)
                    }
                }
            }
        }
    } finally {
        if ($null -ne $inner) { $inner.Dispose() }
        $memory.Dispose()
    }
    return $lines.ToArray()
}

function readEmbeddedObjectLines {
    # readXmlLines が集めた埋め込み（Kind が embed）の参照から、埋め込んだファイルの文字を返す。
    # 戻り値: @{ Number = 通し番号; Lines = 行の配列 }。読まない（参照の先が無い・埋め込んだファイルではない・
    # 読んだことのある部品・読めない）ときは $null。
    #
    # 読めないとき: 部品ごとの上限・出す量の上限（ZipSizeLimitException の LimitKind が Part・Output）・壊れた XML・壊れた ZIP など
    # 何かの例外は、その埋め込みだけを読まなかった扱いにし、$failures に部品の名前、大きさの上限なら $sizeFailures に例外を
    # 追加して続ける（ほかの埋め込み・親の文字は出す。埋め込み 1 つで親のファイルの取り込みを失敗にしない）。
    # 投げ直すのは、1 ファイルの合計の上限（LimitKind が Total。ファイル全体の上限）と、中止・停止・メモリ不足の例外
    param (
        [System.IO.Compression.ZipArchive]$zip,
        [hashtable]$rels,   # readRelationships の結果
        $object,
        [hashtable]$state,  # newEmbeddedState の結果
        [System.Collections.Generic.List[string]]$failures = $null,
        [System.Collections.Generic.List[object]]$sizeFailures = $null
    )

    $rel = $rels[[string]$object.RelId]
    if ($null -eq $rel -or $rel.Type -notlike "*/package") {
        return $null
    }
    if (-not $state.Seen.Add([string]$rel.Target)) {
        return $null
    }
    $number = $state.Next
    $state.Next = $number + 1
    $outputBefore = $state.OutputChars  # 読めなかった埋め込みが出しかけた分は、合計から戻す（あとの埋め込みを巻き込まない）

    try {
        $bytes = readZipEntryBytes $zip $rel.Target
        if ($null -eq $bytes) {
            return $null
        }
        $lines = @(readEmbeddedPackageLines $bytes $state $failures $sizeFailures)
    } catch {
        $exception = $_.Exception
        # .NET のメソッド（XmlReader.Read など）の例外は、MethodInvocationException に包まれて届く
        while ($exception -is [System.Management.Automation.MethodInvocationException] -and $null -ne $exception.InnerException) {
            $exception = $exception.InnerException
        }
        $isSize = ($exception -is [ZipSizeLimitException])
        if ($isSize -and $exception.LimitKind -eq "Total") {
            throw
        }
        if ($exception -is [System.OperationCanceledException] -or $exception -is [System.Management.Automation.PipelineStoppedException] -or
            $exception -is [System.OutOfMemoryException]) {
            throw
        }
        $state.OutputChars = $outputBefore
        if ($isSize -and $null -ne $sizeFailures) {
            $sizeFailures.Add($exception)
        }
        if ($null -ne $failures) {
            $failures.Add([string]$rel.Target)
        }
        return $null
    }
    return @{ Number = $number; Lines = $lines }
}
