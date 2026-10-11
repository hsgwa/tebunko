# PDF ファイルの対象の拡張子・文字の読み取り。PowerShell だけで読む（外の部品・Office を使わない）。
# バイト列 → ページごとの文字列（readPdfPages）は、画面にもファイルにも触らない判断層の関数にする。
# ファイルを開いて読む関数（readPdfFile）だけが状態層（ファイルに触る）。
# 読む範囲と上限は docs/design/indexing/pdf.md、細工した PDF から身を守ることは docs/safety/checks.md。
#   読む    : 相互参照の表・相互参照のストリーム（オブジェクトのストリーム）、圧縮なし・FlateDecode、
#             文字の命令 Tj・TJ・'・"、行送り Td・TD・T*・Tm・TL（縦の位置が変わったときに行を分ける。ET では分けない）、
#             フォントの ToUnicode、ToUnicode の無い単純フォント（WinAnsi 相当）
#   暗号化   : 空の利用者パスワードで開ける PDF は復号して読む（pdf_crypt.ps1）。パスワードが要るもの・未対応の方式は例外にする。
#   読まない: FlateDecode 以外の圧縮、ToUnicode の無い複合フォント（CID）、壊れた相互参照（修復しない）。例外にする。
#             Form XObject（Do）・注釈・フォーム・しおりの文字は例外にせず、読み飛ばす
# 例外の文言は「PDF」で始める（readPdfPages は、それ以外の例外を「壊れている」に言い換える）。

# 取り込み対象にする PDF の拡張子（小文字）。足す・減らすと tests\meta\safety.Tests.ps1 が固定の一覧で確かめる
${pdfExtensions} = @(".pdf")

# 元のファイルがこの大きさ（バイト）を超えたら、読まずに失敗にする
${pdfFileMaxBytes} = 100MB
# 1 ファイルの解凍後の合計がこの大きさ（バイト）を超えたら、失敗にする（圧縮爆弾）
${pdfInflatedMaxBytes} = 100MB
# 相互参照のオブジェクトの数の上限
${pdfObjectMaxCount} = 1000000
# ページツリー・間接参照・配列や辞書の入れ子の深さの上限
${pdfNestMaxDepth} = 64
# ページ数の上限
${pdfPageMaxCount} = 10000
# TJ の数値がこの値以下（1/1000 文字幅。負の向きが右への詰め）で、前後の文字がどちらも半角のときだけ空白を入れる
${pdfGapThreshold} = -250
# 文字の対応表（ToUnicode）の項目の、1 ファイルの合計の上限
${pdfMapMaxEntries} = 1000000
# たどる相互参照（Prev・XRefStm）の数の上限
${pdfXrefMaxCount} = 1000
# ストリームの終わりを探して読み進める文字数の合計の上限（ファイルの大きさの何倍まで）。
# /Length の合わないストリームが順や逆順に並ぶだけなら、探す量は約 1 倍に収まる。その余裕を見て 4 倍にする
${pdfSearchMaxRatio} = 4
# 同じ行とみなす縦位置の差（ポイント）の上限
${pdfLineTolerance} = 1.0

${pdfLatin1} = [System.Text.Encoding]::GetEncoding(28591)
${pdfWinAnsi} = [System.Text.Encoding]::GetEncoding(1252)
${pdfDelimiters} = [char[]]@(0, 9, 10, 12, 13, 32, '(', ')', '<', '>', '[', ']', '{', '}', '/', '%')
${pdfStringStops} = [char[]]@('(', ')', '\')
# readPdfToken の速い道。空白（コード 32 以下）を飛ばして、(1) 数値・キーワード、または (2) 名前・入れ子もエスケープも無い文字列・<< >> [ ] { } を 1 つ取る。
# 区切りは ${pdfDelimiters} と同じ。コメント・16 進の文字列・入れ子やエスケープのある文字列・単独の < > ) は取らない（1 文字ずつ進める道に任せる）
# ページの内容で、文字の位置や文字列に効く命令。これ以外の命令（図形・色など）は、数と文字列を捨てて読み飛ばす
${pdfTextOps} = [System.Collections.Generic.HashSet[string]]::new([string[]]@("Tj", "TJ", "'", '"', "Tf", "BT", "Td", "TD", "TL", "T*", "Tm", "BI"), [System.StringComparer]::Ordinal)
${pdfTokenRegex} =[regex]::new('\G(?>[\x00-\x20]*)(?:([^\x00-\x20()<>\[\]{}/%][^\x00\x09\x0A\x0C\x0D\x20()<>\[\]{}/%]*)|(/[^\x00\x09\x0A\x0C\x0D\x20()<>\[\]{}/%]*|\([^()\\]*\)|<<|>>|[\[\]{}]))', [System.Text.RegularExpressions.RegexOptions]::Compiled)

function testPdfExtension {
    # ファイル名（またはパス）の拡張子が、取り込み対象の PDF の拡張子か（大文字・小文字を区別しない）
    param (
        [string]$path
    )

    $extension = [System.IO.Path]::GetExtension($path).ToLowerInvariant()
    return (${pdfExtensions} -contains $extension)
}

function newPdfLimits {
    # 読み取りの上限。既定は上の値。テストで小さく差し替えるための引数
    param (
        [long]$maxInflated = ${pdfInflatedMaxBytes},
        [int]$maxObjects = ${pdfObjectMaxCount},
        [int]$maxDepth = ${pdfNestMaxDepth},
        [int]$maxPages = ${pdfPageMaxCount},
        [int]$maxMapEntries = ${pdfMapMaxEntries},
        [int]$maxXref = ${pdfXrefMaxCount},
        [double]$maxSearchRatio = ${pdfSearchMaxRatio}
    )

    return @{
        MaxInflated = $maxInflated; MaxObjects = $maxObjects; MaxDepth = $maxDepth; MaxPages = $maxPages
        MaxMapEntries = $maxMapEntries; MaxXref = $maxXref; MaxSearchRatio = $maxSearchRatio
    }
}

function newPdfLexer {
    param (
        [string]$text,
        [int]$position = 0
    )

    return @{ S = $text; P = $position; K = "" }
}

function readPdfToken {
    # 字句を 1 つ読んで値を返す。種類は $st.K に入る:
    #   num（数値。double）/ name（先頭の / を除いた名前）/ str（文字列。丸括弧・16 進。1 バイト 1 文字の文字列）/
    #   kw（キーワード・演算子）/ [ / ] / << / >> / eof
    # 文字列全体に正規表現を当てず、1 文字ずつ進める（壊れた入力で後戻りが増えて固まらないため）
    param (
        [hashtable]$st
    )

    $s = $st.S
    $n = $s.Length
    $p = $st.P
    # 速い道: 空白のあとに、数値・キーワード・名前・入れ子もエスケープも無い文字列・<< >> [ ] { } のどれかが続くときは、
    # 位置指定（\G）の正規表現 1 回で切り出す（後戻りしない形。取れなければ、下の 1 文字ずつ進める道に任せる）
    $m = if ($p -lt $n) { ${pdfTokenRegex}.Match($s, $p) } else { [System.Text.RegularExpressions.Match]::Empty }
    if ($m.Success) {
        $g = $m.Groups[1]
        if ($g.Success) {
            $token = $g.Value
            $st.P = $g.Index + $g.Length
            $c = $token[0]
            if (($c -ge '0' -and $c -le '9') -or $c -eq '-' -or $c -eq '+' -or $c -eq '.') {
                $d = 0.0
                if ([double]::TryParse($token, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d) -and -not [double]::IsInfinity($d) -and -not [double]::IsNaN($d)) {
                    $st.K = "num"
                    return $d
                }
            }
            $st.K = "kw"
            return $token
        }
        $g = $m.Groups[2]
        $token = $g.Value
        $st.P = $g.Index + $g.Length
        $c = $token[0]
        if ($c -eq '/') {
            $name = $token.Substring(1)
            if ($name.Contains("#")) {
                $name = [regex]::Replace($name, '#([0-9A-Fa-f]{2})', { param($m) [string][char][Convert]::ToInt32($m.Groups[1].Value, 16) })
            }
            $st.K = "name"
            return $name
        }
        if ($c -eq '(') { $st.K = "str"; return $token.Substring(1, $token.Length - 2) }
        if ($c -eq '[') { $st.K = "["; return $null }
        if ($c -eq ']') { $st.K = "]"; return $null }
        if ($c -eq '<') { $st.K = "<<"; return $null }
        if ($c -eq '>') { $st.K = ">>"; return $null }
        $st.K = "kw"
        return $token
    }
    while ($true) {
        while ($p -lt $n -and [int]$s[$p] -le 32) { $p++ }
        if ($p -ge $n) {
            $st.P = $p
            $st.K = "eof"
            return $null
        }
        $c = $s[$p]
        if ($c -eq '%') {
            $e = $s.IndexOfAny([char[]]@([char]10, [char]13), $p)
            if ($e -lt 0) { $p = $n } else { $p = $e }
            continue
        }
        if ($c -eq '/') {
            $e = $s.IndexOfAny(${pdfDelimiters}, $p + 1)
            if ($e -lt 0) { $e = $n }
            $name = $s.Substring($p + 1, $e - $p - 1)
            if ($name.Contains("#")) {
                $name = [regex]::Replace($name, '#([0-9A-Fa-f]{2})', { param($m) [string][char][Convert]::ToInt32($m.Groups[1].Value, 16) })
            }
            $st.P = $e
            $st.K = "name"
            return $name
        }
        if ($c -eq '(') {
            $sb = [System.Text.StringBuilder]::new()
            $nest = 1
            $p++
            while ($p -lt $n) {
                $q = $s.IndexOfAny(${pdfStringStops}, $p)
                if ($q -lt 0) {
                    [void]$sb.Append($s, $p, $n - $p)
                    $p = $n
                    break
                }
                [void]$sb.Append($s, $p, $q - $p)
                $ch = $s[$q]
                if ($ch -eq '(') {
                    $nest++
                    [void]$sb.Append('(')
                    $p = $q + 1
                } elseif ($ch -eq ')') {
                    $nest--
                    $p = $q + 1
                    if ($nest -eq 0) { break }
                    [void]$sb.Append(')')
                } else {
                    $q++
                    if ($q -ge $n) { $p = $n; break }
                    $e = $s[$q]
                    $p = $q + 1
                    if ($e -ceq 'n') { [void]$sb.Append([char]10) }
                    elseif ($e -ceq 'r') { [void]$sb.Append([char]13) }
                    elseif ($e -ceq 't') { [void]$sb.Append([char]9) }
                    elseif ($e -ceq 'b') { [void]$sb.Append([char]8) }
                    elseif ($e -ceq 'f') { [void]$sb.Append([char]12) }
                    elseif ($e -ge '0' -and $e -le '7') {
                        $value = [int]$e - 48
                        $digits = 1
                        while ($digits -lt 3 -and $p -lt $n -and $s[$p] -ge '0' -and $s[$p] -le '7') {
                            $value = $value * 8 + ([int]$s[$p] - 48)
                            $p++
                            $digits++
                        }
                        [void]$sb.Append([char]($value -band 255))
                    } elseif ($e -eq "`r") {
                        if ($p -lt $n -and $s[$p] -eq "`n") { $p++ }
                    } elseif ($e -eq "`n") {
                        # 行の継続（何も入れない）
                    } else {
                        [void]$sb.Append($e)
                    }
                }
            }
            $st.P = $p
            $st.K = "str"
            return $sb.ToString()
        }
        if ($c -eq '<') {
            if ($p + 1 -lt $n -and $s[$p + 1] -eq '<') {
                $st.P = $p + 2
                $st.K = "<<"
                return $null
            }
            $e = $s.IndexOf('>', $p + 1)
            if ($e -lt 0) { $e = $n }
            $hex = $s.Substring($p + 1, $e - $p - 1) -replace '[^0-9A-Fa-f]', ''
            if ($hex.Length % 2 -eq 1) { $hex += "0" }
            $sb = [System.Text.StringBuilder]::new()
            for ($i = 0; $i -lt $hex.Length; $i += 2) {
                [void]$sb.Append([char][Convert]::ToInt32($hex.Substring($i, 2), 16))
            }
            $st.P = [Math]::Min($e + 1, $n)
            $st.K = "str"
            return $sb.ToString()
        }
        if ($c -eq '>') {
            if ($p + 1 -lt $n -and $s[$p + 1] -eq '>') {
                $st.P = $p + 2
                $st.K = ">>"
                return $null
            }
            $p++
            continue
        }
        if ($c -eq '[') { $st.P = $p + 1; $st.K = "["; return $null }
        if ($c -eq ']') { $st.P = $p + 1; $st.K = "]"; return $null }
        if ($c -eq ')' ) { $p++; continue }
        if ($c -eq '{' -or $c -eq '}') { $st.P = $p + 1; $st.K = "kw"; return [string]$c }

        $e = $s.IndexOfAny(${pdfDelimiters}, $p)
        if ($e -lt 0) { $e = $n }
        $token = $s.Substring($p, $e - $p)
        $st.P = $e
        $d = 0.0
        if (($c -ge '0' -and $c -le '9') -or $c -eq '-' -or $c -eq '+' -or $c -eq '.') {
            if ([double]::TryParse($token, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d) -and -not [double]::IsInfinity($d) -and -not [double]::IsNaN($d)) {
                $st.K = "num"
                return $d
            }
        }
        $st.K = "kw"
        return $token
    }
}

function readPdfValue {
    # 値（数値・名前 "/名前"・文字列 @{PdfType="str"}・配列 List・辞書 Hashtable・間接参照 @{PdfType="ref"; N}・真偽・null）を 1 つ読む。
    # 読み終わったあと $st.K が ] / >> / eof なら、値でなく終わりの印を読んだ（戻り値は $null）
    param (
        [hashtable]$doc,
        [hashtable]$st,
        [int]$depth,
        [bool]$allowRef = $true
    )

    if ($depth -gt $doc.Limits.MaxDepth) {
        throw "PDFの入れ子が深すぎるため読めません。"
    }
    $v = readPdfToken $st
    $kind = $st.K
    if ($kind -eq "num") {
        if ($allowRef -and $v -ge 0 -and $v -le 2147483647 -and $v -eq [Math]::Floor($v)) {
            $save = $st.P
            $g = readPdfToken $st
            if ($st.K -eq "num" -and $g -ge 0 -and $g -eq [Math]::Floor($g)) {
                $r = readPdfToken $st
                if ($st.K -eq "kw" -and $r -ceq "R") {
                    $st.K = "val"
                    return @{ PdfType = "ref"; N = [int]$v }
                }
            }
            $st.P = $save
            $st.K = "num"
        }
        return $v
    }
    if ($kind -eq "name") { return "/$v" }
    if ($kind -eq "str") { return @{ PdfType = "str"; V = $v } }
    if ($kind -eq "[") {
        $list = [System.Collections.Generic.List[object]]::new()
        while ($true) {
            $item = readPdfValue $doc $st ($depth + 1) $allowRef
            if ($st.K -eq "]" -or $st.K -eq "eof" -or $st.K -eq ">>") { break }
            if ($list.Count -ge $doc.Limits.MaxObjects) {
                throw "PDFの配列の要素の数が上限を超えるため読めません。"
            }
            $list.Add($item)
        }
        $st.K = "val"
        return , $list
    }
    if ($kind -eq "<<") {
        $dict = [System.Collections.Hashtable]::new()
        while ($true) {
            $key = readPdfToken $st
            if ($st.K -eq ">>" -or $st.K -eq "eof") { break }
            if ($st.K -ne "name") { continue }
            $item = readPdfValue $doc $st ($depth + 1) $allowRef
            if ($st.K -eq ">>" -or $st.K -eq "eof") { break }
            $dict[$key] = $item
            if ($dict.Count -gt $doc.Limits.MaxObjects) {
                throw "PDFの辞書の項目の数が上限を超えるため読めません。"
            }
        }
        $st.K = "val"
        return , $dict
    }
    if ($kind -eq "kw") {
        if ($v -ceq "true") { return $true }
        if ($v -ceq "false") { return $false }
        return $null
    }
    return $null
}

function testPdfRef {
    param ($v)
    return ($v -is [hashtable] -and $v.ContainsKey("PdfType") -and $v["PdfType"] -eq "ref")
}

function testPdfStream {
    param ($v)
    return ($v -is [hashtable] -and $v.ContainsKey("PdfType") -and $v["PdfType"] -eq "stream")
}

function testPdfDict {
    param ($v)
    return ($v -is [System.Collections.Hashtable] -and -not $v.ContainsKey("PdfType"))
}

function resolvePdf {
    # 間接参照をたどって値を返す（たどる回数に上限。循環は例外）
    param (
        [hashtable]$doc,
        $v
    )

    $count = 0
    while (testPdfRef $v) {
        $count++
        if ($count -gt $doc.Limits.MaxDepth) {
            throw "PDFの参照が循環しているため読めません。"
        }
        $v = getPdfObject $doc $v["N"]
    }
    return $v
}

function getPdfObject {
    # 番号のオブジェクトを読んで返す（読んだものは控える）。相互参照に無い・解放済みは $null。
    # 読んでいる最中に同じ番号を読もうとしたら（循環）、例外にする
    param (
        [hashtable]$doc,
        [int]$num
    )

    if ($doc.Cache.ContainsKey($num)) { return $doc.Cache[$num] }
    if (-not $doc.Xref.ContainsKey($num)) { return $null }
    $entry = $doc.Xref[$num]
    if ($entry.Type -eq 0) { return $null }
    if ($doc.Loading.ContainsKey($num)) {
        throw "PDFの参照が循環しているため読めません。"
    }
    if ($doc.Loading.Count -ge $doc.Limits.MaxDepth) {
        throw "PDFの参照が深すぎるため読めません。"
    }
    $doc.Loading[$num] = $true
    try {
        if ($entry.Type -eq 1) {
            $value = readPdfIndirectAt $doc $entry.Offset $num
        } else {
            $value = readPdfFromObjectStream $doc $entry.Stream $entry.Index $num
        }
    } finally {
        $doc.Loading.Remove($num)
    }
    $doc.Cache[$num] = $value
    return $value
}

function readPdfIndirectAt {
    # オフセットにある "N G obj ... endobj" を読む。ストリームなら @{PdfType="stream"; Dict; Start; Length} を返す。
    # expected が 0 以上なら、番号が合うことも確かめる
    param (
        [hashtable]$doc,
        [long]$offset,
        [int]$expected
    )

    $text = $doc.Text
    if ($offset -lt 0 -or $offset -ge $text.Length) {
        throw "PDFの相互参照がファイルの外を指しているため読めません。"
    }
    $st = newPdfLexer $text ([int]$offset)
    $a = readPdfToken $st
    $aKind = $st.K
    $gen = readPdfToken $st
    $c = readPdfToken $st
    if ($aKind -ne "num" -or $st.K -ne "kw" -or $c -cne "obj") {
        throw "PDFのオブジェクトが見つからないため読めません。"
    }
    if ($expected -ge 0 -and [int]$a -ne $expected) {
        throw "PDFのオブジェクトの番号が相互参照と合わないため読めません。"
    }
    $value = readPdfValue $doc $st 0
    if (-not (testPdfDict $value)) { return $value }

    $save = $st.P
    $t = readPdfToken $st
    if ($st.K -ne "kw" -or $t -cne "stream") { return $value }
    $start = $st.P
    if ($start -lt $text.Length -and $text[$start] -eq "`r") { $start++ }
    if ($start -lt $text.Length -and $text[$start] -eq "`n") { $start++ }
    $length = -1
    $declared = resolvePdf $doc $value["Length"]
    if ($declared -is [double] -and $declared -ge 0 -and $start + $declared -le $text.Length) {
        $after = [int]($start + $declared)
        $skip = 0
        while ($skip -lt 4 -and $after + $skip -lt $text.Length -and [int]$text[$after + $skip] -le 32) { $skip++ }
        if ([string]::CompareOrdinal($text, $after + $skip, "endstream", 0, 9) -eq 0) { $length = [int]$declared }
    }
    if ($length -lt 0) {
        # 長さが合わないストリームが多くても、同じ範囲を何度も探さないよう、直前に見つけた位置を使い回す
        # [From, At) には "endstream" の始まりが無い。前より手前から探すときは、手前の分だけ探して無ければ At を使う
        $found = $doc.EndSearch
        if ($start -ge $found.From -and $start -le $found.At) {
            $end = $found.At
        } elseif ($start -lt $found.From -and $found.At -ge $found.From) {
            $span = [Math]::Min($found.From - $start + 8, $text.Length - $start)
            $end = $text.IndexOf("endstream", $start, $span, [StringComparison]::Ordinal)
            if ($end -lt 0 -or $end -ge $found.From) {
                $doc.Searched += $span
                $end = $found.At
            } else {
                $doc.Searched += $end - $start
            }
            $found.From = $start
            $found.At = $end
        } else {
            $end = $text.IndexOf("endstream", $start, [StringComparison]::Ordinal)
            if ($end -lt 0) {
                throw "PDFのストリームの終わりが無いため読めません。"
            }
            $doc.Searched += $end - $start
            $found.From = $start
            $found.At = $end
        }
        if ($doc.Searched -gt $doc.Limits.MaxSearchRatio * $text.Length) {
            throw "PDFのストリームの終わりを探す量が上限を超えるため読めません。"
        }
        $length = $end - $start
        if ($length -gt 0 -and $text[$start + $length - 1] -eq "`n") { $length-- }
        if ($length -gt 0 -and $text[$start + $length - 1] -eq "`r") { $length-- }
    }
    return @{ PdfType = "stream"; Dict = $value; Start = $start; Length = $length; Num = [int]$a }
}

function expandPdfFlate {
    # zlib の先頭 2 バイトを飛ばして DeflateStream で解凍する。解凍後の合計が上限を超えたら例外
    param (
        [hashtable]$doc,
        [byte[]]$data
    )

    if ($data.Length -lt 3) {
        throw "PDFの圧縮されたデータが壊れているため読めません。"
    }
    $source = [System.IO.MemoryStream]::new($data, 2, $data.Length - 2)
    $deflate = [System.IO.Compression.DeflateStream]::new($source, [System.IO.Compression.CompressionMode]::Decompress)
    $output = [System.IO.MemoryStream]::new()
    $buffer = New-Object byte[] 65536
    try {
        while ($true) {
            $read = $deflate.Read($buffer, 0, $buffer.Length)
            if ($read -le 0) { break }
            $doc.Inflated += $read
            if ($doc.Inflated -gt $doc.Limits.MaxInflated) {
                throw "PDFの解凍後の大きさが上限を超えるため読めません。"
            }
            $output.Write($buffer, 0, $read)
        }
        return , $output.ToArray()
    } finally {
        $deflate.Dispose()
        $source.Dispose()
        $output.Dispose()
    }
}

function undoPdfPredictor {
    # FlateDecode の Predictor（PNG の 10〜15）を戻す。TIFF（2）は読まない
    param (
        [byte[]]$data,
        $params
    )

    $predictor = 1
    if ((testPdfDict $params) -and $params.ContainsKey("Predictor")) { $predictor = [int]$params["Predictor"] }
    if ($predictor -le 1) { return , $data }
    if ($predictor -lt 10) {
        throw "PDFの未対応の圧縮のため読めません。"
    }
    $columns = 1
    $colors = 1
    $bits = 8
    if ($params.ContainsKey("Columns")) { $columns = [int]$params["Columns"] }
    if ($params.ContainsKey("Colors")) { $colors = [int]$params["Colors"] }
    if ($params.ContainsKey("BitsPerComponent")) { $bits = [int]$params["BitsPerComponent"] }
    $rowLength = [int][Math]::Ceiling($columns * $colors * $bits / 8.0)
    $bpp = [Math]::Max(1, [int][Math]::Ceiling($colors * $bits / 8.0))
    if ($rowLength -le 0) {
        throw "PDFの圧縮の指定が壊れているため読めません。"
    }
    $rows = [int][Math]::Floor($data.Length / ($rowLength + 1))
    $out = New-Object byte[] ($rows * $rowLength)
    for ($r = 0; $r -lt $rows; $r++) {
        $src = $r * ($rowLength + 1)
        $type = $data[$src]
        $dst = $r * $rowLength
        $prev = $dst - $rowLength
        for ($i = 0; $i -lt $rowLength; $i++) {
            $x = [int]$data[$src + 1 + $i]
            $left = 0
            if ($i -ge $bpp) { $left = [int]$out[$dst + $i - $bpp] }
            $up = 0
            if ($r -gt 0) { $up = [int]$out[$prev + $i] }
            $upLeft = 0
            if ($r -gt 0 -and $i -ge $bpp) { $upLeft = [int]$out[$prev + $i - $bpp] }
            if ($type -eq 0) { $v = $x }
            elseif ($type -eq 1) { $v = $x + $left }
            elseif ($type -eq 2) { $v = $x + $up }
            elseif ($type -eq 3) { $v = $x + [int][Math]::Floor(($left + $up) / 2) }
            elseif ($type -eq 4) {
                $pp = $left + $up - $upLeft
                $pa = [Math]::Abs($pp - $left)
                $pb = [Math]::Abs($pp - $up)
                $pc = [Math]::Abs($pp - $upLeft)
                if ($pa -le $pb -and $pa -le $pc) { $pred = $left } elseif ($pb -le $pc) { $pred = $up } else { $pred = $upLeft }
                $v = $x + $pred
            } else {
                throw "PDFの圧縮のデータが壊れているため読めません。"
            }
            $out[$dst + $i] = [byte]($v -band 255)
        }
    }
    return , $out
}

function getPdfStreamBytes {
    # ストリームの中身を、圧縮を戻して返す（圧縮なし・FlateDecode だけ。ほかは例外）
    param (
        [hashtable]$doc,
        [hashtable]$stream
    )

    $dict = $stream.Dict
    $filters = resolvePdf $doc $dict["Filter"]
    # 取り出す前の大きさも、圧縮の有無にかかわらず合計に数える（同じストリームを何度も参照されても増え続けないように）
    $doc.Inflated += $stream.Length
    if ($doc.Inflated -gt $doc.Limits.MaxInflated) {
        throw "PDFの解凍後の大きさが上限を超えるため読めません。"
    }
    $data = New-Object byte[] $stream.Length
    [Array]::Copy($doc.Bytes, $stream.Start, $data, 0, $stream.Length)
    if ($null -ne $doc.Crypt -and $stream.ContainsKey("Num")) {
        # 暗号化された PDF は、圧縮を戻す前に復号する
        $method = getPdfStreamCryptMethod $doc $stream
        $data = unprotectPdfData $doc.Crypt $method $data
    }
    if ($null -eq $filters) { return , $data }
    $filterList = @($filters)
    $paramsValue = resolvePdf $doc $dict["DecodeParms"]
    $paramsList = @($paramsValue)
    for ($i = 0; $i -lt $filterList.Count; $i++) {
        $name = resolvePdf $doc $filterList[$i]
        if ($name -ceq "/Crypt") {
            continue
        } elseif ($name -ceq "/FlateDecode" -or $name -ceq "/Fl") {
            $data = expandPdfFlate $doc $data
            $params = $null
            if ($i -lt $paramsList.Count) { $params = resolvePdf $doc $paramsList[$i] }
            $data = undoPdfPredictor $data $params
        } else {
            throw "PDFの未対応の圧縮のため読めません。"
        }
    }
    return , $data
}

function readPdfFromObjectStream {
    # オブジェクトのストリームの中の、番号のオブジェクトを読む
    param (
        [hashtable]$doc,
        [int]$streamNum,
        [int]$index,
        [int]$num
    )

    if (-not $doc.ObjStreams.ContainsKey($streamNum)) {
        $stream = getPdfObject $doc $streamNum
        if (-not (testPdfStream $stream)) {
            throw "PDFのオブジェクトのストリームが読めないため読めません。"
        }
        $bytes = getPdfStreamBytes $doc $stream
        $text = ${pdfLatin1}.GetString($bytes)
        $count = [int](resolvePdf $doc $stream.Dict["N"])
        $first = [int](resolvePdf $doc $stream.Dict["First"])
        if ($count -lt 0 -or $count -gt $doc.Limits.MaxObjects) {
            throw "PDFのオブジェクトの数が上限を超えるため読めません。"
        }
        # 見出しは 1 件に 4 文字（"1 0 "）以上要る。中身より多い件数は、確保する前に断る
        if ($count -gt [Math]::Floor($text.Length / 4)) {
            throw "PDFのオブジェクトのストリームの件数が中身と合わないため読めません。"
        }
        $st = newPdfLexer $text 0
        $numbers = New-Object int[] $count
        $offsets = New-Object int[] $count
        for ($i = 0; $i -lt $count; $i++) {
            $token = readPdfToken $st
            if ($st.K -ne "num") { throw "PDFのオブジェクトのストリームの見出しが読めないため読めません。" }
            $numbers[$i] = [int]$token
            $token = readPdfToken $st
            if ($st.K -ne "num") { throw "PDFのオブジェクトのストリームの見出しが読めないため読めません。" }
            $offsets[$i] = [int]$token
        }
        $doc.ObjStreams[$streamNum] = @{ Text = $text; First = $first; Numbers = $numbers; Offsets = $offsets }
    }
    $info = $doc.ObjStreams[$streamNum]
    $position = -1
    if ($index -ge 0 -and $index -lt $info.Numbers.Length -and $info.Numbers[$index] -eq $num) {
        $position = $index
    } else {
        # 位置は細工できるため、番号から位置への辞書を 1 度だけ作って引く（毎回先頭から探すと、件数の 2 乗の比較になる）
        if (-not $info.ContainsKey("Positions")) {
            $positions = @{}
            for ($i = 0; $i -lt $info.Numbers.Length; $i++) {
                if (-not $positions.ContainsKey($info.Numbers[$i])) { $positions[$info.Numbers[$i]] = $i }
            }
            $info.Positions = $positions
        }
        $position = if ($info.Positions.ContainsKey([int]$num)) { $info.Positions[[int]$num] } else { -1 }
    }
    if ($position -lt 0) { return $null }
    $st = newPdfLexer $info.Text ($info.First + $info.Offsets[$position])
    return (readPdfValue $doc $st 0)
}

function addPdfXrefEntry {
    param (
        [hashtable]$doc,
        [long]$num,
        [hashtable]$entry
    )

    # 同じ番号を繰り返しても増えないので、読んだ行の総数も数える
    $doc.XrefRows += 1
    if ($doc.XrefRows -gt $doc.Limits.MaxObjects) {
        throw "PDFのオブジェクトの数が上限を超えるため読めません。"
    }
    if ($num -lt 0 -or $num -gt 2147483647) { return }
    if (-not $doc.Xref.ContainsKey([int]$num)) {
        $doc.Xref[[int]$num] = $entry
        if ($doc.Xref.Count -gt $doc.Limits.MaxObjects) {
            throw "PDFのオブジェクトの数が上限を超えるため読めません。"
        }
    } elseif ($doc.Hybrid -and $doc.Xref[[int]$num].Type -eq 0 -and $entry.Type -ne 0) {
        $doc.Xref[[int]$num] = $entry
    }
}

function readPdfXref {
    # オフセットにある相互参照（表または相互参照のストリーム）を読み、前の相互参照（Prev・XRefStm）も順にたどる。
    # 新しいものが先に入る。暗号化（Encrypt）の辞書の場所と ID は doc に覚える。ファイルの外を指す・循環・数が上限を超えるのは例外。
    # 入れ子の呼び出しにせず、待ちの積み上げで読む（Prev が長く続いても呼び出しが深くならない）
    param (
        [hashtable]$doc,
        [long]$offset
    )

    $pending = [System.Collections.Generic.Stack[object]]::new()
    $pending.Push(@{ Offset = $offset; Hybrid = $false })
    while ($pending.Count -gt 0) {
        if ($doc.Visited.Count -ge $doc.Limits.MaxXref) {
            throw "PDFの相互参照の数が上限を超えるため読めません。"
        }
        $job = $pending.Pop()
        $next = readPdfXrefSection $doc $job.Offset $job.Hybrid
        # 先に XRefStm、次に Prev の順に読む（Prev を先に積む）
        if ($next.Prev -ge 0) { $pending.Push(@{ Offset = $next.Prev; Hybrid = $false }) }
        if ($next.Stm -ge 0) { $pending.Push(@{ Offset = $next.Stm; Hybrid = $true }) }
    }
}

function readPdfXrefSection {
    # 相互参照 1 つ分を読み、次に読む位置 @{ Stm; Prev }（無ければ -1）を返す。
    # hybrid（XRefStm で読むもの）は、表で「解放済み」とされた番号を上書きする（旧い読み手から隠すための作り）
    param (
        [hashtable]$doc,
        [long]$offset,
        [bool]$hybrid = $false
    )

    $doc.Hybrid = $hybrid

    $text = $doc.Text
    if ($offset -lt 0 -or $offset -ge $text.Length) {
        throw "PDFの相互参照がファイルの外を指しているため読めません。"
    }
    if ($doc.Visited.ContainsKey($offset)) {
        throw "PDFの相互参照が循環しているため読めません。"
    }
    $doc.Visited[$offset] = $true
    $p = [int]$offset
    while ($p -lt $text.Length -and [int]$text[$p] -le 32) { $p++ }
    $max = $doc.Limits.MaxObjects
    if ([string]::CompareOrdinal($text, $p, "xref", 0, 4) -eq 0) {
        $st = newPdfLexer $text ($p + 4)
        while ($true) {
            $a = readPdfToken $st
            if ($st.K -eq "kw" -and $a -ceq "trailer") { break }
            if ($st.K -ne "num") {
                throw "PDFの相互参照の表が読めないため読めません。"
            }
            $first = [long]$a
            $count = readPdfToken $st
            if ($st.K -ne "num" -or $count -lt 0 -or $count -gt $max) {
                throw "PDFの相互参照の表が読めない、またはオブジェクトの数が上限を超えるため読めません。"
            }
            $pp = $st.P
            for ($k = 0; $k -lt $count; $k++) {
                while ($pp -lt $text.Length -and [int]$text[$pp] -le 32) { $pp++ }
                $entryOffset = 0L
                $kind = ""
                if ($pp + 17 -lt $text.Length -and $text[$pp + 10] -eq " " -and $text[$pp + 16] -eq " " -and
                    [long]::TryParse($text.Substring($pp, 10), [ref]$entryOffset)) {
                    $kind = [string]$text[$pp + 17]
                    $pp += 18
                } else {
                    $st.P = $pp
                    $o = readPdfToken $st
                    $okO = ($st.K -eq "num")
                    [void](readPdfToken $st)
                    $t = readPdfToken $st
                    if (-not $okO -or $st.K -ne "kw") {
                        throw "PDFの相互参照の表が読めないため読めません。"
                    }
                    $entryOffset = [long]$o
                    $kind = [string]$t
                    $pp = $st.P
                }
                if ($kind -ceq "n" -and $entryOffset -gt 0) {
                    addPdfXrefEntry $doc ($first + $k) @{ Type = 1; Offset = $entryOffset }
                } else {
                    addPdfXrefEntry $doc ($first + $k) @{ Type = 0 }
                }
            }
            $st.P = $pp
        }
        $trailer = readPdfValue $doc $st 0
    } else {
        $stream = readPdfIndirectAt $doc $p -1
        if (-not (testPdfStream $stream) -or $stream.Dict["Type"] -cne "/XRef") {
            throw "PDFの相互参照が読めないため読めません。"
        }
        $trailer = $stream.Dict
        $data = getPdfStreamBytes $doc $stream
        $w = @(resolvePdf $doc $trailer["W"])
        if ($w.Count -ne 3) {
            throw "PDFの相互参照のストリームが壊れているため読めません。"
        }
        $w0 = [int]$w[0]
        $w1 = [int]$w[1]
        $w2 = [int]$w[2]
        if ($w0 -lt 0 -or $w1 -lt 0 -or $w2 -lt 0 -or $w0 -gt 8 -or $w1 -gt 8 -or $w2 -gt 8 -or ($w0 + $w1 + $w2) -eq 0) {
            throw "PDFの相互参照のストリームが壊れているため読めません。"
        }
        $rowLength = $w0 + $w1 + $w2
        $index = resolvePdf $doc $trailer["Index"]
        if ($null -eq $index) {
            $index = @(0, [int](resolvePdf $doc $trailer["Size"]))
        }
        $index = @($index)
        $pos = 0
        for ($g = 0; $g + 1 -lt $index.Count; $g += 2) {
            $first = [long]$index[$g]
            $count = [long]$index[$g + 1]
            if ($count -lt 0 -or $count -gt $max) {
                throw "PDFのオブジェクトの数が上限を超えるため読めません。"
            }
            for ($k = 0; $k -lt $count; $k++) {
                if ($pos + $rowLength -gt $data.Length) {
                    throw "PDFの相互参照のストリームが壊れているため読めません。"
                }
                $type = 1L
                if ($w0 -gt 0) {
                    $type = 0L
                    for ($j = 0; $j -lt $w0; $j++) { $type = $type * 256 + $data[$pos + $j] }
                }
                $f2 = 0L
                for ($j = 0; $j -lt $w1; $j++) { $f2 = $f2 * 256 + $data[$pos + $w0 + $j] }
                $f3 = 0L
                for ($j = 0; $j -lt $w2; $j++) { $f3 = $f3 * 256 + $data[$pos + $w0 + $w1 + $j] }
                $pos += $rowLength
                if ($type -eq 1) {
                    addPdfXrefEntry $doc ($first + $k) @{ Type = 1; Offset = $f2 }
                } elseif ($type -eq 2) {
                    addPdfXrefEntry $doc ($first + $k) @{ Type = 2; Stream = [int]$f2; Index = [int]$f3 }
                } else {
                    addPdfXrefEntry $doc ($first + $k) @{ Type = 0 }
                }
            }
        }
    }

    if (-not (testPdfDict $trailer)) {
        throw "PDFの相互参照のあとの情報が読めないため読めません。"
    }
    if ($trailer.ContainsKey("Encrypt") -and $null -eq $doc.EncryptEntry) {
        # 新しい方のトレーラーを使う。辞書は、相互参照がそろってから読む（readPdfPagesCore）
        $doc.EncryptEntry = $trailer["Encrypt"]
    }
    if ($null -eq $doc.Root -and $trailer.ContainsKey("Root")) {
        $doc.Root = $trailer["Root"]
    }
    $next = @{ Stm = -1L; Prev = -1L }
    if ($trailer.ContainsKey("XRefStm")) { $next.Stm = [long]$trailer["XRefStm"] }
    if ($trailer.ContainsKey("Prev")) { $next.Prev = [long]$trailer["Prev"] }
    return $next
}

function addPdfMapEntries {
    # 文字の対応表の項目を数え、1 ファイルの合計が上限を超えたら例外にする（doc が無いときは数えない）
    param (
        $doc,
        [long]$count
    )

    if ($null -eq $doc) { return }
    $doc.MapEntries += $count
    if ($doc.MapEntries -gt $doc.Limits.MaxMapEntries) {
        throw "PDFの文字の対応表が大きすぎるため読めません。"
    }
}

function readPdfToUnicode {
    # ToUnicode の CMap から、文字コード → 文字列の表 Map（キーは int）と、文字コードのバイト数 Width（0 は指定なし）を返す
    param (
        [string]$text,
        [hashtable]$doc = $null
    )

    $map = @{}
    $width = 0
    $st = newPdfLexer $text 0
    $utf16 = [System.Text.Encoding]::BigEndianUnicode
    while ($true) {
        $t = readPdfToken $st
        if ($st.K -eq "eof") { break }
        if ($st.K -ne "kw") { continue }
        if ($t -ceq "begincodespacerange") {
            while ($true) {
                $lo = readPdfToken $st
                if ($st.K -ne "str") { break }
                [void](readPdfToken $st)
                if ($width -eq 0 -and $lo.Length -ge 1 -and $lo.Length -le 2) { $width = $lo.Length }
            }
        } elseif ($t -ceq "beginbfchar") {
            while ($true) {
                $src = readPdfToken $st
                if ($st.K -ne "str") { break }
                $dst = readPdfToken $st
                if ($st.K -ne "str") { break }
                if ($src.Length -eq 0 -or $src.Length -gt 3) { continue }
                addPdfMapEntries $doc 1
                $map[(convertPdfCode $src)] = convertPdfUtf16 $dst $utf16
            }
        } elseif ($t -ceq "beginbfrange") {
            while ($true) {
                $lo = readPdfToken $st
                if ($st.K -ne "str") { break }
                $hi = readPdfToken $st
                if ($st.K -ne "str") { break }
                $dst = readPdfToken $st
                $kind = $st.K
                if ($lo.Length -eq 0 -or $lo.Length -gt 3 -or $hi.Length -eq 0 -or $hi.Length -gt 3) {
                    if ($kind -eq "[") {
                        while ($true) { [void](readPdfToken $st); if ($st.K -eq "]" -or $st.K -eq "eof") { break } }
                    }
                    continue
                }
                $from = convertPdfCode $lo
                $to = [Math]::Min((convertPdfCode $hi), $from + 65535)
                if ($kind -eq "str") {
                    $base = convertPdfUtf16 $dst $utf16
                    if ($base.Length -eq 0) { continue }
                    if ($to -ge $from) { addPdfMapEntries $doc ($to - $from + 1) }
                    $head = $base.Substring(0, $base.Length - 1)
                    $lastCode = [int]$base[$base.Length - 1]
                    for ($c = $from; $c -le $to; $c++) {
                        $map[[int]$c] = $head + [char](($lastCode + $c - $from) -band 65535)
                    }
                } elseif ($kind -eq "[") {
                    $c = $from
                    while ($true) {
                        $item = readPdfToken $st
                        if ($st.K -ne "str") { break }
                        if ($c -le $to) {
                            addPdfMapEntries $doc 1
                            $map[[int]$c] = convertPdfUtf16 $item $utf16
                        }
                        $c++
                    }
                }
            }
        }
    }
    return @{ Map = $map; Width = $width }
}

function convertPdfCode {
    # 1 バイト 1 文字の文字列を、ビッグエンディアンの文字コード（int）にする（3 バイトまで）
    param (
        [string]$s
    )

    $code = 0
    foreach ($ch in $s.ToCharArray()) { $code = $code * 256 + [int]$ch }
    return [int]$code
}

function convertPdfUtf16 {
    # 1 バイト 1 文字の文字列（UTF-16BE のバイト列）を文字列にする。奇数バイトのときは 1 バイトずつ文字にする
    param (
        [string]$s,
        $utf16
    )

    if ($s.Length % 2 -eq 1) { return $s }
    return $utf16.GetString(${pdfLatin1}.GetBytes($s))
}

function newPdfFont {
    # フォントの辞書から、文字列を文字にする規則 @{ Width; Map; NoMap } を作る。
    # ToUnicode があればそれを使う。無い単純フォントは WinAnsi 相当（Windows-1252）。無い複合フォント（CID）は NoMap（使ったときに例外）
    param (
        [hashtable]$doc,
        $dict
    )

    $font = @{ Width = 1; Map = $null; NoMap = $false }
    if (-not (testPdfDict $dict)) { return $font }
    $isComposite = ($dict["Subtype"] -ceq "/Type0")
    if ($isComposite) { $font.Width = 2 }
    $toUnicodeRef = $dict["ToUnicode"]
    $toUnicode = resolvePdf $doc $toUnicodeRef
    if (testPdfStream $toUnicode) {
        # 同じ ToUnicode を使うフォントが多くても、1 度だけ読む（番号ごとに控える）
        $key = $null
        if (testPdfRef $toUnicodeRef) { $key = $toUnicodeRef["N"] }
        if ($null -ne $key -and $doc.CMaps.ContainsKey($key)) {
            $cmap = $doc.CMaps[$key]
        } else {
            $bytes = getPdfStreamBytes $doc $toUnicode
            $cmap = readPdfToUnicode (${pdfLatin1}.GetString($bytes)) $doc
            if ($null -ne $key) { $doc.CMaps[$key] = $cmap }
        }
        $font.Map = $cmap.Map
        if ($cmap.Width -gt 0) { $font.Width = $cmap.Width }
    } elseif ($isComposite) {
        $font.NoMap = $true
    }
    return $font
}

function convertPdfStringToText {
    # 文字列（1 バイト 1 文字）を、フォントの規則で文字列にする
    param (
        $font,
        [string]$s
    )

    if ($null -ne $font -and $font.NoMap) {
        throw "PDFの文字の対応表が無いため読めません。"
    }
    if ($null -eq $font -or $null -eq $font.Map) {
        return ${pdfWinAnsi}.GetString(${pdfLatin1}.GetBytes($s))
    }
    $map = $font.Map
    $width = $font.Width
    $sb = [System.Text.StringBuilder]::new()
    for ($i = 0; $i -lt $s.Length; $i += $width) {
        if ($width -eq 2) {
            if ($i + 1 -ge $s.Length) { break }
            $code = [int]$s[$i] * 256 + [int]$s[$i + 1]
        } else {
            $code = [int]$s[$i]
        }
        if ($map.ContainsKey($code)) {
            [void]$sb.Append($map[$code])
        } elseif ($width -eq 1) {
            [void]$sb.Append(${pdfWinAnsi}.GetString([byte[]]@($code)))
        }
    }
    return $sb.ToString()
}

function testPdfHalfWidth {
    # 半角の文字か（ラテン文字の範囲）。TJ の詰めに空白を入れるかの判断に使う
    param (
        [char]$c
    )

    return ([int]$c -le 0x24F)
}

function getPdfGlyphCount {
    # 文字列が表す字形（フォントの 1 コード）の数。1 コードが 2 文字になる対応（合字）でも 1 と数える
    param (
        $font,
        [string]$s
    )

    if ($null -eq $font -or $null -eq $font.Map -or $font.Width -ne 2) { return $s.Length }
    return [int][Math]::Ceiling($s.Length / 2)
}

function addPdfText {
    # ページの文字に文字列を足す。詰め（Gap）が付いていて、前後の文字がどちらも半角のときだけ空白を入れる
    param (
        [hashtable]$ts,
        [string]$text,
        [int]$glyphs = 1
    )

    if ($text.Length -eq 0) { return }
    # 縦位置が前に出した文字と違うときだけ行を分ける（BT〜ET で分かれていても、同じ縦位置なら同じ行）
    if ($ts.HasOut -and [Math]::Abs($ts.Y - $ts.OutY) -gt ${pdfLineTolerance}) {
        breakPdfLine $ts
    }
    $ts.HasOut = $true
    $ts.OutY = $ts.Y
    $ts.Run += $glyphs
    $sb = $ts.Sb
    if ($ts.Gap -and $sb.Length -gt 0) {
        $last = $sb[$sb.Length - 1]
        if ($last -ne "`n" -and $last -ne " " -and (testPdfHalfWidth $last) -and (testPdfHalfWidth $text[0]) -and $text[0] -ne " ") {
            [void]$sb.Append(" ")
        }
    }
    $ts.Gap = $false
    [void]$sb.Append($text)
}

function breakPdfLine {
    param (
        [hashtable]$ts
    )

    $sb = $ts.Sb
    if ($sb.Length -gt 0 -and $sb[$sb.Length - 1] -ne "`n") {
        [void]$sb.Append("`n")
    }
    $ts.Gap = $false
}

function readPdfPageText {
    # ページの内容（文字の命令）から文字列を作る。行送りは改行にする。
    # fontDict は フォント名 → フォント（辞書または参照）の辞書（無ければ $null）
    param (
        [hashtable]$doc,
        [string]$content,
        $fontDict
    )

    # Y は今の文字の縦位置（行の行列の y）、OutY は最後に出した文字の縦位置、Leading は行送り（TL）
    $ts = @{ Sb = [System.Text.StringBuilder]::new(); Gap = $false; Y = 0.0; OutY = 0.0; HasOut = $false; Leading = 0.0; Scale = 1.0; Run = 0 }
    $st = newPdfLexer $content 0
    $nums = [System.Collections.Generic.List[double]]::new()
    $lastStr = $null
    $lastName = $null
    $lastArray = $null
    $font = $null
    $halfGlyph = $false
    $fontSize = 0.0
    while ($true) {
        $t = readPdfToken $st
        $kind = $st.K
        if ($kind -eq "eof") { break }
        if ($kind -eq "num") {
            if ($nums.Count -ge 64) { $nums.Clear() }
            $nums.Add($t)
            continue
        }
        if ($kind -eq "name") { $lastName = $t; continue }
        if ($kind -eq "str") { $lastStr = $t; continue }
        if ($kind -eq "[") {
            $items = [System.Collections.Generic.List[object]]::new()
            while ($true) {
                $item = readPdfToken $st
                if ($st.K -eq "]" -or $st.K -eq "eof") { break }
                if ($st.K -eq "num" -or $st.K -eq "str") { $items.Add($item) }
            }
            $lastArray = $items
            continue
        }
        if ($kind -eq "<<") {
            $nest = 1
            while ($nest -gt 0) {
                [void](readPdfToken $st)
                if ($st.K -eq "<<") { $nest++ } elseif ($st.K -eq ">>") { $nest-- } elseif ($st.K -eq "eof") { break }
            }
            continue
        }
        if ($kind -ne "kw") { continue }
        if (-not ${pdfTextOps}.Contains($t)) {
            $nums.Clear()
            $lastStr = $null
            $lastName = $null
            $lastArray = $null
            continue
        }

        if ($t -ceq "Tj") {
            if ($null -ne $lastStr) { addPdfText $ts (convertPdfStringToText $font $lastStr) $(if ($halfGlyph) { ($lastStr.Length + 1) -shr 1 } else { $lastStr.Length }) }
        } elseif ($t -ceq "TJ") {
            if ($null -ne $lastArray) {
                foreach ($item in $lastArray) {
                    if ($item -is [string]) {
                        addPdfText $ts (convertPdfStringToText $font $item) $(if ($halfGlyph) { ($item.Length + 1) -shr 1 } else { $item.Length })
                    } elseif ($item -le ${pdfGapThreshold}) {
                        $ts.Gap = $true
                    }
                }
            }
        } elseif ($t -ceq "'" -or $t -ceq '"') {
            $ts.Y -= $ts.Leading * $ts.Scale
            $ts.Run = 0
            breakPdfLine $ts
            if ($null -ne $lastStr) { addPdfText $ts (convertPdfStringToText $font $lastStr) $(if ($halfGlyph) { ($lastStr.Length + 1) -shr 1 } else { $lastStr.Length }) }
        } elseif ($t -ceq "Tf") {
            $fontSize = if ($nums.Count -ge 1) { [Math]::Abs($nums[$nums.Count - 1]) } else { 0.0 }
            $font = $null
            if ($null -ne $lastName -and (testPdfDict $fontDict) -and $fontDict.ContainsKey($lastName)) {
                $font = getPdfFont $doc $fontDict[$lastName]
            }
            # 1 コードが 2 バイトの対応表のあるフォントは、2 文字を 1 字形と数える（getPdfGlyphCount と同じ）
            $halfGlyph = ($null -ne $font -and $null -ne $font.Map -and $font.Width -eq 2)
        } elseif ($t -ceq "BT") {
            # テキストの行列は BT で初期化される（縦位置は 0 から）
            $ts.Y = 0.0
            $ts.Scale = 1.0
            $ts.Run = 0
        } elseif ($t -ceq "Td" -or $t -ceq "TD") {
            if ($nums.Count -ge 2) {
                # 縦の動きは、Tm の拡大（d）を掛けた値にする（`12 0 0 12 x y Tm` のあとの `0 -1 TD` は 12 ポイント動く）
                $ts.Y += $nums[1] * $ts.Scale
                if ($t -ceq "TD") { $ts.Leading = - $nums[1] }
                # 横にだけ動いたときは詰め。縦に動いたときは、次の文字を出すときに行を分ける。
                # 1 文字ずつ Td で並べる PDF（Chromium など）の文字を離さないよう、直前の Td 以降に出した文字が 1 文字以下で、
                # 前向きに 1 em より小さい動きは詰めにしない（空白は文字として入っている）。2 文字以上の語の後ろは空白でつなぐ
                $joined = ($fontSize -gt 0 -and $nums[0] -gt 0 -and $nums[0] -lt $fontSize -and $ts.Run -le 1)
                if ($nums[0] -ne 0 -and -not $joined) { $ts.Gap = $true }
                $ts.Run = 0
            }
        } elseif ($t -ceq "TL") {
            if ($nums.Count -ge 1) { $ts.Leading = $nums[0] }
        } elseif ($t -ceq "T*") {
            $ts.Y -= $ts.Leading * $ts.Scale
            $ts.Run = 0
            breakPdfLine $ts
        } elseif ($t -ceq "Tm") {
            if ($nums.Count -ge 6) {
                $ts.Y = $nums[5]
                # 上下を反転した行列（d が負）は、縦の動きの向きも反転するので、d の符号を残す
                $ts.Scale = [Math]::Sqrt($nums[2] * $nums[2] + $nums[3] * $nums[3])
                if ($nums[3] -lt 0) { $ts.Scale = -$ts.Scale }
                $ts.Run = 0
                $ts.Gap = $true
            }
        } elseif ($t -ceq "BI") {
            # 画像の中身（バイナリ）は読まない。ID から EI までを飛ばす
            $idPos = $content.IndexOf("ID", $st.P, [StringComparison]::Ordinal)
            if ($idPos -lt 0) { break }
            $scan = $idPos + 2
            $endPos = -1
            while ($true) {
                $found = $content.IndexOf("EI", $scan, [StringComparison]::Ordinal)
                if ($found -lt 0) { break }
                $beforeOk = ($found -eq 0 -or [int]$content[$found - 1] -le 32)
                $afterOk = ($found + 2 -ge $content.Length -or [int]$content[$found + 2] -le 32)
                if ($beforeOk -and $afterOk) { $endPos = $found + 2; break }
                $scan = $found + 2
            }
            if ($endPos -lt 0) { break }
            $st.P = $endPos
        }
        $nums.Clear()
        $lastStr = $null
        $lastName = $null
        $lastArray = $null
    }
    return ([regex]::Replace($ts.Sb.ToString(), '[\x00-\x09\x0B-\x1F\x7F]', " ")).TrimEnd("`n")
}

function getPdfFont {
    # フォント（辞書または参照）の規則を返す。参照のものは番号ごとに控える
    param (
        [hashtable]$doc,
        $ref
    )

    if (testPdfRef $ref) {
        $num = $ref["N"]
        if (-not $doc.Fonts.ContainsKey($num)) {
            $doc.Fonts[$num] = newPdfFont $doc (resolvePdf $doc $ref)
        }
        return $doc.Fonts[$num]
    }
    return (newPdfFont $doc $ref)
}

function walkPdfPages {
    # ページツリーをたどり、ページ（@{ Dict; Resources }）を順に pages に足す。
    # 同じノードを 2 度たどる（循環）・深すぎる・ページが多すぎるときは例外
    param (
        [hashtable]$doc,
        $node,
        $inheritedResources,
        [int]$depth,
        $pages
    )

    if ($depth -gt $doc.Limits.MaxDepth) {
        throw "PDFのページの入れ子が深すぎるため読めません。"
    }
    if (testPdfRef $node) {
        if ($doc.SeenNodes.ContainsKey($node["N"])) {
            throw "PDFのページの並びが循環しているため読めません。"
        }
        $doc.SeenNodes[$node["N"]] = $true
    }
    $dict = resolvePdf $doc $node
    if (-not (testPdfDict $dict)) { return }
    $resources = $inheritedResources
    if ($dict.ContainsKey("Resources")) { $resources = $dict["Resources"] }
    if ($dict.ContainsKey("Kids")) {
        $kids = resolvePdf $doc $dict["Kids"]
        foreach ($kid in @($kids)) {
            walkPdfPages $doc $kid $resources ($depth + 1) $pages
        }
        return
    }
    if ($pages.Count -ge $doc.Limits.MaxPages) {
        throw "PDFのページ数が上限を超えるため読めません。"
    }
    $pages.Add(@{ Dict = $dict; Resources = $resources })
}

function readPdfPages {
    # PDF のバイト列から、ページごとの文字列（行は改行で区切る。文字の無いページは空）の配列を返す。
    # 読めない PDF（パスワードが要る・暗号化の方式が未対応・壊れている・上限を超える・読まない圧縮・文字の対応表が無い）は、「PDF」で始まる日本語の例外にする
    param (
        [byte[]]$bytes,
        [hashtable]$limits = $null
    )

    if ($null -eq $limits) { $limits = newPdfLimits }
    try {
        return , (readPdfPagesCore $bytes $limits)
    } catch {
        $message = $_.Exception.Message
        if ($message.StartsWith("PDF")) {
            throw $message
        }
        throw "PDFが壊れているため読めません。"
    }
}

function readPdfPagesCore {
    param (
        [byte[]]$bytes,
        [hashtable]$limits
    )

    $text = ${pdfLatin1}.GetString($bytes)
    if ($text.IndexOf("%PDF-", 0, [Math]::Min(1024, $text.Length), [StringComparison]::Ordinal) -lt 0) {
        throw "PDFではないか、壊れているため読めません。"
    }
    $doc = @{
        Bytes = $bytes; Text = $text; Limits = $limits
        Xref = @{}; Cache = @{}; Loading = @{}; ObjStreams = @{}; Fonts = @{}; Visited = @{}; SeenNodes = @{}
        Inflated = [long]0; XrefRows = [long]0; Searched = [long]0; Root = $null; EndSearch = @{ From = 1; At = 0 }; Hybrid = $false; CMaps = @{}; MapEntries = [long]0
        Crypt = $null; EncryptEntry = $null
    }
    $start = $text.LastIndexOf("startxref", [StringComparison]::Ordinal)
    if ($start -lt 0) {
        throw "PDFの相互参照が見つからないため読めません。"
    }
    $st = newPdfLexer $text ($start + 9)
    $offset = readPdfToken $st
    if ($st.K -ne "num") {
        throw "PDFの相互参照の位置が読めないため読めません。"
    }
    readPdfXref $doc ([long]$offset)
    if ($null -ne $doc.EncryptEntry) {
        $enc = resolvePdf $doc $doc.EncryptEntry
        if (-not (testPdfDict $enc)) {
            throw "PDFの暗号化の情報が壊れているため読めません。"
        }
        $doc.Crypt = newPdfCrypt $doc $enc
    }

    $root = resolvePdf $doc $doc.Root
    if (-not (testPdfDict $root)) {
        throw "PDFの文書の入り口が見つからないため読めません。"
    }
    $pages = [System.Collections.Generic.List[object]]::new()
    walkPdfPages $doc $root["Pages"] $null 0 $pages

    $result = New-Object string[] $pages.Count
    for ($i = 0; $i -lt $pages.Count; $i++) {
        $page = $pages[$i]
        $contents = resolvePdf $doc $page.Dict["Contents"]
        $buffer = [System.IO.MemoryStream]::new()
        foreach ($item in @($contents)) {
            $stream = resolvePdf $doc $item
            if (testPdfStream $stream) {
                $data = getPdfStreamBytes $doc $stream
                $buffer.Write($data, 0, $data.Length)
                $buffer.WriteByte(10)
            }
        }
        $resources = resolvePdf $doc $page.Resources
        $fontDict = $null
        if (testPdfDict $resources) { $fontDict = resolvePdf $doc $resources["Font"] }
        $result[$i] = readPdfPageText $doc (${pdfLatin1}.GetString($buffer.ToArray())) $fontDict
    }
    return , $result
}

function readPdfFile {
    # PDF ファイルを読み、ページごとの文字列の配列を返す（readPdfPages）。
    # 大きさが maxBytes を超えるときは読まずに例外にする
    param (
        [string]$path,
        [long]$maxBytes = ${pdfFileMaxBytes},
        [hashtable]$limits = $null
    )

    if ([System.IO.FileInfo]::new((toLongPath $path)).Length -gt $maxBytes) {
        throw "ファイルサイズが大きすぎるため更新できません。"
    }
    $bytes = [System.IO.File]::ReadAllBytes((toLongPath $path))
    return , (readPdfPages $bytes $limits)
}
