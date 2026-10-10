# Word（.docx / .docm）・PowerPoint（.pptx / .pptm）のファイルと、Excel（.xlsx / .xlsm）の図形・コメントからテキストを読み出す。
# ファイルはZIP（Office Open XML）として直接読むため、Word・PowerPoint・Excelは使わない。
# インデクサ・テストから dot-source して使う。共通の部品（shared.ps1）を先に読み込んでおくこと。
#
# 読み出した結果は「場所 → 行の一覧」の順序付き辞書（ユニット）で返す。
#   Word      : ページ001, ページ001[図形], ページ001[コメント], ページ001[埋め込み1], ページ002, ..., ヘッダー・フッター, 脚注
#   PowerPoint: スライド001, スライド001[図形], スライド001[コメント], スライド001[埋め込み1], スライド001_ノート, スライド002（非表示）, ..., ヘッダー・フッター
#   Excel     : <シート名>[図形]（グラフ・SmartArt の文字を含む。表示のグラフシートも含む）, <シート名>[コメント],
#              <シート名>[ヘッダー・フッター]（セルの値はインデクサが Excel で読む）
# 1行は段落1つ、または表の1行（セルをタブ区切り）。図形・コメントの場所は図形・コメント1つ
# （Word・PowerPoint は文字だけ、Excel は "<セル番地><TAB><文字>"）。
# 埋め込みの場所は、埋め込んだ Office のファイル（Office Open XML）1 つの中の文字をまとめたもの
# （office_embedded.ps1。[埋め込み<N>] の N は 1 ファイルの中の通し番号）。

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

${nsWord}    = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
${nsDrawing} = "http://schemas.openxmlformats.org/drawingml/2006/main"
${nsPresent} = "http://schemas.openxmlformats.org/presentationml/2006/main"
${nsCompat}  = "http://schemas.openxmlformats.org/markup-compatibility/2006"
${nsRel}     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
${nsPkgRel}  = "http://schemas.openxmlformats.org/package/2006/relationships"
${nsDiagram} = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
${nsChart}   = "http://schemas.openxmlformats.org/drawingml/2006/chart"
${nsOffice}  = "urn:schemas-microsoft-com:office:office"

# PowerPointで読み飛ばすプレースホルダー（スライド番号・日付・ヘッダー・フッター・スライド画像）
${skipPlaceholderTypes} = @("sldNum", "dt", "hdr", "ftr", "sldImg")

# ZIPの部品（ファイル）1つの展開後の大きさと、1ファイルの合計（readDocxUnits・readPptxUnits・
# readXlsxObjectUnitsがファイルを開くたびに0から数える）の上限。どちらもバイト（entry.Lengthと同じ単位）。
# ZIP爆弾（小さく圧縮した巨大な部品）・部品を大量に並べたファイルでメモリを使い切るのを防ぐ（readZipEntry）。
# setting.config には出さない定数（値を変えたい場合は docs/design/indexing/known-issues.md を参照）。
# $script: の変数は読み取りのスレッド（runspace）ごとに別になる。ファイルの合計はスレッドごとに独立して
# 数えたいので都合がよい（テストから一時的に値を変えるときは、この $script: を直接上書きする）
$script:zipPartMaxBytes = 100MB
$script:zipTotalMaxBytes = 300MB
# シートの部品を流れで読む（readXlsxSheetHeaderFooter）ときの、展開後の大きさ（バイト）の上限。流れ読みは文字列にしない
# のでメモリを使わず、かかるのは時間だけ。シートは sheetData が大きく、実在のブックでも部品の上限（100MB）を超えうる
# ため、部品の上限より大きい別の値にする（超えたシートは、そのシートのヘッダー・フッターだけ読めなかった扱い）。
# 1ファイルの合計（zipTotalMaxBytes）には数えないので、シートごとに最大この大きさまで読む
$script:zipSheetStreamMaxBytes = 1000MB
# 1ファイルの合計（バイト）。readZipEntry が確保する大きさ（entry.Length）を足していく
$script:zipTotalReadBytes = 0

# サイズの上限を超えた（ZIP爆弾・巨大なXML等の見込み）ときに投げる、利用者向けの短い文言。
# 画面には部品ごとか合計かの区別・部品名・大きさを出さない（分かると悪用のヒントになるため）。
# 原因を調べられるよう、部品名・大きさ・部品ごとか合計かはインデックス作成のログにだけ書く。
# shared/ はツール（インデックス作成のログ）を知らないため、ここでは書かず、
# ZipSizeLimitException のプロパティで呼び出し元（tebunko/indexer）に伝える
$script:zipTooLargeMessage = "ファイルサイズが大きすぎるため更新できません。"

class ZipSizeLimitException : System.Exception {
    # .Message は今までどおり利用者向けの簡潔な文言のまま（$script:zipTooLargeMessage）。
    # 部品名・大きさ・部品ごとか合計かは、ここのプロパティにだけ持ち、呼び出し元がインデックス作成のログに書く
    [string]$PartName     # ZIP内のパス（例: "word/document.xml"）
    [long]$MeasuredBytes  # 超えたと分かった時点の大きさ（バイト）。偽りのヘッダーのときは、実際の大きさまでは分からないため、検出できた時点の値
    [string]$LimitKind    # "Part"（部品ごとの上限。偽りのヘッダーを含む）または "Total"（1ファイルの合計の上限）

    ZipSizeLimitException([string]$message, [string]$partName, [long]$measuredBytes, [string]$limitKind) : base($message) {
        $this.PartName = $partName
        $this.MeasuredBytes = $measuredBytes
        $this.LimitKind = $limitKind
    }
}

function isZipFile {
    # ファイルの先頭がZIPのシグネチャ（PK\x03\x04）かどうか。
    # パスワード付きのOfficeファイルや旧形式は、拡張子が .docx 等でもZIPではない
    param (
        [string]$path
    )

    $stream = [System.IO.File]::OpenRead((toLongPath $path))
    try {
        $head = New-Object byte[] 4
        $read = $stream.Read($head, 0, 4)
        return ($read -eq 4 -and $head[0] -eq 0x50 -and $head[1] -eq 0x4B -and $head[2] -eq 0x03 -and $head[3] -eq 0x04)
    } finally {
        $stream.Dispose()
    }
}

function isCompoundFile {
    # ファイルの先頭が複合ドキュメント形式（旧形式の .doc / .ppt、パスワード付きのOfficeファイル）のシグネチャかどうか
    param (
        [string]$path
    )

    $signature = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1)
    $stream = [System.IO.File]::OpenRead((toLongPath $path))
    try {
        $head = New-Object byte[] 8
        if ($stream.Read($head, 0, 8) -ne 8) {
            return $false
        }
        for ($i = 0; $i -lt 8; $i++) {
            if ($head[$i] -ne $signature[$i]) {
                return $false
            }
        }
        return $true
    } finally {
        $stream.Dispose()
    }
}

# 読み取りのスレッド（runspace）ごとに使い回す、部品の読み取り用の入れ物。大きな部品を読むときに、
# 部品の大きさぶんの配列を毎回確保しない（ラージオブジェクトヒープに乗らない大きさ（85,000 バイト
# 未満。既定 32768 バイト）にとどめ、GC の負担を減らす）
$script:zipReadBufferSize = 32768
$script:zipReadBuffer = $null

function checkZipEntrySize {
    # 部品の展開後の大きさ（entry.Length）を、部品ごとの上限（zipPartMaxBytes）と1ファイルの合計の上限
    # （zipTotalMaxBytes）に照らして数え、超えれば ZipSizeLimitException にする。超えなければ大きさを返す
    # （readZipEntry・readZipEntryBytes が、読む前に呼ぶ。合計は「確保する大きさ」で数える）
    param (
        [System.IO.Compression.ZipArchiveEntry]$entry,
        [string]$entryName
    )

    $length = $entry.Length
    if ($length -gt $script:zipPartMaxBytes) {
        throw [ZipSizeLimitException]::new($script:zipTooLargeMessage, $entryName, $length, "Part")
    }
    $script:zipTotalReadBytes += $length
    if ($script:zipTotalReadBytes -gt $script:zipTotalMaxBytes) {
        throw [ZipSizeLimitException]::new($script:zipTooLargeMessage, $entryName, $script:zipTotalReadBytes, "Total")
    }
    return $length
}

function readZipEntryBytes {
    # ZIP内のファイルをバイト列で返す。無ければ $null（埋め込んだファイルを読むときに使う。ふつうの部品は readZipEntry）。
    # 大きさの数え方と上限は readZipEntry と同じ。バイト列は申告の大きさ（entry.Length。上限の範囲内）だけ確保し、
    # 読み終えたあとにもう 1 バイト読めたら、中身がヘッダーより長い（偽りのヘッダー）として部品ごとの上限の超過にする
    param (
        [System.IO.Compression.ZipArchive]$zip,
        [string]$entryName
    )

    $entry = $zip.GetEntry($entryName)
    if ($null -eq $entry) {
        return $null
    }

    $length = checkZipEntrySize $entry $entryName
    $bytes = New-Object byte[] ([int]$length)
    $filled = 0
    $stream = $entry.Open()
    try {
        while ($filled -lt $length) {
            $read = $stream.Read($bytes, $filled, [int]($length - $filled))
            if ($read -le 0) {
                break
            }
            $filled += $read
        }
        if ($filled -eq $length -and $stream.ReadByte() -ge 0) {
            throw [ZipSizeLimitException]::new($script:zipTooLargeMessage, $entryName, $length + 1, "Part")
        }
    } finally {
        $stream.Dispose()
    }
    if ($filled -lt $length) {
        $shortBytes = New-Object byte[] $filled
        [System.Array]::Copy($bytes, $shortBytes, $filled)
        $bytes = $shortBytes
    }
    return , $bytes
}

function readZipEntry {
    # ZIP内のファイルを文字列で返す。無ければ $null。
    #
    # 展開後の大きさ（entry.Length。中央ディレクトリの値）が部品ごとの上限（zipPartMaxBytes）を超える、
    # または1ファイルの合計（zipTotalReadBytes）が合計の上限（zipTotalMaxBytes）を超えると例外にする。
    # 合計は、実際に読んだ量ではなく「確保する大きさ」（entry.Length）で、読む前に数える
    # （ヘッダーの大きさを偽って中身を小さくした部品を並べても、合計の上限で打ち切るため）。
    #
    # entry.Length の値は、細工して実際より小さく書き換えられる（偽りのヘッダー）。これを見つけるため、
    # 実際に読む量を Length + 1 バイトまでに絞る（入れ物は小さく使い回し、1回ごとの Stream.Read の
    # 要求バイト数を「残り（Length + 1 - 既読量）」で頭打ちにする）。既読量が Length + 1 バイトに届いたら、
    # 中身がヘッダーより長いと分かるので、そこで止めて例外にする。展開はここで必ず止まり、
    # 上限のバイト数を超えて進まない（部品の大きさぶんの配列は作らない）。
    #
    # 読んだバイト列は、今までの StreamReader（UTF-8を指定し、BOMを見て文字コードを決める）と同じ規則で
    # 文字列にする（BOM付きUTF-8・UTF-16 LE/BEを見分け、無ければUTF-8）。入れ物（$script:zipReadBuffer）に
    # 入った分だけ Decoder で文字に直し、StringBuilder に積む（文字の境界がまたがっても Decoder が覚える）
    param (
        [System.IO.Compression.ZipArchive]$zip,
        [string]$entryName
    )

    $entry = $zip.GetEntry($entryName)
    if ($null -eq $entry) {
        return $null
    }

    $length = checkZipEntrySize $entry $entryName

    if ($null -eq $script:zipReadBuffer) {
        # BOM の見分けに先頭 3 バイトが要るため、3 より小さくしない
        $script:zipReadBuffer = New-Object byte[] ([Math]::Max(3, $script:zipReadBufferSize))
    }
    $buffer = $script:zipReadBuffer
    $budget = $length + 1  # 偽りのヘッダーを見つけるため、読む量はこれより増やさない
    $totalRead = 0L
    $sb = New-Object System.Text.StringBuilder  # 容量は渡さない（申告の大きさぶんを先に確保しないため）
    $decoder = $null
    $charBuffer = $null
    $first = $true

    $stream = $entry.Open()
    try {
        while ($totalRead -lt $budget) {
            $want = [Math]::Min($buffer.Length, [long]($budget - $totalRead))
            $filled = 0
            $endOfStream = $false
            while ($filled -lt $want) {
                $read = $stream.Read($buffer, $filled, [int]($want - $filled))
                if ($read -le 0) {
                    $endOfStream = $true
                    break
                }
                $filled += $read
            }
            $totalRead += $filled
            if ($totalRead -gt $length) {
                # 中身がヘッダーの値より長い（偽りのヘッダー）。実際の大きさは分からないため、検出できた時点の値を記録する
                throw [ZipSizeLimitException]::new($script:zipTooLargeMessage, $entryName, $totalRead, "Part")
            }

            $byteOffset = 0
            if ($first) {
                $first = $false
                $encoding = [System.Text.Encoding]::UTF8
                if ($filled -ge 3 -and $buffer[0] -eq 0xEF -and $buffer[1] -eq 0xBB -and $buffer[2] -eq 0xBF) {
                    $byteOffset = 3
                } elseif ($filled -ge 2 -and $buffer[0] -eq 0xFF -and $buffer[1] -eq 0xFE) {
                    $encoding = [System.Text.Encoding]::Unicode
                    $byteOffset = 2
                } elseif ($filled -ge 2 -and $buffer[0] -eq 0xFE -and $buffer[1] -eq 0xFF) {
                    $encoding = [System.Text.Encoding]::BigEndianUnicode
                    $byteOffset = 2
                }
                $decoder = $encoding.GetDecoder()
                $charBuffer = New-Object char[] ($encoding.GetMaxCharCount($buffer.Length))
            }
            $byteCount = $filled - $byteOffset
            if ($byteCount -gt 0) {
                $charCount = $decoder.GetChars($buffer, $byteOffset, $byteCount, $charBuffer, 0, $false)
                if ($charCount -gt 0) {
                    [void]$sb.Append($charBuffer, 0, $charCount)
                }
            }
            if ($endOfStream) {
                break
            }
        }
        if ($null -ne $decoder) {
            # 最後にDecoderの中に残っている分（サロゲートペアの前半など）を吐き出させる
            $charCount = $decoder.GetChars([byte[]]@(), 0, 0, $charBuffer, 0, $true)
            if ($charCount -gt 0) {
                [void]$sb.Append($charBuffer, 0, $charCount)
            }
        }
    } finally {
        $stream.Dispose()
    }

    return $sb.ToString()
}

function resolveZipPath {
    # リレーションシップの Target（相対パス）を、ZIP内のパスに変換する
    param (
        [string]$baseDir,   # 例: "ppt/slides"
        [string]$target     # 例: "../notesSlides/notesSlide1.xml"
    )

    if ($target.StartsWith("/")) {
        return $target.TrimStart("/")
    }

    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($part in (("$baseDir/$target") -split "/")) {
        if ($part -eq "..") {
            if ($parts.Count -gt 0) { $parts.RemoveAt($parts.Count - 1) }
        } elseif ($part -ne "" -and $part -ne ".") {
            $parts.Add($part)
        }
    }
    return ($parts -join "/")
}

function newXmlDocument {
    # XML文字列から XmlDocument を作る。既存の New-Object + LoadXml と同じ挙動
    # （文字の検査はせず、余白は詰める）にしつつ、DTD（<!DOCTYPE ...>）の処理を禁止する。
    # DTDの実体参照を入れ子にして膨張させる攻撃（billion laughs）を防ぐため
    # （readXmlLines と同じ禁止を、LoadXmlを使うすべての呼び出し元にも揃える）
    param (
        [string]$xml
    )

    $settings = New-Object System.Xml.XmlReaderSettings
    $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver = $null
    $settings.CheckCharacters = $false
    $reader = [System.Xml.XmlReader]::Create((New-Object System.IO.StringReader($xml)), $settings)
    try {
        $doc = New-Object System.Xml.XmlDocument
        $doc.PreserveWhitespace = $false
        $doc.Load($reader)
        return , $doc  # XmlDocument は IEnumerable なので、, を付けないとパイプラインで展開される
    } finally {
        $reader.Dispose()
    }
}

function readRelationships {
    # .rels ファイルを読み、Id → @{ Type; Target（ZIP内のパス） } の辞書を返す
    param (
        [System.IO.Compression.ZipArchive]$zip,
        [string]$partName   # 例: "ppt/slides/slide1.xml"
    )

    $dir = [System.IO.Path]::GetDirectoryName($partName).Replace("\", "/")
    $relsName = "$dir/_rels/$([System.IO.Path]::GetFileName($partName)).rels"

    $result = @{}
    $xml = readZipEntry $zip $relsName
    if ($null -eq $xml) {
        return $result
    }

    $doc = newXmlDocument $xml
    foreach ($rel in $doc.GetElementsByTagName("Relationship", ${nsPkgRel})) {
        if ($rel.GetAttribute("TargetMode") -eq "External") {
            continue
        }
        $result[$rel.GetAttribute("Id")] = @{
            Type   = $rel.GetAttribute("Type")
            Target = resolveZipPath $dir $rel.GetAttribute("Target")
        }
    }
    return $result
}

function readXmlLines {
    # WordprocessingML / DrawingML のXMLから、段落・表の行ごとのテキストを文書順に返す。
    # 戻り値: @{ Page = ページ番号; Text = テキスト } の配列
    #
    # ・表の1行はセルをタブ区切りにして1行とする（セル内の複数段落はスペース区切り）
    # ・テキストボックス内の段落は、それぞれ1行とする
    # ・互換用の代替表示（mc:Fallback）は、同じ内容が重複するため読まない
    # ・ページは、手動の改ページ・段落前で改ページ・セクション区切り（連続以外）で数える。
    #   $pageMode = "rendered" の場合は、Wordが保存時に記録したページ区切り（lastRenderedPageBreak）でも数える。
    #   ただし手動の改ページ等の直後（本文が出る前）の lastRenderedPageBreak は同じ区切りなので数えない
    #   （Wordは手動の改ページの後に lastRenderedPageBreak を記録する場合としない場合がある）
    #   $pageMode = "none" の場合はページを数えない（PowerPoint）
    # ・PowerPoint: $onlyPlaceholders を指定すると、その種類のプレースホルダー（例: "ftr"）のテキストだけを読む。
    #   指定しない場合は、スライド番号・日付・ヘッダー・フッター・スライド画像のプレースホルダーを読まない
    # ・$objects（リスト）を渡すと、本文以外の文字の在りかを、ページ付きでそこに集める（渡さなければ今までどおり）:
    #     @{ Kind = "shape";   Page; Text }   Word のテキストボックス・図形内の文字（本文の行には入れない。段落はスペースでつなぐ）
    #     @{ Kind = "diagram"; Page; RelId }  SmartArt（dgm:relIds の r:dm。データはリレーションシップの先）
    #     @{ Kind = "chart";   Page; RelId }  グラフ（c:chart の r:id）
    #     @{ Kind = "comment"; Page; Id }     コメントの参照（w:commentReference の w:id）
    #     @{ Kind = "embed";   Page; RelId }  埋め込んだファイル（Word は o:OLEObject（Type="Embed"）・w:objectEmbed、
    #                                         PowerPoint は p:oleObj の r:id。リレーションシップの先は呼び出し元が読む）
    param (
        [string]$xml,
        [string]$ns,
        [string]$pageMode = "none",
        [string[]]$onlyPlaceholders = $null,
        [System.Collections.Generic.List[object]]$objects = $null
    )

    # 要素・テキストごとに呼ぶため、PowerShell で遅い書き方（スクリプトブロックの呼び出し・switch・型名の解決・
    # PSCustomObject の作成）を避けている（段落 3,000 の文書で約 2 秒かかっていた）。出力は以前と同じ
    $lines = New-Object System.Collections.Generic.List[object]
    # 開いている段落（p）・表の行（tr）・セル（tc）。種類ごとに積み、最も内側は末尾（XmlReader は開始と終了の対応を保証する）
    $pFrames = New-Object System.Collections.Generic.List[hashtable]
    $trFrames = New-Object System.Collections.Generic.List[hashtable]
    $tcFrames = New-Object System.Collections.Generic.List[hashtable]
    $page = 1
    $breakPending = $false  # 手動の改ページ等の後、まだ本文が出ていない
    $inText = $false
    $inSectPr = $false
    $skipShapeText = $false
    # 開いているテキストボックス（$objects を渡したときだけ使う）。TcBase は開いた時点のセルの数（テキストボックスの外のセル）
    $collect = ($null -ne $objects)
    $boxFrames = New-Object System.Collections.Generic.List[hashtable]

    $elementType = [System.Xml.XmlNodeType]::Element
    $endElementType = [System.Xml.XmlNodeType]::EndElement
    $textType = [System.Xml.XmlNodeType]::Text
    $significantWhitespaceType = [System.Xml.XmlNodeType]::SignificantWhitespace
    $whitespaceType = [System.Xml.XmlNodeType]::Whitespace
    $cdataType = [System.Xml.XmlNodeType]::CDATA
    $onlyMode = ($null -ne $onlyPlaceholders)

    $settings = New-Object System.Xml.XmlReaderSettings
    $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $reader = [System.Xml.XmlReader]::Create((New-Object System.IO.StringReader($xml)), $settings)
    try {
        $hasNode = $reader.Read()
        while ($hasNode) {
            $nodeType = $reader.NodeType

            if ($nodeType -eq $elementType) {
                $name = $reader.LocalName
                $uri = $reader.NamespaceURI
                $isEmpty = $reader.IsEmptyElement

                # 読み飛ばす要素（Skip() で次の要素に進むため、Read() はしない）
                $skip = $false
                if ($uri -eq ${nsCompat} -and $name -eq "Fallback") {
                    $skip = $true
                } elseif ($uri -eq ${nsWord} -and $name -eq "moveFrom") {
                    $skip = $true  # 変更履歴の移動元（移動先と重複する）
                } elseif ($uri -eq ${nsPresent} -and $name -eq "txBody" -and $skipShapeText) {
                    $skip = $true
                } elseif ($uri -eq ${nsPresent} -and $name -eq "graphicFrame" -and $onlyMode) {
                    $skip = $true  # 表など（プレースホルダーではない）
                } elseif ($uri -eq $ns -and $name -eq "rPr" -and -not $isEmpty) {
                    $skip = $true  # 文字の書式（テキストを含まない。1 文字ごとに付くことも多く、読むだけで時間がかかる）
                }
                if ($skip) {
                    $reader.Skip()
                    $hasNode = -not $reader.EOF
                    continue
                }

                if ($uri -eq $ns) {
                    if ($name -eq "p") {
                        if (-not $isEmpty) {
                            $pFrames.Add(@{ Text = (New-Object System.Text.StringBuilder); Page = $null; BreakAfter = $false })
                        }
                    } elseif ($name -eq "tr") {
                        if (-not $isEmpty) {
                            $trFrames.Add(@{ Cells = (New-Object System.Collections.Generic.List[string]); Page = $null })
                        }
                    } elseif ($name -eq "tc") {
                        if ($isEmpty) {
                            if ($trFrames.Count -gt 0) { $trFrames[$trFrames.Count - 1].Cells.Add("") }
                        } else {
                            $tcFrames.Add(@{ Text = (New-Object System.Text.StringBuilder); Page = $null })
                        }
                    } elseif ($name -eq "t") {
                        $inText = -not $isEmpty
                    } elseif ($name -eq "tab" -or $name -eq "cr" -or $name -eq "noBreakHyphen") {
                        if ($pFrames.Count -gt 0) {
                            [void]$pFrames[$pFrames.Count - 1].Text.Append($(if ($name -eq "noBreakHyphen") { "-" } else { " " }))
                        }
                    } elseif ($name -eq "br") {
                        $breakType = $reader.GetAttribute("type", $ns)
                        if ($breakType -ne "page" -and $breakType -ne "column") {
                            if ($pFrames.Count -gt 0) {
                                [void]$pFrames[$pFrames.Count - 1].Text.Append(" ")
                            }
                        } elseif ($pageMode -ne "none" -and $breakType -eq "page") {
                            $page++
                            $breakPending = $true
                        }
                    } elseif ($name -eq "lastRenderedPageBreak") {
                        if ($pageMode -eq "rendered") {
                            if ($breakPending) {
                                $breakPending = $false
                            } else {
                                $page++
                            }
                        }
                    } elseif ($name -eq "pageBreakBefore") {
                        if ($pageMode -ne "none" -and $reader.GetAttribute("val", $ns) -notin @("0", "false", "off")) {
                            $page++
                            $breakPending = $true
                        }
                    } elseif ($name -eq "sectPr") {
                        # 段落内のセクション区切り: 既定（次のページから開始）なら、段落の後で改ページ
                        if ($pageMode -ne "none" -and $pFrames.Count -gt 0) {
                            $inSectPr = -not $isEmpty
                            $pFrames[$pFrames.Count - 1].BreakAfter = $true
                        }
                    } elseif ($name -eq "type") {
                        if ($inSectPr -and $reader.GetAttribute("val", $ns) -eq "continuous") {
                            if ($pFrames.Count -gt 0) { $pFrames[$pFrames.Count - 1].BreakAfter = $false }
                        }
                    } elseif ($collect -and $name -eq "txbxContent") {
                        if (-not $isEmpty) {
                            $boxFrames.Add(@{ Lines = (New-Object System.Collections.Generic.List[string]); Page = $page; TcBase = $tcFrames.Count })
                        }
                    } elseif ($collect -and $name -eq "commentReference") {
                        $objects.Add(@{ Kind = "comment"; Page = $page; Id = $reader.GetAttribute("id", $ns) })
                    } elseif ($collect -and $name -eq "objectEmbed") {
                        $objects.Add(@{ Kind = "embed"; Page = $page; RelId = $reader.GetAttribute("id", ${nsRel}) })
                    }
                } elseif ($collect -and $uri -eq ${nsOffice} -and $name -eq "OLEObject") {
                    if ($reader.GetAttribute("Type") -eq "Embed") {
                        $objects.Add(@{ Kind = "embed"; Page = $page; RelId = $reader.GetAttribute("id", ${nsRel}) })
                    }
                } elseif ($collect -and $uri -eq ${nsDiagram} -and $name -eq "relIds") {
                    $objects.Add(@{ Kind = "diagram"; Page = $page; RelId = $reader.GetAttribute("dm", ${nsRel}) })
                } elseif ($collect -and $uri -eq ${nsChart} -and $name -eq "chart") {
                    $objects.Add(@{ Kind = "chart"; Page = $page; RelId = $reader.GetAttribute("id", ${nsRel}) })
                } elseif ($uri -eq ${nsPresent}) {
                    if ($name -eq "sp") {
                        $skipShapeText = $onlyMode
                    } elseif ($name -eq "ph") {
                        if ($onlyMode) {
                            $skipShapeText = ($reader.GetAttribute("type") -notin $onlyPlaceholders)
                        } else {
                            $skipShapeText = ($reader.GetAttribute("type") -in ${skipPlaceholderTypes})
                        }
                    } elseif ($collect -and $name -eq "oleObj") {
                        $objects.Add(@{ Kind = "embed"; Page = $page; RelId = $reader.GetAttribute("id", ${nsRel}) })
                    }
                }
            } elseif ($nodeType -eq $endElementType) {
                if ($reader.NamespaceURI -eq $ns) {
                    $name = $reader.LocalName
                    # 段落・表の行が終わったら、テキストを親のセルに追加する。セルの中でなければ1行として出力する
                    $text = $null
                    if ($name -eq "t") {
                        $inText = $false
                    } elseif ($name -eq "sectPr") {
                        $inSectPr = $false
                    } elseif ($name -eq "p") {
                        $p = $pFrames[$pFrames.Count - 1]
                        $pFrames.RemoveAt($pFrames.Count - 1)
                        $text = $p.Text.ToString().Trim()
                        $textPage = $p.Page
                        $separator = " "
                        if ($p.BreakAfter) {
                            $page++
                            $breakPending = $true
                        }
                    } elseif ($name -eq "tc") {
                        $cell = $tcFrames[$tcFrames.Count - 1]
                        $tcFrames.RemoveAt($tcFrames.Count - 1)
                        if ($trFrames.Count -gt 0) {
                            $row = $trFrames[$trFrames.Count - 1]
                            $row.Cells.Add($cell.Text.ToString())
                            if ($null -eq $row.Page) {
                                $row.Page = $cell.Page
                            }
                        }
                    } elseif ($name -eq "tr") {
                        $row = $trFrames[$trFrames.Count - 1]
                        $trFrames.RemoveAt($trFrames.Count - 1)
                        # 入れ子の表の行は、外側のセルの中ではスペース区切りにする（テキストボックスの外のセルは数えない）
                        $cellBase = $(if ($boxFrames.Count -gt 0) { $boxFrames[$boxFrames.Count - 1].TcBase } else { 0 })
                        $text = ($row.Cells -join $(if ($tcFrames.Count -gt $cellBase) { " " } else { "`t" })).TrimEnd()
                        $textPage = $(if ($null -ne $row.Page) { $row.Page } else { $page })
                        $separator = " "
                    } elseif ($name -eq "txbxContent" -and $boxFrames.Count -gt 0) {
                        # テキストボックスが終わったら、段落をスペースでつないで図形 1 つにする
                        $box = $boxFrames[$boxFrames.Count - 1]
                        $boxFrames.RemoveAt($boxFrames.Count - 1)
                        if ($box.Lines.Count -gt 0) {
                            $objects.Add(@{ Kind = "shape"; Page = $box.Page; Text = ($box.Lines -join " ") })
                        }
                    }
                    if ($text) {
                        # テキストボックスの中（その中のセルは除く）なら、テキストボックスの行にする
                        $box = $null
                        if ($boxFrames.Count -gt 0 -and $tcFrames.Count -le $boxFrames[$boxFrames.Count - 1].TcBase) {
                            $box = $boxFrames[$boxFrames.Count - 1]
                        }
                        if ($box) {
                            $box.Lines.Add($text)
                        } elseif ($tcFrames.Count -gt 0) {
                            $cell = $tcFrames[$tcFrames.Count - 1]
                            if ($cell.Text.Length -gt 0) {
                                [void]$cell.Text.Append($separator)
                            }
                            [void]$cell.Text.Append($text)
                            if ($null -eq $cell.Page) {
                                $cell.Page = $textPage
                            }
                        } else {
                            $lines.Add([pscustomobject]@{ Page = $textPage; Text = $text })
                        }
                    }
                }
            } elseif ($inText -and ($nodeType -eq $textType -or $nodeType -eq $significantWhitespaceType -or
                                    $nodeType -eq $whitespaceType -or $nodeType -eq $cdataType)) {
                if ($pFrames.Count -gt 0) {
                    $p = $pFrames[$pFrames.Count - 1]
                    if ($null -eq $p.Page) {
                        $p.Page = $page
                    }
                    $value = $reader.Value
                    [void]$p.Text.Append($value.Replace("`t", " ").Replace("`r", " ").Replace("`n", " "))
                    if ($value.Trim() -ne "") {
                        $breakPending = $false
                    }
                }
            }

            $hasNode = $reader.Read()
        }
    } finally {
        $reader.Dispose()
    }

    return $lines.ToArray()
}

function addUnitLines {
    # ユニット（順序付き辞書）に行を追加する
    param (
        [System.Collections.Specialized.OrderedDictionary]$units,
        [string]$unitName,
        [string[]]$lines
    )

    if (-not $units.Contains($unitName)) {
        $units[$unitName] = New-Object System.Collections.Generic.List[string]
    }
    foreach ($line in $lines) {
        $units[$unitName].Add($line)
    }
}

function readDiagramText {
    # SmartArt のデータ（diagrams/dataN.xml）の文字を、スペースでつないで 1 つにする。
    # 同じ文字の描画用（diagrams/drawingN.xml）は、重複するため読まない
    param (
        [string]$xml
    )

    return (@(readXmlLines $xml ${nsDrawing} | ForEach-Object { $_.Text }) -join " ")
}

function readChartTxText {
    # c:tx（グラフのタイトル・軸ラベル・系列名で使う、文字を持つ要素）の文字を返す。
    # セル参照（c:strRef/c:strCache/c:pt/c:v）と直値（c:v）の両方を読む。c:rich（a:p）は呼び出し元が別に読む
    param (
        [System.Xml.XmlElement]$parent
    )

    $found = New-Object System.Collections.Generic.List[string]
    foreach ($child in $parent.ChildNodes) {
        if ($child.LocalName -ne "tx") {
            continue
        }
        foreach ($v in $child.GetElementsByTagName("v", ${nsChart})) {
            $text = $v.InnerText.Trim()
            if ($text -ne "") { $found.Add($text) }
        }
        break
    }
    return $found
}

function readChartText {
    # グラフ（charts/chartN.xml）の文字（タイトル・軸ラベル・系列名）を、スペースでつないで 1 つにする。
    # 項目名（横軸の項目。c:cat。点の数だけある）と数値（c:val）は読まない。同じ文字は 1 回だけにする。
    # 系列（c:ser）とタイトル・軸（c:title）の数だけ調べるため、項目の点数が多いグラフでも遅くならない
    param (
        [string]$xml
    )

    $texts = New-Object System.Collections.Generic.List[string]
    $seen = New-Object System.Collections.Generic.HashSet[string]
    # タイトル・軸ラベル（c:rich の a:p。手で直したデータラベルの文字も同じ形のため、ここで読む）
    foreach ($line in (readXmlLines $xml ${nsDrawing})) {
        if ($seen.Add($line.Text)) { $texts.Add($line.Text) }
    }
    $doc = newXmlDocument $xml
    # タイトル・軸ラベルのうち、セル参照の文字（c:title/c:tx/c:strRef/c:strCache/c:pt/c:v。直値の c:rich は上で読み済み）
    foreach ($title in $doc.GetElementsByTagName("title", ${nsChart})) {
        foreach ($text in (readChartTxText $title)) {
            if ($seen.Add($text)) { $texts.Add($text) }
        }
    }
    # 系列名（c:ser/c:tx の文字。セル参照なら c:strRef/c:strCache/c:pt/c:v、直値なら c:v）
    foreach ($ser in $doc.GetElementsByTagName("ser", ${nsChart})) {
        foreach ($text in (readChartTxText $ser)) {
            if ($seen.Add($text)) { $texts.Add($text) }
        }
    }
    return ($texts -join " ")
}

function readObjectText {
    # readXmlLines が集めた SmartArt・グラフの参照（RelId）から、その文字を返す（読めなければ空）
    param (
        [System.IO.Compression.ZipArchive]$zip,
        [hashtable]$rels,   # readRelationships の結果
        $object
    )

    $rel = $rels[[string]$object.RelId]
    if ($null -eq $rel) {
        return ""
    }
    $xml = readZipEntry $zip $rel.Target
    if ($null -eq $xml) {
        return ""
    }
    if ($object.Kind -eq "diagram") {
        return (readDiagramText $xml)
    }
    return (readChartText $xml)
}

function readWordComments {
    # Word のコメント（word/comments.xml）を、w:id → 文字（段落をスペースでつなぐ）の辞書で返す。
    # 返信も別のコメント（別の w:id）として入っている。作成者名は読まない
    param (
        [string]$xml
    )

    $comments = @{}
    if (-not $xml) {
        return $comments  # コメントが無い（[string] の引数は $null を空文字にする）
    }
    $doc = newXmlDocument $xml
    foreach ($comment in $doc.GetElementsByTagName("comment", ${nsWord})) {
        $text = @(readXmlLines $comment.OuterXml ${nsWord} | ForEach-Object { $_.Text }) -join " "
        if ($text -ne "") {
            $comments[$comment.GetAttribute("id", ${nsWord})] = $text
        }
    }
    return $comments
}

function readSlideComments {
    # PowerPoint のスライドのコメントを、コメント・返信ごとの文字の配列で返す（作成者名は読まない）。
    #   旧形式（ppt/comments/commentN.xml）: p:cm の p:text
    #   新形式（ppt/comments/modernComment_*.xml）: p188:cm の p188:txBody と、返信（p188:reply）の p188:txBody
    param (
        [string]$xml
    )

    $texts = New-Object System.Collections.Generic.List[string]
    $doc = newXmlDocument $xml
    foreach ($cm in @($doc.SelectNodes("//*") | Where-Object { $_.LocalName -eq "cm" })) {
        if ($cm.NamespaceURI -eq ${nsPresent}) {
            $body = @($cm.ChildNodes | Where-Object { $_.LocalName -eq "text" })
            if ($body.Count -gt 0 -and $body[0].InnerText.Trim() -ne "") { $texts.Add($body[0].InnerText.Trim()) }
            continue
        }
        # 新形式: XML では返信の一覧（replyLst）が本文（txBody）より前にあるため、本文を先に出してから返信を出す
        $bodies = @($cm.ChildNodes | Where-Object { $_.LocalName -eq "txBody" })
        foreach ($reply in @($cm.ChildNodes | Where-Object { $_.LocalName -eq "replyLst" } | ForEach-Object { $_.ChildNodes } | Where-Object { $_.LocalName -eq "reply" })) {
            $bodies += @($reply.ChildNodes | Where-Object { $_.LocalName -eq "txBody" })
        }
        foreach ($body in $bodies) {
            $text = @(readXmlLines $body.OuterXml ${nsDrawing} | ForEach-Object { $_.Text }) -join " "
            if ($text -ne "") { $texts.Add($text) }
        }
    }
    return $texts.ToArray()
}

function readDocxUnits {
    # Word（.docx / .docm）のテキストを、ページ・ヘッダー/フッター・脚注ごとに返す。
    # 本文のテキストボックス・図形内の文字と SmartArt・グラフの文字は "ページNNN[図形]"、
    # コメントは "ページNNN[コメント]"（コメントを付けた所のページ）に分ける（1 行は図形・コメント 1 つ）。
    # 埋め込んだ Office のファイルの文字は "ページNNN[埋め込み<N>]"（N は 1 ファイルの中の、参照の順の通し番号。
    # 読まなかったものは欠番。office_embedded.ps1）。
    # $failures・$sizeFailures を渡すと、読めなかった埋め込みの部品の名前と、大きさの上限の例外を追加する
    # （呼び出し元がインデックス作成のログに書く。shared/ はツールを知らないため、ここでは書かない）。
    # ファイルを開き、合計（zipTotalReadBytes）を 0 から数え直して、閉じる。読む本体は readDocxUnitsFromZip
    param (
        [string]$path,
        [System.Collections.Generic.List[string]]$failures = $null,
        [System.Collections.Generic.List[object]]$sizeFailures = $null
    )

    $script:zipTotalReadBytes = 0  # このファイル1つ分の、readZipEntryが読む合計（zipTotalMaxBytes）を0から数え直す
    $zip = [System.IO.Compression.ZipFile]::OpenRead((toLongPath $path))
    try {
        return (readDocxUnitsFromZip $zip $true $failures $sizeFailures)
    } finally {
        $zip.Dispose()
    }
}

function readDocxUnitsFromZip {
    # 開いた Word の ZipArchive から、readDocxUnits と同じユニットを返す（開く・閉じる・合計の数え直しはしない）。
    # $readEmbeds が $false なら埋め込みは読まない（埋め込みの中の Word 文書を読むとき。深さは 1 段まで）
    param (
        [System.IO.Compression.ZipArchive]$zip,
        [bool]$readEmbeds,
        [System.Collections.Generic.List[string]]$failures = $null,
        [System.Collections.Generic.List[object]]$sizeFailures = $null
    )

    $units = New-Object System.Collections.Specialized.OrderedDictionary
    $body = readZipEntry $zip "word/document.xml"
    if ($null -eq $body) {
        throw "Word文書の本文（word/document.xml）がありません。"
    }

    # Wordで保存されたファイルには、保存時点のページ区切りが記録されている。無ければ手動の改ページで数える
    $pageMode = $(if ($body.Contains("lastRenderedPageBreak")) { "rendered" } else { "explicit" })
    # 本文は行数が多いため、1行ずつ addUnitLines を呼ばずにページのユニットへ入れる（結果は addUnitLines と同じ）
    $lastPage = $null
    $pageLines = $null
    $objects = New-Object System.Collections.Generic.List[object]
    foreach ($line in (readXmlLines $body ${nsWord} $pageMode $null $objects)) {
        if ($null -eq $pageLines -or $line.Page -ne $lastPage) {
            $unitName = "ページ{0:D3}" -f $line.Page
            if (-not $units.Contains($unitName)) {
                $units[$unitName] = New-Object System.Collections.Generic.List[string]
            }
            $pageLines = $units[$unitName]
            $lastPage = $line.Page
        }
        $pageLines.Add($line.Text)
    }

    # 本文の XML の文字列は、ここから先で使わない。埋め込みを読む前に放す（同時に持つメモリを抑える）
    $body = $null

    # 図形（テキストボックス・SmartArt・グラフ）・コメント・埋め込み。文書の中の順に、そのページの場所へ入れる
    $rels = readRelationships $zip "word/document.xml"
    $comments = readWordComments (readZipEntry $zip "word/comments.xml")
    $usedComments = New-Object System.Collections.Generic.HashSet[string]
    $embedState = newEmbeddedState
    foreach ($object in $objects) {
        $base = "ページ{0:D3}" -f $object.Page
        if ($object.Kind -eq "shape") {
            addUnitLines $units "${base}[図形]" @($object.Text)
        } elseif ($object.Kind -eq "comment") {
            $id = [string]$object.Id
            if ($comments.ContainsKey($id) -and $usedComments.Add($id)) {
                addUnitLines $units "${base}[コメント]" @($comments[$id])
            }
        } elseif ($object.Kind -eq "embed") {
            if ($readEmbeds) {
                $embedded = readEmbeddedObjectLines $zip $rels $object $embedState $failures $sizeFailures
                if ($null -ne $embedded -and $embedded.Lines.Count -gt 0) {
                    addUnitLines $units "${base}[埋め込み$($embedded.Number)]" $embedded.Lines
                }
            }
        } else {
            $text = readObjectText $zip $rels $object
            if ($text -ne "") {
                addUnitLines $units "${base}[図形]" @($text)
            }
        }
    }
    # 本文に参照の無いコメント（ヘッダー・脚注に付けたものなど）は、場所が分からないため "文書[コメント]" にまとめる
    foreach ($id in @($comments.Keys | Sort-Object { $n = 0; [void][int]::TryParse($_, [ref]$n); $n })) {
        if (-not $usedComments.Contains($id)) {
            addUnitLines $units "文書[コメント]" @($comments[$id])
        }
    }

    # ヘッダー・フッター（セクションごとに同じ内容が並ぶため、重複は除く）
    $seen = New-Object System.Collections.Generic.HashSet[string]
    $parts = @($zip.Entries | Where-Object { $_.FullName -match "^word/(header|footer)\d*\.xml$" } |
        Sort-Object { $_.FullName -notmatch "/header" }, FullName)  # ヘッダー → フッターの順
    foreach ($part in $parts) {
        foreach ($line in (readXmlLines (readZipEntry $zip $part.FullName) ${nsWord})) {
            if ($seen.Add($line.Text)) {
                addUnitLines $units "ヘッダー・フッター" @($line.Text)
            }
        }
    }

    # 脚注・文末脚注
    foreach ($partName in @("word/footnotes.xml", "word/endnotes.xml")) {
        $xml = readZipEntry $zip $partName
        if ($null -ne $xml) {
            addUnitLines $units "脚注" @(readXmlLines $xml ${nsWord} | ForEach-Object { $_.Text })
        }
    }

    return $units
}

function readPptxUnits {
    # PowerPoint（.pptx / .pptm）のテキストを、スライド・ノートごと（スライドの表示順）と、
    # スライドのフッター（全スライド分をまとめ、重複を除く）に分けて返す。
    # 埋め込んだ Office のファイルの文字は "スライドNNN[埋め込み<N>]"（N は 1 ファイルの中の、参照の順の通し番号。
    # 読まなかったものは欠番。office_embedded.ps1）。
    # $failures・$sizeFailures の意味と、ファイルの開閉・合計の数え直しは readDocxUnits と同じ。読む本体は readPptxUnitsFromZip
    param (
        [string]$path,
        [System.Collections.Generic.List[string]]$failures = $null,
        [System.Collections.Generic.List[object]]$sizeFailures = $null
    )

    $script:zipTotalReadBytes = 0  # このファイル1つ分の、readZipEntryが読む合計（zipTotalMaxBytes）を0から数え直す
    $zip = [System.IO.Compression.ZipFile]::OpenRead((toLongPath $path))
    try {
        return (readPptxUnitsFromZip $zip $true $failures $sizeFailures)
    } finally {
        $zip.Dispose()
    }
}

function readPptxUnitsFromZip {
    # 開いた PowerPoint の ZipArchive から、readPptxUnits と同じユニットを返す（開く・閉じる・合計の数え直しはしない）。
    # $readEmbeds が $false なら埋め込みは読まない（埋め込みの中の PowerPoint を読むとき。深さは 1 段まで）
    param (
        [System.IO.Compression.ZipArchive]$zip,
        [bool]$readEmbeds,
        [System.Collections.Generic.List[string]]$failures = $null,
        [System.Collections.Generic.List[object]]$sizeFailures = $null
    )

    $units = New-Object System.Collections.Specialized.OrderedDictionary
    $footers = New-Object System.Collections.Generic.List[string]
    $seenFooters = New-Object System.Collections.Generic.HashSet[string]
    $embedState = newEmbeddedState
    $presentationXml = readZipEntry $zip "ppt/presentation.xml"
    if ($null -eq $presentationXml) {
        throw "PowerPointのプレゼンテーション情報（ppt/presentation.xml）がありません。"
    }
    $presentation = newXmlDocument $presentationXml
    $presentationRels = readRelationships $zip "ppt/presentation.xml"

    # スライドの表示順は sldIdLst の順（ファイル名の番号とは一致しないことがある）
    $number = 0
    foreach ($slideId in $presentation.GetElementsByTagName("sldId", ${nsPresent})) {
        $number++
        $rel = $presentationRels[$slideId.GetAttribute("id", ${nsRel})]
        if ($null -eq $rel) {
            continue
        }

        $slideXml = readZipEntry $zip $rel.Target
        if ($null -eq $slideXml) {
            continue
        }

        $unitName = "スライド{0:D3}" -f $number
        if ($slideXml -match '^[\s\S]{0,2000}?<p:sld\b[^>]*\sshow="(0|false)"') {
            $unitName += "（非表示）"
        }
        # テキストボックス・図形の文字はスライドの本文にする（スライドの文字はほとんどが図形のため）。
        # SmartArt・グラフの文字は "スライドNNN[図形]"、コメントは "スライドNNN[コメント]" に分ける
        $objects = New-Object System.Collections.Generic.List[object]
        addUnitLines $units $unitName @(readXmlLines $slideXml ${nsDrawing} "none" $null $objects | ForEach-Object { $_.Text })
        $slideRels = readRelationships $zip $rel.Target
        foreach ($object in $objects) {
            if ($object.Kind -eq "embed") {
                continue  # 埋め込みは、スライドの文字をすべて読んだあとに読む（下）
            }
            $text = readObjectText $zip $slideRels $object
            if ($text -ne "") {
                addUnitLines $units "${unitName}[図形]" @($text)
            }
        }
        foreach ($slideRel in $slideRels.Values) {
            if ($slideRel.Type -like "*/comments") {
                $commentsXml = readZipEntry $zip $slideRel.Target
                if ($null -ne $commentsXml) {
                    addUnitLines $units "${unitName}[コメント]" @(readSlideComments $commentsXml)
                }
            }
        }

        # スライドのフッター（各スライドに同じ内容が並ぶため、重複は除く）
        foreach ($line in (readXmlLines $slideXml ${nsDrawing} "none" @("ftr"))) {
            if ($seenFooters.Add($line.Text)) {
                $footers.Add($line.Text)
            }
        }

        # 発表者ノート
        foreach ($slideRel in $slideRels.Values) {
            if ($slideRel.Type -like "*/notesSlide") {
                $notesXml = readZipEntry $zip $slideRel.Target
                if ($null -ne $notesXml) {
                    addUnitLines $units ("スライド{0:D3}_ノート" -f $number) @(readXmlLines $notesXml ${nsDrawing} | ForEach-Object { $_.Text })
                }
            }
        }

        # 埋め込んだファイル（スライドの中の XML の順）。スライドの XML の文字列は、ここから先で使わない
        # ので、埋め込みを読む前に放す（同時に持つメモリを抑える）
        $slideXml = $null
        if ($readEmbeds) {
            foreach ($object in @($objects | Where-Object { $_.Kind -eq "embed" })) {
                $embedded = readEmbeddedObjectLines $zip $slideRels $object $embedState $failures $sizeFailures
                if ($null -ne $embedded -and $embedded.Lines.Count -gt 0) {
                    addUnitLines $units "${unitName}[埋め込み$($embedded.Number)]" $embedded.Lines
                }
            }
        }
    }

    if ($footers.Count -gt 0) {
        addUnitLines $units "ヘッダー・フッター" $footers.ToArray()
    }

    return $units
}

function toObjectCellText {
    # 図形・コメントの文字を、Excel のテキスト保存と同じ形の 1 セルにする。
    # 改行はセル内改行（$cellNewLine）にし、改行・" ・タブを含むときは " で囲む（中の " は "" にする）
    param (
        [string[]]$lines
    )

    $text = ($lines -join "`n").Trim()
    if ($text.IndexOfAny([char[]]@('"', "`t", "`r", "`n")) -lt 0) {
        return $text
    }
    $text = ($text -replace "\r\n|\r|\n", ${cellNewLine}).Replace('"', '""')
    return "`"$text`""
}

function readXlsxShapeRows {
    # 図形（xl/drawings/drawingN.xml）ごとの文字を @{ Row; Column; Text } の配列で返す。
    # 行・列は図形の左上のセル（1 から数える）。グループ化した図形は、まとめて 1 つの図形とする。
    # $zip・$drawingPath（この図形の部品自身の ZIP 内のパス）を渡すと、グラフ（c:chart）・SmartArt（dgm:relIds）の
    # 参照も解決し、テキストボックスの段落 → グラフ・SmartArt の文字（XML の順）の順に 1 つの図形（1 行）にする。
    # リレーションシップ（_rels）は、参照が 1 つ以上あるときだけ読む。
    # 参照の先・リレーションシップが無い、部品が読めない（XML が壊れているなど）ときは、そのグラフ・SmartArt だけを
    # 空にして続ける（同じ図形の中のほかの文字、ほかの図形は出す）。$failures を渡すと、読めなかった部品の名前を追加する。
    # サイズの上限（ZipSizeLimitException）で読めなかったときは、原因を調べられるよう、その例外自体を
    # $sizeFailures に追加する（呼び出し元がインデックス作成のログに部品名・大きさ・部品ごとか合計かを書く）
    param (
        [string]$xml,
        [System.IO.Compression.ZipArchive]$zip = $null,
        [string]$drawingPath = $null,
        [System.Collections.Generic.List[string]]$failures = $null,
        [System.Collections.Generic.List[object]]$sizeFailures = $null
    )

    $nsSheetDrawing = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    $doc = newXmlDocument $xml
    $rows = New-Object System.Collections.Generic.List[object]
    $rels = $null  # この図形の部品（drawingN.xml）自身のリレーションシップ。参照が1つ以上あるときだけ読む
    foreach ($anchor in @($doc.DocumentElement.SelectNodes("//*")) | Where-Object {
            $_.NamespaceURI -eq $nsSheetDrawing -and $_.LocalName -in @("twoCellAnchor", "oneCellAnchor", "absoluteAnchor") }) {
        # 互換用の代替表示（mc:Fallback）の中は、mc:Choice と同じ図形なので読まない
        $inFallback = $false
        for ($node = $anchor.ParentNode; $null -ne $node; $node = $node.ParentNode) {
            if ($node.NamespaceURI -eq ${nsCompat} -and $node.LocalName -eq "Fallback") {
                $inFallback = $true
                break
            }
        }
        if ($inFallback) {
            continue
        }

        $objects = New-Object System.Collections.Generic.List[object]
        $lines = New-Object System.Collections.Generic.List[string]
        foreach ($line in (readXmlLines $anchor.OuterXml ${nsDrawing} "none" $null $objects)) {
            $lines.Add($line.Text)
        }
        # グラフ（c:chart）・SmartArt（dgm:relIds）の文字を、テキストボックスの段落の後、XML の順に並べる
        foreach ($object in @($objects | Where-Object { $_.Kind -eq "chart" -or $_.Kind -eq "diagram" })) {
            $text = ""
            try {
                # リレーションシップ自体（drawingN.xml.rels）が壊れていても、この図形だけを空にしてほかは捨てない
                if ($null -eq $rels) {
                    $rels = $(if ($zip -and $drawingPath) { readRelationships $zip $drawingPath } else { @{} })
                }
                $text = readObjectText $zip $rels $object
            } catch {
                if ($null -ne $sizeFailures -and $_.Exception -is [ZipSizeLimitException]) {
                    # サイズの上限を超えたときは、例外自体を渡す（部品名は例外の PartName に入っている）
                    $sizeFailures.Add($_.Exception)
                }
                if ($null -ne $failures) {
                    if ($null -eq $rels -and $drawingPath) {
                        # リレーションシップ自体（drawingN.xml.rels）が読めなかった場合は、図形の部品の名前を記録する
                        # （$object.RelId だけでは、どの部品が壊れているのか分からないため）
                        $failures.Add($drawingPath)
                    } else {
                        $rel = $(if ($rels) { $rels[[string]$object.RelId] } else { $null })
                        $failures.Add($(if ($rel) { $rel.Target } else { [string]$object.RelId }))
                    }
                }
            }
            if ($text -ne "") {
                $lines.Add($text)
            }
        }
        if ($lines.Count -eq 0) {
            continue
        }
        # 位置をセルで持たない図形（absoluteAnchor）は A1 とする
        $row = 1
        $column = 1
        $from = $anchor.GetElementsByTagName("from", $nsSheetDrawing)
        if ($from.Count -gt 0) {
            $row = 1 + [int]$from[0].GetElementsByTagName("row", $nsSheetDrawing)[0].InnerText
            $column = 1 + [int]$from[0].GetElementsByTagName("col", $nsSheetDrawing)[0].InnerText
        }
        $rows.Add(@{ Row = $row; Column = $column; Text = (toObjectCellText $lines.ToArray()) })
    }
    return $rows.ToArray()
}

function readXlsxCommentRows {
    # コメント（メモ。xl/commentsN.xml）とスレッド形式のコメント（xl/threadedComments/*.xml）の文字を、
    # セル番地 → 文字の辞書で返す。スレッド形式のコメントがあるセルは、その文字（返信を含む）を使う
    # （同じセルのメモには、古い版の Excel 向けの案内文とコメントが入っているため）
    param (
        [string]$commentsXml,
        [string[]]$threadedXmls
    )

    $nsSheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    $nsThreaded = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
    $result = [ordered]@{}

    foreach ($xml in @($threadedXmls | Where-Object { $_ })) {
        $doc = newXmlDocument $xml
        foreach ($comment in $doc.GetElementsByTagName("threadedComment", $nsThreaded)) {
            $ref = $comment.GetAttribute("ref")
            $text = @($comment.GetElementsByTagName("text", $nsThreaded) | ForEach-Object { $_.InnerText }) -join "`n"
            if (-not $ref -or $text.Trim() -eq "") {
                continue
            }
            if ($result.Contains($ref)) {
                $result[$ref] += "`n" + $text
            } else {
                $result[$ref] = $text
            }
        }
    }

    if ($commentsXml) {
        $doc = newXmlDocument $commentsXml
        foreach ($comment in $doc.GetElementsByTagName("comment", $nsSheet)) {
            $ref = $comment.GetAttribute("ref")
            if (-not $ref -or $result.Contains($ref)) {
                continue
            }
            # ふりがな（rPh）は読まない
            $text = @($comment.GetElementsByTagName("t", $nsSheet) | Where-Object { $_.ParentNode.LocalName -ne "rPh" } |
                ForEach-Object { $_.InnerText }) -join ""
            if ($text.Trim() -ne "") {
                $result[$ref] = $text
            }
        }
    }
    return $result
}

function getCellPosition {
    # セル番地（例: "AB12"）を @(行, 列) にする。読めなければ @(0, 0)
    param (
        [string]$ref
    )

    $m = [regex]::Match($ref, '^\$?([A-Za-z]+)\$?(\d+)$')
    if (-not $m.Success) {
        return @(0, 0)
    }
    $column = 0
    foreach ($ch in $m.Groups[1].Value.ToUpperInvariant().ToCharArray()) {
        $column = $column * 26 + ([int]$ch - [int][char]"A" + 1)
    }
    return @([int]$m.Groups[2].Value, $column)
}

function addSortedShapeLines {
    # 図形の一覧を、上の行から（同じ行は左から、同じセルは XML の順に）並べてユニットに加える。
    # Sort-Object は同じキーの順を保たない（インデックス作成のたびに順が変わる）ため、XML の順もキーにする
    param (
        [System.Collections.Specialized.OrderedDictionary]$units,
        [string]$unitName,
        [System.Collections.Generic.List[object]]$shapes
    )

    for ($i = 0; $i -lt $shapes.Count; $i++) {
        $shapes[$i].Order = $i
    }
    $lines = @($shapes | Sort-Object { $_.Row }, { $_.Column }, { $_.Order } |
        ForEach-Object { "$(toColumnName $_.Column)$($_.Row)`t$($_.Text)" })
    if ($lines.Count -gt 0) {
        addUnitLines $units $unitName $lines
    }
}

function getHeaderFooterLines {
    # Excel のヘッダー・フッターの書式コード付きの文字（例: "&L社外秘&C&"ＭＳ ゴシック,太字"&12見出し&R&P / &N ページ"）を、
    # 検索する行の一覧にする。左（&L）・中央（&C。区切りが無い先頭も中央）・右（&R）の順で、1 区分の 1 段落が 1 行。
    # 印刷時に決まる値（&P ページ・&N 総ページ・&D 日付・&T 時刻・&Z パス・&F ファイル名・&A シート名・&G 画像）と、
    # 書式（フォント・大きさ・色・&B &I &U &E &S &X &Y &O &H）は文字にしない。&& は & 1 文字。知らない & の続きはそのまま残す。
    # タブはスペースにして前後の空白を落とし、空の行は出さない。1 行ずつ toObjectCellText を通す
    # （Excel の行は " で囲んだセルを外す前提のため）
    param (
        [string]$code
    )

    $sections = @{ L = New-Object System.Text.StringBuilder; C = New-Object System.Text.StringBuilder; R = New-Object System.Text.StringBuilder }
    $current = $sections["C"]
    $length = $code.Length
    $i = 0
    while ($i -lt $length) {
        $ch = $code[$i]
        if ($ch -ne '&' -or $i + 1 -ge $length) {
            [void]$current.Append($ch)
            $i++
            continue
        }
        $next = $code[$i + 1]
        if ($next -eq '&') {
            [void]$current.Append('&')
            $i += 2
        } elseif ($next -cin @('L', 'C', 'R')) {
            $current = $sections[[string]$next]
            $i += 2
        } elseif ($next -ceq 'P') {
            # &P+1 &P-1（ページ番号の加減）の数字も取り除く
            $m = [regex]::Match($code.Substring($i), '^&P(?:[+-]\d+)?')
            $i += $m.Length
        } elseif ($next -cin @('N', 'D', 'T', 'Z', 'F', 'A', 'G', 'B', 'I', 'U', 'E', 'S', 'X', 'Y', 'O', 'H')) {
            $i += 2
        } elseif ($next -eq '"') {
            $close = $code.IndexOf('"', $i + 2)
            if ($close -lt 0) {
                [void]$current.Append('&')
                $i++
            } else {
                $i = $close + 1
            }
        } elseif ($next -ceq 'K') {
            $m = [regex]::Match($code.Substring($i), '^&K(?:[0-9A-Fa-f]{6}|[0-9]{2}[+-][0-9]{3})')
            if ($m.Success) {
                $i += $m.Length
            } else {
                [void]$current.Append('&')
                $i++
            }
        } elseif ($next -ge [char]'0' -and $next -le [char]'9') {
            $m = [regex]::Match($code.Substring($i), '^&\d{1,3}')
            $i += $m.Length
        } else {
            [void]$current.Append('&')
            $i++
        }
    }

    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($key in @("L", "C", "R")) {
        foreach ($part in ($sections[$key].ToString() -split "\r\n|\r|\n")) {
            $text = $part.Replace("`t", " ").Trim()
            if ($text -ne "") {
                $lines.Add((toObjectCellText @($text)))
            }
        }
    }
    return $lines.ToArray()
}

function readXlsxHeaderFooterLines {
    # シート（ワークシート・グラフシート）の XML（XmlReader）から、ヘッダー・フッターの検索する行を返す。
    # 並びは ヘッダー → フッター、それぞれ 先頭ページ（differentFirst のとき）→ 奇数（通常）ページ →
    # 偶数ページ（differentOddEven のとき）。同じ文字の行は 1 回だけ。
    # sheetData（セルの値。大きいシートでは数百 MB になる）は読まずに飛ばし、headerFooter を読んだら終わる
    param (
        [System.Xml.XmlReader]$reader
    )

    $nsSheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    $texts = @{}
    $differentFirst = $false
    $differentOddEven = $false
    $hasNode = $reader.Read()
    while ($hasNode) {
        if ($reader.NodeType -eq [System.Xml.XmlNodeType]::Element -and $reader.NamespaceURI -eq $nsSheet) {
            if ($reader.LocalName -eq "sheetData") {
                $reader.Skip()
                $hasNode = -not $reader.EOF
                continue
            }
            if ($reader.LocalName -eq "headerFooter") {
                $differentFirst = $reader.GetAttribute("differentFirst") -in @("1", "true")
                $differentOddEven = $reader.GetAttribute("differentOddEven") -in @("1", "true")
                $sub = $reader.ReadSubtree()
                try {
                    $sub.Read() | Out-Null  # headerFooter 自身
                    $hasChild = $sub.Read()
                    while ($hasChild) {
                        if ($sub.NodeType -eq [System.Xml.XmlNodeType]::Element -and $sub.Depth -eq 1 -and $sub.NamespaceURI -eq $nsSheet) {
                            $childName = $sub.LocalName  # 読んだ後は位置が進むため、先に控える
                            $texts[$childName] = $sub.ReadElementContentAsString()
                            $hasChild = -not $sub.EOF
                        } else {
                            $hasChild = $sub.Read()
                        }
                    }
                } finally {
                    $sub.Dispose()
                }
                break
            }
        }
        $hasNode = $reader.Read()
    }

    $names = New-Object System.Collections.Generic.List[string]
    foreach ($kind in @("Header", "Footer")) {
        if ($differentFirst) { $names.Add("first$kind") }
        $names.Add("odd$kind")
        if ($differentOddEven) { $names.Add("even$kind") }
    }
    $lines = New-Object System.Collections.Generic.List[string]
    $seen = New-Object System.Collections.Generic.HashSet[string]
    foreach ($name in $names) {
        if (-not $texts.ContainsKey($name)) {
            continue
        }
        foreach ($line in (getHeaderFooterLines $texts[$name])) {
            if ($seen.Add($line)) { $lines.Add($line) }
        }
    }
    return $lines.ToArray()
}

function readXlsxSheetHeaderFooter {
    # ZIP の中のシートの部品（$entryName）か、読み込み済みの文字列（$xml。グラフシート）から、
    # ヘッダー・フッターの行を返す。読めない（XML が壊れているなど）ときは空にし、$failures に部品の名前を足す
    param (
        [System.IO.Compression.ZipArchive]$zip,
        [string]$entryName,
        [string]$xml = $null,
        [System.Collections.Generic.List[string]]$failures = $null
    )

    $settings = New-Object System.Xml.XmlReaderSettings
    $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    # シートの部品は sheetData が大きいため、文字列にせず流れで読む（readZipEntry は使わない）。
    # 代わりに、申告の大きさ（entry.Length）が上限（zipSheetStreamMaxBytes）を超える部品は開かずに読めなかった扱いにする。
    # MaxCharactersInDocument は、申告を偽った部品（実際は大きいのに小さく申告）を打ち切る守り（XmlException → 読めなかった扱い）
    $settings.MaxCharactersInDocument = $script:zipSheetStreamMaxBytes
    $stream = $null
    $reader = $null
    try {
        if ($xml) {
            $reader = [System.Xml.XmlReader]::Create((New-Object System.IO.StringReader($xml)), $settings)
        } else {
            $entry = $zip.GetEntry($entryName)
            if ($null -eq $entry) {
                return @()
            }
            if ($entry.Length -gt $script:zipSheetStreamMaxBytes) {
                # 申告の大きさが流れ読みの上限を超える（開かずに読めなかった扱いにし、理由をログに出せるようにする）
                if ($null -ne $failures) {
                    $failures.Add("$entryName（流れ読みの上限 $([long]($script:zipSheetStreamMaxBytes / 1MB))MB を超えています）")
                }
                return @()
            }
            $stream = $entry.Open()
            $reader = [System.Xml.XmlReader]::Create($stream, $settings)
        }
        return @(readXlsxHeaderFooterLines $reader)
    } catch {
        if ($null -ne $failures) {
            if ($_.Exception.Message -like "*MaxCharactersInDocument*") {
                $failures.Add("$entryName（流れ読みの上限 $([long]($script:zipSheetStreamMaxBytes / 1MB))MB を超えています）")
            } else {
                $failures.Add($entryName)
            }
        }
        return @()
    } finally {
        if ($null -ne $reader) { $reader.Dispose() }
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function readXlsxObjectUnits {
    # Excel（.xlsx / .xlsm）の表示シート・表示のグラフシートにある図形（テキストボックス・グループ・WordArt・
    # グラフ・SmartArt を含む）・コメント・ヘッダー／フッターの文字を、"<シート名>[図形]" "<シート名>[コメント]"
    # "<シート名>[ヘッダー・フッター]" の場所ごとに返す（シート名には [ ] を使えないため、実在のシートと重ならない）。
    # 図形・コメントの 1 行は "<セル番地><TAB><文字>"。セル番地は図形の左上・コメントのセルで、シートの上の行から順に並べる。
    # ヘッダー・フッターの 1 行は文字だけ（セル番地は無い。getHeaderFooterLines・readXlsxHeaderFooterLines）。
    # セルの値は Excel のテキスト保存で読むため、ここでは読まない。
    # ZIP の中身が Excel のブック（xl/workbook.xml）でなければ（.xlsb など）、何も返さない。
    # $failures を渡すと、読めなかったグラフ・SmartArt の部品の名前を追加する（呼び出し元でログに書く。
    # shared/ はツールを知らないため、ここでは書かない）。
    # $sizeFailures を渡すと、サイズの上限（ZipSizeLimitException）で読めなかったグラフ・SmartArt の
    # 例外自体を追加する（呼び出し元が部品名・大きさ・部品ごとか合計かをログに書く）
    param (
        [string]$path,
        [System.Collections.Generic.List[string]]$failures = $null,
        [System.Collections.Generic.List[object]]$sizeFailures = $null
    )

    $script:zipTotalReadBytes = 0  # このファイル1つ分の、readZipEntryが読む合計（zipTotalMaxBytes）を0から数え直す
    $units = New-Object System.Collections.Specialized.OrderedDictionary
    $nsSheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    $zip = [System.IO.Compression.ZipFile]::OpenRead((toLongPath $path))
    try {
        $workbookXml = readZipEntry $zip "xl/workbook.xml"
        if ($null -eq $workbookXml) {
            return $units
        }
        $workbook = newXmlDocument $workbookXml
        $workbookRels = readRelationships $zip "xl/workbook.xml"

        foreach ($sheet in $workbook.GetElementsByTagName("sheet", $nsSheet)) {
            # 非表示（hidden）・完全に非表示（veryHidden）のシートは、セルと同じく読まない
            if ($sheet.GetAttribute("state") -in @("hidden", "veryHidden")) {
                continue
            }
            $rel = $workbookRels[$sheet.GetAttribute("id", ${nsRel})]
            if ($null -eq $rel) {
                continue
            }
            $sheetName = $sheet.GetAttribute("name")

            if ($rel.Type -like "*/chartsheet") {
                # 表示のグラフシート: グラフ自体を absoluteAnchor で置いた図形の部品を、通常のシートと同じ形で読む。
                # 位置をセルで持たないため、セル番地は今の決まりどおり A1 になる
                $chartsheetXml = readZipEntry $zip $rel.Target
                if ($null -eq $chartsheetXml) {
                    continue
                }
                $shapes = New-Object System.Collections.Generic.List[object]
                foreach ($csRel in (readRelationships $zip $rel.Target).Values) {
                    if ($csRel.Type -notlike "*/drawing") {
                        continue
                    }
                    $xml = readZipEntry $zip $csRel.Target
                    if ($null -ne $xml) {
                        $shapes.AddRange([object[]]@(readXlsxShapeRows $xml $zip $csRel.Target $failures $sizeFailures))
                    }
                }
                addSortedShapeLines $units "${sheetName}[図形]" $shapes
                $headerFooter = @(readXlsxSheetHeaderFooter $zip $rel.Target $chartsheetXml $failures)
                if ($headerFooter.Count -gt 0) {
                    addUnitLines $units "${sheetName}[ヘッダー・フッター]" $headerFooter
                }
                continue
            }
            if ($rel.Type -notlike "*/worksheet") {
                continue  # ダイアログシートなど
            }

            $shapes = New-Object System.Collections.Generic.List[object]
            $commentsXml = $null
            $threadedXmls = New-Object System.Collections.Generic.List[string]
            foreach ($sheetRel in (readRelationships $zip $rel.Target).Values) {
                # 要る 3 種類だけ読む（埋め込み・背景の画像など、関係のない大きな部品でサイズの上限に当たらないように）
                if ($sheetRel.Type -notlike "*/drawing" -and $sheetRel.Type -notlike "*/comments" -and $sheetRel.Type -notlike "*/threadedComment") {
                    continue
                }
                $xml = readZipEntry $zip $sheetRel.Target
                if ($null -eq $xml) {
                    continue
                }
                if ($sheetRel.Type -like "*/drawing") {
                    $shapes.AddRange([object[]]@(readXlsxShapeRows $xml $zip $sheetRel.Target $failures $sizeFailures))
                } elseif ($sheetRel.Type -like "*/comments") {
                    $commentsXml = $xml
                } elseif ($sheetRel.Type -like "*/threadedComment") {
                    $threadedXmls.Add($xml)
                }
            }
            addSortedShapeLines $units "${sheetName}[図形]" $shapes

            $comments = readXlsxCommentRows $commentsXml $threadedXmls.ToArray()
            $lines = @($comments.Keys | Sort-Object { (getCellPosition $_)[0] }, { (getCellPosition $_)[1] } |
                ForEach-Object { "$($_.Replace('$', ''))`t$(toObjectCellText @($comments[$_] -split "\r\n|\r|\n"))" })
            if ($lines.Count -gt 0) {
                addUnitLines $units "${sheetName}[コメント]" $lines
            }

            $headerFooter = @(readXlsxSheetHeaderFooter $zip $rel.Target $null $failures)
            if ($headerFooter.Count -gt 0) {
                addUnitLines $units "${sheetName}[ヘッダー・フッター]" $headerFooter
            }
        }
    } finally {
        $zip.Dispose()
    }
    return $units
}

function writeUnits {
    # ユニットごとに "<場所>.tsv" を出力し、出力したファイル数を返す（空のユニットは出力しない）。
    # 元のファイル名は、出力先のフォルダ名（インデクサが作業フォルダから移すときのフォルダ）になる
    param (
        [System.Collections.Specialized.OrderedDictionary]$units,
        [string]$outDir
    )

    $count = 0
    foreach ($unitName in $units.Keys) {
        $lines = New-Object System.Collections.Generic.List[string]
        foreach ($line in $units[$unitName]) {
            $line = $line.TrimEnd()
            if ($line -ne "") {
                $lines.Add($line)
            }
        }
        if ($lines.Count -eq 0) {
            continue
        }

        $path = Join-Path $outDir (toIndexFileName $unitName)
        [System.IO.File]::WriteAllLines((toLongPath $path), $lines, ${utf8Bom})
        $count++
    }
    return $count
}
