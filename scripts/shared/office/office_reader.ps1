# Word（.docx / .docm）・PowerPoint（.pptx / .pptm）のファイルと、Excel（.xlsx / .xlsm）の図形・コメントからテキストを読み出す。
# ファイルはZIP（Office Open XML）として直接読むため、Word・PowerPoint・Excelは使わない。
# インデクサ・テストから dot-source して使う。共通の部品（shared.ps1）を先に読み込んでおくこと。
#
# 読み出した結果は「場所 → 行の一覧」の順序付き辞書（ユニット）で返す。
#   Word      : ページ001, ページ001[図形], ページ001[コメント], ページ002, ..., ヘッダー・フッター, 脚注
#   PowerPoint: スライド001, スライド001[図形], スライド001[コメント], スライド001_ノート, スライド002（非表示）, ..., ヘッダー・フッター
#   Excel     : <シート名>[図形], <シート名>[コメント]（セルの値はインデクサが Excel で読む）
# 1行は段落1つ、または表の1行（セルをタブ区切り）。図形・コメントの場所は図形・コメント1つ
# （Word・PowerPoint は文字だけ、Excel は "<セル番地><TAB><文字>"）。

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

# PowerPointで読み飛ばすプレースホルダー（スライド番号・日付・ヘッダー・フッター・スライド画像）
${skipPlaceholderTypes} = @("sldNum", "dt", "hdr", "ftr", "sldImg")

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

function readZipEntry {
    # ZIP内のファイルを文字列で返す。無ければ $null
    param (
        [System.IO.Compression.ZipArchive]$zip,
        [string]$entryName
    )

    $entry = $zip.GetEntry($entryName)
    if ($null -eq $entry) {
        return $null
    }
    $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    try {
        return $reader.ReadToEnd()
    } finally {
        $reader.Dispose()
    }
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

    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($xml)
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

function readChartText {
    # グラフ（charts/chartN.xml）の文字（タイトル・軸ラベル・系列名・項目名）を、スペースでつないで 1 つにする。
    # 数値（numCache）は読まない。同じ文字は 1 回だけにする
    param (
        [string]$xml
    )

    $texts = New-Object System.Collections.Generic.List[string]
    $seen = New-Object System.Collections.Generic.HashSet[string]
    # タイトル・軸ラベル（c:rich の a:p）
    foreach ($line in (readXmlLines $xml ${nsDrawing})) {
        if ($seen.Add($line.Text)) { $texts.Add($line.Text) }
    }
    # 系列名・項目名（c:strCache / c:multiLvlStrCache の c:pt/c:v）
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($xml)
    foreach ($v in $doc.GetElementsByTagName("v", ${nsChart})) {
        $cache = $v.ParentNode.ParentNode
        if ($cache.LocalName -eq "lvl") { $cache = $cache.ParentNode }
        if ($cache.LocalName -notin @("strCache", "multiLvlStrCache")) { continue }
        $text = $v.InnerText.Trim()
        if ($text -ne "" -and $seen.Add($text)) { $texts.Add($text) }
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
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($xml)
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
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($xml)
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
    # コメントは "ページNNN[コメント]"（コメントを付けた所のページ）に分ける（1 行は図形・コメント 1 つ）
    param (
        [string]$path
    )

    $units = New-Object System.Collections.Specialized.OrderedDictionary
    $zip = [System.IO.Compression.ZipFile]::OpenRead((toLongPath $path))
    try {
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

        # 図形（テキストボックス・SmartArt・グラフ）とコメント。文書の中の順に、そのページの場所へ入れる
        $rels = readRelationships $zip "word/document.xml"
        $comments = readWordComments (readZipEntry $zip "word/comments.xml")
        $usedComments = New-Object System.Collections.Generic.HashSet[string]
        foreach ($object in $objects) {
            $base = "ページ{0:D3}" -f $object.Page
            if ($object.Kind -eq "shape") {
                addUnitLines $units "${base}[図形]" @($object.Text)
            } elseif ($object.Kind -eq "comment") {
                $id = [string]$object.Id
                if ($comments.ContainsKey($id) -and $usedComments.Add($id)) {
                    addUnitLines $units "${base}[コメント]" @($comments[$id])
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
    } finally {
        $zip.Dispose()
    }
    return $units
}

function readPptxUnits {
    # PowerPoint（.pptx / .pptm）のテキストを、スライド・ノートごと（スライドの表示順）と、
    # スライドのフッター（全スライド分をまとめ、重複を除く）に分けて返す
    param (
        [string]$path
    )

    $units = New-Object System.Collections.Specialized.OrderedDictionary
    $footers = New-Object System.Collections.Generic.List[string]
    $seenFooters = New-Object System.Collections.Generic.HashSet[string]
    $zip = [System.IO.Compression.ZipFile]::OpenRead((toLongPath $path))
    try {
        $presentationXml = readZipEntry $zip "ppt/presentation.xml"
        if ($null -eq $presentationXml) {
            throw "PowerPointのプレゼンテーション情報（ppt/presentation.xml）がありません。"
        }
        $presentation = New-Object System.Xml.XmlDocument
        $presentation.LoadXml($presentationXml)
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
        }

        if ($footers.Count -gt 0) {
            addUnitLines $units "ヘッダー・フッター" $footers.ToArray()
        }
    } finally {
        $zip.Dispose()
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
    # 行・列は図形の左上のセル（1 から数える）。グループ化した図形は、まとめて 1 つの図形とする
    param (
        [string]$xml
    )

    $nsSheetDrawing = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($xml)
    $rows = New-Object System.Collections.Generic.List[object]
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

        $lines = @(readXmlLines $anchor.OuterXml ${nsDrawing} | ForEach-Object { $_.Text })
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
        $rows.Add(@{ Row = $row; Column = $column; Text = (toObjectCellText $lines) })
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
        $doc = New-Object System.Xml.XmlDocument
        $doc.LoadXml($xml)
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
        $doc = New-Object System.Xml.XmlDocument
        $doc.LoadXml($commentsXml)
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

function readXlsxObjectUnits {
    # Excel（.xlsx / .xlsm）の表示シートにある図形（テキストボックス・グループ・WordArt を含む）とコメントの文字を、
    # "<シート名>[図形]" "<シート名>[コメント]" の場所ごとに返す（シート名には [ ] を使えないため、実在のシートと重ならない）。
    # 1 行は "<セル番地><TAB><文字>"。セル番地は図形の左上・コメントのセルで、シートの上の行から順に並べる。
    # セルの値は Excel のテキスト保存で読むため、ここでは読まない。
    # ZIP の中身が Excel のブック（xl/workbook.xml）でなければ（.xlsb など）、何も返さない
    param (
        [string]$path
    )

    $units = New-Object System.Collections.Specialized.OrderedDictionary
    $nsSheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    $zip = [System.IO.Compression.ZipFile]::OpenRead((toLongPath $path))
    try {
        $workbookXml = readZipEntry $zip "xl/workbook.xml"
        if ($null -eq $workbookXml) {
            return $units
        }
        $workbook = New-Object System.Xml.XmlDocument
        $workbook.LoadXml($workbookXml)
        $workbookRels = readRelationships $zip "xl/workbook.xml"

        foreach ($sheet in $workbook.GetElementsByTagName("sheet", $nsSheet)) {
            # 非表示（hidden）・完全に非表示（veryHidden）のシートは、セルと同じく読まない
            if ($sheet.GetAttribute("state") -in @("hidden", "veryHidden")) {
                continue
            }
            $rel = $workbookRels[$sheet.GetAttribute("id", ${nsRel})]
            if ($null -eq $rel -or $rel.Type -notlike "*/worksheet") {
                continue  # グラフシートなど
            }
            $sheetName = $sheet.GetAttribute("name")

            $shapes = New-Object System.Collections.Generic.List[object]
            $commentsXml = $null
            $threadedXmls = New-Object System.Collections.Generic.List[string]
            foreach ($sheetRel in (readRelationships $zip $rel.Target).Values) {
                $xml = readZipEntry $zip $sheetRel.Target
                if ($null -eq $xml) {
                    continue
                }
                if ($sheetRel.Type -like "*/drawing") {
                    $shapes.AddRange([object[]]@(readXlsxShapeRows $xml))
                } elseif ($sheetRel.Type -like "*/comments") {
                    $commentsXml = $xml
                } elseif ($sheetRel.Type -like "*/threadedComment") {
                    $threadedXmls.Add($xml)
                }
            }

            # 上の行から順に（同じ行は左から、同じセルは XML の順に）並べる。
            # Sort-Object は同じキーの順を保たない（インデックス作成のたびに順が変わる）ため、XML の順もキーにする
            for ($i = 0; $i -lt $shapes.Count; $i++) {
                $shapes[$i].Order = $i
            }
            $lines = @($shapes | Sort-Object { $_.Row }, { $_.Column }, { $_.Order } |
                ForEach-Object { "$(toColumnName $_.Column)$($_.Row)`t$($_.Text)" })
            if ($lines.Count -gt 0) {
                addUnitLines $units "${sheetName}[図形]" $lines
            }

            $comments = readXlsxCommentRows $commentsXml $threadedXmls.ToArray()
            $lines = @($comments.Keys | Sort-Object { (getCellPosition $_)[0] }, { (getCellPosition $_)[1] } |
                ForEach-Object { "$($_.Replace('$', ''))`t$(toObjectCellText @($comments[$_] -split "\r\n|\r|\n"))" })
            if ($lines.Count -gt 0) {
                addUnitLines $units "${sheetName}[コメント]" $lines
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
