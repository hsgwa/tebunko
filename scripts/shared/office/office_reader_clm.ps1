# Office ファイル（Office Open XML）の文字を、制限言語モード（ConstrainedLanguage）で読む。
#
# いつものインデクサは office_reader.ps1（System.IO.Compression の ZipArchive・System.Xml.XmlReader・List）で読むが、
# 制限言語モードではどれも使えない。ここでは次の方法で、同じ場所（ユニット）に同じ行を返す:
#   ・ZIP は Windows に付いている tar.exe で作業フォルダに取り出す（openOfficeZip）。
#     tar.exe は引数を ANSI のコードページで受け取るため、元のファイルは cmd のリダイレクトで標準入力から渡し、
#     取り出し先には cmd で移ってから取り出す（日本語・CP932 に無い文字のパスでも読める）
#   ・XML は正規表現で字句（開始・終了のタグ、文字）に分け、office_reader.ps1 の readXmlLines と同じ状態機械で読む
#     （readXmlLinesClm）。名前空間の接頭辞（xmlns）と文字参照（&amp; など）もここで解く
#   ・関係（.rels）・コメント・図形の位置などの小さな XML は [xml] で読む。要素のメソッド（GetAttribute など）は
#     制限言語モードで使えないため、属性は getXmlAttribute、子孫は getXmlDescendants でたどる
# 同じ結果になることは tests\shared\office\office_reader_clm.Tests.ps1 が、テスト用のファイルで office_reader.ps1 と突き合わせる。
# ユニットは newUnitsClm で作り、getUnitsClm で「場所 → 行の配列」の順序付き辞書にする。

${nsWordClm}     = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
${nsDrawingClm}  = "http://schemas.openxmlformats.org/drawingml/2006/main"
${nsPresentClm}  = "http://schemas.openxmlformats.org/presentationml/2006/main"
${nsCompatClm}   = "http://schemas.openxmlformats.org/markup-compatibility/2006"
${nsRelClm}      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
${nsDiagramClm}  = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
${nsChartClm}    = "http://schemas.openxmlformats.org/drawingml/2006/chart"
${nsSheetClm}    = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
${nsSheetDrawingClm} = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
${nsThreadedClm} = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"

# PowerPoint で読み飛ばすプレースホルダー（office_reader.ps1 の skipPlaceholderTypes と同じ）
${skipPlaceholderTypesClm} = @("sldNum", "dt", "hdr", "ftr", "sldImg")

# XML の字句（コメント・処理命令・CDATA・DOCTYPE・終了タグ・開始タグ・文字）
${xmlTokenPattern} = New-Object regex -ArgumentList @(
    '<!--.*?-->|<\?.*?\?>|<!\[CDATA\[(?<cdata>.*?)\]\]>|<!DOCTYPE[^>]*>|</(?<end>[^\s>]+)\s*>|<(?<name>[^\s/>]+)(?<attrs>(?:\s+[^\s=/>]+\s*=\s*(?:"[^"]*"|''[^'']*''))*)\s*(?<empty>/)?>|(?<text>[^<]+)',
    'Singleline')
${xmlAttributePattern} = New-Object regex -ArgumentList @('([^\s=/>]+)\s*=\s*(?:"([^"]*)"|''([^'']*)'')', 'None')

# tar.exe（PATH の先に Git の GNU tar があることがあるため、フルパスで呼ぶ）
${officeTarExe} = Join-Path $env:windir "System32\tar.exe"


function decodeXmlText {
    # XML の文字参照・実体参照（&lt; &gt; &amp; &quot; &apos; &#N; &#xH;）を文字に戻す
    param (
        [string]$text
    )

    if ($text.IndexOf("&") -lt 0) {
        return $text
    }
    $parts = [regex]::Split($text, '(&(?:lt|gt|amp|quot|apos|#[0-9]+|#x[0-9A-Fa-f]+);)')
    for ($i = 1; $i -lt $parts.Count; $i += 2) {
        $entity = $parts[$i]
        switch ($entity) {
            "&lt;" { $parts[$i] = "<" }
            "&gt;" { $parts[$i] = ">" }
            "&amp;" { $parts[$i] = "&" }
            "&quot;" { $parts[$i] = '"' }
            "&apos;" { $parts[$i] = "'" }
            default {
                $code = if ($entity.StartsWith("&#x")) { [int]("0x" + $entity.Substring(3, $entity.Length - 4)) } else { [int]$entity.Substring(2, $entity.Length - 3) }
                if ($code -ge 0x10000) {
                    # サロゲートペア
                    $code -= 0x10000
                    $parts[$i] = [string][char](0xD800 + ($code -shr 10)) + [string][char](0xDC00 + ($code -band 0x3FF))
                } else {
                    $parts[$i] = [string][char]$code
                }
            }
        }
    }
    return ($parts -join "")
}

function getXmlAttribute {
    # [xml] で読んだ要素の属性の値（名前空間つき）。無ければ空。要素の GetAttribute は制限言語モードで使えないため、属性を順に見る
    param (
        $node,
        [string]$localName,
        [string]$namespace = ""
    )

    foreach ($attribute in $node.Attributes) {
        if ($attribute.LocalName -eq $localName -and $attribute.NamespaceURI -eq $namespace) {
            return $attribute.Value
        }
    }
    return ""
}

function getXmlDescendants {
    # [xml] で読んだ要素の子孫のうち、名前（と名前空間）が一致する要素を文書の順に返す（要素の GetElementsByTagName の代わり）
    param (
        $node,
        [string]$localName,
        $namespace = $null  # $null は名前空間を問わない（[string] にすると $null が空文字列になる）
    )

    $found = @()
    $stack = @($node.ChildNodes)
    # 文書の順にたどる（後ろから積み、前から取り出す）
    $pending = @{}
    $count = 0
    for ($i = $stack.Count - 1; $i -ge 0; $i--) { $pending[$count] = $stack[$i]; $count++ }
    while ($count -gt 0) {
        $count--
        $current = $pending[$count]
        $pending.Remove($count)
        if ($current.NodeType -ne "Element") { continue }
        if ($current.LocalName -eq $localName -and ($null -eq $namespace -or $current.NamespaceURI -eq $namespace)) {
            $found += $current
        }
        $children = @($current.ChildNodes)
        for ($i = $children.Count - 1; $i -ge 0; $i--) { $pending[$count] = $children[$i]; $count++ }
    }
    return $found
}

function openOfficeZip {
    # Office ファイル（ZIP）を作業フォルダに取り出し、@{ Dir（作業フォルダ） } を返す（使い終わったら closeOfficeZip）。
    # 画像・埋め込みのファイル（media・embeddings）は文字が無く大きいため取り出さない。ZIP として読めなければ例外
    param (
        [string]$path,
        [string]$workRoot
    )

    $dir = Join-Path $workRoot ("zip_" + $PID + "_" + (Get-Date).Ticks)
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    $source = fromLongPath $path
    $ErrorActionPreference = "Continue"
    $output = cmd.exe /d /c "cd /d `"$dir`" && `"${officeTarExe}`" -xf - --exclude `"*/media/*`" --exclude `"*/embeddings/*`" < `"$source`"" 2>&1
    $code = $LASTEXITCODE
    $ErrorActionPreference = "Stop"
    if ($code -ne 0) {
        Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
        throw "ZIP として読めません（tar.exe の終了コード $code）: $(@($output) -join ' ')"
    }
    return @{ Dir = $dir }
}

function closeOfficeZip {
    param (
        $zip
    )

    if ($zip -and $zip.Dir -and (Test-Path -LiteralPath $zip.Dir)) {
        Remove-Item -LiteralPath $zip.Dir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function readOfficeZipEntry {
    # 取り出した ZIP の中のファイルを文字列で返す（ZIP の中のパスは / 区切り）。無ければ $null
    param (
        $zip,
        [string]$entryName
    )

    $path = Join-Path $zip.Dir $entryName.Replace("/", "\")
    if (!(Test-Path -LiteralPath $path -PathType Leaf)) {
        return $null
    }
    $text = Get-Content -LiteralPath $path -Raw -Encoding UTF8
    if ($null -eq $text) {
        return ""
    }
    return $text
}

function getOfficeZipEntries {
    # 取り出した ZIP の中のファイルの一覧（/ 区切りのパス）
    param (
        $zip
    )

    $length = $zip.Dir.Length + 1
    return @(Get-ChildItem -LiteralPath $zip.Dir -Recurse -File | ForEach-Object { $_.FullName.Substring($length).Replace("\", "/") })
}

function readRelationshipsClm {
    # .rels を読み、Id → @{ Type; Target（ZIP の中のパス） } を返す（office_reader.ps1 の readRelationships と同じ）
    param (
        $zip,
        [string]$partName
    )

    $slash = $partName.LastIndexOf("/")
    $dir = if ($slash -ge 0) { $partName.Substring(0, $slash) } else { "" }
    $file = $partName.Substring($slash + 1)
    $result = @{}
    $xml = readOfficeZipEntry $zip "$dir/_rels/$file.rels"
    if ($null -eq $xml) {
        return $result
    }
    [xml]$doc = $xml
    foreach ($rel in $doc.GetElementsByTagName("Relationship", "http://schemas.openxmlformats.org/package/2006/relationships")) {
        if ((getXmlAttribute $rel "TargetMode") -eq "External") {
            continue
        }
        $result[(getXmlAttribute $rel "Id")] = @{
            Type   = getXmlAttribute $rel "Type"
            Target = resolveZipPath $dir (getXmlAttribute $rel "Target")
        }
    }
    return $result
}

function readXmlLinesClm {
    # office_reader.ps1 の readXmlLines と同じ行を返す（@{ Page; Text } の配列）。XmlReader の代わりに、正規表現で分けた字句を順に読む。
    # 引数と規則は readXmlLines を参照。$objects には、読んだ図形・コメントの参照を 0 から番号で入れる（ハッシュテーブル）
    param (
        [string]$xml,
        [string]$ns,
        [string]$pageMode = "none",
        [string[]]$onlyPlaceholders = $null,
        [hashtable]$objects = $null
    )

    # XML の決まりどおり、改行を LF にそろえてから読む（XmlReader も同じ）
    $xml = $xml.Replace("`r`n", "`n").Replace("`r", "`n")

    # 積むものは番号 → 中身のハッシュテーブルで持つ（List は制限言語モードで使えない）
    $lines = @{}; $lineCount = 0
    $pFrames = @{}; $pCount = 0
    $trFrames = @{}; $trCount = 0
    $tcFrames = @{}; $tcCount = 0
    $boxFrames = @{}; $boxCount = 0
    $page = 1
    $breakPending = $false
    $inText = $false
    $inSectPr = $false
    $skipShapeText = $false
    $collect = ($null -ne $objects)
    $onlyMode = ($null -ne $onlyPlaceholders)
    # 名前空間の接頭辞 → URI。要素の深さごとに、宣言があれば新しい対応表を積む
    $scopes = @{}
    $scopes[0] = @{ "xml" = "http://www.w3.org/XML/1998/namespace" }
    $depth = 0
    $skipDepth = -1   # 読み飛ばしている要素の深さ（-1 は読み飛ばしていない）

    foreach ($m in ${xmlTokenPattern}.Matches($xml)) {
        $groups = $m.Groups
        $textGroup = $groups["text"]
        $cdataGroup = $groups["cdata"]
        $endGroup = $groups["end"]
        $nameGroup = $groups["name"]

        if ($nameGroup.Success) {
            # 開始タグ
            $isEmpty = $groups["empty"].Success
            if ($skipDepth -ge 0) {
                if (!$isEmpty) { $depth++ }
                continue
            }
            # 属性と、名前空間の宣言
            $attributes = @{}
            $map = $scopes[$depth]
            $newMap = $null
            $attrText = $groups["attrs"].Value
            if ($attrText -ne "") {
                foreach ($a in ${xmlAttributePattern}.Matches($attrText)) {
                    $qname = $a.Groups[1].Value
                    $value = if ($a.Groups[2].Success) { $a.Groups[2].Value } else { $a.Groups[3].Value }
                    if ($qname -eq "xmlns" -or $qname.StartsWith("xmlns:")) {
                        if ($null -eq $newMap) { $newMap = @{}; foreach ($k in $map.Keys) { $newMap[$k] = $map[$k] } }
                        $prefix = if ($qname -eq "xmlns") { "" } else { $qname.Substring(6) }
                        $newMap[$prefix] = decodeXmlText $value
                    } else {
                        $attributes[$qname] = $value
                    }
                }
            }
            if ($null -ne $newMap) { $map = $newMap }
            $qualified = $nameGroup.Value
            $colon = $qualified.IndexOf(":")
            if ($colon -ge 0) {
                $name = $qualified.Substring($colon + 1)
                $uri = [string]$map[$qualified.Substring(0, $colon)]
            } else {
                $name = $qualified
                $uri = [string]$map[""]
            }
            # 属性を「名前空間 URI|ローカル名」で引けるようにする（接頭辞の無い属性は名前空間なし）
            $attr = @{}
            foreach ($key in $attributes.Keys) {
                $c = $key.IndexOf(":")
                if ($c -ge 0) {
                    $attr[[string]$map[$key.Substring(0, $c)] + "|" + $key.Substring($c + 1)] = decodeXmlText $attributes[$key]
                } else {
                    $attr["|" + $key] = decodeXmlText $attributes[$key]
                }
            }
            if (!$isEmpty) {
                $depth++
                $scopes[$depth] = $map
            }

            # 読み飛ばす要素
            $skip = $false
            if ($uri -eq ${nsCompatClm} -and $name -eq "Fallback") {
                $skip = $true
            } elseif ($uri -eq ${nsWordClm} -and $name -eq "moveFrom") {
                $skip = $true
            } elseif ($uri -eq ${nsPresentClm} -and $name -eq "txBody" -and $skipShapeText) {
                $skip = $true
            } elseif ($uri -eq ${nsPresentClm} -and $name -eq "graphicFrame" -and $onlyMode) {
                $skip = $true
            } elseif ($uri -eq $ns -and $name -eq "rPr" -and -not $isEmpty) {
                $skip = $true
            }
            if ($skip) {
                if (!$isEmpty) { $skipDepth = $depth }
                continue
            }

            if ($uri -eq $ns) {
                if ($name -eq "p") {
                    if (-not $isEmpty) {
                        $pFrames[$pCount] = @{ Text = ""; Page = $null; BreakAfter = $false }; $pCount++
                    }
                } elseif ($name -eq "tr") {
                    if (-not $isEmpty) {
                        $trFrames[$trCount] = @{ Cells = @{}; CellCount = 0; Page = $null }; $trCount++
                    }
                } elseif ($name -eq "tc") {
                    if ($isEmpty) {
                        if ($trCount -gt 0) { $row = $trFrames[$trCount - 1]; $row.Cells[$row.CellCount] = ""; $row.CellCount++ }
                    } else {
                        $tcFrames[$tcCount] = @{ Text = ""; Page = $null }; $tcCount++
                    }
                } elseif ($name -eq "t") {
                    $inText = -not $isEmpty
                } elseif ($name -eq "tab" -or $name -eq "cr" -or $name -eq "noBreakHyphen") {
                    if ($pCount -gt 0) {
                        $pFrames[$pCount - 1].Text += $(if ($name -eq "noBreakHyphen") { "-" } else { " " })
                    }
                } elseif ($name -eq "br") {
                    $breakType = $attr["$ns|type"]
                    if ($breakType -ne "page" -and $breakType -ne "column") {
                        if ($pCount -gt 0) { $pFrames[$pCount - 1].Text += " " }
                    } elseif ($pageMode -ne "none" -and $breakType -eq "page") {
                        $page++
                        $breakPending = $true
                    }
                } elseif ($name -eq "lastRenderedPageBreak") {
                    if ($pageMode -eq "rendered") {
                        if ($breakPending) { $breakPending = $false } else { $page++ }
                    }
                } elseif ($name -eq "pageBreakBefore") {
                    if ($pageMode -ne "none" -and $attr["$ns|val"] -notin @("0", "false", "off")) {
                        $page++
                        $breakPending = $true
                    }
                } elseif ($name -eq "sectPr") {
                    if ($pageMode -ne "none" -and $pCount -gt 0) {
                        $inSectPr = -not $isEmpty
                        $pFrames[$pCount - 1].BreakAfter = $true
                    }
                } elseif ($name -eq "type") {
                    if ($inSectPr -and $attr["$ns|val"] -eq "continuous") {
                        if ($pCount -gt 0) { $pFrames[$pCount - 1].BreakAfter = $false }
                    }
                } elseif ($collect -and $name -eq "txbxContent") {
                    if (-not $isEmpty) {
                        $boxFrames[$boxCount] = @{ Lines = @{}; LineCount = 0; Page = $page; TcBase = $tcCount }; $boxCount++
                    }
                } elseif ($collect -and $name -eq "commentReference") {
                    $objects[$objects.Count] = @{ Kind = "comment"; Page = $page; Id = $attr["$ns|id"] }
                }
            } elseif ($collect -and $uri -eq ${nsDiagramClm} -and $name -eq "relIds") {
                $objects[$objects.Count] = @{ Kind = "diagram"; Page = $page; RelId = $attr["${nsRelClm}|dm"] }
            } elseif ($collect -and $uri -eq ${nsChartClm} -and $name -eq "chart") {
                $objects[$objects.Count] = @{ Kind = "chart"; Page = $page; RelId = $attr["${nsRelClm}|id"] }
            } elseif ($uri -eq ${nsPresentClm}) {
                if ($name -eq "sp") {
                    $skipShapeText = $onlyMode
                } elseif ($name -eq "ph") {
                    if ($onlyMode) {
                        $skipShapeText = ($attr["|type"] -notin $onlyPlaceholders)
                    } else {
                        $skipShapeText = ($attr["|type"] -in ${skipPlaceholderTypesClm})
                    }
                }
            }

            # 中身の無い要素（<w:tab/> など）は、XmlReader では終了タグが来ないため、ここで終わりにする
            continue
        }

        if ($endGroup.Success) {
            # 終了タグ
            if ($skipDepth -ge 0) {
                if ($depth -eq $skipDepth) { $skipDepth = -1 }
                $depth--
                continue
            }
            $map = $scopes[$depth]
            $qualified = $endGroup.Value
            $colon = $qualified.IndexOf(":")
            if ($colon -ge 0) {
                $name = $qualified.Substring($colon + 1)
                $uri = [string]$map[$qualified.Substring(0, $colon)]
            } else {
                $name = $qualified
                $uri = [string]$map[""]
            }
            $scopes.Remove($depth)
            $depth--
            if ($uri -ne $ns) { continue }

            $text = $null
            if ($name -eq "t") {
                $inText = $false
            } elseif ($name -eq "sectPr") {
                $inSectPr = $false
            } elseif ($name -eq "p") {
                $pCount--
                $p = $pFrames[$pCount]
                $pFrames.Remove($pCount)
                $text = $p.Text.Trim()
                $textPage = $p.Page
                if ($p.BreakAfter) {
                    $page++
                    $breakPending = $true
                }
            } elseif ($name -eq "tc") {
                $tcCount--
                $cell = $tcFrames[$tcCount]
                $tcFrames.Remove($tcCount)
                if ($trCount -gt 0) {
                    $row = $trFrames[$trCount - 1]
                    $row.Cells[$row.CellCount] = $cell.Text; $row.CellCount++
                    if ($null -eq $row.Page) { $row.Page = $cell.Page }
                }
            } elseif ($name -eq "tr") {
                $trCount--
                $row = $trFrames[$trCount]
                $trFrames.Remove($trCount)
                $cellBase = if ($boxCount -gt 0) { $boxFrames[$boxCount - 1].TcBase } else { 0 }
                $separator = if ($tcCount -gt $cellBase) { " " } else { "`t" }
                $cells = @(for ($i = 0; $i -lt $row.CellCount; $i++) { $row.Cells[$i] })
                $text = ($cells -join $separator).TrimEnd()
                $textPage = if ($null -ne $row.Page) { $row.Page } else { $page }
            } elseif ($name -eq "txbxContent" -and $boxCount -gt 0) {
                $boxCount--
                $box = $boxFrames[$boxCount]
                $boxFrames.Remove($boxCount)
                if ($box.LineCount -gt 0) {
                    $boxLines = @(for ($i = 0; $i -lt $box.LineCount; $i++) { $box.Lines[$i] })
                    $objects[$objects.Count] = @{ Kind = "shape"; Page = $box.Page; Text = ($boxLines -join " ") }
                }
            }
            if ($text) {
                $box = $null
                if ($boxCount -gt 0 -and $tcCount -le $boxFrames[$boxCount - 1].TcBase) {
                    $box = $boxFrames[$boxCount - 1]
                }
                if ($box) {
                    $box.Lines[$box.LineCount] = $text; $box.LineCount++
                } elseif ($tcCount -gt 0) {
                    $cell = $tcFrames[$tcCount - 1]
                    if ($cell.Text.Length -gt 0) { $cell.Text += " " }
                    $cell.Text += $text
                    if ($null -eq $cell.Page) { $cell.Page = $textPage }
                } else {
                    $lines[$lineCount] = @{ Page = $textPage; Text = $text }; $lineCount++
                }
            }
            continue
        }

        if ($skipDepth -ge 0 -or !$inText) {
            continue
        }
        if ($textGroup.Success) {
            $value = decodeXmlText $textGroup.Value
        } elseif ($cdataGroup.Success) {
            $value = $cdataGroup.Value
        } else {
            continue  # コメント・処理命令
        }
        if ($pCount -gt 0) {
            $p = $pFrames[$pCount - 1]
            if ($null -eq $p.Page) { $p.Page = $page }
            $p.Text += $value.Replace("`t", " ").Replace("`r", " ").Replace("`n", " ")
            if ($value.Trim() -ne "") { $breakPending = $false }
        }
    }

    return @(for ($i = 0; $i -lt $lineCount; $i++) { $lines[$i] })
}

function newUnitsClm {
    # ユニット（場所 → 行）の入れ物。場所は足した順に並ぶ
    return @{ Names = @{}; NameCount = 0; Lines = @{} }
}

function addUnitLinesClm {
    # ユニットに行を足す（office_reader.ps1 の addUnitLines と同じ）
    param (
        [hashtable]$units,
        [string]$unitName,
        [string[]]$lines
    )

    if (!$units.Lines.ContainsKey($unitName)) {
        $units.Names[$units.NameCount] = $unitName
        $units.NameCount++
        $units.Lines[$unitName] = @{ N = 0 }
    }
    $list = $units.Lines[$unitName]
    foreach ($line in $lines) {
        $list[$list.N] = $line
        $list.N++
    }
}

function getUnitsClm {
    # ユニットを「場所 → 行の配列」の順序付き辞書にする
    param (
        [hashtable]$units
    )

    $result = [ordered]@{}
    for ($i = 0; $i -lt $units.NameCount; $i++) {
        $name = $units.Names[$i]
        $list = $units.Lines[$name]
        $result[$name] = [string[]]@(for ($j = 0; $j -lt $list.N; $j++) { $list[$j] })
    }
    return $result
}

function readDiagramTextClm {
    param (
        [string]$xml
    )

    return (@(readXmlLinesClm $xml ${nsDrawingClm} | ForEach-Object { $_.Text }) -join " ")
}

function readChartTextClm {
    # グラフの文字（タイトル・軸ラベル・系列名・項目名）。office_reader.ps1 の readChartText と同じ
    param (
        [string]$xml
    )

    $texts = @()
    # HashSet（大文字・小文字を区別する）と同じにするため、@{} ではなく区別するハッシュテーブルを使う
    $seen = New-Object System.Collections.Hashtable
    foreach ($line in (readXmlLinesClm $xml ${nsDrawingClm})) {
        if (!$seen.ContainsKey($line.Text)) { $seen[$line.Text] = $true; $texts += $line.Text }
    }
    [xml]$doc = $xml
    foreach ($v in $doc.GetElementsByTagName("v", ${nsChartClm})) {
        $cache = $v.ParentNode.ParentNode
        if ($cache.LocalName -eq "lvl") { $cache = $cache.ParentNode }
        if ($cache.LocalName -notin @("strCache", "multiLvlStrCache")) { continue }
        $text = $v.InnerText.Trim()
        if ($text -ne "" -and !$seen.ContainsKey($text)) { $seen[$text] = $true; $texts += $text }
    }
    return ($texts -join " ")
}

function readObjectTextClm {
    param (
        $zip,
        [hashtable]$rels,
        $object
    )

    $rel = $rels[[string]$object.RelId]
    if ($null -eq $rel) {
        return ""
    }
    $xml = readOfficeZipEntry $zip $rel.Target
    if ($null -eq $xml) {
        return ""
    }
    if ($object.Kind -eq "diagram") {
        return (readDiagramTextClm $xml)
    }
    return (readChartTextClm $xml)
}

function readWordCommentsClm {
    # Word のコメントを w:id → 文字の辞書で返す（office_reader.ps1 の readWordComments と同じ）
    param (
        [string]$xml
    )

    $comments = @{}
    if (-not $xml) {
        return $comments
    }
    [xml]$doc = $xml
    foreach ($comment in $doc.GetElementsByTagName("comment", ${nsWordClm})) {
        $text = @(readXmlLinesClm $comment.OuterXml ${nsWordClm} | ForEach-Object { $_.Text }) -join " "
        if ($text -ne "") {
            $comments[(getXmlAttribute $comment "id" ${nsWordClm})] = $text
        }
    }
    return $comments
}

function readSlideCommentsClm {
    # PowerPoint のスライドのコメント（office_reader.ps1 の readSlideComments と同じ）
    param (
        [string]$xml
    )

    $texts = @()
    [xml]$doc = $xml
    foreach ($cm in @($doc.SelectNodes("//*") | Where-Object { $_.LocalName -eq "cm" })) {
        if ($cm.NamespaceURI -eq ${nsPresentClm}) {
            $body = @($cm.ChildNodes | Where-Object { $_.LocalName -eq "text" })
            if ($body.Count -gt 0 -and $body[0].InnerText.Trim() -ne "") { $texts += $body[0].InnerText.Trim() }
            continue
        }
        $bodies = @($cm.ChildNodes | Where-Object { $_.LocalName -eq "txBody" })
        foreach ($reply in @($cm.ChildNodes | Where-Object { $_.LocalName -eq "replyLst" } | ForEach-Object { $_.ChildNodes } | Where-Object { $_.LocalName -eq "reply" })) {
            $bodies += @($reply.ChildNodes | Where-Object { $_.LocalName -eq "txBody" })
        }
        foreach ($body in $bodies) {
            $text = @(readXmlLinesClm $body.OuterXml ${nsDrawingClm} | ForEach-Object { $_.Text }) -join " "
            if ($text -ne "") { $texts += $text }
        }
    }
    return $texts
}

function readDocxUnitsClm {
    # Word（.docx / .docm）の文字を、ページ・ヘッダー/フッター・脚注ごとに返す（office_reader.ps1 の readDocxUnits と同じ）
    param (
        [string]$path,
        [string]$workRoot
    )

    $units = newUnitsClm
    $zip = openOfficeZip $path $workRoot
    try {
        $body = readOfficeZipEntry $zip "word/document.xml"
        if ($null -eq $body) {
            throw "Word文書の本文（word/document.xml）がありません。"
        }
        $pageMode = $(if ($body.Contains("lastRenderedPageBreak")) { "rendered" } else { "explicit" })
        $objects = @{}
        foreach ($line in (readXmlLinesClm $body ${nsWordClm} $pageMode $null $objects)) {
            addUnitLinesClm $units ("ページ{0:D3}" -f $line.Page) @($line.Text)
        }

        $rels = readRelationshipsClm $zip "word/document.xml"
        $comments = readWordCommentsClm (readOfficeZipEntry $zip "word/comments.xml")
        $usedComments = @{}
        for ($i = 0; $i -lt $objects.Count; $i++) {
            $object = $objects[$i]
            $base = "ページ{0:D3}" -f $object.Page
            if ($object.Kind -eq "shape") {
                addUnitLinesClm $units "${base}[図形]" @($object.Text)
            } elseif ($object.Kind -eq "comment") {
                $id = [string]$object.Id
                if ($comments.ContainsKey($id) -and !$usedComments.ContainsKey($id)) {
                    $usedComments[$id] = $true
                    addUnitLinesClm $units "${base}[コメント]" @($comments[$id])
                }
            } else {
                $text = readObjectTextClm $zip $rels $object
                if ($text -ne "") {
                    addUnitLinesClm $units "${base}[図形]" @($text)
                }
            }
        }
        # 本文に参照の無いコメントは "文書[コメント]" にまとめる（番号順。[int]::TryParse と同じく、読めなければ 0 とみなす）
        foreach ($id in @($comments.Keys | Sort-Object { $n = 0; if ($_ -match '^\s*[+-]?[0-9]+\s*$') { try { $n = [int]$_ } catch { $n = 0 } }; $n })) {
            if (!$usedComments.ContainsKey($id)) {
                addUnitLinesClm $units "文書[コメント]" @($comments[$id])
            }
        }

        # ヘッダー・フッター（重複は除く。ヘッダー → フッターの順）
        $seen = New-Object System.Collections.Hashtable
        $parts = @(getOfficeZipEntries $zip | Where-Object { $_ -match "^word/(header|footer)\d*\.xml$" } |
            Sort-Object { $_ -notmatch "/header" }, { $_ })
        foreach ($part in $parts) {
            foreach ($line in (readXmlLinesClm (readOfficeZipEntry $zip $part) ${nsWordClm})) {
                if (!$seen.ContainsKey($line.Text)) {
                    $seen[$line.Text] = $true
                    addUnitLinesClm $units "ヘッダー・フッター" @($line.Text)
                }
            }
        }

        foreach ($partName in @("word/footnotes.xml", "word/endnotes.xml")) {
            $xml = readOfficeZipEntry $zip $partName
            if ($null -ne $xml) {
                addUnitLinesClm $units "脚注" @(readXmlLinesClm $xml ${nsWordClm} | ForEach-Object { $_.Text })
            }
        }
    } finally {
        closeOfficeZip $zip
    }
    return (getUnitsClm $units)
}

function readPptxUnitsClm {
    # PowerPoint（.pptx / .pptm）の文字を、スライド・ノートごとと、スライドのフッターに分けて返す（office_reader.ps1 の readPptxUnits と同じ）
    param (
        [string]$path,
        [string]$workRoot
    )

    $units = newUnitsClm
    $footers = @()
    $seenFooters = New-Object System.Collections.Hashtable
    $zip = openOfficeZip $path $workRoot
    try {
        $presentationXml = readOfficeZipEntry $zip "ppt/presentation.xml"
        if ($null -eq $presentationXml) {
            throw "PowerPointのプレゼンテーション情報（ppt/presentation.xml）がありません。"
        }
        [xml]$presentation = $presentationXml
        $presentationRels = readRelationshipsClm $zip "ppt/presentation.xml"

        $number = 0
        foreach ($slideId in $presentation.GetElementsByTagName("sldId", ${nsPresentClm})) {
            $number++
            $rel = $presentationRels[(getXmlAttribute $slideId "id" ${nsRelClm})]
            if ($null -eq $rel) { continue }
            $slideXml = readOfficeZipEntry $zip $rel.Target
            if ($null -eq $slideXml) { continue }

            $unitName = "スライド{0:D3}" -f $number
            if ($slideXml -match '^[\s\S]{0,2000}?<p:sld\b[^>]*\sshow="(0|false)"') {
                $unitName += "（非表示）"
            }
            $objects = @{}
            addUnitLinesClm $units $unitName @(readXmlLinesClm $slideXml ${nsDrawingClm} "none" $null $objects | ForEach-Object { $_.Text })
            $slideRels = readRelationshipsClm $zip $rel.Target
            for ($i = 0; $i -lt $objects.Count; $i++) {
                $text = readObjectTextClm $zip $slideRels $objects[$i]
                if ($text -ne "") {
                    addUnitLinesClm $units "${unitName}[図形]" @($text)
                }
            }
            foreach ($slideRel in $slideRels.Values) {
                if ($slideRel.Type -like "*/comments") {
                    $commentsXml = readOfficeZipEntry $zip $slideRel.Target
                    if ($null -ne $commentsXml) {
                        addUnitLinesClm $units "${unitName}[コメント]" @(readSlideCommentsClm $commentsXml)
                    }
                }
            }
            foreach ($line in (readXmlLinesClm $slideXml ${nsDrawingClm} "none" @("ftr"))) {
                if (!$seenFooters.ContainsKey($line.Text)) {
                    $seenFooters[$line.Text] = $true
                    $footers += $line.Text
                }
            }
            foreach ($slideRel in $slideRels.Values) {
                if ($slideRel.Type -like "*/notesSlide") {
                    $notesXml = readOfficeZipEntry $zip $slideRel.Target
                    if ($null -ne $notesXml) {
                        addUnitLinesClm $units ("スライド{0:D3}_ノート" -f $number) @(readXmlLinesClm $notesXml ${nsDrawingClm} | ForEach-Object { $_.Text })
                    }
                }
            }
        }
        if ($footers.Count -gt 0) {
            addUnitLinesClm $units "ヘッダー・フッター" $footers
        }
    } finally {
        closeOfficeZip $zip
    }
    return (getUnitsClm $units)
}

function readXlsxShapeRowsClm {
    # 図形ごとの文字を @{ Row; Column; Text } の配列で返す（office_reader.ps1 の readXlsxShapeRows と同じ）
    param (
        [string]$xml
    )

    [xml]$doc = $xml
    $rows = @()
    foreach ($anchor in @($doc.SelectNodes("//*") | Where-Object {
                $_.NamespaceURI -eq ${nsSheetDrawingClm} -and $_.LocalName -in @("twoCellAnchor", "oneCellAnchor", "absoluteAnchor") })) {
        $inFallback = $false
        for ($node = $anchor.ParentNode; $null -ne $node; $node = $node.ParentNode) {
            if ($node.NamespaceURI -eq ${nsCompatClm} -and $node.LocalName -eq "Fallback") {
                $inFallback = $true
                break
            }
        }
        if ($inFallback) { continue }

        $lines = @(readXmlLinesClm $anchor.OuterXml ${nsDrawingClm} | ForEach-Object { $_.Text })
        if ($lines.Count -eq 0) { continue }
        $row = 1
        $column = 1
        $from = @(getXmlDescendants $anchor "from" ${nsSheetDrawingClm})
        if ($from.Count -gt 0) {
            $row = 1 + [int]@(getXmlDescendants $from[0] "row" ${nsSheetDrawingClm})[0].InnerText
            $column = 1 + [int]@(getXmlDescendants $from[0] "col" ${nsSheetDrawingClm})[0].InnerText
        }
        $rows += @{ Row = $row; Column = $column; Text = (toObjectCellText $lines) }
    }
    return $rows
}

function readXlsxCommentRowsClm {
    # コメント（メモ）とスレッド形式のコメントを、セル番地 → 文字の辞書で返す（office_reader.ps1 の readXlsxCommentRows と同じ）
    param (
        [string]$commentsXml,
        [string[]]$threadedXmls
    )

    $result = @{}
    $order = @()
    foreach ($xml in @($threadedXmls | Where-Object { $_ })) {
        [xml]$doc = $xml
        foreach ($comment in $doc.GetElementsByTagName("threadedComment", ${nsThreadedClm})) {
            $ref = getXmlAttribute $comment "ref"
            $text = @(getXmlDescendants $comment "text" ${nsThreadedClm} | ForEach-Object { $_.InnerText }) -join "`n"
            if (-not $ref -or $text.Trim() -eq "") { continue }
            if ($result.ContainsKey($ref)) {
                $result[$ref] += "`n" + $text
            } else {
                $result[$ref] = $text
                $order += $ref
            }
        }
    }
    if ($commentsXml) {
        [xml]$doc = $commentsXml
        foreach ($comment in $doc.GetElementsByTagName("comment", ${nsSheetClm})) {
            $ref = getXmlAttribute $comment "ref"
            if (-not $ref -or $result.ContainsKey($ref)) { continue }
            $text = @(getXmlDescendants $comment "t" ${nsSheetClm} | Where-Object { $_.ParentNode.LocalName -ne "rPh" } | ForEach-Object { $_.InnerText }) -join ""
            if ($text.Trim() -ne "") {
                $result[$ref] = $text
                $order += $ref
            }
        }
    }
    $ordered = [ordered]@{}
    foreach ($ref in $order) { $ordered[$ref] = $result[$ref] }
    return $ordered
}

function readXlsxObjectUnitsClm {
    # Excel（.xlsx / .xlsm）の表示シートの図形とコメントの文字を、"<シート名>[図形]" "<シート名>[コメント]" に返す
    # （office_reader.ps1 の readXlsxObjectUnits と同じ）
    param (
        [string]$path,
        [string]$workRoot
    )

    $zip = openOfficeZip $path $workRoot
    try {
        return (readXlsxObjectUnitsFromZipClm $zip)
    } finally {
        closeOfficeZip $zip
    }
}

function readXlsxObjectUnitsFromZipClm {
    # 開いてある ZIP から、図形とコメントの文字を読む（セルと一緒に読むとき、ZIP を 2 回展開しないため）
    param (
        $zip
    )

    $units = newUnitsClm
    $workbookXml = readOfficeZipEntry $zip "xl/workbook.xml"
    if ($null -eq $workbookXml) {
        return (getUnitsClm $units)
    }
    [xml]$workbook = $workbookXml
    $workbookRels = readRelationshipsClm $zip "xl/workbook.xml"

    foreach ($sheet in $workbook.GetElementsByTagName("sheet", ${nsSheetClm})) {
        if ((getXmlAttribute $sheet "state") -in @("hidden", "veryHidden")) { continue }
        $rel = $workbookRels[(getXmlAttribute $sheet "id" ${nsRelClm})]
        if ($null -eq $rel -or $rel.Type -notlike "*/worksheet") { continue }
        $sheetName = getXmlAttribute $sheet "name"

        $shapes = @()
        $commentsXml = $null
        $threadedXmls = @()
        foreach ($sheetRel in (readRelationshipsClm $zip $rel.Target).Values) {
            $xml = readOfficeZipEntry $zip $sheetRel.Target
            if ($null -eq $xml) { continue }
            if ($sheetRel.Type -like "*/drawing") {
                $shapes += @(readXlsxShapeRowsClm $xml)
            } elseif ($sheetRel.Type -like "*/comments") {
                $commentsXml = $xml
            } elseif ($sheetRel.Type -like "*/threadedComment") {
                $threadedXmls += $xml
            }
        }

        for ($i = 0; $i -lt $shapes.Count; $i++) { $shapes[$i].Order = $i }
        $lines = @($shapes | Sort-Object { $_.Row }, { $_.Column }, { $_.Order } |
            ForEach-Object { "$(toColumnName $_.Column)$($_.Row)`t$($_.Text)" })
        if ($lines.Count -gt 0) {
            addUnitLinesClm $units "${sheetName}[図形]" $lines
        }

        $comments = readXlsxCommentRowsClm $commentsXml $threadedXmls
        $lines = @($comments.Keys | Sort-Object { (getCellPosition $_)[0] }, { (getCellPosition $_)[1] } |
            ForEach-Object { "$($_.Replace('$', ''))`t$(toObjectCellText @($comments[$_] -split "\r\n|\r|\n"))" })
        if ($lines.Count -gt 0) {
            addUnitLinesClm $units "${sheetName}[コメント]" $lines
        }
    }
    return (getUnitsClm $units)
}

function writeUnitsClm {
    # ユニットごとに "<場所>.tsv" を書き、書いたファイルの数を返す（office_reader.ps1 の writeUnits と同じ中身。BOM 付き UTF-8・CRLF）
    param (
        [System.Collections.IDictionary]$units,
        [string]$outDir
    )

    $count = 0
    foreach ($unitName in @($units.Keys)) {
        $lines = @(foreach ($line in $units[$unitName]) {
                $line = $line.TrimEnd()
                if ($line -ne "") { $line }
            })
        if ($lines.Count -eq 0) { continue }
        $path = Join-Path $outDir (toIndexFileName $unitName)
        Set-Content -LiteralPath (toLongPath $path) -Value $lines -Encoding UTF8
        $count++
    }
    return $count
}

# ----------------------------------------------------------------------------
# Excel のセル（Excel を使わずに xl/worksheets/*.xml から読む）
# ----------------------------------------------------------------------------

function readSharedStringsClm {
    # xl/sharedStrings.xml の文字列を、番号順の配列で返す。
    # ふりがな（rPh）は読まず、書式で分かれた <r> はつなぐ（Excel のセルの文字と同じ）
    param (
        [string]$xml
    )

    if (!$xml) {
        return @()
    }
    # <si> ごとに、<rPh> の中を除いた <t> をつなぐ
    $items = @{}
    foreach ($match in [regex]::Matches($xml, '<si>(.*?)</si>', "Singleline")) {
        $inner = $match.Groups[1].Value
        # ふりがなは読まない（office_reader.ps1 のコメントと同じ扱い）
        $inner = [regex]::Replace($inner, '<rPh\b.*?</rPh>', "", "Singleline")
        $text = ""
        foreach ($t in [regex]::Matches($inner, '<t(?:\s[^>]*)?>(.*?)</t>', "Singleline")) {
            $text += decodeXmlText $t.Groups[1].Value
        }
        $items[$items.Count] = decodeXlsxEscapes $text
    }
    return @(for ($i = 0; $i -lt $items.Count; $i++) { $items[$i] })
}

function decodeXlsxEscapes {
    # Excel が文字列に書く _xHHHH_（CR など、XML にそのまま書けない文字）を元の文字に戻す。
    # 元から _xHHHH_ という文字だったものは _x005F_xHHHH_ と書かれるため、先に守ってから戻す
    param (
        [string]$text
    )

    if ($text.IndexOf("_x") -lt 0) {
        return $text
    }
    $guard = [string][char]0xE0FE
    $text = $text.Replace("_x005F_", $guard)
    foreach ($match in [regex]::Matches($text, '_x([0-9A-Fa-f]{4})_')) {
        $code = [int]("0x" + $match.Groups[1].Value)
        $text = $text.Replace($match.Value, [string][char]$code)
    }
    return $text.Replace($guard, "_x005F_")
}

function readCellStylesClm {
    # xl/styles.xml から、セルの書式（s 属性）→ 表示形式の書式 の配列を返す
    param (
        [string]$xml
    )

    if (!$xml) {
        return @()
    }
    # ブックに書かれた表示形式（番号 → 書式）
    $formats = @{}
    foreach ($match in [regex]::Matches($xml, '<numFmt\b[^>]*/?>')) {
        $tag = $match.Value
        if ($tag -match 'numFmtId="([^"]*)"' ) {
            $id = $Matches[1]
            if ($tag -match 'formatCode="([^"]*)"') {
                $formats[$id] = decodeXmlText $Matches[1]
            }
        }
    }
    # セルの書式（cellXfs）の並び順が s 属性の番号になる
    $styles = @{}
    if ($xml -match '<cellXfs\b.*?</cellXfs>') {
        foreach ($match in [regex]::Matches($Matches[0], '<xf\b[^>]*>')) {
            $id = "0"
            if ($match.Value -match 'numFmtId="([^"]*)"') {
                $id = $Matches[1]
            }
            $styles[$styles.Count] = getNumberFormatCode $id $formats
        }
    }
    return @(for ($i = 0; $i -lt $styles.Count; $i++) { $styles[$i] })
}

function toExcelCellText {
    # セルの表示文字を、Excel のテキスト保存と同じ 1 セルにする。
    # 改行はセル内改行（$cellNewLine）にし、改行・" ・タブ・カンマを含むときは " で囲む（中の " は "" にする）
    param (
        [string]$text
    )

    if ($text -eq "") {
        return ""
    }
    if ($text.IndexOfAny([char[]]@('"', "`t", "`r", "`n", ",")) -lt 0) {
        return $text
    }
    $text = ($text -replace "\r\n|\r|\n", ${cellNewLine}).Replace('"', '""')
    return "`"$text`""
}

function readXlsxSheetTextClm {
    # 1 シートの XML を、Excel のテキスト保存と同じ TSV（整形前）にする。
    # 1 行目がシートの 1 行目、1 列目が A 列になるよう、間の空行・空セルはタブと改行で埋める
    param (
        [string]$xml,
        [object[]]$sharedStrings = @(),
        [object[]]$styles = @(),
        [bool]$date1904 = $false
    )

    if (!$xml) {
        return ""
    }
    # 行ごと・セルごとに走査する（シートが大きいと [xml] の組み立てが重いため、正規表現で読む）
    $rows = @{}
    $maxRow = 0
    foreach ($rowMatch in [regex]::Matches($xml, '<row\b[^>]*?(?:/>|>(.*?)</row>)', "Singleline")) {
        $rowTag = $rowMatch.Value
        $rowNumber = 0
        if ($rowTag -match '^<row\b[^>]*\br="([0-9]{1,9})"') {
            $rowNumber = [int]$Matches[1]
        }
        $cells = @{}
        $maxColumn = 0
        $nextColumn = 1  # r（セル番地）の無いファイルのために、左から順に数える
        foreach ($cellMatch in [regex]::Matches($rowMatch.Groups[1].Value, '<c\b[^>]*?(?:/>|>(.*?)</c>)', "Singleline")) {
            $cellTag = $cellMatch.Value
            $column = 0
            if ($cellTag -match '^<c\b[^>]*\br="([A-Za-z]{1,3})([0-9]{1,9})"') {
                $position = getCellPosition ($Matches[1] + $Matches[2])
                $column = $position[1]
                if ($rowNumber -eq 0) { $rowNumber = $position[0] }
            } else {
                $column = $nextColumn  # r が無いファイル（書き出し側によっては省く）は左から順に並ぶ
            }
            $nextColumn = $column + 1
            $kind = "number"
            if ($cellTag -match '^<c\b[^>]*\bt="([^"]*)"') {
                $kind = $Matches[1]
            }
            $style = ""
            if ($cellTag -match '^<c\b[^>]*\bs="([0-9]{1,9})"') {
                $style = $Matches[1]
            }
            $inner = $cellMatch.Groups[1].Value
            $text = getXlsxCellDisplayText $inner $kind $style $sharedStrings $styles $date1904
            if ($text -ne "") {
                $cells[$column] = toExcelCellText $text
                if ($column -gt $maxColumn) { $maxColumn = $column }
            }
        }
        if ($rowNumber -gt 0 -and $maxColumn -gt 0) {
            $line = @(for ($i = 1; $i -le $maxColumn; $i++) { $(if ($cells.ContainsKey($i)) { $cells[$i] } else { "" }) }) -join "`t"
            $rows[$rowNumber] = $line
            if ($rowNumber -gt $maxRow) { $maxRow = $rowNumber }
        }
    }
    if ($maxRow -eq 0) {
        return ""
    }
    return (@(for ($i = 1; $i -le $maxRow; $i++) { $(if ($rows.ContainsKey($i)) { $rows[$i] } else { "" }) }) -join "`r`n")
}

function getXlsxCellDisplayText {
    # 1 セルの中身（<c> の中）から、Excel が画面に出す文字を作る
    param (
        [string]$inner,
        [string]$kind,
        [string]$style,
        [object[]]$sharedStrings = @(),
        [object[]]$styles = @(),
        [bool]$date1904 = $false
    )

    if ($kind -eq "inlineStr") {
        $text = ""
        $body = [regex]::Replace($inner, '<rPh\b.*?</rPh>', "", "Singleline")
        foreach ($t in [regex]::Matches($body, '<t(?:\s[^>]*)?>(.*?)</t>', "Singleline")) {
            $text += decodeXmlText $t.Groups[1].Value
        }
        return $text
    }
    $value = ""
    # 改行を含む値（数式の結果など）もあるため、改行をまたいで探す
    $found = [regex]::Match($inner, '<v(?:\s[^>]*)?>(.*?)</v>', "Singleline")
    if ($found.Success) {
        $value = decodeXmlText $found.Groups[1].Value
    }
    if ($value -eq "") {
        return ""
    }
    if ($kind -eq "s") {
        # 共有文字列（番号で引く）
        if ($value -match '^[0-9]{1,9}$' -and [int]$value -lt $sharedStrings.Count) {
            return [string]$sharedStrings[[int]$value]
        }
        return ""
    }
    if ($kind -eq "str") {
        return $value  # 数式の結果の文字列
    }
    $format = "General"
    if ($style -ne "" -and [int]$style -lt $styles.Count) {
        $format = [string]$styles[[int]$style]
    }
    if ($kind -eq "e") {
        return (formatExcelCellText $value "error" $format $date1904)
    }
    if ($kind -eq "b") {
        return (formatExcelCellText $value "boolean" $format $date1904)
    }
    return (formatExcelCellText $value "number" $format $date1904)
}

function readXlsxCellTextsClm {
    # Excel（.xlsx / .xlsm）の表示シートのセルを、シート名 → TSV（整形前）で返す。
    # 非表示シート・グラフシート・値の無いシートは返さない（いつものインデクサと同じ）
    param (
        [string]$path,
        [string]$workRoot
    )

    $zip = openOfficeZip $path $workRoot
    try {
        return (readXlsxCellTextsFromZipClm $zip)
    } finally {
        closeOfficeZip $zip
    }
}

function readXlsxCellTextsFromZipClm {
    # 開いてある ZIP から、表示シートのセルを シート名 → TSV（整形前）で読む
    param (
        $zip
    )

    $texts = [ordered]@{}
    $workbookXml = readOfficeZipEntry $zip "xl/workbook.xml"
    if ($null -eq $workbookXml) {
        return $texts
    }
    [xml]$workbook = $workbookXml
    $date1904 = $false
    foreach ($pr in $workbook.GetElementsByTagName("workbookPr", ${nsSheetClm})) {
        $flag = getXmlAttribute $pr "date1904"
        if ($flag -eq "1" -or $flag -eq "true") { $date1904 = $true }
    }
    $sharedStrings = readSharedStringsClm (readOfficeZipEntry $zip "xl/sharedStrings.xml")
    $styles = readCellStylesClm (readOfficeZipEntry $zip "xl/styles.xml")
    $workbookRels = readRelationshipsClm $zip "xl/workbook.xml"

    foreach ($sheet in $workbook.GetElementsByTagName("sheet", ${nsSheetClm})) {
        if ((getXmlAttribute $sheet "state") -in @("hidden", "veryHidden")) { continue }
        $rel = $workbookRels[(getXmlAttribute $sheet "id" ${nsRelClm})]
        if ($null -eq $rel -or $rel.Type -notlike "*/worksheet") { continue }
        $sheetXml = readOfficeZipEntry $zip $rel.Target
        if ($null -eq $sheetXml) { continue }
        $text = readXlsxSheetTextClm $sheetXml $sharedStrings $styles $date1904
        if ($text -ne "") {
            # シート名のタブ・改行は _x0009_ のように書かれるため、元の文字に戻す
            $texts[(decodeXlsxEscapes (getXmlAttribute $sheet "name"))] = $text
        }
    }
    return $texts
}

function writeXlsxTsvClm {
    # Excel（.xlsx / .xlsm）のセル・図形・コメントを、いつものインデクサと同じ形の TSV にして書き、書いた数を返す。
    # ZIP は 1 回だけ開く（tar.exe での展開は時間がかかるため）
    param (
        [string]$path,
        [string]$workRoot,
        [string]$outDir
    )

    $zip = openOfficeZip $path $workRoot
    try {
        $units = readXlsxObjectUnitsFromZipClm $zip
        $texts = readXlsxCellTextsFromZipClm $zip
    } finally {
        closeOfficeZip $zip
    }

    $count = writeUnitsClm $units $outDir
    foreach ($sheetName in @($texts.Keys)) {
        # 行末の空セル・末尾の空行を取り除き、1 行目がシートの 1 行目になるようにそろえる（いつものインデクサと同じ formatTsv）
        $content = formatTsv $texts[$sheetName] 1 1
        if ($content -eq "") {
            continue
        }
        writeUtf8BomLines (toLongPath (Join-Path $outDir (toIndexFileName $sheetName))) ($content -split "`r`n")
        $count++
    }
    return $count
}
