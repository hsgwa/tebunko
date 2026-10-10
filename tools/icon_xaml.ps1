# アイコンの元データ（SVG）を、画面が読む XAML（DrawingImage）に変える。tools\new_icon.ps1 が読み込んで使う。
#
# 受け付ける SVG の形（外れたら理由を出して止まる。黙って落とさない）:
#   - ルートは svg で、viewBox は "0 0 <幅> <高さ>"
#   - 中身は path だけ（コメントは読み飛ばす）。path が持てる属性は d と fill だけ
#   - d は M・L・H・V・C・Q・Z（大文字と小文字）と数字だけ。fill は #RRGGBB
#   - transform・stroke・グラデーション・クリップ・円などの図形は受け付けない
# 受け付けるものを増やすときは、変換とテスト（tests\tools\icon_xaml.Tests.ps1）を一緒に直す。

function convertSvgToIconXaml {
    param (
        [string]$SvgText
    )

    $xml = New-Object System.Xml.XmlDocument
    try {
        $xml.LoadXml($SvgText)
    } catch {
        throw "SVG を読めません: $($_.Exception.Message)"
    }
    $root = $xml.DocumentElement
    $svgNamespace = "http://www.w3.org/2000/svg"
    if ($root.LocalName -ne "svg" -or $root.NamespaceURI -ne $svgNamespace) {
        throw "ルートが svg（$svgNamespace）ではありません。"
    }
    if ($root.GetAttribute("viewBox") -notmatch '^0 0 ([0-9]+) ([0-9]+)$') {
        throw "viewBox は `"0 0 <幅> <高さ>`" の形だけ使えます（今は `"$($root.GetAttribute('viewBox'))`"）。"
    }
    $width = $Matches[1]
    $height = $Matches[2]

    # 最初に、viewBox の枠いっぱいの透明な四角を置く。絵の大きさ（Bounds）が枠になり、余白も含めて拡大・縮小される
    $drawings = New-Object System.Collections.Generic.List[string]
    $drawings.Add("      <GeometryDrawing Brush=`"Transparent`" Geometry=`"M0,0 L$width,0 L$width,$height L0,$height Z`" />")
    foreach ($node in $root.ChildNodes) {
        if ($node.NodeType -eq [System.Xml.XmlNodeType]::Comment -or $node.NodeType -eq [System.Xml.XmlNodeType]::Whitespace -or $node.NodeType -eq [System.Xml.XmlNodeType]::SignificantWhitespace) {
            continue
        }
        if ($node.NodeType -ne [System.Xml.XmlNodeType]::Element -or $node.LocalName -ne "path" -or $node.NamespaceURI -ne $svgNamespace) {
            throw "svg の中は path だけ使えます（見つかったもの: $($node.Name)）。"
        }
        foreach ($attribute in $node.Attributes) {
            if ($attribute.Name -ne "d" -and $attribute.Name -ne "fill") {
                throw "path の属性は d と fill だけ使えます（見つかったもの: $($attribute.Name)）。"
            }
        }
        $d = $node.GetAttribute("d")
        $fill = $node.GetAttribute("fill")
        if ($d -notmatch '^[MmLlHhVvCcQqZz0-9.,\s-]+$') {
            throw "path の d に使えない文字があります（M L H V C Q Z と数字だけ使えます）: $d"
        }
        if ($fill -notmatch '^#[0-9A-Fa-f]{6}$') {
            throw "path の fill は #RRGGBB の形だけ使えます（今は `"$fill`"）。"
        }
        $drawings.Add("      <GeometryDrawing Brush=`"$($fill.ToUpper())`" Geometry=`"$d`" />")
    }
    if ($drawings.Count -eq 1) {
        throw "path がありません。"
    }

    $lines = @(
        "<!-- 画面のアイコン（窓・スプラッシュ・バージョン情報で使うベクターの絵）。",
        "     docs\images\logo.svg から tools\new_icon.ps1 が作る。手で編集しない（図柄を変えるときは SVG を直して作り直す）。 -->",
        "<DrawingImage xmlns=`"http://schemas.microsoft.com/winfx/2006/xaml/presentation`">",
        "  <DrawingImage.Drawing>",
        "    <DrawingGroup>"
    ) + $drawings + @(
        "    </DrawingGroup>",
        "  </DrawingImage.Drawing>",
        "</DrawingImage>"
    )
    return ($lines -join "`r`n") + "`r`n"
}
