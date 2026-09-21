# GitHub の social preview 用の画像（1280×640 の PNG）を作る。
#
#   .\tools\make_social_preview.ps1              docs\images\social_preview.png に書き出す
#   .\tools\make_social_preview.ps1 -Out .\x.png
#
# 書き出した画像は、リポジトリの Settings → Social preview から手で登録する（API は無い）。
# 右側の検索画面は見本で、社名・人名は架空のもの。
param([string]$Out = (Join-Path $PSScriptRoot '..\docs\images\social_preview.png'))
Add-Type -AssemblyName System.Drawing
$Out = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Out)

$W = 1280; $H = 640
$bmp = New-Object System.Drawing.Bitmap $W, $H
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = 'AntiAlias'
$g.TextRenderingHint = 'AntiAliasGridFit'
$g.InterpolationMode = 'HighQualityBicubic'

function C([string]$hex, [int]$a = 255) {
    $c = [System.Drawing.ColorTranslator]::FromHtml($hex)
    [System.Drawing.Color]::FromArgb($a, $c)
}
function Brush([string]$hex, [int]$a = 255) { New-Object System.Drawing.SolidBrush (C $hex $a) }
function Font([float]$size, [string]$style = 'Regular', [string]$family = 'Yu Gothic UI') {
    New-Object System.Drawing.Font $family, $size, ([System.Drawing.FontStyle]$style), ([System.Drawing.GraphicsUnit]::Pixel)
}
function RoundRect([float]$x, [float]$y, [float]$w, [float]$h, [float]$r) {
    $p = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $r * 2
    $p.AddArc($x, $y, $d, $d, 180, 90)
    $p.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
    $p.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90)
    $p.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
    $p.CloseFigure()
    $p
}
function Text([string]$s, $font, $brush, [float]$x, [float]$y) {
    $g.DrawString($s, $font, $brush, $x, $y, [System.Drawing.StringFormat]::GenericTypographic)
}
function TextWidth([string]$s, $font) {
    $g.MeasureString($s, $font, 10000, [System.Drawing.StringFormat]::GenericTypographic).Width
}

# 背景
$g.Clear((C '#F6F7F9'))
$g.FillRectangle((Brush '#2563EB'), 0, 0, $W, 10)

# 左: 見出し
$x0 = 80
Text 'Windows 用 Office 全文検索ツール' (Font 26 'Bold') (Brush '#2563EB') $x0 118
$fTitle = Font 104 'Bold' 'Segoe UI'
Text 'tebunko' $fTitle (Brush '#1F2937') ($x0 - 4) 158
$tw = TextWidth 'tebunko' $fTitle
Text '手文庫' (Font 38) (Brush '#78828F') ($x0 + $tw + 14) 216

$fTag = Font 38 'Bold'
Text 'Excel・Word・PowerPoint を' $fTag (Brush '#414B5A') $x0 312
Text '中身の文字で横断検索' $fTag (Brush '#414B5A') $x0 362

# 左: 特長のチップ（1 段目は速さをアクセント色で、2 段目は手軽さ）
$fChip = Font 23 'Bold'
$chips = @(
    @{ Text = '数千ファイルを数秒で検索'; Row = 0; Strong = $true },
    @{ Text = 'インストール不要';         Row = 1; Strong = $false },
    @{ Text = '管理者権限不要';           Row = 1; Strong = $false }
)
$cx = @{ 0 = $x0; 1 = $x0 }
foreach ($c in $chips) {
    $y = 440 + $c.Row * 64
    $w = (TextWidth $c.Text $fChip) + 40
    $p = RoundRect $cx[$c.Row] $y $w 50 25
    if ($c.Strong) {
        $g.FillPath((Brush '#2563EB'), $p)
        Text $c.Text $fChip (Brush '#FFFFFF') ($cx[$c.Row] + 20) ($y + 11)
    } else {
        $g.FillPath((Brush '#EFF4FF'), $p)
        $g.DrawPath((New-Object System.Drawing.Pen (C '#BFD3FE'), 2), $p)
        Text $c.Text $fChip (Brush '#1D4ED8') ($cx[$c.Row] + 20) ($y + 11)
    }
    $cx[$c.Row] += $w + 14
}

Text 'github.com/hsgwa/tebunko' (Font 22 'Regular' 'Segoe UI') (Brush '#A3ABB5') $x0 588

# 右: 検索画面の見本カード
$kx = 720; $ky = 92; $kw = 480; $kh = 456
for ($i = 1; $i -le 8; $i++) {
    $sp = RoundRect ($kx - $i + 4) ($ky - $i + 10) ($kw + 2 * $i) ($kh + 2 * $i) (16 + $i)
    $g.FillPath((Brush '#1F2937' 6), $sp)
}
$card = RoundRect $kx $ky $kw $kh 16
$g.FillPath((Brush '#FFFFFF'), $card)
$g.DrawPath((New-Object System.Drawing.Pen (C '#E9ECF0'), 1.5), $card)

# 検索欄
$sb = RoundRect ($kx + 24) ($ky + 24) ($kw - 48) 52 10
$g.FillPath((Brush '#FFFFFF'), $sb)
$g.DrawPath((New-Object System.Drawing.Pen (C '#2563EB'), 2.5), $sb)
$lens = New-Object System.Drawing.Pen (C '#2563EB'), 3
$g.DrawEllipse($lens, $kx + 44, $ky + 39, 18, 18)
$g.DrawLine($lens, $kx + 60, $ky + 55, $kx + 68, $ky + 63)
Text '見積書' (Font 24 'Bold') (Brush '#1F2937') ($kx + 84) ($ky + 35)
Text '1,284 件' (Font 20) (Brush '#78828F') ($kx + $kw - 124) ($ky + 38)

# 結果の行
$rows = @(
    @{ Tag = 'X'; Color = '#1D6F42'; Name = '2024年度_見積一覧.xlsx'; Place = 'シート: 4月'; Pre = '(株)山田商事 '; Hit = '見積書'; Post = ' No.1024' },
    @{ Tag = 'W'; Color = '#2B579A'; Name = '提案書_営業部.docx';     Place = '3 ページ';    Pre = '別紙の';       Hit = '見積書'; Post = 'をご参照ください' },
    @{ Tag = 'P'; Color = '#C43E1C'; Name = '定例会議資料.pptx';      Place = 'スライド 7';  Pre = '';             Hit = '見積書'; Post = 'の提出期限を確認' },
    @{ Tag = 'X'; Color = '#1D6F42'; Name = '発注管理.xlsm';          Place = 'シート: 集計'; Pre = '佐藤様 ';      Hit = '見積書'; Post = ' 受領済み' }
)
$fName = Font 20 'Bold'; $fPlace = Font 17; $fSnip = Font 19; $fSnipB = Font 19 'Bold'; $fIcon = Font 20 'Bold' 'Segoe UI'
$ry = $ky + 100
foreach ($r in $rows) {
    $g.DrawLine((New-Object System.Drawing.Pen (C '#E9ECF0'), 1.5), $kx + 24, $ry - 8, $kx + $kw - 24, $ry - 8)
    $ic = RoundRect ($kx + 24) ($ry + 6) 36 36 7
    $g.FillPath((Brush $r.Color), $ic)
    $iw = TextWidth $r.Tag $fIcon
    Text $r.Tag $fIcon (Brush '#FFFFFF') ($kx + 42 - $iw / 2) ($ry + 12)
    Text $r.Name $fName (Brush '#1F2937') ($kx + 76) ($ry + 2)
    $pw = TextWidth $r.Place $fPlace
    Text $r.Place $fPlace (Brush '#78828F') ($kx + $kw - 24 - $pw) ($ry + 5)
    $sx = $kx + 76; $sy = $ry + 36
    if ($r.Pre) { Text $r.Pre $fSnip (Brush '#414B5A') $sx $sy; $sx += TextWidth $r.Pre $fSnip }
    $hw = TextWidth $r.Hit $fSnipB
    $g.FillRectangle((Brush '#FDE68A'), $sx - 2, $sy - 1, $hw + 4, 26)
    Text $r.Hit $fSnipB (Brush '#1F2937') $sx $sy; $sx += $hw
    Text $r.Post $fSnip (Brush '#414B5A') $sx $sy
    $ry += 86
}

$g.Dispose()
$bmp.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
