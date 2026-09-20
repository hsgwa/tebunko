# モダンなデザインのサンプルを表示する（見本専用。本体の動作には関係しない）。
#   表示       : powershell -NoProfile -STA -File scripts\theme_sample.ps1
#   PNG 書き出し: powershell -NoProfile -STA -File scripts\theme_sample.ps1 -Shot <出力フォルダ>
param(
  [string]$Shot
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# XAML を読み込んで Window を返す
function loadWindow([string]$path) {
  $reader = New-Object System.Xml.XmlTextReader($path)
  try { [System.Windows.Markup.XamlReader]::Load($reader) } finally { $reader.Close() }
}

# 表の中身（見本用のダミー）
function indexRows {
  @(
    [pscustomobject]@{ Enabled = $true;  Name = '見積書';     Path = 'D:\共有\見積書';       StatusText = '最新です';         FileCountText = '184'; LastConvertedText = '09-19 18:42' }
    [pscustomobject]@{ Enabled = $true;  Name = '議事録';     Path = 'D:\共有\会議\議事録'; StatusText = '9 件が未変換です'; FileCountText = '203'; LastConvertedText = '09-12 09:05' }
    [pscustomobject]@{ Enabled = $false; Name = '提案資料';   Path = '\\fs01\提案\2026';     StatusText = '見つかりません';   FileCountText = '25';  LastConvertedText = '08-30 14:20' }
  )
}
function resultRows {
  @(
    [pscustomobject]@{ RelDir = '2026\04'; Book = '04_見積_東日本支社.xlsx'; Location = '見積明細';   LineNumber = 42;  CellText = 'F12'; Line = '消費税（10%）を含む合計金額' }
    [pscustomobject]@{ RelDir = '2026\04'; Book = '04_見積_東日本支社.xlsx'; Location = '条件';       LineNumber = 8;   CellText = 'B8';  Line = '価格は消費税抜きの表示です' }
    [pscustomobject]@{ RelDir = '2026\03'; Book = '03_見積_中部支社.xlsx';   Location = '見積明細';   LineNumber = 41;  CellText = 'F12'; Line = '消費税額' }
    [pscustomobject]@{ RelDir = '2026\03'; Book = '議事録_価格改定.docx';     Location = '2 ページ';   LineNumber = 117; CellText = '';    Line = '消費税の扱いを次回までに確認する' }
    [pscustomobject]@{ RelDir = '2025\12'; Book = '12_見積_九州支社.xlsx';   Location = '表紙';       LineNumber = 5;   CellText = 'C5';  Line = '税込（消費税 10%）' }
    [pscustomobject]@{ RelDir = '2025\12'; Book = '価格改定のご案内.pptx';    Location = 'スライド 3'; LineNumber = 3;   CellText = '';    Line = '消費税率の変更に伴う改定について' }
    [pscustomobject]@{ RelDir = '2025\11'; Book = '11_見積_北海道支社.xlsx'; Location = '見積明細';   LineNumber = 42;  CellText = 'F12'; Line = '消費税（10%）' }
  )
}

$window = loadWindow (Join-Path $PSScriptRoot 'theme_sample.xaml')
$window.FindName('IndexGrid').ItemsSource  = indexRows
$window.FindName('ResultGrid').ItemsSource = resultRows

if (-not $Shot) {
  $window.ShowDialog() | Out-Null
  return
}

# --- PNG 書き出し（新デザインと現行デザインを並べて比べるため） ---
New-Item -ItemType Directory -Force -Path $Shot | Out-Null

# 画面の描画が終わるまで待つ
function waitIdle($w) {
  $w.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::ApplicationIdle) | Out-Null
  $w.UpdateLayout()
  $w.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::ApplicationIdle) | Out-Null
}

# Window の中身を PNG にする
function savePng($w, [string]$file) {
  waitIdle $w
  $width  = [int][Math]::Ceiling($w.ActualWidth)
  $height = [int][Math]::Ceiling($w.ActualHeight)
  $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($width, $height, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
  $rtb.Render($w)
  $encoder = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
  $encoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb)) | Out-Null
  $stream = [System.IO.File]::Create($file)
  try { $encoder.Save($stream) } finally { $stream.Dispose() }
  Write-Host "saved: $file"
}

# 新デザイン
$window.Show()
$tabs = $window.FindName('Tabs')
$tabs.SelectedIndex = 1
savePng $window (Join-Path $Shot 'after_検索.png')
$tabs.SelectedIndex = 0
savePng $window (Join-Path $Shot 'after_インデックス管理.png')
$window.Close()

# 現行デザイン（config_gui.xaml をそのまま読み込む。イベントは付けないので見た目だけ）
$current = loadWindow (Join-Path $PSScriptRoot 'config_gui.xaml')
$current.Width = 1000
$current.Height = 720
$current.Show()
$currentTabs = $current.FindName('Tabs')
$currentTabs.SelectedIndex = 1
savePng $current (Join-Path $Shot 'before_検索.png')
$currentTabs.SelectedIndex = 0
savePng $current (Join-Path $Shot 'before_インデックス管理.png')
$current.Close()
