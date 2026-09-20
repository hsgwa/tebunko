# win_grep の動作確認用テストデータ（Excel・Word・PowerPoint のファイル群と設定ファイル例）を生成する。
# Excel・Word・PowerPoint（COM）が必要。出力フォルダは削除して作り直す。
# PowerPoint は起動中のインスタンスに接続してしまうため、PowerPoint を終了してから実行すること。
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\testdata\make_testdata.ps1
# ケースの一覧と期待結果は同じフォルダの README.md を参照。
param (
    [string]$OutDir = "$PSScriptRoot\excel",
    [string]$ConfigDir = "$PSScriptRoot\設定例",
    [int]$BigRows = 20000,
    [int]$ManyFiles = 200,
    [switch]$WithRisky  # フォルダのループ（ジャンクション）と 260 文字を超えるパスも作る
)

$ErrorActionPreference = "Stop"

# Excel は [ ] を含むパスや長いパスに保存できないため、TEMP に保存してから移動する
$stageDir = Join-Path ([System.IO.Path]::GetTempPath()) "win_grep_testdata"
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$missing = [Type]::Missing

# Excel の定数
$xlFormats = @{
    51 = ".xlsx"  # xlOpenXMLWorkbook
    52 = ".xlsm"  # xlOpenXMLWorkbookMacroEnabled
    56 = ".xls"   # xlExcel8
    50 = ".xlsb"  # xlExcel12
    54 = ".xltx"  # xlOpenXMLTemplate
    6  = ".csv"   # xlCSV
}
$xlSheetHidden = 0
$xlSheetVeryHidden = 2

# Word の定数
$wdFormats = @{
    12 = ".docx"  # wdFormatXMLDocument
    13 = ".docm"  # wdFormatXMLDocumentMacroEnabled
    0  = ".doc"   # wdFormatDocument97
    6  = ".rtf"   # wdFormatRTF
    14 = ".dotx"  # wdFormatXMLTemplate
    23 = ".odt"   # wdFormatOpenDocumentText
    17 = ".pdf"   # wdFormatPDF
}
$wdStyleNormal = -1
$wdStyleHeading1 = -2
$wdStyleHeading2 = -3

# PowerPoint の定数
$ppFormats = @{
    24 = ".pptx"  # ppSaveAsOpenXMLPresentation
    25 = ".pptm"  # ppSaveAsOpenXMLPresentationMacroEnabled
    1  = ".ppt"   # ppSaveAsPresentation
    28 = ".ppsx"  # ppSaveAsOpenXMLShow
    7  = ".pps"   # ppSaveAsShow
    26 = ".potx"  # ppSaveAsOpenXMLTemplate
    35 = ".odp"   # ppSaveAsOpenDocumentPresentation
    32 = ".pdf"   # ppSaveAsPDF
}
$ppLayoutTitle = 1
$ppLayoutText = 2
$ppLayoutTitleOnly = 11

# CP932 で表せない文字
$yoshi = [char]::ConvertFromUtf32(0x20BB7)  # 𠮷（つちよし）
$sushi = [char]::ConvertFromUtf32(0x1F363)  # 🍣

$failed = New-Object System.Collections.Generic.List[string]
$noBom = New-Object System.Text.UTF8Encoding($false)

$excel = $null
$word = $null
$ppt = $null

function newBook {
    # 指定した名前のシートを順に持つブックを作る
    param (
        [string[]]$sheetNames
    )

    $wb = $excel.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) {
        $wb.Worksheets.Item($wb.Worksheets.Count).Delete()
    }
    $wb.Worksheets.Item(1).Name = $sheetNames[0]
    for ($i = 1; $i -lt $sheetNames.Count; $i++) {
        $ws = $wb.Worksheets.Add($missing, $wb.Worksheets.Item($wb.Worksheets.Count))
        $ws.Name = $sheetNames[$i]
    }
    $wb.Worksheets.Item(1).Activate()

    return $wb
}

function setRows {
    # 行ごとの配列をまとめてセルに書き込む（"=" 始まりは数式になる）
    param (
        $ws,
        [string]$topLeft,
        [object[]]$rows
    )

    $colCount = ($rows | ForEach-Object { @($_).Count } | Measure-Object -Maximum).Maximum
    $data = New-Object 'object[,]' $rows.Count, $colCount
    for ($r = 0; $r -lt $rows.Count; $r++) {
        $row = @($rows[$r])
        for ($c = 0; $c -lt $row.Count; $c++) {
            $data[$r, $c] = $row[$c]
        }
    }
    $ws.Range($topLeft).Resize($rows.Count, $colCount).Value2 = $data
}

function setCell {
    # 1 セルに値を書き込む。書式を指定すると値より先に設定する
    param (
        $ws,
        [string]$address,
        $value,
        [string]$format
    )

    $cell = $ws.Range($address)
    if ($format) {
        $cell.NumberFormat = $format
    }
    $cell.Value2 = $value
}

function saveBook {
    # ブックを保存して閉じ、出力フォルダの relPath に置く
    param (
        $wb,
        [string]$relPath,
        [int]$format = 51,
        [string]$password,
        [string]$writePassword
    )

    $stage = Join-Path $stageDir ([guid]::NewGuid().ToString("N") + $xlFormats[$format])
    $pw = if ($password) { $password } else { $missing }
    $wpw = if ($writePassword) { $writePassword } else { $missing }
    $wb.SaveAs($stage, $format, $pw, $wpw)
    $wb.Close($false)

    moveToOut $stage $relPath
}

function moveToOut {
    param (
        [string]$src,
        [string]$relPath
    )

    $dest = Join-Path $OutDir $relPath
    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($dest)) | Out-Null
    if ([System.IO.File]::Exists($dest)) {
        [System.IO.File]::Delete($dest)
    }
    [System.IO.File]::Move($src, $dest)
    Write-Host "  $relPath"
}

function writeText {
    # テキストファイルを書く（出力フォルダからの相対パス、または絶対パス）
    param (
        [string]$path,
        [string]$text,
        [System.Text.Encoding]$encoding = $utf8Bom
    )

    if (![System.IO.Path]::IsPathRooted($path)) {
        $path = Join-Path $OutDir $path
        Write-Host "  $($path.Substring($OutDir.Length + 1))"
    }
    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
    [System.IO.File]::WriteAllText($path, $text, $encoding)
}

function simpleBook {
    # 1 シートに数行だけ書いたブックを保存する
    param (
        [string]$relPath,
        [string]$marker,
        [int]$format = 51,
        [string]$sheetName = "Sheet1"
    )

    $wb = newBook @($sheetName)
    setRows $wb.Worksheets.Item(1) "A1" @(
        @("ID", "内容"),
        @($marker, "$([System.IO.Path]::GetFileName($relPath)) のデータ")
    )
    saveBook $wb $relPath $format
}

function writeLockFile {
    # Office がファイルを開いているときに作るロックファイル（~$〜）に似せたもの
    param (
        [string]$relPath,
        [bool]$hidden = $true
    )

    $lock = New-Object byte[] 165
    $owner = [System.Text.Encoding]::ASCII.GetBytes("tester")
    $lock[0] = $owner.Length
    [Array]::Copy($owner, 0, $lock, 1, $owner.Length)
    $path = Join-Path $OutDir $relPath
    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
    [System.IO.File]::WriteAllBytes($path, $lock)
    if ($hidden) {
        [System.IO.File]::SetAttributes($path, "Hidden")
    }
    Write-Host "  $relPath$(if ($hidden) { '（隠し属性）' })"
}

function addParagraph {
    # 文書の末尾に段落を追加し、追加した文字列の範囲を返す
    param (
        $doc,
        [string]$text,
        [int]$style = $wdStyleNormal
    )

    $range = $doc.Content
    $range.Collapse(0)  # wdCollapseEnd
    $range.Text = $text
    $range.Style = $style
    $start = $range.Start
    $end = $range.End
    $range.InsertParagraphAfter()

    return , $doc.Range($start, $end)
}

function saveDoc {
    # 文書を保存して閉じ、出力フォルダの relPath に置く
    param (
        $doc,
        [string]$relPath,
        [int]$format = 12,
        [string]$password,
        [string]$writePassword
    )

    # Join-Path の戻り値（PSObject）のまま Word・PowerPoint に渡すと、保存が終わらなくなるため文字列にする
    $stage = [string](Join-Path $stageDir ([guid]::NewGuid().ToString("N") + $wdFormats[$format]))
    if ($password) {
        $doc.Password = $password
    }
    if ($writePassword) {
        $doc.WritePassword = $writePassword
    }
    # ページ割りを確定させてから保存する（しないと長い文書でもページ区切りの情報が保存されない）
    $doc.Repaginate()
    $doc.SaveAs2($stage, $format)
    $doc.Close(0)  # wdDoNotSaveChanges

    moveToOut $stage $relPath
}

function simpleDoc {
    # 数段落だけの文書を保存する
    param (
        [string]$relPath,
        [string]$marker,
        [int]$format = 12,
        [string]$password
    )

    $doc = $word.Documents.Add()
    [void](addParagraph $doc $marker)
    [void](addParagraph $doc "$([System.IO.Path]::GetFileName($relPath)) のデータ")
    saveDoc $doc $relPath $format $password
}

function addSlide {
    # 末尾にスライドを追加する。title を指定するとタイトルに設定する
    param (
        $pres,
        [int]$layout,
        [string]$title
    )

    $slide = $pres.Slides.Add($pres.Slides.Count + 1, $layout)
    if ($title) {
        $slide.Shapes.Title.TextFrame.TextRange.Text = $title
    }

    return , $slide
}

function addTextBox {
    param (
        $slide,
        [string]$text,
        [float]$left,
        [float]$top,
        [float]$width = 400,
        [float]$height = 40
    )

    $box = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)  # msoTextOrientationHorizontal
    $box.TextFrame.TextRange.Text = $text

    return , $box
}

function savePres {
    # プレゼンテーションを保存して閉じ、出力フォルダの relPath に置く
    param (
        $pres,
        [string]$relPath,
        [int]$format = 24,
        [string]$password
    )

    # Join-Path の戻り値（PSObject）のまま Word・PowerPoint に渡すと、保存が終わらなくなるため文字列にする
    $stage = [string](Join-Path $stageDir ([guid]::NewGuid().ToString("N") + $ppFormats[$format]))
    if ($password) {
        $pres.Password = $password
    }
    $pres.SaveAs($stage, $format)
    $pres.Close()

    moveToOut $stage $relPath
}

function simplePres {
    # タイトルスライド 1 枚のプレゼンテーションを保存する
    param (
        [string]$relPath,
        [string]$marker,
        [int]$format = 24,
        [string]$password
    )

    $pres = $ppt.Presentations.Add(0)  # ウィンドウなし
    $slide = addSlide $pres $ppLayoutTitle $marker
    $slide.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "$([System.IO.Path]::GetFileName($relPath)) のデータ"
    savePres $pres $relPath $format $password
}

function closeAll {
    # 開いたままのブック・文書・プレゼンテーションを閉じる（ケースが途中で失敗したとき用）
    if ($null -ne $excel) {
        foreach ($wb in @($excel.Workbooks)) {
            $wb.Close($false)
        }
    }
    if ($null -ne $word) {
        foreach ($doc in @($word.Documents)) {
            $doc.Close(0)
        }
    }
    if ($null -ne $ppt) {
        foreach ($pres in @($ppt.Presentations)) {
            $pres.Close()
        }
    }
}

function releaseApp {
    param (
        $app
    )

    closeAll
    $app.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

function makePng {
    # テスト用の PNG 画像を作り、そのパスを返す
    param (
        [string]$text
    )

    Add-Type -AssemblyName System.Drawing
    $path = Join-Path $stageDir ([guid]::NewGuid().ToString("N") + ".png")
    $bitmap = New-Object System.Drawing.Bitmap 400, 120
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.Clear([System.Drawing.Color]::White)
        $font = New-Object System.Drawing.Font("Yu Gothic UI", 14)
        $graphics.DrawString($text, $font, [System.Drawing.Brushes]::Black, 10, 40)
        $font.Dispose()
    } finally {
        $graphics.Dispose()
    }
    $bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bitmap.Dispose()

    return $path
}

function copyOut {
    # 出力フォルダ内のファイルを別の名前でコピーする
    param (
        [string]$srcRelPath,
        [string]$destRelPath,
        [bool]$quiet = $false
    )

    $dest = Join-Path $OutDir $destRelPath
    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($dest)) | Out-Null
    [System.IO.File]::Copy((Join-Path $OutDir $srcRelPath), $dest, $true)
    if (-not $quiet) {
        Write-Host "  $destRelPath"
    }
}

function runCase {
    # 1 ケースを実行する。失敗しても残りのケースは続ける
    param (
        [string]$id,
        [scriptblock]$body
    )

    Write-Host "[$id]" -ForegroundColor Cyan
    try {
        & $body
    } catch {
        Write-Host "  失敗: $($_.Exception.Message)" -ForegroundColor Red
        $failed.Add($id)
        closeAll
    }
}

# 出力フォルダを作り直す（読み取り専用・隠し属性のファイルも消す）
foreach ($dir in @($OutDir, $ConfigDir, $stageDir)) {
    if (Test-Path -LiteralPath $dir) {
        Get-ChildItem -LiteralPath $dir -Recurse -Force | ForEach-Object { $_.Attributes = "Normal" }
        Remove-Item -LiteralPath $dir -Recurse -Force
    }
    [System.IO.Directory]::CreateDirectory($dir) | Out-Null
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false

try {
    runCase "TC01 基本" {
        $wb = newBook @("Sheet1", "売上", "2024_上期")

        setRows $wb.Worksheets.Item(1) "A1" @(
            @("社員ID", "氏名", "部署", "内線", "電話番号", "メール", "備考"),
            @("TC01-001", "山田 太郎", "営業部", 1234, "03-1234-5678", "taro.yamada@example.com", ""),
            @("TC01-002", "佐藤 花子", "総務部", 2345, "06-2345-6789", "hanako.sato@example.com", "産休中"),
            @("TC01-003", "鈴木 一郎", "開発部", 3456, "045-345-6789", "ichiro.suzuki@example.com", ""),
            @("TC01-004", "高橋 美咲", "開発部", 3457, "0120-111-222", "misaki.takahashi@example.com", "リーダー"),
            @("TC01-005", "田中 健", "経理部", 4567, "03-9876-5432", "ken.tanaka@example.com", ""),
            @("TC01-006", "Smith John", "海外事業部", 5678, "+81-3-1111-2222", "john.smith@example.com", "English OK")
        )

        $ws = $wb.Worksheets.Item(2)
        $rows = @(, @("月", "売上金額", "前年比", "担当", "マーカー"))
        for ($m = 1; $m -le 12; $m++) {
            $rows += , @("2024年${m}月", (1000000 + $m * 123456), (0.9 + $m * 0.02), "山田", "TC01-売上")
        }
        setRows $ws "A1" $rows
        $ws.Range("B2:B13").NumberFormat = "#,##0"
        $ws.Range("C2:C13").NumberFormat = "0.0%"

        setRows $wb.Worksheets.Item(3) "A1" @(
            @("期間", "項目", "金額", "マーカー"),
            @("2024上期", "売上高", 50000000, "TC01-上期"),
            @("2024上期", "営業利益", 7500000, "TC01-上期")
        )

        saveBook $wb "基本.xlsx"
    }

    runCase "TC02 設計書の例（サブフォルダ）" {
        $wb = newBook @("表紙", "明細")
        setRows $wb.Worksheets.Item(1) "B2" @(
            @("御見積書"),
            @("A社 御中"),
            @("見積番号", "TC02-2024-001")
        )
        setRows $wb.Worksheets.Item(2) "A1" @(
            @("No", "品名", "数量", "単価", "金額"),
            @(1, "ノートPC", 10, 150000, "=C2*D2"),
            @(2, "モニター 27インチ", 10, 40000, "=C3*D3"),
            @(3, "保守サービス（1年）", 1, 200000, "=C4*D4"),
            @("", "合計", "", "", "=SUM(E2:E4)")
        )
        $wb.Worksheets.Item(2).Range("D2:E5").NumberFormat = "#,##0"
        saveBook $wb "2024\見積\A社.xlsx"
    }

    runCase "TC03 別フォルダの同名ブック" {
        $wb = newBook @("表紙", "明細")
        setRows $wb.Worksheets.Item(1) "B2" @(
            @("御見積書（2025年度）"),
            @("A社 御中"),
            @("見積番号", "TC03-2025-001")
        )
        setRows $wb.Worksheets.Item(2) "A1" @(
            @("No", "品名", "数量", "単価", "金額"),
            @(1, "ノートPC", 12, 160000, "=C2*D2")
        )
        saveBook $wb "2025\見積\A社.xlsx"
    }

    runCase "TC04 旧形式 .xls" {
        $wb = newBook @("旧形式シート", "Sheet2")
        setRows $wb.Worksheets.Item(1) "A1" @(
            @("ID", "内容"),
            @("TC04", "Excel 97-2003 形式のデータ"),
            @("TC04", "日本語　全角スペース入り")
        )
        setRows $wb.Worksheets.Item(2) "A1" @(, @("TC04", "2枚目のシート"))
        saveBook $wb "旧形式.xls" 56
    }

    runCase "TC05 マクロ有効ブック .xlsm・バイナリブック .xlsb" {
        simpleBook "マクロ有効.xlsm" "TC05" 52 "マクロシート"
        simpleBook "バイナリ.xlsb" "TC05" 50 "バイナリシート"
    }

    runCase "TC06 大文字の拡張子" {
        simpleBook "大文字拡張子.XLSX" "TC06"
        simpleBook "大文字拡張子旧形式.XLS" "TC06" 56
    }

    runCase "TC07 変換対象外の拡張子" {
        simpleBook "対象外\対象外.xltx" "TC07" 54
        simpleBook "対象外\対象外.csv" "TC07" 6
        writeText "対象外\対象外.txt" "TC07`tテキストファイル"
        writeText "対象外\対象外.tsv" "TC07`t変換対象フォルダ内の tsv"
        # 拡張子が .xls で始まるだけのもの
        writeText "対象外\対象外.xlsx.bak" "TC07"
        writeText "対象外\対象外.xls~" "TC07"
    }

    runCase "TC08 シートの表示状態・種類" {
        $wb = newBook @("表示", "非表示", "超非表示", "空シート")
        setRows $wb.Worksheets.Item(1) "A1" @(
            @("項目", "値"),
            @("TC08 表示シートのデータ", 10),
            @("りんご", 30),
            @("みかん", 20)
        )
        setRows $wb.Worksheets.Item(2) "A1" @(, @("TC08 非表示シートのデータ"))
        setRows $wb.Worksheets.Item(3) "A1" @(, @("TC08 超非表示シートのデータ"))

        $chart = $wb.Charts.Add($missing, $wb.Sheets.Item($wb.Sheets.Count))
        $chart.SetSourceData($wb.Worksheets.Item(1).Range("A2:B4"))
        $chart.HasTitle = $true
        $chart.ChartTitle.Text = "TC08 グラフシートのタイトル"
        $chart.Name = "グラフ"

        $wb.Worksheets.Item("非表示").Visible = $xlSheetHidden
        $wb.Worksheets.Item("超非表示").Visible = $xlSheetVeryHidden
        $wb.Worksheets.Item(1).Activate()
        saveBook $wb "シート表示状態.xlsx"
    }

    runCase "TC09 シート名の記号" {
        # Excel のシート名に使えない : \ / ? * [ ] 以外の記号を使う
        $names = @(
            "不等号<>",
            "引用符`"x`"",
            "縦棒|",
            "スペース 入り",
            "全角　スペース",
            "R&D's #1 (改)",
            "旧.xls_版",
            "Ｓｈｅｅｔ全角英字",
            "${sushi}寿司",
            "1234567890123456789012345678901",  # 31 文字（上限）
            "衝突`"",
            "衝突$([char]0x201D)",  # ” は以前の版（記号を全角に置換）で " と同じ TSV 名になっていた組。別の TSV になることを確かめる（Excel のシート名は全角半角を区別しないため ＜ と < の組は作れない）
            "タブ`tあり",   # Excel のシート名にはタブ・改行も付けられる（以前の版は、ファイル名に使えないためそのまま使って失敗していた）
            "改行`nあり"
        )
        $wb = newBook $names
        for ($i = 1; $i -le $names.Count; $i++) {
            setRows $wb.Worksheets.Item($i) "A1" @(, @("TC09-$('{0:D2}' -f $i)", "シート名: $($names[$i - 1])"))
        }
        saveBook $wb "シート名記号.xlsx"
    }

    runCase "TC10 セルの内容" {
        $wb = newBook @("文字列", "数値と日付", "数式", "レイアウト", "オブジェクト")

        # --- 文字列 ---
        $ws = $wb.Worksheets.Item(1)
        $cases = @(
            @("TC10-01", "セル内改行", "1行目`n2行目`n3行目"),
            @("TC10-02", "コロンを含む", "URL: https://example.com:8080/path?q=1&r=2"),
            @("TC10-03", "ダブルクォート", '彼は"こんにちは"と言った'),
            @("TC10-04", "ダブルクォートと改行", "`"引用`"`n改行後"),
            @("TC10-05", "セル内タブ", "左`t右"),
            @("TC10-06", "前後の半角空白", "   前後に空白   "),
            @("TC10-07", "全角スペース", "全角　スペース　区切り"),
            @("TC10-08", "正規表現の記号", 'a.b(c)[d]{e}*f+g?h^i$j|k\l'),
            @("TC10-09", "大文字小文字・全角英字", "ABC abc Ａｂｃ ａｂｃ"),
            @("TC10-10", "CP932外の文字", "${yoshi}野家 ${sushi} 한국어 Ümlaut ñ"),
            @("TC10-11", "機種依存文字", "①②③ ㈱ ㌔ Ⅷ ～ − ∥ ￢"),
            @("TC10-12", "半角カナ", "ﾊﾝｶｸｶﾀｶﾅ ﾃｽﾄ"),
            @("TC10-13", "先頭ゼロの文字列", "00123"),
            @("TC10-14", "先頭アポストロフィ", "'=SUM(A1:A3)"),
            @("TC10-15", "HTML風", '<a href="x">&amp;</a>'),
            @("TC10-16", "Windowsパス", 'C:\Users\test\ファイル.txt'),
            @("TC10-17", "UNCパス", '\\server\share\フォルダ'),
            @("TC10-18", "長い文字列", (("あいうえお" * 1000) + "TC10長文末尾")),
            @("TC10-19", "カンマ・セミコロン", "a,b;c"),
            @("TC10-20", "CRのみの改行", "CR`rで改行"),
            @("TC10-21", "メールアドレス", "test.user+tag@example.co.jp"),
            @("TC10-22", "郵便番号・電話番号", "〒100-0001 03-1234-5678"),
            @("TC10-23", "英文", "The quick brown fox jumps over the lazy dog.")
        )
        setRows $ws "A1" @(, @("ID", "説明", "値", "右隣"))
        for ($i = 0; $i -lt $cases.Count; $i++) {
            $row = $i + 2
            setRows $ws "A$row" @(, @($cases[$i][0], $cases[$i][1]))
            $format = if ($cases[$i][0] -eq "TC10-13") { "@" } else { "" }
            setCell $ws "C$row" $cases[$i][2] $format
            setCell $ws "D$row" "$($cases[$i][0])右隣"
        }
        $ws.Range("C2").WrapText = $true
        # 空白だけの行（変換時に削除される）、完全な空行
        $row = $cases.Count + 3
        setRows $ws "A$row" @(, @(" ", "　", "`t"))
        $row += 3
        setRows $ws "A$row" @(, @("TC10-99", "空行のあと", "最終行"))

        # --- 数値と日付 ---
        $ws = $wb.Worksheets.Item(2)
        $time = (12 * 3600 + 34 * 60 + 56) / 86400
        $numbers = @(
            @("TC10-N01", "標準の整数", 1234567, ""),
            @("TC10-N02", "桁区切り", 1234567, "#,##0"),
            @("TC10-N03", "円表記", 1234567, '#,##0"円"'),
            @("TC10-N04", "小数2桁", 3.14159, "0.00"),
            @("TC10-N05", "百分率", 0.125, "0.0%"),
            @("TC10-N06", "負数を▲表示", -1500, '#,##0;"▲"#,##0'),
            @("TC10-N07", "指数", 1.23E+20, ""),
            @("TC10-N08", "15桁超の整数", 12345678901234567890, ""),
            @("TC10-N09", "分数", 0.75, "# ?/?"),
            @("TC10-N10", "日付", 45383, "yyyy/mm/dd"),
            @("TC10-N11", "和暦", 45383, '[$-ja-JP]ggge"年"m"月"d"日"'),
            @("TC10-N11b", "和暦(LCID指定)", 45383, '[$-411]ggge"年"m"月"d"日"'),
            @("TC10-N12", "曜日付き", 45383, '[$-ja-JP]yyyy/m/d(aaa)'),
            @("TC10-N13", "時刻", $time, "hh:mm:ss"),
            @("TC10-N14", "日時", (45383 + $time), "yyyy/mm/dd hh:mm"),
            @("TC10-N15", "経過時間", 1.5, "[h]:mm"),
            @("TC10-N16", "郵便番号書式", 1000001, "000-0000"),
            @("TC10-N17", "真偽値", $true, ""),
            @("TC10-N18", "ゼロ", 0, ""),
            @("TC10-N19", "ゼロを非表示", 0, '#,##0;-#,##0;'),
            @("TC10-N20", "文字列書式の数値", "12345", "@")
        )
        setRows $ws "A1" @(, @("ID", "説明", "値", "書式"))
        for ($i = 0; $i -lt $numbers.Count; $i++) {
            $row = $i + 2
            setRows $ws "A$row" @(, @($numbers[$i][0], $numbers[$i][1]))
            setCell $ws "C$row" $numbers[$i][2] $numbers[$i][3]
            setCell $ws "D$row" $(if ($numbers[$i][3]) { $numbers[$i][3] } else { "標準" }) "@"
        }
        $ws.Columns.Item(3).ColumnWidth = 30
        # 列幅が足りず ##### 表示になる日付
        setCell $ws "F2" "TC10-N21 列幅不足"
        setCell $ws "G2" 45383 "yyyy/mm/dd"
        $ws.Columns.Item(7).ColumnWidth = 3

        # --- 数式 ---
        $ws = $wb.Worksheets.Item(3)
        setRows $ws "A1" @(
            @("ID", "説明", "結果", "数式"),
            @("TC10-F01", "足し算", "=1+1"),
            @("TC10-F02", "文字列連結", '="連結"&"結果"'),
            @("TC10-F03", "0除算", "=1/0"),
            @("TC10-F04", "NA", "=NA()"),
            @("TC10-F05", "VLOOKUP", '=VLOOKUP("TC10-F01",A:C,2,FALSE)'),
            @("TC10-F06", "空文字", '=""'),
            @("TC10-F07", "繰り返し", '=REPT("繰返",3)'),
            @("TC10-F08", "TEXT関数", '=TEXT(45383,"yyyy年m月d日")'),
            @("TC10-F08b", "TEXT関数(和暦)", '=TEXT(45383,"[$-411]ggge年m月d日")'),
            @("TC10-F09", "他シート参照", "=文字列!C3"),
            @("TC10-F10", "HYPERLINK関数", '=HYPERLINK("https://example.com/","HYPERLINK関数の表示")'),
            @("TC10-F11", "改行を含む結果", '="数式の"&CHAR(10)&"改行"'),
            @("TC10-F12", "名前エラー", "=存在しない関数()")
        )
        # D 列に数式の文字列を置く（数式そのものは検索対象外だが、文字列として置いたものは検索できる）
        for ($row = 2; $row -le 14; $row++) {
            setCell $ws "D$row" ("'" + $ws.Range("C$row").Formula)
        }

        # --- レイアウト ---
        $ws = $wb.Worksheets.Item(4)
        # 先頭の行・列を空けて D5 から表を置く
        setRows $ws "D5" @(
            @("TC10-L01", "左上が空いた表", "値1"),
            @("TC10-L02", "", "間に空セル"),
            @("", "", "TC10-L03 先頭が空")
        )
        # 結合セル
        setCell $ws "B2" "TC10-L04 結合セルの見出し"
        $ws.Range("B2:F2").Merge()
        # 空行を挟んだ後の表（非表示行・非表示列・オートフィルタ）
        setRows $ws "A11" @(
            @("区分", "内容", "備考"),
            @("表示", "TC10-L05 表示行", ""),
            @("非表示", "TC10-L06 非表示行", ""),
            @("絞込", "TC10-L07 フィルタで隠れた行", ""),
            @("表示", "TC10-L08 表示行2", "")
        )
        setCell $ws "H12" "TC10-L09 非表示列"
        # 空白だけの行
        setRows $ws "A18" @(, @(" ", "　"))
        # 横に長い行（300列）
        $wide = @()
        for ($c = 1; $c -le 300; $c++) {
            $wide += "列$('{0:D3}' -f $c)"
        }
        $wide[299] = "TC10-L10 300列目"
        setRows $ws "A20" @(, $wide)
        # 値が無く書式だけ設定された遠いセル（使用範囲が広がる）
        $ws.Range("AZ300").Interior.Color = 65535
        # 非表示・フィルタは値を書き終えてから設定する（設定後に配列を書くと値がずれる）
        $ws.Rows.Item(13).Hidden = $true
        $ws.Columns.Item(8).Hidden = $true
        [void]$ws.Range("A11:C15").AutoFilter(1, "<>絞込")

        # --- オブジェクト（検索対象外の場所にある文字列）---
        $ws = $wb.Worksheets.Item(5)
        setRows $ws "A1" @(
            @("ID", "種類", "セル値"),
            @("TC10-O01", "コメント付きセル", "コメントがあるセル"),
            @("TC10-O02", "テキストボックス", "右にテキストボックス"),
            @("TC10-O03", "ハイパーリンク", ""),
            @("TC10-O04", "入力規則", "選択肢A")
        )
        [void]$ws.Range("C2").AddComment("TC10 コメント内のテキスト")
        $box = $ws.Shapes.AddTextbox(1, 300, 20, 220, 40)
        $box.TextFrame2.TextRange.Text = "TC10 テキストボックス内のテキスト"
        [void]$ws.Hyperlinks.Add($ws.Range("C4"), "https://example.com/TC10-hidden-url", $missing, $missing, "リンクの表示文字列")
        [void]$ws.Range("C5").Validation.Add(3, 1, 1, "選択肢A,選択肢B")
        try {
            $ws.PageSetup.CenterHeader = "TC10 ヘッダーのテキスト"
        } catch {
            Write-Host "  （プリンタが無いためヘッダーは設定できませんでした）" -ForegroundColor Yellow
        }

        saveBook $wb "セル内容.xlsx"
    }

    runCase "TC11 大量データ" {
        $wb = newBook @("大量", "横長")
        $products = @("りんご", "みかん", "バナナ", "ぶどう", "もも", "メロン", "いちご")
        $cols = 12
        $data = New-Object 'object[,]' ($BigRows + 1), $cols
        $data[0, 0] = "ID"; $data[0, 1] = "商品"; $data[0, 2] = "数量"
        for ($c = 3; $c -lt $cols; $c++) {
            $data[0, $c] = "項目$c"
        }
        for ($r = 1; $r -le $BigRows; $r++) {
            $data[$r, 0] = "TC11-{0:D5}" -f $r
            $data[$r, 1] = $products[$r % $products.Count]
            $data[$r, 2] = ($r * 7) % 1000
            for ($c = 3; $c -lt $cols; $c++) {
                $data[$r, $c] = "値$r-$c"
            }
        }
        $data[$BigRows, ($cols - 1)] = "TC11 大量データ最終行"
        $wb.Worksheets.Item(1).Range("A1").Resize($BigRows + 1, $cols).Value2 = $data

        $wideCols = 2000
        $wide = New-Object 'object[,]' 1, $wideCols
        for ($c = 0; $c -lt $wideCols; $c++) {
            $wide[0, $c] = "C$($c + 1)"
        }
        $wide[0, ($wideCols - 1)] = "TC11 横長最終列"
        $wb.Worksheets.Item(2).Range("A1").Resize(1, $wideCols).Value2 = $wide

        saveBook $wb "大量データ.xlsx"
    }

    runCase "TC29 使用範囲が膨らんだシート" {
        # 最終行・右端のセルに書式だけが残り、使用範囲がシート全体（約 172 億セル）になったブック。
        # データの範囲だけを一時シートにコピーしてから書き出す処理（copyDataRangeToTempSheet）が無いと、
        # Excel のテキスト保存が空セルのタブだけで数GBを書き出し、制限時間（10 分）を超えて失敗する
        $wb = newBook @("肥大", "肥大_C3から", "普通")
        $ws = $wb.Worksheets.Item(1)
        setRows $ws "A1" @(
            @("ID", "内容", "数量"),
            @("TC29-01", "使用範囲が膨らんだシートのデータ", 10),
            @("TC29-02", "りんご", 20)
        )
        setCell $ws "E2" '=COUNTA(A200:A300)'   # すべて空の範囲を参照する数式（行を削除すると #REF! になる）
        setCell $ws "E3" '=普通!A1'             # 別シート参照
        $ws.Range("XFD1").Interior.Color = 65535
        $ws.Range("A1048576").Interior.Color = 65535

        # データが A1 から始まらない場合も、TSV の N 行目・k 列目がシートの N 行目・k 列目になることを確かめる
        $ws = $wb.Worksheets.Item(2)
        setRows $ws "C3" @(
            @("TC29-03", "C3 から始まる表"),
            @("TC29-04", "みかん")
        )
        $ws.Range("XFD1").Interior.Color = 65535
        $ws.Range("A1048576").Interior.Color = 65535

        setRows $wb.Worksheets.Item(3) "A1" @(, @("TC29-05", "普通のシート"))
        $wb.Worksheets.Item(1).Activate()
        saveBook $wb "使用範囲肥大.xlsx"
    }

    runCase "TC12 空のブック" {
        $wb = newBook @("空")
        saveBook $wb "空ブック.xlsx"
    }

    runCase "TC13 シートが多いブック" {
        $names = @()
        for ($i = 1; $i -le 60; $i++) {
            $names += "S$('{0:D2}' -f $i)"
        }
        $wb = newBook $names
        for ($i = 1; $i -le 60; $i++) {
            setRows $wb.Worksheets.Item($i) "A1" @(, @("TC13", "シート $($names[$i - 1])"))
        }
        saveBook $wb "多数シート.xlsx"
    }

    runCase "TC14 保護" {
        $wb = newBook @("保護シート", "通常シート")
        setRows $wb.Worksheets.Item(1) "A1" @(, @("TC14-1", "保護されたシートのデータ"))
        setRows $wb.Worksheets.Item(2) "A1" @(, @("TC14-1", "通常シートのデータ"))
        $wb.Worksheets.Item(1).Protect("sheetpw")
        saveBook $wb "保護\シート保護.xlsx"

        $wb = newBook @("構成保護1", "構成保護2")
        setRows $wb.Worksheets.Item(1) "A1" @(, @("TC14-2", "ブックの構成が保護されたシート1"))
        setRows $wb.Worksheets.Item(2) "A1" @(, @("TC14-2", "ブックの構成が保護されたシート2"))
        $wb.Protect("bookpw", $true, $false)
        saveBook $wb "保護\ブック構成保護.xlsx"

        $wb = newBook @("書込保護")
        setRows $wb.Worksheets.Item(1) "A1" @(, @("TC14-3", "書き込みパスワード付き"))
        saveBook $wb "保護\書き込みパスワード付き.xlsx" 51 "" "writepw"
    }

    runCase "TC15 異常系" {
        # 読み取りパスワード付き（開けないので変換失敗になる想定）
        $wb = newBook @("秘密")
        setRows $wb.Worksheets.Item(1) "A1" @(, @("TC15-1", "読み取りパスワード付き"))
        saveBook $wb "異常系\読み取りパスワード付き.xlsx" 51 "openpw"

        writeText "異常系\壊れたファイル.xlsx" "TC15-2 これは Excel ファイルではありません"
        writeText "異常系\空ファイル.xlsx" "" (New-Object System.Text.UTF8Encoding($false))

        writeLockFile "異常系\~`$ロックファイル.xlsx"
        writeLockFile "異常系\~`$ロックファイル_隠し属性なし.xlsx" $false

        # 拡張子と中身の形式が違うもの
        $wb = newBook @("偽装")
        setRows $wb.Worksheets.Item(1) "A1" @(, @("TC15-3", "中身は xls 形式"))
        saveBook $wb "異常系\中身はxls.xlsx" 56

        $wb = newBook @("偽装")
        setRows $wb.Worksheets.Item(1) "A1" @(, @("TC15-4", "中身は xlsx 形式"))
        saveBook $wb "異常系\中身はxlsx.xls" 51

        writeText "異常系\中身はCSV.xls" "ID,内容`r`nTC15-5,中身は CSV`r`n" ([System.Text.Encoding]::GetEncoding(932))
        writeText "異常系\中身はHTML.xls" @"
<html><head><meta charset="utf-8"></head><body>
<table><tr><td>ID</td><td>内容</td></tr><tr><td>TC15-6</td><td>中身は HTML の表</td></tr></table>
</body></html>
"@
    }

    runCase "TC16 ファイル名" {
        simpleBook "ファイル名\スペース 入り ブック.xlsx" "TC16-01"
        simpleBook "ファイル名\[確定]報告書.xlsx" "TC16-02"
        simpleBook "ファイル名\R&D_'24 #1 (最終版).xlsx" "TC16-03"
        simpleBook "ファイル名\v1.2.3.xlsx" "TC16-04"
        simpleBook "ファイル名\コピー.xls_old.xlsx" "TC16-05"
        simpleBook "ファイル名\${sushi}メニュー.xlsx" "TC16-06"
        simpleBook "ファイル名\한글파일.xlsx" "TC16-07"
        simpleBook "ファイル名\ＺＥＮＫＡＫＵ全角英字.xlsx" "TC16-08"
        simpleBook "ファイル名\ﾊﾝｶｸｶﾅ.xlsx" "TC16-09"
        simpleBook "ファイル名\ア_イ_ウ.xlsx" "TC16-10" 51 "シ_ー_ト"
        simpleBook "ファイル名\$('長い名前' * 15).xlsx" "TC16-11"
        simpleBook "ファイル名\$('とても長い名前のファイル' * 12).xlsx" "TC16-12"
        simpleBook "ファイル名\読み取り専用属性.xlsx" "TC16-13"
        # 以前の版では、この 2 つは同じ TSV 名（A.xlsx_old.xlsx_1.tsv）になり、互いの TSV を上書き・削除していた
        simpleBook "ファイル名\A.xlsx" "TC16-14" 51 "old.xlsx_1"
        simpleBook "ファイル名\A.xlsx_old.xlsx" "TC16-15" 51 "1"
        simpleBook "ファイル名\終わりが_.xlsx" "TC16-16" 51 "_"
        simpleBook "ファイル名\100%.xlsx" "TC16-17" 51 "100%"
        [System.IO.File]::SetAttributes((Join-Path $OutDir "ファイル名\読み取り専用属性.xlsx"), "ReadOnly")
    }

    runCase "TC17 フォルダ" {
        simpleBook "フォルダ\スペース あり\スペースフォルダ内.xlsx" "TC17-1"
        simpleBook "フォルダ\[角括弧]\角括弧フォルダ内.xlsx" "TC17-2"
        simpleBook "フォルダ\深い\階層\の\フォルダ\テスト\深い階層.xlsx" "TC17-3"
        simpleBook "フォルダ\ドット.付き.フォルダ\ドットフォルダ内.xlsx" "TC17-4"
        simpleBook "フォルダ\拡張子付きフォルダ.xlsx\中身.xlsx" "TC17-5"
        simpleBook "フォルダ\隠しフォルダ\隠しフォルダ内.xlsx" "TC17-6"
        $hidden = New-Object System.IO.DirectoryInfo (Join-Path $OutDir "フォルダ\隠しフォルダ")
        $hidden.Attributes = $hidden.Attributes -bor [System.IO.FileAttributes]::Hidden
        simpleBook "フォルダ\隠しファイル.xlsx" "TC17-7"
        [System.IO.File]::SetAttributes((Join-Path $OutDir "フォルダ\隠しファイル.xlsx"), "Hidden")
        [System.IO.Directory]::CreateDirectory((Join-Path $OutDir "フォルダ\空フォルダ")) | Out-Null
        Write-Host "  フォルダ\空フォルダ"
    }

    runCase "TC26 ファイル名・フォルダ（追加）" {
        # ファイル名の正規化（NFC / NFD）。見た目は同じだがファイル名のコードが違う
        simpleBook "ファイル名\ガ_合成済み.xlsx" "TC26-01"
        copyOut "ファイル名\ガ_合成済み.xlsx" ("ファイル名\カ" + [char]0x3099 + "_結合文字.xlsx")
        # 拡張子が無いファイル（中身は xlsx）
        copyOut "基本.xlsx" "ファイル名\拡張子なし"
        # システム属性・アーカイブ属性
        simpleBook "ファイル名\システム属性.xlsx" "TC26-02"
        [System.IO.File]::SetAttributes((Join-Path $OutDir "ファイル名\システム属性.xlsx"), "System")
        # インターネットからダウンロードした印（Zone.Identifier）。Excel は保護ビューで開く
        simpleBook "フォルダ\ダウンロード\ネットから取得.xlsx" "TC26-03"
        Set-Content -LiteralPath (Join-Path $OutDir "フォルダ\ダウンロード\ネットから取得.xlsx") -Stream "Zone.Identifier" -Value "[ZoneTransfer]`r`nZoneId=3`r`nHostUrl=https://example.com/ネットから取得.xlsx"
        Write-Host "  フォルダ\ダウンロード\ネットから取得.xlsx（Zone.Identifier 付き）"
        # ファイル数の多いフォルダ
        simpleBook "フォルダ\ファイル数が多い\ファイル001.xlsx" "TC26-04"
        for ($i = 2; $i -le $ManyFiles; $i++) {
            copyOut "フォルダ\ファイル数が多い\ファイル001.xlsx" ("フォルダ\ファイル数が多い\ファイル{0:D3}.xlsx" -f $i) $true
        }
        Write-Host "  フォルダ\ファイル数が多い\ファイル001.xlsx 〜 ファイル$('{0:D3}' -f $ManyFiles).xlsx"

        if ($WithRisky) {
            # フォルダのループ（ジャンクション）。再帰の検索が終わらなくなる可能性がある
            $loopDir = Join-Path $OutDir "フォルダ\ループ"
            [System.IO.Directory]::CreateDirectory($loopDir) | Out-Null
            simpleBook "フォルダ\ループ\ループ内.xlsx" "TC26-05"
            cmd /c mklink /J "${loopDir}\自分の親へのジャンクション" "${loopDir}" | Out-Null
            Write-Host "  フォルダ\ループ\自分の親へのジャンクション（ジャンクション）"
            # 260 文字を超えるパス
            $deep = Join-Path $OutDir "フォルダ\長いパス"
            $name = "階層" * 20
            for ($i = 1; $i -le 4; $i++) {
                $deep = Join-Path $deep $name
            }
            [System.IO.Directory]::CreateDirectory("\\?\$deep") | Out-Null
            [System.IO.File]::Copy((Join-Path $OutDir "基本.xlsx"), "\\?\$deep\長いパスの先.xlsx", $true)
            Write-Host "  フォルダ\長いパス\…（260 文字を超えるパス）"
        }
    }

    runCase "TC18 同名で拡張子違い" {
        simpleBook "同名ブック\同名.xlsx" "TC18-xlsx"
        simpleBook "同名ブック\同名.xls" "TC18-xls" 56
        simpleBook "同名ブック\同名.xlsm" "TC18-xlsm" 52
    }

    runCase "TC19 外部リンク" {
        $wb = newBook @("リンク")
        $ws = $wb.Worksheets.Item(1)
        setRows $ws "A1" @(, @("TC19", "外部ブックの値→"))
        $ws.Range("C1").Formula = "='$OutDir\[基本.xlsx]Sheet1'!`$B`$2"
        saveBook $wb "外部リンク.xlsx"
    }

    runCase "TC22 同名で種類違い（Excel）" {
        simpleBook "混在\同名.xlsx" "TC22-xlsx"
    }

    runCase "TC23 Unicode・特殊な文字" {
        $wb = newBook @("Unicode", "長い文字列")

        # 合成済み・結合文字・異体字セレクタなど、見た目が同じでもコードが違う文字
        $ws = $wb.Worksheets.Item(1)
        $cases = @(
            @("TC23-01", "合成済みと結合文字（NFC / NFD）", "ガ", ("カ" + [char]0x3099)),
            @("TC23-02", "異体字セレクタ", "葛", ("葛" + [char]::ConvertFromUtf32(0xE0100))),
            @("TC23-03", "ゼロ幅スペース", "山田太郎", ("山田" + [char]0x200B + "太郎")),
            @("TC23-04", "ノーブレークスペース", "山田 太郎", ("山田" + [char]0x00A0 + "太郎")),
            @("TC23-05", "ソフトハイフン", "TC23-05データ", ("TC23-05" + [char]0x00AD + "データ")),
            @("TC23-06", "BOM（U+FEFF）が途中にある", "TC23-06データ", ("TC23-06" + [char]0xFEFF + "データ")),
            @("TC23-07", "右から左に書く文字", "ヘブライ語・アラビア語", "שלום مرحبا"),
            @("TC23-08", "全角英数・ローマ数字・丸数字", "１２３ＡＢＣ", "１２３ＡＢＣ Ⅷ ⅷ ①②"),
            @("TC23-09", "旧字体・異体字", "髙﨑濵", "髙橋 﨑田 濵田 齋藤 邊"),
            @("TC23-10", "タイ語・デーヴァナーガリー", "多言語", "สวัสดี नमस्ते Привет Ελληνικά"),
            @("TC23-11", "絵文字（ZWJ・肌の色・国旗）", "絵文字", ([char]::ConvertFromUtf32(0x1F468) + [char]0x200D + [char]::ConvertFromUtf32(0x1F469) + [char]0x200D + [char]::ConvertFromUtf32(0x1F467) + " " + [char]::ConvertFromUtf32(0x1F44D) + [char]::ConvertFromUtf32(0x1F3FD) + " " + [char]::ConvertFromUtf32(0x1F1EF) + [char]::ConvertFromUtf32(0x1F1F5))),
            @("TC23-12", "サロゲートペアの漢字", "魚へんの漢字", ([char]::ConvertFromUtf32(0x29E3D) + [char]::ConvertFromUtf32(0x20BB7) + [char]::ConvertFromUtf32(0x21237))),
            @("TC23-13", "上付き・下付き・数学記号", "数式風", "x² H₂O ∑ ∫ ≠ ± √2 ∞"),
            @("TC23-14", "通貨記号", "通貨", "€ £ ¥ ₩ ₹ ¢ ＄"),
            @("TC23-15", "ハイフンに見える文字", "ハイフン類", "- ‐ ‑ – — − ― ｰ"),
            @("TC23-16", "引用符に見える文字", "引用符類", ("'apostrophe' ""quote"" " + [char]0x2018 + "single" + [char]0x2019 + " " + [char]0x201C + "double" + [char]0x201D + " 「かぎ」『二重かぎ』")),
            @("TC23-17", "半角カナの濁点", "半角カナ", "ｶﾞｷﾞｸﾞ ﾊﾟﾋﾟﾌﾟ"),
            @("TC23-18", "囲み文字・単位記号", "単位", "㎡ ㎏ ㌢ ℡ № ㈲ ㊙"),
            @("TC23-19", "改行に見える文字（行区切り U+2028）", "行区切り", ("TC23-19前" + [char]0x2028 + "TC23-19後")),
            @("TC23-20", "連続する空白とタブ", "空白", "前        後	タブ")
        )
        setRows $ws "A1" @(, @("ID", "説明", "比較用（ふつうの文字）", "特殊な文字"))
        for ($i = 0; $i -lt $cases.Count; $i++) {
            $row = $i + 2
            setRows $ws "A$row" @(, @($cases[$i][0], $cases[$i][1], $cases[$i][2]))
            setCell $ws "D$row" $cases[$i][3]
        }

        # 文字数の上限（32767 文字）ちょうどのセル
        $ws = $wb.Worksheets.Item(2)
        $long = "TC23-21先頭" + ("あ" * 32749) + "TC23-21末尾"  # 合計 32767 文字（セルの上限）
        setRows $ws "A1" @(, @("ID", "説明", "文字数"))
        setRows $ws "A2" @(, @("TC23-21", "32767 文字ちょうどのセル", $long.Length))
        setCell $ws "D2" $long
        setRows $ws "A3" @(, @("TC23-22", "長い数式の結果", ""))
        $ws.Range("D3").Formula = '=REPT("繰り返し",200)&"TC23-22末尾"'
        saveBook $wb "特殊文字.xlsx"
    }

    runCase "TC24 Excel の機能" {
        $wb = newBook @("名前と数式", "アウトライン", "テーブルとピボット")

        # --- 名前の定義・配列数式・エラー ---
        $ws = $wb.Worksheets.Item(1)
        setRows $ws "A1" @(
            @("ID", "説明", "結果", "数式（文字列）"),
            @("TC24-01", "名前の定義を使う数式", "", ""),
            @("TC24-02", "配列数式", "", ""),
            @("TC24-03", "循環参照", "", ""),
            @("TC24-04", "#REF! エラー", "", ""),
            @("TC24-05", "#VALUE! エラー", "", ""),
            @("TC24-06", "#NUM! エラー", "", ""),
            @("TC24-07", "#NULL! エラー", "", "")
        )
        [void]$wb.Names.Add("消費税率", "=0.1")
        setCell $ws "F1" 1000
        setCell $ws "F2" 2000
        setCell $ws "G1" 2
        setCell $ws "G2" 3
        $ws.Range("C2").Formula = "=F1*消費税率"
        $ws.Range("C3").FormulaArray = "=SUM(F1:F2*G1:G2)"
        $ws.Range("C5").Formula = "=SUM(H1:H2)"
        $ws.Range("C6").Formula = '=1+"文字列"'
        $ws.Range("C7").Formula = "=SQRT(-1)"
        $ws.Range("C8").Formula = "=SUM(F1:F2 G1:G2)"
        # 列を削除して #REF! を起こす
        $ws.Columns.Item("H").Delete()
        for ($row = 2; $row -le 8; $row++) {
            setCell $ws "D$row" ("'" + $ws.Range("C$row").Formula)
        }
        # 循環参照（反復計算は既定のままなので結果は 0 になる）
        try {
            $excel.Iteration = $false
            $ws.Range("C4").Formula = "=C4+1"
        } catch {
            Write-Host "  （循環参照を作れませんでした）" -ForegroundColor Yellow
        }
        # スピル（動的配列）。古い Excel では #NAME? になる
        setRows $ws "A10" @(, @("TC24-08", "スピル（動的配列）", ""))
        try {
            $ws.Range("C10").Formula2 = "=SEQUENCE(3,2)"
        } catch {
            $ws.Range("C10").Formula = "=SEQUENCE(3,2)"
        }

        # --- アウトライン・非表示・書式 ---
        $ws = $wb.Worksheets.Item(2)
        $rows = @(, @("区分", "内容"))
        $rows += , @("TC24-10", "アウトラインの外")
        for ($i = 1; $i -le 4; $i++) {
            $rows += , @("TC24-11", "折りたたんだ行$i")
        }
        $rows += , @("TC24-12", "行の高さ 0 の行")
        $rows += , @("TC24-13", "ふつうの行")
        setRows $ws "A1" $rows
        $ws.Rows.Item("3:6").Group()
        $ws.Outline.ShowLevels(1) | Out-Null
        $ws.Rows.Item(7).RowHeight = 0
        setCell $ws "D1" "TC24-14 列幅 0 の列"
        $ws.Columns.Item("D").ColumnWidth = 0
        # 条件付き書式・タブ色・印刷範囲・ウィンドウ枠固定（表示値は変わらない）
        [void]$ws.Range("B2:B8").FormatConditions.Add(1, 3, "=""TC24-13""")
        $ws.Tab.Color = 255
        $ws.PageSetup.PrintArea = "A1:B8"

        # --- テーブル（ListObject）とピボットテーブル ---
        $ws = $wb.Worksheets.Item(3)
        $rows = @(, @("日付", "支店", "商品", "金額"))
        $shops = @("東京", "大阪", "名古屋")
        $items = @("りんご", "みかん", "ぶどう")
        for ($i = 0; $i -lt 12; $i++) {
            $rows += , @("2024/{0:D2}/01" -f (($i % 12) + 1), $shops[$i % 3], $items[$i % 3], (10000 + $i * 1000))
        }
        setRows $ws "A1" $rows
        $table = $ws.ListObjects.Add(1, $ws.Range("A1:D13"), $missing, 1)
        $table.Name = "売上テーブル"
        setCell $ws "F1" "TC24-15 テーブル名: 売上テーブル"
        try {
            $pivotSheet = $wb.Worksheets.Add($missing, $wb.Worksheets.Item($wb.Worksheets.Count))
            $pivotSheet.Name = "ピボット"
            $cache = $wb.PivotCaches().Create(1, $ws.Range("A1:D13"))
            $pivot = $cache.CreatePivotTable($pivotSheet.Range("A3"), "TC24ピボット")
            $pivot.PivotFields("支店").Orientation = 1  # xlRowField
            $pivot.PivotFields("商品").Orientation = 2  # xlColumnField
            $pivot.AddDataField($pivot.PivotFields("金額"), "合計 / 金額", -4157) | Out-Null
            setCell $pivotSheet "A1" "TC24-16 ピボットテーブル"
        } catch {
            Write-Host "  （ピボットテーブルを作れませんでした: $($_.Exception.Message)）" -ForegroundColor Yellow
        }

        saveBook $wb "機能.xlsx"
    }

    runCase "TC25 日付の境界・1904 年形式" {
        $wb = newBook @("日付")
        $rows = @(
            @("ID", "説明", "値", "書式"),
            @("TC25-01", "1900/1/1（シリアル値 1）", 1, "yyyy/mm/dd"),
            @("TC25-02", "Excel にだけ存在する 1900/2/29（シリアル値 60）", 60, "yyyy/mm/dd"),
            @("TC25-03", "1900/3/1（シリアル値 61）", 61, "yyyy/mm/dd"),
            @("TC25-04", "9999/12/31（シリアル値の上限）", 2958465, "yyyy/mm/dd"),
            @("TC25-05", "うるう日 2024/2/29", 45351, "yyyy/mm/dd"),
            @("TC25-06", "24 時間を超える時刻", 1.5, "[h]:mm:ss"),
            @("TC25-07", "小数の日付（時刻付き）", 45383.756944, "yyyy/mm/dd hh:mm:ss")
        )
        $ws = $wb.Worksheets.Item(1)
        for ($i = 0; $i -lt $rows.Count; $i++) {
            $row = $i + 1
            setRows $ws "A$row" @(, @($rows[$i][0], $rows[$i][1]))
            if ($i -eq 0) {
                setCell $ws "C$row" $rows[$i][2]
                setCell $ws "D$row" $rows[$i][3]
            } else {
                setCell $ws "C$row" $rows[$i][2] $rows[$i][3]
                setCell $ws "D$row" $rows[$i][3] "@"
            }
        }
        # 負の時刻は ##### 表示になる
        setRows $ws "A9" @(, @("TC25-08", "負の時刻（##### 表示）", ""))
        $ws.Range("C9").Formula = "=TIME(0,0,1)-TIME(1,0,0)"
        $ws.Range("C9").NumberFormat = "hh:mm:ss"
        saveBook $wb "日付.xlsx"

        # 1904 年から数える日付形式のブック（同じシリアル値でも 4 年ずれる）
        $wb = newBook @("1904")
        $wb.Date1904 = $true
        $ws = $wb.Worksheets.Item(1)
        setRows $ws "A1" @(
            @("ID", "説明", "値"),
            @("TC25-09", "1904年形式のシリアル値 45383", "")
        )
        setCell $ws "C2" 45383 "yyyy/mm/dd"
        saveBook $wb "日付1904.xlsx"
    }

    # 設定ファイルの例（config\ にコピーして使う）
    Write-Host "[設定例]" -ForegroundColor Cyan
    $sjis = [System.Text.Encoding]::GetEncoding(932)
    writeText "$ConfigDir\変換対象フォルダパス.txt" "$OutDir`r`n"
    writeText "$ConfigDir\変換対象フォルダパス_前後空白と空行.txt" "`r`n   $OutDir   `r`n`r`n"
    writeText "$ConfigDir\変換対象フォルダパス_BOMなし.txt" $OutDir $noBom
    writeText "$ConfigDir\変換対象フォルダパス_ShiftJIS.txt" "$OutDir\ファイル名" $sjis  # 日本語部分が文字化けする
    writeText "$ConfigDir\変換対象フォルダパス_複数.txt" "$OutDir\2024`r`n$OutDir\2025`r`n"
    # 末尾のフォルダ名が同じ（インデックス名が「見積」と「見積(2)」になる）
    writeText "$ConfigDir\変換対象フォルダパス_同名フォルダ.txt" "$OutDir\2024\見積`r`n$OutDir\2025\見積`r`n"
    writeText "$ConfigDir\変換対象フォルダパス_チェックなし.txt" "$OutDir\2024`r`n# $OutDir\2025`r`n"
    writeText "$ConfigDir\変換対象フォルダパス_すべてチェックなし.txt" "# $OutDir`r`n"
    # 同じフォルダを " で囲んだもの・末尾に \ を付けたもので重ねて書く（最初の行だけ使われる）
    writeText "$ConfigDir\変換対象フォルダパス_重複.txt" "$OutDir\2024`r`n`"$OutDir\2024\`"`r`n$OutDir\2024\`r`n"
    writeText "$ConfigDir\変換対象フォルダパス_空.txt" "`r`n  `r`n"
    writeText "$ConfigDir\変換対象フォルダパス_存在しない.txt" "C:\存在しないフォルダ\excel`r`n$OutDir\2024`r`n"
    writeText "$ConfigDir\変換対象フォルダパス_サブフォルダ.txt" "$OutDir\2024`r`n"

    writeText "$ConfigDir\検索ワード.txt" ((@(
        "TC01",
        "山田",
        "\d{2,4}-\d{2,4}-\d{4}",
        "https?://",
        "12:34",
        "こんにちは",
        "1行目.2行目.3行目",
        "abc",
        "\(c\)\[d\]",
        "${yoshi}野家",
        $sushi,
        "非表示シート",
        "超非表示シート",
        "グラフシート",
        "コメント内",
        "テキストボックス内",
        "hidden-url",
        "VLOOKUP",
        "連結結果",
        "#DIV/0!",
        "令和6年4月1日",
        "1,234,567",
        "TC10-L",
        "TC10長文末尾",
        "TC11 大量データ最終行",
        "TC11 横長最終列",
        "TC07",
        "TC15",
        "TC16",
        "TC17",
        "TC20",
        "TC21",
        "TC22",
        "TC23",
        "TC24",
        "TC25",
        "TC26",
        "TC27",
        "TC28",
        "TC29",
        "該当なしの文字列XYZ"
    ) -join "`r`n") + "`r`n")
    writeText "$ConfigDir\検索ワード_前後空白と空行.txt" "`r`n  山田  `r`n`r`n`t佐藤`t`r`n`r`n"
    writeText "$ConfigDir\検索ワード_不正な正規表現.txt" "[閉じていない`r`n(`r`n*`r`n山田`r`n"
    writeText "$ConfigDir\検索ワード_空.txt" ""
    writeText "$ConfigDir\検索ワード_全件ヒット.txt" ".`r`n"
    writeText "$ConfigDir\検索ワード_高度な正規表現.txt" ((@(
        "(?-i)ABC",                  # 大文字・小文字を区別する
        "\bTC01\b",                  # 単語の区切り
        "^TC10-01",                  # 行の先頭
        '右隣$',                     # 行の末尾
        "TC10-L\d+",                 # 数字の繰り返し
        "山田(?= 太郎)",             # 先読み
        "(?<=TC01-)00[1-3]",         # 後読み
        "[\p{IsHiragana}]{6,}",      # ひらがなが6文字以上
        "[\p{IsCJKUnifiedIdeographs}]{4,}",  # 漢字が4文字以上
        "TC01-001`t山田",            # タブ区切り（列をまたぐ検索）
        "りんご|みかん",             # または
        "(?i)tc23-0[1-5]"            # 明示的に大文字・小文字を区別しない
    ) -join "`r`n") + "`r`n")
    $manyWords = New-Object System.Collections.Generic.List[string]
    for ($i = 1; $i -le 100; $i++) {
        $manyWords.Add("TC11-{0:D5}" -f $i)
    }
    writeText "$ConfigDir\検索ワード_100件.txt" (($manyWords -join "`r`n") + "`r`n")
    writeText "$ConfigDir\検索ワード_重複.txt" "山田`r`n山田`r`n  山田  `r`n佐藤`r`n"
} finally {
    releaseApp $excel
    $excel = $null
}

# ----------------------------------------------------------------------------
# Word
# ----------------------------------------------------------------------------

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0  # wdAlertsNone

try {
    runCase "TC20-1 Word 基本" {
        $doc = $word.Documents.Add()
        [void](addParagraph $doc "TC20-01 見出し1" $wdStyleHeading1)
        [void](addParagraph $doc "TC20-02 本文の段落です。担当は山田 太郎（営業部、03-1234-5678）です。")
        [void](addParagraph $doc "TC20-03 見出し2" $wdStyleHeading2)
        $first = addParagraph $doc "TC20-04 箇条書きの項目1"
        $last = addParagraph $doc "TC20-05 箇条書きの項目2"
        $doc.Range($first.Start, $last.End).ListFormat.ApplyBulletDefault()
        $first = addParagraph $doc "TC20-06 番号付きの項目1"
        $last = addParagraph $doc "TC20-07 番号付きの項目2"
        $doc.Range($first.Start, $last.End).ListFormat.ApplyNumberDefault()
        [void](addParagraph $doc "TC20-08 段落内改行の前$([char]11)段落内改行の後")
        [void](addParagraph $doc "TC20-09 タブの左`tタブの右")
        [void](addParagraph $doc "TC20-10 コロン URL: https://example.com:8080/path 時刻 12:34:56")
        [void](addParagraph $doc 'TC20-11 記号 "引用" a.b(c)[d]{e}*f+g?h^i$j|k\l')
        [void](addParagraph $doc "TC20-12 CP932外の文字 ${yoshi}野家 ${sushi} 한국어")
        [void](addParagraph $doc "TC20-13 全角　スペース ﾊﾝｶｸｶﾅ ①②㈱ ABC abc Ａｂｃ")
        # 手動の改ページ
        $range = $doc.Content
        $range.Collapse(0)
        $range.InsertBreak(7)  # wdPageBreak
        [void](addParagraph $doc "TC20-14 改ページ後（2ページ目）の段落")
        # 1 ページに収まらずに自然に改ページされる分量
        for ($i = 1; $i -le 60; $i++) {
            [void](addParagraph $doc ("TC20-15 ページからあふれる段落{0:D2}" -f $i))
        }
        [void](addParagraph $doc "TC20-16 最終段落")
        saveDoc $doc "Word\基本.docx"
    }

    runCase "TC20-2 Word 表" {
        $doc = $word.Documents.Add()
        [void](addParagraph $doc "TC20-T01 表の前の段落")
        $range = $doc.Content
        $range.Collapse(0)
        $table = $doc.Tables.Add($range, 5, 3)
        $table.Borders.Enable = $true
        $cells = @(
            @("ID", "氏名", "備考"),
            @("TC20-T02", "山田 太郎", "セル内の段落1`rセル内の段落2"),
            @("TC20-T03", "佐藤 花子", "セル内改行の前$([char]11)セル内改行の後")
        )
        for ($r = 0; $r -lt $cells.Count; $r++) {
            for ($c = 0; $c -lt 3; $c++) {
                $table.Cell($r + 1, $c + 1).Range.Text = $cells[$r][$c]
            }
        }
        $table.Cell(4, 1).Merge($table.Cell(4, 3))
        $table.Cell(4, 1).Range.Text = "TC20-T04 結合したセル"
        $table.Cell(5, 1).Range.Text = "TC20-T05 入れ子の表→"
        $inCell = $table.Cell(5, 2).Range
        $inCell.Collapse(1)  # wdCollapseStart
        $inner = $doc.Tables.Add($inCell, 1, 2)
        $inner.Borders.Enable = $true
        $inner.Cell(1, 1).Range.Text = "TC20-T06 入れ子のセルA"
        $inner.Cell(1, 2).Range.Text = "TC20-T07 入れ子のセルB"
        [void](addParagraph $doc "TC20-T08 表の後の段落")
        saveDoc $doc "Word\表.docx"
    }

    runCase "TC20-3 Word 本文以外の文字列" {
        $doc = $word.Documents.Add()
        $r = addParagraph $doc "TC20-O01 脚注・文末脚注・コメントを付けた段落"
        [void]$doc.Footnotes.Add($doc.Range($r.End, $r.End), $missing, "TC20 脚注のテキスト")
        [void]$doc.Endnotes.Add($doc.Range($r.End, $r.End), $missing, "TC20 文末脚注のテキスト")
        [void]$doc.Comments.Add($doc.Range($r.Start, $r.Start + 8), "TC20 コメントのテキスト")

        $r = addParagraph $doc "TC20-O02 ハイパーリンク→"
        [void]$doc.Hyperlinks.Add($doc.Range($r.End, $r.End), "https://example.com/TC20-hidden-url", $missing, $missing, "TC20 リンクの表示文字列")

        $r = addParagraph $doc "TC20-O03 フィールド→"
        $field = $doc.Fields.Add($doc.Range($r.End, $r.End), -1, 'QUOTE "TC20 フィールドの結果"', $false)  # wdFieldEmpty
        [void]$field.Update()

        $label = "TC20-O04 隠し文字→"
        $r = addParagraph $doc "${label}TC20 隠し文字のテキスト"
        $doc.Range($r.Start + $label.Length, $r.End).Font.Hidden = $true

        $r = addParagraph $doc "TC20-O05 ルビ→漢字"
        try {
            $doc.Range($r.End - 2, $r.End).PhoneticGuide("かんじ")
        } catch {
            Write-Host "  （ルビを設定できませんでした）" -ForegroundColor Yellow
        }

        $section = $doc.Sections.Item(1)
        $section.Headers.Item(1).Range.Text = "TC20 ヘッダーのテキスト"  # wdHeaderFooterPrimary
        $section.Footers.Item(1).Range.Text = "TC20 フッターのテキスト"

        $box = $doc.Shapes.AddTextbox(1, 100, 400, 300, 40)
        $box.TextFrame.TextRange.Text = "TC20 テキストボックスのテキスト"
        $shape = $doc.Shapes.AddShape(1, 100, 460, 300, 40)  # msoShapeRectangle
        $shape.TextFrame.TextRange.Text = "TC20 図形のテキスト"
        $shape.AlternativeText = "TC20 代替テキスト"
        saveDoc $doc "Word\本文以外.docx"
    }

    runCase "TC20-4 Word 変更履歴" {
        $doc = $word.Documents.Add()
        [void](addParagraph $doc "TC20-R01 変更履歴のある文書。TC20 削除された文字列。残る文字列。")
        $doc.TrackRevisions = $true
        $range = $doc.Content
        if ($range.Find.Execute("TC20 削除された文字列。")) {
            [void]$range.Delete()
        }
        [void](addParagraph $doc "TC20-R02 変更履歴で追加された段落")
        $doc.TrackRevisions = $false
        saveDoc $doc "Word\変更履歴.docx"
    }

    runCase "TC20-5 Word 長文・空の文書" {
        $doc = $word.Documents.Add()
        $lines = New-Object System.Collections.Generic.List[string]
        for ($i = 1; $i -le 3000; $i++) {
            $lines.Add(("TC20-L{0:D4} 長文の段落です。いろはにほへと ちりぬるを わかよたれそ つねならむ。" -f $i))
        }
        $lines.Add("TC20 長文の最終段落")
        $doc.Content.Text = $lines -join "`r"
        saveDoc $doc "Word\長文.docx"

        $doc = $word.Documents.Add()
        saveDoc $doc "Word\空文書.docx"
    }

    runCase "TC20-6 Word 形式" {
        simpleDoc "Word\形式\旧形式.doc" "TC20-F01" 0
        simpleDoc "Word\形式\マクロ有効.docm" "TC20-F02" 13
        simpleDoc "Word\形式\大文字拡張子.DOCX" "TC20-F03"
        simpleDoc "Word\形式\大文字拡張子旧形式.DOC" "TC20-F04" 0
        simpleDoc "Word\形式\リッチテキスト.rtf" "TC20-F05" 6
        simpleDoc "Word\形式\テンプレート.dotx" "TC20-F06" 14
        simpleDoc "Word\形式\OpenDocument.odt" "TC20-F07" 23
        simpleDoc "Word\形式\PDF.pdf" "TC20-F08" 17
        writeText "Word\形式\テキスト.txt" "TC20-F09`r`nテキストファイル"
    }

    runCase "TC20-7 Word 異常系" {
        simpleDoc "Word\異常系\読み取りパスワード付き.docx" "TC20-E01" 12 "openpw"
        simpleDoc "Word\異常系\読み取りパスワード付き旧形式.doc" "TC20-E02" 0 "openpw"
        $doc = $word.Documents.Add()
        [void](addParagraph $doc "TC20-E03 書き込みパスワード付き")
        saveDoc $doc "Word\異常系\書き込みパスワード付き.docx" 12 "" "writepw"
        writeText "Word\異常系\壊れたファイル.docx" "TC20-E04 これは Word ファイルではありません"
        writeText "Word\異常系\空ファイル.docx" "" $noBom
        writeLockFile "Word\異常系\~`$ロックファイル.docx"
        simpleDoc "Word\異常系\中身はdoc.docx" "TC20-E05" 0
        simpleDoc "Word\異常系\中身はdocx.doc" "TC20-E06" 12
        simpleDoc "Word\異常系\中身はRTF.doc" "TC20-E07" 6
        writeText "Word\異常系\中身はHTML.doc" "<html><head><meta charset=`"utf-8`"></head><body><p>TC20-E08 中身は HTML</p></body></html>"
    }


    runCase "TC27 Word 高度な文書" {
        $doc = $word.Documents.Add()

        # 目次（見出しから作るフィールド）
        [void](addParagraph $doc "TC27-01 目次" $wdStyleHeading1)
        $range = $doc.Content
        $range.Collapse(0)
        $toc = $doc.TablesOfContents.Add($range, $true, 1, 3)
        [void](addParagraph $doc "")

        # セクション 1: 見出しと本文、ブックマーク、相互参照
        [void](addParagraph $doc "TC27-02 第1章" $wdStyleHeading1)
        $r = addParagraph $doc "TC27-03 第1章の本文。ここにブックマークを付ける。"
        [void]$doc.Bookmarks.Add("TC27ブックマーク", $r)
        [void](addParagraph $doc "TC27-04 第1章の小見出し" $wdStyleHeading2)
        $r = addParagraph $doc "TC27-05 相互参照→"
        $field = $doc.Fields.Add($doc.Range($r.End, $r.End), -1, "REF TC27ブックマーク", $false)
        [void]$field.Update()

        # 多階層の箇条書き
        $first = addParagraph $doc "TC27-06 第1階層"
        $second = addParagraph $doc "TC27-07 第2階層"
        $third = addParagraph $doc "TC27-08 第3階層"
        try {
            # 多階層のリスト書式（アウトライン番号）を適用してから、段落ごとの階層を指定する
            $listTemplate = $word.ListGalleries.Item(3).ListTemplates.Item(1)  # wdOutlineNumberGallery
            $doc.Range($first.Start, $third.End).ListFormat.ApplyListTemplateWithLevel($listTemplate, $false, 0, 2)
            $doc.Range($second.Start, $second.End).ListFormat.ListLevelNumber = 2
            $doc.Range($third.Start, $third.End).ListFormat.ListLevelNumber = 3
        } catch {
            Write-Host "  （多階層の箇条書きにできませんでした: $($_.Exception.Message)）" -ForegroundColor Yellow
            $doc.Range($first.Start, $third.End).ListFormat.ApplyBulletDefault()
        }

        # 画像（代替テキスト付き）
        $r = addParagraph $doc "TC27-09 画像→"
        $pngPath = makePng "TC27 画像の中の文字（検索できない）"
        $picture = $doc.InlineShapes.AddPicture($pngPath, $false, $true, $doc.Range($r.End, $r.End))
        $picture.AlternativeText = "TC27 画像の代替テキスト"

        # コメントと、そのコメントへの返信
        $r = addParagraph $doc "TC27-10 コメントと返信を付けた段落"
        $comment = $doc.Comments.Add($r, "TC27 コメント本体のテキスト")
        try {
            [void]$comment.Replies.Add($comment.Range, "TC27 コメントへの返信のテキスト")
        } catch {
            Write-Host "  （コメントへの返信を追加できませんでした）" -ForegroundColor Yellow
        }

        # 書式だけの変更履歴（文字は変わらない）
        $r = addParagraph $doc "TC27-11 太字に変えた変更履歴のある段落"
        $doc.TrackRevisions = $true
        $r.Bold = $true
        $doc.TrackRevisions = $false

        # コンテンツコントロール（入力欄）
        $r = addParagraph $doc "TC27-12 コンテンツコントロール→"
        try {
            $control = $doc.ContentControls.Add(1, $doc.Range($r.End, $r.End))  # wdContentControlText
            $control.Range.Text = "TC27 コンテンツコントロールの中のテキスト"
        } catch {
            Write-Host "  （コンテンツコントロールを追加できませんでした）" -ForegroundColor Yellow
        }

        # 索引（索引項目のマークと索引）
        $r = addParagraph $doc "TC27-13 索引項目を付けた段落"
        try {
            [void]$doc.Indexes.MarkEntry($r, "TC27 索引の項目")
        } catch {
            Write-Host "  （索引項目を付けられませんでした）" -ForegroundColor Yellow
        }

        # 数式
        $r = addParagraph $doc "TC27-14 数式→ a^2 + b^2 = c^2"
        try {
            [void]$doc.OMaths.Add($doc.Range($r.Start + 14, $r.End))
            $doc.OMaths.Item(1).BuildUp()
        } catch {
            Write-Host "  （数式に変換できませんでした）" -ForegroundColor Yellow
        }

        # セクション 2: 横向き・2 段組み・独自のヘッダー
        $range = $doc.Content
        $range.Collapse(0)
        $range.InsertBreak(2)  # wdSectionBreakNextPage
        [void](addParagraph $doc "TC27-15 第2章（横向き・2段組み）" $wdStyleHeading1)
        for ($i = 1; $i -le 20; $i++) {
            [void](addParagraph $doc ("TC27-16 2段組みの段落{0:D2}。いろはにほへと ちりぬるを わかよたれそ。" -f $i))
        }
        $section = $doc.Sections.Item(2)
        $section.PageSetup.Orientation = 1  # wdOrientLandscape
        $section.PageSetup.TextColumns.SetCount(2)
        $header = $section.Headers.Item(1)
        $header.LinkToPrevious = $false
        $header.Range.Text = "TC27 第2章のヘッダー"
        $doc.Sections.Item(1).Headers.Item(1).Range.Text = "TC27 第1章のヘッダー"

        # 目次を更新してから保存する
        $toc.Update()
        saveDoc $doc "Word\高度.docx"
    }

    runCase "TC22 同名で種類違い（Word）" {
        simpleDoc "混在\同名.docx" "TC22-docx"
    }
} finally {
    releaseApp $word
    $word = $null
}

# ----------------------------------------------------------------------------
# PowerPoint
# ----------------------------------------------------------------------------

if (@(Get-Process powerpnt -ErrorAction SilentlyContinue).Count -gt 0) {
    # 起動中の PowerPoint に接続すると、開いているプレゼンテーションを閉じてしまうため作らない
    Write-Host "[TC21] PowerPoint が起動中のため作成しません。PowerPoint を終了してから再実行してください。" -ForegroundColor Red
    $failed.Add("TC21 PowerPoint（起動中のためスキップ）")
} else {
    $ppt = New-Object -ComObject PowerPoint.Application

    try {
        runCase "TC21-1 PowerPoint 基本" {
            $pres = $ppt.Presentations.Add(0)
            $slide = addSlide $pres $ppLayoutTitle "TC21-01 タイトルスライド"
            $slide.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "TC21-02 サブタイトル 山田 太郎"

            $slide = addSlide $pres $ppLayoutText "TC21-03 箇条書きのスライド"
            $body = $slide.Shapes.Placeholders.Item(2).TextFrame.TextRange
            $body.Text = "TC21-04 項目1`rTC21-05 項目2`rTC21-06 インデントした項目"
            $body.Paragraphs(3).IndentLevel = 2
            $slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "TC21 ノートのテキスト"

            $slide = addSlide $pres $ppLayoutText "TC21-07 特殊な文字"
            $slide.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = @(
                "TC21-08 段落内改行の前$([char]11)段落内改行の後",
                "TC21-09 コロン URL: https://example.com:8080/path 時刻 12:34:56",
                'TC21-10 記号 "引用" a.b(c)[d]{e}*f+g?h^i$j|k\l',
                "TC21-11 CP932外の文字 ${yoshi}野家 ${sushi} 한국어",
                "TC21-12 全角　スペース ﾊﾝｶｸｶﾅ ①②㈱ タブの左`tタブの右"
            ) -join "`r"

            # 作成順と表示順が違うスライド
            $slide = addSlide $pres $ppLayoutTitleOnly "TC21-13 最後に作って2枚目に移動したスライド"
            $slide.MoveTo(2)
            savePres $pres "PowerPoint\基本.pptx"

            # PowerPoint は保存時にスライドのファイル名（slide1.xml〜）を表示順に振り直すため、
            # presentation.xml を直接書き換えて、ファイル名の順と表示順が逆のプレゼンテーションを作る
            $relPath = "PowerPoint\並べ替え（XML編集）.pptx"
            $dest = Join-Path $OutDir $relPath
            [System.IO.File]::Copy((Join-Path $OutDir "PowerPoint\基本.pptx"), $dest)
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            $zip = [System.IO.Compression.ZipFile]::Open($dest, "Update")
            try {
                $entry = $zip.GetEntry("ppt/presentation.xml")
                $reader = New-Object System.IO.StreamReader($entry.Open())
                $xml = $reader.ReadToEnd()
                $reader.Close()
                $ids = @([regex]::Matches($xml, '<p:sldId [^>]*/>') | ForEach-Object { $_.Value })
                [array]::Reverse($ids)
                $xml = [regex]::Replace($xml, '(?<=<p:sldIdLst>).*?(?=</p:sldIdLst>)', ($ids -join ""))
                $entry.Delete()
                $writer = New-Object System.IO.StreamWriter($zip.CreateEntry("ppt/presentation.xml").Open(), $noBom)
                $writer.Write($xml)
                $writer.Close()
            } finally {
                $zip.Dispose()
            }
            Write-Host "  $relPath"
        }

        runCase "TC21-2 PowerPoint 本文以外の文字列" {
            $pres = $ppt.Presentations.Add(0)

            $slide = addSlide $pres $ppLayoutTitleOnly "TC21-O00 図形"
            [void](addTextBox $slide "TC21-O01 テキストボックス" 50 120)
            $shape = $slide.Shapes.AddShape(1, 50, 180, 400, 60)  # msoShapeRectangle
            $shape.TextFrame.TextRange.Text = "TC21-O02 図形内のテキスト"
            $shape.AlternativeText = "TC21 代替テキスト"
            $a = addTextBox $slide "TC21-O03 グループ内の図形A" 50 260 280
            $b = addTextBox $slide "TC21-O04 グループ内の図形B" 350 260 280
            [void]$slide.Shapes.Range([object[]]@($a.Name, $b.Name)).Group()
            $box = addTextBox $slide "TC21-O05 縦書きのテキスト" 650 100 40 300
            $box.TextFrame.Orientation = 4  # msoTextOrientationVerticalFarEast

            $slide = addSlide $pres $ppLayoutTitleOnly "TC21-T00 表"
            $table = $slide.Shapes.AddTable(3, 3, 50, 120, 600, 150).Table
            $cells = @(
                @("ID", "氏名", "備考"),
                @("TC21-T01", "山田 太郎", "セル内の段落1`rセル内の段落2")
            )
            for ($r = 0; $r -lt $cells.Count; $r++) {
                for ($c = 0; $c -lt 3; $c++) {
                    $table.Cell($r + 1, $c + 1).Shape.TextFrame.TextRange.Text = $cells[$r][$c]
                }
            }
            $table.Cell(3, 1).Merge($table.Cell(3, 3))
            $table.Cell(3, 1).Shape.TextFrame.TextRange.Text = "TC21-T02 結合したセル"

            $slide = addSlide $pres $ppLayoutTitleOnly "TC21-L00 リンク・フッターなど"
            $box = addTextBox $slide "TC21-O06 リンクの表示文字列" 50 120
            $box.TextFrame.TextRange.ActionSettings.Item(1).Hyperlink.Address = "https://example.com/TC21-hidden-url"
            try {
                $slide.HeadersFooters.Footer.Visible = -1
                $slide.HeadersFooters.Footer.Text = "TC21 フッターのテキスト"
            } catch {
                Write-Host "  （フッターを設定できませんでした）" -ForegroundColor Yellow
            }
            try {
                [void]$slide.Comments.Add(10, 10, "tester", "T", "TC21 コメントのテキスト")
            } catch {
                Write-Host "  （コメントを追加できませんでした）" -ForegroundColor Yellow
            }
            try {
                $smartArt = $slide.Shapes.AddSmartArt($ppt.SmartArtLayouts.Item(1), 50, 200, 500, 250)
                $smartArt.SmartArt.AllNodes.Item(1).TextFrame2.TextRange.Text = "TC21 SmartArtのテキスト"
            } catch {
                Write-Host "  （SmartArt を追加できませんでした）" -ForegroundColor Yellow
            }

            # 何も入力していないプレースホルダー（「クリックしてタイトルを入力」は文字列ではない）
            [void](addSlide $pres $ppLayoutText)
            savePres $pres "PowerPoint\本文以外.pptx"
        }

        runCase "TC21-3 PowerPoint 非表示スライド・セクション" {
            $pres = $ppt.Presentations.Add(0)
            [void](addSlide $pres $ppLayoutTitleOnly "TC21-S01 表示スライド")
            $slide = addSlide $pres $ppLayoutTitleOnly "TC21-S02 非表示スライド"
            $slide.SlideShowTransition.Hidden = -1  # msoTrue
            [void](addSlide $pres $ppLayoutTitleOnly "TC21-S03 2つ目のセクションのスライド")
            [void]$pres.SectionProperties.AddBeforeSlide(1, "TC21 セクション名1")
            [void]$pres.SectionProperties.AddBeforeSlide(3, "TC21 セクション名2")
            savePres $pres "PowerPoint\非表示スライド.pptx"
        }

        runCase "TC21-4 PowerPoint 多数スライド・空" {
            $pres = $ppt.Presentations.Add(0)
            for ($i = 1; $i -le 100; $i++) {
                [void](addSlide $pres $ppLayoutTitleOnly ("TC21-M{0:D3} スライド{0}" -f $i))
            }
            $pres.Slides.Item(100).Shapes.Title.TextFrame.TextRange.Text = "TC21 最終スライド"
            savePres $pres "PowerPoint\多数スライド.pptx"

            $pres = $ppt.Presentations.Add(0)
            savePres $pres "PowerPoint\空.pptx"
        }

        runCase "TC21-5 PowerPoint 形式" {
            simplePres "PowerPoint\形式\旧形式.ppt" "TC21-F01" 1
            simplePres "PowerPoint\形式\マクロ有効.pptm" "TC21-F02" 25
            simplePres "PowerPoint\形式\大文字拡張子.PPTX" "TC21-F03"
            simplePres "PowerPoint\形式\大文字拡張子旧形式.PPT" "TC21-F04" 1
            simplePres "PowerPoint\形式\スライドショー.ppsx" "TC21-F05" 28
            simplePres "PowerPoint\形式\旧形式スライドショー.pps" "TC21-F06" 7
            simplePres "PowerPoint\形式\テンプレート.potx" "TC21-F07" 26
            simplePres "PowerPoint\形式\OpenDocument.odp" "TC21-F08" 35
            simplePres "PowerPoint\形式\PDF.pdf" "TC21-F09" 32
        }

        runCase "TC21-6 PowerPoint 異常系" {
            simplePres "PowerPoint\異常系\読み取りパスワード付き.pptx" "TC21-E01" 24 "openpw"
            simplePres "PowerPoint\異常系\読み取りパスワード付き旧形式.ppt" "TC21-E02" 1 "openpw"
            writeText "PowerPoint\異常系\壊れたファイル.pptx" "TC21-E03 これは PowerPoint ファイルではありません"
            writeText "PowerPoint\異常系\空ファイル.pptx" "" $noBom
            writeLockFile "PowerPoint\異常系\~`$ロックファイル.pptx"
            simplePres "PowerPoint\異常系\中身はppt.pptx" "TC21-E04" 1
            simplePres "PowerPoint\異常系\中身はpptx.ppt" "TC21-E05" 24
        }


        runCase "TC28 PowerPoint 高度なプレゼンテーション" {
            $pres = $ppt.Presentations.Add(0)
            $pres.PageSetup.SlideSize = 1  # ppSlideSizeOnScreen（4:3）

            # 多階層の箇条書き
            $slide = addSlide $pres $ppLayoutText "TC28-01 多階層の箇条書き"
            $body = $slide.Shapes.Placeholders.Item(2).TextFrame.TextRange
            $body.Text = "TC28-02 第1階層`rTC28-03 第2階層`rTC28-04 第3階層`rTC28-05 第4階層`rTC28-06 第5階層"
            for ($i = 2; $i -le 5; $i++) {
                $body.Paragraphs($i).IndentLevel = $i
            }
            # ノート（複数段落・長文）
            $notes = $slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange
            $notesText = New-Object System.Collections.Generic.List[string]
            $notesText.Add("TC28 ノートの1段落目")
            $notesText.Add("TC28 ノートの2段落目")
            for ($i = 1; $i -le 30; $i++) {
                $notesText.Add(("TC28 ノートの長い段落{0:D2}。いろはにほへと ちりぬるを わかよたれそ。" -f $i))
            }
            $notesText.Add("TC28 ノートの最終段落")
            $notes.Text = $notesText -join "`r"

            # WordArt・画像・日付・スライド番号
            $slide = addSlide $pres $ppLayoutTitleOnly "TC28-07 WordArt と画像"
            try {
                [void]$slide.Shapes.AddTextEffect(0, "TC28 WordArtのテキスト", "Arial", 28, 0, 0, 50, 120)
            } catch {
                Write-Host "  （WordArt を追加できませんでした）" -ForegroundColor Yellow
            }
            $pngPath = makePng "TC28 画像の中の文字（検索できない）"
            $picture = $slide.Shapes.AddPicture($pngPath, 0, -1, 50, 220, 400, 120)
            $picture.AlternativeText = "TC28 画像の代替テキスト"
            try {
                $slide.HeadersFooters.DateAndTime.Visible = -1
                $slide.HeadersFooters.DateAndTime.UseFormat = 0
                $slide.HeadersFooters.DateAndTime.Text = "TC28 日付プレースホルダのテキスト"
                $slide.HeadersFooters.SlideNumber.Visible = -1
            } catch {
                Write-Host "  （日付・スライド番号を設定できませんでした）" -ForegroundColor Yellow
            }

            # グラフ（Excel のデータを埋め込む）
            $slide = addSlide $pres $ppLayoutTitleOnly "TC28-08 グラフ"
            try {
                $chartShape = $slide.Shapes.AddChart2(-1, 51, 50, 120, 500, 300)  # 51 = 集合縦棒
                $chart = $chartShape.Chart
                $chart.HasTitle = $true
                $chart.ChartTitle.Text = "TC28 グラフのタイトル"
                $chart.ChartData.Activate()
                $chartBook = $chart.ChartData.Workbook
                $chartSheet = $chartBook.Worksheets.Item(1)
                $chartSheet.Range("A2").Value2 = "TC28 グラフの項目名"
                $chartBook.Application.Quit()
            } catch {
                Write-Host "  （グラフを追加できませんでした: $($_.Exception.Message)）" -ForegroundColor Yellow
            }

            # 非表示スライドにもノートを付ける
            $slide = addSlide $pres $ppLayoutTitleOnly "TC28-09 非表示スライド（ノート付き）"
            $slide.SlideShowTransition.Hidden = -1
            $slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "TC28 非表示スライドのノート"

            # スライドマスターとレイアウトに置いた文字（スライドには表示されない）。スライドを作った後に置く
            try {
                $master = $pres.SlideMaster
                $box = $master.Shapes.AddTextbox(1, 20, 20, 400, 30)
                $box.TextFrame.TextRange.Text = "TC28 スライドマスターのテキスト"
                $layout = $pres.Slides.Item(1).CustomLayout
                $box = $layout.Shapes.AddTextbox(1, 20, 60, 400, 30)
                $box.TextFrame.TextRange.Text = "TC28 レイアウトのテキスト"
            } catch {
                Write-Host "  （マスター・レイアウトに文字を置けませんでした）" -ForegroundColor Yellow
            }

            savePres $pres "PowerPoint\高度.pptx"
        }

        runCase "TC22 同名で種類違い（PowerPoint）" {
            simplePres "混在\同名.pptx" "TC22-pptx"
        }
    } finally {
        releaseApp $ppt
        $ppt = $null
    }
}

if (Test-Path -LiteralPath $stageDir) {
    Remove-Item -LiteralPath $stageDir -Recurse -Force
}

Write-Host ""
if ($failed.Count -gt 0) {
    Write-Host "作成に失敗したケース: $($failed -join ', ')" -ForegroundColor Red
    exit 1
}
Write-Host "テストデータを作成しました: $OutDir"
