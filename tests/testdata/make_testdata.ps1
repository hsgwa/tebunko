# win_grep の動作確認用テストデータ（Excel ファイル群と設定ファイル例）を生成する。
# Excel（COM）が必要。出力フォルダは削除して作り直す。
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\testdata\make_testdata.ps1
# ケースの一覧と期待結果は同じフォルダの README.md を参照。
param (
    [string]$OutDir = "$PSScriptRoot\excel",
    [string]$ConfigDir = "$PSScriptRoot\設定例",
    [int]$BigRows = 20000
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

# CP932 で表せない文字
$yoshi = [char]::ConvertFromUtf32(0x20BB7)  # 𠮷（つちよし）
$sushi = [char]::ConvertFromUtf32(0x1F363)  # 🍣

$failed = New-Object System.Collections.Generic.List[string]

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
        foreach ($wb in @($excel.Workbooks)) {
            $wb.Close($false)
        }
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
            "衝突$([char]0x201D)"  # ” は toSafeFileName 後の " と同じ文字なのでファイル名が衝突する（Excel のシート名は全角半角を区別しないため ＜ と < の組は作れない）
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

        # Excel がブックを開いているときに作るロックファイルに似せたもの
        $lock = New-Object byte[] 165
        $owner = [System.Text.Encoding]::ASCII.GetBytes("tester")
        $lock[0] = $owner.Length
        [Array]::Copy($owner, 0, $lock, 1, $owner.Length)
        foreach ($name in @("~`$ロックファイル.xlsx", "~`$ロックファイル_隠し属性なし.xlsx")) {
            $path = Join-Path $OutDir "異常系\$name"
            [System.IO.File]::WriteAllBytes($path, $lock)
            Write-Host "  異常系\$name"
        }
        [System.IO.File]::SetAttributes((Join-Path $OutDir "異常系\~`$ロックファイル.xlsx"), "Hidden")

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

    # 設定ファイルの例（config\ にコピーして使う）
    Write-Host "[設定例]" -ForegroundColor Cyan
    $noBom = New-Object System.Text.UTF8Encoding($false)
    $sjis = [System.Text.Encoding]::GetEncoding(932)
    writeText "$ConfigDir\変換対象フォルダパス.txt" "$OutDir`r`n"
    writeText "$ConfigDir\変換対象フォルダパス_前後空白と空行.txt" "`r`n   $OutDir   `r`n`r`n"
    writeText "$ConfigDir\変換対象フォルダパス_BOMなし.txt" $OutDir $noBom
    writeText "$ConfigDir\変換対象フォルダパス_ShiftJIS.txt" $OutDir $sjis
    writeText "$ConfigDir\変換対象フォルダパス_2行.txt" "$OutDir`r`n$OutDir\2024`r`n"
    writeText "$ConfigDir\変換対象フォルダパス_空.txt" "`r`n  `r`n"
    writeText "$ConfigDir\変換対象フォルダパス_存在しない.txt" "C:\存在しないフォルダ\excel`r`n"
    writeText "$ConfigDir\変換対象フォルダパス_サブフォルダ.txt" "$OutDir\2024`r`n"

    $rootDir = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    writeText "$ConfigDir\検索対象インデックスパス_複数.txt" "$rootDir\work\index\2024`r`n$rootDir\work\index\2025`r`n"
    writeText "$ConfigDir\検索対象インデックスパス_存在しないフォルダを含む.txt" "$rootDir\work\index\2024`r`nC:\存在しないフォルダ\index`r`n"
    writeText "$ConfigDir\検索対象インデックスパス_空.txt" ""

    writeText "$ConfigDir\検索ワード.txt" ((@(
        "TC01",
        "山田",
        "\d{2,4}-\d{2,4}-\d{4}",
        "https?://",
        "12:34",
        "こんにちは",
        "1行目2行目3行目",
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
        "該当なしの文字列XYZ"
    ) -join "`r`n") + "`r`n")
    writeText "$ConfigDir\検索ワード_前後空白と空行.txt" "`r`n  山田  `r`n`r`n`t佐藤`t`r`n`r`n"
    writeText "$ConfigDir\検索ワード_不正な正規表現.txt" "[閉じていない`r`n(`r`n*`r`n山田`r`n"
    writeText "$ConfigDir\検索ワード_空.txt" ""
    writeText "$ConfigDir\検索ワード_全件ヒット.txt" ".`r`n"
} finally {
    foreach ($wb in @($excel.Workbooks)) {
        $wb.Close($false)
    }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $stageDir) {
        Remove-Item -LiteralPath $stageDir -Recurse -Force
    }
}

Write-Host ""
if ($failed.Count -gt 0) {
    Write-Host "作成に失敗したケース: $($failed -join ', ')" -ForegroundColor Red
    exit 1
}
Write-Host "テストデータを作成しました: $OutDir"
