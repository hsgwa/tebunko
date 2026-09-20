# 画面（WPF）
#
# ［1 インデックス管理］［2 検索］［9 プロセス停止］の3タブ。画面の定義は config_gui.xaml。
# 変換は office_to_tsv.ps1 をウィンドウを出さずに起動して進み具合を表示し、検索・プロセス停止は画面内で行う（処理は common.ps1 と共通）。

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

# zip 展開で付く Mark-of-the-Web（外部由来の印）を、このフォルダから消す。印が残っていると
# RemoteSigned で common.ps1 などの読み込みがブロックされるため。通常は win_grep.bat が起動前に消すが、
# ショートカットから直接起動したときや、あとでファイルを差し替えたときのために、ここでも消しておく。
# （この config_gui.ps1 自身が印付きだと、この行に来る前にブロックされる。その場合は win_grep.bat から起動する）
try {
    Get-ChildItem -LiteralPath $PSScriptRoot -Recurse -File -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue
} catch { }

. "$PSScriptRoot\common.ps1"

$ErrorActionPreference = "Stop"

${appTitle}    = "win_grep"
${appId}       = "win_grep"  # タスクバーのボタン・ショートカットを結び付ける ID（AppUserModelID）
${searchLimit} = 10000
${previewLines} = 3  # 選択行のプレビューに出す前後の行数
${commonPath}  = "$PSScriptRoot\common.ps1"

trap {
    [System.Windows.MessageBox]::Show("予期しないエラーが発生しました。`n$($_.Exception.Message)", ${appTitle}, "OK", "Error") | Out-Null
    exit 1
}

# ---- 多重起動の防止（ツールの配置フォルダごと） ----
#
# すでに開いているときは、その画面のウィンドウを前面に出して終わる（もう一度起動するのは、
# たいてい「開いたつもりのウィンドウが他のウィンドウの裏にある」ときのため）。
# 知らせるのは名前付きイベントで行う。ここは C# の型をコンパイルする前のため、.NET の機能だけを使う。

$md5 = New-Object System.Security.Cryptography.MD5CryptoServiceProvider
$instanceKey = [BitConverter]::ToString($md5.ComputeHash([System.Text.Encoding]::UTF8.GetBytes(${rootDir}.ToLowerInvariant()))).Replace("-", "")
$mutexName = "Local\win_grep_gui_" + $instanceKey
$activateName = "Local\win_grep_gui_activate_" + $instanceKey
$createdNew = $false
$mutex = New-Object System.Threading.Mutex($true, $mutexName, [ref]$createdNew)
if (!$createdNew) {
    $running = $null
    if ([System.Threading.EventWaitHandle]::TryOpenExisting($activateName, [ref]$running)) {
        [void]$running.Set()
        $running.Close()
        exit
    }
    # 以前の版の画面が開いている等で知らせられないときだけ、メッセージを出す
    [System.Windows.MessageBox]::Show("すでに開いています。", ${appTitle}, "OK", "Information") | Out-Null
    exit
}
$activateEvent = New-Object System.Threading.EventWaitHandle($false, [System.Threading.EventResetMode]::AutoReset, $activateName)

# ---- 画面で使う型（PowerShell class。実行時コンパイル（csc.exe）を出さないため C# から移した） ----
# INotifyPropertyChanged は NotifyBase を継承して実装する（PS class はプロパティのセッターに
# ロジックを書けないため、値を変える側が Set*/Raise を呼ぶ）。HitRow は件数が多いので生成時は生データだけ持ち、
# 表示用（強調セグメント・セル解析）は画面に見えた行だけ Prepare() で作る（LoadingRow から呼ぶ）。

class NotifyBase : System.ComponentModel.INotifyPropertyChanged {
    hidden [System.ComponentModel.PropertyChangedEventHandler] $handler
    [void] add_PropertyChanged([System.ComponentModel.PropertyChangedEventHandler]$h) { $this.handler = [Delegate]::Combine($this.handler, $h) }
    [void] remove_PropertyChanged([System.ComponentModel.PropertyChangedEventHandler]$h) { $this.handler = [Delegate]::Remove($this.handler, $h) }
    [void] Raise([string]$name) { if ($this.handler) { $this.handler.Invoke($this, (New-Object System.ComponentModel.PropertyChangedEventArgs $name)) } }
}

class Segment {
    [string]$Text
    [bool]$IsHit
}

# 選択行のプレビューの列。見出しと各行のセルが同じものを参照し、幅を変えると列全体に反映する。
# セルは Width に OneWay バインド。幅の変更は SetWidth（見出しのドラッグから呼ぶ）で行う
class PreviewColumn : NotifyBase {
    static [double] $MinWidth = 24   # ドラッグで狭くできる下限
    [string]$Label
    [double]$Width
    [void] SetWidth([double]$value) {
        $newWidth = [Math]::Max([PreviewColumn]::MinWidth, $value)
        if ($this.Width -eq $newWidth) { return }
        $this.Width = $newWidth
        $this.Raise("Width")
    }
}

# 選択行のプレビューのセル1つ（IsSelected はコピーのために選んだ範囲）
class PreviewCell : NotifyBase {
    [string]$Text        # セルの値そのまま（コピーに使う）
    [string]$Display     # 画面に出す文字列（長すぎるセルは切り詰める）
    [string]$ToolTip
    [PreviewColumn]$Column
    [bool]$IsHit
    [int]$RowIndex
    [int]$ColumnIndex
    [bool]$IsSelected
    [void] SetSelected([bool]$value) {
        if ($this.IsSelected -eq $value) { return }
        $this.IsSelected = $value
        $this.Raise("IsSelected")
    }
}

# 選択行のプレビューの1行
class PreviewRow {
    [string]$Number
    [bool]$IsHitRow
    [System.Collections.Generic.List[PreviewCell]]$Cells
}

# 選択行のプレビュー。HitOffset・HitWidth は選択行で最初に一致したセルの左端と幅（横スクロール用）
class PreviewTable {
    [System.Collections.Generic.List[PreviewColumn]]$Columns
    [System.Collections.Generic.List[PreviewRow]]$Rows
    [double]$HitOffset
    [double]$HitWidth
    [int]$TotalColumns
    [int]$ShownColumns
    [string]$RangeLabel

    hidden [int]$anchorRow = -1
    hidden [int]$anchorColumn = -1
    hidden [int]$focusRow = -1
    hidden [int]$focusColumn = -1

    [bool] HasSelection() { return $this.anchorRow -ge 0 }

    [void] Select([PreviewCell]$cell, [bool]$extend) {
        if ($null -eq $cell) { return }
        if (-not $extend -or $this.anchorRow -lt 0) {
            $this.anchorRow = $cell.RowIndex
            $this.anchorColumn = $cell.ColumnIndex
        }
        $this.focusRow = $cell.RowIndex
        $this.focusColumn = $cell.ColumnIndex
        $this.ApplySelection()
    }

    [void] SelectRow([PreviewCell]$cell) {
        if ($null -eq $cell -or $this.Columns.Count -eq 0) { return }
        $this.anchorRow = $cell.RowIndex
        $this.focusRow = $cell.RowIndex
        $this.anchorColumn = 0
        $this.focusColumn = $this.Columns.Count - 1
        $this.ApplySelection()
    }

    [void] ClearSelection() {
        $this.anchorRow = -1; $this.anchorColumn = -1; $this.focusRow = -1; $this.focusColumn = -1
        $this.ApplySelection()
    }

    hidden [void] ApplySelection() {
        $rowFrom = [Math]::Min($this.anchorRow, $this.focusRow); $rowTo = [Math]::Max($this.anchorRow, $this.focusRow)
        $columnFrom = [Math]::Min($this.anchorColumn, $this.focusColumn); $columnTo = [Math]::Max($this.anchorColumn, $this.focusColumn)
        foreach ($row in $this.Rows) {
            foreach ($cell in $row.Cells) {
                $cell.SetSelected($this.anchorRow -ge 0 -and
                    $cell.RowIndex -ge $rowFrom -and $cell.RowIndex -le $rowTo -and
                    $cell.ColumnIndex -ge $columnFrom -and $cell.ColumnIndex -le $columnTo)
            }
        }
    }

    [string] GetSelectionText() {
        if ($this.anchorRow -lt 0) { return "" }
        $rowFrom = [Math]::Min($this.anchorRow, $this.focusRow); $rowTo = [Math]::Max($this.anchorRow, $this.focusRow)
        $columnFrom = [Math]::Min($this.anchorColumn, $this.focusColumn); $columnTo = [Math]::Max($this.anchorColumn, $this.focusColumn)
        if ($rowFrom -eq $rowTo -and $columnFrom -eq $columnTo) {
            return $this.CellText($rowFrom, $columnFrom)
        }
        $text = New-Object System.Text.StringBuilder
        for ($r = $rowFrom; $r -le $rowTo; $r++) {
            if ($r -gt $rowFrom) { [void]$text.Append("`r`n") }
            for ($c = $columnFrom; $c -le $columnTo; $c++) {
                if ($c -gt $columnFrom) { [void]$text.Append("`t") }
                [void]$text.Append([PreviewTable]::QuoteForExcel($this.CellText($r, $c)))
            }
        }
        return $text.ToString()
    }

    [int] SelectedCount() {
        if ($this.anchorRow -lt 0) { return 0 }
        return ([Math]::Abs($this.anchorRow - $this.focusRow) + 1) * ([Math]::Abs($this.anchorColumn - $this.focusColumn) + 1)
    }

    hidden [string] CellText([int]$rowIndex, [int]$columnIndex) {
        if ($rowIndex -lt 0 -or $rowIndex -ge $this.Rows.Count) { return "" }
        $cells = $this.Rows[$rowIndex].Cells
        if ($columnIndex -ge 0 -and $columnIndex -lt $cells.Count) { return $cells[$columnIndex].Text }
        return ""
    }

    static [string] QuoteForExcel([string]$text) {
        if ($text.IndexOf("`n") -lt 0 -and $text.IndexOf("`t") -lt 0 -and $text.IndexOf('"') -lt 0) { return $text }
        return '"' + $text.Replace('"', '""') + '"'
    }
}

# 検索結果の1行。生成時は生データのみ。表示用（DisplayLine・Segments・CellText）は Prepare() で作る（可視行だけ）。
# 表示用の各項目は Prepare() 後に PropertyChanged を出す（LoadingRow より前にバインドされても更新されるように）
class HitRow : NotifyBase {
    static [regex] $CellRegex  = [regex]::new("\t(?:`"(?:[^`"]|`"`")*`"[^\t]*|[^\t]*)")
    static [regex] $QuoteRegex = [regex]::new("^`"((?:[^`"]|`"`")*)`"(.*)`$", [System.Text.RegularExpressions.RegexOptions]::Singleline)
    static [regex] $ExcelRegex = [regex]::new("\.xls[a-z]?`$", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    static [char] $CellNewLine = [char]0x2028   # TSV のセル内改行（common.ps1 の cellNewLine）
    static [int] $LeadLength = 40
    static [int] $MaxDisplay = 600
    static [double] $NumberWidth = 44
    static [double] $MinCellWidth = 48
    static [double] $MaxCellWidth = 260
    static [double] $MaxParagraphWidth = 640
    static [int] $MaxPreviewColumns = 200
    static [int] $MaxCellChars = 300
    static [int] $MaxToolTipChars = 1000
    static [int] $MaxWidthChars = 100

    [string]$IndexName
    [string]$Root
    [string]$RelPath
    [string]$RelDir
    [string]$FileName
    [string]$Book
    [string]$Location
    [int]$LineNumber
    [string]$Line
    [bool]$IsExcel
    [string]$MatchCell
    [string]$CellText
    [string]$DisplayLine
    [System.Collections.Generic.List[Segment]]$Segments
    [bool]$Prepared

    hidden [string]$word
    hidden [regex]$pattern

    # 生成（検索ヒットごと。生データの代入のみ＝軽い）
    static [HitRow] Create([string]$indexName, [string]$root, [string]$relPath, [string]$relDir, [string]$fileName,
                           [string]$book, [string]$location, [int]$lineNumber, [string]$line, [string]$word, [regex]$pattern) {
        $row = [HitRow]::new()
        $row.IndexName = $indexName
        $row.Root = $root
        $row.RelPath = $relPath
        $row.RelDir = $(if ($null -eq $relDir) { "" } else { $relDir })
        $row.FileName = $fileName
        $row.Book = $book
        $row.Location = $location
        $row.LineNumber = $lineNumber
        $row.Line = $(if ($null -eq $line) { "" } else { $line })
        $row.word = $word
        $row.pattern = $pattern
        $row.IsExcel = [HitRow]::ExcelRegex.IsMatch($(if ($null -eq $book) { "" } else { $book }))
        return $row
    }

    # 表示用（強調セグメント・DisplayLine・セル列）を作る。可視行になったときに1回だけ呼ぶ（LoadingRow）
    [void] Prepare() {
        if ($this.Prepared) { return }
        $this.Prepared = $true
        $this.DisplayLine = [HitRow]::ToDisplay($this.Line)
        $this.Segments = $this.BuildSegments()
        $this.SetMatchCell()
        $this.Raise("DisplayLine"); $this.Raise("Segments"); $this.Raise("CellText"); $this.Raise("MatchCell")
    }

    hidden [void] SetMatchCell() {
        $this.MatchCell = ""
        $this.CellText = ""
        if (-not $this.IsExcel) { return }
        $cells = [HitRow]::SplitCells($this.Line, $true)
        $count = 0
        for ($i = 0; $i -lt $cells.Count; $i++) {
            if ([HitRow]::FindMatches($cells[$i], $this.word, $this.pattern).Count -eq 0) { continue }
            if ($count -eq 0) { $this.MatchCell = [HitRow]::ColumnName($i + 1) + $this.LineNumber }
            $count++
        }
        $this.CellText = $(if ($count -gt 1) { $this.MatchCell + " ほか " + ($count - 1) } else { $this.MatchCell })
    }

    # 絞り込み。生データ（相対フォルダ・元ファイル名・場所・行番号・生の行）での部分一致（大文字小文字を区別しない）。
    # ※以前は表示用（セル番地・タブ表示）も対象にしていたが、遅延生成のため生データのみを対象にした。
    [bool] Contains([string]$text) {
        $ci = [System.StringComparison]::CurrentCultureIgnoreCase
        if ($this.RelDir.IndexOf($text, $ci) -ge 0) { return $true }
        if (("" + $this.Book).IndexOf($text, $ci) -ge 0) { return $true }
        if (("" + $this.Location).IndexOf($text, $ci) -ge 0) { return $true }
        if ($this.LineNumber.ToString().IndexOf($text, $ci) -ge 0) { return $true }
        if ($this.Line.IndexOf($text, $ci) -ge 0) { return $true }
        return $false
    }

    static [string] ToDisplay([string]$text) {
        if ($null -eq $text) { return "" }
        return $text.Replace("`t", " │ ").Replace([HitRow]::CellNewLine, [char]0x21b5)
    }

    static [System.Collections.Generic.List[int[]]] FindMatches([string]$text, [string]$word, [regex]$pattern) {
        $list = New-Object System.Collections.Generic.List[int[]]
        if ([string]::IsNullOrEmpty($text)) { return $list }
        if ($null -ne $pattern) {
            foreach ($m in $pattern.Matches($text)) { if ($m.Length -gt 0) { $list.Add(@($m.Index, $m.Length)) } }
        } elseif (-not [string]::IsNullOrEmpty($word)) {
            $i = 0
            while ($i -lt $text.Length) {
                $i = $text.IndexOf($word, $i, [System.StringComparison]::CurrentCultureIgnoreCase)
                if ($i -lt 0) { break }
                $list.Add(@($i, $word.Length)); $i += [Math]::Max(1, $word.Length)
            }
        }
        return $list
    }

    hidden [System.Collections.Generic.List[Segment]] BuildSegments() {
        $segs = New-Object System.Collections.Generic.List[Segment]
        $pos = 0; $shown = 0
        foreach ($m in [HitRow]::FindMatches($this.Line, $this.word, $this.pattern)) {
            if ($m[0] -lt $pos) { continue }
            if ($m[0] + $m[1] -gt $this.Line.Length) { break }
            $before = $this.Line.Substring($pos, $m[0] - $pos)
            if ($segs.Count -eq 0 -and $before.Length -gt [HitRow]::LeadLength) {
                $before = [char]0x2026 + $before.Substring($before.Length - [HitRow]::LeadLength)
            }
            if ($before.Length -gt 0) { $segs.Add([Segment]@{ Text = [HitRow]::ToDisplay($before); IsHit = $false }) }
            $segs.Add([Segment]@{ Text = [HitRow]::ToDisplay($this.Line.Substring($m[0], $m[1])); IsHit = $true })
            $shown += $before.Length + $m[1]
            $pos = $m[0] + $m[1]
            if ($shown -gt [HitRow]::MaxDisplay) { break }
        }
        if ($pos -lt $this.Line.Length) {
            $rest = $this.Line.Substring($pos)
            if ($rest.Length -gt [HitRow]::MaxDisplay) { $rest = $rest.Substring(0, [HitRow]::MaxDisplay) + [char]0x2026 }
            $segs.Add([Segment]@{ Text = [HitRow]::ToDisplay($rest); IsHit = $false })
        }
        return $segs
    }

    static [System.Collections.Generic.List[string]] SplitCells([string]$line, [bool]$isExcel) {
        $cells = New-Object System.Collections.Generic.List[string]
        $src = $(if ($null -eq $line) { "" } else { $line })
        if (-not $isExcel) { $cells.AddRange($src.Split([char]9)); return $cells }
        foreach ($m in [HitRow]::CellRegex.Matches("`t" + $src)) {
            $cell = $m.Value.Substring(1)
            $q = [HitRow]::QuoteRegex.Match($cell)
            if ($q.Success) { $cell = $q.Groups[1].Value.Replace('""', '"') + $q.Groups[2].Value }
            $cells.Add($cell)
        }
        return $cells
    }

    # 選択行のプレビュー（前後の行をセルに分けた表）を作る。numbers・lines は readTsvContext の結果
    [PreviewTable] BuildPreview([int[]]$numbers, [string[]]$lines) {
        if ($null -eq $numbers -or $null -eq $lines -or $numbers.Length -eq 0 -or $numbers.Length -ne $lines.Length) {
            $numbers = @($this.LineNumber)
            $lines = @($this.Line)
        }
        $rowCells = New-Object 'System.Collections.Generic.List[System.Collections.Generic.List[string]]'
        $columnCount = 0
        foreach ($ln in $lines) {
            $cells = [HitRow]::SplitCells($ln, $this.IsExcel)
            $rowCells.Add($cells)
            $columnCount = [Math]::Max($columnCount, $cells.Count)
        }

        $hitColumn = -1
        for ($r = 0; $r -lt $rowCells.Count -and $hitColumn -lt 0; $r++) {
            if ($numbers[$r] -ne $this.LineNumber) { continue }
            for ($c = 0; $c -lt $rowCells[$r].Count; $c++) {
                if ([HitRow]::FindMatches($rowCells[$r][$c], $this.word, $this.pattern).Count -gt 0) { $hitColumn = $c; break }
            }
        }
        $firstColumn = 0
        $shownColumns = $columnCount
        if ($columnCount -gt [HitRow]::MaxPreviewColumns) {
            $shownColumns = [HitRow]::MaxPreviewColumns
            if ($hitColumn -ge 0) {
                $firstColumn = [Math]::Max(0, [Math]::Min($hitColumn - [int]([HitRow]::MaxPreviewColumns / 2), $columnCount - [HitRow]::MaxPreviewColumns))
            }
        }

        $maxWidth = $(if ($this.IsExcel) { [HitRow]::MaxCellWidth } else { [HitRow]::MaxParagraphWidth })
        $table = [PreviewTable]::new()
        $table.Columns = New-Object System.Collections.Generic.List[PreviewColumn]
        $table.Rows = New-Object System.Collections.Generic.List[PreviewRow]
        $table.HitOffset = -1
        $table.HitWidth = 0
        $table.TotalColumns = $columnCount
        $table.ShownColumns = $shownColumns
        $table.RangeLabel = $(if ($shownColumns -gt 0) { $this.ColumnLabel($firstColumn + 1) + "〜" + $this.ColumnLabel($firstColumn + $shownColumns) } else { "" })
        for ($i = 0; $i -lt $shownColumns; $i++) {
            $c = $firstColumn + $i
            $width = [HitRow]::TextWidth($this.ColumnLabel($c + 1))
            foreach ($cells in $rowCells) {
                if ($c -lt $cells.Count) { $width = [Math]::Max($width, [HitRow]::CellWidth($cells[$c])) }
            }
            $col = [PreviewColumn]::new()
            $col.Label = $this.ColumnLabel($c + 1)
            $col.Width = [Math]::Min($maxWidth, [Math]::Max([HitRow]::MinCellWidth, $width))
            $table.Columns.Add($col)
        }

        for ($r = 0; $r -lt $rowCells.Count; $r++) {
            $isHitRow = $numbers[$r] -eq $this.LineNumber
            $row = [PreviewRow]::new()
            $row.Number = $numbers[$r].ToString()
            $row.IsHitRow = $isHitRow
            $row.Cells = New-Object System.Collections.Generic.List[PreviewCell]
            $left = [HitRow]::NumberWidth
            for ($i = 0; $i -lt $shownColumns; $i++) {
                $c = $firstColumn + $i
                $cell = $(if ($c -lt $rowCells[$r].Count) { $rowCells[$r][$c] } else { "" })
                $isHit = [HitRow]::FindMatches($cell, $this.word, $this.pattern).Count -gt 0
                $text = $cell.Replace([HitRow]::CellNewLine, "`n")
                $pc = [PreviewCell]::new()
                $pc.Text = $text
                $pc.Display = [HitRow]::Shorten($text, [HitRow]::MaxCellChars)
                $pc.ToolTip = $(if ($text.Length -gt 0) { [HitRow]::Shorten($text, [HitRow]::MaxToolTipChars) } else { $null })
                $pc.Column = $table.Columns[$i]
                $pc.IsHit = $isHit
                $pc.RowIndex = $r
                $pc.ColumnIndex = $i
                $row.Cells.Add($pc)
                if ($isHitRow -and $isHit -and $table.HitOffset -lt 0) {
                    $table.HitOffset = $left
                    $table.HitWidth = $table.Columns[$i].Width
                }
                $left += $table.Columns[$i].Width
            }
            $table.Rows.Add($row)
        }
        if ($table.HitOffset -lt 0) { $table.HitOffset = 0 }
        return $table
    }

    static [double] CellWidth([string]$cell) {
        $width = 0.0
        foreach ($line in $(if ($null -eq $cell) { "" } else { $cell }).Split(@([HitRow]::CellNewLine, "`n"))) {
            $seg = $(if ($line.Length -gt [HitRow]::MaxWidthChars) { $line.Substring(0, [HitRow]::MaxWidthChars) } else { $line })
            $width = [Math]::Max($width, [HitRow]::TextWidth($seg))
            if ($width -ge [HitRow]::MaxParagraphWidth) { break }
        }
        return $width
    }

    static [string] Shorten([string]$text, [int]$max) {
        if ($null -eq $text -or $text.Length -le $max) { return $text }
        return $text.Substring(0, $max) + [char]0x2026
    }

    hidden [string] ColumnLabel([int]$number) {
        return $(if ($this.IsExcel) { [HitRow]::ColumnName($number) } else { $number.ToString() })
    }

    static [double] TextWidth([string]$text) {
        $width = 14.0
        foreach ($ch in $(if ($null -eq $text) { "" } else { $text }).ToCharArray()) {
            $code = [int]$ch
            $width += $(if ($code -lt 0x0100 -or ($code -ge 0xFF61 -and $code -le 0xFF9F)) { 7 } else { 12 })
        }
        return $width
    }

    static [string] ColumnName([int]$number) {
        $name = ""
        while ($number -gt 0) {
            $number--
            $name = [char]([int][char]'A' + $number % 26) + $name
            $number = [int]($number / 26)
        }
        return $name
    }
}

# ［9 プロセス停止］の1行
class ProcRow {
    [int]$Id
    [string]$AppName
    [bool]$Background
    [string]$StateText
    [string]$StartText
    [string]$MemoryText
    [string]$TitleText
}

# ［1 インデックス管理］の変換に失敗したファイル1件
class FailRow {
    [string]$RelPath
    [string]$Reason
    [string]$ConvertedText
    [string]$SourcePath
}

# ［1 インデックス管理］のインデックス一覧 1 件。プログラムから変えたときに画面へ反映するため通知する。
# ［変換］チェックの TwoWay バインドは値の往復に使い、保存はチェックボックスの Click で行う（PS class はセッターにロジックを書けないため）
class FolderItem : NotifyBase {
    [string]$Name          # インデックス名（work\index 直下のフォルダ名）
    [string]$Path
    [bool]$Enabled
    [string]$StatusText
    [object]$StatusBrush
    [string]$FileCountText
    [string]$FileCountToolTip
    [string]$LastConvertedText

    [void] SetEnabled([bool]$value) { if ($this.Enabled -ne $value) { $this.Enabled = $value; $this.Raise("Enabled") } }
    [void] SetName([string]$value) { if ($this.Name -ne $value) { $this.Name = $value; $this.Raise("Name") } }
    [void] SetPath([string]$value) { if ($this.Path -ne $value) { $this.Path = $value; $this.Raise("Path") } }
    [void] SetStatus([string]$text, [object]$brush) { $this.StatusText = $text; $this.StatusBrush = $brush; $this.Raise("StatusText"); $this.Raise("StatusBrush") }
    [void] SetStats([string]$countText, [string]$toolTip, [string]$lastConverted) {
        $this.FileCountText = $countText; $this.FileCountToolTip = $toolTip; $this.LastConvertedText = $lastConverted
        $this.Raise("FileCountText"); $this.Raise("FileCountToolTip"); $this.Raise("LastConvertedText")
    }
}

# 検索対象の1件（common.ps1 の getIndexTsvFiles に渡す）
class SearchTarget {
    [string]$Root
    [string]$RelPath
    [bool]$Recurse
}

# 検索対象から外したフォルダ（common.ps1 の readSearchExcludes / writeSearchExcludes と同じ項目）
class SearchExclude {
    [string]$Path
    [bool]$Subfolders
}

# 検索対象インデックスのツリーの1項目。3状態チェック（true/false/null）は子・親の状態から決まる。
# チェックはツリーのチェックボックスの Click で Toggle() を呼んで変える（TwoWay バインドはしない）
class IndexNode : NotifyBase {
    [string]$Name
    [string]$Root
    [string]$RelPath
    [bool]$IsFiles
    [bool]$IsPlaceholder
    [bool]$Exists
    [string]$SourcePath
    [string]$ToolTip
    [IndexNode]$Parent
    [System.Collections.ObjectModel.ObservableCollection[IndexNode]]$Children

    [Nullable[bool]]$IsChecked = $true   # チェックボックスは OneWay バインド。変更は Toggle/SetChecked で行う
    [bool]$IsExpanded                    # TreeViewItem.IsExpanded は OneWay。展開は Expanded イベント/SetExpanded で読み込む
    hidden [bool]$loaded

    IndexNode([IndexNode]$parent, [string]$name, [string]$root, [string]$relPath, [bool]$isFiles) {
        $this.Parent = $parent
        $this.Name = $name
        $this.Root = $root
        $this.RelPath = $relPath
        $this.IsFiles = $isFiles
        $this.Exists = $true
        $this.Children = New-Object System.Collections.ObjectModel.ObservableCollection[IndexNode]
        if ($null -ne $parent) { $this.IsChecked = ($parent.IsChecked -ne $false) }
    }

    static [IndexNode] CreateRoot([string]$root, [string]$name, [string]$relPath, [string]$sourcePath) {
        $node = [IndexNode]::new($null, $name, $root, $relPath, $false)
        $node.SourcePath = $sourcePath
        $dir = $node.FullPath()
        $node.Exists = [System.IO.Directory]::Exists([IndexNode]::LongPath($dir))
        $node.ToolTip = $(if ($null -ne $sourcePath) { "元のフォルダ：" + $sourcePath + "`nインデックス：" + $dir } else { $dir })
        if (-not $node.Exists) { $node.ToolTip += "`n（フォルダが見つかりません。検索時はスキップします）" }
        if ($node.Exists -and [IndexNode]::HasSubfolders($dir)) { $node.Children.Add([IndexNode]::NewPlaceholder($node)) }
        return $node
    }

    [string] FullPath() {
        return $(if ($this.RelPath -eq "") { $this.Root } else { $this.Root.TrimEnd('\') + "\" + $this.RelPath })
    }

    # 展開する（プログラムから・Expanded イベントから）。子を1回だけ読み込む
    [void] SetExpanded([bool]$value) {
        if ($value) { $this.LoadChildren() }
        if ($this.IsExpanded -eq $value) { return }
        $this.IsExpanded = $value
        $this.Raise("IsExpanded")
    }

    [void] SetChecked([bool]$value) {
        $this.SetTree($value)
        if ($null -ne $this.Parent) { $this.Parent.UpdateFromChildren() }
    }

    [void] Toggle() { $this.SetChecked($this.IsChecked -ne $true) }

    hidden [void] SetTree([bool]$value) {
        $this.SetState($value)
        foreach ($child in $this.Children) { if (-not $child.IsPlaceholder) { $child.SetTree($value) } }
    }

    hidden [void] SetState([Nullable[bool]]$value) {
        if ($this.IsChecked -eq $value) { return }
        $this.IsChecked = $value
        $this.Raise("IsChecked")
    }

    hidden [void] UpdateFromChildren() {
        $any = $false
        [Nullable[bool]]$state = $null
        foreach ($child in $this.Children) {
            if ($child.IsPlaceholder) { continue }
            if (-not $any) { $state = $child.IsChecked; $any = $true }
            elseif ($state -ne $child.IsChecked) { $state = $null; break }
        }
        if (-not $any) { return }
        $this.SetState($state)
        if ($null -ne $this.Parent) { $this.Parent.UpdateFromChildren() }
    }

    [void] LoadChildren() {
        if ($this.loaded -or $this.IsFiles -or $this.IsPlaceholder) { return }
        $this.loaded = $true
        $this.Children.Clear()
        if (-not $this.Exists) { return }
        $dir = $this.FullPath()
        $names = New-Object System.Collections.Generic.List[string]
        try {
            foreach ($sub in [System.IO.Directory]::EnumerateDirectories([IndexNode]::LongPath($dir))) {
                $n = [System.IO.Path]::GetFileName($sub)
                if ([IndexNode]::IsBookDir($n)) { continue }
                $names.Add($n)
            }
        } catch {
        }
        $names.Sort([System.StringComparer]::CurrentCultureIgnoreCase)
        if ($names.Count -gt 0 -and [IndexNode]::HasFiles($dir)) {
            $files = [IndexNode]::new($this, "（このフォルダ直下のファイル）", $this.Root, $this.RelPath, $true)
            $files.ToolTip = "サブフォルダを除く、" + $(if ($null -ne $this.SourcePath) { $this.SourcePath } else { $dir }) + " の直下のファイル"
            $this.Children.Add($files)
        }
        foreach ($n in $names) {
            $childRel = $(if ($this.RelPath -eq "") { $n } else { $this.RelPath + "\" + $n })
            $child = [IndexNode]::new($this, $n, $this.Root, $childRel, $false)
            if ($null -ne $this.SourcePath) { $child.SourcePath = $this.SourcePath.TrimEnd('\') + "\" + $n }
            $child.ToolTip = $(if ($null -ne $child.SourcePath) { "元のフォルダ：" + $child.SourcePath } else { $child.FullPath() })
            if ([IndexNode]::HasSubfolders($child.FullPath())) { $child.Children.Add([IndexNode]::NewPlaceholder($child)) }
            $this.Children.Add($child)
        }
    }

    [IndexNode] Find([string]$path) {
        if ($this.IsFiles -or $this.IsPlaceholder) { return $null }
        $full = $this.FullPath().TrimEnd('\')
        $path = $path.TrimEnd('\')
        if ([string]::Equals($path, $full, [System.StringComparison]::OrdinalIgnoreCase)) { return $this }
        if (-not $path.StartsWith($full + "\", [System.StringComparison]::OrdinalIgnoreCase)) { return $null }
        $this.LoadChildren()
        foreach ($child in $this.Children) {
            $found = $child.Find($path)
            if ($null -ne $found) { return $found }
        }
        return $null
    }

    [void] ApplyExclude([string]$path, [bool]$subfolders) {
        $node = $this.Find($path)
        if ($null -eq $node) { return }
        if (-not $subfolders) {
            $node.LoadChildren()
            foreach ($child in $node.Children) {
                if ($child.IsFiles) { $child.SetChecked($false); return }
            }
            if ($node.Children.Count -gt 0) { return }
        }
        $node.SetChecked($false)
    }

    [void] AddTargets([System.Collections.Generic.List[SearchTarget]]$targets) {
        if ($this.IsPlaceholder -or $this.IsChecked -eq $false) { return }
        if ($this.IsChecked -eq $true) {
            $targets.Add([SearchTarget]@{ Root = $this.Root; RelPath = $this.RelPath; Recurse = (-not $this.IsFiles) })
            return
        }
        foreach ($child in $this.Children) { $child.AddTargets($targets) }
    }

    [void] AddExcludes([System.Collections.Generic.List[SearchExclude]]$excludes) {
        if ($this.IsPlaceholder -or $this.IsChecked -eq $true) { return }
        if ($this.IsChecked -eq $false) {
            $excludes.Add([SearchExclude]@{ Path = $this.FullPath(); Subfolders = (-not $this.IsFiles) })
            return
        }
        foreach ($child in $this.Children) { $child.AddExcludes($excludes) }
    }

    [void] AddExpanded([System.Collections.Generic.List[string]]$paths) {
        if ($this.IsFiles -or $this.IsPlaceholder) { return }
        if ($this.IsExpanded) { $paths.Add($this.FullPath()) }
        foreach ($child in $this.Children) { $child.AddExpanded($paths) }
    }

    static [IndexNode] NewPlaceholder([IndexNode]$parent) {
        $node = [IndexNode]::new($parent, "読み込み中…", $parent.Root, $parent.RelPath, $false)
        $node.IsPlaceholder = $true
        return $node
    }

    static [string] LongPath([string]$path) {
        if ($path.StartsWith("\\?\")) { return $path }
        if ($path.EndsWith(":")) { $path += "\" }
        if ($path.StartsWith("\\")) { return "\\?\UNC\" + $path.Substring(2) }
        return "\\?\" + $path
    }

    static [bool] IsBookDir([string]$name) {
        $ext = [System.IO.Path]::GetExtension($name).ToLowerInvariant()
        if ($ext.Length -lt 4 -or $ext.Length -gt 5) { return $false }
        return $ext.StartsWith(".xls") -or $ext.StartsWith(".doc") -or $ext.StartsWith(".ppt")
    }

    static [bool] HasSubfolders([string]$dir) {
        try {
            foreach ($sub in [System.IO.Directory]::EnumerateDirectories([IndexNode]::LongPath($dir))) {
                if (-not [IndexNode]::IsBookDir([System.IO.Path]::GetFileName($sub))) { return $true }
            }
            return $false
        } catch { return $false }
    }

    static [bool] HasFiles([string]$dir) {
        try {
            foreach ($f in [System.IO.Directory]::EnumerateFiles([IndexNode]::LongPath($dir), "*.tsv")) { return $true }
            foreach ($sub in [System.IO.Directory]::EnumerateDirectories([IndexNode]::LongPath($dir))) {
                if (-not [IndexNode]::IsBookDir([System.IO.Path]::GetFileName($sub))) { continue }
                foreach ($f in [System.IO.Directory]::EnumerateFiles($sub, "*.tsv")) { return $true }
            }
            return $false
        } catch { return $false }
    }
}

# ---- フォルダ選択ダイアログ（エクスプローラー風）で使う型 ----

# 左のツリーの1項目。データだけを持ち、中身の読み込みは loadFolderNode（common.ps1 の getFolderEntries を呼ぶ）で行う。
# 展開（IsExpanded）・選択（IsSelected）は TreeViewItem と TwoWay バインドし、
# 展開したときの読み込みは TreeView の Expanded イベントで駆動する（PS class はセッターにロジックを書けないため）
class FolderNode : NotifyBase {
    [string]$Name
    [string]$Path                 # 開くフォルダ。見出し（「よく使う場所」「PC」）は ""
    [string]$ToolTip
    [bool]$IsHeader               # 見出し（フォルダではないので、選んでも移動しない）
    [bool]$IsPlaceholder          # 「読み込み中…」（▷ を出すためだけの子）
    [bool]$Loaded                 # 子を読み込み済みか
    [FolderNode]$Parent
    [System.Collections.ObjectModel.ObservableCollection[FolderNode]]$Children

    [bool]$IsExpanded
    [bool]$IsSelected

    FolderNode([FolderNode]$parent, [string]$name, [string]$path) {
        $this.Parent = $parent
        $this.Name = $name
        $this.Path = $path
        $this.ToolTip = $path
        $this.Children = New-Object System.Collections.ObjectModel.ObservableCollection[FolderNode]
    }

    [void] SetExpanded([bool]$value) {
        if ($this.IsExpanded -eq $value) { return }
        $this.IsExpanded = $value
        $this.Raise("IsExpanded")
    }

    [void] SetSelected([bool]$value) {
        if ($this.IsSelected -eq $value) { return }
        $this.IsSelected = $value
        $this.Raise("IsSelected")
    }

    [void] AddPlaceholder() {
        $node = [FolderNode]::new($this, "読み込み中…", "")
        $node.IsPlaceholder = $true
        $this.Children.Add($node)
    }

    # スクリーンリーダー・自動化ツールにはフォルダ名で見えるようにする（既定では型名になる）
    [string] ToString() { return $this.Name }
}

# 右の一覧の1行。フォルダは選べ、ファイルは「目的のフォルダかどうか」を確かめるために出すだけ（選べない）
class FolderEntry {
    [string]$Name
    [string]$Path
    [bool]$IsFolder
    [bool]$IsOffice
    [string]$Kind
    [string]$UpdatedText

    # スクリーンリーダー・自動化ツールには名前で見えるようにする（既定では型名になる）
    [string] ToString() { return $this.Name }
}

# ---- アイコン ----

${iconFile} = "$PSScriptRoot\win_grep.ico"  # タイトルバーとタスクバーに出すアイコン
# アイコンは Window.Icon（loadWindow）でタイトルバー・タスクバーに出る。
# ※以前は SetAppId（P/Invoke）でタスクバーのボタンを PowerShell と分けていたが、
#   実行時コンパイル（csc.exe）を無くすため廃止した（アイコン自体は Window.Icon で出るため残る）。

function loadWindow {
    param (
        [string]$path
    )

    [xml]$xaml = [System.IO.File]::ReadAllText($path)
    $loaded = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))

    # アイコンは XAML に書かず、ここで読み込む（XamlReader.Load は XAML 内の相対パスを解決できないため）。
    # ファイルを掴んだままにしないよう OnLoad で読み切る。アイコンが無くても画面は開けるようにする。
    if (Test-Path ${iconFile}) {
        $loaded.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create(
            (New-Object Uri ${iconFile}),
            [System.Windows.Media.Imaging.BitmapCreateOptions]::None,
            [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
    }

    return $loaded
}

$window = loadWindow "$PSScriptRoot\config_gui.xaml"
$ui = @{}
foreach ($name in @(
        "Tabs", "IndexTab", "SearchTab", "KillTab", "IndexTabHeader", "KillTabHeader", "StatusText", "CloseButton",
        "IndexGrid", "IndexGridPlaceholder", "NewIndexButton", "EditIndexButton", "RebuildIndexButton", "RemoveIndexButton",
        "IndexSummaryText", "ConversionStateText", "ConvertButton", "ConvertHint",
        "FailedPanel", "FailedHeading", "FailedGrid",
        "ConvertProgressPanel", "ConvertProgressText", "ConvertProgressEta", "ConvertProgress", "ConvertProgressDetail", "ConvertStopButton", "ConvertLogButton",
        "WordBox", "SearchButton", "RegexCheck", "CaseCheck", "FileFilterBox", "FileFilterPlaceholder", "WordNotice", "SearchTargetText", "GoIndexTabButton",
        "IndexTree", "IndexTreePlaceholder", "CheckAllIndexButton", "UncheckAllIndexButton",
        "SummaryText", "SearchProgress", "FilterBox", "FilterPlaceholder", "ResultGrid", "IndexColumn",
        "MenuOpen", "MenuOpenReadOnly", "MenuOpenNew", "MenuOpenFolder", "MenuCopy", "MenuCopyPath", "DetailPanel", "DetailTitle", "OpenButton", "OpenModeCombo", "OpenFolderButton", "PreviewScroll", "PreviewHeader", "PreviewRows", "PreviewNote", "MenuPreviewCopy", "MenuPreviewCopyRow", "ExportButton",
        "ProcessGrid", "ProcessSummaryText", "RefreshProcessButton", "KillAllButton", "KillSelectedButton", "KillBackgroundButton")) {
    $ui[$name] = $window.FindName($name)
}
$taskbar = $window.TaskbarItemInfo

function toBrush {
    param (
        [string]$hex
    )

    return New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString($hex))
}

${okBrush}   = toBrush "#2E8B57"
${warnBrush} = toBrush "#B45309"
${ngBrush}   = toBrush "#DC2626"

# ---- 共通の部品 ----

function setStatus {
    param (
        [string]$text
    )

    $ui.StatusText.Text = $text
    $ui.StatusText.ToolTip = $text
}

function showMessage {
    param (
        [string]$message,
        [string]$buttons = "OK",
        [string]$icon = "Information",
        [string]$default = "None",
        [System.Windows.Window]$owner = $window  # ダイアログを開いているときは、そのダイアログを親にする
    )

    return [System.Windows.MessageBox]::Show($owner, $message, ${appTitle}, $buttons, $icon, $default)
}

function safe {
    # イベント処理で例外が起きても画面を落とさず、内容を表示する
    param (
        [scriptblock]$block
    )

    try {
        & $block
    } catch {
        setStatus "エラーが発生しました：$($_.Exception.Message)"
        showMessage "エラーが発生しました。`n$($_.Exception.Message)" "OK" "Error" | Out-Null
    }
}

function newTimer {
    param (
        [int]$milliseconds,
        [scriptblock]$onTick
    )

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds($milliseconds)
    $timer.Add_Tick($onTick)
    return $timer
}

function selectFolder {
    # エクスプローラー風のフォルダ選択ダイアログ（config_gui_folder_select.xaml）を開き、選んだフォルダを返す（キャンセルなら $null）。
    #   ・左：よく使う場所（デスクトップ・ドキュメント・ダウンロード）と PC のドライブのツリー
    #   ・右：今のフォルダの中身（フォルダと、そのフォルダにあるファイル。Office ファイルは色を変える）
    #   ・上：アドレスバー（パスの入力・貼り付けで移動）と［←］［→］［↑］
    #   ・下：選ぶフォルダのパス（一覧でフォルダを選ぶ・ドラッグ＆ドロップでも入る）
    # ※Windows 標準のフォルダ選択（WinForms の FolderBrowserDialog）はツリーだけでファイルが見えず、
    #   目的のフォルダにたどり着きにくいため、画面として作る。
    #   エクスプローラー形式の COM ダイアログ（IFileOpenDialog）は実行時コンパイル（csc.exe）が要るため使わない（12.2）
    param (
        [string]$description,
        [string]$initialPath,
        [System.Windows.Window]$owner = $window
    )

    $dialog = loadWindow "$PSScriptRoot\config_gui_folder_select.xaml"
    if ($owner) {
        $dialog.Owner = $owner
    }
    $ctrl = @{}
    foreach ($name in @(
            "DescriptionText", "BackButton", "ForwardButton", "UpButton", "AddressBox",
            "FolderTree", "EntryList", "EntryPlaceholder", "StatusText", "FolderBox", "OkButton", "ErrorText")) {
        $ctrl[$name] = $dialog.FindName($name)
    }
    $script:folderSelect = @{
        Window  = $dialog
        Ctrl    = $ctrl
        Current = ""                                                     # 今開いているフォルダ
        History = (New-Object System.Collections.Generic.List[string])   # ［←］［→］でたどる履歴
        Index   = -1                                                     # 履歴の今の位置
        All     = @()                                                    # 今のフォルダの中身（絞り込み前）
        Roots   = (New-Object 'System.Collections.ObjectModel.ObservableCollection[object]')
        Syncing = $false                                                 # ツリーの選択を合わせている間（移動を起こさない）
    }
    $ctrl.DescriptionText.Text = $description
    $ctrl.FolderTree.ItemsSource = $script:folderSelect.Roots
    loadFolderTreeRoots

    # ---- 操作 ----
    $ctrl.BackButton.Add_Click({ safe { moveFolderHistory -1 } })
    $ctrl.ForwardButton.Add_Click({ safe { moveFolderHistory 1 } })
    $ctrl.UpButton.Add_Click({ safe { goParentFolder } })
    $ctrl.AddressBox.Add_PreviewKeyDown({
        param ($sender, $e)
        if ($e.Key -eq "Return") {
            # Enter は［選択］（既定のボタン）ではなく、入力したパスへの移動にする
            safe { goFolder $script:folderSelect.Ctrl.AddressBox.Text | Out-Null }
            $e.Handled = $true
        }
    })
    $ctrl.EntryList.Add_SelectionChanged({ safe { onFolderEntrySelected } })
    $ctrl.EntryList.Add_MouseDoubleClick({ safe { openSelectedFolderEntry } })
    $ctrl.EntryList.Add_PreviewKeyDown({
        param ($sender, $e)
        if ($e.Key -eq "Return") {
            safe { openSelectedFolderEntry }
            $e.Handled = $true
        }
    })
    $ctrl.FolderTree.Add_SelectedItemChanged({
        safe {
            $d = $script:folderSelect
            $node = $d.Ctrl.FolderTree.SelectedItem
            if ($d.Syncing -or $null -eq $node -or $node.Path -eq "") {
                return
            }
            if (-not (testSamePath $node.Path $d.Current)) {
                goFolder $node.Path | Out-Null
            }
        }
    })
    # 展開したときに、そのフォルダのサブフォルダを読み込む（読み込みはイベントで駆動する。12 章）
    $ctrl.FolderTree.AddHandler([System.Windows.Controls.TreeViewItem]::ExpandedEvent, [System.Windows.RoutedEventHandler] {
        param ($s, $e)
        safe {
            $node = $e.OriginalSource.DataContext
            if ($node -is [FolderNode]) {
                loadFolderNode $node
            }
        }
    })
    # 選んだ項目が画面の外にあるとき（アドレスバーからの移動など）に見えるようにする
    $ctrl.FolderTree.AddHandler([System.Windows.Controls.TreeViewItem]::SelectedEvent, [System.Windows.RoutedEventHandler] {
        param ($s, $e)
        if ($e.OriginalSource -is [System.Windows.Controls.TreeViewItem]) {
            $e.OriginalSource.BringIntoView()
        }
    })
    $ctrl.FolderBox.Add_PreviewDragOver({ onFolderDragOver @args })
    $ctrl.FolderBox.Add_PreviewDrop({
        param ($sender, $e)
        safe {
            $folders = @(getDroppedFolders $e)
            if ($folders.Count -gt 0) {
                goFolder $folders[0] | Out-Null
            }
        }
        $e.Handled = $true
    })
    $ctrl.OkButton.Add_Click({
        safe {
            $d = $script:folderSelect
            $path = normalizeFolderPath $d.Ctrl.FolderBox.Text
            if ($path -eq "") {
                setFolderSelectError "フォルダを選んでください。"
                return
            }
            if (-not (Test-Path -LiteralPath (toLongPath $path) -PathType Container)) {
                setFolderSelectError "「${path}」は見つかりません。一覧から選ぶか、パスを確かめてください。"
                return
            }
            $d.Window.DialogResult = $true
        }
    })
    # キーボード操作はエクスプローラーに合わせる
    #   Alt+← / Alt+→：戻る・進む、Alt+↑ / BackSpace：1 つ上へ、F5 / Ctrl+R：読み直す、
    #   F4 / Alt+D / Ctrl+L：アドレスバーへ、Esc：キャンセル（IsCancel）
    $dialog.Add_PreviewKeyDown({
        param ($sender, $e)
        $key = if ($e.Key -eq "System") { $e.SystemKey } else { $e.Key }
        $modifiers = $e.KeyboardDevice.Modifiers
        $alt = (($modifiers -band [System.Windows.Input.ModifierKeys]::Alt) -ne 0)
        $ctrl = (($modifiers -band [System.Windows.Input.ModifierKeys]::Control) -ne 0)
        $inTextBox = ($e.OriginalSource -is [System.Windows.Controls.TextBox])
        if ($key -eq "F5" -or ($ctrl -and $key -eq "R")) {
            safe { reloadFolder }
            $e.Handled = $true
        } elseif ($alt -and $key -eq "Left") {
            safe { moveFolderHistory -1 }
            $e.Handled = $true
        } elseif ($alt -and $key -eq "Right") {
            safe { moveFolderHistory 1 }
            $e.Handled = $true
        } elseif (($alt -and $key -eq "Up") -or ($key -eq "Back" -and -not $inTextBox)) {
            # BackSpace は入力欄では文字を消すため、入力欄以外のときだけ 1 つ上へ
            safe { goParentFolder }
            $e.Handled = $true
        } elseif ($key -eq "F4" -or ($alt -and $key -eq "D") -or ($ctrl -and $key -eq "L")) {
            safe { focusAddressBox }
            $e.Handled = $true
        }
    })
    # マウスの戻る・進むボタン（サイドボタン）でも履歴をたどる
    $dialog.Add_PreviewMouseDown({
        param ($sender, $e)
        if ($e.ChangedButton -eq [System.Windows.Input.MouseButton]::XButton1) {
            safe { moveFolderHistory -1 }
            $e.Handled = $true
        } elseif ($e.ChangedButton -eq [System.Windows.Input.MouseButton]::XButton2) {
            safe { moveFolderHistory 1 }
            $e.Handled = $true
        }
    })

    # ---- 最初に開くフォルダ ----
    $start = normalizeFolderPath $initialPath
    if ($start -ne "" -and -not (Test-Path -LiteralPath (toLongPath $start) -PathType Container)) {
        # 指定のフォルダが無ければ、その上の、今もあるフォルダを開く
        $start = getExistingFolder $start
    }
    $opened = $false
    foreach ($candidate in @($start) + @(getQuickFolders | ForEach-Object { $_.Path }) + @(getComputerFolders | ForEach-Object { $_.Path })) {
        if ($candidate -eq "") {
            continue
        }
        if (goFolder $candidate) {
            $opened = $true
            break
        }
    }
    if (-not $opened) {
        setFolderSelectError "開けるフォルダが見つかりません。上の欄にフォルダのパスを入力してください。"
        updateFolderSelectButtons
    }
    $ctrl.EntryList.Focus() | Out-Null

    $result = $null
    if ($dialog.ShowDialog()) {
        $result = normalizeFolderPath $ctrl.FolderBox.Text
    }
    $script:folderSelect = $null
    return $result
}

function testSamePath {
    # フォルダ選択ダイアログの中で、2つのパスが同じ書き方かを見る（末尾の \ ・大文字と小文字の違いは無視する）
    param (
        [string]$a,
        [string]$b
    )

    return [string]::Equals(([string]$a).TrimEnd("\"), ([string]$b).TrimEnd("\"), [System.StringComparison]::OrdinalIgnoreCase)
}

function setFolderSelectError {
    # フォルダ選択ダイアログの下に出す、直してほしい内容（空なら消す）
    param (
        [string]$message
    )

    $ctrl = $script:folderSelect.Ctrl
    $ctrl.ErrorText.Text = $message
    $ctrl.ErrorText.Visibility = if ($message -eq "") { "Collapsed" } else { "Visible" }
}

function loadFolderTreeRoots {
    # ツリーの一番上（「よく使う場所」「PC」）を作る。どちらも最初から開いておく
    $d = $script:folderSelect
    $d.Roots.Clear()

    $quick = [FolderNode]::new($null, "よく使う場所", "")
    $quick.IsHeader = $true
    $quick.Loaded = $true
    foreach ($place in @(getQuickFolders)) {
        addFolderTreeChild $quick $place.Name $place.Path | Out-Null
    }
    if ($quick.Children.Count -gt 0) {
        $quick.SetExpanded($true)
        $d.Roots.Add($quick)
    }

    $computer = [FolderNode]::new($null, "PC", "")
    $computer.IsHeader = $true
    $computer.Loaded = $true
    foreach ($drive in @(getComputerFolders)) {
        addFolderTreeChild $computer $drive.Name $drive.Path | Out-Null
    }
    $computer.SetExpanded($true)
    $d.Roots.Add($computer)
}

function addFolderTreeChild {
    # ツリーに子（フォルダ）を1つ足す。サブフォルダがあれば ▷ を出すための仮の子を入れておく
    param (
        [FolderNode]$parent,
        [string]$name,
        [string]$path
    )

    $node = [FolderNode]::new($parent, $name, $path)
    if (testHasSubFolders $path) {
        $node.AddPlaceholder()
    }
    $parent.Children.Add($node)
    return $node
}

function loadFolderNode {
    # ツリーのノードのサブフォルダを読み込む（1回だけ）
    param (
        [FolderNode]$node
    )

    if ($null -eq $node -or $node.Loaded -or $node.IsPlaceholder -or $node.Path -eq "") {
        return
    }
    $node.Loaded = $true
    $node.Children.Clear()
    foreach ($entry in @((getFolderEntries $node.Path -foldersOnly).Entries)) {
        addFolderTreeChild $node $entry.Name $entry.Path | Out-Null
    }
}

function findFolderNode {
    # ツリーから path のノードを探す。**まだ開いていないフォルダは開かない**（エクスプローラーと同じく、
    # 移動しただけで左のツリーが勝手に展開されないようにする）。見つからなければ $null
    param (
        [FolderNode]$node,
        [string]$path
    )

    if ($node.IsPlaceholder -or $node.Path -eq "") {
        return $null
    }
    if (testSamePath $node.Path $path) {
        return $node
    }
    if (-not $node.Loaded) {
        return $null
    }
    if (-not ([string]$path).StartsWith($node.Path.TrimEnd("\") + "\", [System.StringComparison]::OrdinalIgnoreCase)) {
        return $null
    }
    foreach ($child in $node.Children) {
        $found = findFolderNode $child $path
        if ($null -ne $found) {
            return $found
        }
    }
    return $null
}

function findFolderNodeInTree {
    # ツリー全体（読み込んである範囲）から path のノードを探す
    param (
        [string]$path
    )

    foreach ($root in $script:folderSelect.Roots) {
        foreach ($child in $root.Children) {
            $found = findFolderNode $child $path
            if ($null -ne $found) {
                return $found
            }
        }
    }
    return $null
}

function revealFolderTree {
    # 開いているフォルダをツリーでも選んだ状態にする（ツリーの選択が移動を起こさないよう Syncing を立てる）。
    # ツリーは、すでに開いてあるところと、その 1 つ下までを追随させる（右で下りると 1 段ずつ開く）。
    # 深いパスを貼り付けたときに途中のフォルダをすべて開くことはしない
    param (
        [string]$path
    )

    $d = $script:folderSelect
    $d.Syncing = $true
    try {
        $selected = $d.Ctrl.FolderTree.SelectedItem
        if ($null -ne $selected -and (testSamePath $selected.Path $path)) {
            return
        }
        $found = findFolderNodeInTree $path
        if ($null -eq $found) {
            # 1 つ上のフォルダがツリーにあれば、そこだけ開いて中を出す
            $parentPath = getParentFolderPath $path
            if ($parentPath -ne "") {
                $parent = findFolderNodeInTree $parentPath
                if ($null -ne $parent) {
                    loadFolderNode $parent
                    $parent.SetExpanded($true)
                    $found = findFolderNodeInTree $path
                }
            }
        }
        if ($null -ne $selected) {
            $selected.SetSelected($false)
        }
        if ($null -ne $found) {
            $found.SetSelected($true)
        }
        # ツリーに無いフォルダ（開いていない深いパス・ネットワークのパスなど）は、ツリーの選択を外すだけにする
    } finally {
        $d.Syncing = $false
    }
}

function goFolder {
    # フォルダを開く（一覧・アドレスバー・ツリーをそのフォルダに合わせる）。開けたかを返す
    param (
        [string]$path,
        [bool]$addHistory = $true
    )

    $d = $script:folderSelect
    $path = normalizeFolderPath $path
    if ($path -eq "") {
        setFolderSelectError "フォルダのパスを入力してください。"
        return $false
    }
    $entries = getFolderEntries $path
    if ($entries.Error -ne "") {
        setFolderSelectError "「${path}」を開けません。$($entries.Error)"
        return $false
    }

    setFolderSelectError ""
    $d.Current = $path
    if ($addHistory) {
        pushFolderHistory $path
    }
    $d.Ctrl.AddressBox.Text = $path
    $d.Ctrl.FolderBox.Text = $path
    $d.All = @($entries.Entries | ForEach-Object { newFolderEntry $_ })
    updateFolderEntryList
    $d.Ctrl.StatusText.Text = describeFolderEntries $entries
    revealFolderTree $path
    updateFolderSelectButtons
    return $true
}

function describeFolderEntries {
    # 一覧の下に出す件数（目的のフォルダかどうかの目安にする）
    param (
        $entries
    )

    $text = "フォルダー {0:#,0} 個 ・ Office ファイル {1:#,0} 個" -f $entries.FolderCount, $entries.OfficeCount
    if ($entries.Truncated) {
        $text += "（中身が多いため、先頭だけを表示しています）"
    }
    return $text
}

function newFolderEntry {
    # 一覧の1行を作る（common.ps1 の getFolderEntries が返した中身から）
    param (
        $entry
    )

    $row = [FolderEntry]::new()
    $row.Name = $entry.Name
    $row.Path = $entry.Path
    $row.IsFolder = $entry.IsFolder
    $row.IsOffice = $entry.IsOffice
    $row.Kind = getFolderEntryKind $entry
    $row.UpdatedText = if ($null -ne $entry.Updated) { $entry.Updated.ToString("yyyy/MM/dd H:mm") } else { "" }
    return $row
}

function getFolderEntryKind {
    # 一覧の「種類」列に出す名前
    param (
        $entry
    )

    if ($entry.IsFolder) {
        return "フォルダー"
    }
    if (-not $entry.IsOffice) {
        return "ファイル"
    }
    $ext = [System.IO.Path]::GetExtension($entry.Name).ToLowerInvariant()
    if ($ext.StartsWith(".xls")) {
        return "Excel ブック"
    }
    if ($ext.StartsWith(".doc")) {
        return "Word 文書"
    }
    return "PowerPoint プレゼンテーション"
}

function updateFolderEntryList {
    # 今のフォルダの中身を一覧に出す。
    # ItemsSource には必ず配列を渡す（1 件のときに配列が展開されると渡せないため @() で包む）
    $d = $script:folderSelect
    $rows = @($d.All)
    $d.Ctrl.EntryList.ItemsSource = $rows
    $d.Ctrl.EntryPlaceholder.Text = "このフォルダの中にはフォルダもファイルもありません。このフォルダでよければ［選択］を押してください。"
    $d.Ctrl.EntryPlaceholder.Visibility = if ($rows.Count -eq 0) { "Visible" } else { "Collapsed" }
}

function onFolderEntrySelected {
    # 一覧でフォルダを選んだら、下の欄をそのフォルダにする（選んでいなければ今のフォルダ）
    $d = $script:folderSelect
    $row = $d.Ctrl.EntryList.SelectedItem
    $path = if ($null -ne $row -and $row.IsFolder) { $row.Path } else { $d.Current }
    if ($path -ne "" -and -not (testSamePath $d.Ctrl.FolderBox.Text $path)) {
        $d.Ctrl.FolderBox.Text = $path
        setFolderSelectError ""
    }
}

function openSelectedFolderEntry {
    # 一覧で選んでいるフォルダを開く（ダブルクリック・Enter）
    $row = $script:folderSelect.Ctrl.EntryList.SelectedItem
    if ($null -ne $row -and $row.IsFolder) {
        goFolder $row.Path | Out-Null
    }
}

function focusAddressBox {
    # アドレスバーへ移り、今のパスを選んだ状態にする（F4 / Alt+D / Ctrl+L）
    $address = $script:folderSelect.Ctrl.AddressBox
    $address.Focus() | Out-Null
    $address.SelectAll()
}

function goParentFolder {
    # 1つ上のフォルダへ
    $parent = getParentFolderPath $script:folderSelect.Current
    if ($parent -ne "") {
        goFolder $parent | Out-Null
    }
}

function reloadFolder {
    # 今のフォルダを読み直す（ツリーの下も読み込み直す）
    $d = $script:folderSelect
    if ($d.Current -eq "") {
        return
    }
    $node = $d.Ctrl.FolderTree.SelectedItem
    if ($node -is [FolderNode] -and $node.Loaded) {
        $node.Loaded = $false
        $node.Children.Clear()
        loadFolderNode $node
    }
    goFolder $d.Current $false | Out-Null
}

function pushFolderHistory {
    # ［←］［→］でたどる履歴に足す（今の位置より先は捨てる）
    param (
        [string]$path
    )

    $d = $script:folderSelect
    if ($d.Index -ge 0 -and (testSamePath $d.History[$d.Index] $path)) {
        return
    }
    while ($d.History.Count -gt ($d.Index + 1)) {
        $d.History.RemoveAt($d.History.Count - 1)
    }
    $d.History.Add($path)
    $d.Index = $d.History.Count - 1
}

function moveFolderHistory {
    # 履歴を1つ戻る・進む
    param (
        [int]$step
    )

    $d = $script:folderSelect
    $next = $d.Index + $step
    if ($next -lt 0 -or $next -ge $d.History.Count) {
        return
    }
    $before = $d.Index
    $d.Index = $next
    if (-not (goFolder $d.History[$next] $false)) {
        $d.Index = $before   # 消えたフォルダなどで開けなければ、位置は戻す
    }
    updateFolderSelectButtons
}

function updateFolderSelectButtons {
    # ［←］［→］［↑］の使える・使えないを合わせる
    $d = $script:folderSelect
    $d.Ctrl.BackButton.IsEnabled = ($d.Index -gt 0)
    $d.Ctrl.ForwardButton.IsEnabled = ($d.Index -ge 0 -and $d.Index -lt ($d.History.Count - 1))
    $d.Ctrl.UpButton.IsEnabled = ((getParentFolderPath $d.Current) -ne "")
}

function getDroppedFolders {
    param (
        [System.Windows.DragEventArgs]$e
    )

    if (!$e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
        return @()
    }
    return @($e.Data.GetData([System.Windows.DataFormats]::FileDrop) | Where-Object { Test-Path -LiteralPath $_ -PathType Container })
}

function onFolderDragOver {
    param ($sender, [System.Windows.DragEventArgs]$e)

    $e.Effects = if ((getDroppedFolders $e).Count -gt 0) { "Copy" } else { "None" }
    $e.Handled = $true
}

function getTargetsKey {
    # 変換対象フォルダの一覧（@{ Path; Enabled } の配列）を比べるための文字列
    param (
        [object[]]$folders
    )

    return (@($folders | Where-Object { $_ } | ForEach-Object { "$($_.Enabled)`t$($_.Path)" }) -join "`n")
}

function formatTime {
    # 当日なら HH:mm、それ以前は M/d HH:mm
    param (
        $time
    )

    if ($null -eq $time) {
        return ""
    }
    if ($time.Date -eq (Get-Date).Date) {
        return $time.ToString("H:mm")
    }
    return $time.ToString("M/d H:mm")
}

function readTextShared {
    # 変換側が書き込み中でも妨げないよう、共有を許して読む
    param (
        [string]$path
    )

    if (!(Test-Path -LiteralPath $path)) {
        return ""
    }
    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $stream = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
    $reader = New-Object System.IO.StreamReader($stream, ${utf8Bom})
    try {
        return $reader.ReadToEnd()
    } finally {
        $reader.Dispose()
    }
}

# ---- 別スレッドの処理（インデックスの件数など） ----

$script:jobs = New-Object System.Collections.ArrayList

function startJob {
    # scriptBlock を別スレッドで実行し、終わったら画面のスレッドで onDone { param($output, $errorText) } を呼ぶ
    param (
        [scriptblock]$scriptBlock,
        [object[]]$arguments,
        [scriptblock]$onDone
    )

    $ps = [powershell]::Create()
    [void]$ps.AddScript($scriptBlock.ToString())
    foreach ($argument in $arguments) {
        [void]$ps.AddArgument($argument)
    }
    [void]$script:jobs.Add(@{ PS = $ps; Handle = $ps.BeginInvoke(); OnDone = $onDone })
    $script:jobTimer.Start()
}

$script:jobTimer = newTimer 200 {
    safe {
        foreach ($job in @($script:jobs.ToArray())) {
            if (!$job.Handle.IsCompleted) {
                continue
            }
            $script:jobs.Remove($job)
            $output = $null
            $errorText = $null
            try {
                $output = $job.PS.EndInvoke($job.Handle)
                if ($job.PS.Streams.Error.Count -gt 0) {
                    $errorText = $job.PS.Streams.Error[0].ToString()
                }
            } catch {
                $errorText = $_.Exception.Message
            } finally {
                $job.PS.Dispose()
            }
            & $job.OnDone $output $errorText
        }
        if ($script:jobs.Count -eq 0) {
            $script:jobTimer.Stop()
        }
    }
}

# ============================================================================
# ［1 インデックス管理］（インデックスの作成・編集・削除と、変換の実行）
# ============================================================================

$script:targetItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$ui.IndexGrid.ItemsSource = $script:targetItems
$script:loadingTargets = $false
# ［変換］チェックのクリックで保存する（TwoWay バインドで Enabled は更新済み。PS class のプレーンな
# プロパティは PropertyChanged を出さないため、購読ではなくここで保存する）
$ui.IndexGrid.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler] {
    param ($s, $e)
    safe {
        $cb = $e.OriginalSource
        if ($cb -is [System.Windows.Controls.CheckBox] -and $cb.DataContext -is [FolderItem] -and !$script:loadingTargets) {
            saveTargets
            updateConvertButton
        }
    }
})
$script:savedTargets = $null  # 最後に読み込み・保存したインデックス一覧（getTargetsKey）。ほかでの変更の検出に使う
$script:editDialog = $null    # 新規作成・編集のダイアログ（開いている間だけ）
$script:convertProcess = $null
$script:convertStart = $null
$script:convertFailed = 0  # 変換中に一覧へ反映済みの失敗件数
$script:conversionState = $null
$script:indexSummary = $null

function isConverting {
    return ($null -ne $script:convertProcess) -and !$script:convertProcess.HasExited
}

function updateFolderItemStatus {
    param (
        $item
    )

    if (Test-Path -LiteralPath $item.Path -PathType Container) {
        $item.SetStatus("✓ フォルダがあります", ${okBrush})
    } else {
        $item.SetStatus("✗ フォルダが見つかりません", ${ngBrush})
    }
}

function newFolderItem {
    param (
        [string]$path,
        [bool]$enabled,
        [string]$name = ""
    )

    $item = New-Object FolderItem
    $item.Name = $name
    $item.Path = $path
    $item.Enabled = $enabled
    $item.FileCountText = "－"
    $item.LastConvertedText = ""
    updateFolderItemStatus $item
    # ［変換］チェックの保存は、一覧のチェックボックスの Click（IndexGrid.AddHandler）で行う。
    # PS class のプレーンなプロパティは TwoWay セットで PropertyChanged を出さないため、購読では拾えない。
    return $item
}

function getUsedIndexNames {
    # 一覧のインデックス名の集合（大文字・小文字を区別しない）。except に渡した行の名前は含めない
    param (
        $except = $null
    )

    $used = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($item in $script:targetItems) {
        if ($item -ne $except -and $item.Name) {
            [void]$used.Add($item.Name)
        }
    }
    # , を付けて、集合そのものを返す（付けないと PowerShell が中身を展開し、
    # 空なら $null・1 個なら文字列になって Contains の意味が変わる）
    return ,$used
}

function loadTargets {
    $script:loadingTargets = $true
    try {
        $script:targetItems.Clear()
        $folders = @(getTargetFolders)
        # 名前の決まっていないインデックス（以前の版の設定から移した直後など）には、ここで名前を割り当てて確定する。
        # 一覧・編集・削除はインデックス名で扱うため、画面に出す時点で名前があるようにする（変換側と同じ assignIndexNames を使う）
        if (@($folders | Where-Object { $_ -and !$_.Name }).Count -gt 0) {
            $folders = @(assignIndexNames $folders (readStatusFile).Folders)
            writeTargetFolders $folders
        }
        foreach ($folder in $folders) {
            $script:targetItems.Add((newFolderItem $folder.Path $folder.Enabled $folder.Name))
        }
        $script:savedTargets = getTargetsKey @(getTargetFolders)
    } finally {
        $script:loadingTargets = $false
    }
    updateIndexListView
}

function saveTargets {
    writeTargetFolders @($script:targetItems | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Path = $_.Path; Enabled = $_.Enabled } })
    $script:savedTargets = getTargetsKey @(getTargetFolders)
    setStatus "インデックス一覧を保存しました（$(Get-Date -Format 'H:mm')）"
}

function updateIndexSourceFile {
    # インデックスのフォルダの 元のフォルダ.txt を今の一覧に合わせて書き直す。
    # 次の変換を待たずに、検索結果から元のファイルを開けるようにする（インデックスが無ければ何もしない）
    if (!(Test-Path -LiteralPath ${indexDir} -PathType Container)) {
        return
    }
    writeSourceFolderFile @($script:targetItems | Where-Object { $_.Name } | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Path = $_.Path } })
}

function refreshIndexViews {
    # インデックスを作成・編集・削除した後、検索タブ（検索対象のツリー・件数）も読み直す
    $script:sourceFolderMaps = @{}
    $script:indexSummary = $null
    loadIndexTree
    refreshIndexSummary
    refreshConversionState
}

function applyIndexStats {
    # 変換一覧の集計（getIndexStats）を一覧の各行のファイル数・最終変換に反映する
    param (
        $stats
    )

    foreach ($item in $script:targetItems) {
        $stat = $null
        if ($item.Name -and $null -ne $stats -and $stats.ContainsKey($item.Name)) {
            $stat = $stats[$item.Name]
        }
        if ($null -eq $stat) {
            $item.SetStats("－", "まだ変換していません", "")
            continue
        }
        $converted = [datetime]::MinValue
        $lastText = if ($stat.LastConverted -and [datetime]::TryParseExact($stat.LastConverted, "yyyy/MM/dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$converted)) {
            formatTime $converted
        } else {
            ""
        }
        $item.SetStats(("{0:#,0}" -f $stat.Total), ("済 {0:#,0} 件 ・ 未変換 {1:#,0} 件 ・ 失敗 {2:#,0} 件" -f $stat.Done, $stat.Pending, $stat.Failed), $lastText)
    }
}

function updateIndexListView {
    $ui.IndexGridPlaceholder.Visibility = if ($script:targetItems.Count -eq 0) { "Visible" } else { "Collapsed" }
    updateConvertButton
}

function testIndexOperable {
    # 変換中はインデックスの作成・編集・削除をしない（インデックスのフォルダ・変換一覧を変換側が使っているため）
    param (
        [string]$operation
    )

    if (isConverting) {
        showMessage "変換中はインデックスを${operation}できません。変換が終わるまでお待ちください（［中止］で止められます）。" "OK" "Warning" | Out-Null
        return $false
    }
    if ($script:indexBusy) {
        # 前のインデックスの TSV を削除している最中（別スレッド）
        showMessage "前のインデックスの削除が終わるまでお待ちください。" "OK" "Warning" | Out-Null
        return $false
    }
    return $true
}

function showIndexEditDialog {
    # インデックスの新規作成・編集のダイアログ。決めた内容 @{ Path; Name } を返す（キャンセルは $null）。
    #   item: 編集するインデックス（$null なら新規作成）
    param (
        $item = $null
    )

    $dialog = loadWindow "$PSScriptRoot\config_gui_index_edit.xaml"
    $dialog.Owner = $window
    $ctrl = @{}
    foreach ($name in @("OkButton", "BrowseButton", "FolderBox", "NameBox", "IntroText", "NoticeText", "ErrorText")) {
        $ctrl[$name] = $dialog.FindName($name)
    }
    $script:editDialog = @{ Window = $dialog; Ctrl = $ctrl; Item = $item; Suggested = "" }

    if ($null -eq $item) {
        $dialog.Title = "インデックスの新規作成"
        $ctrl.IntroText.Text = "Office ファイル（Excel・Word・PowerPoint）のあるフォルダを 1 つ指定すると、そのフォルダのインデックスを作ります。" +
            "一覧に加えるだけで、中身の変換は［変換を開始］を押してから始まります。"
    } else {
        $dialog.Title = "インデックスの編集"
        $ctrl.IntroText.Text = "インデックスの名前と、元のフォルダの場所を変えられます。"
        $ctrl.FolderBox.Text = $item.Path
        $ctrl.NameBox.Text = $item.Name
        $ctrl.NoticeText.Visibility = "Visible"
        $ctrl.NoticeText.Text = "変換済みのインデックスは作り直しません（名前を変えるときは work\index のフォルダごと名前を変えます）。" +
            "フォルダを別のドライブ・共有フォルダへ移した場合は、ここで場所を変えてください。次の変換では、更新されたファイルだけを変換します。"
    }

    $ctrl.FolderBox.Add_TextChanged({
        safe {
            # 新規作成のときは、フォルダ名からインデックス名を自動で入れる（利用者が名前を変えた後は触らない）
            $d = $script:editDialog
            if ($null -ne $d.Item -or ($d.Ctrl.NameBox.Text -ne "" -and $d.Ctrl.NameBox.Text -ne $d.Suggested)) {
                return
            }
            $path = normalizeFolderPath $d.Ctrl.FolderBox.Text
            $d.Suggested = if ($path -eq "") { "" } else { newIndexName $path (getUsedIndexNames) }
            $d.Ctrl.NameBox.Text = $d.Suggested
        }
    })
    $ctrl.BrowseButton.Add_Click({
        safe {
            $d = $script:editDialog
            $initial = normalizeFolderPath $d.Ctrl.FolderBox.Text
            $path = selectFolder "インデックスにする、Office ファイルのあるフォルダを選んでください" $initial $d.Window
            if ($path) {
                $d.Ctrl.FolderBox.Text = $path
            }
        }
    })
    $ctrl.FolderBox.Add_PreviewDragOver({ onFolderDragOver @args })
    $ctrl.FolderBox.Add_PreviewDrop({
        param ($sender, $e)
        safe {
            $folders = @(getDroppedFolders $e)
            if ($folders.Count -gt 0) {
                $script:editDialog.Ctrl.FolderBox.Text = $folders[0]
            }
        }
        $e.Handled = $true
    })
    $ctrl.OkButton.Add_Click({
        safe {
            $d = $script:editDialog
            $message = checkIndexEditInput
            if ($message -ne "") {
                $d.Ctrl.ErrorText.Text = $message
                $d.Ctrl.ErrorText.Visibility = "Visible"
                return
            }
            $d.Window.DialogResult = $true
        }
    })

    $result = $null
    if ($dialog.ShowDialog()) {
        $result = @{ Path = (normalizeFolderPath $ctrl.FolderBox.Text); Name = $ctrl.NameBox.Text.Trim() }
    }
    $script:editDialog = $null
    return $result
}

function checkIndexEditInput {
    # 新規作成・編集のダイアログの入力を調べ、直してほしい内容を返す（問題なければ空文字列）
    $d = $script:editDialog
    $path = normalizeFolderPath $d.Ctrl.FolderBox.Text
    if ($path -eq "") {
        return "元のフォルダを指定してください。"
    }
    foreach ($other in $script:targetItems) {
        if ($other -eq $d.Item) {
            continue
        }
        if (testSameFolder $other.Path $path) {
            return "「${path}」のインデックス [$($other.Name)] が既にあります。"
        }
        # 入れ子のフォルダは、同じファイルが2つのインデックスに入り、変換も検索結果も二重になるため登録しない
        if (testFolderUnder $path $other.Path) {
            return "「${path}」は、インデックス [$($other.Name)]（$($other.Path)）の中のフォルダです。" +
                "同じファイルが二重に変換されるため、登録できません。検索する範囲を絞るときは［2 検索］の検索対象で外してください。"
        }
        if (testFolderUnder $other.Path $path) {
            return "「${path}」の中には、インデックス [$($other.Name)]（$($other.Path)）があります。" +
                "同じファイルが二重に変換されるため、登録できません。まとめるときは、先に [$($other.Name)] を削除してください。"
        }
    }
    return (testIndexName ($d.Ctrl.NameBox.Text.Trim()) @(getUsedIndexNames $d.Item))
}

function addIndexItem {
    # インデックスを一覧に加えて保存する
    param (
        [string]$path,
        [string]$name
    )

    $item = newFolderItem $path $true $name
    $script:targetItems.Add($item)
    $ui.IndexGrid.SelectedItem = $item
    $ui.IndexGrid.ScrollIntoView($item)
    saveTargets
    updateIndexSourceFile
    updateIndexListView
    refreshConversionState
    if (Test-Path -LiteralPath $path -PathType Container) {
        setStatus "インデックス [${name}] を作成しました。［変換を開始］を押すと中身を変換します"
    } else {
        setStatus "インデックス [${name}] を作成しましたが、フォルダが見つかりません：${path}"
    }
}

function newIndex {
    # ［新規作成…］。フォルダとインデックス名を決めて一覧に加える（変換はしない）
    if (!(testIndexOperable "作成")) {
        return
    }
    $result = showIndexEditDialog $null
    if ($null -eq $result) {
        return
    }
    addIndexItem $result.Path $result.Name
}

function addIndexForFolder {
    # 一覧へのドラッグ＆ドロップでインデックスを作る（名前はフォルダ名から自動で決める）
    param (
        [string]$path
    )

    if (!(testIndexOperable "作成")) {
        return
    }
    $path = normalizeFolderPath $path
    if ($path -eq "") {
        return
    }
    foreach ($item in $script:targetItems) {
        # 書き方が違うだけで同じフォルダ（ネットワークドライブと UNC パスなど）も、すでにあるとみなす
        if (testSameFolder $item.Path $path) {
            $ui.IndexGrid.SelectedItem = $item
            setStatus "「$($item.Path)」のインデックス [$($item.Name)] は既にあります"
            return
        }
    }
    addIndexItem $path (newIndexName $path (getUsedIndexNames))
}

function editIndex {
    # ［編集…］。インデックス名と元のフォルダの場所を変える。インデックスは作り直さない
    $item = $ui.IndexGrid.SelectedItem
    if ($null -eq $item -or !(testIndexOperable "編集")) {
        return
    }
    $result = showIndexEditDialog $item
    if ($null -eq $result) {
        return
    }

    $changes = New-Object System.Collections.Generic.List[string]
    if ($result.Name -ne $item.Name) {
        # インデックスのフォルダ（work\index\<名前>）と変換一覧の記録も名前を変える（中身は作り直さない）
        renameIndex $item.Name $result.Name
        $changes.Add("名前 [$($item.Name)] → [$($result.Name)]")
        $item.SetName($result.Name)
    }
    if ($result.Path -ne $item.Path) {
        $changes.Add("場所 $($item.Path) → $($result.Path)")
        $item.SetPath($result.Path)
        updateFolderItemStatus $item
    }
    if ($changes.Count -eq 0) {
        return
    }

    saveTargets
    updateIndexSourceFile
    refreshIndexViews
    setStatus ("インデックスを変更しました（" + ($changes -join " / ") + "）")
}

function rebuildIndex {
    # ［作り直す…］。変換した TSV と変換一覧の記録を消してから変換を始め、インデックスを一から作り直す。
    # 差分変換（更新日時とサイズで判定する）では変換し直さない場合に使う:
    #   ・更新日時・サイズが変わらないまま中身が変わった（同じ秒に同じ大きさで保存した・更新日時を保つツールで書き換えた）
    #   ・TSV の中身が壊れた（0 バイトのTSVは変換側が見つけて作り直すが、中身の書き換えまでは分からない）
    $item = $ui.IndexGrid.SelectedItem
    if ($null -eq $item -or !(testIndexOperable "作り直し")) {
        return
    }
    if (!$item.Name) {
        setStatus "このインデックスはまだ変換していません。［変換を開始］で作成してください"
        return
    }
    if (!(Test-Path -LiteralPath $item.Path -PathType Container)) {
        # 消してから変換できないと、インデックスが無いだけの状態になる
        showMessage ("元のフォルダが見つからないため、インデックス [$($item.Name)] を作り直せません。`n`n" +
            "元のフォルダ：$($item.Path)`n`nフォルダを使えるようにするか、［編集…］で場所を変えてください。") "OK" "Warning" | Out-Null
        return
    }

    $answer = showMessage ("インデックス [$($item.Name)] を作り直します。`n`n元のフォルダ：$($item.Path)`n`n" +
        "変換した TSV（work\index\$($item.Name)）を削除し、フォルダの中の Office ファイルをすべて変換し直します（件数によっては時間がかかります）。`n" +
        "ふだんは、更新されたファイルだけを変換する［変換を開始］で足ります。元のファイルの更新日時が変わらないまま中身が変わった場合などに使ってください。`n" +
        "［変換］のチェックが外れている場合は付けます。`n`n作り直しますか？") "YesNo" "Question" "No"
    if ($answer -ne "Yes") {
        return
    }

    $item.SetEnabled($true)  # チェックが外れていると変換されず、インデックスが無いだけになる
    saveTargets
    updateIndexSourceFile
    updateIndexListView
    # TSV の削除は数万フォルダで数十秒かかることがあるため、別スレッドで行う（画面は固まらない）
    startIndexRemoveJob $item.Name "作り直し" {
        startConversion
        $name = $script:indexJobName
        if (isConverting) {
            setStatus "インデックス [${name}] を作り直します（変換を開始しました）"
        } else {
            # 変換を始めるときの確認（失敗分の再変換）でキャンセルした場合。TSV は削除済みのため、次の変換で作り直す
            setStatus "インデックス [${name}] の TSV を削除しました。［変換を開始］を押すと作り直します"
        }
    }
}

function startIndexRemoveJob {
    # インデックス（work\index\<名前>）と変換一覧の記録の削除を別スレッドで行う。
    # 数万フォルダの削除は数十秒かかることがあり、画面のスレッドで行うと「応答なし」になるため。
    # 終わるまでインデックスの操作・変換の開始はできないようにし、何をしているかをステータスに出す
    param (
        [string]$name,
        [string]$operation,   # "削除" / "作り直し"（表示に使う）
        [scriptblock]$onDone  # 削除が終わった後に画面のスレッドで行うこと（$script:indexJobName で名前を参照できる）
    )

    $script:indexBusy = $true
    $script:indexJobName = $name
    $script:indexJobOnDone = $onDone
    $script:indexJobOperation = $operation
    updateConvertButton
    setStatus "インデックス [${name}] の TSV を削除しています…（件数によっては少し時間がかかります）"
    startJob {
        param ($commonPath, $name)
        . $commonPath
        removeIndex $name
    } @(${commonPath}, $name) {
        param ($output, $errorText)
        $script:indexBusy = $false
        updateConvertButton
        $name = $script:indexJobName
        if ($errorText) {
            setStatus "インデックス [${name}] の $($script:indexJobOperation)に失敗しました：${errorText}"
            refreshIndexViews
            return
        }
        refreshIndexViews
        if ($script:indexJobOnDone) {
            & $script:indexJobOnDone
        }
    }
}

function deleteIndex {
    # ［削除］。一覧から削除し、変換した TSV（work\index\<名前>）と変換一覧の記録も削除する
    $item = $ui.IndexGrid.SelectedItem
    if ($null -eq $item -or !(testIndexOperable "削除")) {
        return
    }

    $answer = showMessage ("インデックス [$($item.Name)] を削除します。`n`n元のフォルダ：$($item.Path)`n`n" +
        "変換した TSV（work\index\$($item.Name)）と変換一覧の記録を削除します。元のフォルダと Office ファイルは削除しません。`n" +
        "一時的に変換しないだけなら、削除せずに［変換］のチェックを外してください。`n`n削除しますか？") "YesNo" "Question" "No"
    if ($answer -ne "Yes") {
        return
    }

    # 一覧からはすぐ消し、TSV の削除（時間がかかることがある）は別スレッドで行う
    $script:targetItems.Remove($item)
    saveTargets
    updateIndexSourceFile
    updateIndexListView
    startIndexRemoveJob $item.Name "削除" {
        setStatus "インデックス [$($script:indexJobName)] を削除しました"
    }
}

function updateConvertButton {
    $ready = $false
    foreach ($item in $script:targetItems) {
        if ($item.Enabled -and (Test-Path -LiteralPath $item.Path -PathType Container)) {
            $ready = $true
            break
        }
    }

    $state = $script:conversionState
    if ($script:indexBusy) {
        # インデックスの削除中（別スレッド）は、変換もインデックスの操作も始めない
        $ready = $false
    }
    if (isConverting) {
        $ui.ConvertButton.Content = "変換中…"
        $ui.ConvertButton.IsEnabled = $false
    } else {
        $ui.ConvertButton.Content = if ($state -and $state.Pending -gt 0) { "続きから再開（残り $($state.Pending) 件）" } else { "変換を開始" }
        $ui.ConvertButton.IsEnabled = $ready
    }

    $hint = "変換中も検索できます。途中でやめるときは［中止］を押してください（次回、続きから再開できます）。"
    if (!$ready -and !(isConverting)) {
        $hint = "インデックスを作成して、チェックを付けてください。"
    } elseif ($state -and $state.Failed -gt 0 -and !(isConverting)) {
        $hint = "前回失敗したファイルがあります。変換を始めるときに、再変換するかを選べます。" + $hint
    }
    $ui.ConvertHint.Text = $hint

    # インデックスの作成・編集・削除は、選んでいるかどうかと変換中かどうかで切り替える
    # （変換中はインデックスのフォルダ・変換一覧を変換側が使っているため触らない）
    $selected = $null -ne $ui.IndexGrid.SelectedItem
    $editable = !(isConverting) -and !$script:indexBusy
    $ui.NewIndexButton.IsEnabled = $editable
    $ui.EditIndexButton.IsEnabled = $selected -and $editable
    $ui.RebuildIndexButton.IsEnabled = $selected -and $editable
    $ui.RemoveIndexButton.IsEnabled = $selected -and $editable
}

function refreshConversionState {
    # 変換一覧の集計は、ファイルが数万行になると数秒〜十数秒かかる。
    # 画面のスレッドで行うと、起動時・タブの切り替え時に画面が固まる（応答なしになる）ため別スレッドで数える
    if ($script:stateRunning) {
        $script:stateAgain = $true
        return
    }
    $script:stateRunning = $true
    $script:stateAgain = $false
    startJob {
        param ($commonPath)
        . $commonPath
        getConversionState
    } @(${commonPath}) {
        param ($output, $errorText)
        $script:stateRunning = $false
        # 変換側が書き込んでいる瞬間などは、次の機会に読み直す
        if ($output -and $output.Count -gt 0 -and $output[0]) {
            applyConversionState $output[0]
        }
        if ($script:stateAgain) {
            refreshConversionState
        }
    }
}

function applyConversionState {
    # 集計（別スレッド）の結果を画面に反映する
    param (
        $state  # getConversionState の結果
    )

    $script:conversionState = $state

    # 失敗したファイルは下の一覧に原因とともに表示する
    $ui.ConversionStateText.Text = if ($state.Pending -gt 0 -and !(isConverting)) { "⏸ 前回の変換が中断しています（残り $($state.Pending) 件）" } else { "" }
    $ui.IndexTabHeader.Text = if ($state.Failed -gt 0) { "⚠ 1 インデックス管理" } else { "1 インデックス管理" }
    applyIndexStats $state.IndexStats
    updateFailedList $state
    updateIndexSummaryText
    updateConvertButton
}

function updateFailedList {
    # 変換に失敗したファイルと原因（変換一覧のエラー列）を一覧に表示する
    param (
        $state  # getConversionState の結果
    )

    $folderPaths = @{}  # インデックス名 → 変換対象フォルダ（大文字・小文字を区別しない）
    foreach ($folder in $state.Folders) {
        if ($folder.Name) {
            $folderPaths[$folder.Name] = $folder.Path
        }
    }

    $rows = New-Object System.Collections.ArrayList
    foreach ($status in $state.FailedRows) {
        $row = New-Object FailRow
        $row.RelPath = $status.相対パス
        $row.Reason = if ($status.エラー) { $status.エラー } else { "（原因は記録されていません）" }
        $converted = [datetime]::MinValue
        if ([datetime]::TryParseExact([string]$status.変換日時, "yyyy/MM/dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$converted)) {
            $row.ConvertedText = formatTime $converted
        }
        $parts = splitIndexRelPath $status.相対パス
        if ($folderPaths.ContainsKey($parts.Name)) {
            $row.SourcePath = Join-Path $folderPaths[$parts.Name] $parts.Rest
        }
        [void]$rows.Add($row)
    }

    $ui.FailedGrid.ItemsSource = $rows
    $ui.FailedHeading.Text = "⚠ 変換に失敗したファイル $($rows.Count) 件"
    $ui.FailedPanel.Visibility = if ($rows.Count -gt 0) { "Visible" } else { "Collapsed" }
}

function openFailedFileFolder {
    # 失敗したファイルの場所をエクスプローラーで開く（ファイルを選択した状態）
    $row = $ui.FailedGrid.SelectedItem
    if ($null -eq $row) {
        return
    }
    if (!$row.SourcePath) {
        setStatus "元のファイルの場所が分かりません（変換一覧に変換対象フォルダの記録がありません）：$($row.RelPath)"
        return
    }
    if (Test-Path -LiteralPath $row.SourcePath -PathType Leaf) {
        Start-Process -FilePath "explorer.exe" -ArgumentList "/select,`"$($row.SourcePath)`""
        return
    }
    $dir = Split-Path $row.SourcePath -Parent
    if (Test-Path -LiteralPath $dir -PathType Container) {
        Start-Process -FilePath "explorer.exe" -ArgumentList "`"${dir}`""
        setStatus "ファイルが見つからないため、フォルダを開きました（移動・削除された可能性があります）：$($row.SourcePath)"
        return
    }
    setStatus "ファイルが見つかりません（移動・削除された可能性があります）：$($row.SourcePath)"
}

function updateIndexSummaryText {
    $summary = $script:indexSummary
    if ($null -eq $summary) {
        $ui.IndexSummaryText.Text = "確認中…"
        return
    }
    if ($summary["Count"] -eq 0) {
        $ui.IndexSummaryText.Text = "まだインデックスがありません。"
        return
    }
    $text = "TSV $($summary['Count'].ToString('N0')) 件 ・ 最終変換 $(formatTime $summary['LastWrite'])"
    $state = $script:conversionState
    if ($state -and $state.Done -gt 0) {
        $text = "変換済み $($state.Done.ToString('N0')) ファイル（$text）"
    }
    $ui.IndexSummaryText.Text = $text
}

function refreshIndexSummary {
    # TSV の件数は数えるのに時間がかかることがあるため、別スレッドで数える
    if ($script:summaryRunning) {
        $script:summaryAgain = $true
        return
    }
    $script:summaryRunning = $true
    $script:summaryAgain = $false
    $folders = @(${indexDir})
    startJob {
        param ($commonPath, $folders)
        . $commonPath
        getIndexSummary $folders
    } @(${commonPath}, $folders) {
        param ($output, $errorText)
        $script:summaryRunning = $false
        if ($output -and $output.Count -gt 0) {
            $script:indexSummary = $output[0]
        }
        updateIndexSummaryText
        updateSearchTarget
        if ($script:summaryAgain) {
            refreshIndexSummary
        }
    }
}

# ---- 変換の起動と進み具合 ----

function getConversionProgress {
    # 変換の進み具合を返す（変換側が書く 変換進捗.txt の1行を読む）。
    # 変換一覧（数万行）を読み直すと1回に数秒かかり、毎秒読むと画面が固まるため、この1行だけを読む
    param (
        [datetime]$since
    )

    $progress = @{ Scanned = $false; Processed = 0; Failed = 0; Remaining = 0; Current = ""; Detail = ""; Finishing = $false }
    $current = readConvertProgress
    if ($null -eq $current) {
        return $progress
    }

    $progress.Scanned = ($current.Phase -ne ${convertPhaseScan})
    $progress.Finishing = ($current.Phase -eq ${convertPhaseFinish})
    $progress.Processed = $current.Processed
    $progress.Remaining = $current.Remaining
    $progress.Failed = $current.Failed
    $progress.Detail = $current.Detail
    if ($progress.Scanned) {
        $progress.Current = $current.Detail  # 変換中のファイルの相対パス
    }
    return $progress
}

function findRunningConversion {
    # このツールの変換（office_to_tsv.ps1）が実行中なら、そのプロセスを返す（画面を閉じて開き直した場合など）
    $script = "${PSScriptRoot}\office_to_tsv.ps1"
    foreach ($process in @(Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" -ErrorAction SilentlyContinue)) {
        if ($process.CommandLine -and $process.CommandLine.IndexOf($script, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
            try {
                return Get-Process -Id $process.ProcessId -ErrorAction Stop
            } catch {
            }
        }
    }
    return $null
}

function showConversionPanel {
    $ui.ConvertProgressPanel.Visibility = "Visible"
    $ui.ConvertProgress.IsIndeterminate = $true
    $ui.ConvertProgressText.Text = "変換の準備をしています…"
    $ui.ConvertProgressEta.Text = ""
    $ui.ConvertProgressDetail.Text = "変換対象のファイルを確認しています。"
    $ui.ConvertStopButton.Visibility = "Visible"
    $ui.ConvertStopButton.IsEnabled = $true
    $ui.ConvertLogButton.Visibility = "Collapsed"
    $taskbar.ProgressState = "Indeterminate"
}

function startConversion {
    if (isConverting) {
        return
    }
    $existing = findRunningConversion
    if ($existing) {
        adoptConversion $existing
        setStatus "実行中の変換があるため、その進み具合を表示します"
        return
    }

    # 前回失敗し、その後更新されていないファイルを再変換するか聞く
    $retryFailed = $false
    $state = $script:conversionState
    if ($state -and $state.Failed -gt 0) {
        $answer = showMessage ("前回変換に失敗し、その後更新されていないファイルが $($state.Failed) 件あります（パスワード付きなど）。`n`n" +
            "これらも再変換しますか？`n（「いいえ」の場合はスキップして、新しいファイル・更新されたファイルだけを変換します）") "YesNoCancel" "Question" "No"
        if ($answer -eq "Cancel") {
            return
        }
        $retryFailed = $answer -eq "Yes"
    }

    saveTargets
    $arguments = "-NoProfile -ExecutionPolicy RemoteSigned -WindowStyle Hidden -File `"${PSScriptRoot}\office_to_tsv.ps1`""
    if ($retryFailed) {
        $arguments += " -RetryFailed"
    }
    $script:convertStart = Get-Date
    $script:convertRate = $null
    $script:convertProcess = Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -WorkingDirectory ${rootDir} -WindowStyle Hidden -PassThru
    # PowerShell 5.1 では、起動直後にハンドルを取っておかないと終了コードを取得できないことがある
    $null = $script:convertProcess.Handle
    $script:convertAdopted = $false

    showConversionPanel
    setStatus "変換を開始しました"
    updateConvertButton
    updateKillBadge
    $script:convertTimer.Start()
}

function adoptConversion {
    # 画面の外で起動された（または前回の画面で起動した）変換の進み具合を表示する
    param (
        [System.Diagnostics.Process]$process
    )

    $script:convertProcess = $process
    try {
        $null = $process.Handle
    } catch {
    }
    $script:convertStart = $process.StartTime
    $script:convertRate = $null
    $script:convertAdopted = $true
    showConversionPanel
    updateConvertButton
    $script:convertTimer.Start()
}

function stopConversion {
    if (!(isConverting)) {
        return
    }
    $answer = showMessage "変換を中止しますか？`n変換中のファイルが終わったところで止まります。次回は続きから再開できます。" "YesNo" "Question" "No"
    if ($answer -ne "Yes") {
        return
    }
    [System.IO.File]::WriteAllText(${stopRequestFile}, "", ${utf8Bom})
    $ui.ConvertStopButton.IsEnabled = $false
    $ui.ConvertProgressDetail.Text = "中止しています…（変換中のファイルが終わるまでお待ちください）"
    setStatus "変換の中止を要求しました"
}

function updateConversionProgress {
    if (!(isConverting)) {
        finishConversion
        return
    }

    try {
        $progress = getConversionProgress $script:convertStart
    } catch {
        return
    }
    $stopping = !$ui.ConvertStopButton.IsEnabled

    $total = $progress.Processed + $progress.Remaining
    if (!$progress.Scanned) {
        # 変換対象を探している間（大きいフォルダ・ネットワーク越しでは数分かかることがある）。
        # 何を見ているかが分かるよう、変換側が書いた内容をそのまま出す
        $ui.ConvertProgress.IsIndeterminate = $true
        $ui.ConvertProgressText.Text = "変換対象のファイルを確認しています…"
        if (!$stopping) {
            $ui.ConvertProgressDetail.Text = [string]$progress.Detail
        }
        $taskbar.ProgressState = "Indeterminate"
        return
    }
    if ($progress.Finishing) {
        # 後片付け（Officeアプリの終了・変換一覧の書き直し）。止まって見えないよう、何をしているかを出す
        $ui.ConvertProgress.IsIndeterminate = $true
        $ui.ConvertProgressText.Text = "変換を終えています…"
        $ui.ConvertProgressDetail.Text = [string]$progress.Detail
        $taskbar.ProgressState = "Indeterminate"
        return
    }
    if ($progress.Processed -eq 0) {
        $ui.ConvertProgress.IsIndeterminate = $true
        $ui.ConvertProgressText.Text = if ($progress.Remaining -gt 0) { "$($progress.Remaining) 件のファイルを変換します" } else { "変換が必要なファイルを確認しています…" }
        if (!$stopping) {
            $ui.ConvertProgressDetail.Text = if ($progress.Current) { "変換中のファイル：$($progress.Current)" } else { "" }
        }
        $taskbar.ProgressState = "Indeterminate"
        return
    }

    $ratio = if ($total -gt 0) { $progress.Processed / $total } else { 1 }
    $ui.ConvertProgress.IsIndeterminate = $false
    $ui.ConvertProgress.Value = $ratio
    $taskbar.ProgressState = if ($progress.Failed -gt 0) { "Paused" } else { "Normal" }
    $taskbar.ProgressValue = $ratio

    $text = "変換中… $($progress.Processed.ToString('N0')) / $($total.ToString('N0')) 件"
    if ($progress.Failed -gt 0) {
        $text += "（失敗 $($progress.Failed) 件）"
    }
    $ui.ConvertProgressText.Text = $text

    # 残り時間の目安（最初の1件が終わってからの速さで計算する）
    $now = Get-Date
    if ($null -eq $script:convertRate) {
        $script:convertRate = @{ Time = $now; Processed = $progress.Processed }
    }
    $done = $progress.Processed - $script:convertRate.Processed
    if ($progress.Remaining -eq 0) {
        $ui.ConvertProgressEta.Text = ""
    } elseif ($done -gt 0) {
        $seconds = ($now - $script:convertRate.Time).TotalSeconds / $done * $progress.Remaining
        $ui.ConvertProgressEta.Text = if ($seconds -lt 60) { "残り 1 分未満" } else { "残り約 $([math]::Ceiling($seconds / 60)) 分" }
    }
    if (!$stopping) {
        $ui.ConvertProgressDetail.Text = if ($progress.Current) { "変換中のファイル：$($progress.Current)" } else { "" }
    }
}

function finishConversion {
    $script:convertTimer.Stop()
    $taskbar.ProgressState = "None"
    # 変換完了の通知。以前はタスクバーのボタンを光らせていたが（FlashWindowEx）、P/Invoke は
    # 実行時コンパイル（csc.exe）を無くすため廃止した。完了は進捗表示・ステータスで分かる。
    $exitCode = $null
    try {
        $script:convertProcess.WaitForExit()
        $exitCode = $script:convertProcess.ExitCode
    } catch {
    }
    $progress = $null
    try {
        $progress = getConversionProgress $script:convertStart
    } catch {
    }

    $counts = ""
    if ($progress -and $progress.Processed -gt 0) {
        $counts = "成功 $($progress.Processed - $progress.Failed) 件 / 失敗 $($progress.Failed) 件"
        if ($progress.Remaining -gt 0) {
            $counts += " / 残り $($progress.Remaining) 件"
        }
        $ui.ConvertProgress.IsIndeterminate = $false
        $ui.ConvertProgress.Value = $progress.Processed / ($progress.Processed + $progress.Remaining)
    } else {
        $ui.ConvertProgress.IsIndeterminate = $false
        $ui.ConvertProgress.Value = 0
    }

    if ($exitCode -eq 1) {
        # 変換を続けられないエラー（変換対象フォルダが無い など）
        $message = (readTextShared ${convertErrorFile}).Trim()
        if ($message -eq "") {
            $message = "詳しくはログを確認してください。"
        }
        $ui.ConvertProgressText.Text = "変換できませんでした"
        $ui.ConvertProgressDetail.Text = $message
        setStatus "変換できませんでした：$message"
        showMessage "変換できませんでした。`n`n$message" "OK" "Error" | Out-Null
    } elseif ($exitCode -eq 2) {
        $ui.ConvertProgressText.Text = if ($counts) { "変換を中止しました（$counts）" } else { "変換を中止しました" }
        $ui.ConvertProgressDetail.Text = "次回は続きから再開できます。"
        setStatus $ui.ConvertProgressText.Text
    } elseif ($progress -and $progress.Processed -gt 0) {
        $ui.ConvertProgressText.Text = "変換が終わりました（$counts）"
        $ui.ConvertProgressDetail.Text = if ($progress.Failed -gt 0) { "失敗したファイルと原因は「変換に失敗したファイル」の一覧で確認できます。" } else { "" }
        setStatus $ui.ConvertProgressText.Text
    } else {
        $ui.ConvertProgressText.Text = "変換が必要なファイルはありませんでした"
        $ui.ConvertProgressDetail.Text = ""
        setStatus $ui.ConvertProgressText.Text
    }
    $ui.ConvertProgressEta.Text = ""
    $ui.ConvertStopButton.Visibility = "Collapsed"
    $ui.ConvertLogButton.Visibility = if (Test-Path -LiteralPath ${convertLogFile}) { "Visible" } else { "Collapsed" }

    $script:convertProcess = $null
    $script:sourceFolderMaps = @{}
    refreshConversionState
    refreshIndexSummary
    loadIndexTree  # 新しいインデックス・フォルダをツリーに出す
    updateKillBadge
}
$script:convertTimer = newTimer 1000 { safe { updateConversionProgress } }

# ---- イベント ----

$ui.NewIndexButton.Add_Click({ safe { newIndex } })
$ui.EditIndexButton.Add_Click({ safe { editIndex } })
$ui.RebuildIndexButton.Add_Click({ safe { rebuildIndex } })
$ui.RemoveIndexButton.Add_Click({ safe { deleteIndex } })
$ui.IndexGrid.Add_SelectionChanged({ safe { updateIndexListView } })
$ui.IndexGrid.Add_MouseDoubleClick({ safe { editIndex } })
$ui.IndexGrid.Add_KeyDown({
    param ($sender, $e)
    if ($e.Key -eq "Delete") {
        safe { deleteIndex }
    }
})
$ui.IndexGrid.Add_PreviewDragOver({ onFolderDragOver @args })
$ui.IndexGrid.Add_PreviewDrop({
    param ($sender, $e)
    safe {
        foreach ($folder in (getDroppedFolders $e)) {
            addIndexForFolder $folder
        }
    }
    $e.Handled = $true
})
$ui.FailedGrid.Add_MouseDoubleClick({
    param ($sender, $e)
    # 行の上でのダブルクリックだけを対象にする（列見出し・スクロールバーは除く）
    $element = $e.OriginalSource
    while ($element -and !($element -is [System.Windows.Controls.DataGridRow])) {
        if ($element -is [System.Windows.Controls.Primitives.DataGridColumnHeader] -or $element -is [System.Windows.Controls.Primitives.ScrollBar]) {
            return
        }
        $element = [System.Windows.Media.VisualTreeHelper]::GetParent($element)
    }
    if ($element) {
        safe { openFailedFileFolder }
    }
})
$ui.ConvertButton.Add_Click({ safe { startConversion } })
$ui.ConvertStopButton.Add_Click({ safe { stopConversion } })
$ui.ConvertLogButton.Add_Click({
    safe {
        if (Test-Path -LiteralPath ${convertLogFile}) {
            Invoke-Item -LiteralPath ${convertLogFile}
        }
    }
})

# ============================================================================
# ［2 検索］
# ============================================================================

$script:hitRows = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$ui.ResultGrid.ItemsSource = $script:hitRows
$script:hitView = [System.Windows.Data.CollectionViewSource]::GetDefaultView($script:hitRows)
# 表示用（強調セグメント・DisplayLine・セル列）は、行が画面に出るときだけ作る（件数が多くても軽い）。
# HitRow.Prepare は1回だけ実行し、作った値は PropertyChanged で反映する
$ui.ResultGrid.Add_LoadingRow({
    param ($s, $e)
    if ($e.Row.Item -is [HitRow]) { $e.Row.Item.Prepare() }
})
$script:search = $null
$script:lastSearch = $null
$script:sourceFolderMaps = @{}  # インデックスのフォルダ → インデックス名と変換対象フォルダの対応（getSourceLocation のキャッシュ）
$script:filterText = ""

# 別スレッドで実行する検索（結果は $shared.Queue に少しずつ入れる）
${searchScript} = {
    param ($commonPath, $word, $simpleMatch, $folders, $limit, $shared)
    try {
        . $commonPath
        # TSV が多いと数え上げだけで数秒かかるため、途中の件数を画面に伝える（止まって見えないように）
        $index = getIndexTsvFiles $folders { param ($count) $shared.Scanned = $count }
        $shared.Folders = $index.Folders
        $shared.Total = $index.Files.Count
        $shared.IndexTotal = $index.Files.Count
        # 検索条件（大文字・小文字の区別・対象ファイル）は startSearch が $shared に入れる
        $result = searchIndex $word $index.Files $simpleMatch $limit 50 -caseSensitive $shared.CaseSensitive -fileFilter $shared.FileFilter {
            param ($done, $total, $newHits)
            foreach ($hit in $newHits) {
                $shared.Queue.Enqueue($hit)
            }
            $shared.Total = $total
            $shared.Done = $done
        } { $shared.Stop }
        $shared.Total = $result.Total
        $shared.Truncated = $result.Truncated
        $shared.Cancelled = $result.Cancelled
    } catch {
        $shared.Error = $_.Exception.Message
    } finally {
        $shared.Finished = $true
    }
}

function getWordText {
    return $ui.WordBox.Text.Trim()
}

function getSearchOptionFromUi {
    # 画面の検索条件を readSearchOption と同じ形で返す
    return @{
        UseRegex      = [bool]$ui.RegexCheck.IsChecked
        CaseSensitive = [bool]$ui.CaseCheck.IsChecked
        FileFilter    = $ui.FileFilterBox.Text.Trim()
    }
}

function setSearchOptionToUi {
    param (
        [hashtable]$option
    )

    $ui.RegexCheck.IsChecked = [bool]$option.UseRegex
    $ui.CaseCheck.IsChecked = [bool]$option.CaseSensitive
    $ui.FileFilterBox.Text = [string]$option.FileFilter
}

function describeSearchOption {
    # 既定から変えた検索条件を「大文字と小文字を区別・対象ファイル：*.xlsx」のように返す（無ければ空）
    param (
        [hashtable]$option
    )

    $items = @()
    if ($option.CaseSensitive) {
        $items += "大文字と小文字を区別"
    }
    if ($option.FileFilter) {
        $items += "対象ファイル：$($option.FileFilter)"
    }
    return ($items -join "・")
}

function updateWordNotice {
    $word = getWordText
    if ($ui.RegexCheck.IsChecked -and $word -ne "" -and !(isValidRegex $word)) {
        $ui.WordNotice.Text = "正規表現として不正なため、文字どおり検索します。"
        $ui.WordNotice.Visibility = "Visible"
    } else {
        $ui.WordNotice.Visibility = "Collapsed"
    }
    updateSearchButton
}

function updateSearchButton {
    if ($script:search) {
        $ui.SearchButton.Content = "中止"
        $ui.SearchButton.IsEnabled = !$script:search.Shared.Stop
        return
    }
    $ui.SearchButton.Content = "検索"
    $noIndex = $script:indexSummary -and $script:indexSummary["Count"] -eq 0
    $ui.SearchButton.IsEnabled = (getWordText) -ne "" -and !$noIndex -and @(getSearchTargets).Count -gt 0
}

function updateSearchTarget {
    $targets = @(getSearchTargets)
    $summary = $script:indexSummary
    if ($summary -and $summary["Count"] -eq 0) {
        $ui.SearchTargetText.Text = "検索対象：なし（インデックスがありません。先に［1 インデックス管理］で作成してください）"
    } elseif ($targets.Count -eq 0) {
        $ui.SearchTargetText.Text = "検索対象：なし（左の一覧で、検索するインデックス・フォルダにチェックを付けてください）"
    } elseif (!(isAllIndexChecked)) {
        $ui.SearchTargetText.Text = "検索対象：$(describeSearchTargets $targets)"
    } elseif ($null -eq $summary) {
        $ui.SearchTargetText.Text = "検索対象：すべて（確認中…）"
    } else {
        $ui.SearchTargetText.Text = "検索対象：すべて（TSV $($summary['Count'].ToString('N0')) 件 ・ 最終変換 $(formatTime $summary['LastWrite'])）"
    }
    $ui.SearchTargetText.ToolTip = $ui.SearchTargetText.Text
    $ui.GoIndexTabButton.Visibility = if ($summary -and $summary["Count"] -eq 0) { "Visible" } else { "Collapsed" }
    updateSearchButton
}

function startSearch {
    if ($script:search) {
        cancelSearch
        return
    }
    $word = getWordText
    if ($word -eq "") {
        setStatus "検索ワードを入力してください。"
        return
    }

    $option = getSearchOptionFromUi
    writeSearchOption $option
    $useRegex = $option.UseRegex
    # 一致箇所の強調にも、検索と同じ正規表現を使う
    $searchRegex = newSearchRegex $word (!$useRegex) $option.CaseSensitive
    $simpleMatch = $searchRegex.SimpleMatch
    $pattern = $searchRegex.Regex

    $ui.FilterBox.Text = ""
    $script:filterText = ""
    $script:hitView.Filter = $null
    $script:hitRows.Clear()
    $ui.DetailPanel.Visibility = "Collapsed"
    # 検索対象ツリーでチェックしたフォルダだけを検索する（結果の相対パスは、インデックスのフォルダからのまま）
    $folders = @(getSearchTargets)
    if ($folders.Count -eq 0) {
        setStatus "検索するフォルダに、左の「検索対象」でチェックを付けてください。"
        return
    }
    $ui.IndexColumn.Visibility = if (@($folders | ForEach-Object { (splitIndexRelPath ([string]$_.RelPath)).Name } | Sort-Object -Unique).Count -gt 1) { "Visible" } else { "Collapsed" }

    $shared = [hashtable]::Synchronized(@{
        Queue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
        Stop = $false; Finished = $false; Done = 0; Total = -1; IndexTotal = -1; Folders = $null; Scanned = 0
        Truncated = $false; Cancelled = $false; Error = $null
        CaseSensitive = $option.CaseSensitive; FileFilter = $option.FileFilter
    })
    $ps = [powershell]::Create()
    [void]$ps.AddScript(${searchScript}.ToString())
    foreach ($argument in @(${commonPath}, $word, $simpleMatch, $folders, ${searchLimit}, $shared)) {
        [void]$ps.AddArgument($argument)
    }
    $script:search = @{
        PS = $ps; Handle = $ps.BeginInvoke(); Shared = $shared
        Word = $word; Pattern = $pattern; SimpleMatch = $simpleMatch; UseRegex = $useRegex; Option = $option; Start = Get-Date
    }

    $ui.SearchProgress.Visibility = "Visible"
    $ui.SearchProgress.IsIndeterminate = $true
    $ui.SummaryText.Text = "検索中…"
    $taskbar.ProgressState = "Indeterminate"
    setStatus "検索しています：${word}"
    updateSearchButton
    $script:searchTimer.Start()
}

function cancelSearch {
    if ($script:search) {
        $script:search.Shared.Stop = $true
        $ui.SummaryText.Text = "中止しています…"
        updateSearchButton
    }
}

function pumpSearch {
    # 検索スレッドの結果を表に移し、進み具合を表示する
    $s = $script:search
    if ($null -eq $s) {
        $script:searchTimer.Stop()
        return
    }
    $shared = $s.Shared
    $hit = $null
    $added = 0
    while ($added -lt 3000 -and $shared.Queue.TryDequeue([ref]$hit)) {
        # インデックスのフォルダ（work\index）からの相対パスの先頭がインデックス名
        $script:hitRows.Add([HitRow]::Create((splitIndexRelPath ([string]$hit.RelDir)).Name, $hit.Root, $hit.RelPath, $hit.RelDir, $hit.FileName,
                $hit.Book, $hit.Location, [int]$hit.LineNumber, $hit.Line, $s.Word, $s.Pattern))
        $added++
    }

    if ($shared.Total -gt 0) {
        $ratio = $shared.Done / $shared.Total
        $ui.SearchProgress.IsIndeterminate = $false
        $ui.SearchProgress.Value = $ratio
        $taskbar.ProgressState = "Normal"
        $taskbar.ProgressValue = $ratio
        if (!$shared.Stop) {
            $ui.SummaryText.Text = "検索中… $($shared.Done.ToString('N0')) / $($shared.Total.ToString('N0')) ファイル（$($script:hitRows.Count.ToString('N0')) 件）"
        }
    } elseif ($shared.Total -lt 0 -and !$shared.Stop) {
        # 数え上げの途中。件数が増えていくのが見えれば、止まっていないことが分かる
        $scanned = [int]$shared.Scanned
        $ui.SummaryText.Text = if ($scanned -gt 0) {
            "検索対象のファイルを確認しています…（$($scanned.ToString('N0')) 件）"
        } else {
            "検索対象のファイルを確認しています…"
        }
    }

    if ($shared.Finished -and $shared.Queue.IsEmpty) {
        finishSearch
    }
}

function finishSearch {
    $s = $script:search
    $script:search = $null
    $script:searchTimer.Stop()
    try {
        [void]$s.PS.EndInvoke($s.Handle)
    } catch {
    }
    $s.PS.Dispose()

    $shared = $s.Shared
    $ui.SearchProgress.Visibility = "Collapsed"
    $taskbar.ProgressState = if (isConverting) { $taskbar.ProgressState } else { "None" }
    $script:lastSearch = $s
    $seconds = ((Get-Date) - $s.Start).TotalSeconds
    updateSearchButton

    if ($shared.Error) {
        $ui.SummaryText.Text = "検索できませんでした"
        setStatus "検索できませんでした：$($shared.Error)"
        return
    }

    $count = $script:hitRows.Count
    $files = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $script:hitRows) {
        [void]$files.Add("$($row.Root)\$($row.RelDir)\$($row.Book)")
    }

    if ($shared.IndexTotal -gt 0 -and $shared.Total -eq 0) {
        $ui.SummaryText.Text = "対象ファイル（$($s.Option.FileFilter)）に一致するファイルがありません。"
    } elseif ($shared.Total -eq 0) {
        $ui.SummaryText.Text = "検索対象の TSV がありません。先にインデックスを作成してください。"
    } elseif ($count -eq 0) {
        $text = "見つかりませんでした。"
        if (!$s.UseRegex -and $s.Word -match '[\\()\[\]{}.*+?^$|]') {
            $text += "（正規表現として探す場合は［正規表現を使う］をオンにしてください）"
        } elseif (describeSearchOption $s.Option) {
            $text += "（検索条件：$(describeSearchOption $s.Option)）"
        }
        $ui.SummaryText.Text = $text
    } else {
        $ui.SummaryText.Text = "$($count.ToString('N0')) 件（$($files.Count.ToString('N0')) ファイル） ・ $($seconds.ToString('0.0')) 秒"
    }

    $status = "検索しました（$($s.Word)：$($count.ToString('N0')) 件）"
    if (describeSearchOption $s.Option) {
        $status = "検索しました（$($s.Word)：$($count.ToString('N0')) 件　条件：$(describeSearchOption $s.Option)）"
    }
    if ($shared.Truncated) {
        # パスの順に検索して打ち切るため、この先のファイルのヒットは結果に出ない。そのことが分かる文面にする
        $status = "$(${searchLimit}.ToString('N0')) 件を超えたため、ここで打ち切りました。この先のファイルは検索していないため、ワード・対象ファイル・検索対象で絞り込んでください。"
    } elseif ($shared.Cancelled) {
        $status = "中止しました（$($count.ToString('N0')) 件まで表示）"
    }
    $missing = @($shared.Folders | Where-Object { !$_.Exists } | ForEach-Object { $_.Path })
    if ($missing.Count -gt 0) {
        $status += "　見つからない検索対象フォルダ：$($missing -join '、')"
    }
    if (isConverting) {
        $status += "　変換中のため、作成途中のインデックスを検索しています。"
    }
    setStatus $status
}

$script:searchTimer = newTimer 100 { safe { pumpSearch } }

# ---- 絞り込み・選択行の詳細 ----

function applyFilter {
    $script:filterText = $ui.FilterBox.Text.Trim()
    if ($script:filterText -eq "") {
        $script:hitView.Filter = $null
    } else {
        $script:hitView.Filter = [Predicate[object]] { param ($row) $row.Contains($script:filterText) }
    }
    if ($script:lastSearch -and !$script:search -and $script:hitRows.Count -gt 0) {
        $shown = 0
        foreach ($row in $script:hitView) {
            $shown++
        }
        if ($script:filterText -eq "") {
            finishSummaryText
        } else {
            $ui.SummaryText.Text = "$($script:hitRows.Count.ToString('N0')) 件中 $($shown.ToString('N0')) 件を表示"
        }
    }
}

function finishSummaryText {
    $files = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $script:hitRows) {
        [void]$files.Add("$($row.Root)\$($row.RelDir)\$($row.Book)")
    }
    $ui.SummaryText.Text = "$($script:hitRows.Count.ToString('N0')) 件（$($files.Count.ToString('N0')) ファイル）"
}

$script:filterTimer = newTimer 300 {
    $script:filterTimer.Stop()
    safe { applyFilter }
}

# 選択行のプレビューは、↑↓で続けて選択が変わったときは最後の1回だけ読む（巨大なTSVでも操作が重くならないようにする）
$script:detailTimer = newTimer 120 {
    $script:detailTimer.Stop()
    safe { showDetail }
}

function getViewRows {
    # 表示中（絞り込み・並べ替え後）の行
    $rows = New-Object System.Collections.ArrayList
    foreach ($row in $script:hitView) {
        [void]$rows.Add($row)
    }
    return , $rows.ToArray()
}

function getSelectedRows {
    # 選択行を表示の順に並べて返す
    $rows = New-Object System.Collections.ArrayList
    foreach ($row in $ui.ResultGrid.SelectedItems) {
        [void]$rows.Add($row)
    }
    return , @($rows.ToArray() | Sort-Object { $ui.ResultGrid.Items.IndexOf($_) })
}

function showDetail {
    $row = $ui.ResultGrid.SelectedItem
    if ($null -eq $row) {
        $ui.DetailPanel.Visibility = "Collapsed"
        return
    }
    $path = if ($row.RelDir) { "$($row.RelDir)\$($row.Book)" } else { $row.Book }
    $place = if ($row.MatchCell) { "セル $($row.MatchCell)" } else { "$($row.LineNumber) 行目" }
    $ui.OpenButton.Content = if ($row.IsExcel) { "Excel で開く" } else { "開く" }

    # 前後の行をインデックスのTSVから読む（読めなければ選択行だけを出す）
    $context = @(readTsvContext ([System.IO.Path]::Combine($row.Root, $row.RelPath)) $row.LineNumber ${previewLines} ${previewLines})
    $table = $row.BuildPreview([int[]]@($context | ForEach-Object { $_.LineNumber }), [string[]]@($context | ForEach-Object { $_.Line }))

    $title = "${path} ・ $($row.Location) ・ ${place}"
    $ui.DetailTitle.Text = $title
    $ui.DetailTitle.ToolTip = $title

    # 横に長い行は一部の列だけを表示するため、その範囲を知らせる
    if ($table.TotalColumns -gt $table.ShownColumns) {
        $note = "表示は $($table.RangeLabel) の $($table.ShownColumns.ToString('N0')) 列（全 $($table.TotalColumns.ToString('N0')) 列）"
        $ui.PreviewNote.Text = $note
        $ui.PreviewNote.ToolTip = "$note　一致したセルを中心に表示しています。ほかの列は元のファイルで確認してください。"
        $ui.PreviewNote.Visibility = "Visible"
    } else {
        $ui.PreviewNote.Visibility = "Collapsed"
    }
    $ui.PreviewHeader.ItemsSource = $table.Columns
    $ui.PreviewRows.ItemsSource = $table.Rows
    $script:previewTable = $table
    $ui.DetailPanel.Visibility = "Visible"

    # 一致したセルが見えるよう横にスクロールする（左端から見えていればそのまま）
    $ui.PreviewScroll.UpdateLayout()
    $offset = 0
    if ($table.HitOffset + $table.HitWidth -gt $ui.PreviewScroll.ViewportWidth) {
        $offset = [math]::Max(0, $table.HitOffset - 120)
    }
    $ui.PreviewScroll.ScrollToHorizontalOffset($offset)
}

function getPreviewCell {
    # マウスの下（またはイベントの発生元）のプレビューのセル。セルの上でなければ $null
    param (
        $source
    )

    $element = $source
    while ($element) {
        if ($element -is [System.Windows.FrameworkElement] -and $element.DataContext -is [PreviewCell]) {
            return $element.DataContext
        }
        $element = [System.Windows.Media.VisualTreeHelper]::GetParent($element)
    }
    return $null
}

function copyPreviewSelection {
    # プレビューで選んだセルの値をクリップボードに入れる（1 セルならその値のまま、複数ならタブ区切り）
    if ($null -eq $script:previewTable -or !$script:previewTable.HasSelection()) {
        setStatus "プレビューでコピーするセルをクリックしてください（Shift＋クリック・ドラッグで複数選べます）。"
        return
    }
    $text = $script:previewTable.GetSelectionText()
    if ($text -eq "") {
        [System.Windows.Clipboard]::Clear()
    } else {
        [System.Windows.Clipboard]::SetText($text)
    }
    $count = $script:previewTable.SelectedCount()
    if ($count -le 1) {
        setStatus "セルの値をコピーしました：$(toStatusText $text)"
    } else {
        setStatus "${count} 個のセルをコピーしました（Excel に貼り付けると、元の位置に並びます）"
    }
}

function toStatusText {
    # ステータスに出す短い文字列（改行・タブはスペースにし、長ければ末尾を省略）
    param (
        [string]$text
    )

    $text = ($text -replace "[`r`n`t]+", " ")
    if ($text.Length -gt 40) {
        return $text.Substring(0, 40) + "…"
    }
    return $text
}

# ---- 元のファイルを開く・コピー・出力 ----

function getSourcePath {
    # 元のファイルのパス（ファイルがあるかは確かめない）。元の場所が分からなければ $null
    param (
        $row
    )

    return resolveSourcePath $row $script:sourceFolderMaps
}

function getExistingFolder {
    # path の上のフォルダのうち、存在する最も深いフォルダ（フォルダ選択の初期位置）。無ければ空
    param (
        [string]$path
    )

    $dir = Split-Path $path -Parent
    while ($dir) {
        if (Test-Path -LiteralPath $dir -PathType Container) {
            return $dir
        }
        $dir = Split-Path $dir -Parent
    }
    return ""
}

function findSourceFile {
    # 元のファイルのパスを返す。見つからなければ、元のファイルのあるフォルダを選んでもらって探し、
    # 見つかればそのインデックスの元のフォルダ（今の置き場所）として設定に記録する
    # （インデックス名に対して 1 か所を記録するため、同じインデックスのほかのファイルも次からそのまま開ける）。
    # 見つからない・選ばなかった場合は $null
    param (
        $row
    )

    $location = getSourceLocation $row $script:sourceFolderMaps
    $relPath = if ($location.Rest) { "$($location.Rest)\$($row.Book)" } else { $row.Book }
    if ($location.Known) {
        $path = joinSourcePath $location.Folder $location.Rest $row.Book
        if (Test-Path -LiteralPath (toLongPath $path) -PathType Leaf) {
            return $path
        }
        # 記録した場所に無くても、書き方が違うだけで同じ場所を指すパスで開けることがある
        # （ネットワークドライブと UNC パス）。別の PC でドライブの割り当てが違う場合に、聞かずに開けるようにする
        foreach ($alias in @(getFolderPathAliases $location.Folder | Select-Object -Skip 1)) {
            $candidate = joinSourcePath $alias $location.Rest $row.Book
            if (!(Test-Path -LiteralPath (toLongPath $candidate) -PathType Leaf)) {
                continue
            }
            if ($location.Name) {
                setIndexSourceFolder $location.Name $alias
                $script:sourceFolderMaps = @{}
                setStatus "インデックス [$($location.Name)] の元のフォルダを ${alias} に変えました"
            }
            return $candidate
        }
        $message = "元のファイルが見つかりません。`n${path}`n`n" +
                   "インデックスを別の PC に持ってきた場合や、フォルダを移した場合は、今の場所のフォルダを選ぶと開けます。"
        $description = "「$($location.Folder)」に当たるフォルダ（または $($row.Book) のあるフォルダ）を選んでください"
        $initial = getExistingFolder $path
    } else {
        $path = "$($row.Root)\$($row.RelDir)\$($row.Book)"
        $message = "元のファイルの場所が分かりません（インデックス [$($location.Name)] の元のフォルダが記録されていません）。`n${relPath}`n`n" +
                   "元のファイルのあるフォルダを選ぶと開けます。"
        $description = "$($row.Book) のあるフォルダ（またはインデックス [$($location.Name)] の元のフォルダ）を選んでください"
        $initial = ""
    }
    $message += "`n（選んだフォルダはインデックス [$($location.Name)] の元のフォルダとして記録し、同じインデックスのほかのファイルも開けるようにします）`n`nフォルダを選びますか？"

    while ($true) {
        if ((showMessage $message "YesNo" "Question" "Yes") -ne "Yes") {
            setStatus "元のファイルが見つかりません：${path}"
            return $null
        }
        $picked = selectFolder $description $initial
        if (!$picked) {
            setStatus "元のファイルが見つかりません：${path}"
            return $null
        }

        $found = findMovedSource $picked $location.Rest $row.Book
        if ($found) {
            if ($found.Root -and $location.Name) {
                setIndexSourceFolder $location.Name $found.Root
                $script:sourceFolderMaps = @{}
                setStatus "インデックス [$($location.Name)] の元のフォルダを $($found.Root) に変えました"
            } else {
                setStatus "開きました：$($found.Path)"
            }
            return $found.Path
        }
        $message = "選んだフォルダの中に、元のファイルが見つかりませんでした。`n選んだフォルダ：${picked}`n探したファイル：${relPath}`n`n別のフォルダを選びますか？"
        $initial = $picked
    }
}

function openWithShell {
    # ファイルを既定のアプリで開く。開き方（mode）は、エクスプローラーの右クリックメニューと同じ動詞で行う。
    #   読み取り専用 → OpenAsReadOnly、新規 → New（元のファイルを基にした無題の文書。元のファイルを占有しない）
    # その動詞が登録されていない種類のファイルは、そのまま開いて $false を返す
    param (
        [string]$path,
        [string]$mode
    )

    $verb = switch ($mode) {
        ${openModeReadOnly} { "OpenAsReadOnly" }
        ${openModeNew}      { "New" }
        default             { $null }
    }
    if ($verb) {
        # Start-Process は [ ] をワイルドカードとして扱うため使わない。
        # ProcessStartInfo.Verbs は New を一覧に含めないため、実行してみて、無ければ例外で分かる
        $info = New-Object System.Diagnostics.ProcessStartInfo($path)
        $info.UseShellExecute = $true
        $info.Verb = $verb
        try {
            [void][System.Diagnostics.Process]::Start($info)
            return $true
        } catch [System.ComponentModel.Win32Exception] {
            # その動詞が登録されていない
        }
    }
    Invoke-Item -LiteralPath $path
    return (-not $verb)
}

function openInExcel {
    # 表示中の Excel（無ければ新しく起動）でブックを開き、該当シートの該当セルを選択する。
    # 開き方（mode）: 通常 = そのまま開く / 読み取り専用 = ReadOnly で開く /
    #                 新規 = 元のファイルを基にした新しいブック（無題）として開く（元のファイルを占有しない）
    param (
        [string]$path,
        [string]$location,
        [string]$cell,
        [string]$mode = ${openModeNormal}
    )

    $excel = $null
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        # 変換処理がバックグラウンドで使っている Excel は使わない
        if (!$excel.Visible) {
            $excel = $null
        }
    } catch {
        $excel = $null
    }
    if ($null -eq $excel) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.UserControl = $true
    }

    $book = $null
    if ($mode -eq ${openModeNew}) {
        # 元のファイルをテンプレートとして新しいブックを作る（読み込んだ後は元のファイルを開いたままにしない）
        $book = $excel.Workbooks.Add($path)
    } else {
        # 既に開いているブックがあれば、そのまま使う（同じブックを二重に開けないため。開き方も既に開いたときのまま）
        foreach ($openBook in $excel.Workbooks) {
            if ($openBook.FullName -eq $path) {
                $book = $openBook
                break
            }
        }
        if ($null -eq $book) {
            #   引数: Filename, UpdateLinks, ReadOnly
            $book = $excel.Workbooks.Open($path, [Type]::Missing, ($mode -eq ${openModeReadOnly}))
        }
    }
    $book.Activate()

    # 場所はシート名。名前が同じシートを選ぶ。
    # 以前の版のインデックスは、ファイル名に使えない文字を全角に置き換えてあるため、同じ名前のシートが無ければ
    # 全角に置き換えて一致するシートを選ぶ（`衝突"` と `衝突”` のように、置き換えると重なるシートがあるため、同じ名前を優先する）
    $target = $null
    $sameSafeName = $null
    foreach ($sheet in $book.Worksheets) {
        if ($sheet.Name -eq $location) {
            $target = $sheet
            break
        }
        if ($null -eq $sameSafeName -and (toSafeFileName $sheet.Name) -eq $location) {
            $sameSafeName = $sheet
        }
    }
    if ($null -eq $target) {
        $target = $sameSafeName
    }
    if ($null -ne $target) {
        $target.Activate()
        if ($cell) {
            $target.Range($cell).Select()
        }
    }
    if ($excel.WindowState -eq -4140) {
        # 最小化されていれば元に戻す（xlMinimized → xlNormal）
        $excel.WindowState = -4143
    }
    # Excel は Visible にして前面に出す（P/Invoke の SetForegroundWindow は実行時コンパイル（csc.exe）を無くすため廃止）
    $excel.Visible = $true
    try { $excel.ActiveWindow.Activate() } catch { }
}

function getOpenMode {
    # ダブルクリック・Enter・［開く］での開き方（［開き方］の選択。${openModes} のいずれか）
    $item = $ui.OpenModeCombo.SelectedItem
    if ($null -ne $item -and (${openModes} -contains $item.Tag)) {
        return [string]$item.Tag
    }
    return ${openModeNormal}
}

function setOpenMode {
    # ［開き方］の選択を設定の値に合わせる（起動時。選んだことにはしないため、設定は保存しない）
    param (
        [string]$mode
    )

    $script:loadingOpenMode = $true
    try {
        foreach ($item in $ui.OpenModeCombo.Items) {
            if ($item.Tag -eq $mode) {
                $ui.OpenModeCombo.SelectedItem = $item
                return
            }
        }
        $ui.OpenModeCombo.SelectedIndex = 0
    } finally {
        $script:loadingOpenMode = $false
    }
}

function updateOpenMenu {
    # 右クリックメニューは3つの開き方をすべて出し、既定の開き方（ダブルクリック・Enter と同じ）に Enter を表示する
    $mode = getOpenMode
    $ui.MenuOpen.InputGestureText         = $(if ($mode -eq ${openModeNormal})   { "Enter" } else { "" })
    $ui.MenuOpenReadOnly.InputGestureText = $(if ($mode -eq ${openModeReadOnly}) { "Enter" } else { "" })
    $ui.MenuOpenNew.InputGestureText      = $(if ($mode -eq ${openModeNew})      { "Enter" } else { "" })
}

function openSource {
    # 選択行の元のファイルを開く。mode で開き方（通常・読み取り専用・新規）を指定する
    param (
        [string]$mode = (getOpenMode)
    )

    $row = $ui.ResultGrid.SelectedItem
    if ($null -eq $row) {
        return
    }
    $path = findSourceFile $row
    if (!$path) {
        return
    }
    $how = switch ($mode) {
        ${openModeReadOnly} { "読み取り専用で開きました" }
        ${openModeNew}      { "新規で開きました" }
        default             { "開きました" }
    }
    $fallback = switch ($mode) {
        ${openModeReadOnly} { "読み取り専用で開けなかったため、元のファイルを開きました" }
        default             { "新規で開けなかったため、元のファイルを開きました" }
    }

    if ($row.IsExcel) {
        setStatus "Excel で開いています：${path}"
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            openInExcel $path $row.Location $row.MatchCell $mode
            setStatus "${how}：${path}"
        } catch {
            # Excel を操作できない場合（ダイアログを表示中など）は、ファイルを開くだけにする
            if (openWithShell $path $mode) {
                setStatus "${how}（該当セルへの移動はできませんでした）：${path}"
            } else {
                setStatus "${fallback}（該当セルへの移動はできませんでした）：${path}"
            }
        } finally {
            $window.Cursor = $null
        }
    } elseif (openWithShell $path $mode) {
        setStatus "${how}：${path}"
    } else {
        setStatus "${fallback}：${path}"
    }
}

function openSourceFolder {
    $row = $ui.ResultGrid.SelectedItem
    if ($null -eq $row) {
        return
    }
    $path = findSourceFile $row
    if (!$path) {
        return
    }
    Start-Process -FilePath "explorer.exe" -ArgumentList "/select,`"${path}`""
}

function copySelectedRows {
    $rows = getSelectedRows
    if ($rows.Count -eq 0) {
        return
    }
    $result = toSearchResultLines $rows
    $text = ((@($result.Header) + $result.Lines.ToArray()) -join "`r`n") + "`r`n"
    [System.Windows.Clipboard]::SetText($text)
    setStatus "$($rows.Count) 行をコピーしました（Excel に貼り付けると、元の列の位置に並びます）"
}

function copySourcePath {
    $row = $ui.ResultGrid.SelectedItem
    if ($null -eq $row) {
        return
    }
    $path = getSourcePath $row
    if (!$path) {
        $path = "$($row.Root)\$($row.RelPath)"
    }
    [System.Windows.Clipboard]::SetText($path)
    setStatus "パスをコピーしました：${path}"
}

function exportResults {
    if ($null -eq $script:lastSearch) {
        setStatus "先に検索してください。"
        return
    }
    $rows = getViewRows
    try {
        [System.IO.Directory]::CreateDirectory(${workDir}) | Out-Null
        $writer = New-Object System.IO.StreamWriter(${resultFile}, $false, ${utf8Bom})
        try {
            writeSearchResult $writer $script:lastSearch.Word $rows
        } finally {
            $writer.Close()
        }
    } catch [System.IO.IOException] {
        setStatus "検索結果.txt に書き込めません。開いているアプリを閉じてから、もう一度出力してください。"
        return
    }
    Invoke-Item -LiteralPath ${resultFile}
    if ($rows.Count -lt $script:hitRows.Count) {
        setStatus "絞り込み後の $($rows.Count.ToString('N0')) 件を検索結果.txt に出力しました"
    } else {
        setStatus "検索結果.txt に出力しました（$($rows.Count.ToString('N0')) 件）"
    }
}

# ---- 検索対象インデックスのツリー ----

$script:indexRoots = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$ui.IndexTree.ItemsSource = $script:indexRoots

function loadIndexTree {
    # インデックスの一覧（getSearchIndexes）をツリーに読み込む。一番上の項目がインデックス 1 件で、
    # ［1 インデックス管理］で作ったインデックスがすべて並ぶ。
    # 保存したチェックなしのフォルダと、読み込み前の展開の状態は戻す
    $expanded = New-Object 'System.Collections.Generic.List[string]'
    foreach ($node in $script:indexRoots) {
        $node.AddExpanded($expanded)
    }

    $script:indexRoots.Clear()
    $root = ${indexDir}
    if (Test-Path -LiteralPath $root -PathType Container) {
        $root = (Resolve-Path -LiteralPath $root).ProviderPath.TrimEnd("\")
    }
    foreach ($index in @(getSearchIndexes)) {
        $sourcePath = if ($index.SourcePath) { $index.SourcePath } else { $null }
        $script:indexRoots.Add([IndexNode]::CreateRoot($root, $index.Name, $index.Name, $sourcePath))
    }
    $ui.IndexTreePlaceholder.Visibility = if ($script:indexRoots.Count -eq 0) { "Visible" } else { "Collapsed" }

    foreach ($exclude in @(readSearchExcludes)) {
        foreach ($node in $script:indexRoots) {
            $node.ApplyExclude($exclude.Path, $exclude.Subfolders)
        }
    }
    foreach ($node in $script:indexRoots) {
        foreach ($path in $expanded) {
            $found = $node.Find($path)
            if ($found) {
                $found.SetExpanded($true)
            }
        }
    }
    updateSearchTarget
}

function getSearchTargets {
    # 検索対象ツリーでチェックしたフォルダ（SearchTarget の配列。getIndexTsvFiles に渡す）
    $targets = New-Object 'System.Collections.Generic.List[SearchTarget]'
    foreach ($node in $script:indexRoots) {
        $node.AddTargets($targets)
    }
    return $targets.ToArray()
}

function isAllIndexChecked {
    return @($script:indexRoots | Where-Object { $_.IsChecked -ne $true }).Count -eq 0
}

function describeSearchTargets {
    # 検索対象の表示（先頭の 3 件まで）。インデックスのフォルダ（work\index）からの相対パスは
    # 「インデックス名\その下のフォルダ」のため、そのまま表示に使う
    param (
        [object[]]$targets
    )

    $names = @($targets | ForEach-Object {
        $target = $_
        $name = ([string]$target.RelPath).Trim("\")
        if (!$target.Recurse) {
            $name += "（直下のファイル）"
        }
        $name
    })
    if ($names.Count -gt 3) {
        return "$($names[0..2] -join '、') ほか $($names.Count - 3) か所"
    }
    return $names -join "、"
}

function saveSearchExcludes {
    # チェックなしのフォルダを設定に保存する。見つからないインデックスのフォルダ（ネットワークのドライブが切れているなど）の記録は残す
    $excludes = New-Object 'System.Collections.Generic.List[SearchExclude]'
    foreach ($node in $script:indexRoots) {
        $node.AddExcludes($excludes)
    }
    $roots = @($script:indexRoots | Where-Object { $_.Exists } | ForEach-Object { $_.FullPath.TrimEnd("\") })
    $kept = @(readSearchExcludes | Where-Object {
        $path = $_.Path
        @($roots | Where-Object { $path -eq $_ -or $path.StartsWith("$_\", [System.StringComparison]::OrdinalIgnoreCase) }).Count -eq 0
    })
    writeSearchExcludes (@($kept) + @($excludes.ToArray()))
}

function onIndexTreeChecked {
    saveSearchExcludes
    updateSearchTarget
}

function setAllIndexChecked {
    param (
        [bool]$checked
    )

    foreach ($node in $script:indexRoots) {
        $node.SetChecked($checked)
    }
    onIndexTreeChecked
}

# ---- イベント ----

$ui.WordBox.Add_TextChanged({ safe { updateWordNotice } })
$ui.WordBox.Add_PreviewKeyDown({
    param ($sender, $e)
    if ($e.Key -eq "Return") {
        safe {
            if (!$script:search) {
                startSearch
            }
        }
        $e.Handled = $true
    }
})
$ui.SearchButton.Add_Click({ safe { startSearch } })
$ui.RegexCheck.Add_Click({
    safe {
        writeSearchOption @{ UseRegex = [bool]$ui.RegexCheck.IsChecked }
        updateWordNotice
    }
})
$ui.CaseCheck.Add_Click({ safe { writeSearchOption @{ CaseSensitive = [bool]$ui.CaseCheck.IsChecked } } })
$ui.FileFilterBox.Add_TextChanged({
    $ui.FileFilterPlaceholder.Visibility = if ($ui.FileFilterBox.Text -eq "") { "Visible" } else { "Collapsed" }
})
# 対象ファイルは入力を終えたとき（フォーカスが外れたとき・検索したとき）に保存する
$ui.FileFilterBox.Add_LostFocus({ safe { writeSearchOption @{ FileFilter = $ui.FileFilterBox.Text.Trim() } } })
$ui.FileFilterBox.Add_KeyDown({
    param ($sender, $e)
    if ($e.Key -eq "Return") {
        safe {
            if (!$script:search) {
                startSearch
            }
        }
        $e.Handled = $true
    }
})
# ツリーのチェックボックスのクリック。チェックは OneWay バインドのため、クリックされたノードの Toggle() で
# 3状態（子・親への伝播）を反映してから保存する（PS class はセッターにロジックを書けないため、ここで行う）
$ui.IndexTree.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler] {
    param ($s, $e)
    safe {
        $cb = $e.OriginalSource
        if ($cb -is [System.Windows.Controls.CheckBox] -and $cb.DataContext -is [IndexNode]) {
            $cb.DataContext.Toggle()
            onIndexTreeChecked
        }
    }
})
# フォルダを展開したときに子を読み込む（IsExpanded は OneWay/プレーンなので、ここで LoadChildren する）
$ui.IndexTree.AddHandler([System.Windows.Controls.TreeViewItem]::ExpandedEvent, [System.Windows.RoutedEventHandler] {
    param ($s, $e)
    safe {
        $node = $e.OriginalSource.DataContext
        if ($node -is [IndexNode]) { $node.LoadChildren() }
    }
})
$ui.IndexTree.Add_PreviewKeyDown({
    param ($sender, $e)
    # スペースで選択中のフォルダのチェックを切り替える
    $node = $ui.IndexTree.SelectedItem
    if ($e.Key -eq "Space" -and $node -and !$node.IsPlaceholder) {
        safe {
            $node.Toggle()
            onIndexTreeChecked
        }
        $e.Handled = $true
    }
})
$ui.CheckAllIndexButton.Add_Click({ safe { setAllIndexChecked $true } })
$ui.UncheckAllIndexButton.Add_Click({ safe { setAllIndexChecked $false } })
$ui.GoIndexTabButton.Add_Click({ $ui.Tabs.SelectedItem = $ui.IndexTab })
$ui.FilterBox.Add_TextChanged({
    $ui.FilterPlaceholder.Visibility = if ($ui.FilterBox.Text -eq "") { "Visible" } else { "Collapsed" }
    $script:filterTimer.Stop()
    $script:filterTimer.Start()
})
$ui.ResultGrid.Add_SelectionChanged({
    $script:detailTimer.Stop()
    $script:detailTimer.Start()
})
$ui.ResultGrid.Add_MouseDoubleClick({
    param ($sender, $e)
    # 行の上でのダブルクリックだけを対象にする（列見出し・スクロールバーは除く）
    $element = $e.OriginalSource
    while ($element -and !($element -is [System.Windows.Controls.DataGridRow])) {
        if ($element -is [System.Windows.Controls.Primitives.DataGridColumnHeader] -or $element -is [System.Windows.Controls.Primitives.ScrollBar]) {
            return
        }
        $element = [System.Windows.Media.VisualTreeHelper]::GetParent($element)
    }
    if ($element) {
        safe { openSource }
    }
})
$ui.ResultGrid.Add_PreviewKeyDown({
    param ($sender, $e)
    if ($e.Key -eq "Return") {
        safe { openSource }
        $e.Handled = $true
    } elseif ($e.Key -eq "C" -and [System.Windows.Input.Keyboard]::Modifiers -eq "Control") {
        safe { copySelectedRows }
        $e.Handled = $true
    }
})
$ui.MenuOpen.Add_Click({ safe { openSource ${openModeNormal} } })
$ui.MenuOpenReadOnly.Add_Click({ safe { openSource ${openModeReadOnly} } })
$ui.MenuOpenNew.Add_Click({ safe { openSource ${openModeNew} } })
$ui.OpenModeCombo.Add_SelectionChanged({
    safe {
        # 起動時の読み込みでは保存しない（設定していない利用者の setting.config を作らないため）
        if (-not $script:loadingOpenMode) {
            writeOpenMode (getOpenMode)
        }
        updateOpenMenu
    }
})
$ui.MenuOpenFolder.Add_Click({ safe { openSourceFolder } })
$ui.OpenButton.Add_Click({ safe { openSource } })
$ui.OpenFolderButton.Add_Click({ safe { openSourceFolder } })
# プレビューのセルをクリックすると、その値をコピーできるように選ぶ（Shift＋クリック・ドラッグで範囲、Ctrl+C でコピー）
$script:previewTable = $null
$ui.PreviewRows.Add_PreviewMouseLeftButtonDown({
    param ($sender, $e)
    safe {
        $cell = getPreviewCell $e.OriginalSource
        if ($cell -and $script:previewTable) {
            $extend = [System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Shift
            $script:previewTable.Select($cell, [bool]$extend)
            [void]$ui.PreviewScroll.Focus()
        }
    }
})
$ui.PreviewRows.Add_MouseMove({
    param ($sender, $e)
    if ($e.LeftButton -ne "Pressed") {
        return
    }
    safe {
        $cell = getPreviewCell $e.OriginalSource
        if ($cell -and $script:previewTable) {
            $script:previewTable.Select($cell, $true)
        }
    }
})
$ui.PreviewRows.Add_PreviewMouseRightButtonDown({
    param ($sender, $e)
    safe {
        # 右クリックしたセルが選ばれていなければ、そのセルだけを選ぶ
        $cell = getPreviewCell $e.OriginalSource
        if ($cell -and $script:previewTable -and !$cell.IsSelected) {
            $script:previewTable.Select($cell, $false)
        }
    }
})
$ui.PreviewScroll.Add_PreviewKeyDown({
    param ($sender, $e)
    if ($e.Key -eq "C" -and [System.Windows.Input.Keyboard]::Modifiers -eq "Control") {
        safe { copyPreviewSelection }
        $e.Handled = $true
    }
})
$ui.MenuPreviewCopy.Add_Click({ safe { copyPreviewSelection } })
$ui.MenuPreviewCopyRow.Add_Click({
    safe {
        if ($script:previewTable -and $script:previewTable.HasSelection()) {
            # 選んでいるセルのある行をすべて選んでからコピーする
            $selected = $null
            foreach ($row in $script:previewTable.Rows) {
                $selected = @($row.Cells | Where-Object { $_.IsSelected })[0]
                if ($selected) {
                    break
                }
            }
            if ($selected) {
                $script:previewTable.SelectRow($selected)
            }
        }
        copyPreviewSelection
    }
})
# プレビューの列見出しの右端をドラッグすると、その列（PreviewColumn）の幅が変わる（各行のセルも同じ列を参照しているため一緒に変わる）
$ui.PreviewHeader.AddHandler(
    [System.Windows.Controls.Primitives.Thumb]::DragDeltaEvent,
    [System.Windows.Controls.Primitives.DragDeltaEventHandler] {
        param ($sender, $e)
        safe {
            $column = $e.OriginalSource.DataContext
            if ($column -is [PreviewColumn]) {
                $column.SetWidth($column.Width + $e.HorizontalChange)
            }
        }
    })
$ui.MenuCopy.Add_Click({ safe { copySelectedRows } })
$ui.MenuCopyPath.Add_Click({ safe { copySourcePath } })
$ui.ExportButton.Add_Click({ safe { exportResults } })

# ============================================================================
# ［9 プロセス停止］
# ============================================================================

$script:processes = @()

function refreshProcesses {
    $selectedIds = @($ui.ProcessGrid.SelectedItems | ForEach-Object { $_.Id })
    $script:processes = @(getOfficeProcesses | Sort-Object @{ Expression = { !$_.Background } }, @{ Expression = { $_.StartTime } })

    $rows = New-Object System.Collections.ArrayList
    foreach ($process in $script:processes) {
        $row = New-Object ProcRow
        $row.Id = $process.Id
        $row.AppName = $process.AppName
        $row.Background = $process.Background
        $row.StateText = if ($process.Background) { "⚠ バックグラウンド" } else { "画面に表示中" }
        $row.StartText = formatTime $process.StartTime
        $row.MemoryText = "$($process.MemoryMB.ToString('N0')) MB"
        $row.TitleText = if ($process.Title) { $process.Title } else { "（なし）" }
        [void]$rows.Add($row)
    }
    $ui.ProcessGrid.ItemsSource = $rows
    foreach ($row in $rows) {
        if ($selectedIds -contains $row.Id) {
            [void]$ui.ProcessGrid.SelectedItems.Add($row)
        }
    }

    $background = @($script:processes | Where-Object { $_.Background }).Count
    $visible = $script:processes.Count - $background
    if ($script:processes.Count -eq 0) {
        $ui.ProcessSummaryText.Text = "実行中の Excel・Word・PowerPoint はありません。（$(Get-Date -Format 'H:mm:ss') 時点）"
    } else {
        $ui.ProcessSummaryText.Text = "$($script:processes.Count) 件（バックグラウンド $background 件・画面に表示中 $visible 件） ・ $(Get-Date -Format 'H:mm:ss') 時点"
    }
    $ui.KillBackgroundButton.Content = "バックグラウンドのみ終了（$background 件）"
    $ui.KillBackgroundButton.IsEnabled = $background -gt 0
    $ui.KillAllButton.IsEnabled = $script:processes.Count -gt 0
    $ui.KillSelectedButton.IsEnabled = $ui.ProcessGrid.SelectedItems.Count -gt 0
    updateKillBadge $background
}

function updateKillBadge {
    param (
        $background = $null
    )

    if ($null -eq $background) {
        $background = @(getOfficeProcesses | Where-Object { $_.Background }).Count
    }
    # 変換中はバックグラウンドの Excel 等があって当然なので、印を付けない
    $ui.KillTabHeader.Text = if ($background -gt 0 -and !(isConverting)) { "⚠ 9 プロセス停止" } else { "9 プロセス停止" }
}

function killProcesses {
    param (
        [object[]]$targets
    )

    if ($targets.Count -eq 0) {
        return
    }

    $describe = {
        param ([object[]]$list)
        @(${officeProcessNames}.Values | ForEach-Object {
            $name = $_
            $count = @($list | Where-Object { $_.AppName -eq $name }).Count
            if ($count -gt 0) { "$name $count 件" }
        }) -join "・"
    }
    $visibleTargets = @($targets | Where-Object { !$_.Background })
    $message = if ($visibleTargets.Count -gt 0) {
        "画面に表示中の $(& $describe $visibleTargets) を含む $($targets.Count) 件を、保存せずに終了します。`n保存していない内容は失われます。よろしいですか？"
    } else {
        "バックグラウンドの $(& $describe $targets) を終了します。よろしいですか？"
    }
    $default = if ($visibleTargets.Count -gt 0) { "No" } else { "Yes" }
    if (isConverting) {
        $message = "変換中です。バックグラウンドのプロセスを終了すると、変換中のファイルは失敗扱いになります。`n`n" + $message
        $default = "No"
    }
    if ((showMessage $message "YesNo" "Warning" $default) -ne "Yes") {
        return
    }

    $results = @(stopOfficeProcesses @($targets | ForEach-Object { $_.Id }))
    $stopped = @($results | Where-Object { $_.Stopped }).Count
    $failures = @($results | Where-Object { !$_.Stopped } | ForEach-Object { "PID $($_.Id) を終了できませんでした：$($_.Message)" })
    setStatus (@("$stopped 個のプロセスを終了しました。") + $failures -join "　")
    Start-Sleep -Milliseconds 300
    refreshProcesses
}

$script:processTimer = newTimer 5000 { safe { refreshProcesses } }

$ui.RefreshProcessButton.Add_Click({ safe { refreshProcesses } })
$ui.ProcessGrid.Add_SelectionChanged({ $ui.KillSelectedButton.IsEnabled = $ui.ProcessGrid.SelectedItems.Count -gt 0 })
$ui.KillBackgroundButton.Add_Click({ safe { refreshProcesses; killProcesses @($script:processes | Where-Object { $_.Background }) } })
$ui.KillSelectedButton.Add_Click({
    safe {
        $ids = @($ui.ProcessGrid.SelectedItems | ForEach-Object { $_.Id })
        killProcesses @($script:processes | Where-Object { $ids -contains $_.Id })
    }
})
$ui.KillAllButton.Add_Click({ safe { refreshProcesses; killProcesses $script:processes } })

# ============================================================================
# ウィンドウ全体
# ============================================================================

$ui.CloseButton.Add_Click({ $window.Close() })

$ui.Tabs.Add_SelectionChanged({
    param ($sender, $e)
    # 中の表・一覧の選択変更も伝わってくるため、タブの切り替えだけを扱う
    if ($e.OriginalSource -ne $ui.Tabs) {
        return
    }
    safe {
        if ($ui.Tabs.SelectedItem -eq $ui.KillTab) {
            refreshProcesses
            $script:processTimer.Start()
        } else {
            $script:processTimer.Stop()
        }
        if ($ui.Tabs.SelectedItem -eq $ui.IndexTab) {
            refreshConversionState
        }
    }
})

$window.Add_Activated({
    safe {
        # 変換対象フォルダがほかの画面で変更されていれば読み直す
        if ((getTargetsKey @(getTargetFolders)) -ne $script:savedTargets) {
            loadTargets
            setStatus "インデックス一覧がほかで変更されたため、読み直しました"
        }
        foreach ($item in $script:targetItems) {
            updateFolderItemStatus $item
        }
        if (!(isConverting)) {
            refreshConversionState
        }
        updateSearchTarget
        if ($ui.Tabs.SelectedItem -eq $ui.KillTab) {
            refreshProcesses
        } else {
            updateKillBadge
        }
    }
})

$window.Add_PreviewKeyDown({
    param ($sender, $e)
    $modifiers = [System.Windows.Input.Keyboard]::Modifiers
    if ($e.Key -eq "F" -and $modifiers -eq "Control") {
        $ui.Tabs.SelectedItem = $ui.SearchTab
        $ui.WordBox.Focus() | Out-Null
        $ui.WordBox.SelectAll()
        $e.Handled = $true
    } elseif ($e.Key -eq "F" -and $modifiers -eq ([System.Windows.Input.ModifierKeys]::Control -bor [System.Windows.Input.ModifierKeys]::Shift)) {
        $ui.Tabs.SelectedItem = $ui.SearchTab
        $ui.FilterBox.Focus() | Out-Null
        $e.Handled = $true
    } elseif ($e.Key -eq "F5") {
        safe {
            if ($ui.Tabs.SelectedItem -eq $ui.KillTab) {
                refreshProcesses
            } else {
                refreshConversionState
                refreshIndexSummary
                loadIndexTree
            }
        }
        $e.Handled = $true
    } elseif ($e.Key -eq "Escape" -and $script:search) {
        cancelSearch
        $e.Handled = $true
    }
})

$window.Add_Closing({
    param ($sender, $e)
    # 変換はウィンドウを出さずに動いているため、閉じる前にどうするか聞く
    if (isConverting) {
        $answer = showMessage ("変換中です。`n`n［はい］変換を中止してから閉じる（変換中のファイルが終わったところで止まります）`n" +
            "［いいえ］変換を続けたまま閉じる（もう一度開くと進み具合を表示します）`n［キャンセル］閉じない") "YesNoCancel" "Question" "Cancel"
        if ($answer -eq "Cancel") {
            $e.Cancel = $true
            return
        }
        if ($answer -eq "Yes") {
            [System.IO.File]::WriteAllText(${stopRequestFile}, "", ${utf8Bom})
        }
    }
    if ($script:search) {
        $script:search.Shared.Stop = $true
    }
})

$window.Add_Loaded({
    safe {
        if ($ui.Tabs.SelectedItem -eq $ui.SearchTab) {
            $ui.WordBox.Focus() | Out-Null
        }
    }
})

# ---- 起動 ----

loadTargets
setSearchOptionToUi (readSearchOption)
setOpenMode (readOpenMode)
updateOpenMenu
refreshConversionState
loadIndexTree
updateWordNotice
updateKillBadge
refreshIndexSummary

# 前回の画面で起動した変換が続いていれば、進み具合を表示する
$runningConversion = findRunningConversion
if ($runningConversion) {
    adoptConversion $runningConversion
}

# 起動時のタブ：変換中・中断中、またはインデックスが無ければ［1 インデックス管理］、それ以外は［2 検索］
$openIndexTab = $runningConversion -or ($script:conversionState -and $script:conversionState.Pending -gt 0) -or !(testIndexExists)
$ui.Tabs.SelectedItem = if ($openIndexTab) { $ui.IndexTab } else { $ui.SearchTab }
setStatus ""

# 多重起動したとき（2つ目のプロセスが $activateEvent を合図）に、この画面を前面へ出す。
# 画面のスレッドで一定間隔にイベントを確認する（P/Invoke を使わず、WPF の Activate で前面化する）。
# ※以前は C# の SingleInstance（AttachThreadInput 等の P/Invoke）で行っていたが、
#   実行時コンパイル（csc.exe）を無くすため、DispatcherTimer＋Window.Activate に置き換えた。
$activateTimer = New-Object System.Windows.Threading.DispatcherTimer
$activateTimer.Interval = [TimeSpan]::FromMilliseconds(300)
$activateTimer.Add_Tick({
    if ($activateEvent.WaitOne(0)) {
        if ($window.WindowState -eq [System.Windows.WindowState]::Minimized) {
            $window.WindowState = [System.Windows.WindowState]::Normal
        }
        [void]$window.Activate()
        # ほかのプロセスが前面のときは Activate が無視されることがあるため、最前面を一瞬立ててから戻す
        $window.Topmost = $true
        $window.Topmost = $false
    }
})
$activateTimer.Start()

try {
    [void]$window.ShowDialog()
} finally {
    if ($script:search) {
        $script:search.PS.Stop()
    }
    $activateTimer.Stop()
    $activateEvent.Close()
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}
