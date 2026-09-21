# tebunko_grep の画面で使う型（検索結果・プレビュー・インデックス一覧など）。
# shared\ui\types.ps1 を先に読み込んでおくこと（NotifyBase を継承する）。

class Segment {
    [string]$Text
    [bool]$IsHit

    Segment() {}

    # 表示する行ごとに作るため、ハッシュテーブルからの変換（[Segment]@{ … }）より速いコンストラクタで作る
    Segment([string]$text, [bool]$isHit) {
        $this.Text = $text
        $this.IsHit = $isHit
    }
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
        $text = [System.Text.StringBuilder]::new()
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
    static [char] $CellNewLine = [char]0x2028   # TSV のセル内改行（shared\core\text.ps1 の cellNewLine）
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
            if (-not [HitRow]::HasMatch($cells[$i], $this.word, $this.pattern)) { continue }
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

    # FindMatches が 1 件以上返すか（一致の一覧を作らずに調べる。表示する行ごとにセルの数だけ呼ぶため）
    static [bool] HasMatch([string]$text, [string]$word, [regex]$pattern) {
        if ([string]::IsNullOrEmpty($text)) { return $false }
        if ($null -ne $pattern) {
            $m = $pattern.Match($text)
            while ($m.Success) {
                if ($m.Length -gt 0) { return $true }
                $m = $m.NextMatch()
            }
            return $false
        }
        if ([string]::IsNullOrEmpty($word)) { return $false }
        return $text.IndexOf($word, [System.StringComparison]::CurrentCultureIgnoreCase) -ge 0
    }

    static [System.Collections.Generic.List[int[]]] FindMatches([string]$text, [string]$word, [regex]$pattern) {
        $list = [System.Collections.Generic.List[int[]]]::new()
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
        $segs = [System.Collections.Generic.List[Segment]]::new()
        $pos = 0; $shown = 0
        foreach ($m in [HitRow]::FindMatches($this.Line, $this.word, $this.pattern)) {
            if ($m[0] -lt $pos) { continue }
            if ($m[0] + $m[1] -gt $this.Line.Length) { break }
            $before = $this.Line.Substring($pos, $m[0] - $pos)
            if ($segs.Count -eq 0 -and $before.Length -gt [HitRow]::LeadLength) {
                $before = [char]0x2026 + $before.Substring($before.Length - [HitRow]::LeadLength)
            }
            if ($before.Length -gt 0) { $segs.Add([Segment]::new([HitRow]::ToDisplay($before), $false)) }
            $segs.Add([Segment]::new([HitRow]::ToDisplay($this.Line.Substring($m[0], $m[1])), $true))
            $shown += $before.Length + $m[1]
            $pos = $m[0] + $m[1]
            if ($shown -gt [HitRow]::MaxDisplay) { break }
        }
        if ($pos -lt $this.Line.Length) {
            $rest = $this.Line.Substring($pos)
            if ($rest.Length -gt [HitRow]::MaxDisplay) { $rest = $rest.Substring(0, [HitRow]::MaxDisplay) + [char]0x2026 }
            $segs.Add([Segment]::new([HitRow]::ToDisplay($rest), $false))
        }
        return $segs
    }

    static [System.Collections.Generic.List[string]] SplitCells([string]$line, [bool]$isExcel) {
        $cells = [System.Collections.Generic.List[string]]::new()
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
        $rowCells = [System.Collections.Generic.List[System.Collections.Generic.List[string]]]::new()
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
                if ([HitRow]::HasMatch($rowCells[$r][$c], $this.word, $this.pattern)) { $hitColumn = $c; break }
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
        $table.Columns = [System.Collections.Generic.List[PreviewColumn]]::new()
        $table.Rows = [System.Collections.Generic.List[PreviewRow]]::new()
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
            $row.Cells = [System.Collections.Generic.List[PreviewCell]]::new()
            $left = [HitRow]::NumberWidth
            for ($i = 0; $i -lt $shownColumns; $i++) {
                $c = $firstColumn + $i
                $cell = $(if ($c -lt $rowCells[$r].Count) { $rowCells[$r][$c] } else { "" })
                $isHit = [HitRow]::HasMatch($cell, $this.word, $this.pattern)
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
    [string]$StartText
    [string]$MemoryText
    [string]$TitleText
}

# ［1 インデックス管理］の取り込みに失敗したファイル1件
class FailRow {
    [string]$RelPath
    [string]$Reason
    [string]$IngestedText
    [string]$SourcePath
}

# インデックス作成の確認ダイアログに出すインデックス1件（取り込み予定.tsv の1行）
class PlanRow {
    [string]$Name
    [string]$Path
    [string]$TargetText   # 取り込み対象の件数（"12 件" / "更新不要" / "取り込みません"）
    [object]$TargetBrush
    [string]$DetailText   # 内訳（新規 N 件 / 更新あり N 件 …）
    [string]$TotalText    # 見つかった Office ファイルの数
}

# ［1 インデックス管理］のインデックス一覧 1 件。プログラムから変えたときに画面へ反映するため通知する。
# ［作成］チェックの TwoWay バインドは値の往復に使い、保存はチェックボックスの Click で行う（PS class はセッターにロジックを書けないため）
class FolderItem : NotifyBase {
    [string]$Name          # インデックス名（work\index 直下のフォルダ名）
    [string]$Path
    [bool]$Enabled
    [string]$StatusText
    [object]$StatusBrush
    [string]$FileCountText
    [string]$FileCountToolTip
    [string]$LastIngestedText
    [bool]$StatusChecked   # フォルダの有無を調べ終えたか（別スレッドで調べる。refreshFolderStatus）
    [bool]$FolderExists    # 調べた結果、フォルダがあったか

    [void] SetEnabled([bool]$value) { if ($this.Enabled -ne $value) { $this.Enabled = $value; $this.Raise("Enabled") } }
    [void] SetName([string]$value) { if ($this.Name -ne $value) { $this.Name = $value; $this.Raise("Name") } }
    [void] SetPath([string]$value) { if ($this.Path -ne $value) { $this.Path = $value; $this.Raise("Path") } }
    [void] SetStatus([string]$text, [object]$brush) { $this.StatusText = $text; $this.StatusBrush = $brush; $this.Raise("StatusText"); $this.Raise("StatusBrush") }
    [void] SetStats([string]$countText, [string]$toolTip, [string]$lastIngested) {
        $this.FileCountText = $countText; $this.FileCountToolTip = $toolTip; $this.LastIngestedText = $lastIngested
        $this.Raise("FileCountText"); $this.Raise("FileCountToolTip"); $this.Raise("LastIngestedText")
    }
}

# 検索対象の1件（tebunko_grep\search\search_run.ps1 の getIndexTsvFiles に渡す）
class SearchTarget {
    [string]$Root
    [string]$RelPath
    [bool]$Recurse
}

# 検索対象から外したフォルダ（tebunko_grep\core\settings_grep.ps1 の readSearchExcludes / writeSearchExcludes と同じ項目）
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
