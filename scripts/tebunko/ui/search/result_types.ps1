# tebunko の画面で使う型: 検索結果とプレビュー。
# shared\ui\types.ps1 を先に読み込んでおくこと（NotifyBase を継承する）。lib.ps1（index_name.ps1）も先に読み込んでおくこと（HitRow.ObjectPlaceRegex の値を作る）。
# Preview* と HitRow は互いを参照するため、同じファイルに置く。

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
    static [int] $WidthStamp = 0     # どれかの列の幅が SetWidth で変わるたびに増える（PreviewTable.Prefix の作り直しの判断）
    [string]$Label
    [double]$Width
    PreviewColumn() {}   # 既定のコンストラクタを明示する（理由は shared\ui\types.ps1 の NotifyBase）
    [void] SetWidth([double]$value) {
        $newWidth = [Math]::Max([PreviewColumn]::MinWidth, $value)
        if ($this.Width -eq $newWidth) { return }
        $this.Width = $newWidth
        [PreviewColumn]::WidthStamp++
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
    [int]$LineNumber     # 元のファイルの場所の中の行の番号（行を前後に足し・捨てても、選んだ範囲がずれない）
    [int]$ColumnIndex
    [bool]$IsSelected
    PreviewCell() {}   # 既定のコンストラクタを明示する（理由は shared\ui\types.ps1 の NotifyBase）
    [void] SetSelected([bool]$value) {
        if ($this.IsSelected -eq $value) { return }
        $this.IsSelected = $value
        $this.Raise("IsSelected")
    }
}

# 選択行のプレビューの、見えている列の範囲（表の全部の行が同じものを持つ）。First・Last は見えている最初と最後の列
# （左右の余白の列を含む。0 から）、LeftSpace・RightSpace はその左右の、作らない列の幅の合計。
# 行の左右の空きの幅がこれに束ねてあり、範囲が変わると、見えている行すべての空きが一度に変わる
class PreviewWindow : NotifyBase {
    [int]$First
    [int]$Last = -1
    [double]$LeftSpace
    [double]$RightSpace
    PreviewWindow() {}   # 既定のコンストラクタを明示する（理由は shared\ui\types.ps1 の NotifyBase）
    [void] SetSpaces([double]$left, [double]$right) {
        if ($this.LeftSpace -ne $left) { $this.LeftSpace = $left; $this.Raise("LeftSpace") }
        if ($this.RightSpace -ne $right) { $this.RightSpace = $right; $this.Raise("RightSpace") }
    }
}

# 選択行のプレビューの1行。「読み込み中…」などを出す行（IsMessage）は、行番号・セルを持たず、選択・コピーの対象にしない。
# 行の全部のセルの値は Texts に持ち、画面に出すセル（Cells）は、見えている列の範囲（Window）の分だけを、見えている行について作る
class PreviewRow : NotifyBase {
    [string]$Number
    [int]$LineNumber     # 元のファイルの場所の中の行の番号（メッセージの行は 0）
    [bool]$IsHitRow
    [bool]$IsMessage
    [string]$Message
    [System.Collections.Generic.List[string]]$Texts   # 行の全部のセルの値（TSV から切り出したまま）
    [PreviewWindow]$Window
    [int]$CellsVersion = -1   # Cells を作ったときの範囲の版（-1 は未作成）
    [System.Collections.Generic.List[PreviewCell]]$Cells = [System.Collections.Generic.List[PreviewCell]]::new()
    PreviewRow() {}   # 既定のコンストラクタを明示する（理由は shared\ui\types.ps1 の NotifyBase）
    [void] SetCells([System.Collections.Generic.List[PreviewCell]]$cells, [int]$version) {
        $this.Cells = $cells
        $this.CellsVersion = $version
        $this.Raise("Cells")
    }
}

# 選択行のプレビュー。HitOffset・HitWidth は選択行で最初に一致したセルの左端と幅（横スクロール用）。
# Rows は前後に足し・捨てられる集まり（行の仮想化で見えている所だけ要素を作る）。読んでいる所の端には、
# 読み込み中・読めなかったことを出すメッセージの行（UpMessage・DownMessage）を置く
class PreviewTable {
    [System.Collections.Generic.List[PreviewColumn]]$Columns
    [System.Collections.ObjectModel.ObservableCollection[PreviewRow]]$Rows
    [double]$HitOffset
    [double]$HitWidth
    [PreviewWindow]$Window = [PreviewWindow]::new()
    [int]$WindowVersion       # 見えている列の範囲が変わるたびに増える（行のセルを作り直すかの判断）
    [HitRow]$Owner            # 行のセルを作る（一致の印に、検索の語を使うため）
    hidden [double[]]$widthSums              # Prefix() が作った累積
    hidden [int]$prefixCount = -1            # 累積を作ったときの列の数
    hidden [int]$prefixStamp = -1            # 累積を作ったときの PreviewColumn.WidthStamp
    hidden [System.Collections.Generic.List[PreviewRow]]$built = [System.Collections.Generic.List[PreviewRow]]::new()   # セルを作ってある行
    [bool]$IsExcel   # Excel の表か（Excel 以外は列が 1 つで、幅を枠いっぱいにする。preview.ps1 の fitPreviewWidth）
    # 読み足し（preview.ps1）の状態。TopDone・BottomDone はその向きにもう行が無い（または読めなかった）こと。
    # Marks は readContentIndexLines の控え（続きを、先頭から読み飛ばさず読むため）
    [bool]$TopDone
    [bool]$BottomDone
    $Marks
    [int]$Request             # この表を作った読み込みの番号（読み足しの結果が、この表のものかを見分ける）
    [PreviewRow]$UpMessage
    [PreviewRow]$DownMessage

    # 行の集まりの中の、メッセージの行を除いた行（本物の行）の数・最初と最後の行の番号（無ければ 0）
    [int] LeadCount() { return $(if ($null -ne $this.UpMessage) { 1 } else { 0 }) }
    [int] DataCount() {
        return $this.Rows.Count - $this.LeadCount() - $(if ($null -ne $this.DownMessage) { 1 } else { 0 })
    }
    [int] FirstLine() {
        if ($this.DataCount() -le 0) { return 0 }
        return $this.Rows[$this.LeadCount()].LineNumber
    }
    [int] LastLine() {
        if ($this.DataCount() -le 0) { return 0 }
        return $this.Rows[$this.LeadCount() + $this.DataCount() - 1].LineNumber
    }

    # side（"up"・"down"）の端にメッセージの行を出す（あれば置き換える）
    [void] ShowMessage([string]$side, [string]$text) {
        $this.HideMessage($side)
        $message = [PreviewRow]::new()
        $message.IsMessage = $true
        $message.Message = $text
        if ($side -eq "up") {
            $this.Rows.Insert(0, $message)
            $this.UpMessage = $message
        } else {
            $this.Rows.Add($message)
            $this.DownMessage = $message
        }
    }

    [void] HideMessage([string]$side) {
        if ($side -eq "up" -and $null -ne $this.UpMessage) {
            [void]$this.Rows.Remove($this.UpMessage)
            $this.UpMessage = $null
        } elseif ($side -eq "down" -and $null -ne $this.DownMessage) {
            [void]$this.Rows.Remove($this.DownMessage)
            $this.DownMessage = $null
        }
    }

    # 読み足した行を、side（"up"・"down"）の端（メッセージの行の内側）に足す。rows は行の番号の順
    [void] InsertRows([string]$side, [System.Collections.Generic.List[PreviewRow]]$rows) {
        if ($side -eq "up") {
            $at = $this.LeadCount()
            foreach ($row in $rows) { $this.Rows.Insert($at++, $row) }
        } else {
            $at = $this.Rows.Count - $(if ($null -ne $this.DownMessage) { 1 } else { 0 })
            foreach ($row in $rows) { $this.Rows.Insert($at++, $row) }
        }
    }

    # 持つ行を減らすために、side（"top"・"bottom"）の端から本物の行を count 行捨てる。
    # 捨てた行を含む範囲を選んでいたら、選択を外す
    [void] TrimRows([string]$side, [int]$count) {
        $count = [Math]::Min($count, $this.DataCount())
        if ($count -le 0) { return }
        $from = $(if ($side -eq "top") { $this.LeadCount() } else { $this.LeadCount() + $this.DataCount() - $count })
        $low = $this.Rows[$from].LineNumber
        $high = $this.Rows[$from + $count - 1].LineNumber
        for ($i = 0; $i -lt $count; $i++) {
            [void]$this.built.Remove($this.Rows[$from])
            $this.Rows.RemoveAt($from)
        }
        if ($this.anchorRow -ge 0) {
            $selectedFrom = [Math]::Min($this.anchorRow, $this.focusRow)
            $selectedTo = [Math]::Max($this.anchorRow, $this.focusRow)
            if ($selectedTo -ge $low -and $selectedFrom -le $high) { $this.ClearSelection() }
        }
    }

    hidden [int]$anchorRow = -1
    hidden [int]$anchorColumn = -1
    hidden [int]$focusRow = -1
    hidden [int]$focusColumn = -1

    [bool] HasSelection() { return $this.anchorRow -ge 0 }

    [void] Select([PreviewCell]$cell, [bool]$extend) {
        if ($null -eq $cell) { return }
        if (-not $extend -or $this.anchorRow -lt 0) {
            $this.anchorRow = $cell.LineNumber
            $this.anchorColumn = $cell.ColumnIndex
        }
        $this.focusRow = $cell.LineNumber
        $this.focusColumn = $cell.ColumnIndex
        $this.ApplySelection()
    }

    [void] SelectRow([PreviewCell]$cell) {
        if ($null -eq $cell) { return }
        $this.SelectLine($cell.LineNumber)
    }

    # 行番号が lineNumber の行を、全部の列を選ぶ
    [void] SelectLine([int]$lineNumber) {
        if ($this.Columns.Count -eq 0) { return }
        $this.anchorRow = $lineNumber
        $this.focusRow = $lineNumber
        $this.anchorColumn = 0
        $this.focusColumn = $this.Columns.Count - 1
        $this.ApplySelection()
    }

    # 選んでいる範囲の、一番上の行の番号（選んでいなければ 0）
    [int] SelectedFirstLine() {
        if ($this.anchorRow -lt 0) { return 0 }
        return [Math]::Min($this.anchorRow, $this.focusRow)
    }

    # 行番号 lineNumber・列 columnIndex のセルが、選んでいる範囲の中か
    [bool] IsInSelection([int]$lineNumber, [int]$columnIndex) {
        return $this.anchorRow -ge 0 -and
            $lineNumber -ge [Math]::Min($this.anchorRow, $this.focusRow) -and $lineNumber -le [Math]::Max($this.anchorRow, $this.focusRow) -and
            $columnIndex -ge [Math]::Min($this.anchorColumn, $this.focusColumn) -and $columnIndex -le [Math]::Max($this.anchorColumn, $this.focusColumn)
    }

    [void] ClearSelection() {
        $this.anchorRow = -1; $this.anchorColumn = -1; $this.focusRow = -1; $this.focusColumn = -1
        $this.ApplySelection()
    }

    # 作ってあるセル（見えている範囲）の選択の印を、選んでいる範囲に合わせる。見えていないセルは、作るときに合わせる
    hidden [void] ApplySelection() {
        foreach ($row in $this.built) {
            foreach ($cell in $row.Cells) {
                $cell.SetSelected($this.IsInSelection($cell.LineNumber, $cell.ColumnIndex))
            }
        }
    }

    # 列の幅の累積（getPreviewColumnWindow に渡す。prefix[i] は先頭から i 列分の幅の合計で、長さは列の数 + 1）。
    # 列が増える・SetWidth で幅が変わるまでは、前に作ったものを使い回す
    [double[]] Prefix() {
        $count = $this.Columns.Count
        if ($this.prefixCount -eq $count -and $this.prefixStamp -eq [PreviewColumn]::WidthStamp) { return $this.widthSums }
        $sums = New-Object 'double[]' ($count + 1)
        $total = 0.0
        for ($i = 0; $i -lt $count; $i++) {
            $total += $this.Columns[$i].Width
            $sums[$i + 1] = $total
        }
        $this.widthSums = $sums
        $this.prefixCount = $count
        $this.prefixStamp = [PreviewColumn]::WidthStamp
        return $sums
    }

    # 見えている列の範囲（first～last）と、その左右の空きの幅を決める。範囲が変わったら、作ってあるセルは古い範囲のものなので捨てる
    # （EnsureCells が、見えている行について今の範囲で作り直す）。範囲が変わったら $true
    [bool] SetWindow([int]$first, [int]$last, [double]$left, [double]$right) {
        $this.Window.SetSpaces($left, $right)
        if ($first -eq $this.Window.First -and $last -eq $this.Window.Last) { return $false }
        $this.Window.First = $first
        $this.Window.Last = $last
        $this.WindowVersion++
        foreach ($row in $this.built) { $row.SetCells([System.Collections.Generic.List[PreviewCell]]::new(), -1) }
        $this.built.Clear()
        return $true
    }

    # 見えている列の範囲の列（列見出しに出す）
    [PreviewColumn[]] WindowColumns() {
        $last = [Math]::Min($this.Window.Last, $this.Columns.Count - 1)
        if ($last -lt $this.Window.First) { return @() }
        return $this.Columns.GetRange($this.Window.First, $last - $this.Window.First + 1).ToArray()
    }

    # 行の集まりの中の firstIndex～lastIndex の行（メッセージの行を除く）に、今の範囲のセルが無ければ作る。
    # 範囲の外の行のセルは捨てる（作ったセルの数が、見えている範囲の分に収まるように）
    [void] EnsureCells([int]$firstIndex, [int]$lastIndex) {
        $lead = $this.LeadCount()
        $firstIndex = [Math]::Max($lead, $firstIndex)
        $lastIndex = [Math]::Min($lead + $this.DataCount() - 1, $lastIndex)
        if ($lastIndex -lt $firstIndex) { return }
        $low = $this.Rows[$firstIndex].LineNumber
        $high = $this.Rows[$lastIndex].LineNumber
        for ($i = $this.built.Count - 1; $i -ge 0; $i--) {
            $row = $this.built[$i]
            if ($row.LineNumber -lt $low -or $row.LineNumber -gt $high) {
                $row.SetCells([System.Collections.Generic.List[PreviewCell]]::new(), -1)
                $this.built.RemoveAt($i)
            }
        }
        for ($i = $firstIndex; $i -le $lastIndex; $i++) {
            $row = $this.Rows[$i]
            if ($row.CellsVersion -eq $this.WindowVersion) { continue }
            $this.Owner.BuildCells($this, $row)
            $this.built.Add($row)
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
        $first = $true
        foreach ($row in $this.Rows) {
            if ($row.IsMessage -or $row.LineNumber -lt $rowFrom -or $row.LineNumber -gt $rowTo) { continue }
            if (!$first) { [void]$text.Append("`r`n") }
            $first = $false
            for ($c = $columnFrom; $c -le $columnTo; $c++) {
                if ($c -gt $columnFrom) { [void]$text.Append("`t") }
                $cellText = $(if ($c -lt $row.Texts.Count) { $row.Texts[$c].Replace([HitRow]::CellNewLine, "`n") } else { "" })
                [void]$text.Append([PreviewTable]::QuoteForExcel($cellText))
            }
        }
        return $text.ToString()
    }

    [int] SelectedCount() {
        if ($this.anchorRow -lt 0) { return 0 }
        return ([Math]::Abs($this.anchorRow - $this.focusRow) + 1) * ([Math]::Abs($this.anchorColumn - $this.focusColumn) + 1)
    }

    hidden [string] CellText([int]$lineNumber, [int]$columnIndex) {
        foreach ($row in $this.Rows) {
            if ($row.IsMessage -or $row.LineNumber -ne $lineNumber) { continue }
            if ($columnIndex -ge 0 -and $columnIndex -lt $row.Texts.Count) { return $row.Texts[$columnIndex].Replace([HitRow]::CellNewLine, "`n") }
            return ""
        }
        return ""
    }

    static [string] QuoteForExcel([string]$text) {
        if ($text.IndexOf("`n") -lt 0 -and $text.IndexOf("`t") -lt 0 -and $text.IndexOf('"') -lt 0) { return $text }
        return '"' + $text.Replace('"', '""') + '"'
    }
}

# 検索結果の、元のファイル 1 つ分（結果の表の見出しの行）。検索のヒットは生のまま Hits に持ち、表の行（HitRow）は
# 開いたとき・絞り込み・並べ替え・出力のときに初めて作って Rows に入れる。開いているときだけ、見出しの下に表の行として並べる（result_list.ps1）。
# 文言（AppKind・LocationText）は画面側で判断層（search\result_list_view.ps1）の関数から作って入れる
class FileGroup : NotifyBase {
    [bool]$IsFileHeader = $true   # 結果の表で見出しの形にする（search/result_list.xaml の FileHeaderRow）
    [int]$Order                   # 見つかった順（並べ替えで同じ値のときの順）
    [string]$Book
    [string]$RelDir
    [string]$FullPath
    [string]$AppKind        # Excel / Word / PowerPoint（アイコンの色と文字を決める。どれでもなければ空）
    [string]$LocationText   # ヒットした場所の要約（「[シート]4月 ほか 2 か所」。読み上げに使う）
    [string]$LocationPrefix # 見出しの右端の件数の前に付ける場所（「[シート]4月 ほか 2 か所 ・ 」）。場所が 1 か所だけのときは空（行の「場所」列と重なるため）
    [bool]$IsExpanded       # 見出しの下にヒットした行を並べるか（検索した直後は閉じている）
    [int]$ShownCount        # 見出しに出す件数（絞り込みに合うヒットの数）
    # 検索のヒット（searchContentIndex の結果そのまま。見つかった順）
    [System.Collections.Generic.List[object]]$Hits = [System.Collections.Generic.List[object]]::new()
    # 作った表の行（Hits の先頭から順に作る。並べ替えたらその順）と、そのうち絞り込みに合う行
    [System.Collections.Generic.List[object]]$Rows = [System.Collections.Generic.List[object]]::new()
    [System.Collections.Generic.List[object]]$ShownRows = [System.Collections.Generic.List[object]]::new()
    # 結果の表での状態（result_list.ps1 が使う）。InView は見出しを表に入れたか、DisplayedCount は見出しの下に入れた行の数
    [bool]$InView
    [int]$DisplayedCount
    hidden [System.Collections.Generic.HashSet[string]]$locations = [System.Collections.Generic.HashSet[string]]::new()
    hidden [System.Collections.Generic.List[string]]$labels = [System.Collections.Generic.List[string]]::new()

    # ヒットした場所（結果の「場所」そのまま）を記録する。初めての場所なら $true
    # （呼び出し側が表記を AddLabel で足し、LocationText を作り直す。ヒットごとに表記を作らないため）
    [bool] AddLocation([string]$location) {
        return $this.locations.Add($location)
    }

    # 場所の表記を足す（図形・コメントは元の場所と同じ表記になるため、同じ表記は 1 つにする）。足したら $true
    [bool] AddLabel([string]$label) {
        if ($this.labels.Contains($label)) { return $false }
        $this.labels.Add($label)
        return $true
    }

    [string[]] GetLocations() {
        return $this.labels.ToArray()
    }

    [void] SetLocationText([string]$text) {
        $this.LocationText = $text
        $this.Raise("LocationText")
    }

    [void] SetLocationText([string]$text, [string]$prefix) {
        $this.LocationPrefix = $prefix
        $this.Raise("LocationPrefix")
        $this.SetLocationText($text)
    }

    [void] SetExpanded([bool]$value) {
        if ($this.IsExpanded -eq $value) { return }
        $this.IsExpanded = $value
        $this.Raise("IsExpanded")
    }

    [void] SetCount([int]$count) {
        if ($this.ShownCount -eq $count) { return }
        $this.ShownCount = $count
        $this.Raise("ShownCount")
    }
}

# 検索結果の1行。生成時は生データのみ。表示用（DisplayLine・Segments・MatchCell・MatchCount）は Prepare() で作る（可視行だけ）。
# 表示用の各項目は Prepare() 後に PropertyChanged を出す（LoadingRow より前にバインドされても更新されるように）
class HitRow : NotifyBase {
    static [regex] $CellRegex  = [regex]::new("\t(?:`"(?:[^`"]|`"`")*`"[^\t]*|[^\t]*)")
    static [regex] $QuoteRegex = [regex]::new("^`"((?:[^`"]|`"`")*)`"(.*)`$", [System.Text.RegularExpressions.RegexOptions]::Singleline)
    # 図形・コメントの場所 "<元の場所>[<種類>]"。クラスからはスクリプトの関数が見えないため、値はクラスの定義の後で入れる（newObjectPlaceKindPattern）。
    # Excel の図形・コメントの 1 行は "<セル番地><TAB><文字>"。ヘッダー・フッターの 1 行は文字だけ（セル番地が無い）
    static [regex] $ObjectPlaceRegex
    static [char] $CellNewLine = [char]0x2028   # TSV のセル内改行（shared\core\text.ps1 の cellNewLine）
    static [int] $LeadLength = 40
    static [int] $MaxDisplay = 600
    static [double] $NumberWidth = 44
    static [double] $MinCellWidth = 48
    static [double] $MaxCellWidth = 260
    static [double] $MaxParagraphWidth = 640
    static [int] $WidthSampleRows = 10   # 列の幅を決めるために、選択行の前後何行のセルを測るか
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
    [string]$PlaceText    # 場所ごとの表記（describePlace。例: [シート]売上）。生成した側が入れる
    [string]$PlaceDisplay # 表の「場所」の列の表記（describeHitPlace。例: [シート]売上!B12）。画面に出るとき SetPlaceDisplay で入れる
    [string]$Kind         # 画面の「種別」（セル・図形・コメント・本文・ノート）
    [int]$LineNumber
    [string]$Line
    [bool]$IsExcel         # Excel のファイルの行か（行を作る側が getContentIndexFileKind で決めて入れる）
    [bool]$IsText          # テキストの拡張子（.txt 等）の行か（describeHitPlace・open_source.ps1 が使う。呼び出し側が入れる）
    [bool]$IsObjectPlace
    [bool]$IsCellPrefixed   # 行の先頭がセル番地 + タブの場所（Excel の図形・コメント）か（testCellPrefixedPlace。行を作る側が入れる。既定は $false = セル番地を出さない・選ばない）
    [string]$MatchCell
    [int]$MatchCount        # 1 行のうち一致したセルの数（Excel のセルの行だけ。図形・コメントは 0 か 1）
    [string]$DisplayLine
    [System.Collections.Generic.List[Segment]]$Segments
    [bool]$Prepared
    [bool]$IsFileHeader     # 常に $false（結果の表で、見出しの行と区別する）
    [int]$Order             # そのファイルの中で見つかった順（並べ替えで同じ値のときの順）
    [FileGroup]$FileGroup   # 元のファイルの見出し（同じファイルの行で共有する）

    hidden [string]$word
    hidden [regex]$pattern

    HitRow() {}   # 既定のコンストラクタを明示する（理由は shared\ui\types.ps1 の NotifyBase）

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
        $row.IsObjectPlace = [HitRow]::ObjectPlaceRegex.IsMatch($(if ($null -eq $location) { "" } else { $location }))
        return $row
    }

    # 表示用（強調セグメント・DisplayLine・セル番地・一致したセルの数）を作る。可視行になったときに1回だけ呼ぶ（LoadingRow）
    [void] Prepare() {
        if ($this.Prepared) { return }
        $this.Prepared = $true
        $text = $this.ShownText()
        $this.DisplayLine = [HitRow]::ToDisplay($text)
        $this.Segments = $this.BuildSegments($text)
        $this.SetMatchCell()
        $this.Raise("DisplayLine"); $this.Raise("Segments"); $this.Raise("MatchCell"); $this.Raise("MatchCount")
    }

    hidden [void] SetMatchCell() {
        $this.MatchCell = ""
        $this.MatchCount = 0
        if (-not $this.IsExcel) { return }
        $cells = [HitRow]::SplitCells($this.Line, $true)
        if ($this.IsObjectPlace) {
            # Excel の図形・コメントの行は、先頭のセルが図形の左上・コメントのセルの番地
            $hit = $false
            foreach ($cell in $cells) { if ([HitRow]::HasMatch($cell, $this.word, $this.pattern)) { $hit = $true; break } }
            # ヘッダー・フッターの行は文字だけでセル番地が無い（MatchCell は空のまま。場所は [シート]<名前> になる）
            if ($hit -and $cells.Count -gt 0 -and $this.IsCellPrefixed) { $this.MatchCell = $cells[0] }
            $this.MatchCount = $(if ($hit) { 1 } else { 0 })
            return
        }
        $count = 0
        for ($i = 0; $i -lt $cells.Count; $i++) {
            if (-not [HitRow]::HasMatch($cells[$i], $this.word, $this.pattern)) { continue }
            if ($count -eq 0) { $this.MatchCell = [HitRow]::ColumnName($i + 1) + $this.LineNumber }
            $count++
        }
        $this.MatchCount = $count
    }

    # 表の「場所」の列に出す文字を入れる（describeHitPlace の結果）。画面に出るとき 1 回入れる
    [void] SetPlaceDisplay([string]$value) {
        if ($this.PlaceDisplay -eq $value) { return }
        $this.PlaceDisplay = $value
        $this.Raise("PlaceDisplay")
    }

    # 絞り込み。生データ（相対フォルダ・元ファイル名・場所・種別・生の行）での部分一致。行番号は列に出さないため対象にしない（大文字小文字を区別しない）。
    # ※以前は表示用（セル番地・タブ表示）も対象にしていたが、遅延生成のため生データのみを対象にした。
    [bool] Contains([string]$text) {
        $ci = [System.StringComparison]::CurrentCultureIgnoreCase
        if ($this.RelDir.IndexOf($text, $ci) -ge 0) { return $true }
        if (("" + $this.Book).IndexOf($text, $ci) -ge 0) { return $true }
        if (("" + $this.Location).IndexOf($text, $ci) -ge 0) { return $true }
        if (("" + $this.PlaceText).IndexOf($text, $ci) -ge 0) { return $true }
        if (("" + $this.Kind).IndexOf($text, $ci) -ge 0) { return $true }
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

    # 「該当行」列に出す文字。Excel の図形・コメントの行は、先頭のセル番地を「場所」列に出すため除き、
    # 囲みの " を外した文字にする（"納期は<改行>別途" → 納期は<改行>別途）。ほかは TSV の行のまま
    hidden [string] ShownText() {
        if (-not ($this.IsExcel -and $this.IsObjectPlace)) { return $this.Line }
        $cells = [HitRow]::SplitCells($this.Line, $true)
        if (-not $this.IsCellPrefixed) { return ($cells -join "`t") }
        if ($cells.Count -lt 2) { return $this.Line }
        return ($cells.GetRange(1, $cells.Count - 1) -join "`t")
    }

    hidden [System.Collections.Generic.List[Segment]] BuildSegments([string]$line) {
        $segs = [System.Collections.Generic.List[Segment]]::new()
        $pos = 0; $shown = 0
        foreach ($m in [HitRow]::FindMatches($line, $this.word, $this.pattern)) {
            if ($m[0] -lt $pos) { continue }
            if ($m[0] + $m[1] -gt $line.Length) { break }
            $before = $line.Substring($pos, $m[0] - $pos)
            if ($segs.Count -eq 0 -and $before.Length -gt [HitRow]::LeadLength) {
                $before = [char]0x2026 + $before.Substring($before.Length - [HitRow]::LeadLength)
            }
            if ($before.Length -gt 0) { $segs.Add([Segment]::new([HitRow]::ToDisplay($before), $false)) }
            $segs.Add([Segment]::new([HitRow]::ToDisplay($line.Substring($m[0], $m[1])), $true))
            $shown += $before.Length + $m[1]
            $pos = $m[0] + $m[1]
            if ($shown -gt [HitRow]::MaxDisplay) { break }
        }
        if ($pos -lt $line.Length) {
            $rest = $line.Substring($pos)
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

    # 選択行のプレビュー（前後の行をセルに分けた表）を作る。numbers・lines は readContentIndexLines の結果。
    # 列は全部の列を持つ。画面に出すセルは作らない（見えている範囲を SetWindow で決め、EnsureCells で作る）。
    # 列の幅は、列見出しの文字と、選択行の前後 WidthSampleRows 行のセルから決める（全部の行・列は測らない）
    [PreviewTable] BuildPreview([int[]]$numbers, [string[]]$lines) {
        if ($null -eq $numbers -or $null -eq $lines -or $numbers.Length -eq 0 -or $numbers.Length -ne $lines.Length) {
            $numbers = @($this.LineNumber)
            $lines = @($this.Line)
        }
        $rowTexts = [System.Collections.Generic.List[System.Collections.Generic.List[string]]]::new()
        $columnCount = 0
        foreach ($ln in $lines) {
            $cells = [HitRow]::SplitCells($ln, $this.IsExcel)
            $rowTexts.Add($cells)
            $columnCount = [Math]::Max($columnCount, $cells.Count)
        }

        $hitColumn = -1
        for ($r = 0; $r -lt $rowTexts.Count -and $hitColumn -lt 0; $r++) {
            if ($numbers[$r] -ne $this.LineNumber) { continue }
            for ($c = 0; $c -lt $rowTexts[$r].Count; $c++) {
                if ([HitRow]::HasMatch($rowTexts[$r][$c], $this.word, $this.pattern)) { $hitColumn = $c; break }
            }
        }

        $table = [PreviewTable]::new()
        $table.Owner = $this
        $table.Columns = [System.Collections.Generic.List[PreviewColumn]]::new()
        $table.Rows = [System.Collections.ObjectModel.ObservableCollection[PreviewRow]]::new()
        $table.HitOffset = 0
        $table.HitWidth = 0
        $table.IsExcel = $this.IsExcel
        $sample = [System.Collections.Generic.List[System.Collections.Generic.List[string]]]::new()
        for ($r = 0; $r -lt $rowTexts.Count; $r++) {
            if ([Math]::Abs($numbers[$r] - $this.LineNumber) -le [HitRow]::WidthSampleRows) { $sample.Add($rowTexts[$r]) }
        }
        $this.AddColumns($table, $columnCount, $sample)

        foreach ($row in $this.MakeRows($table, $numbers, $rowTexts)) { $table.Rows.Add($row) }
        if ($hitColumn -ge 0) {
            $left = [HitRow]::NumberWidth
            for ($c = 0; $c -lt $hitColumn; $c++) { $left += $table.Columns[$c].Width }
            $table.HitOffset = $left
            $table.HitWidth = $table.Columns[$hitColumn].Width
        }
        $table.TopDone = ($numbers.Length -gt 0 -and $numbers[0] -le 1)
        # 見えている範囲が決まるまでは、行の右の空きに全部の列の幅を入れておく（横にスクロールできる広さにする）
        $total = 0.0
        foreach ($column in $table.Columns) { $total += $column.Width }
        $table.Window.SetSpaces(0, $total)
        return $table
    }

    # 読み足した行（numbers・lines は readContentIndexLines の結果）から、表の行を作る。
    # それまでより多い列がある行は、列を右に足す（足した列の幅は、この行のセルから決める。前からある列の幅は変えない）
    [System.Collections.Generic.List[PreviewRow]] BuildPreviewRows([PreviewTable]$table, [int[]]$numbers, [string[]]$lines) {
        $rowTexts = [System.Collections.Generic.List[System.Collections.Generic.List[string]]]::new()
        $columnCount = 0
        foreach ($ln in $lines) {
            $cells = [HitRow]::SplitCells($ln, $this.IsExcel)
            $rowTexts.Add($cells)
            $columnCount = [Math]::Max($columnCount, $cells.Count)
        }
        if ($columnCount -gt $table.Columns.Count) {
            $sample = [System.Collections.Generic.List[System.Collections.Generic.List[string]]]::new()
            for ($r = 0; $r -lt $rowTexts.Count -and $sample.Count -lt 2 * [HitRow]::WidthSampleRows; $r++) { $sample.Add($rowTexts[$r]) }
            $this.AddColumns($table, $columnCount, $sample)
        }
        return $this.MakeRows($table, $numbers, $rowTexts)
    }

    # table の列を columnCount 列まで増やす。足す列の幅は、列見出しの文字と sample の行のセルから決める
    hidden [void] AddColumns([PreviewTable]$table, [int]$columnCount, $sample) {
        $maxWidth = $(if ($this.IsExcel) { [HitRow]::MaxCellWidth } else { [HitRow]::MaxParagraphWidth })
        for ($c = $table.Columns.Count; $c -lt $columnCount; $c++) {
            $width = [HitRow]::TextWidth($this.ColumnLabel($c + 1))
            foreach ($cells in $sample) {
                if ($c -lt $cells.Count) { $width = [Math]::Max($width, [HitRow]::CellWidth($cells[$c])) }
            }
            $col = [PreviewColumn]::new()
            $col.Label = $this.ColumnLabel($c + 1)
            $col.Width = [Math]::Min($maxWidth, [Math]::Max([HitRow]::MinCellWidth, $width))
            $table.Columns.Add($col)
        }
    }

    hidden [System.Collections.Generic.List[PreviewRow]] MakeRows([PreviewTable]$table, [int[]]$numbers, $rowTexts) {
        $rows = [System.Collections.Generic.List[PreviewRow]]::new()
        for ($r = 0; $r -lt $rowTexts.Count; $r++) {
            $row = [PreviewRow]::new()
            $row.Number = $numbers[$r].ToString()
            $row.LineNumber = $numbers[$r]
            $row.IsHitRow = $numbers[$r] -eq $this.LineNumber
            $row.Texts = $rowTexts[$r]
            $row.Window = $table.Window
            $rows.Add($row)
        }
        return $rows
    }

    # 行 row の、見えている列の範囲（table.Window）のセルを作って、行に入れる（PreviewTable.EnsureCells から呼ぶ）
    [void] BuildCells([PreviewTable]$table, [PreviewRow]$row) {
        $cells = [System.Collections.Generic.List[PreviewCell]]::new()
        $last = [Math]::Min($table.Window.Last, $table.Columns.Count - 1)
        for ($c = $table.Window.First; $c -le $last; $c++) {
            $raw = $(if ($c -lt $row.Texts.Count) { $row.Texts[$c] } else { "" })
            $text = $raw.Replace([HitRow]::CellNewLine, "`n")
            $pc = [PreviewCell]::new()
            $pc.Text = $text
            $pc.Display = [HitRow]::Shorten($text, [HitRow]::MaxCellChars)
            $pc.ToolTip = $(if ($text.Length -gt 0) { [HitRow]::Shorten($text, [HitRow]::MaxToolTipChars) } else { $null })
            $pc.Column = $table.Columns[$c]
            $pc.IsHit = [HitRow]::HasMatch($raw, $this.word, $this.pattern)
            $pc.LineNumber = $row.LineNumber
            $pc.ColumnIndex = $c
            $pc.IsSelected = $table.IsInSelection($row.LineNumber, $c)
            $cells.Add($pc)
        }
        $row.SetCells($cells, $table.WindowVersion)
    }

    static [double] CellWidth([string]$cell) {
        if ([string]::IsNullOrEmpty($cell)) { return [HitRow]::TextWidth("") }
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
        if ($this.IsExcel -and $this.IsObjectPlace -and -not $this.IsCellPrefixed -and $number -eq 1) { return "文字" }
        if ($this.IsExcel -and $this.IsObjectPlace) {
            return $(switch ($number) { 1 { "セル" } 2 { "文字" } default { $number.ToString() } })
        }
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
            $number = [int][Math]::Floor($number / 26)   # [int] だけでは四捨五入になる（26 → "AZ"）
        }
        return $name
    }
}

# index_name.ps1 の newObjectPlaceKindPattern を使うため、lib.ps1 を先に読み込んでおくこと
[HitRow]::ObjectPlaceRegex = [regex]::new((newObjectPlaceKindPattern))
