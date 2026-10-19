# ［検索］の選択行のプレビュー（tebunko\ui\preview.ps1）のテスト。
# 画面の部品（$ui.PreviewRows など）は偽物にして、表示する中身と状態の変化を確かめる。
# クリップボードは利用者の PC のものを書き換えるため、コピーは「選んでいないとき」だけを確かめる。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    . "${scriptsDir}\shared\ui\types.ps1"
    . "${scriptsDir}\tebunko\ui\search\result_types.ps1"
    . "${scriptsDir}\tebunko\ui\preview_view.ps1"
    . "${scriptsDir}\tebunko\ui\search\search_bar_view.ps1"
    . "${scriptsDir}\tebunko\ui\search\result_list_view.ps1"
    . "${scriptsDir}\tebunko\ui\search\open_source_view.ps1"

    # gui.ps1 で決める値
    ${previewFirstLines}    = 50
    ${previewReadAheadLines} = 20
    ${previewChunkLines}    = 200
    ${previewMaxRows}       = 3000
    ${previewColumnMargin}  = 3
    ${previewCellRowMargin} = 10

    # ---- 画面の偽物 ----
    # 読み込み時に登録されるイベントの処理は $handlers に、呼ばれた操作は $fake に取っておく
    $handlers = @{}
    $fake = @{ Status = ""; Current = $null; TimerStarts = 0; TimerStops = 0; Scrolls = @(); VerticalScrolls = @(); HeaderScrolls = @() }

    function newFakePart {
        # 画面の部品の偽物。events に書いた Add_<イベント> を呼ぶと、処理を $handlers["<名前>.<イベント>"] に取っておく
        param (
            [string]$name,
            [hashtable]$properties = @{},
            [string[]]$events = @()
        )

        $part = [pscustomobject]$properties
        $part | Add-Member NoteProperty PartName $name
        foreach ($event in $events) {
            $part | Add-Member ScriptMethod "Add_$event" ([scriptblock]::Create("param (`$block) `$handlers['$name.$event'] = `$block"))
        }
        return $part
    }

    $previewScroll = newFakePart "PreviewScroll" @{ ViewportWidth = 400.0; HorizontalOffset = 0.0; VerticalOffset = 0.0; ViewportHeight = 20.0 } @("ScrollChanged", "PreviewKeyDown")
    $previewScroll | Add-Member ScriptMethod UpdateLayout { }
    $previewScroll | Add-Member ScriptMethod Focus { $true }
    $previewScroll | Add-Member ScriptMethod ScrollToHorizontalOffset { param ($offset) $fake.Scrolls += $offset }
    $previewScroll | Add-Member ScriptMethod ScrollToVerticalOffset { param ($offset) $fake.VerticalScrolls += $offset; $this.VerticalOffset = [double]$offset }
    $previewSplitter = newFakePart "PreviewSplitter" @{ IsDragging = $false } @("DragCompleted")
    $resultListHost = newFakePart "ResultListHost" @{ ActualHeight = 0.0; ActualWidth = 0.0 } @("SizeChanged")
    $headerScroll = newFakePart "PreviewHeaderScroll" @{ Visibility = "Collapsed" }
    $headerScroll | Add-Member ScriptMethod ScrollToHorizontalOffset { param ($offset) $fake.HeaderScrolls += $offset }
    $previewHeader = newFakePart "PreviewHeader" @{ ItemsSource = $null }
    $previewHeader | Add-Member ScriptMethod AddHandler { param ($event, $handler) $handlers["PreviewHeader.$($event.Name)"] = $handler }

    $ui = [pscustomobject]@{
        ResultGrid          = newFakePart "ResultGrid" @{ SelectedItem = $null } @("SelectionChanged")
        PreviewHeader       = $previewHeader
        PreviewRows         = newFakePart "PreviewRows" @{ ItemsSource = $null } @("PreviewMouseLeftButtonDown", "MouseMove", "PreviewMouseRightButtonDown", "ContextMenuOpening")
        PreviewScroll       = $previewScroll
        PreviewSplitter     = $previewSplitter
        ResultListHost      = $resultListHost
        PreviewHeaderScroll = $headerScroll
        PreviewHeaderPanel  = newFakePart "PreviewHeaderPanel" @{ DataContext = $null }
        PreviewPlaceholder  = newFakePart "PreviewPlaceholder" @{ Visibility = "Visible" }
        DetailTitle         = newFakePart "DetailTitle" @{ Text = ""; ToolTip = $null }
        OpenButton          = newFakePart "OpenButton" @{ IsEnabled = $false; Content = "" }
        OpenMenuButton      = newFakePart "OpenMenuButton" @{ IsEnabled = $false }
        OpenFolderButton    = newFakePart "OpenFolderButton" @{ IsEnabled = $false }
        PreviewMenu         = newFakePart "PreviewMenu"
        MenuPreviewOpen     = newFakePart "MenuPreviewOpen" @{} @("Click")
        MenuPreviewCopy     = newFakePart "MenuPreviewCopy" @{} @("Click")
        MenuPreviewCopyRow  = newFakePart "MenuPreviewCopyRow" @{} @("Click")
    }

    # 画面の共通部品（shared\ui\shell.ps1）と結果の表（result_list.ps1）の代わり
    function safe {
        param ([scriptblock]$block)
        & $block
    }
    function setStatus {
        param ([string]$text)
        $fake.Status = $text
    }
    function newTimer {
        param ([int]$milliseconds, [scriptblock]$onTick)
        $handlers["Timer.Tick"] = $onTick
        $timer = [pscustomobject]@{ Interval = $milliseconds }
        $timer | Add-Member ScriptMethod Start { $fake.TimerStarts++ }
        $timer | Add-Member ScriptMethod Stop { $fake.TimerStops++ }
        return $timer
    }
    function getCurrentHitRow {
        return $fake.Current
    }
    function openSource { $fake.Opened++ }
    function setContextMenuItems {
        # 判断層が決めた並びを、メニューに組むところの代わり（受け取った並びを取っておく）
        param ($menu, $items, $parts)
        $fake.MenuItems = $items
        $fake.MenuParts = $parts
    }
    function startJob {
        # 画面の裏の仕事（BackgroundQueue）の代わりに、その場で実行して結果を渡す。
        # 返すまでの間に選択が変わった場合を試すため、$fake.BeforeDone があれば結果を渡す前に呼ぶ
        param ([scriptblock]$scriptBlock, [object[]]$arguments, [scriptblock]$onDone)
        $output = $null
        $errorText = $null
        try {
            $output = @(& $scriptBlock @arguments)
        } catch {
            $errorText = $_.Exception.Message
        }
        if ($fake.BeforeDone) {
            & $fake.BeforeDone
        }
        & $onDone $output $errorText
    }

    . "${scriptsDir}\tebunko\ui\preview.ps1"

    # ---- テストの準備 ----

    function resetPreview {
        $fake.Status = ""
        $fake.Current = $null
        $fake.TimerStarts = 0
        $fake.TimerStops = 0
        $fake.Scrolls = @()
        $fake.VerticalScrolls = @()
        $fake.HeaderScrolls = @()
        $ui.PreviewScroll.ViewportWidth = 400.0
        $ui.PreviewScroll.VerticalOffset = 0.0
        $ui.ResultGrid.SelectedItem = $null
        clearDetail
    }

    function writeHitContentIndex {
        # インデックスの本文インデックスのファイル（TestDrive に作る）に、元のファイル book の場所 location の行を書く
        param ([string]$path, [string]$book, [string]$location, [string[]]$lines)
        [void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path))
        writeContentIndexFile $path (convertToContentIndexText @(@{ Name = $book; Places = @(@{ Place = $location; Text = (($lines -join "`r`n") + "`r`n") }) }))
    }

    function newHitRow {
        # インデックスの本文インデックスのファイル（TestDrive に作る）の lineNumber 行目にヒットした行
        param (
            [string]$book,
            [string]$location,
            [int]$lineNumber,
            [string[]]$lines,
            [string]$relDir = "営業部"
        )

        $root = "$TestDrive\index"
        $name = getContentIndexFileName (getContentIndexExtension $book)
        $relPath = if ($relDir) { "$relDir\$name" } else { $name }
        writeHitContentIndex "$root\$relPath" $book $location $lines
        $row = [HitRow]::Create("営業部", $root, $relPath, $relDir, $book, $book, $location, $lineNumber, $lines[$lineNumber - 1], "見積", [regex]"見積")
        $described = describePlace $book $location
        $row.PlaceText = $described.Place
        $row.Kind = $described.Kind
        $row.IsExcel = ((getContentIndexFileKind $book) -eq "Excel")
        return $row
    }
}

Describe "clearDetail" -Tag Unit {
    BeforeEach { resetPreview }

    It "中身を空にして案内を出し、開くボタンを使えなくする" {
        $ui.DetailTitle.Text = "前の行"
        $ui.PreviewHeader.ItemsSource = @(1)
        $script:previewTable = [PreviewTable]::new()

        clearDetail

        $ui.DetailTitle.Text | Should -Be ""
        $null -eq $ui.PreviewHeader.ItemsSource | Should -Be $true
        $null -eq $script:previewTable | Should -Be $true
        $ui.PreviewPlaceholder.Visibility | Should -Be "Visible"
        $ui.PreviewHeaderScroll.Visibility | Should -Be "Collapsed"
        $ui.OpenButton.IsEnabled | Should -Be $false
        $ui.OpenMenuButton.IsEnabled | Should -Be $false
        $ui.OpenFolderButton.IsEnabled | Should -Be $false
    }
}

Describe "updateResultPaneHeight" -Tag Unit {
    BeforeEach {
        resetPreview
        $script:previewHeightDragged = $false
        $content = [pscustomobject]@{ Needs = 150.0; DesiredSize = [pscustomobject]@{ Height = 0.0 }; Measures = 0 }
        $content | Add-Member ScriptMethod Measure { param ($size) $this.Measures++; $this.DesiredSize.Height = $this.Needs }
        $content | Add-Member ScriptMethod InvalidateMeasure { }
        $ui.ResultListHost | Add-Member NoteProperty Content $content -Force
        $ui.ResultListHost | Add-Member ScriptMethod UpdateLayout { } -Force
        $ui.ResultListHost.ActualHeight = 300.0
        $ui.ResultListHost.ActualWidth = 500.0
        $detailRow = [pscustomobject]@{ ActualHeight = 180.0; Height = [System.Windows.GridLength]::new(180) }
        $ui | Add-Member NoteProperty DetailRow $detailRow -Force
    }
    AfterEach {
        $ui.ResultListHost.ActualHeight = 0.0
        $ui.ResultListHost.ActualWidth = 0.0
    }

    It "表の内容が低いとき、表を内容の高さまで縮めて、プレビューを広げる" {
        updateResultPaneHeight 2
        $ui.DetailRow.Height.Value | Should -Be 330
    }

    It "表に行がまだ無いとき（起動直後）は、広げずに元の高さのまま" {
        updateResultPaneHeight 0
        $ui.DetailRow.Height.Value | Should -Be 180
    }

    It "表の内容が増えて、表が足りなくなったら、プレビューを元の高さに戻す" {
        updateResultPaneHeight 2
        $ui.ResultListHost.Content.Needs = 400.0
        updateResultPaneHeight 20
        $ui.DetailRow.Height.Value | Should -Be 180
    }

    It "行が 40 ちょうどまでは内容を測る" {
        updateResultPaneHeight 40
        $ui.ResultListHost.Content.Measures | Should -Be 1
    }

    It "行が多いときは、内容を測らずに元の高さにする" {
        updateResultPaneHeight 41
        $ui.ResultListHost.Content.Measures | Should -Be 0
        $ui.DetailRow.Height.Value | Should -Be 180
    }

    It "利用者が境目を動かしたあとは、高さを変えない" {
        $script:previewHeightDragged = $true
        updateResultPaneHeight 2
        $ui.DetailRow.Height.Value | Should -Be 180
    }

    It "境目をドラッグしている最中は、大きさが変わっても高さを書き戻さない" {
        $ui.PreviewSplitter.IsDragging = $true
        try {
            updateResultPaneHeight 2
            $ui.DetailRow.Height.Value | Should -Be 180
        }
        finally { $ui.PreviewSplitter.IsDragging = $false }
    }

    It "境目のドラッグを終えたあと、<Name>" -TestCases @(
        @{ Name = "クリックしただけ（動かしていない）では、自動の高さ合わせを止めない"; Canceled = $false; VerticalChange = 0.0; Expected = 330 }
        @{ Name = "取り消したドラッグでは、自動の高さ合わせを止めない"; Canceled = $true; VerticalChange = -80.0; Expected = 330 }
        @{ Name = "動かして離したら、それ以後は自動で変えない"; Canceled = $false; VerticalChange = -80.0; Expected = 180 }
    ) {
        param ($Name, $Canceled, $VerticalChange, $Expected)
        & $handlers["PreviewSplitter.DragCompleted"] $ui.PreviewSplitter ([pscustomobject]@{ Canceled = $Canceled; VerticalChange = $VerticalChange })
        updateResultPaneHeight 2
        $ui.DetailRow.Height.Value | Should -Be $Expected
    }

    It "大きさがまだ取れないときは、何もしない" {
        $ui.ResultListHost.ActualWidth = 0.0
        updateResultPaneHeight 2
        $ui.DetailRow.Height.Value | Should -Be 180
    }
}

Describe "showDetail" -Tag Io {
    BeforeEach { resetPreview }

    It "行を選んでいなければ中身を空にする" {
        $ui.DetailTitle.Text = "前の行"
        showDetail
        $ui.DetailTitle.Text | Should -Be ""
        $ui.PreviewPlaceholder.Visibility | Should -Be "Visible"
    }

    It "Excel の行は前後の行をセルに分けて出し、見出しにセル番地を出す" {
        # 閉じている見出しを選んだときの先頭の行のように、まだ画面に出ていない（Prepare していない）行でもセル番地を出す
        $row = newHitRow "見積.xlsx" "4月" 3 @("`t品名`t金額", "`tりんご`t100", "`t見積書`t200", "`tみかん`t300")
        $row.Prepared | Should -Be $false
        $fake.Current = $row

        showDetail

        $ui.PreviewPlaceholder.Visibility | Should -Be "Collapsed"
        $ui.PreviewHeaderScroll.Visibility | Should -Be "Visible"
        $ui.OpenButton.IsEnabled | Should -Be $true
        $ui.OpenMenuButton.IsEnabled | Should -Be $true
        $ui.OpenFolderButton.IsEnabled | Should -Be $true
        $ui.DetailTitle.Text | Should -Be "営業部\見積.xlsx ・ [シート]4月!B3 ・ セル"
        $ui.DetailTitle.ToolTip | Should -Be $ui.DetailTitle.Text
        @($ui.PreviewRows.ItemsSource).Count | Should -Be 4
        @($ui.PreviewHeader.ItemsSource).Count | Should -Be 3
        $script:previewTable.Rows.Count | Should -Be 4
        # 一致したセルが左端から見えているので、横にはスクロールしない
        $fake.Scrolls[-1] | Should -Be 0
    }

    It "Excel 以外は行番号を出し、ボタンは「開く」にする" {
        $fake.Current = newHitRow "議事録.docx" "ページ001" 2 @("はじめに", "見積の件", "おわりに") ""

        showDetail

        $ui.DetailTitle.Text | Should -Be "議事録.docx ・ 1 ページ（目安） ・ 本文"
    }

    It "Word の 1 列は、行番号を除いた見える幅（400 - 44）いっぱいにする" {
        $fake.Current = newHitRow "議事録.docx" "ページ001" 1 @("見積の件") ""

        showDetail

        $script:previewTable.IsExcel | Should -Be $false
        $script:previewTable.Columns[0].Width | Should -Be 356
    }

    It "Excel の列は中身の幅のまま（枠いっぱいにしない）" {
        $fake.Current = newHitRow "見積.xlsx" "4月" 1 @("`t見積の件") "営業部"

        showDetail

        $script:previewTable.IsExcel | Should -Be $true
        $script:previewTable.Columns[0].Width | Should -Not -Be 356
        $script:previewTable.Columns[0].Width | Should -BeLessThan 356
    }

    It "読んでいる間に別の行を選んだら、読み終えた古い行の結果は出さない" {
        $fake.Current = newHitRow "見積.xlsx" "4月" 1 @("`t見積", "`t次")
        $other = newHitRow "請求.xlsx" "5月" 1 @("`t請求")
        $fake.BeforeDone = { $fake.Current = $other }
        try {
            showDetail
        } finally {
            $fake.BeforeDone = $null
        }
        $ui.DetailTitle.Text | Should -Be ""
        $script:previewTable | Should -BeNullOrEmpty
    }

    It "後から頼んだ読み込みがあれば、先に頼んだ分の結果は捨てる" {
        $fake.Current = newHitRow "見積.xlsx" "4月" 1 @("`t見積")
        $fake.BeforeDone = { $script:previewRequest++ }
        try {
            showDetail
        } finally {
            $fake.BeforeDone = $null
        }
        $script:previewTable | Should -BeNullOrEmpty
    }

    It "最後に読んだ 1 つのファイルの控えを次の読みに渡す。別のファイルに移れば控えは入れ替わり、選び直したら捨てる" {
        $seen = [System.Collections.Generic.List[object]]::new()
        $originalLines = Get-Command readContentIndexLines -CommandType Function
        Mock readContentIndexLines { $seen.Add($marks); & $originalLines $path $book $location $firstLine $count $cache $marks }
        $lines = @("`t一", "`t二", "`t三", "`t四", "`t五")
        $first = newHitRow "見積.xlsx" "4月" 2 $lines
        $second = [HitRow]::Create("営業部", $first.Root, $first.RelPath, $first.RelDir, $first.Book, $first.Book, $first.Location, 4, $lines[3], "見積", [regex]"見積")
        $second.IsExcel = $first.IsExcel
        $other = newHitRow "請求.xlsx" "5月" 2 $lines "経理部"

        $fake.Current = $first
        showDetail
        $seen[0] | Should -BeNullOrEmpty
        $script:previewMarks.Path | Should -Be (toLongPath ([System.IO.Path]::Combine($first.Root, $first.RelPath)))

        $fake.Current = $second
        showDetail
        $seen[1].Path | Should -Be $script:previewMarks.Path

        $fake.Current = $other
        showDetail
        $seen[2].Path | Should -Not -Be (toLongPath ([System.IO.Path]::Combine($other.Root, $other.RelPath)))
        $script:previewMarks.Path | Should -Be (toLongPath ([System.IO.Path]::Combine($other.Root, $other.RelPath)))

        $fake.Current = $null
        showDetail
        $script:previewMarks | Should -BeNullOrEmpty
        $fake.Current = $second
        showDetail
        $seen[3] | Should -BeNullOrEmpty
    }

    Context "前後の行を読めないとき" {
        It "選んだ行だけを出す" {
            Mock readContentIndexLines { throw "読めません" }
            $fake.Current = newHitRow "議事録.docx" "ページ001" 2 @("はじめに", "見積の件", "おわりに") ""

            showDetail

            $script:previewTable.Rows.Count | Should -Be 1
            $ui.DetailTitle.Text | Should -Be "議事録.docx ・ 1 ページ（目安） ・ 本文"
        }
    }

    It "列が多くても全部の列を持ち、見えている列の見出しとセルだけを作る" {
        $cells = @(1..250 | ForEach-Object { "値$_" })
        $cells[229] = "見積"
        $fake.Current = newHitRow "見積.xlsx" "4月" 1 @("`t" + ($cells -join "`t"))

        showDetail

        $table = $script:previewTable
        $table.Columns.Count | Should -Be 251
        @($ui.PreviewHeader.ItemsSource).Count | Should -BeLessThan 251
        $ui.PreviewHeaderPanel.DataContext | Should -Be $table.Window
        @($table.Rows | ForEach-Object { $_.Cells.Count } | Measure-Object -Sum).Sum | Should -BeLessThan 251
    }

    It "一致したセルが見えないときは、その少し左まで横にスクロールする" {
        $cells = @(1..30 | ForEach-Object { "値$_" })
        $cells[25] = "見積"
        $fake.Current = newHitRow "見積.xlsx" "4月" 1 @("`t" + ($cells -join "`t"))

        showDetail

        $table = $script:previewTable
        $fake.Scrolls[-1] | Should -Be ([math]::Max(0, $table.HitOffset - 120))
        $fake.Scrolls[-1] | Should -BeGreaterThan 0
    }

    It "最初は選択行の前後 50 行ずつ（101 行）を読み、選択行が上から 3 行目に来るよう縦にスクロールする" {
        $lines = @(1..200 | ForEach-Object { "`t行$_" })
        $lines[119] = "`t見積"
        $fake.Current = newHitRow "見積.xlsx" "4月" 120 $lines

        showDetail

        $rows = $script:previewTable.Rows
        $rows.Count | Should -Be 101
        $rows[0].Number | Should -Be "70"
        $rows[100].Number | Should -Be "170"
        $rows[50].IsHitRow | Should -Be $true
        $fake.VerticalScrolls[-1] | Should -Be 48
    }

    It "先頭に近い行は、前の行が無い分だけ読み、先頭から縦に出す" {
        $fake.Current = newHitRow "見積.xlsx" "4月" 2 @("`t前", "`t見積", "`t後")

        showDetail

        $script:previewTable.Rows.Count | Should -Be 3
        $fake.VerticalScrolls[-1] | Should -Be 0
    }

    It "検索のあとにインデックスが短くなり、選んだ行が無くなっていても落ちない" {
        $row = newHitRow "見積.xlsx" "4月" 5 @("`ta", "`tb", "`tc", "`td", "`t見積")
        writeHitContentIndex ([System.IO.Path]::Combine($row.Root, $row.RelPath)) $row.Book $row.Location @("`ta", "`tb")
        $fake.Current = $row

        showDetail

        @($script:previewTable.Rows | Where-Object { $_.IsHitRow }).Count | Should -Be 0
        $script:previewTable.Rows.Count | Should -Be 2
        $fake.Scrolls[-1] | Should -Be 0
        $fake.VerticalScrolls[-1] | Should -Be 0
    }

    It "インデックスの本文インデックスのファイルが読めなければ、選んだ行だけを出す" {
        $row = newHitRow "見積.xlsx" "4月" 2 @("`t前", "`t見積", "`t後")
        Remove-Item -LiteralPath ([System.IO.Path]::Combine($row.Root, $row.RelPath))
        $fake.Current = $row

        showDetail

        $script:previewTable.Rows.Count | Should -Be 1
        $script:previewTable.Rows[0].Number | Should -Be "2"
    }
}

Describe "プレビューの右クリックメニュー" -Tag Unit {
    BeforeEach {
        resetPreview
        $fake.Opened = 0
        $fake.MenuItems = $null
    }

    It "プレビューが無いときは、メニューを出さない" {
        $script:previewTable = $null
        $e = [pscustomobject]@{ Handled = $false }

        & $handlers["PreviewRows.ContextMenuOpening"] $null $e

        $e.Handled | Should -Be $true
        $fake.MenuItems | Should -BeNullOrEmpty
    }

    It "プレビューがあるときは、判断層の並びでメニューを組む" {
        $script:previewTable = [PreviewTable]::new()
        $e = [pscustomobject]@{ Handled = $false }

        & $handlers["PreviewRows.ContextMenuOpening"] $null $e

        $e.Handled | Should -Be $false
        ($fake.MenuItems | ForEach-Object { $_.Id }) -join "/" | Should -Be "openHere/separator/copyCell/copyRow"
        $fake.MenuParts["openHere"].PartName | Should -Be "MenuPreviewOpen"
        $fake.MenuParts["copyCell"].PartName | Should -Be "MenuPreviewCopy"
        $fake.MenuParts["copyRow"].PartName | Should -Be "MenuPreviewCopyRow"
    }

    It "［元のファイルのこの場所を開く］は、選んでいる結果の行を開く処理を呼ぶ" {
        & $handlers["MenuPreviewOpen.Click"]

        $fake.Opened | Should -Be 1
    }
}

Describe "copyPreviewSelection" -Tag Unit {
    BeforeEach { resetPreview }

    It "プレビューが無いときは、セルを選ぶよう案内する" {
        copyPreviewSelection
        $fake.Status | Should -BeExactly "プレビューでコピーするセルをクリックしてください。"
    }

    It "<name>" -TestCases @(
        @{ name = "セルを選んでいないときも、案内だけを出す"; menu = "MenuPreviewCopy.Click" }
        @{ name = "［行をコピー］もセルを選んでいなければ案内だけを出す"; menu = "MenuPreviewCopyRow.Click" }
    ) {
        param ($name, $menu)
        $script:previewTable = [PreviewTable]::new()

        & $handlers[$menu]

        $fake.Status | Should -BeExactly "プレビューでコピーするセルをクリックしてください。"
    }
}

Describe "イベント" -Tag Unit {
    BeforeEach { resetPreview }

    It "別の行を選ぶと、タイマーをかけ直す" {
        & $handlers["ResultGrid.SelectionChanged"]

        $fake.TimerStops | Should -Be 1
        $fake.TimerStarts | Should -Be 1
    }

    It "タイマーが来たら止めてプレビューを出す" {
        $fake.Current = $null
        $ui.DetailTitle.Text = "前の行"

        & $handlers["Timer.Tick"]

        $fake.TimerStops | Should -Be 1
        $ui.DetailTitle.Text | Should -Be ""
    }

    It "<name>" -TestCases @(
        @{ name = "見える幅が変わると、列の幅を合わせ直す"; views = @(400); manual = $null; excel = $false; expected = 356 }
        @{ name = "見える幅が狭まると、列の幅も狭まる"; views = @(400, 300); manual = $null; excel = $false; expected = 256 }
        @{ name = "狭まりきったら、初めの幅（100）まで戻る"; views = @(400, 120); manual = $null; excel = $false; expected = 100 }
        @{ name = "手で変えた幅は、見える幅が変わっても保つ"; views = @(400, 300, 290, 280); manual = 500; excel = $false; expected = 500 }
        @{ name = "Excel の列は合わせ直さない"; views = @(400); manual = $null; excel = $true; expected = 100 }
    ) {
        param ($name, $views, $manual, $excel, $expected)
        $column = [PreviewColumn]::new()
        $column.Width = 100
        $table = [PreviewTable]::new()
        $table.Columns = [System.Collections.Generic.List[PreviewColumn]]::new()
        $table.Columns.Add($column)
        $table.IsExcel = $excel
        $script:previewTable = $table
        $script:previewBaseWidth = 100
        $script:previewFitWidth = 100

        foreach ($view in $views) {
            $ui.PreviewScroll.ViewportWidth = [double]$view
            & $handlers["PreviewScroll.ScrollChanged"] $ui.PreviewScroll ([pscustomobject]@{ ViewportWidthChange = -17.0 })
            if ($null -ne $manual -and $view -eq $views[0]) {
                $column.SetWidth($manual)
            }
        }

        $column.Width | Should -Be $expected
    }

    It "見える幅が変わらないスクロールでは、列の幅を合わせ直さない" {
        $column = [PreviewColumn]::new()
        $column.Width = 100
        $table = [PreviewTable]::new()
        $table.Columns = [System.Collections.Generic.List[PreviewColumn]]::new()
        $table.Columns.Add($column)
        $script:previewTable = $table
        $script:previewBaseWidth = 100
        $script:previewFitWidth = 100

        & $handlers["PreviewScroll.ScrollChanged"] $ui.PreviewScroll ([pscustomobject]@{ ViewportWidthChange = 0.0 })

        $column.Width | Should -Be 100
    }

    It "行を横にスクロールしたら、列見出しも同じ位置にする" {
        $ui.PreviewScroll.HorizontalOffset = 150.0

        & $handlers["PreviewScroll.ScrollChanged"] $ui.PreviewScroll ([pscustomobject]@{ ViewportWidthChange = 0.0 })

        $fake.HeaderScrolls[-1] | Should -Be 150.0
    }

    It "Ctrl+C 以外のキーではコピーしない" {
        $e = [pscustomobject]@{ Key = "A"; Handled = $false }

        & $handlers["PreviewScroll.PreviewKeyDown"] $ui.PreviewScroll $e

        $e.Handled | Should -Be $false
        $fake.Status | Should -Be ""
    }

    It "セルの外をクリック・ドラッグしても選ばない" {
        $script:previewTable = [PreviewTable]::new()
        $down = [pscustomobject]@{ OriginalSource = $null; LeftButton = "Pressed" }

        & $handlers["PreviewRows.PreviewMouseLeftButtonDown"] $ui.PreviewRows $down
        & $handlers["PreviewRows.MouseMove"] $ui.PreviewRows $down
        & $handlers["PreviewRows.PreviewMouseRightButtonDown"] $ui.PreviewRows $down

        $script:previewTable.HasSelection() | Should -Be $false
    }

    It "列見出しの右端をドラッグすると、その列の幅が変わる" {
        $column = [PreviewColumn]::new()
        $column.Width = 100
        $e = New-Object System.Windows.Controls.Primitives.DragDeltaEventArgs 30.0, 0.0
        $e.RoutedEvent = [System.Windows.Controls.Primitives.Thumb]::DragDeltaEvent
        $e.Source = [pscustomobject]@{ DataContext = $column }

        $handlers["PreviewHeader.DragDelta"].Invoke($ui.PreviewHeader, $e)

        $column.Width | Should -Be 130
    }

    It "列が多い表で見出しをドラッグしている間は、見えている列の範囲も見出しも替えず、終わったら範囲を決め直す" {
        $cells = @(1..60 | ForEach-Object { "列$_" }) -join "`t"
        $fake.Current = newHitRow "見積.xlsx" "4月" 1 @("`t$cells")
        showDetail
        $table = $script:previewTable
        $table.Columns.Count | Should -BeGreaterThan 20
        $lastBefore = $table.Window.Last
        $headerBefore = $ui.PreviewHeader.ItemsSource
        $versionBefore = $table.WindowVersion

        $handlers["PreviewHeader.DragStarted"].Invoke($ui.PreviewHeader, (New-Object System.Windows.Controls.Primitives.DragStartedEventArgs 0.0, 0.0))
        $e = New-Object System.Windows.Controls.Primitives.DragDeltaEventArgs 2000.0, 0.0
        $e.RoutedEvent = [System.Windows.Controls.Primitives.Thumb]::DragDeltaEvent
        $e.Source = [pscustomobject]@{ DataContext = $table.Columns[0] }
        $handlers["PreviewHeader.DragDelta"].Invoke($ui.PreviewHeader, $e)

        $table.Columns[0].Width | Should -BeGreaterThan 2000
        $table.Window.Last | Should -Be $lastBefore
        $table.WindowVersion | Should -Be $versionBefore
        [object]::ReferenceEquals($ui.PreviewHeader.ItemsSource, $headerBefore) | Should -Be $true

        $handlers["PreviewHeader.DragCompleted"].Invoke($ui.PreviewHeader, (New-Object System.Windows.Controls.Primitives.DragCompletedEventArgs 0.0, 0.0, $false))

        $table.Window.Last | Should -BeLessThan $lastBefore
        $table.WindowVersion | Should -BeGreaterThan $versionBefore
        [object]::ReferenceEquals($ui.PreviewHeader.ItemsSource, $headerBefore) | Should -Be $false
    }

    It "見出しをドラッグしている間に別の行の表に替わっても、新しい表の見出しとセルの範囲が決まる" {
        $cells = @(1..60 | ForEach-Object { "列$_" }) -join "`t"
        $fake.Current = newHitRow "見積.xlsx" "4月" 1 @("`t$cells")
        showDetail
        $handlers["PreviewHeader.DragStarted"].Invoke($ui.PreviewHeader, (New-Object System.Windows.Controls.Primitives.DragStartedEventArgs 0.0, 0.0))

        # ドラッグの終わりが届かないまま、別の行を選んで表が替わる
        $fake.Current = newHitRow "別.xlsx" "5月" 1 @("`t$cells") "経理部"
        showDetail

        $table = $script:previewTable
        $table.Window.Last | Should -BeGreaterThan 0
        @($ui.PreviewHeader.ItemsSource).Count | Should -BeGreaterThan 0
        $table.Rows[0].Cells.Count | Should -BeGreaterThan 0
    }
}

Describe "プレビューの読み足し（スクロールに合わせて、上と下に行を足す）" -Tag Io {
    BeforeEach {
        resetPreview
        $fake.BeforeDone = $null
        # 1,000 行のファイルの 500 行目を選んだ表（450〜550 行の 101 行）
        $script:linesA = @(1..1000 | ForEach-Object { "`t行$_" })
        $script:rowA = newHitRow "見積.xlsx" "4月" 500 $script:linesA
        $script:linesB = @(1..1000 | ForEach-Object { "`t別$_" })
        $script:rowB = newHitRow "別.xlsx" "5月" 500 $script:linesB "経理部"
        $fake.Current = $script:rowA
        showDetail
        $script:tableA = $script:previewTable
    }
    AfterEach { $fake.BeforeDone = $null }

    BeforeAll {
        function newMoreResult {
            # 読み足しの結果（startReadMore の仕事が返す形）。first から last までの行
            param ($table, [string]$side, [int]$first, [int]$last, [bool]$end = $false, [string]$prefix = "行")
            $rows = @()
            if ($first -le $last) { $rows = @($first..$last | ForEach-Object { [pscustomobject]@{ LineNumber = $_; Line = "`t$prefix$_" } }) }
            return @{ Id = $table.Request; Side = $side; Failed = $false; First = $first; Rows = $rows; End = $end; Marks = $null }
        }
    }

    It "最初は選択行の前後 50 行ずつ（101 行）を読み、表に読み込みの番号を持たせる" {
        $script:tableA.Rows.Count | Should -Be 101
        $script:tableA.FirstLine() | Should -Be 450
        $script:tableA.LastLine() | Should -Be 550
        $script:tableA.Request | Should -Be $script:previewRequest
    }

    It "読み足しの結果の控えを、表と次の読みに使う控えに入れる" {
        $result = newMoreResult $script:tableA "down" 551 560
        $result.Marks = @{ Path = "控え" }
        applyReadMore @($result) $null
        $script:tableA.Marks.Path | Should -Be "控え"
        $script:previewMarks.Path | Should -Be "控え"
    }

    It "下に足した行が今の行につながるなら足す。終わりまで読んだときだけ、下は終わりにする" {
        applyReadMore @(newMoreResult $script:tableA "down" 551 560) $null
        $script:tableA.LastLine() | Should -Be 560
        $script:tableA.BottomDone | Should -Be $false

        applyReadMore @(newMoreResult $script:tableA "down" 561 570 $true) $null
        $script:tableA.LastLine() | Should -Be 570
        $script:tableA.BottomDone | Should -Be $true
    }

    It "今の行につながらない結果は足さない。終わりまで読んでいなければ終わりにせず、また近づいたときに読み直す" {
        applyReadMore @(newMoreResult $script:tableA "down" 700 710) $null
        $script:tableA.LastLine() | Should -Be 550
        $script:tableA.BottomDone | Should -Be $false

        applyReadMore @(newMoreResult $script:tableA "up" 300 310) $null
        $script:tableA.FirstLine() | Should -Be 450
        $script:tableA.TopDone | Should -Be $false
    }

    It "上につながらず、終わりまで読んだ結果は、上を終わりにする（同じ範囲を頼み直し続けない）" {
        applyReadMore @(newMoreResult $script:tableA "up" 300 310 $true) $null
        $script:tableA.FirstLine() | Should -Be 450
        $script:tableA.TopDone | Should -Be $true
    }

    It "頼んだあとに下の端が動いた古い結果は、End でも空でも下を終わりにしない（捨てて戻した終わりの印を立て直さない）" -ForEach @(
        @{ name = "End"; first = 700; last = 710; end = $true }
        @{ name = "空"; first = 700; last = 699; end = $true }
        @{ name = "空（End なし）"; first = 700; last = 699; end = $false }
    ) {
        applyReadMore @(newMoreResult $script:tableA "down" $first $last $end) $null
        $script:tableA.LastLine() | Should -Be 550
        $script:tableA.BottomDone | Should -Be $false
    }

    It "つながらず終わりまで読んだ向きは、続けてスクロールしても頼み直さない" {
        $script:asked = 0
        $fake.BeforeDone = { $script:asked++ }
        $ui.PreviewScroll.VerticalOffset = 0.0   # 上の端にいる
        applyReadMore @(newMoreResult $script:tableA "up" 300 310 $true) $null
        $script:asked = 0
        updatePreviewReadMore
        updatePreviewReadMore
        $script:asked | Should -Be 0

        # 対照: 終わりにしていなければ、同じ位置で頼む
        $script:tableA.TopDone = $false
        updatePreviewReadMore
        $script:asked | Should -BeGreaterThan 0
    }

    It "頼んだときのままで行が 1 つも無い結果は、End の有無にかかわらず、その向きにもう行が無いことにする" -ForEach @(
        @{ end = $true }
        @{ end = $false }
    ) {
        $empty = newMoreResult $script:tableA "down" 551 550 $end
        $empty.Rows.Count | Should -Be 0
        applyReadMore @($empty) $null
        $script:tableA.LastLine() | Should -Be 550
        $script:tableA.BottomDone | Should -Be $true
    }

    It "読み足している間に別の行を選んで表が替わったら、古い結果は新しい表に入れず、終わりの印も付けない" {
        $fake.BeforeDone = {
            $fake.BeforeDone = $null
            $fake.Current = $script:rowB
            showDetail
        }
        startReadMore @{ Side = "down"; FirstLine = 551; Count = 200 }

        $tableB = $script:previewTable
        [object]::ReferenceEquals($tableB, $script:tableA) | Should -Be $false
        $tableB.Rows.Count | Should -Be 101
        $tableB.LastLine() | Should -Be 550
        $tableB.BottomDone | Should -Be $false
        @($tableB.Rows | Where-Object { ($_.Texts -join "") -like "*行*" }).Count | Should -Be 0
        # 古い表にも足されない（表を替えたあとに届いた結果は捨てる）
        $script:tableA.LastLine() | Should -Be 550
    }

    It "表が替わる前（選び直したあと、読み終わる前）に読み足すときは、古い表の行と番号で読み、古い表に足す" {
        # showDetail が番号と行を替えたあと、applyDetail で表を替えるまでの間
        $script:previewRequest++
        $script:previewRow = $script:rowB

        updatePreviewReadMore
        startReadMore @{ Side = "down"; FirstLine = 551; Count = 5 }

        $script:tableA.LastLine() | Should -Be 555
        $script:tableA.Rows[$script:tableA.Rows.Count - 1].Texts -join "" | Should -BeLike "*行555*"
    }

    It "上限を超えて下を捨てるとき、下の終わりの印を戻す（また下に近づいたら読む）" {
        applyReadMore @(newMoreResult $script:tableA "down" 551 3600 $true) $null
        $script:tableA.DataCount() | Should -Be 2951
        $script:tableA.FirstLine() | Should -Be 450
        $script:tableA.LastLine() | Should -Be 3400
        $script:tableA.BottomDone | Should -Be $false
    }

    It "上限を超えて上を捨てるとき、上の終わりの印を戻し、見ていた行が同じ所に見えるよう縦の位置を直す" {
        $script:tableA.TopDone = $true
        $ui.PreviewScroll.VerticalOffset = 2900.0
        applyReadMore @(newMoreResult $script:tableA "down" 551 3600) $null
        $script:tableA.FirstLine() | Should -Be 650
        $script:tableA.TopDone | Should -Be $false
        $fake.VerticalScrolls[-1] | Should -Be 2700
    }

    It "上に足したあとも、見ていた行が同じ所に見える（読み込み中の行を見ていた先頭でも。見ていた位置 <offset>）" -ForEach @(
        @{ offset = 0; expected = 99 }
        @{ offset = 5; expected = 104 }
    ) {
        $script:tableA.ShowMessage("up", "読み込み中…")
        $ui.PreviewScroll.VerticalOffset = [double]$offset
        applyReadMore @(newMoreResult $script:tableA "up" 350 449) $null
        $script:tableA.FirstLine() | Should -Be 350
        $script:tableA.LeadCount() | Should -Be 0
        $fake.VerticalScrolls[-1] | Should -Be $expected -Because "足す前に 450 行目（または見ていた行）が見えていた所に、足したあとも同じ行が見える"
    }

    It "読み足しが失敗したら、その向きを終わりにして「読めませんでした」を出す" {
        $script:tableA.ShowMessage("down", "読み込み中…")
        applyReadMore @(@{ Id = $script:tableA.Request; Side = "down"; Failed = $true }) $null
        $script:tableA.BottomDone | Should -Be $true
        $script:tableA.DownMessage.Message | Should -Be "読めませんでした"
    }
}
