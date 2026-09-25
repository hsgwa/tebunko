# ［2 検索］の選択行のプレビュー（tebunko_grep\ui\preview.ps1）のテスト。
# 画面の部品（$ui.PreviewRows など）は偽物にして、表示する中身と状態の変化を確かめる。
# クリップボードは利用者の PC のものを書き換えるため、コピーは「選んでいないとき」だけを確かめる。
. "$PSScriptRoot\..\..\helpers\load.ps1"
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
. "${scriptsDir}\shared\ui\types.ps1"
. "${scriptsDir}\tebunko_grep\ui\types_grep.ps1"
. "${scriptsDir}\tebunko_grep\ui\preview_view.ps1"

# gui.ps1 で決める値
${previewRowHeight}     = 22
${previewScrollBarSize} = 18
${maxPreviewRows}       = 101

# ---- 画面の偽物 ----
# 読み込み時に登録されるイベントの処理は $handlers に、呼ばれた操作は $fake に取っておく
$handlers = @{}
$fake = @{ Status = ""; Current = $null; TimerStarts = 0; TimerStops = 0; Scrolls = @(); HeaderScrolls = @() }

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

$previewScroll = newFakePart "PreviewScroll" @{ ViewportHeight = 220.0; ActualHeight = 0.0; ViewportWidth = 400.0; HorizontalOffset = 0.0 } @("ScrollChanged", "SizeChanged", "PreviewKeyDown")
$previewScroll | Add-Member ScriptMethod UpdateLayout { }
$previewScroll | Add-Member ScriptMethod Focus { $true }
$previewScroll | Add-Member ScriptMethod ScrollToHorizontalOffset { param ($offset) $fake.Scrolls += $offset }
$headerScroll = newFakePart "PreviewHeaderScroll" @{ Visibility = "Collapsed" }
$headerScroll | Add-Member ScriptMethod ScrollToHorizontalOffset { param ($offset) $fake.HeaderScrolls += $offset }
$previewHeader = newFakePart "PreviewHeader" @{ ItemsSource = $null }
$previewHeader | Add-Member ScriptMethod AddHandler { param ($event, $handler) $handlers["PreviewHeader.DragDelta"] = $handler }

$ui = [pscustomobject]@{
    ResultGrid          = newFakePart "ResultGrid" @{ SelectedItem = $null } @("SelectionChanged")
    PreviewHeader       = $previewHeader
    PreviewRows         = newFakePart "PreviewRows" @{ ItemsSource = $null } @("PreviewMouseLeftButtonDown", "MouseMove", "PreviewMouseRightButtonDown")
    PreviewScroll       = $previewScroll
    PreviewHeaderScroll = $headerScroll
    PreviewNote         = newFakePart "PreviewNote" @{ Text = ""; ToolTip = $null; Visibility = "Collapsed" }
    PreviewPlaceholder  = newFakePart "PreviewPlaceholder" @{ Visibility = "Visible" }
    DetailTitle         = newFakePart "DetailTitle" @{ Text = ""; ToolTip = $null }
    OpenButton          = newFakePart "OpenButton" @{ IsEnabled = $false; Content = "" }
    OpenFolderButton    = newFakePart "OpenFolderButton" @{ IsEnabled = $false }
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

. "${scriptsDir}\tebunko_grep\ui\preview.ps1"

# ---- テストの準備 ----

function resetPreview {
    $fake.Status = ""
    $fake.Current = $null
    $fake.TimerStarts = 0
    $fake.TimerStops = 0
    $fake.Scrolls = @()
    $fake.HeaderScrolls = @()
    $ui.PreviewScroll.ViewportHeight = 220.0
    $ui.PreviewScroll.ActualHeight = 0.0
    $ui.PreviewScroll.ViewportWidth = 400.0
    $ui.ResultGrid.SelectedItem = $null
    $script:detailKeepScroll = $false
    clearDetail
}

function writeHitPack {
    # インデックスの集約ファイル（TestDrive に作る）に、元のファイル book の場所 location の行を書く
    param ([string]$path, [string]$book, [string]$location, [string[]]$lines)
    [void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path))
    writePackFile $path (convertToPackText @(@{ Name = $book; Places = @(@{ Place = $location; Text = (($lines -join "`r`n") + "`r`n") }) }))
}

function newHitRow {
    # インデックスの集約ファイル（TestDrive に作る）の lineNumber 行目にヒットした行
    param (
        [string]$book,
        [string]$location,
        [int]$lineNumber,
        [string[]]$lines,
        [string]$relDir = "営業部"
    )

    $root = "$TestDrive\index"
    $name = getPackFileName (getPackExtension $book)
    $relPath = if ($relDir) { "$relDir\$name" } else { $name }
    writeHitPack "$root\$relPath" $book $location $lines
    $row = [HitRow]::Create("営業部", $root, $relPath, $relDir, $book, $book, $location, $lineNumber, $lines[$lineNumber - 1], "見積", [regex]"見積")
    $described = describePlace $book $location
    $row.PlaceText = $described.Place
    $row.Kind = $described.Kind
    return $row
}

Describe "読み込み" -Tag Unit {
    It "選択が変わったときにまとめて読むタイマーと、イベントを登録する" {
        $handlers.ContainsKey("Timer.Tick") | Should Be $true
        $handlers.ContainsKey("ResultGrid.SelectionChanged") | Should Be $true
        $handlers.ContainsKey("PreviewRows.PreviewMouseLeftButtonDown") | Should Be $true
        $handlers.ContainsKey("MenuPreviewCopyRow.Click") | Should Be $true
        $handlers["PreviewHeader.DragDelta"] -is [System.Windows.Controls.Primitives.DragDeltaEventHandler] | Should Be $true
    }
}

Describe "clearDetail" -Tag Unit {
    BeforeEach { resetPreview }

    It "中身を空にして案内を出し、開くボタンを使えなくする" {
        $ui.DetailTitle.Text = "前の行"
        $ui.PreviewHeader.ItemsSource = @(1)
        $script:previewTable = [PreviewTable]::new()

        clearDetail

        $ui.DetailTitle.Text | Should Be ""
        $null -eq $ui.PreviewHeader.ItemsSource | Should Be $true
        $null -eq $script:previewTable | Should Be $true
        $ui.PreviewPlaceholder.Visibility | Should Be "Visible"
        $ui.PreviewHeaderScroll.Visibility | Should Be "Collapsed"
        $ui.OpenButton.IsEnabled | Should Be $false
        $ui.OpenFolderButton.IsEnabled | Should Be $false
    }
}

Describe "getPreviewContextLines" -Tag Unit {
    BeforeEach { resetPreview }

    It "見えている高さから、前後に読む行数を決める" {
        $ui.PreviewScroll.ViewportHeight = 220.0
        $counts = getPreviewContextLines
        $counts[0] | Should Be 4
        $counts[1] | Should Be 5
    }

    It "見えている高さが取れないときは、全体の高さから横スクロールバーの分を引く" {
        $ui.PreviewScroll.ViewportHeight = 0.0
        $ui.PreviewScroll.ActualHeight = 260.0
        $counts = getPreviewContextLines
        ($counts[0] + $counts[1] + 1) | Should Be 11
    }

    It "高さがまったく取れないときは、選んだ行だけを読む" {
        $ui.PreviewScroll.ViewportHeight = 0.0
        $ui.PreviewScroll.ActualHeight = 0.0
        $counts = getPreviewContextLines
        $counts[0] | Should Be 0
        $counts[1] | Should Be 0
    }
}

Describe "showDetail" -Tag Io {
    BeforeEach { resetPreview }

    It "行を選んでいなければ中身を空にする" {
        $ui.DetailTitle.Text = "前の行"
        showDetail
        $ui.DetailTitle.Text | Should Be ""
        $ui.PreviewPlaceholder.Visibility | Should Be "Visible"
    }

    It "Excel の行は前後の行をセルに分けて出し、見出しにセル番地を出す" {
        # 閉じている見出しを選んだときの先頭の行のように、まだ画面に出ていない（Prepare していない）行でもセル番地を出す
        $row = newHitRow "見積.xlsx" "4月" 3 @("`t品名`t金額", "`tりんご`t100", "`t見積書`t200", "`tみかん`t300")
        $row.Prepared | Should Be $false
        $fake.Current = $row

        showDetail

        $ui.PreviewPlaceholder.Visibility | Should Be "Collapsed"
        $ui.PreviewHeaderScroll.Visibility | Should Be "Visible"
        $ui.OpenButton.IsEnabled | Should Be $true
        $ui.OpenFolderButton.IsEnabled | Should Be $true
        $ui.OpenButton.Content | Should Be "Excel で開く"
        $ui.DetailTitle.Text | Should Be "営業部\見積.xlsx ・ [シート] 4月 ・ セル ・ セル B3"
        $ui.DetailTitle.ToolTip | Should Be $ui.DetailTitle.Text
        $ui.PreviewNote.Visibility | Should Be "Collapsed"
        @($ui.PreviewRows.ItemsSource).Count | Should Be 4
        @($ui.PreviewHeader.ItemsSource).Count | Should Be 3
        $script:previewTable.Rows.Count | Should Be 4
        # 一致したセルが左端から見えているので、横にはスクロールしない
        $fake.Scrolls[-1] | Should Be 0
    }

    It "Excel 以外は行番号を出し、ボタンは「開く」にする" {
        $fake.Current = newHitRow "議事録.docx" "ページ001" 2 @("はじめに", "見積の件", "おわりに") ""

        showDetail

        $ui.OpenButton.Content | Should Be "開く"
        $ui.DetailTitle.Text | Should Be "議事録.docx ・ [ページ] 1（目安） ・ 本文 ・ 2 行目"
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
        $ui.DetailTitle.Text | Should Be ""
        $script:previewTable | Should BeNullOrEmpty
    }

    It "後から頼んだ読み込みがあれば、先に頼んだ分の結果は捨てる" {
        $fake.Current = newHitRow "見積.xlsx" "4月" 1 @("`t見積")
        $fake.BeforeDone = { $script:previewRequest++ }
        try {
            showDetail
        } finally {
            $fake.BeforeDone = $null
        }
        $script:previewTable | Should BeNullOrEmpty
    }

    # Pester 3 の Mock は Describe の中の後のテストにも効くため、Context で囲む
    Context "前後の行を読めないとき" {
        It "選んだ行だけを出す" {
            Mock readPackContext { throw "読めません" }
            $fake.Current = newHitRow "議事録.docx" "ページ001" 2 @("はじめに", "見積の件", "おわりに") ""

            showDetail

            $script:previewTable.Rows.Count | Should Be 1
            $ui.DetailTitle.Text | Should Be "議事録.docx ・ [ページ] 1（目安） ・ 本文 ・ 2 行目"
        }
    }

    It "列が多すぎるときは、表示した列の範囲を知らせる" {
        $cells = @(1..250 | ForEach-Object { "値$_" })
        $cells[229] = "見積"
        $fake.Current = newHitRow "見積.xlsx" "4月" 1 @("`t" + ($cells -join "`t"))

        showDetail

        $ui.PreviewNote.Visibility | Should Be "Visible"
        $ui.PreviewNote.Text | Should Be "表示は $($script:previewTable.RangeLabel) の 200 列（全 251 列）"
        $ui.PreviewNote.ToolTip | Should Match "一致したセルを中心に表示しています"
    }

    It "一致したセルが見えないときは、その少し左まで横にスクロールする" {
        $cells = @(1..30 | ForEach-Object { "値$_" })
        $cells[25] = "見積"
        $fake.Current = newHitRow "見積.xlsx" "4月" 1 @("`t" + ($cells -join "`t"))

        showDetail

        $table = $script:previewTable
        $fake.Scrolls[-1] | Should Be ([math]::Max(0, $table.HitOffset - 120))
        $fake.Scrolls[-1] | Should BeGreaterThan 0
    }

    It "高さを変えただけのときは、横の位置をそのままにする" {
        $fake.Current = newHitRow "見積.xlsx" "4月" 1 @("`t見積")
        $script:detailKeepScroll = $true

        showDetail

        $fake.Scrolls.Count | Should Be 0
        $script:detailKeepScroll | Should Be $false
    }

    It "検索のあとにインデックスが短くなり、選んだ行が無くなっていても落ちない" {
        $row = newHitRow "見積.xlsx" "4月" 5 @("`ta", "`tb", "`tc", "`td", "`t見積")
        writeHitPack ([System.IO.Path]::Combine($row.Root, $row.RelPath)) $row.Book $row.Location @("`ta", "`tb")
        $fake.Current = $row

        showDetail

        @($script:previewTable.Rows | Where-Object { $_.IsHitRow }).Count | Should Be 0
        $script:previewTable.Rows.Count | Should Be 2
        $fake.Scrolls[-1] | Should Be 0
    }

    It "インデックスの集約ファイルが読めなければ、選んだ行だけを出す" {
        $row = newHitRow "見積.xlsx" "4月" 2 @("`t前", "`t見積", "`t後")
        Remove-Item -LiteralPath ([System.IO.Path]::Combine($row.Root, $row.RelPath))
        $fake.Current = $row

        showDetail

        $script:previewTable.Rows.Count | Should Be 1
        $script:previewTable.Rows[0].Number | Should Be "2"
    }
}

Describe "getPreviewCell" -Tag Unit {
    It "セルの上でなければ無し" {
        $null -eq (getPreviewCell $null) | Should Be $true
    }
}

Describe "copyPreviewSelection" -Tag Unit {
    BeforeEach { resetPreview }

    It "プレビューが無いときは、セルを選ぶよう案内する" {
        copyPreviewSelection
        $fake.Status | Should Match "^プレビューでコピーするセルをクリックしてください"
    }

    It "セルを選んでいないときも、案内だけを出す" {
        $script:previewTable = [PreviewTable]::new()
        $fake.Status = ""

        & $handlers["MenuPreviewCopy.Click"]

        $fake.Status | Should Match "^プレビューでコピーするセルをクリックしてください"
    }

    It "［行をコピー］もセルを選んでいなければ案内だけを出す" {
        $script:previewTable = [PreviewTable]::new()

        & $handlers["MenuPreviewCopyRow.Click"]

        $fake.Status | Should Match "^プレビューでコピーするセルをクリックしてください"
    }
}

Describe "イベント" -Tag Unit {
    BeforeEach { resetPreview }

    It "別の行を選ぶと、タイマーをかけ直し、一致したセルまでスクロールし直す" {
        $script:detailKeepScroll = $true

        & $handlers["ResultGrid.SelectionChanged"]

        $script:detailKeepScroll | Should Be $false
        $fake.TimerStops | Should Be 1
        $fake.TimerStarts | Should Be 1
    }

    It "タイマーが来たら止めてプレビューを出す" {
        $fake.Current = $null
        $ui.DetailTitle.Text = "前の行"

        & $handlers["Timer.Tick"]

        $fake.TimerStops | Should Be 1
        $ui.DetailTitle.Text | Should Be ""
    }

    It "行を選んでいるときに高さが変わったら、横の位置をそのままにして読み直す" {
        $ui.ResultGrid.SelectedItem = "行"

        & $handlers["PreviewScroll.SizeChanged"] $ui.PreviewScroll ([pscustomobject]@{ HeightChanged = $true })

        $script:detailKeepScroll | Should Be $true
        $fake.TimerStarts | Should Be 1
    }

    It "幅だけが変わったとき・行を選んでいないときは読み直さない" {
        $ui.ResultGrid.SelectedItem = "行"
        & $handlers["PreviewScroll.SizeChanged"] $ui.PreviewScroll ([pscustomobject]@{ HeightChanged = $false })
        $ui.ResultGrid.SelectedItem = $null
        & $handlers["PreviewScroll.SizeChanged"] $ui.PreviewScroll ([pscustomobject]@{ HeightChanged = $true })

        $script:detailKeepScroll | Should Be $false
        $fake.TimerStarts | Should Be 0
    }

    It "行を横にスクロールしたら、列見出しも同じ位置にする" {
        $ui.PreviewScroll.HorizontalOffset = 150.0

        & $handlers["PreviewScroll.ScrollChanged"]

        $fake.HeaderScrolls[-1] | Should Be 150.0
    }

    It "Ctrl+C 以外のキーではコピーしない" {
        $e = [pscustomobject]@{ Key = "A"; Handled = $false }

        & $handlers["PreviewScroll.PreviewKeyDown"] $ui.PreviewScroll $e

        $e.Handled | Should Be $false
        $fake.Status | Should Be ""
    }

    It "セルの外をクリック・ドラッグしても選ばない" {
        $script:previewTable = [PreviewTable]::new()
        $down = [pscustomobject]@{ OriginalSource = $null; LeftButton = "Pressed" }

        & $handlers["PreviewRows.PreviewMouseLeftButtonDown"] $ui.PreviewRows $down
        & $handlers["PreviewRows.MouseMove"] $ui.PreviewRows $down
        & $handlers["PreviewRows.PreviewMouseRightButtonDown"] $ui.PreviewRows $down

        $script:previewTable.HasSelection() | Should Be $false
    }

    It "ボタンを押していないマウスの移動は無視する" {
        $script:previewTable = [PreviewTable]::new()

        & $handlers["PreviewRows.MouseMove"] $ui.PreviewRows ([pscustomobject]@{ OriginalSource = $null; LeftButton = "Released" })

        $script:previewTable.HasSelection() | Should Be $false
    }

    It "列見出しの右端をドラッグすると、その列の幅が変わる" {
        $column = [PreviewColumn]::new()
        $column.Width = 100
        $e = New-Object System.Windows.Controls.Primitives.DragDeltaEventArgs 30.0, 0.0
        $e.RoutedEvent = [System.Windows.Controls.Primitives.Thumb]::DragDeltaEvent
        $e.Source = [pscustomobject]@{ DataContext = $column }

        $handlers["PreviewHeader.DragDelta"].Invoke($ui.PreviewHeader, $e)

        $column.Width | Should Be 130
    }

    It "列見出し以外のドラッグでは何もしない" {
        $e = New-Object System.Windows.Controls.Primitives.DragDeltaEventArgs 30.0, 0.0
        $e.RoutedEvent = [System.Windows.Controls.Primitives.Thumb]::DragDeltaEvent
        $e.Source = [pscustomobject]@{ DataContext = "列ではない" }

        { $handlers["PreviewHeader.DragDelta"].Invoke($ui.PreviewHeader, $e) } | Should Not Throw
    }
}
