# 検索結果から元のファイルを開く・パスをコピーする・結果を書き出す（tebunko\ui\open_source.ps1）のテスト。
# 画面層のため、$ui・$window は偽物にする。Excel・既定のアプリ・エクスプローラーの起動は Mock し、実際には開かない。
# コピーはクリップボードを書き換えるため、ここでは確かめない（画面の操作からの呼び出しだけを確かめる）。
BeforeDiscovery {
    # WPF の要素は STA のスレッドでしか作れない（powershell.exe 5.1 は既定で STA）。-Skip は発見の段階で決めるため、ここで調べる
    $sta = [System.Threading.Thread]::CurrentThread.GetApartmentState() -eq "STA"
}

BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    . "${scriptsDir}\shared\ui\types.ps1"
    . "${scriptsDir}\tebunko\ui\search\result_types.ps1"
    . "${scriptsDir}\shared\ui\data_grid.ps1"

    function newFakeControl {
        # イベントの登録（Add_<イベント>）を Handlers に記録するだけの偽のコントロール
        param (
            [string[]]$events,
            [hashtable]$properties = @{}
        )

        $control = New-Object PSObject -Property (@{ Handlers = @{} } + $properties)
        foreach ($name in $events) {
            $control | Add-Member ScriptMethod "Add_$name" ([scriptblock]::Create("param(`$handler) `$this.Handlers['$name'] = `$handler"))
        }
        return $control
    }

    function newFakeUi {
        $menu = { newFakeControl @("Click") }
        return @{
            ResultGrid       = newFakeControl @("MouseDoubleClick", "PreviewKeyDown", "PreviewMouseRightButtonDown", "ContextMenuOpening") @{ SelectedItem = $null; SelectedItems = @(); Columns = @() }
            ResultMenu       = newFakeControl @()
            MenuToggleGroup  = & $menu
            MenuOpen         = & $menu
            MenuOpenReadOnly = & $menu
            MenuOpenNew      = & $menu
            MenuOpenFolder   = & $menu
            OpenButton       = newFakeControl @("Click")
            OpenMenuButton   = newFakeControl @("Click")
            MenuOpenModeNormal   = newFakeControl @("Click")
            MenuOpenModeNew      = newFakeControl @("Click")
            MenuOpenModeReadOnly = newFakeControl @("Click")
            OpenFolderButton = newFakeControl @("Click")
            MenuCopy         = & $menu
            MenuCopyPath     = & $menu
            ExportButton     = newFakeControl @("Click") @{ ToolTip = "" }
            ExportMenuPart   = New-Object PSObject -Property @{ Visibility = "Visible" }
            ExportGroup      = New-Object PSObject -Property @{ ActualWidth = 100 }
            ExportMenuButton = newFakeControl @("Click") @{ ContextMenu = (New-Object PSObject -Property @{ MinWidth = 216; PlacementTarget = $null; Placement = $null; HorizontalOffset = 0; VerticalOffset = 0; IsOpen = $false }) }
            ExportButtonText = New-Object PSObject -Property @{ Text = "出力" }
            MenuExportXlsx   = newFakeControl @("Click")
            MenuExportCsv    = newFakeControl @("Click")
            MenuExportTsv    = newFakeControl @("Click")
        }
    }

    # 画面のほかのファイル（shell.ps1・result_list.ps1・folder_dialog.ps1）にある関数の代わり。テストごとに Mock する
    function setStatus { param([string]$text) $script:statuses.Add($text) }
    function showNoticeBanner { param([string]$prefix, [string]$kind, [string]$message) $script:banners.Add("${prefix}|${kind}|${message}") }
    function hideNoticeBanner { param([string]$prefix) $script:banners.Add("${prefix}|hidden") }
    function safe { param([scriptblock]$block) & $block }
    function getCurrentHitRow { }
    function getSelectedRows { }
    function getViewRows { }
    function getExportHits { param([bool]$truncated, [string]$filterText, [bool]$sorted) }
    function newTimer {
        # 画面のタイマーの代わり（開始・停止を受けるだけ。進めるのはテストが pumpExport を呼ぶ）
        param ([int]$milliseconds, [scriptblock]$block)
        $timer = New-Object PSObject
        $timer | Add-Member ScriptMethod Start { }
        $timer | Add-Member ScriptMethod Stop { }
        return $timer
    }
    function toggleFileGroup { param($group) }
    function activateResultItem { param($item) }
    function testResultSorted { return $false }
    function showConfirm { param([string]$title, [string]$heading, [string]$hint, [object[]]$facts, [object[]]$choices) }
    function showMessage { param([string]$message, [string]$buttons, [string]$icon, [string]$default) }
    function selectFolder { param([string]$description, [string]$initial, $owner, [bool]$knownExisting) }
    function factGone { param([string]$title, [string]$detail) "Gone ${title}" }
    function factNext { param([string]$title, [string]$detail) "Next ${title}" }
    function factWarn { param([string]$title, [string]$detail) "! ${title}" }
    function testNetworkPath { param([string]$path) $false }  # 既定はローカル。テストごとに Mock で切り替える
    function normalizeFolderPath { param([string]$path) $path.Trim() }  # 前後の空白を取るだけの代わり（本物は shared\core\folder.ps1）
    function startJob {
        # 画面の裏の仕事（BackgroundQueue）の代わりに、その場で実行して結果を渡す。
        # 返すまでの間に選択が変わった場合を試すため、$fake.BeforeDone があれば結果を渡す前に呼ぶ
        param ([scriptblock]$scriptBlock, [object[]]$arguments, [scriptblock]$onDone, [string]$queue = "default")
        $script:startJobCalls++
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

    $ui = newFakeUi
    $window = New-Object PSObject -Property @{ Cursor = $null }
    . "${scriptsDir}\tebunko\ui\search\search_bar_view.ps1"
    . "${scriptsDir}\tebunko\ui\search\result_list_view.ps1"
    . "${scriptsDir}\tebunko\ui\search\open_source_view.ps1"
    . "${scriptsDir}\tebunko\ui\open_source.ps1"

    function newRow {
        param (
            [string]$book = "見積.xlsx",
            [bool]$isExcel = $true,
            [bool]$isText = $false
        )

        [pscustomobject]@{
            Root = "C:\index"; RelDir = "営業\sub"; RelPath = "営業\sub\$book"; Book = $book
            Location = "見積[図形]"; MatchCell = "B2"; IsExcel = $isExcel; IsText = $isText
        }
    }

    function lastStatus { $script:statuses[$script:statuses.Count - 1] }

    $fake = @{ BeforeDone = $null }
}

Describe "findSourceFile" -Tag Io {
    BeforeEach {
        $script:statuses = New-Object System.Collections.Generic.List[string]
        $script:sourceFolderMaps = @{ 記録 = 1 }
        $script:startJobCalls = 0
        $script:foundPaths = New-Object System.Collections.Generic.List[string]
        $script:openSourcePendingRow.Value = $null
        $script:openSourcePendingPath.Value = ""
        $fake.BeforeDone = $null
        Mock getConfirmedSourceNames { @{ Confirmed = @("営業"); Crawled = @("営業") } }
    }

    It "ローカルのパスは、画面のスレッドでその場で確かめて開く（裏の仕事は使わない）" {
        newTsv "$TestDrive\known\sub\見積.xlsx" @("x")
        Mock getSourceLocation { @{ Name = "営業"; Folder = "$TestDrive\known"; Rest = "sub"; Known = $true } }
        Mock showConfirm { }

        findSourceFile (newRow) { param ($path) $script:foundPaths.Add($path) }
        $script:foundPaths -join "," | Should -Be "$TestDrive\known\sub\見積.xlsx"
        $script:startJobCalls | Should -Be 0
        Should -Invoke showConfirm -Times 0 -Exactly
    }

    It "記録した場所に無くても、同じ場所を指す別の書き方（別名）で見つかれば、その場所を記録して開く" {
        newTsv "$TestDrive\alias\sub\見積.xlsx" @("x")
        Mock getSourceLocation { @{ Name = "営業"; Folder = "Z:\営業"; Rest = "sub"; Known = $true } }
        Mock testNetworkPath { $false }
        Mock findSourceFileState { @{ State = "Found"; Path = "$TestDrive\alias\sub\見積.xlsx"; Alias = "$TestDrive\alias" } }
        Mock setIndexSourceFolder { }

        findSourceFile (newRow) { param ($path) $script:foundPaths.Add($path) }
        $script:foundPaths -join "," | Should -Be "$TestDrive\alias\sub\見積.xlsx"
        Should -Invoke setIndexSourceFolder -Times 1 -Exactly -ParameterFilter { $name -eq "営業" -and $folder -eq "$TestDrive\alias" }
        $script:sourceFolderMaps.Count | Should -Be 0
        lastStatus | Should -Be "インデックス [営業] の元のフォルダを $TestDrive\alias に変えました"
    }

    It "別の書き方（別名）で見つかっても、インデックス名が分からなければ記録しない" {
        newTsv "$TestDrive\alias2\sub\見積.xlsx" @("x")
        Mock getSourceLocation { @{ Name = ""; Folder = "Z:\営業"; Rest = "sub"; Known = $true } }
        Mock testNetworkPath { $false }
        Mock showConfirm { "use" }  # 名前が空だと確かめ済みと言えないので毎回確かめる。使うと答えても名前が無いので記録しない
        Mock findSourceFileState { @{ State = "Found"; Path = "$TestDrive\alias2\sub\見積.xlsx"; Alias = "$TestDrive\alias2" } }
        Mock setIndexSourceFolder { }

        findSourceFile (newRow) { param ($path) $script:foundPaths.Add($path) }
        $script:foundPaths -join "," | Should -Be "$TestDrive\alias2\sub\見積.xlsx"
        Should -Invoke setIndexSourceFolder -Times 0 -Exactly
    }

    It "ネットワークのパスでは、画面のスレッドから直接 getPathState・Test-Path を呼ばない（呼んだら失敗にする）" {
        Mock getSourceLocation { @{ Name = "営業"; Folder = "\\server\share\営業"; Rest = "sub"; Known = $true } }
        Mock testNetworkPath { param ($path) ([string]$path).StartsWith("\\") }
        Mock findSourceFileState { @{ State = "Found"; Path = "\\server\share\営業\sub\見積.xlsx" } }
        Mock getPathState { throw "画面のスレッドから getPathState を呼んでいる" }
        Mock Test-Path { throw "画面のスレッドから Test-Path を呼んでいる" }

        findSourceFile (newRow) { param ($path) $script:foundPaths.Add($path) }
        $script:foundPaths -join "," | Should -Be "\\server\share\営業\sub\見積.xlsx"
    }

    It "待っている間に同じ行をもう一度開いても、新しい依頼は出さない" {
        Mock getSourceLocation { @{ Name = "営業"; Folder = "\\server\share\営業"; Rest = "sub"; Known = $true } }
        Mock testNetworkPath { param ($path) ([string]$path).StartsWith("\\") }
        # 応答が来ない仕事を真似る（onDone を呼ばないため、依頼は待ったままになる）
        Mock startJob { $script:startJobCalls++ }

        $row = newRow
        findSourceFile $row { param ($path) $script:foundPaths.Add($path) }
        findSourceFile $row { param ($path) $script:foundPaths.Add($path) }

        $script:startJobCalls | Should -Be 1
        Should -Invoke getSourceLocation -Times 1 -Exactly
    }

    It "cancelPendingSourceLookup を呼ぶと、待っていた行をもう一度開いたときに新しい依頼を出す" {
        Mock getSourceLocation { @{ Name = "営業"; Folder = "\\server\share\営業"; Rest = "sub"; Known = $true } }
        Mock testNetworkPath { param ($path) ([string]$path).StartsWith("\\") }
        Mock startJob { $script:startJobCalls++ }

        $row = newRow
        findSourceFile $row { param ($path) $script:foundPaths.Add($path) }
        cancelPendingSourceLookup
        findSourceFile $row { param ($path) $script:foundPaths.Add($path) }

        $script:startJobCalls | Should -Be 2
    }

    It "cancelPendingSourceLookup は依頼の番号を進め、カーソルと待っている行を戻す" {
        $window.Cursor = [System.Windows.Input.Cursors]::AppStarting
        $script:openSourcePendingRow.Value = "dummy"
        $before = $script:openSourceRequest.Value

        cancelPendingSourceLookup

        $script:openSourceRequest.Value | Should -Be ($before + 1)
        $script:openSourcePendingRow.Value | Should -Be $null
        $window.Cursor | Should -Be $null
    }

    It "結果が届いた後、同じ行をもう一度開くと新しい依頼を出す" {
        # GetNewClosure() の中で待っている行の記録を戻すため、結果が届いた後に片づいていることを確かめる
        Mock getSourceLocation { @{ Name = "営業"; Folder = "\\server\share\営業"; Rest = "sub"; Known = $true } }
        Mock testNetworkPath { param ($path) ([string]$path).StartsWith("\\") }
        Mock findSourceFileState { @{ State = "Found"; Path = "\\server\share\営業\sub\見積.xlsx" } }

        $row = newRow
        findSourceFile $row { param ($path) $script:foundPaths.Add($path) }
        $script:startJobCalls | Should -Be 1

        findSourceFile $row { param ($path) $script:foundPaths.Add($path) }

        $script:startJobCalls | Should -Be 2
        $script:foundPaths.Count | Should -Be 2
    }

    It "届かない行を待っている間に、別のローカルの行を開いても正しく開ける" {
        Mock getSourceLocation {
            param ($hit, $maps)
            if ($hit.Book -eq "ローカル.xlsx") {
                return @{ Name = "営業"; Folder = "$TestDrive\local"; Rest = ""; Known = $true }
            }
            return @{ Name = "営業"; Folder = "\\server\share\営業"; Rest = "sub"; Known = $true }
        }
        Mock testNetworkPath { param ($path) ([string]$path).StartsWith("\\") }
        Mock findSourceFileState { @{ State = "Found"; Path = "$TestDrive\local\ローカル.xlsx" } }
        # 応答が来ない仕事を真似る（届かない行の依頼は待ったままになる）
        Mock startJob { $script:startJobCalls++ }

        findSourceFile (newRow) { param ($path) $script:foundPaths.Add("A:$path") }
        findSourceFile (newRow "ローカル.xlsx") { param ($path) $script:foundPaths.Add("B:$path") }

        $script:foundPaths -join "," | Should -Be "B:$TestDrive\local\ローカル.xlsx"
        $script:startJobCalls | Should -Be 1
    }

    It "ネットワークのパスは裏の仕事（'network' の列）で確かめ、届くまで確かめている間のステータスを出す" {
        Mock getSourceLocation { @{ Name = "営業"; Folder = "\\server\share\営業"; Rest = "sub"; Known = $true } }
        Mock testNetworkPath { param ($path) ([string]$path).StartsWith("\\") }
        Mock findSourceFileState { @{ State = "Found"; Path = "\\server\share\営業\sub\見積.xlsx" } }

        findSourceFile (newRow) { param ($path) $script:foundPaths.Add($path) }
        $script:startJobCalls | Should -Be 1
        $script:foundPaths -join "," | Should -Be "\\server\share\営業\sub\見積.xlsx"
        $script:statuses[0] | Should -Be "元のファイルを確かめています…：\\server\share\営業\sub\見積.xlsx（共有フォルダに接続できないときは、しばらくかかります）"
        $window.Cursor | Should -Be $null
    }

    It "待っている間に別の行を開く等をすると、前の依頼の結果は捨てる" {
        Mock getSourceLocation { @{ Name = "営業"; Folder = "\\server\share\営業"; Rest = "sub"; Known = $true } }
        Mock testNetworkPath { param ($path) ([string]$path).StartsWith("\\") }
        Mock findSourceFileState { @{ State = "Found"; Path = "\\server\share\営業\sub\見積.xlsx" } }
        # 結果が届く前に、別の依頼（openSourceRequest が進む）が出たことを真似る
        $fake.BeforeDone = { $script:openSourceRequest.Value++ }

        findSourceFile (newRow) { param ($path) $script:foundPaths.Add($path) }
        $script:foundPaths.Count | Should -Be 0
    }

    It "接続できないときは確認のダイアログを出し、ステータスに知らせる" {
        Mock getSourceLocation { @{ Name = "営業"; Folder = "\\server\share\営業"; Rest = "sub"; Known = $true } }
        Mock testNetworkPath { param ($path) ([string]$path).StartsWith("\\") }
        Mock findSourceFileState { @{ State = "Unreachable"; Message = "" } }
        Mock showConfirm { $null }

        findSourceFile (newRow) { param ($path) $script:foundPaths.Add($path) }
        $script:foundPaths.Count | Should -Be 0
        lastStatus | Should -Be "元のフォルダに接続できません：\\server\share\営業"
        Should -Invoke showConfirm -Times 1 -Exactly -ParameterFilter { $heading -eq "見積.xlsx を開けません" }
    }

    It "接続できないダイアログで［フォルダを選ぶ］を選ぶと、見つからないときと同じ流れに進む" {
        newTsv "$TestDrive\moved\sub\見積.xlsx" @("x")
        Mock getSourceLocation { @{ Name = "営業"; Folder = "\\server\share\営業"; Rest = "sub"; Known = $true } }
        Mock testNetworkPath { param ($path) ([string]$path).StartsWith("\\") }
        Mock findSourceFileState { @{ State = "Unreachable"; Message = "" } }
        Mock showConfirm { "pick" }
        Mock selectFolder { "$TestDrive\moved" }
        Mock setIndexSourceFolder { }

        findSourceFile (newRow) { param ($path) $script:foundPaths.Add($path) }
        $script:foundPaths -join "," | Should -Be "$TestDrive\moved\sub\見積.xlsx"
        Should -Invoke setIndexSourceFolder -Times 1 -Exactly -ParameterFilter { $name -eq "営業" -and $folder -eq "$TestDrive\moved" }
    }

    It "その他の失敗のときは、例外の文面を出す" {
        Mock getSourceLocation { @{ Name = "営業"; Folder = "\\server\share\営業"; Rest = "sub"; Known = $true } }
        Mock testNetworkPath { param ($path) ([string]$path).StartsWith("\\") }
        Mock findSourceFileState { @{ State = "Other"; Message = "アクセスが拒否されました。" } }
        Mock showConfirm { $null }

        findSourceFile (newRow) { param ($path) $script:foundPaths.Add($path) }
        lastStatus | Should -Be "元のファイルを確かめられませんでした：アクセスが拒否されました。"
        Should -Invoke showConfirm -Times 1 -Exactly -ParameterFilter { $facts -and $facts[0] -eq "Gone 元のファイルを確かめられませんでした" }
    }

    It "裏の仕事が予期せず失敗したときも、その他として知らせる" {
        Mock getSourceLocation { @{ Name = "営業"; Folder = "\\server\share\営業"; Rest = "sub"; Known = $true } }
        Mock testNetworkPath { param ($path) ([string]$path).StartsWith("\\") }
        Mock findSourceFileState { throw "バグ" }
        Mock showConfirm { $null }

        findSourceFile (newRow) { param ($path) $script:foundPaths.Add($path) }
        lastStatus | Should -Be "元のファイルを確かめられませんでした：バグ"
    }

    It "記録が無く、元のファイルがインデックスの直下（相対フォルダが空）なら、知らせるパスに \ を重ねない" {
        Mock getSourceLocation { @{ Name = "営業"; Folder = ""; Rest = ""; Known = $false } }
        Mock showConfirm { $null }
        $row = newRow
        $row.RelDir = ""

        findSourceFile $row { param ($path) $script:foundPaths.Add($path) }
        $script:foundPaths.Count | Should -Be 0
        lastStatus | Should -Be "元のファイルが見つかりません：C:\index\見積.xlsx"
    }

    It "見つからず、フォルダを選ばなければ開かない（記録した場所の近くを選ぶ画面の初期位置にする）" {
        [System.IO.Directory]::CreateDirectory("$TestDrive\gone") | Out-Null
        Mock getSourceLocation { @{ Name = "営業"; Folder = "$TestDrive\gone"; Rest = "sub"; Known = $true } }
        Mock testNetworkPath { $false }
        Mock findSourceFileState { @{ State = "Missing"; Initial = "$TestDrive\gone" } }
        Mock showConfirm { "pick" }
        Mock selectFolder { "" }

        findSourceFile (newRow) { param ($path) $script:foundPaths.Add($path) }
        $script:foundPaths.Count | Should -Be 0
        Should -Invoke selectFolder -Times 1 -Exactly -ParameterFilter { $initial -eq "$TestDrive\gone" }
        lastStatus | Should -Be "元のファイルが見つかりません：$TestDrive\gone\sub\見積.xlsx"
    }

    It "確認で選ぶのをやめれば開かない" {
        Mock getSourceLocation { @{ Name = "営業"; Folder = ""; Rest = "sub"; Known = $false } }
        Mock showConfirm { $null }
        Mock selectFolder { }

        findSourceFile (newRow) { param ($path) $script:foundPaths.Add($path) }
        $script:foundPaths.Count | Should -Be 0
        Should -Invoke selectFolder -Times 0 -Exactly
        lastStatus | Should -Be "元のファイルが見つかりません：C:\index\営業\sub\見積.xlsx"
    }

    It "記録が無くても、選んだフォルダの中で見つかれば、元のフォルダを記録して開く" {
        newTsv "$TestDrive\moved\営業\sub\見積.xlsx" @("x")
        Mock getSourceLocation { @{ Name = "営業"; Folder = ""; Rest = "sub"; Known = $false } }
        Mock showConfirm { "pick" }
        Mock selectFolder { "$TestDrive\moved\営業\sub" }  # ファイルのあるフォルダを選んだ
        Mock setIndexSourceFolder { }

        findSourceFile (newRow) { param ($path) $script:foundPaths.Add($path) }
        $script:foundPaths -join "," | Should -Be "$TestDrive\moved\営業\sub\見積.xlsx"
        Should -Invoke setIndexSourceFolder -Times 1 -Exactly -ParameterFilter { $name -eq "営業" -and $folder -eq "$TestDrive\moved\営業" }
        $script:sourceFolderMaps.Count | Should -Be 0
    }

    It "選んだフォルダで見つかっても、インデックス名が分からなければ記録せずに開く" {
        newTsv "$TestDrive\moved2\sub\見積.xlsx" @("x")
        Mock getSourceLocation { @{ Name = ""; Folder = ""; Rest = "sub"; Known = $false } }
        Mock showConfirm { "pick" }
        Mock selectFolder { "$TestDrive\moved2" }
        Mock setIndexSourceFolder { }

        findSourceFile (newRow) { param ($path) $script:foundPaths.Add($path) }
        $script:foundPaths -join "," | Should -Be "$TestDrive\moved2\sub\見積.xlsx"
        Should -Invoke setIndexSourceFolder -Times 0 -Exactly
        lastStatus | Should -Be "開きました：$TestDrive\moved2\sub\見積.xlsx"
    }

    It "選んだフォルダに無ければ、選んだフォルダを初期位置にしてもう一度聞く" {
        [System.IO.Directory]::CreateDirectory("$TestDrive\empty") | Out-Null
        $script:confirmCount = 0
        Mock getSourceLocation { @{ Name = "営業"; Folder = ""; Rest = ""; Known = $false } }
        Mock showConfirm { $script:confirmCount++; if ($script:confirmCount -eq 1) { "pick" } }
        Mock selectFolder { "$TestDrive\empty" }

        findSourceFile (newRow) { param ($path) $script:foundPaths.Add($path) }
        $script:foundPaths.Count | Should -Be 0
        Should -Invoke showConfirm -Times 2 -Exactly
        Should -Invoke showConfirm -Times 1 -Exactly -ParameterFilter { $facts -and $facts[0] -eq "Gone 選んだフォルダの中にありませんでした" }
    }
}

Describe "元の場所の対応の読み込み（ネットワークのワークスペース）" -Tag Unit {
    BeforeAll {
        $script:workspace = newTestWorkspace @{ IndexDir = "\\fileserver\共有\ws\work\index"; StatusFile = "\\fileserver\共有\ws\work\ingest_status.tsv" }
        function newNetworkRow {
            [pscustomobject]@{ Root = "\\fileserver\共有\ws\work\index"; RelDir = "営業\sub"; RelPath = "営業\sub\見積.xlsx"; Book = "見積.xlsx" }
        }
        function newSourceMap {
            $map = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
            $map["営業"] = "\\fileserver\共有\営業"
            return , $map
        }
    }

    BeforeEach {
        $script:statuses = New-Object System.Collections.Generic.List[string]
        $script:sourceFolderMaps = @{}
        $script:startJobCalls = 0
        $script:foundPaths = New-Object System.Collections.Generic.List[string]
        $script:openSourcePendingRow.Value = $null
        $script:openSourcePendingPath.Value = ""
        $fake.BeforeDone = $null
        Mock testNetworkPath { ([string]$path).StartsWith("\\") }
        Mock getSourceLocation { throw "画面のスレッドで対応を読んだ" }
        Mock getConfirmedSourceNames { @{ Confirmed = @("営業"); Crawled = @("営業") } }
    }

    It "開く: 対応が未読なら、画面のスレッドで読まず、裏の仕事で読んでから元のファイルを確かめる" {
        Mock getSourceFolderMap { newSourceMap }
        Mock findSourceFileState { @{ State = "Found"; Path = "\\fileserver\共有\営業\sub\見積.xlsx" } }

        findSourceFile (newNetworkRow) { param ($path) $script:foundPaths.Add($path) }

        $script:startJobCalls | Should -Be 2
        $script:statuses[0] | Should -Be (getSourceLookingStatus)
        $script:foundPaths -join "," | Should -Be "\\fileserver\共有\営業\sub\見積.xlsx"
        $script:sourceFolderMaps.ContainsKey("\\fileserver\共有\ws\work\index") | Should -Be $true
        $window.Cursor | Should -Be $null
        Should -Invoke getSourceLocation -Times 0
    }

    It "開く: 対応がキャッシュにあれば、対応を読む裏の仕事を出さない" {
        $script:sourceFolderMaps = @{ "\\fileserver\共有\ws\work\index" = (newSourceMap) }
        Mock findSourceFileState { @{ State = "Found"; Path = "\\fileserver\共有\営業\sub\見積.xlsx" } }

        findSourceFile (newNetworkRow) { param ($path) $script:foundPaths.Add($path) }

        $script:startJobCalls | Should -Be 1
        $script:statuses[0] | Should -BeLike "元のファイルを確かめています…*"
    }

    It "開く: 読めなかったときは、ステータスに知らせて開かない" {
        Mock getSourceFolderMap { throw "届きません" }

        findSourceFile (newNetworkRow) { param ($path) $script:foundPaths.Add($path) }

        $script:foundPaths.Count | Should -Be 0
        $script:statuses[$script:statuses.Count - 1] | Should -Be (getSourceLookupFailedStatus "届きません")
        $window.Cursor | Should -Be $null
    }

    It "開く: 待っている間に同じ行をもう一度開いても、新しい依頼は出さず、調べている最中であることを出し直す" {
        Mock startJob { $script:startJobCalls++ }

        $row = newNetworkRow
        findSourceFile $row { param ($path) $script:foundPaths.Add($path) }
        findSourceFile $row { param ($path) $script:foundPaths.Add($path) }

        $script:startJobCalls | Should -Be 1
        $script:statuses | Should -Be @((getSourceLookingStatus), (getSourceLookingStatus))
    }

    It "開く: 待っている間に別の依頼が出たら、読んだ結果は捨てる" {
        Mock getSourceFolderMap { newSourceMap }
        Mock findSourceFileState { @{ State = "Found"; Path = "x" } }
        $fake.BeforeDone = { $script:openSourceRequest.Value++ }

        findSourceFile (newNetworkRow) { param ($path) $script:foundPaths.Add($path) }

        $script:foundPaths.Count | Should -Be 0
        $script:sourceFolderMaps.Count | Should -Be 0
    }

    It "パスをコピー: 対応が未読なら、裏の仕事で読んでから写す" {
        Mock getSourceFolderMap { newSourceMap }
        Mock getCurrentHitRow { newNetworkRow }
        Mock setClipboardText { }

        copySourcePath

        $script:startJobCalls | Should -Be 1
        $script:statuses[0] | Should -Be (getSourceLookingStatus)
        Should -Invoke setClipboardText -Times 1 -Exactly -ParameterFilter { $text -eq "\\fileserver\共有\営業\sub\見積.xlsx" }
        Should -Invoke getSourceLocation -Times 0
    }

    It "パスをコピー: 対応がキャッシュにあれば、その場で写す（裏の仕事は出さない）" {
        $script:sourceFolderMaps = @{ "\\fileserver\共有\ws\work\index" = (newSourceMap) }
        Mock getCurrentHitRow { newNetworkRow }
        Mock setClipboardText { }

        copySourcePath

        $script:startJobCalls | Should -Be 0
        Should -Invoke setClipboardText -Times 1 -Exactly
    }

    It "パスをコピー: 待っている間に別のコピーの依頼が出たら、写さない" {
        Mock getSourceFolderMap { newSourceMap }
        Mock getCurrentHitRow { newNetworkRow }
        Mock setClipboardText { }
        $fake.BeforeDone = { $script:copySourceRequest.Value++ }

        copySourcePath

        Should -Invoke setClipboardText -Times 0 -Exactly
    }

    It "パスをコピー: [開く] の依頼が出ても、コピーの結果は捨てない（番号を分けている）" {
        Mock getSourceFolderMap { newSourceMap }
        Mock getCurrentHitRow { newNetworkRow }
        Mock setClipboardText { }
        $fake.BeforeDone = { $script:openSourceRequest.Value++ }

        copySourcePath

        Should -Invoke setClipboardText -Times 1 -Exactly
    }

    It "パスをコピー: 検索し直し・ワークスペースの変更（cancelPendingSourceLookup）では、コピーの結果も捨てる" {
        Mock getSourceFolderMap { newSourceMap }
        Mock getCurrentHitRow { newNetworkRow }
        Mock setClipboardText { }
        $fake.BeforeDone = { cancelPendingSourceLookup }

        copySourcePath

        Should -Invoke setClipboardText -Times 0 -Exactly
    }

    It "パスをコピー: 読めなかったときは、ステータスに知らせて写さない" {
        Mock getSourceFolderMap { throw "届きません" }
        Mock getCurrentHitRow { newNetworkRow }
        Mock setClipboardText { }

        copySourcePath

        Should -Invoke setClipboardText -Times 0 -Exactly
        $script:statuses[$script:statuses.Count - 1] | Should -Be (getSourceLookupFailedStatus "届きません")
    }
}

Describe "openWithShell" -Tag Io {
    It "通常の開き方は、既定のアプリでそのまま開く" {
        Mock Invoke-Item { }

        openWithShell "$TestDrive\見積.xlsx" ${openModeNormal} | Should -Be $true
        Should -Invoke Invoke-Item -Times 1 -Exactly -ParameterFilter { $LiteralPath -eq "$TestDrive\見積.xlsx" }
    }

    It "読み取り専用で開けなければ、そのまま開いて `$false を返す" {
        # 無いファイルは動詞で開けない（Win32Exception）。エラーのダイアログは出ない（ErrorDialog の既定は $false）
        Mock Invoke-Item { }

        openWithShell "$TestDrive\無い_open_source.xlsx" ${openModeReadOnly} | Should -Be $false
        Should -Invoke Invoke-Item -Times 1 -Exactly
    }

    It "取り込み対象の名前でないものは、<mode> でも開かず、ステータスに知らせる（<name>）" -ForEach @(
        @{ name = "hta"; leaf = "evil.hta"; mode = ${openModeNormal} }
        @{ name = "lnk"; leaf = "a.lnk"; mode = ${openModeNormal} }
        @{ name = "exe"; leaf = "run.exe"; mode = ${openModeReadOnly} }
        @{ name = "向きを変える文字"; leaf = "a$([char]0x202E)b.xlsx"; mode = ${openModeNormal} }
    ) {
        Mock Invoke-Item { }
        $script:statuses.Clear()

        openWithShell "$TestDrive\$leaf" $mode | Should -Be $false
        Should -Invoke Invoke-Item -Times 0 -Exactly
        $script:statuses.Count | Should -Be 1
        lastStatus | Should -Be (getSourceBookRejectedStatus $leaf)
    }

    It "知らない開き方は、通常の開き方と同じにする" {
        Mock Invoke-Item { }

        openWithShell "$TestDrive\[確定]見積.xlsx" "unknown" | Should -Be $true
        Should -Invoke Invoke-Item -Times 1 -Exactly -ParameterFilter { $LiteralPath -eq "$TestDrive\[確定]見積.xlsx" }
    }
}

Describe "openWithShell（-noFallback）" -Tag Io {
    It "読み取り専用で開けなくても、既定のアプリで開き直さず `$false を返す" {
        Mock Invoke-Item { }

        openWithShell "$TestDrive\無い_open_source.docm" ${openModeReadOnly} -noFallback | Should -Be $false
        Should -Invoke Invoke-Item -Times 0 -Exactly
    }
}

Describe "openWithNotepad" -Tag Io {
    It "固定のパス（%SystemRoot%\System32\notepad.exe）で、対象のパスを引数にして開く" {
        Mock Start-Process { }

        openWithNotepad "$TestDrive\起動.bat"
        Should -Invoke Start-Process -Times 1 -Exactly -ParameterFilter {
            $FilePath -eq "$env:SystemRoot\System32\notepad.exe" -and $ArgumentList -eq "`"$TestDrive\起動.bat`""
        }
    }

    It "[ ] や空白を含むパスも、引用符で囲んだ 1 つの引数として渡す" {
        Mock Start-Process { }

        openWithNotepad "$TestDrive\[確定] 起動 スクリプト.js"
        Should -Invoke Start-Process -Times 1 -Exactly -ParameterFilter {
            $ArgumentList -eq "`"$TestDrive\[確定] 起動 スクリプト.js`""
        }
    }

    It "開けたときは true を返す" {
        Mock Start-Process { }

        openWithNotepad "$TestDrive\起動.bat" | Should -Be $true
    }

    It "メモ帳が無い等で起動できないときは、例外を投げずに false を返す（既定のアプリには戻さない）" {
        Mock Start-Process { throw "指定されたファイルが見つかりません。" }

        openWithNotepad "$TestDrive\起動.bat" | Should -Be $false
    }
}

Describe "getOpenMode / setOpenMode" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "設定した開き方を返す"; set = "readOnly"; expected = "readOnly" }
        @{ name = "新規も返す"; set = "new"; expected = "new" }
        @{ name = "知らない値なら通常"; set = "unknown"; expected = "normal" }
        @{ name = "空なら通常"; set = ""; expected = "normal" }
    ) {
        param ($name, $set, $expected)
        setOpenMode $set
        getOpenMode | Should -Be $expected
    }

    It "選んでいなければ通常" {
        $script:openMode = $null
        getOpenMode | Should -Be ${openModeNormal}
    }
}

Describe "openSource" -Tag Unit {
    BeforeEach {
        $script:statuses = New-Object System.Collections.Generic.List[string]
        $script:openMode = $null
        Mock getConfirmedSourceNames { @{ Confirmed = @("営業"); Crawled = @("営業") } }
    }

    It "元のファイルが見つからなければ開かない" {
        Mock getCurrentHitRow { newRow }
        Mock findSourceFile { param ($row, $onFound) }  # 見つからないときは onFound を呼ばない
        Mock openInExcel { }
        Mock openWithShell { $true }

        openSource
        Should -Invoke openInExcel -Times 0 -Exactly
        Should -Invoke openWithShell -Times 0 -Exactly
    }

    It "Excel は該当シート（図形・コメントは元のシート）の該当セルを開く" {
        Mock getCurrentHitRow { newRow }
        Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\見積.xlsx" }
        Mock openInExcel { }

        openSource ${openModeReadOnly}
        Should -Invoke openInExcel -Times 1 -Exactly -ParameterFilter {
            $path -eq "C:\data\見積.xlsx" -and $location -eq "見積" -and $cell -eq "B2" -and $mode -eq ${openModeReadOnly}
        }
        lastStatus | Should -Be "読み取り専用で開きました：C:\data\見積.xlsx"
        $window.Cursor | Should -Be $null
    }

    It "Excel を操作できなければ、ファイルを開くだけにする" {
        Mock getCurrentHitRow { newRow }
        Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\見積.xlsx" }
        Mock openInExcel { throw "ダイアログを表示中" }
        Mock openWithShell { $true }

        openSource ${openModeNormal}
        lastStatus | Should -Be "開きました（該当セルへの移動はできませんでした）：C:\data\見積.xlsx"
    }

    It "Excel を操作できず、指定の開き方でも開けなければ、元のファイルを開いたと知らせる" {
        Mock getCurrentHitRow { newRow }
        Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\見積.xlsx" }
        Mock openInExcel { throw "ダイアログを表示中" }
        Mock openWithShell { $false }

        openSource ${openModeReadOnly}
        lastStatus | Should -Be "読み取り専用で開けなかったため、元のファイルを開きました（該当セルへの移動はできませんでした）：C:\data\見積.xlsx"
    }

    It "Excel 以外は既定のアプリで開く（開き方は既定の開き方）" {
        Mock getCurrentHitRow { newRow "報告.docx" $false }
        Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\報告.docx" }
        Mock openWithShell { $true }
        $script:openMode = ${openModeNew}

        openSource
        Should -Invoke openWithShell -Times 1 -Exactly -ParameterFilter { $mode -eq ${openModeNew} }
        lastStatus | Should -Be "新規で開きました：C:\data\報告.docx"
    }

    It "指定の開き方で開けなければ、元のファイルを開いたと知らせる" {
        Mock getCurrentHitRow { newRow "報告.docx" $false }
        Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\報告.docx" }
        Mock openWithShell { $false }

        openSource ${openModeNew}
        lastStatus | Should -Be "新規で開けなかったため、元のファイルを開きました：C:\data\報告.docx"
    }

    It "テキストは開き方の動詞を試さず、そのまま開いて「開きました」と出す（読み取り専用・新規を選んでいても）" {
        Mock getCurrentHitRow { newRow "議事メモ.txt" $false $true }
        Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\議事メモ.txt" }
        Mock openWithShell { $true }

        openSource ${openModeReadOnly}
        # 選んだ開き方（読み取り専用）ではなく、常に通常の開き方で openWithShell を呼ぶ（動詞を試さない）
        Should -Invoke openWithShell -Times 1 -Exactly -ParameterFilter { $path -eq "C:\data\議事メモ.txt" -and $mode -eq ${openModeNormal} }
        lastStatus | Should -Be "開きました：C:\data\議事メモ.txt"
    }

    It "PDF も開き方の動詞を試さず、そのまま開いて「開きました」と出す（読み取り専用・新規を選んでいても）" {
        Mock getCurrentHitRow { newRow "仕様書.pdf" $false $true }
        Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\仕様書.pdf" }
        Mock openWithShell { $true }

        openSource ${openModeReadOnly}
        Should -Invoke openWithShell -Times 1 -Exactly -ParameterFilter { $path -eq "C:\data\仕様書.pdf" -and $mode -eq ${openModeNormal} }
        lastStatus | Should -Be "開きました：C:\data\仕様書.pdf"
    }

    It "開くと実行・登録になる拡張子（.bat など）のテキストは、既定のアプリではなくメモ帳で開く" {
        Mock getCurrentHitRow { newRow "起動.bat" $false $true }
        Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\起動.bat" }
        Mock openWithNotepad { $true }
        Mock openWithShell { $true }

        openSource ${openModeReadOnly}
        Should -Invoke openWithNotepad -Times 1 -Exactly -ParameterFilter { $path -eq "C:\data\起動.bat" }
        Should -Invoke openWithShell -Times 0 -Exactly
        lastStatus | Should -Be "メモ帳で開きました：C:\data\起動.bat"
    }

    It "メモ帳で開けなかったときは、既定のアプリには戻さず失敗を伝える" {
        Mock getCurrentHitRow { newRow "起動.bat" $false $true }
        Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\起動.bat" }
        Mock openWithNotepad { $false }
        Mock openWithShell { $true }

        openSource ${openModeReadOnly}
        Should -Invoke openWithShell -Times 0 -Exactly
        lastStatus | Should -Be "メモ帳で開けませんでした：C:\data\起動.bat"
    }

    It "既定のアプリで開くテキスト（.html など）は、openWithNotepad を呼ばない" {
        Mock getCurrentHitRow { newRow "解析.html" $false $true }
        Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\解析.html" }
        Mock openWithNotepad { }
        Mock openWithShell { $true }

        openSource ${openModeReadOnly}
        Should -Invoke openWithShell -Times 1 -Exactly -ParameterFilter { $path -eq "C:\data\解析.html" }
        Should -Invoke openWithNotepad -Times 0 -Exactly
        lastStatus | Should -Be "開きました：C:\data\解析.html"
    }

    It "既定のアプリで開くのを断られたとき（<book>）は、「開きました」で上書きしない" -ForEach @(
        @{ book = "解析.html"; isText = $true }
        @{ book = "資料.pdf"; isText = $false }
    ) {
        Mock getCurrentHitRow { newRow $book $false $isText }
        Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\$($row.Book)" }
        Mock openWithNotepad { }
        Mock openWithShell { setStatus "断りました"; $false }

        openSource ${openModeReadOnly}
        Should -Invoke openWithShell -Times 1 -Exactly
        lastStatus | Should -Be "断りました"
    }

    It "取り込み対象でないファイル名（<book>）は、元のファイルを探さず、何も開かない" -ForEach @(
        @{ book = "evil.hta" }
        @{ book = "..\y.xlsx" }
        @{ book = "a.lnk" }
    ) {
        Mock getCurrentHitRow { newRow $book $false $true }
        Mock findSourceFile { }
        Mock openWithNotepad { }
        Mock openWithShell { $true }

        openSource ${openModeReadOnly}
        Should -Invoke findSourceFile -Times 0 -Exactly
        Should -Invoke openWithShell -Times 0 -Exactly
        Should -Invoke openWithNotepad -Times 0 -Exactly
        lastStatus | Should -Be "取り込み対象のファイル名ではないため、開きません：$book"
    }

    Context "新規で開く（「新規」の動詞が無い種類はコピーを開く）" {
        BeforeEach {
            Mock getOpenCopyDir { "C:\work\tmp\key\open_copy" }
            Mock openWithNotepad { $true }
            Mock openWithShell { $true }
            Mock openInExcel { }
            Mock getWorkspaceJobQueue { "default" }
            # 裏の仕事をその場で動かし、終わりの処理も続けて呼ぶ（コピーは裏の仕事で作る）
            Mock startJob {
                $out = @(& $scriptBlock @arguments)
                & $onDone $out ""
            }
        }

        It "PDF は、コピーを作って既定のアプリで開く（元のファイルは開かない）" {
            Mock getCurrentHitRow { newRow "仕様書.pdf" $false $true }
            Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\仕様書.pdf" }
            Mock copySourceForOpen { @{ Path = "C:\work\tmp\key\open_copy\abc\仕様書.pdf"; Reason = "" } }

            openSource ${openModeNew}
            Should -Invoke copySourceForOpen -Times 1 -Exactly -ParameterFilter { $sourcePath -eq "C:\data\仕様書.pdf" -and $book -eq "仕様書.pdf" -and $copyDir -eq "C:\work\tmp\key\open_copy" }
            Should -Invoke openWithShell -Times 1 -Exactly -ParameterFilter { $path -eq "C:\work\tmp\key\open_copy\abc\仕様書.pdf" -and $mode -eq ${openModeNormal} }
            Should -Invoke getWorkspaceJobQueue -Times 1 -Exactly -ParameterFilter { $paths.Count -eq 2 -and $paths[0] -eq "C:\data" -and $paths[1] -eq "C:\work\tmp\key\open_copy" }
            lastStatus | Should -Be (getOpenCopyOpenedStatus "C:\data\仕様書.pdf" "C:\work\tmp\key\open_copy\abc\仕様書.pdf")
            lastStatus | Should -BeLike "*コピー：C:\work\tmp\key\open_copy\abc\仕様書.pdf*"
        }

        It "メモ帳で開くテキスト（.bat など）は、コピーをメモ帳で開く（既定のアプリでは開かない）" {
            Mock getCurrentHitRow { newRow "起動.bat" $false $true }
            Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\起動.bat" }
            Mock copySourceForOpen { @{ Path = "C:\work\tmp\key\open_copy\abc\起動.bat"; Reason = "" } }

            openSource ${openModeNew}
            Should -Invoke openWithNotepad -Times 1 -Exactly -ParameterFilter { $path -eq "C:\work\tmp\key\open_copy\abc\起動.bat" }
            Should -Invoke openWithShell -Times 0 -Exactly
            lastStatus | Should -Be (getOpenCopyOpenedStatus "C:\data\起動.bat" "C:\work\tmp\key\open_copy\abc\起動.bat" "メモ帳")
        }

        It "コピーを作れなければ、何も開かず、理由（<reason>）に合わせて［開く］を使うよう伝える" -TestCases @(
            @{ reason = "Place" }
            @{ reason = "Failed" }
        ) {
            param ($reason)
            Mock getCurrentHitRow { newRow "仕様書.pdf" $false $true }
            Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\仕様書.pdf" }
            Mock copySourceForOpen { @{ Path = ""; Reason = $reason } }

            openSource ${openModeNew}
            Should -Invoke openWithShell -Times 0 -Exactly
            Should -Invoke openWithNotepad -Times 0 -Exactly
            lastStatus | Should -Be (getOpenCopyFailedStatus "C:\data\仕様書.pdf" $reason)
        }

        It "仕事そのものが失敗したときは、作るのに失敗した扱いにする" {
            Mock getCurrentHitRow { newRow "仕様書.pdf" $false $true }
            Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\仕様書.pdf" }
            Mock copySourceForOpen { @{ Path = "x"; Reason = "" } }
            Mock startJob { & $onDone @() "失敗" }

            openSource ${openModeNew}
            Should -Invoke openWithShell -Times 0 -Exactly
            lastStatus | Should -Be (getOpenCopyFailedStatus "C:\data\仕様書.pdf" "Failed")
        }

        It "名前に [ ] を含むファイルも、そのままのパスでコピーを作って開く（断らない）" {
            Mock getCurrentHitRow { newRow "報告[1].pdf" $false $true }
            Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\報告[1].pdf" }
            Mock copySourceForOpen { @{ Path = "C:\work\tmp\key\open_copy\abc\報告[1]（コピー 20261012-1530）.pdf"; Reason = "" } }

            openSource ${openModeNew}
            Should -Invoke copySourceForOpen -Times 1 -Exactly -ParameterFilter { $sourcePath -eq "C:\data\報告[1].pdf" -and $book -eq "報告[1].pdf" }
            Should -Invoke openWithShell -Times 1 -Exactly -ParameterFilter { $path -eq "C:\work\tmp\key\open_copy\abc\報告[1]（コピー 20261012-1530）.pdf" }
        }

        It "名前が長すぎるファイルは、コピーの仕事を始めず、理由をステータスで伝える" {
            $long = ("あ" * 250) + ".pdf"
            Mock getCurrentHitRow { newRow $long $false $true }
            Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\$($row.Book)" }
            Mock copySourceForOpen { @{ Path = "x"; Reason = "" } }

            openSource ${openModeNew}
            Should -Invoke copySourceForOpen -Times 0 -Exactly
            Should -Invoke startJob -Times 0 -Exactly
            Should -Invoke openWithShell -Times 0 -Exactly
            lastStatus | Should -Be (getOpenCopyFailedStatus "C:\data\$long" "TooLong")
        }

        It "Word・PowerPoint・Excel は、コピーを作らない（動詞・Workbooks.Add のまま）" -TestCases @(
            @{ book = "報告.docx"; isExcel = $false }
            @{ book = "発表.pptx"; isExcel = $false }
            @{ book = "見積.xlsx"; isExcel = $true }
        ) {
            param ($book, $isExcel)
            Mock getCurrentHitRow { newRow $book $isExcel } -ParameterFilter { $true }
            Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\$($row.Book)" }
            Mock copySourceForOpen { @{ Path = ""; Reason = "" } }

            openSource ${openModeNew}
            Should -Invoke copySourceForOpen -Times 0 -Exactly
        }

        It "新規以外の開き方では、コピーを作らない" {
            Mock getCurrentHitRow { newRow "仕様書.pdf" $false $true }
            Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\仕様書.pdf" }
            Mock copySourceForOpen { @{ Path = ""; Reason = "" } }

            openSource ${openModeNormal}
            Should -Invoke copySourceForOpen -Times 0 -Exactly
        }
    }
}

Describe "findSourceFile（もらったインデックスの元のフォルダの確認）" -Tag Io {
    BeforeEach {
        $script:statuses = New-Object System.Collections.Generic.List[string]
        $script:sourceFolderMaps = @{ 記録 = 1 }
        $script:startJobCalls = 0
        $script:foundPaths = New-Object System.Collections.Generic.List[string]
        $script:order = New-Object System.Collections.Generic.List[string]
        $script:hints = New-Object System.Collections.Generic.List[string]
        $script:confirmedNow = @("営業")
        $script:openSourcePendingRow.Value = $null
        $script:openSourcePendingPath.Value = ""
        $fake.BeforeDone = $null
        # 設定に無い名前「受取」（content_index\ に手でコピーしたインデックス）。元のフォルダは共有フォルダ
        Mock getSourceLocation { @{ Name = "受取"; Folder = "\\host\share\営業"; Rest = "sub"; Known = $true } }
        Mock getConfirmedSourceNames { @{ Confirmed = $script:confirmedNow; Crawled = $script:confirmedNow } }
        Mock testNetworkPath { ([string]$path).StartsWith("\\") }
        Mock findSourceFileState {
            $script:order.Add("state")
            @{ State = "Found"; Path = "\\host\share\営業\sub\見積.xlsx" }
        }
        Mock setIndexSourceFolder { $script:order.Add("set"); $script:confirmedNow = @($script:confirmedNow) + $name }
        Mock showConfirm {
            $script:order.Add("confirm")
            $script:hints.Add($hint)
            "use"
        }
    }

    It "設定に無い名前の元のフォルダは、フォルダに触れる前に確認する（有無の確認・接続の前）" {
        findSourceFile (newRow) { param ($path) $script:foundPaths.Add($path) }

        Should -Invoke showConfirm -Times 1 -Exactly -ParameterFilter { $title -eq "元のフォルダを確かめてください" }
        $script:order -join "," | Should -Be "confirm,set,state"
        $script:foundPaths -join "," | Should -Be "\\host\share\営業\sub\見積.xlsx"
        Should -Invoke setIndexSourceFolder -Times 1 -Exactly -ParameterFilter { $name -eq "受取" -and $folder -eq "\\host\share\営業" }
    }

    It "共有フォルダのときは、サインイン情報が送られることがあると添える" {
        findSourceFile (newRow) { param ($path) }

        $script:hints[0] | Should -Match "サインイン情報"
    }

    It "ローカルのフォルダでも確認する（サインイン情報の注意は添えない）" {
        Mock testNetworkPath { $false }
        Mock getSourceLocation { @{ Name = "受取"; Folder = "C:\営業"; Rest = "sub"; Known = $true } }
        Mock findSourceFileState { @{ State = "Found"; Path = "C:\営業\sub\見積.xlsx" } }

        findSourceFile (newRow) { param ($path) $script:foundPaths.Add($path) }

        $script:hints.Count | Should -Be 1
        $script:hints[0] | Should -Not -Match "サインイン情報"
        $script:foundPaths -join "," | Should -Be "C:\営業\sub\見積.xlsx"
        $script:startJobCalls | Should -Be 0
    }

    It "キャンセルしたら、何も確かめず、何も記録せず、開くのをやめたと出す" {
        Mock showConfirm { $script:order.Add("confirm"); $null }

        findSourceFile (newRow) { param ($path) $script:foundPaths.Add($path) }

        Should -Invoke findSourceFileState -Times 0 -Exactly
        Should -Invoke setIndexSourceFolder -Times 0 -Exactly
        $script:startJobCalls | Should -Be 0
        $script:foundPaths.Count | Should -Be 0
        lastStatus | Should -Be (getSourceConfirmCanceledStatus)
    }

    It "［フォルダを選ぶ］は、「見つかりません」の確認を挟まずにフォルダ選択から始める（選ばなければ何も記録しない）" {
        Mock showConfirm { $script:order.Add("confirm:$title"); "pick" }
        Mock selectFolder { $script:order.Add("select"); $null }

        findSourceFile (newRow) { param ($path) $script:foundPaths.Add($path) }

        $script:order -join "," | Should -Be "confirm:元のフォルダを確かめてください,select"
        Should -Invoke findSourceFileState -Times 0 -Exactly
        Should -Invoke setIndexSourceFolder -Times 0 -Exactly
        $script:foundPaths.Count | Should -Be 0
        # フォルダ選択をやめたのは「開くのをやめた」。「元のファイルが見つかりません」とは出さない
        lastStatus | Should -Be (getSourceConfirmCanceledStatus)
    }

    It "設定にある名前（自分で作ったインデックス・確認済み）は、確認しない" {
        $script:confirmedNow = @("営業", "受取")

        findSourceFile (newRow) { param ($path) $script:foundPaths.Add($path) }

        Should -Invoke showConfirm -Times 0 -Exactly
        $script:foundPaths.Count | Should -Be 1
    }

    It "［このフォルダを使う］で記録したら、同じインデックスは次から聞かない" {
        findSourceFile (newRow) { param ($path) $script:foundPaths.Add($path) }
        findSourceFile (newRow) { param ($path) $script:foundPaths.Add($path) }

        Should -Invoke showConfirm -Times 1 -Exactly
        $script:foundPaths.Count | Should -Be 2
    }

    It "前後に空白のある共有フォルダの書き方でも、サインイン情報の注意を添える（整えてから共有フォルダか調べる）" {
        Mock getSourceLocation { @{ Name = "受取"; Folder = " \\host\share\営業 "; Rest = "sub"; Known = $true } }

        findSourceFile (newRow) { param ($path) }

        $script:hints[0] | Should -Match "サインイン情報"
    }

    It "確認で <answer> を選んだら、変えていたカーソルを戻す" -TestCases @(
        @{ answer = $null }
        @{ answer = "pick" }
    ) {
        param ($answer)
        $script:answer = $answer
        Mock showConfirm { $script:answer }
        Mock selectFolder { $null }
        $window.Cursor = [System.Windows.Input.Cursors]::AppStarting

        findSourceFile (newRow) { param ($path) }

        $window.Cursor | Should -Be $null
    }

    It "確かめている間に別の行の確認をキャンセルしたら、前の行の結果は捨てて開かない" {
        $script:confirmedNow = @("営業", "受取")
        Mock showConfirm { $null }
        $fake.BeforeDone = {
            # 前の行を裏で確かめている間に、確認が要る別の行を開いてキャンセルする
            $fake.BeforeDone = $null
            $script:confirmedNow = @("営業")
            findSourceFile (newRow "別.xlsx") { param ($path) $script:foundPaths.Add("別:$path") }
        }

        findSourceFile (newRow) { param ($path) $script:foundPaths.Add($path) }

        $script:foundPaths.Count | Should -Be 0
        lastStatus | Should -Be (getSourceConfirmCanceledStatus)
        $script:openSourcePendingRow.Value | Should -Be $null
    }
}

Describe "openSource（もらったインデックスのマクロを持てる形式）" -Tag Unit {
    BeforeEach {
        $script:statuses = New-Object System.Collections.Generic.List[string]
        $script:openMode = $null
        $script:indexName = "受取"
        Mock getConfirmedSourceNames { @{ Confirmed = @("営業"); Crawled = @("営業") } }
    }

    It "もらったインデックスの .xlsm は、通常の開き方でも読み取り専用で開き、知らせを添える" {
        Mock getCurrentHitRow { newRow "マクロ.xlsm" }
        Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\マクロ.xlsm" $script:indexName }
        Mock openInExcel { }

        openSource ${openModeNormal}

        Should -Invoke openInExcel -Times 1 -Exactly -ParameterFilter { $mode -eq ${openModeReadOnly} }
        lastStatus | Should -Be "読み取り専用で開きました：C:\data\マクロ.xlsm（もらったインデックスのため、マクロを持てる形式は読み取り専用で開きます）"
    }

    It "自分で作ったインデックスの .xlsm は、選んだ開き方のまま" {
        $script:indexName = "営業"
        Mock getCurrentHitRow { newRow "マクロ.xlsm" }
        Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\マクロ.xlsm" $script:indexName }
        Mock openInExcel { }

        openSource ${openModeNormal}

        Should -Invoke openInExcel -Times 1 -Exactly -ParameterFilter { $mode -eq ${openModeNormal} }
        lastStatus | Should -Be "開きました：C:\data\マクロ.xlsm"
    }

    It "もらったインデックスでも、マクロを持てない形式（.xlsx）は選んだ開き方のまま" {
        Mock getCurrentHitRow { newRow "見積.xlsx" }
        Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\見積.xlsx" $script:indexName }
        Mock openInExcel { }

        openSource ${openModeNew}

        Should -Invoke openInExcel -Times 1 -Exactly -ParameterFilter { $mode -eq ${openModeNew} }
        lastStatus | Should -Be "新規で開きました：C:\data\見積.xlsx"
    }

    It "もらったインデックスの .docm は、読み取り専用で開き、開けなければ通常には落とさず開かない" {
        Mock getCurrentHitRow { newRow "報告.docm" $false }
        Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\報告.docm" $script:indexName }
        Mock openWithShell { $false }

        openSource ${openModeNormal}

        Should -Invoke openWithShell -Times 1 -Exactly -ParameterFilter { $mode -eq ${openModeReadOnly} -and [bool]$noFallback }
        lastStatus | Should -Be (getSourceReadOnlyFailedStatus "C:\data\報告.docm")
    }

    It "もらったインデックスの .docm を読み取り専用で開けたら、知らせを添える" {
        Mock getCurrentHitRow { newRow "報告.docm" $false }
        Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\報告.docm" $script:indexName }
        Mock openWithShell { $true }

        openSource ${openModeNew}

        lastStatus | Should -Be "読み取り専用で開きました：C:\data\報告.docm（もらったインデックスのため、マクロを持てる形式は読み取り専用で開きます）"
    }

    It "Excel を操作できなくても、もらったインデックスの .xlsm は通常の開き方には落とさない" {
        Mock getCurrentHitRow { newRow "マクロ.xlsm" }
        Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\マクロ.xlsm" $script:indexName }
        Mock openInExcel { throw "ダイアログを表示中" }
        Mock openWithShell { $false }

        openSource ${openModeNormal}

        Should -Invoke openWithShell -Times 1 -Exactly -ParameterFilter { $mode -eq ${openModeReadOnly} -and [bool]$noFallback }
        lastStatus | Should -Be (getSourceReadOnlyFailedStatus "C:\data\マクロ.xlsm")
    }

    It "自分で作ったインデックスの .docm は、これまでどおり開けなければ元のファイルを開く" {
        $script:indexName = "営業"
        Mock getCurrentHitRow { newRow "報告.docm" $false }
        Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\報告.docm" $script:indexName }
        Mock openWithShell { $false }

        openSource ${openModeReadOnly}

        Should -Invoke openWithShell -Times 1 -Exactly -ParameterFilter { -not [bool]$noFallback }
        lastStatus | Should -Be "読み取り専用で開けなかったため、元のファイルを開きました：C:\data\報告.docm"
    }

    It "インデックスの名前に見えない文字が混じっていれば、自分で作った名前とは別物（もらったインデックス）として読み取り専用で開く" {
        $script:indexName = "営" + [string][char]0xAD + "業"
        Mock getCurrentHitRow { newRow "マクロ.xlsm" }
        Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\マクロ.xlsm" $script:indexName }
        Mock openInExcel { }

        openSource ${openModeNormal}

        Should -Invoke openInExcel -Times 1 -Exactly -ParameterFilter { $mode -eq ${openModeReadOnly} }
    }

    It "インデックスの名前が渡されないときも、もらったインデックスとして読み取り専用で開く" {
        Mock getCurrentHitRow { newRow "マクロ.xlsm" }
        Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\マクロ.xlsm" }
        Mock openInExcel { }

        openSource ${openModeNormal}

        Should -Invoke openInExcel -Times 1 -Exactly -ParameterFilter { $mode -eq ${openModeReadOnly} }
    }
}

Describe "setExcelAutomationSecurity" -Tag Unit {
    It "設定して、前の値を返す" {
        $excel = [pscustomobject]@{ AutomationSecurity = 1 }

        setExcelAutomationSecurity $excel 2 | Should -Be 1
        $excel.AutomationSecurity | Should -Be 2
    }

    It "設定できなければ例外を投げる（マクロが動く設定のまま開かない）" {
        $excel = New-Object PSObject
        $excel | Add-Member ScriptProperty AutomationSecurity { 1 } { throw "設定できません" }

        { setExcelAutomationSecurity $excel 2 } | Should -Throw "*設定できません*"
    }
}

Describe "openSourceFolder" -Tag Unit {
    It "行を選んでいない・元のファイルが見つからなければ開かない" {
        Mock Start-Process { }
        Mock getCurrentHitRow { $null }
        openSourceFolder
        Mock getCurrentHitRow { newRow }
        Mock findSourceFile { param ($row, $onFound) }
        openSourceFolder
        Should -Invoke Start-Process -Times 0 -Exactly
    }

    It "エクスプローラーで元のファイルを選んだ状態で開く" {
        Mock Start-Process { }
        Mock getCurrentHitRow { newRow }
        Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\見積.xlsx" }

        openSourceFolder
        Should -Invoke Start-Process -Times 1 -Exactly -ParameterFilter {
            $FilePath -eq "explorer.exe" -and $ArgumentList -eq "/select,`"C:\data\見積.xlsx`""
        }
    }
}

Describe "showInExplorer" -Tag Unit {
    BeforeAll {
        function writeErrorLog { param ([string]$context, $record) }
    }

    BeforeEach {
        Mock writeErrorLog {}
    }

    It "explorer.exe に /select, とパスを渡し、終了コードは見ない（起動できれば true）" {
        Mock Start-Process { $global:LASTEXITCODE = 1 }

        showInExplorer "C:\w\a b.xlsx" | Should -Be $true
        Should -Invoke Start-Process -Times 1 -Exactly -ParameterFilter { $FilePath -eq "explorer.exe" -and $ArgumentList -eq '/select,"C:\w\a b.xlsx"' }
        Should -Invoke writeErrorLog -Times 0 -Exactly
    }

    It "起動で例外になれば false を返し、ログに残す" {
        Mock Start-Process { throw "起動できません" }

        showInExplorer "C:\w\a.xlsx" | Should -Be $false
        Should -Invoke writeErrorLog -Times 1 -Exactly
    }
}

Describe "openExportedFile" -Tag Unit {
    BeforeAll {
        function writeErrorLog { param ([string]$context, $record) }
    }

    It "Excel ブックは、そのパスのまま（[ ] をワイルドカードにせず）既定のアプリで開く" {
        Mock Invoke-Item { }

        openExportedFile "C:\w\検索結果[1].xlsx" | Should -Be $true
        Should -Invoke Invoke-Item -Times 1 -Exactly -ParameterFilter { $LiteralPath -eq "C:\w\検索結果[1].xlsx" }
    }

    It "<name>: CSV・TSV はメモ帳で開く（表計算アプリで数式として動かさない）。既定のアプリでは開かない" -ForEach @(
        @{ name = "csv"; path = "C:\w\検索結果.csv" }
        @{ name = "tsv"; path = "C:\w\検索結果.TSV" }
    ) {
        Mock Invoke-Item { }
        Mock openWithNotepad { $true }
        $script:wanted = $path

        openExportedFile $path | Should -Be $true
        Should -Invoke openWithNotepad -Times 1 -Exactly -ParameterFilter { $path -eq $script:wanted }
        Should -Invoke Invoke-Item -Times 0 -Exactly
    }

    It "メモ帳が使えないとき（CSV・TSV）は false を返す（既定のアプリには戻さない）" {
        Mock Invoke-Item { }
        Mock openWithNotepad { $false }

        openExportedFile "C:\w\a.csv" | Should -Be $false
        Should -Invoke Invoke-Item -Times 0 -Exactly
    }

    It "開けなければ、false を返し、例外にしない（既定のアプリが無いときなど）" {
        Mock Invoke-Item { throw "アプリが関連付けられていません" }
        Mock writeErrorLog { }

        openExportedFile "C:\w\a.xlsx" | Should -Be $false
    }
}

Describe "exportResults" -Tag Io {
    BeforeAll {
        $hits = @(
            [pscustomobject]@{ Book = "見積.xlsx"; Location = "Sheet1"; LineNumber = 3; Line = "りんご`t100"; RelDir = "sub" },
            [pscustomobject]@{ Book = "報告.docx"; Location = "ページ001"; LineNumber = 1; Line = "りんごの報告"; RelDir = "" }
        )

        function newFakeService {
            # 検索の司令の代わり。Export で受けた要求を、その場で invokeSearchExport で実行する（$run が偽なら受けるだけ）
            param ([bool]$run = $true)
            $service = New-Object PSObject -Property @{ Requests = (New-Object System.Collections.Generic.List[object]); Run = $run; Running = $true; Failure = "" }
            $service | Add-Member ScriptMethod Export { param ($request) $this.Requests.Add($request); if ($this.Run) { invokeSearchExport $request }; return $request }
            $service | Add-Member ScriptMethod IsRunning { return $this.Running }
            $service | Add-Member ScriptMethod GetFailure { return $this.Failure }
            return $service
        }

        function newExportSearch {
            param ([string]$workDir, [string]$word = "りんご", [bool]$truncated = $false)
            return @{ Word = $word; Shared = @{ WorkDir = $workDir; Truncated = $truncated } }
        }

        function readZipEntryNames {
            param ([string]$path)
            $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
            try { return @($zip.Entries | ForEach-Object { $_.FullName }) } finally { $zip.Dispose() }
        }
    }

    BeforeEach {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $script:statuses = New-Object System.Collections.Generic.List[string]
        $script:banners = New-Object System.Collections.Generic.List[string]
        $script:search = $null
        $script:filterText = ""
        $script:exportJob = $null
        $script:exportPending = $null
        $script:lastExportFormat = "xlsx"
        $ui.ExportButtonText.Text = "出力"
        $ui.ExportMenuPart.Visibility = "Visible"
        $ui.ExportMenuButton.ContextMenu.IsOpen = $false
        $script:searchService = newFakeService
        $script:opened = New-Object System.Collections.Generic.List[string]
        Mock getTargetFolders { @() }
        # 出力先の置き場の確かめ（リンクでないこと）は search_export.Tests で確かめる。ここでは、テストが決めた出力先に書かせる
        Mock testOpenCopyDir { $true }
        Mock openExportedFile { $script:opened.Add($path); $true }
    }

    It "検索する前は出力先を決めず、出力もしない" {
        $script:lastSearch = $null
        Mock newExportTarget { @{ Path = "$TestDrive
one.xlsx"; Format = $format } }

        exportResults "xlsx"
        exportResultsAs "csv"
        lastStatus | Should -Be "先に検索してください。"
        Should -Invoke newExportTarget -Times 0 -Exactly
        $script:lastExportFormat | Should -Be "xlsx"
        $script:searchService.Requests.Count | Should -Be 0
    }

    It "出力すると、出力先をダイアログなしで決めて書き、書けたファイルを開く" {
        $workDir = "$TestDrive\export_xlsx\work"
        $out = "$TestDrive\export_xlsx\検索結果.xlsx"
        [System.IO.Directory]::CreateDirectory("$TestDrive\export_xlsx") | Out-Null
        $workspace = newTestWorkspace @{} $workDir
        $script:lastSearch = newExportSearch $workDir
        Mock getExportHits { , $hits }
        Mock newExportTarget { @{ Path = $out; Format = $format } }
        $script:searchService = newFakeService $false

        exportResults "xlsx"
        $ui.ExportButtonText.Text | Should -Be "出力を中止"
        $ui.ExportMenuPart.Visibility | Should -Be "Collapsed"
        $ui.ExportButton.ToolTip | Should -Be "出力を中止します"
        lastStatus | Should -Be "検索結果を出力中…（0 件）"
        $script:searchService.Requests.Count | Should -Be 1
        $script:searchService.Requests[0].Format | Should -Be "xlsx"
        $script:searchService.Requests[0].TmpRoot | Should -Be $workspace.TmpRoot
        invokeSearchExport $script:searchService.Requests[0]
        pumpExport

        $out | Should -Exist
        readZipEntryNames $out | Should -Contain "xl/worksheets/sheet1.xml"
        lastStatus | Should -Be "検索結果.xlsx に 2 件を出力しました（出力したファイルを開きます）"
        ($script:opened -join "|") | Should -Be $out
        $ui.ExportButtonText.Text | Should -Be "出力"
        $ui.ExportMenuPart.Visibility | Should -Be "Visible"
        $script:exportJob | Should -BeNullOrEmpty
    }

    It "［▾］を押すと、形式のメニューがボタンの右端にそろえて開く" {
        openExportMenu

        $menu = $ui.ExportMenuButton.ContextMenu
        $menu.IsOpen | Should -Be $true
        $menu.PlacementTarget | Should -Be $ui.ExportGroup
        $menu.HorizontalOffset | Should -Be (100 - 216)
    }

    It "［出力］（本体）を押すと、最後に出力できた種類で出力先を尋ねる" {
        $workDir = "$TestDrive/export_body/work"
        [System.IO.Directory]::CreateDirectory("$TestDrive/export_body") | Out-Null
        $workspace = newTestWorkspace @{} $workDir
        $script:lastSearch = newExportSearch $workDir
        $script:lastExportFormat = "tsv"
        Mock newExportTarget { @{ Path = "$TestDrive/b.tsv"; Format = $format } }
        $script:searchService = newFakeService $false

        clickExportButton

        Should -Invoke newExportTarget -Times 1 -Exactly -ParameterFilter { $format -eq "tsv" }
    }

    It "待ち・書き込み中にメニューの種類を選んでも、中止せず、新しい要求も出さない（右クリックなどでメニューが開いたとき）" {
        $workDir = "$TestDrive/export_busy/work"
        [System.IO.Directory]::CreateDirectory("$TestDrive/export_busy") | Out-Null
        $workspace = newTestWorkspace @{} $workDir
        $script:lastSearch = newExportSearch $workDir
        Mock newExportTarget { @{ Path = "$TestDrive/export_busy/a.xlsx"; Format = $format } }
        $script:search = @{ Shared = @{} }
        exportResultsAs "xlsx"
        $pending = $script:exportPending
        $pending | Should -Not -BeNullOrEmpty

        exportResultsAs "csv"

        Should -Invoke newExportTarget -Times 1 -Exactly
        $script:exportPending | Should -Be $pending
        $ui.ExportButtonText.Text | Should -Be "出力を中止"
        $script:search = $null
        $script:exportPending = $null
    }

    It "待っている出力があるときに［出力を中止］を押すと取りやめ、▾ を押してもメニューは開かない" {
        $script:exportPending = @{ Path = "$TestDrive/x.xlsx"; Format = "xlsx" }

        openExportMenu
        $ui.ExportMenuButton.ContextMenu.IsOpen | Should -Be $false

        clickExportButton

        $ui.ExportMenuButton.ContextMenu.IsOpen | Should -Be $false
        $script:exportPending | Should -BeNullOrEmpty
        lastStatus | Should -Be (getExportCancelledText)
    }

    It "メニューから CSV を選んで書けると、その形式で出力し、本体の［出力］も CSV で出力する" {
        $workDir = "$TestDrive\export_csv\work"
        $out = "$TestDrive\export_csv\結果.csv"
        [System.IO.Directory]::CreateDirectory("$TestDrive\export_csv") | Out-Null
        $workspace = newTestWorkspace @{} $workDir
        $script:lastSearch = newExportSearch $workDir
        Mock getExportHits { , $hits }
        Mock newExportTarget { @{ Path = $out; Format = $format } }

        exportResultsAs "csv"
        pumpExport
        Should -Invoke newExportTarget -Times 1 -Exactly -ParameterFilter { $format -eq "csv" }
        $text = [System.IO.File]::ReadAllText($out)
        $text | Should -BeLike "*ファイル名,場所,種別,行,A,B*"
        $script:lastExportFormat | Should -Be "csv"
        $ui.ExportButton.ToolTip | Should -Be "CSV (.csv) で出力します"
    }

    It "出力先は、ダイアログなしで、作業フォルダの open_copy の下に決める（文字列だけで、ファイルには触れない）" {
        $workDir = "$TestDrive\export_target\work"
        $workspace = newTestWorkspace @{} $workDir
        $script:lastSearch = newExportSearch $workDir "りんご"
        Mock Test-Path { throw "出力先を決めるときは確かめない" }

        $a = newExportTarget "csv"
        $b = newExportTarget "csv"

        $root = (getOpenCopyDir $workspace).TrimEnd("\") + "\"
        $a.Path | Should -BeLike "${root}*\検索結果_りんご_*.csv"
        $a.Format | Should -Be "csv"
        $a.Path | Should -Not -Be $b.Path
        (Split-Path (Split-Path $a.Path -Parent) -Leaf).Length | Should -Be 12
        (newExportTarget "x").Format | Should -Be "xlsx"
        (newExportTarget "tsv").Path | Should -BeLike "*.tsv"
    }

    It "出力を始めても、保存のダイアログは出さない" {
        $workDir = "$TestDrive\export_nodialog\work"
        $workspace = newTestWorkspace @{} $workDir
        $script:lastSearch = newExportSearch $workDir
        $script:searchService = newFakeService $false
        Mock getExportHits { , $hits }

        exportResults "xlsx"

        $script:searchService.Requests.Count | Should -Be 1
        $script:searchService.Requests[0].OutputPath | Should -BeLike "$($workspace.TmpRoot)\*\open_copy\*\検索結果_りんご_*.xlsx"
    }

    It "ファイルを開けなければ、そのパスを知らせる" {
        $workDir = "$TestDrive\export_noexplorer\work"
        $out = "$TestDrive\export_noexplorer\結果.csv"
        [System.IO.Directory]::CreateDirectory("$TestDrive\export_noexplorer") | Out-Null
        $workspace = newTestWorkspace @{} $workDir
        $script:lastSearch = newExportSearch $workDir
        Mock getExportHits { , $hits }
        Mock newExportTarget { @{ Path = $out; Format = $format } }
        Mock openExportedFile { $false }

        exportResults "xlsx"
        pumpExport
        lastStatus | Should -Be "結果.csv に 2 件を出力しました。ファイルを開けなかったため、${out} を開いてください"
    }

    It "届かない共有を出力先に選んでも、画面はすぐ戻り（確かめない）、司令の失敗がステータスの文になる" {
        $workDir = "$TestDrive\export_unc\work"
        $workspace = newTestWorkspace @{} $workDir
        $script:lastSearch = newExportSearch $workDir
        Mock getExportHits { , $hits }
        Mock newExportTarget { @{ Path = "\\nohost.invalid\share\a.xlsx"; Format = $format } }
        Mock Test-Path { throw "画面のスレッドでは確かめない" }
        Mock getPathState { throw "画面のスレッドでは確かめない" }
        $script:searchService = newFakeService $false

        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        exportResults "xlsx"
        $watch.Stop()
        $watch.ElapsedMilliseconds | Should -BeLessThan 3000
        $script:searchService.Requests.Count | Should -Be 1
        $script:searchService.Requests[0].OutputPath | Should -Be "\\nohost.invalid\share\a.xlsx"

        invokeSearchExport $script:searchService.Requests[0]
        $script:searchService.Requests[0].Outcome | Should -Be "Failed"
        pumpExport
        lastStatus | Should -BeLike "検索結果を出力できませんでした*"
        $script:exportJob | Should -BeNullOrEmpty
    }

    It "絞り込んでいれば、絞り込み後の件数を知らせ、絞り込みを書き出しの要求に渡す" {
        $workDir = "$TestDrive\export_filtered\work"
        $workspace = newTestWorkspace @{} $workDir
        $script:lastSearch = newExportSearch $workDir
        $script:filterText = "見積"
        Mock getExportHits { , @($hits[0]) }
        Mock newExportTarget { @{ Path = "$TestDrive\export_filtered\f.csv"; Format = $format } }
        [System.IO.Directory]::CreateDirectory("$TestDrive\export_filtered") | Out-Null

        exportResults "xlsx"
        pumpExport
        lastStatus | Should -Be "絞り込み後の 1 件を f.csv に出力しました"
        $script:searchService.Requests[0].FilterText | Should -Be "見積"
        Should -Invoke getExportHits -Times 1 -Exactly -ParameterFilter { $truncated -eq $false -and $filterText -eq "見積" }
    }

    It "上限を超えた検索の出力は、表に入らなかった分も含めたと知らせる" {
        $workDir = "$TestDrive\export_overflow\work"
        $workspace = newTestWorkspace @{} $workDir
        $script:lastSearch = newExportSearch $workDir "りんご" $true
        $script:lastSearch.Shared.CountTimeout = $false
        $script:lastSearch.Shared.OverflowFailed = $false
        $script:lastSearch.Shared.Cancelled = $false
        $script:lastSearch.Shared.OverflowDir = $TestDrive
        Mock getExportHits { , $hits }
        Mock newExportTarget { @{ Path = "$TestDrive\export_overflow\o.csv"; Format = $format } }
        [System.IO.Directory]::CreateDirectory("$TestDrive\export_overflow") | Out-Null

        exportResults "xlsx"
        pumpExport
        lastStatus | Should -Be "表に入らなかった分も含めて、全 2 件を o.csv に出力しました"
        Should -Invoke getExportHits -Times 1 -Exactly -ParameterFilter { $truncated -eq $true }
    }

    It "検索の途中で押すと、出力先を決めて、終わるのを待って書く。待っている間にもう一度押すと取りやめる" {
        $workDir = "$TestDrive\export_wait\work"
        $workspace = newTestWorkspace @{} $workDir
        [System.IO.Directory]::CreateDirectory("$TestDrive\export_wait") | Out-Null
        $script:lastSearch = newExportSearch $workDir
        $script:search = @{ Shared = @{} }
        Mock getExportHits { , $hits }
        Mock newExportTarget { @{ Path = "$TestDrive\export_wait\w.csv"; Format = $format } }

        exportResults "xlsx"
        lastStatus | Should -Be "検索が終わってから出力します…"
        $ui.ExportButtonText.Text | Should -Be "出力を中止"
        $ui.ExportMenuPart.Visibility | Should -Be "Collapsed"
        $script:searchService.Requests.Count | Should -Be 0

        $script:search = $null
        startPendingExport
        $script:searchService.Requests.Count | Should -Be 1
        pumpExport
        lastStatus | Should -Be "w.csv に 2 件を出力しました（出力したファイルを開きます）"

        $script:search = @{ Shared = @{} }
        exportResults "xlsx"
        clickExportButton
        lastStatus | Should -Be "出力を中止しました。出力先のファイルは変えていません。"
        $ui.ExportButtonText.Text | Should -Be "出力"
        $script:exportPending | Should -BeNullOrEmpty
        $script:search = $null
        startPendingExport
        $script:searchService.Requests.Count | Should -Be 1
    }

    It "書いている間に押すと中止を伝え、止まったら前のファイルは残して知らせる" {
        $workDir = "$TestDrive\export_cancel\work"
        $out = "$TestDrive\export_cancel\c.csv"
        [System.IO.Directory]::CreateDirectory("$TestDrive\export_cancel") | Out-Null
        [System.IO.File]::WriteAllText($out, "前回の結果")
        $workspace = newTestWorkspace @{} $workDir
        $script:lastSearch = newExportSearch $workDir
        Mock getExportHits { , $hits }
        Mock newExportTarget { @{ Path = $out; Format = $format } }
        $script:searchService = newFakeService $false

        exportResults "xlsx"
        clickExportButton
        $request = $script:searchService.Requests[0]
        $request.Stop | Should -Be $true
        $ui.ExportButtonText.Text | Should -Be "出力を中止"
        invokeSearchExport $request
        pumpExport

        lastStatus | Should -Be "出力を中止しました。出力先のファイルは変えていません。"
        $ui.ExportButtonText.Text | Should -Be "出力"
        [System.IO.File]::ReadAllText($out) | Should -Be "前回の結果"
        Should -Invoke openExportedFile -Times 0 -Exactly
        $script:lastExportFormat | Should -Be "xlsx"
    }

    It "書いている途中は、書いた件数を知らせる" {
        $workDir = "$TestDrive\export_progress\work"
        $workspace = newTestWorkspace @{} $workDir
        $script:lastSearch = newExportSearch $workDir
        Mock getExportHits { , $hits }
        Mock newExportTarget { @{ Path = "$TestDrive\export_progress\p.csv"; Format = $format } }
        $script:searchService = newFakeService $false

        exportResults "xlsx"
        $script:searchService.Requests[0].Written = 1500
        pumpExport
        lastStatus | Should -Be "検索結果を出力中…（1,500 件）"
        $script:exportJob | Should -Not -BeNullOrEmpty
    }

    It "出力先をほかのアプリが開いていれば、閉じるよう知らせる" {
        $workDir = "$TestDrive\export_locked\work"
        $out = "$TestDrive\export_locked\l.xlsx"
        [System.IO.Directory]::CreateDirectory("$TestDrive\export_locked") | Out-Null
        [System.IO.File]::WriteAllText($out, "前回の結果")
        $workspace = newTestWorkspace @{} $workDir
        $script:lastSearch = newExportSearch $workDir
        Mock getExportHits { , $hits }
        Mock newExportTarget { @{ Path = $out; Format = $format } }

        $lock = [System.IO.File]::Open($out, "Open", "ReadWrite", "None")
        try {
            exportResults "xlsx"
        } finally {
            $lock.Dispose()
        }
        pumpExport
        $message = "l.xlsx に書き込めません。開いているアプリを閉じてから、もう一度出力してください。"
        lastStatus | Should -Be $message
        $script:banners[$script:banners.Count - 1] | Should -Be "SearchNotice|ResultSaveFailed|${message}"
        $ui.ExportButtonText.Text | Should -Be "出力"
        Should -Invoke openExportedFile -Times 0 -Exactly
    }

    It "出力先が読み取り専用なら、予期しないエラーにせず、読み取り専用を外すよう知らせる" {
        $workDir = "$TestDrive\export_readonly\work"
        $out = "$TestDrive\export_readonly\r.csv"
        [System.IO.Directory]::CreateDirectory("$TestDrive\export_readonly") | Out-Null
        [System.IO.File]::WriteAllText($out, "前回の結果")
        [System.IO.File]::SetAttributes($out, "ReadOnly")
        $workspace = newTestWorkspace @{} $workDir
        $script:lastSearch = newExportSearch $workDir
        Mock getExportHits { , $hits }
        Mock newExportTarget { @{ Path = $out; Format = $format } }

        try {
            exportResults "xlsx"
            pumpExport
        } finally {
            [System.IO.File]::SetAttributes($out, "Normal")
        }
        $message = "r.csv に書き込む権限がありません（読み取り専用など）。$out を確かめてから、もう一度出力してください。"
        lastStatus | Should -Be $message
        $script:banners[$script:banners.Count - 1] | Should -Be "SearchNotice|ResultSaveFailed|${message}"
        [System.IO.File]::ReadAllText($out) | Should -Be "前回の結果"
        Should -Invoke openExportedFile -Times 0 -Exactly
    }

    It "司令のスレッドが止まっていたら、理由を知らせて出力を終える" {
        $workDir = "$TestDrive\export_dead\work"
        $workspace = newTestWorkspace @{} $workDir
        $script:lastSearch = newExportSearch $workDir
        Mock getExportHits { , $hits }
        Mock newExportTarget { @{ Path = "$TestDrive\export_dead\d.csv"; Format = $format } }
        $script:searchService = newFakeService $false
        $script:searchService.Running = $false
        $script:searchService.Failure = "検索のスレッドが止まりました。"

        exportResults "xlsx"
        pumpExport
        lastStatus | Should -Be "検索結果を出力できませんでした：検索のスレッドが止まりました。"
        $ui.ExportButtonText.Text | Should -Be "出力"
        $script:exportJob | Should -BeNullOrEmpty
    }

    It "ワークスペースを変えたとき（clearExport）、待っている出力は取りやめ、書いている出力には中止を伝える" {
        $workDir = "$TestDrive\export_clear\work"
        $workspace = newTestWorkspace @{} $workDir
        $script:lastSearch = newExportSearch $workDir
        Mock getExportHits { , $hits }
        Mock newExportTarget { @{ Path = "$TestDrive\export_clear\c.csv"; Format = $format } }
        $script:searchService = newFakeService $false

        exportResults "xlsx"
        clearExport
        $script:searchService.Requests[0].Stop | Should -Be $true

        $script:exportJob = $null
        $script:exportPending = @{ Path = "x.csv"; Format = "csv" }
        clearExport
        $script:exportPending | Should -BeNullOrEmpty
        $ui.ExportButtonText.Text | Should -Be "出力"
    }
}

Describe "画面の操作" -Tag Unit {
    BeforeEach {
        $script:statuses = New-Object System.Collections.Generic.List[string]
    }

    It "右クリックメニューとボタンは、それぞれの開き方・操作を呼ぶ" {
        Mock openSource { }
        Mock openSourceFolder { }
        Mock copySelectedRows { }
        Mock copySourcePath { }
        Mock clickExportButton { }
        Mock openExportMenu { }
        Mock exportResultsAs { }

        & $ui.MenuOpen.Handlers["Click"]
        & $ui.MenuOpenReadOnly.Handlers["Click"]
        & $ui.MenuOpenNew.Handlers["Click"]
        & $ui.OpenButton.Handlers["Click"]
        & $ui.MenuOpenFolder.Handlers["Click"]
        & $ui.OpenFolderButton.Handlers["Click"]
        & $ui.MenuCopy.Handlers["Click"]
        & $ui.MenuCopyPath.Handlers["Click"]
        & $ui.ExportButton.Handlers["Click"]
        & $ui.ExportMenuButton.Handlers["Click"]
        & $ui.MenuExportXlsx.Handlers["Click"]
        & $ui.MenuExportCsv.Handlers["Click"]
        & $ui.MenuExportTsv.Handlers["Click"]

        Should -Invoke openSource -Times 1 -Exactly -ParameterFilter { $mode -eq ${openModeNormal} }
        Should -Invoke openSource -Times 1 -Exactly -ParameterFilter { $mode -eq ${openModeReadOnly} }
        Should -Invoke openSource -Times 1 -Exactly -ParameterFilter { $mode -eq ${openModeNew} }
        Should -Invoke openSource -Times 4 -Exactly
        Should -Invoke openSourceFolder -Times 2 -Exactly
        Should -Invoke copySelectedRows -Times 1 -Exactly
        Should -Invoke copySourcePath -Times 1 -Exactly
        Should -Invoke clickExportButton -Times 1 -Exactly
        Should -Invoke openExportMenu -Times 1 -Exactly
        Should -Invoke exportResultsAs -Times 1 -Exactly -ParameterFilter { $format -eq "xlsx" }
        Should -Invoke exportResultsAs -Times 1 -Exactly -ParameterFilter { $format -eq "csv" }
        Should -Invoke exportResultsAs -Times 1 -Exactly -ParameterFilter { $format -eq "tsv" }
    }

    It "［開く］の［▾］のメニューで選ぶと、その開き方を既定にして保存し、開く" {
        Mock openSource { }
        Mock writeOpenMode { }

        & $ui.MenuOpenModeReadOnly.Handlers["Click"]
        getOpenMode | Should -Be ${openModeReadOnly}
        Should -Invoke writeOpenMode -Times 1 -Exactly -ParameterFilter { $mode -eq ${openModeReadOnly} }
        Should -Invoke openSource -Times 1 -Exactly -ParameterFilter { $mode -eq ${openModeReadOnly} }

        & $ui.MenuOpenModeNew.Handlers["Click"]
        getOpenMode | Should -Be ${openModeNew}
        Should -Invoke writeOpenMode -Times 1 -Exactly -ParameterFilter { $mode -eq ${openModeNew} }
        Should -Invoke openSource -Times 1 -Exactly -ParameterFilter { $mode -eq ${openModeNew} }
    }

    It "Enter は、見出し・続きの行では閉じる開く・続きを出す（activateResultItem）、ほかの行では元のファイルを開く" {
        Mock openSource { }
        Mock activateResultItem { $item -is [FileGroup] }
        $group = [FileGroup]::new()

        $ui.ResultGrid.SelectedItem = $group
        $e = [pscustomobject]@{ Key = "Return"; Handled = $false }
        & $ui.ResultGrid.Handlers["PreviewKeyDown"] $null $e
        $e.Handled | Should -Be $true
        Should -Invoke activateResultItem -Times 1 -Exactly -ParameterFilter { [object]::ReferenceEquals($item, $group) }
        Should -Invoke openSource -Times 0 -Exactly

        $ui.ResultGrid.SelectedItem = newRow
        & $ui.ResultGrid.Handlers["PreviewKeyDown"] $null ([pscustomobject]@{ Key = "Return"; Handled = $false })
        Should -Invoke openSource -Times 1 -Exactly
    }

    It "メニューの［この結果を折りたたむ］は、選んでいる見出しだけを閉じる・開く" {
        Mock toggleFileGroup { }

        $ui.ResultGrid.SelectedItem = newRow
        & $ui.MenuToggleGroup.Handlers["Click"]
        Should -Invoke toggleFileGroup -Times 0 -Exactly

        $expectedGroup = [FileGroup]::new()
        $ui.ResultGrid.SelectedItem = $expectedGroup
        & $ui.MenuToggleGroup.Handlers["Click"]
        Should -Invoke toggleFileGroup -Times 1 -Exactly -ParameterFilter { [object]::ReferenceEquals($group, $expectedGroup) }
    }

    It "ほかのキーは扱わない" {
        Mock openSource { }
        $e = [pscustomobject]@{ Key = "A"; Handled = $false }
        & $ui.ResultGrid.Handlers["PreviewKeyDown"] $null $e
        $e.Handled | Should -Be $false
        Should -Invoke openSource -Times 0 -Exactly
    }

    It "行の外（クリックした要素が無い）でのダブルクリックでは開かない" {
        Mock openSource { }
        & $ui.ResultGrid.Handlers["MouseDoubleClick"] $null ([pscustomobject]@{ OriginalSource = $null })
        Should -Invoke openSource -Times 0 -Exactly
    }

    It "行の上のダブルクリックで元のファイルを開き、見出しの行・列見出し・スクロールバー・行の外では開かない" -Skip:(!$sta) {
        Mock openSource { }
        $double = { param($source) & $ui.ResultGrid.Handlers["MouseDoubleClick"] $null ([pscustomobject]@{ OriginalSource = $source }) }

        $row = New-Object System.Windows.Controls.DataGridRow
        $row.Item = newRow
        & $double $row
        Should -Invoke openSource -Times 1 -Exactly

        $header = New-Object System.Windows.Controls.DataGridRow
        $header.Item = [FileGroup]::new()
        & $double $header
        & $double (New-Object System.Windows.Controls.Primitives.DataGridColumnHeader)
        & $double (New-Object System.Windows.Controls.Primitives.ScrollBar)
        & $double (New-Object System.Windows.Controls.TextBlock)  # 行の中に無い要素（親をたどっても行に着かない）
        Should -Invoke openSource -Times 1 -Exactly
    }
}

Describe "右クリックのメニュー" -Tag Unit {
    BeforeEach {
        $script:statuses = New-Object System.Collections.Generic.List[string]
        $ui.ResultGrid.SelectedItem = $null
        $ui.ResultGrid.SelectedItems = @()
    }

    Context "showResultMenu" {
        BeforeEach {
            Mock setContextMenuItems { }
        }

        It "マウスで開いたとき、行の無い所では出さない" {
            $ui.ResultGrid.SelectedItem = newRow
            showResultMenu $null 10 | Should -Be $false
            Should -Invoke setContextMenuItems -Times 0 -Exactly
        }

        It "選んでいる項目が無ければ出さない" {
            showResultMenu $null -1 | Should -Be $false
            Should -Invoke setContextMenuItems -Times 0 -Exactly
        }

        It "「続き」の行を選んでいるときは出さない（開く元のファイルが無い）" {
            $ui.ResultGrid.SelectedItem = [MoreRowsItem]::new()
            showResultMenu $null -1 | Should -Be $false
            Should -Invoke setContextMenuItems -Times 0 -Exactly
        }

        It "キーボードで開いたとき（CursorLeft が負）は、行の確かめを飛ばして選んでいる行に出す" {
            $ui.ResultGrid.SelectedItem = newRow
            showResultMenu $null -1 | Should -Be $true
            Should -Invoke setContextMenuItems -Times 1 -Exactly -ParameterFilter { @($items | Where-Object { $_.Id -eq "openFolder" }).Count -eq 1 -and @($items | Where-Object { $_.Id -eq "toggleGroup" }).Count -eq 0 }
        }

        It "見出しでは group のメニューにし、開いているかを渡す（<name>）" -TestCases @(
            @{ name = "開いている"; expanded = $true }
            @{ name = "閉じている"; expanded = $false }
        ) {
            $group = [FileGroup]::new()
            $group.IsExpanded = $expanded
            $group.Book = "見積.xlsx"
            $ui.ResultGrid.SelectedItem = $group
            showResultMenu $null -1 | Should -Be $true
            $expectedJson = getResultMenuItems "group" (getOpenMode) $expanded $true | ConvertTo-Json -Compress
            Should -Invoke setContextMenuItems -Times 1 -Exactly -ParameterFilter { ($items | ConvertTo-Json -Compress) -eq $expectedJson }
        }

        It "読み取り専用で開けないファイル（テキストなど）では、［読み取り専用で開く］を出さない（<name>）" -TestCases @(
            @{ name = "行"; book = "議事メモ.txt"; target = "row" }
            @{ name = "見出し"; book = "仕様書.pdf"; target = "group" }
        ) {
            if ($target -eq "group") {
                $item = [FileGroup]::new()
                $item.Book = $book
            } else {
                $item = newRow $book $false $true
            }
            $ui.ResultGrid.SelectedItem = $item
            showResultMenu $null -1 | Should -Be $true
            Should -Invoke setContextMenuItems -Times 1 -Exactly -ParameterFilter { @($items | Where-Object { $_.Id -eq "openReadOnly" }).Count -eq 0 -and @($items | Where-Object { $_.Id -eq "openNew" }).Count -eq 1 }
        }
    }

    Context "画面の受け口" {
        It "行の無い所の右クリックでは、メニューを出さない（Handled）" {
            $e = [pscustomobject]@{ OriginalSource = $null; CursorLeft = 10; Handled = $false }
            & $ui.ResultGrid.Handlers["ContextMenuOpening"] $null $e
            $e.Handled | Should -Be $true
        }

        It "選んでいる項目が無ければ、キーボードで開いてもメニューを出さない（Handled）" {
            $e = [pscustomobject]@{ OriginalSource = $null; CursorLeft = -1; Handled = $false }
            & $ui.ResultGrid.Handlers["ContextMenuOpening"] $null $e
            $e.Handled | Should -Be $true
        }

        It "出せるときは Handled にしない" {
            Mock setContextMenuItems { }
            $ui.ResultGrid.SelectedItem = newRow
            $e = [pscustomobject]@{ OriginalSource = $null; CursorLeft = -1; Handled = $false }
            & $ui.ResultGrid.Handlers["ContextMenuOpening"] $null $e
            $e.Handled | Should -Be $false
        }

        It "右クリックで押した所に行が無ければ、選びを変えない" {
            Mock selectResultRowForMenu { }
            & $ui.ResultGrid.Handlers["PreviewMouseRightButtonDown"] $null ([pscustomobject]@{ OriginalSource = $null })
            Should -Invoke selectResultRowForMenu -Times 0 -Exactly
        }

        It "右クリックで押した所に行があれば、その行を選ぶ" -Skip:(!$sta) {
            Mock selectResultRowForMenu { }
            $row = New-Object System.Windows.Controls.DataGridRow
            $row.Item = newRow
            & $ui.ResultGrid.Handlers["PreviewMouseRightButtonDown"] $null ([pscustomobject]@{ OriginalSource = $row })
            Should -Invoke selectResultRowForMenu -Times 1 -Exactly
        }
    }

    Context "selectResultRowForMenu" {
        It "選ばれていない行を右クリックすると、その行だけを選ぶ" -Skip:(!$sta) {
            $ui.ResultGrid.SelectedItem = newRow "他.xlsx"
            $row = New-Object System.Windows.Controls.DataGridRow
            $row.Item = newRow
            selectResultRowForMenu $row
            [object]::ReferenceEquals($ui.ResultGrid.SelectedItem, $row.Item) | Should -Be $true
        }

        It "複数選んだ行の中を右クリックしても、選びは変わらない（1 件だけの操作は、先に選んだ行 SelectedItem に効く）" -Skip:(!$sta) {
            $first = newRow "先.xlsx"
            $second = newRow "後.xlsx"
            $ui.ResultGrid.SelectedItem = $first
            $ui.ResultGrid.SelectedItems = @($first, $second)
            $row = New-Object System.Windows.Controls.DataGridRow
            $row.Item = $second
            $row.IsSelected = $true
            selectResultRowForMenu $row
            [object]::ReferenceEquals($ui.ResultGrid.SelectedItem, $first) | Should -Be $true
        }

        It "見出しを含む選びの中の行を右クリックすると、その行だけを選ぶ" -Skip:(!$sta) {
            $item = newRow
            $heading = [FileGroup]::new()
            $ui.ResultGrid.SelectedItem = $heading
            $ui.ResultGrid.SelectedItems = @($heading, $item)
            $row = New-Object System.Windows.Controls.DataGridRow
            $row.Item = $item
            $row.IsSelected = $true
            selectResultRowForMenu $row
            [object]::ReferenceEquals($ui.ResultGrid.SelectedItem, $item) | Should -Be $true
        }
    }

    Context "setContextMenuItems" {
        It "判断層が決めた並びで、区切り・太字・文言を並べる" -Skip:(!$sta) {
            $menu = New-Object System.Windows.Controls.ContextMenu
            $parts = @{ a = New-Object System.Windows.Controls.MenuItem; b = New-Object System.Windows.Controls.MenuItem; c = New-Object System.Windows.Controls.MenuItem }
            $parts.c.Header = "前の文言"
            $parts.c.FontWeight = [System.Windows.FontWeights]::Bold
            $menu.Items.Add($parts.a) | Out-Null  # 前の中身は消える
            $items = @(
                @{ Id = "b"; Header = "二番目の文言"; Bold = $false }
                @{ Id = "separator"; Header = ""; Bold = $false }
                @{ Id = "a"; Header = "先頭の文言"; Bold = $true }
                @{ Id = "c"; Header = "三番目"; Bold = $false }
            )
            setContextMenuItems $menu $items $parts

            $menu.Items.Count | Should -Be 4
            $menu.Items[0].Header | Should -Be "二番目の文言"
            $menu.Items[1] | Should -BeOfType [System.Windows.Controls.Separator]
            $menu.Items[2].Header | Should -Be "先頭の文言"
            $menu.Items[2].FontWeight | Should -Be ([System.Windows.FontWeights]::Bold)
            $menu.Items[3].Header | Should -Be "三番目"
            $menu.Items[3].FontWeight | Should -Be ([System.Windows.FontWeights]::Normal)
        }
    }
}
