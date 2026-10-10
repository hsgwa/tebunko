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
    . "${scriptsDir}\tebunko\ui\types.ps1"
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
            ResultGrid       = newFakeControl @("MouseDoubleClick", "PreviewKeyDown", "PreviewMouseRightButtonDown", "ContextMenuOpening") @{ SelectedItem = $null; SelectedItems = @() }
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
            ExportButton     = newFakeControl @("Click")
        }
    }

    # 画面のほかのファイル（shell.ps1・result_list.ps1・folder_dialog.ps1）にある関数の代わり。テストごとに Mock する
    function setStatus { param([string]$text) $script:statuses.Add($text) }
    function safe { param([scriptblock]$block) & $block }
    function getCurrentHitRow { }
    function getSelectedRows { }
    function getViewRows { }
    function toggleFileGroup { param($group) }
    function showConfirm { param([string]$heading, [object[]]$facts, [object[]]$choices) }
    function selectFolder { param([string]$description, [string]$initial, $owner, [bool]$knownExisting) }
    function factGone { param([string]$title, [string]$detail) "✗ ${title}" }
    function factNext { param([string]$title, [string]$detail) "→ ${title}" }
    function testNetworkPath { param([string]$path) $false }  # 既定はローカル。テストごとに Mock で切り替える
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
        Should -Invoke showConfirm -Times 1 -Exactly -ParameterFilter { $facts -and $facts[0] -eq "✗ 元のファイルを確かめられませんでした" }
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
        Should -Invoke showConfirm -Times 1 -Exactly -ParameterFilter { $facts -and $facts[0] -eq "✗ 選んだフォルダの中にありませんでした" }
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

    It "知らない開き方は、通常の開き方と同じにする" {
        Mock Invoke-Item { }

        openWithShell "$TestDrive\[確定]見積.xlsx" "unknown" | Should -Be $true
        Should -Invoke Invoke-Item -Times 1 -Exactly -ParameterFilter { $LiteralPath -eq "$TestDrive\[確定]見積.xlsx" }
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
        Mock openWithShell { $false }

        openSource ${openModeReadOnly}
        # 選んだ開き方（読み取り専用）ではなく、常に通常の開き方で openWithShell を呼ぶ（動詞を試さない）
        Should -Invoke openWithShell -Times 1 -Exactly -ParameterFilter { $path -eq "C:\data\議事メモ.txt" -and $mode -eq ${openModeNormal} }
        lastStatus | Should -Be "開きました：C:\data\議事メモ.txt"
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

    It "既定のアプリで開くテキスト（.py など）は、openWithNotepad を呼ばない" {
        Mock getCurrentHitRow { newRow "解析.py" $false $true }
        Mock findSourceFile { param ($row, $onFound) & $onFound "C:\data\解析.py" }
        Mock openWithNotepad { }
        Mock openWithShell { $true }

        openSource ${openModeReadOnly}
        Should -Invoke openWithShell -Times 1 -Exactly -ParameterFilter { $path -eq "C:\data\解析.py" }
        Should -Invoke openWithNotepad -Times 0 -Exactly
        lastStatus | Should -Be "開きました：C:\data\解析.py"
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

Describe "exportResults" -Tag Io {
    BeforeAll {
        $hits = @(
            [pscustomobject]@{ Book = "見積.xlsx"; Location = "Sheet1"; LineNumber = 3; Line = "りんご`t100"; RelDir = "sub" },
            [pscustomobject]@{ Book = "報告.docx"; Location = "ページ001"; LineNumber = 1; Line = "りんごの報告"; RelDir = "" }
        )
    }

    BeforeEach {
        $script:statuses = New-Object System.Collections.Generic.List[string]
    }

    It "検索する前は出力しない" {
        $script:lastSearch = $null
        Mock Invoke-Item { }

        exportResults
        lastStatus | Should -Be "先に検索してください。"
        Should -Invoke Invoke-Item -Times 0 -Exactly
    }

    It "表示中の行を search_results.txt に書き出して開く" {
        $workDir = "$TestDrive\export\work"
        $resultFile = "$workDir\search_results.txt"
        $workspace = newTestWorkspace @{} $workDir
        $script:lastSearch = @{ Word = "りんご" }
        $script:hitCount = 2
        Mock getViewRows { , $hits }
        Mock Invoke-Item { }

        exportResults
        $lines = [System.IO.File]::ReadAllLines($resultFile)
        $lines[0] | Should -Be "【検索文字列　りんご】 2 件"
        $lines[2] | Should -Be "sub\見積.xlsx`t[シート]Sheet1`tセル`t3`tりんご`t100"
        Should -Invoke Invoke-Item -Times 1 -Exactly -ParameterFilter { $LiteralPath -eq $resultFile }
        lastStatus | Should -Be "search_results.txt に出力しました（2 件）"
    }

    It "前回の search_results.txt は追記せずに置き換え、BOM 付き UTF-8 で書く（0 件でも書き出す）" {
        $workDir = "$TestDrive\export_empty\work"
        $resultFile = "$workDir\search_results.txt"
        $workspace = newTestWorkspace @{} $workDir
        [System.IO.Directory]::CreateDirectory($workDir) | Out-Null
        [System.IO.File]::WriteAllText($resultFile, "前回の結果`r`n前回の結果`r`n前回の結果`r`n前回の結果`r`n")
        $script:lastSearch = @{ Word = "無い言葉" }
        $script:hitCount = 0
        Mock getViewRows { , @() }
        Mock Invoke-Item { }

        exportResults
        $bytes = [System.IO.File]::ReadAllBytes($resultFile)
        @($bytes[0..2]) | Should -Be @(0xEF, 0xBB, 0xBF)
        $lines = [System.IO.File]::ReadAllLines($resultFile)
        $lines[0] | Should -Be "【検索文字列　無い言葉】 0 件"
        $lines -contains "前回の結果" | Should -Be $false
        lastStatus | Should -Be "search_results.txt に出力しました（0 件）"
    }

    It "絞り込んでいれば、絞り込み後の件数を知らせる" {
        $workDir = "$TestDrive\export_filtered\work"
        $resultFile = "$workDir\search_results.txt"
        $workspace = newTestWorkspace @{} $workDir
        $script:lastSearch = @{ Word = "りんご" }
        $script:hitCount = 1234
        Mock getViewRows { , $hits }
        Mock Invoke-Item { }

        exportResults
        lastStatus | Should -Be "絞り込み後の 2 件を search_results.txt に出力しました"
    }

    It "search_results.txt をほかのアプリが開いていれば、閉じるよう知らせる" {
        $workDir = "$TestDrive\export_locked\work"
        $resultFile = "$workDir\search_results.txt"
        $workspace = newTestWorkspace @{} $workDir
        [System.IO.Directory]::CreateDirectory($workDir) | Out-Null
        $script:lastSearch = @{ Word = "りんご" }
        Mock getViewRows { , $hits }
        Mock Invoke-Item { }

        $lock = [System.IO.File]::Open($resultFile, "Create", "ReadWrite", "None")
        try {
            exportResults
        } finally {
            $lock.Dispose()
        }
        lastStatus | Should -Be "search_results.txt に書き込めません。開いているアプリを閉じてから、もう一度出力してください。"
        Should -Invoke Invoke-Item -Times 0 -Exactly
    }

    It "search_results.txt が読み取り専用なら、予期しないエラーにせず、読み取り専用を外すよう知らせる" {
        $workDir = "$TestDrive\export_readonly\work"
        $resultFile = "$workDir\search_results.txt"
        $workspace = newTestWorkspace @{} $workDir
        [System.IO.Directory]::CreateDirectory($workDir) | Out-Null
        [System.IO.File]::WriteAllText($resultFile, "前回の結果")
        [System.IO.File]::SetAttributes($resultFile, "ReadOnly")
        $script:lastSearch = @{ Word = "りんご" }
        Mock getViewRows { , $hits }
        Mock Invoke-Item { }

        try {
            exportResults
        } finally {
            [System.IO.File]::SetAttributes($resultFile, "Normal")
        }
        lastStatus | Should -Be "search_results.txt に書き込む権限がありません（読み取り専用など）。$resultFile を確かめてから、もう一度出力してください。"
        [System.IO.File]::ReadAllText($resultFile) | Should -Be "前回の結果"
        Should -Invoke Invoke-Item -Times 0 -Exactly
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
        Mock exportResults { }

        & $ui.MenuOpen.Handlers["Click"]
        & $ui.MenuOpenReadOnly.Handlers["Click"]
        & $ui.MenuOpenNew.Handlers["Click"]
        & $ui.OpenButton.Handlers["Click"]
        & $ui.MenuOpenFolder.Handlers["Click"]
        & $ui.OpenFolderButton.Handlers["Click"]
        & $ui.MenuCopy.Handlers["Click"]
        & $ui.MenuCopyPath.Handlers["Click"]
        & $ui.ExportButton.Handlers["Click"]

        Should -Invoke openSource -Times 1 -Exactly -ParameterFilter { $mode -eq ${openModeNormal} }
        Should -Invoke openSource -Times 1 -Exactly -ParameterFilter { $mode -eq ${openModeReadOnly} }
        Should -Invoke openSource -Times 1 -Exactly -ParameterFilter { $mode -eq ${openModeNew} }
        Should -Invoke openSource -Times 4 -Exactly
        Should -Invoke openSourceFolder -Times 2 -Exactly
        Should -Invoke copySelectedRows -Times 1 -Exactly
        Should -Invoke copySourcePath -Times 1 -Exactly
        Should -Invoke exportResults -Times 1 -Exactly
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

    It "Enter は、見出しの行では閉じる・開く、ほかの行では元のファイルを開く" {
        Mock openSource { }
        Mock toggleFileGroup { }
        $group = [FileGroup]::new()

        $ui.ResultGrid.SelectedItem = $group
        $e = [pscustomobject]@{ Key = "Return"; Handled = $false }
        & $ui.ResultGrid.Handlers["PreviewKeyDown"] $null $e
        $e.Handled | Should -Be $true
        Should -Invoke toggleFileGroup -Times 1 -Exactly
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
            $ui.ResultGrid.SelectedItem = $group
            showResultMenu $null -1 | Should -Be $true
            $expectedJson = getResultMenuItems "group" (getOpenMode) $expanded | ConvertTo-Json -Compress
            Should -Invoke setContextMenuItems -Times 1 -Exactly -ParameterFilter { ($items | ConvertTo-Json -Compress) -eq $expectedJson }
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
