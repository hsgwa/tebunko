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
        $menu = { newFakeControl @("Click") @{ InputGestureText = "" } }
        return @{
            ResultGrid       = newFakeControl @("MouseDoubleClick", "PreviewKeyDown") @{ SelectedItem = $null }
            MenuOpen         = & $menu
            MenuOpenReadOnly = & $menu
            MenuOpenNew      = & $menu
            OpenModeCombo    = newFakeControl @("SelectionChanged") @{ SelectedItem = $null; SelectedIndex = -1; Items = @() }
            MenuOpenFolder   = & $menu
            OpenButton       = newFakeControl @("Click")
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
    function selectFolder { param([string]$description, [string]$initial) }
    function factGone { param([string]$title, [string]$detail) "✗ ${title}" }
    function factNext { param([string]$title, [string]$detail) "→ ${title}" }

    $ui = newFakeUi
    $window = New-Object PSObject -Property @{ Cursor = $null }
    . "${scriptsDir}\tebunko\ui\open_source.ps1"

    function newRow {
        param (
            [string]$book = "見積.xlsx",
            [bool]$isExcel = $true
        )

        [pscustomobject]@{
            Root = "C:\index"; RelDir = "営業\sub"; RelPath = "営業\sub\$book"; Book = $book
            Location = "見積[図形]"; MatchCell = "B2"; IsExcel = $isExcel
        }
    }

    function lastStatus { $script:statuses[$script:statuses.Count - 1] }
}

Describe "getExistingFolder" -Tag Io {
    It "上のフォルダのうち、存在する最も深いフォルダを返す" {
        [System.IO.Directory]::CreateDirectory("$TestDrive\exist\a") | Out-Null
        getExistingFolder "$TestDrive\exist\a\b\c\見積.xlsx" | Should -Be "$TestDrive\exist\a"
    }

    It "フォルダ名に [ ] があっても（ワイルドカードとして扱わず）確かめる" {
        [System.IO.Directory]::CreateDirectory("$TestDrive\exist\[旧]営業") | Out-Null
        getExistingFolder "$TestDrive\exist\[旧]営業\無い\見積.xlsx" | Should -Be "$TestDrive\exist\[旧]営業"
    }

    It "どこも存在しなければ空" {
        getExistingFolder "無いフォルダ_open_source\見積.xlsx" | Should -Be ""
    }
}

Describe "findSourceFile" -Tag Io {
    BeforeEach {
        $script:statuses = New-Object System.Collections.Generic.List[string]
        $script:sourceFolderMaps = @{ 記録 = 1 }
    }

    It "記録した場所にあれば、そのパスを返す" {
        newTsv "$TestDrive\known\sub\見積.xlsx" @("x")
        Mock getSourceLocation { @{ Name = "営業"; Folder = "$TestDrive\known"; Rest = "sub"; Known = $true } }
        Mock showConfirm { }

        findSourceFile (newRow) | Should -Be "$TestDrive\known\sub\見積.xlsx"
        Should -Invoke showConfirm -Times 0 -Exactly
    }

    It "ファイル名に [ ] があっても（ワイルドカードとして扱わず）見つける" {
        newTsv "$TestDrive\bracket\[確定]見積.xlsx" @("x")
        Mock getSourceLocation { @{ Name = "営業"; Folder = "$TestDrive\bracket"; Rest = ""; Known = $true } }
        Mock showConfirm { }

        findSourceFile (newRow "[確定]見積.xlsx") | Should -Be "$TestDrive\bracket\[確定]見積.xlsx"
        Should -Invoke showConfirm -Times 0 -Exactly
    }

    It "260 文字を超えるパスのファイルも見つける" {
        $deep = "$TestDrive\long\" + ("深いフォルダ" * 10) + "\" + ("もっと深いフォルダ" * 10) + "\" + ("さらに深いフォルダ" * 10)
        newTsv (toLongPath "$deep\見積.xlsx") @("x")
        $script:deep = $deep
        Mock getSourceLocation { @{ Name = "営業"; Folder = $script:deep; Rest = ""; Known = $true } }
        Mock showConfirm { }

        try {
            ("$deep\見積.xlsx").Length | Should -BeGreaterThan 260
            findSourceFile (newRow) | Should -Be "$deep\見積.xlsx"
            Should -Invoke showConfirm -Times 0 -Exactly
        } finally {
            # TestDrive の後片付けは長いパスを消せないため、ここで消す
            removeDirectoryRetry "$TestDrive\long"
        }
    }

    It "記録が無く、元のファイルがインデックスの直下（相対フォルダが空）なら、知らせるパスに \ を重ねない" {
        Mock getSourceLocation { @{ Name = "営業"; Folder = ""; Rest = ""; Known = $false } }
        Mock showConfirm { $null }
        $row = newRow
        $row.RelDir = ""

        findSourceFile $row | Should -Be $null
        lastStatus | Should -Be "元のファイルが見つかりません：C:\index\見積.xlsx"
    }

    It "記録した場所に無くても、同じ場所を指す別の書き方で見つかれば、その場所を記録して返す" {
        newTsv "$TestDrive\alias\sub\見積.xlsx" @("x")
        Mock getSourceLocation { @{ Name = "営業"; Folder = "Z:\営業"; Rest = "sub"; Known = $true } }
        Mock getFolderPathAliases { @("Z:\営業", "$TestDrive\missing", "$TestDrive\alias") }
        Mock setIndexSourceFolder { }

        findSourceFile (newRow) | Should -Be "$TestDrive\alias\sub\見積.xlsx"
        Should -Invoke setIndexSourceFolder -Times 1 -Exactly -ParameterFilter { $name -eq "営業" -and $folder -eq "$TestDrive\alias" }
        $script:sourceFolderMaps.Count | Should -Be 0
        lastStatus | Should -Be "インデックス [営業] の元のフォルダを $TestDrive\alias に変えました"
    }

    It "別の書き方で見つかっても、インデックス名が分からなければ記録しない" {
        newTsv "$TestDrive\alias2\sub\見積.xlsx" @("x")
        Mock getSourceLocation { @{ Name = ""; Folder = "Z:\営業"; Rest = "sub"; Known = $true } }
        Mock getFolderPathAliases { @("Z:\営業", "$TestDrive\alias2") }
        Mock setIndexSourceFolder { }

        findSourceFile (newRow) | Should -Be "$TestDrive\alias2\sub\見積.xlsx"
        Should -Invoke setIndexSourceFolder -Times 0 -Exactly
    }

    It "見つからず、フォルダを選ばなければ `$null（記録した場所の近くを選ぶ画面の初期位置にする）" {
        [System.IO.Directory]::CreateDirectory("$TestDrive\gone") | Out-Null
        Mock getSourceLocation { @{ Name = "営業"; Folder = "$TestDrive\gone"; Rest = "sub"; Known = $true } }
        Mock getFolderPathAliases { @("$TestDrive\gone") }
        Mock showConfirm { "pick" }
        Mock selectFolder { "" }

        findSourceFile (newRow) | Should -Be $null
        Should -Invoke selectFolder -Times 1 -Exactly -ParameterFilter { $initial -eq "$TestDrive\gone" }
        lastStatus | Should -Be "元のファイルが見つかりません：$TestDrive\gone\sub\見積.xlsx"
    }

    It "確認で選ぶのをやめれば `$null" {
        Mock getSourceLocation { @{ Name = "営業"; Folder = ""; Rest = "sub"; Known = $false } }
        Mock showConfirm { $null }
        Mock selectFolder { }

        findSourceFile (newRow) | Should -Be $null
        Should -Invoke selectFolder -Times 0 -Exactly
        lastStatus | Should -Be "元のファイルが見つかりません：C:\index\営業\sub\見積.xlsx"
    }

    It "記録が無くても、選んだフォルダの中で見つかれば、元のフォルダを記録して返す" {
        newTsv "$TestDrive\moved\営業\sub\見積.xlsx" @("x")
        Mock getSourceLocation { @{ Name = "営業"; Folder = ""; Rest = "sub"; Known = $false } }
        Mock showConfirm { "pick" }
        Mock selectFolder { "$TestDrive\moved\営業\sub" }  # ファイルのあるフォルダを選んだ
        Mock setIndexSourceFolder { }

        findSourceFile (newRow) | Should -Be "$TestDrive\moved\営業\sub\見積.xlsx"
        Should -Invoke setIndexSourceFolder -Times 1 -Exactly -ParameterFilter { $name -eq "営業" -and $folder -eq "$TestDrive\moved\営業" }
        $script:sourceFolderMaps.Count | Should -Be 0
    }

    It "選んだフォルダで見つかっても、インデックス名が分からなければ記録せずに開く" {
        newTsv "$TestDrive\moved2\sub\見積.xlsx" @("x")
        Mock getSourceLocation { @{ Name = ""; Folder = ""; Rest = "sub"; Known = $false } }
        Mock showConfirm { "pick" }
        Mock selectFolder { "$TestDrive\moved2" }
        Mock setIndexSourceFolder { }

        findSourceFile (newRow) | Should -Be "$TestDrive\moved2\sub\見積.xlsx"
        Should -Invoke setIndexSourceFolder -Times 0 -Exactly
        lastStatus | Should -Be "開きました：$TestDrive\moved2\sub\見積.xlsx"
    }

    It "選んだフォルダに無ければ、選んだフォルダを初期位置にしてもう一度聞く" {
        [System.IO.Directory]::CreateDirectory("$TestDrive\empty") | Out-Null
        $script:confirmCount = 0
        Mock getSourceLocation { @{ Name = "営業"; Folder = ""; Rest = ""; Known = $false } }
        Mock showConfirm { $script:confirmCount++; if ($script:confirmCount -eq 1) { "pick" } }
        Mock selectFolder { "$TestDrive\empty" }

        findSourceFile (newRow) | Should -Be $null
        Should -Invoke showConfirm -Times 2 -Exactly
        Should -Invoke showConfirm -Times 1 -Exactly -ParameterFilter { $facts[0] -eq "✗ 選んだフォルダの中にありませんでした" }
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

Describe "getOpenMode / setOpenMode / updateOpenMenu" -Tag Unit {
    BeforeAll {
        $items = @(
            [pscustomobject]@{ Tag = ${openModeNormal} },
            [pscustomobject]@{ Tag = ${openModeReadOnly} },
            [pscustomobject]@{ Tag = ${openModeNew} }
        )
    }

    It "［開き方］で選んだものを返す" {
        $ui.OpenModeCombo.SelectedItem = $items[1]
        getOpenMode | Should -Be ${openModeReadOnly}
    }

    It "選んでいない・知らない値なら通常" {
        $ui.OpenModeCombo.SelectedItem = $null
        getOpenMode | Should -Be ${openModeNormal}
        $ui.OpenModeCombo.SelectedItem = [pscustomobject]@{ Tag = "unknown" }
        getOpenMode | Should -Be ${openModeNormal}
    }

    It "設定の値の項目を選ぶ（選んでいる間だけ読み込み中の印を立てる）" {
        $ui.OpenModeCombo.Items = $items
        setOpenMode ${openModeNew}
        $ui.OpenModeCombo.SelectedItem | Should -Be $items[2]
        $script:loadingOpenMode | Should -Be $false
    }

    It "知らない値なら先頭を選ぶ" {
        $ui.OpenModeCombo.Items = $items
        $ui.OpenModeCombo.SelectedIndex = 2
        setOpenMode "unknown"
        $ui.OpenModeCombo.SelectedIndex | Should -Be 0
    }

    It "右クリックメニューの、既定の開き方にだけ Enter を出す" {
        $ui.OpenModeCombo.SelectedItem = $items[1]
        updateOpenMenu
        $ui.MenuOpen.InputGestureText | Should -Be ""
        $ui.MenuOpenReadOnly.InputGestureText | Should -Be "Enter"
        $ui.MenuOpenNew.InputGestureText | Should -Be ""

        $ui.OpenModeCombo.SelectedItem = $null  # 選んでいなければ通常
        updateOpenMenu
        $ui.MenuOpen.InputGestureText | Should -Be "Enter"
        $ui.MenuOpenReadOnly.InputGestureText | Should -Be ""
        $ui.MenuOpenNew.InputGestureText | Should -Be ""

        $ui.OpenModeCombo.SelectedItem = $items[2]
        updateOpenMenu
        $ui.MenuOpen.InputGestureText | Should -Be ""
        $ui.MenuOpenNew.InputGestureText | Should -Be "Enter"
    }
}

Describe "openSource" -Tag Unit {
    BeforeEach {
        $script:statuses = New-Object System.Collections.Generic.List[string]
        $ui.OpenModeCombo.SelectedItem = $null
    }

    It "元のファイルが見つからなければ開かない" {
        Mock getCurrentHitRow { newRow }
        Mock findSourceFile { $null }
        Mock openInExcel { }
        Mock openWithShell { $true }

        openSource
        Should -Invoke openInExcel -Times 0 -Exactly
        Should -Invoke openWithShell -Times 0 -Exactly
    }

    It "Excel は該当シート（図形・コメントは元のシート）の該当セルを開く" {
        Mock getCurrentHitRow { newRow }
        Mock findSourceFile { "C:\data\見積.xlsx" }
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
        Mock findSourceFile { "C:\data\見積.xlsx" }
        Mock openInExcel { throw "ダイアログを表示中" }
        Mock openWithShell { $true }

        openSource ${openModeNormal}
        lastStatus | Should -Be "開きました（該当セルへの移動はできませんでした）：C:\data\見積.xlsx"
    }

    It "Excel を操作できず、指定の開き方でも開けなければ、元のファイルを開いたと知らせる" {
        Mock getCurrentHitRow { newRow }
        Mock findSourceFile { "C:\data\見積.xlsx" }
        Mock openInExcel { throw "ダイアログを表示中" }
        Mock openWithShell { $false }

        openSource ${openModeReadOnly}
        lastStatus | Should -Be "読み取り専用で開けなかったため、元のファイルを開きました（該当セルへの移動はできませんでした）：C:\data\見積.xlsx"
    }

    It "Excel 以外は既定のアプリで開く（開き方は［開き方］の選択）" {
        Mock getCurrentHitRow { newRow "報告.docx" $false }
        Mock findSourceFile { "C:\data\報告.docx" }
        Mock openWithShell { $true }
        $ui.OpenModeCombo.SelectedItem = [pscustomobject]@{ Tag = ${openModeNew} }

        openSource
        Should -Invoke openWithShell -Times 1 -Exactly -ParameterFilter { $mode -eq ${openModeNew} }
        lastStatus | Should -Be "新規で開きました：C:\data\報告.docx"
    }

    It "指定の開き方で開けなければ、元のファイルを開いたと知らせる" {
        Mock getCurrentHitRow { newRow "報告.docx" $false }
        Mock findSourceFile { "C:\data\報告.docx" }
        Mock openWithShell { $false }

        openSource ${openModeNew}
        lastStatus | Should -Be "新規で開けなかったため、元のファイルを開きました：C:\data\報告.docx"
    }
}

Describe "openSourceFolder" -Tag Unit {
    It "行を選んでいない・元のファイルが見つからなければ開かない" {
        Mock Start-Process { }
        Mock getCurrentHitRow { $null }
        openSourceFolder
        Mock getCurrentHitRow { newRow }
        Mock findSourceFile { $null }
        openSourceFolder
        Should -Invoke Start-Process -Times 0 -Exactly
    }

    It "エクスプローラーで元のファイルを選んだ状態で開く" {
        Mock Start-Process { }
        Mock getCurrentHitRow { newRow }
        Mock findSourceFile { "C:\data\見積.xlsx" }

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

    It "表示中の行を検索結果.txt に書き出して開く" {
        $workDir = "$TestDrive\export\work"
        $resultFile = "$workDir\検索結果.txt"
        $workspace = newTestWorkspace @{} $workDir
        $script:lastSearch = @{ Word = "りんご" }
        $script:hitCount = 2
        Mock getViewRows { , $hits }
        Mock Invoke-Item { }

        exportResults
        $lines = [System.IO.File]::ReadAllLines($resultFile)
        $lines[0] | Should -Be "【検索文字列　りんご】 2 件"
        $lines[2] | Should -Be "sub\見積.xlsx`t[シート] Sheet1`tセル`t3`tりんご`t100"
        Should -Invoke Invoke-Item -Times 1 -Exactly -ParameterFilter { $LiteralPath -eq $resultFile }
        lastStatus | Should -Be "検索結果.txt に出力しました（2 件）"
    }

    It "前回の検索結果.txt は追記せずに置き換え、BOM 付き UTF-8 で書く（0 件でも書き出す）" {
        $workDir = "$TestDrive\export_empty\work"
        $resultFile = "$workDir\検索結果.txt"
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
        lastStatus | Should -Be "検索結果.txt に出力しました（0 件）"
    }

    It "絞り込んでいれば、絞り込み後の件数を知らせる" {
        $workDir = "$TestDrive\export_filtered\work"
        $resultFile = "$workDir\検索結果.txt"
        $workspace = newTestWorkspace @{} $workDir
        $script:lastSearch = @{ Word = "りんご" }
        $script:hitCount = 1234
        Mock getViewRows { , $hits }
        Mock Invoke-Item { }

        exportResults
        lastStatus | Should -Be "絞り込み後の 2 件を検索結果.txt に出力しました"
    }

    It "検索結果.txt をほかのアプリが開いていれば、閉じるよう知らせる" {
        $workDir = "$TestDrive\export_locked\work"
        $resultFile = "$workDir\検索結果.txt"
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
        lastStatus | Should -Be "検索結果.txt に書き込めません。開いているアプリを閉じてから、もう一度出力してください。"
        Should -Invoke Invoke-Item -Times 0 -Exactly
    }

    It "検索結果.txt が読み取り専用なら、予期しないエラーにせず、読み取り専用を外すよう知らせる" {
        $workDir = "$TestDrive\export_readonly\work"
        $resultFile = "$workDir\検索結果.txt"
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
        lastStatus | Should -Be "検索結果.txt に書き込む権限がありません（読み取り専用など）。$resultFile を確かめてから、もう一度出力してください。"
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

    It "［開き方］を選ぶと設定に保存する（起動時の読み込みでは保存しない）" {
        Mock writeOpenMode { }
        $ui.OpenModeCombo.SelectedItem = [pscustomobject]@{ Tag = ${openModeReadOnly} }

        $script:loadingOpenMode = $true
        & $ui.OpenModeCombo.Handlers["SelectionChanged"]
        Should -Invoke writeOpenMode -Times 0 -Exactly

        $script:loadingOpenMode = $false
        & $ui.OpenModeCombo.Handlers["SelectionChanged"]
        Should -Invoke writeOpenMode -Times 1 -Exactly -ParameterFilter { $mode -eq ${openModeReadOnly} }
        $ui.MenuOpenReadOnly.InputGestureText | Should -Be "Enter"
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
