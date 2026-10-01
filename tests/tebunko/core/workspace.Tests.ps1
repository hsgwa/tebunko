# ワークスペースの中の場所（tebunko\core\workspace.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "Workspace" -Tag Unit {
    It "中の場所をワークスペースのフォルダから組み立てる（出力用のフォルダはプロセスごと）" {
        $target = [Workspace]::new("C:\Users\test\Documents\tebunko_ws")

        $target.Dir | Should -Be "C:\Users\test\Documents\tebunko_ws"
        $target.IndexDir | Should -Be "C:\Users\test\Documents\tebunko_ws\content_index"
        $target.LegacyIndexDir | Should -Be "C:\Users\test\Documents\tebunko_ws\index"
        $target.SystemIndexDir | Should -Be "C:\Users\test\Documents\tebunko_ws\system_index"
        $target.SystemIndexStateFile | Should -Be "C:\Users\test\Documents\tebunko_ws\システムインデックスの状態.tsv"
        $target.PublishDir | Should -Be "C:\Users\test\Documents\tebunko_ws\取り込み出力\$PID"
        $target.StatusFile | Should -Be "C:\Users\test\Documents\tebunko_ws\取り込み一覧.tsv"
        $target.IngestingFile | Should -Be "C:\Users\test\Documents\tebunko_ws\取り込み中.txt"
        $target.ResultFile | Should -Be "C:\Users\test\Documents\tebunko_ws\検索結果.txt"
        $target.IndexingLogFile | Should -Be "C:\Users\test\Documents\tebunko_ws\インデックス作成ログ.txt"
        $target.GuiErrorLogFile | Should -Be "C:\Users\test\Documents\tebunko_ws\画面エラー.txt"
        $target.TmpRoot | Should -Be "C:\Users\test\Documents\tebunko_ws\tmp"
    }

    It "関数は呼んだときの `$workspace の場所を使う（差し替えれば、読み込み直さずに別のワークスペースを使う）" {
        $workspace = newTestWorkspace @{} "$TestDrive\別"

        @((getIndexSummary).Missing) | Should -Be @("$TestDrive\別\content_index")
    }

    It "Entries() が前の版の index（LegacyIndexDir）と一時フォルダ（TmpRoot）を含む" {
        $target = [Workspace]::new("$TestDrive\entries")

        @($target.Entries()) | Should -Contain $target.LegacyIndexDir
        @($target.Entries()) | Should -Contain $target.IndexDir
        @($target.Entries()) | Should -Contain $target.TmpRoot
    }
}

Describe "getMachineKey" -Tag Unit {
    It "この PC の鍵（getFolderKey の先頭 8 文字）を返す" {
        getMachineKey | Should -Be (getFolderKey ([Environment]::MachineName)).Substring(0, 8)
    }
}

Describe "getWorkspaceTmpDir" -Tag Unit {
    It "ワークスペースの tmp の下に、PC の鍵とプロセスIDで組み立てる" {
        $workspace = newTestWorkspace @{} "$TestDrive\getwtd"

        getWorkspaceTmpDir $workspace | Should -Be "$TestDrive\getwtd\tmp\$(getMachineKey)\$PID"
    }
}

Describe "selectTmpDir" -Tag Unit {
    BeforeAll {
        # 候補（ワークスペースの tmp の下）の長さと、取り込みのスレッドが下に作る最も長い名前の分（tmpNameReserve）から、
        # ちょうど境目の長さのワークスペースを組み立てるための下ごしらえ
        $suffix = "\tmp\$(getMachineKey)\$PID"
        $thresholdLen = $excelMaxPath - ${tmpNameReserve}

        function newWorkspaceOfCandidateLength([int]$candidateLen) {
            $dirLen = $candidateLen - $suffix.Length
            $dir = "C:\" + ("あ" * ($dirLen - 3))
            return (newTestWorkspace @{} $dir)
        }
    }

    It "<name>" -TestCases @(
        @{ name = "[ ] を含まず、長さも十分短ければ、そのまま使う"; kind = "path"; dir = "C:\Users\test\ws"; expectedReason = "" }
        @{ name = "候補のパスに [ と ] の両方があれば、前の版の場所（%TEMP%\tebunko\<PID>）を使う"; kind = "path"; dir = "C:\Users\test\[共有]フォルダ"; expectedReason = "Brackets" }
        @{ name = "候補のパスに [ だけあっても、前の版の場所を使う"; kind = "path"; dir = "C:\Users\test\[共有フォルダ"; expectedReason = "Brackets" }
        @{ name = "候補のパスに ] だけあっても、前の版の場所を使う"; kind = "path"; dir = "C:\Users\test\共有]フォルダ"; expectedReason = "Brackets" }
        @{ name = "候補の長さが境目より短ければ、そのまま使う"; kind = "length"; offset = -1; expectedReason = "" }
        @{ name = "候補の長さが境目以上なら、前の版の場所（%TEMP%\tebunko\<PID>）を使う"; kind = "length"; offset = 0; expectedReason = "TooLong" }
    ) {
        param ($name, $kind, $dir, $offset, $expectedReason)
        $workspace = if ($kind -eq "length") {
            newWorkspaceOfCandidateLength ($thresholdLen + $offset)
        } else {
            newTestWorkspace @{} $dir
        }

        $result = selectTmpDir $workspace
        if ($expectedReason) {
            $result.Dir | Should -Be (Join-Path ${legacyTmpParent} $PID)
        } else {
            $result.Dir | Should -Be (getWorkspaceTmpDir $workspace)
        }
        $result.Reason | Should -Be $expectedReason
    }
}

Describe "moveWorkspace" -Tag Io {
    BeforeAll {
        function newWorkspaceFiles([string]$dir) {
            # tebunko のファイル・フォルダと、利用者のファイル（移さない）を置く。
            # content_index（今の版）と、前の版の残骸の index の両方を置いて、どちらも移す・消すことを確かめる
            newTsv "$dir\content_index\営業\見積\content_index.xlsx.001.tsv" @("a")
            newTsv "$dir\index\営業\見積\content.xlsx.001.tsv" @("a-前の版")
            newTsv "$dir\system_index\営業\見積\system_index.txt" @("b")
            newTsv "$dir\取り込み一覧.tsv" @("c")
            newTsv "$dir\インデックス作成ログ.txt" @("d")
            newTsv "$dir\利用者のメモ.txt" @("e")
        }
    }

    It "tebunko のファイル・フォルダだけを移し、移した数を返す（利用者のファイルは残す）" {
        newWorkspaceFiles "$TestDrive\move\from"

        moveWorkspace "$TestDrive\move\from" "$TestDrive\move\to" | Should -Be 5

        Test-Path -LiteralPath "$TestDrive\move\to\content_index\営業\見積\content_index.xlsx.001.tsv" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\move\to\index\営業\見積\content.xlsx.001.tsv" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\move\to\system_index\営業\見積\system_index.txt" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\move\to\取り込み一覧.tsv" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\move\to\インデックス作成ログ.txt" | Should -Be $true
        @(getWorkspaceEntries "$TestDrive\move\from").Count | Should -Be 0
        Test-Path -LiteralPath "$TestDrive\move\from\利用者のメモ.txt" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\move\to\利用者のメモ.txt" | Should -Be $false
    }

    It "移し先に同じ名前があれば、何も移さずに例外にする" {
        newWorkspaceFiles "$TestDrive\conflict\from"
        newTsv "$TestDrive\conflict\to\取り込み一覧.tsv" @("別のワークスペース")

        { moveWorkspace "$TestDrive\conflict\from" "$TestDrive\conflict\to" } | Should -Throw "*すでに 取り込み一覧.tsv があります*"
        @(getWorkspaceEntries "$TestDrive\conflict\from").Count | Should -Be 5
        Test-Path -LiteralPath "$TestDrive\conflict\to\content_index" | Should -Be $false
    }

    It "途中で移せなければ、移した分を戻して例外にする（中身が 2 つに分かれない）" {
        newWorkspaceFiles "$TestDrive\rollback\from"
        # system_index の中のファイルを開いておき、フォルダを移せなくする（content_index・index は先に移る）
        $stream = [System.IO.File]::Open("$TestDrive\rollback\from\system_index\営業\見積\system_index.txt", "Open", "Read", "None")
        try {
            { moveWorkspace "$TestDrive\rollback\from" "$TestDrive\rollback\to" } | Should -Throw "*中身は「$TestDrive\rollback\from」に残しています*"
        } finally {
            $stream.Dispose()
        }
        @(getWorkspaceEntries "$TestDrive\rollback\from").Count | Should -Be 5
        @(getWorkspaceEntries "$TestDrive\rollback\to").Count | Should -Be 0
    }

    It "中身が無ければ何も移さない" {
        moveWorkspace "$TestDrive\empty\from" "$TestDrive\empty\to" | Should -Be 0
    }
}

Describe "copyDirectoryTree" -Tag Io {
    It "フォルダを中身ごと写す（別のドライブへ移すとき）" {
        newTsv "$TestDrive\copy\from\a\b.tsv" @("b")
        newTsv "$TestDrive\copy\from\c.txt" @("c")

        copyDirectoryTree (toLongPath "$TestDrive\copy\from") (toLongPath "$TestDrive\copy\to")

        Test-Path -LiteralPath "$TestDrive\copy\to\a\b.tsv" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\copy\to\c.txt" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\copy\from\c.txt" | Should -Be $true
    }
}

Describe "moveSearchExcludes" -Tag Io {
    It "前のワークスペースのインデックスの下を、移した先のインデックスの下に付け替える（ほかは変えない）" {
        $settings = "$TestDrive\excludes\setting.config"
        writeSearchExcludes @(
            [pscustomobject]@{ Path = "C:\old\content_index"; Subfolders = $true },
            [pscustomobject]@{ Path = "C:\OLD\content_index\営業\見積"; Subfolders = $false },
            [pscustomobject]@{ Path = "C:\old\content_index2\技術"; Subfolders = $true }
        ) $settings

        moveSearchExcludes "C:\old" "D:\新しい場所" $settings | Should -Be 2

        $excludes = @(readSearchExcludes $settings)
        @($excludes | ForEach-Object { $_.Path }) -join "|" | Should -Be "D:\新しい場所\content_index|D:\新しい場所\content_index\営業\見積|C:\old\content_index2\技術"
        @($excludes | ForEach-Object { $_.Subfolders }) -join "," | Should -Be "True,False,True"
    }
}

Describe "removeWorkspaceEntries" -Tag Io {
    It "tebunko のファイル・フォルダだけを削除し、削除した数を返す（content_index・前の版の index も含めて消し、利用者のファイルは残す）" {
        newTsv "$TestDrive\remove_ws\content_index\営業\見積\content_index.xlsx.001.tsv" @("a")
        newTsv "$TestDrive\remove_ws\index\営業\見積\content.xlsx.001.tsv" @("a-前の版")
        newTsv "$TestDrive\remove_ws\取り込み一覧.tsv" @("b")
        newTsv "$TestDrive\remove_ws\利用者のメモ.txt" @("c")

        removeWorkspaceEntries "$TestDrive\remove_ws" | Should -Be 3

        @(getWorkspaceEntries "$TestDrive\remove_ws").Count | Should -Be 0
        Test-Path -LiteralPath "$TestDrive\remove_ws\content_index" | Should -Be $false
        Test-Path -LiteralPath "$TestDrive\remove_ws\index" | Should -Be $false
        Test-Path -LiteralPath "$TestDrive\remove_ws\利用者のメモ.txt" | Should -Be $true
    }
}

Describe "useWorkspaceTargets" -Tag Io {
    It "ワークスペースの取り込み一覧にあるクロール対象フォルダを、インデックスの一覧にする" {
        $settings = "$TestDrive\use\setting.config"
        writeTargetFolders @([pscustomobject]@{ Name = "自分"; Path = "C:\自分のフォルダ"; Enabled = $true }) $settings
        writeStatusFile @(
            [pscustomobject]@{ Path = "\\server\共有\営業部"; Name = "営業" },
            [pscustomobject]@{ Path = "\\server\共有\技術部"; Name = "技術" }
        ) @() "$TestDrive\use\ws\取り込み一覧.tsv"

        useWorkspaceTargets "$TestDrive\use\ws" $settings | Should -Be 2

        $targets = @(getTargetFolders $settings)
        @($targets | ForEach-Object { "$($_.Name)=$($_.Path)=$($_.Enabled)" }) -join "|" | Should -Be "営業=\\server\共有\営業部=True|技術=\\server\共有\技術部=True"
    }

    It "取り込み一覧が無ければ、一覧は変えない" {
        $settings = "$TestDrive\use_none\setting.config"
        writeTargetFolders @([pscustomobject]@{ Name = "自分"; Path = "C:\自分のフォルダ"; Enabled = $true }) $settings

        useWorkspaceTargets "$TestDrive\use_none\ws" $settings | Should -Be 0

        @(getTargetFolders $settings)[0].Name | Should -Be "自分"
    }
}

Describe "getLegacyIndexState" -Tag Io {
    It "<name>" -TestCases @(
        @{ name = "index が無い"; hasDir = $false; hasMark = $false; content = "none"; hasTxt = $false
           hasLegacy = $false; contentEmpty = $true; hasLegacyTxt = $false }
        @{ name = "index はあるがしるしが無い"; hasDir = $true; hasMark = $false; content = "none"; hasTxt = $false
           hasLegacy = $false; contentEmpty = $true; hasLegacyTxt = $false }
        @{ name = "しるしあり・content_index 無し・前の名前の txt 無し"; hasDir = $true; hasMark = $true; content = "none"; hasTxt = $false
           hasLegacy = $true; contentEmpty = $true; hasLegacyTxt = $false }
        @{ name = "しるしあり・content_index 無し・前の名前の txt あり"; hasDir = $true; hasMark = $true; content = "none"; hasTxt = $true
           hasLegacy = $true; contentEmpty = $true; hasLegacyTxt = $true }
        @{ name = "しるしあり・content_index は空のフォルダだけ・txt 無し"; hasDir = $true; hasMark = $true; content = "emptyFolder"; hasTxt = $false
           hasLegacy = $true; contentEmpty = $true; hasLegacyTxt = $false }
        @{ name = "しるしあり・content_index は空のフォルダだけ・txt あり"; hasDir = $true; hasMark = $true; content = "emptyFolder"; hasTxt = $true
           hasLegacy = $true; contentEmpty = $true; hasLegacyTxt = $true }
        @{ name = "しるしあり・content_index にファイルがある・txt 無し"; hasDir = $true; hasMark = $true; content = "hasFile"; hasTxt = $false
           hasLegacy = $true; contentEmpty = $false; hasLegacyTxt = $false }
        @{ name = "しるしあり・content_index にファイルがある・txt あり（空でないので調べない）"; hasDir = $true; hasMark = $true; content = "hasFile"; hasTxt = $true
           hasLegacy = $true; contentEmpty = $false; hasLegacyTxt = $false }
    ) {
        param ($name, $hasDir, $hasMark, $content, $hasTxt, $hasLegacy, $contentEmpty, $hasLegacyTxt)
        $dir = "$TestDrive\legacy_$([Guid]::NewGuid().ToString('N'))"
        if ($hasDir) {
            [System.IO.Directory]::CreateDirectory("$dir\index\営業") | Out-Null
            if ($hasMark) {
                newTsv "$dir\index\営業\元のフォルダ.txt" @("営業`tC:\元")
            }
        }
        switch ($content) {
            "emptyFolder" { [System.IO.Directory]::CreateDirectory("$dir\content_index\営業") | Out-Null }
            "hasFile"     { newTsv "$dir\content_index\営業\content_index.xlsx.001.tsv" @("a") }
        }
        if ($hasTxt) {
            newTsv "$dir\system_index\営業\システムインデックス.txt" @("x00000000")
        }

        $state = getLegacyIndexState $dir
        $state.HasLegacyIndex | Should -Be $hasLegacy
        $state.ContentEmpty | Should -Be $contentEmpty
        $state.HasLegacySystemIndex | Should -Be $hasLegacyTxt
    }

    It "調べ終えたら、下のフォルダを掴んだままにしない（すぐに移動できる）" {
        # 見つけたところで列挙をやめても、列挙子が下のフォルダを開いたまま残らないこと（残るとインポートの上書きで移動できない）
        $dir = "$TestDrive\legacy_handle"
        [System.IO.Directory]::CreateDirectory("$dir\index\営業") | Out-Null
        newTsv "$dir\index\営業\元のフォルダ.txt" @("営業`tC:\元")
        newTsv "$dir\content_index\営業\見積\content_index.xlsx.001.tsv" @("a")
        newTsv "$dir\content_index\営業\見積\content_index.xlsx.002.tsv" @("b")
        newTsv "$dir\system_index\営業\システムインデックス.txt" @("x00000000")

        $state = getLegacyIndexState $dir
        $state.ContentEmpty | Should -Be $false

        { [System.IO.Directory]::Move("$dir\content_index\営業", "$dir\moved_content") } | Should -Not -Throw
        { [System.IO.Directory]::Move("$dir\index\営業", "$dir\moved_index") } | Should -Not -Throw

        # content_index が空のときに調べる system_index も同じ
        $empty = "$TestDrive\legacy_handle_empty"
        newTsv "$empty\system_index\営業\見積\システムインデックス.txt" @("x00000000")
        newTsv "$empty\system_index\営業\見積\システムインデックス2.txt" @("x00000001")

        (getLegacyIndexState $empty).HasLegacySystemIndex | Should -Be $true

        { [System.IO.Directory]::Move("$empty\system_index\営業", "$empty\moved_system") } | Should -Not -Throw
    }
}

Describe "testLegacyCleanupNeeded" -Tag Unit {
    It "content_index が空で、しるしか前の名前の txt があるときだけ片付ける" -TestCases @(
        @{ contentEmpty = $true;  hasLegacy = $true;  hasTxt = $false; expected = $true }
        @{ contentEmpty = $true;  hasLegacy = $false; hasTxt = $true;  expected = $true }
        @{ contentEmpty = $true;  hasLegacy = $false; hasTxt = $false; expected = $false }
        @{ contentEmpty = $false; hasLegacy = $true;  hasTxt = $true;  expected = $false }
    ) {
        param ($contentEmpty, $hasLegacy, $hasTxt, $expected)
        testLegacyCleanupNeeded @{ ContentEmpty = $contentEmpty; HasLegacyIndex = $hasLegacy; HasLegacySystemIndex = $hasTxt } | Should -Be $expected
    }
}

Describe "getLegacyIndexMessage" -Tag Unit {
    It "しるしがあれば文言を、無ければ空を返す（既定・既定でない場所で同じ）" -TestCases @(
        @{ dir = "C:\Users\test\Documents\tebunko_ws" }
        @{ dir = "D:\ほかの場所\ws" }
    ) {
        param ($dir)
        getLegacyIndexMessage $dir $false | Should -Be ""
        $message = getLegacyIndexMessage $dir $true
        $message | Should -Match ([regex]::Escape("$dir\index"))
        $message | Should -Match "取り込み直します"
    }
}

Describe "clearLegacySystemIndex" -Tag Io {
    It "system_index の txt を開いたままだと Ok=`$false` を理由付きで返すが、状態ファイルの中身は空にする。閉じてもう一度で Ok=`$true` になる" {
        $dir = "$TestDrive\clear1"
        $ws = [Workspace]::new($dir)
        $txt = "$($ws.SystemIndexDir)\営業\システムインデックス.txt"
        newTsv $txt @("x00000000")
        [void](updateSystemIndexState { param ($s) [void]$s.Covered.Add("営業") } $ws.SystemIndexStateFile)

        $stream = [System.IO.File]::Open($txt, "Open", "Read", "None")
        try {
            $result = clearLegacySystemIndex $dir
            $result.Ok | Should -Be $false
            $result.Reason | Should -Not -BeNullOrEmpty
        } finally {
            $stream.Dispose()
        }
        (readSystemIndexState $ws.SystemIndexStateFile).Covered.Count | Should -Be 0
        Test-Path -LiteralPath $txt | Should -Be $true

        $result = clearLegacySystemIndex $dir
        $result.Ok | Should -Be $true
        $result.Reason | Should -Be ""
        Test-Path -LiteralPath $ws.SystemIndexDir | Should -Be $false
    }

    It "状態ファイルを開けなければ Ok=`$false` を理由付きで返し、system_index には触らない" {
        $dir = "$TestDrive\clear2"
        $ws = [Workspace]::new($dir)
        newTsv "$($ws.SystemIndexDir)\営業\システムインデックス.txt" @("x00000000")
        [void](updateSystemIndexState { param ($s) } $ws.SystemIndexStateFile)
        $stream = [System.IO.FileStream]::new($ws.SystemIndexStateFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        try {
            $result = clearLegacySystemIndex $dir
            $result.Ok | Should -Be $false
            $result.Reason | Should -Not -BeNullOrEmpty
        } finally {
            $stream.Dispose()
        }
        Test-Path -LiteralPath $ws.SystemIndexDir | Should -Be $true
    }
}
