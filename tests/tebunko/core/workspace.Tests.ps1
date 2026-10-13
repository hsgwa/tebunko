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
        $target.SystemIndexStateFile | Should -Be "C:\Users\test\Documents\tebunko_ws\system_index_state.tsv"
        $target.PublishDir | Should -Be "C:\Users\test\Documents\tebunko_ws\publish\$PID"
        $target.StatusFile | Should -Be "C:\Users\test\Documents\tebunko_ws\ingest_status.tsv"
        $target.IngestingFile | Should -Be "C:\Users\test\Documents\tebunko_ws\ingesting.txt"
        $target.ResultFile | Should -Be "C:\Users\test\Documents\tebunko_ws\search_results.txt"
        $target.IndexingLogFile | Should -Be "C:\Users\test\Documents\tebunko_ws\indexing_log.txt"
        $target.GuiErrorLogFile | Should -Be "C:\Users\test\Documents\tebunko_ws\gui_error_log.txt"
        $target.IndexingLogPreviousFile | Should -Be "C:\Users\test\Documents\tebunko_ws\indexing_log_previous.txt"
        $target.GuiErrorLogPreviousFile | Should -Be "C:\Users\test\Documents\tebunko_ws\gui_error_log_previous.txt"
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
        @($target.Entries()) | Should -Contain $target.OfficePidRoot
        @($target.Entries()) | Should -Contain $target.IndexingLogPreviousFile
        @($target.Entries()) | Should -Contain $target.GuiErrorLogPreviousFile
        # 前回の分の名前は getPreviousLogPath の結果と同じ
        $target.IndexingLogPreviousFile | Should -Be (getPreviousLogPath $target.IndexingLogFile)
        $target.GuiErrorLogPreviousFile | Should -Be (getPreviousLogPath $target.GuiErrorLogFile)
    }

    It "文字列のプロパティ（Dir・Legacy で始まるものを除く）は、Dir より後ろが ASCII（再発防止。後から足したプロパティも捕まえる）" {
        $target = [Workspace]::new("C:\Users\test\Documents\tebunko_ws")
        $names = $target.PSObject.Properties |
            Where-Object { $_.TypeNameOfValue -eq "System.String" -and $_.Name -ne "Dir" -and $_.Name -notlike "Legacy*" } |
            ForEach-Object { $_.Name }

        $names.Count | Should -BeGreaterThan 0
        foreach ($name in $names) {
            $value = $target.$name
            $rest = $value.Substring($target.Dir.Length)
            $rest | Should -Match "^[\x20-\x7E]+$"
        }

        ${sourceFolderFileName} | Should -Match "^[\x20-\x7E]+$"
    }
}

Describe "getMachineKey" -Tag Unit {
    It "この PC の鍵（getFolderKey の先頭 8 文字）を返す" {
        getMachineKey | Should -Be (getFolderKey ([Environment]::MachineName)).Substring(0, 8)
    }
}

Describe "getOfficePidDir" -Tag Unit {
    It "ワークスペースの office_pids の下に、PC の鍵で組み立てる（文字列だけで、フォルダは作らない）" {
        $workspace = newTestWorkspace @{} "$TestDrive\getopd"

        getOfficePidDir $workspace | Should -Be "$TestDrive\getopd\office_pids\$(getMachineKey)"
        Test-Path "$TestDrive\getopd" | Should -Be $false
    }
}

Describe "assertWorkspaceReachable" -Tag Unit {
    It "<label>ときの判定" -TestCases @(
        @{ label = "見つかった"; state = "Found"; throws = $false }
        @{ label = "まだ無い（インデックスを作っていない）"; state = "Missing"; throws = $false }
        @{ label = "届かない"; state = "Unreachable"; throws = $true }
        @{ label = "確かめられない（アクセス拒否など）"; state = "Other"; throws = $true }
    ) {
        param ($label, $state, $throws)
        $script:fakeState = switch ($state) {
            "Found" { ${pathStateFound} }
            "Missing" { ${pathStateMissing} }
            "Unreachable" { ${pathStateUnreachable} }
            default { ${pathStateOther} }
        }
        Mock getPathState { @{ State = $script:fakeState; IsDirectory = $true; Message = "理由" } }

        if ($throws) {
            { assertWorkspaceReachable "\\fileserver\共有\ws" } | Should -Throw "*ワークスペース*"
        } else {
            { assertWorkspaceReachable "\\fileserver\共有\ws" } | Should -Not -Throw
        }
    }

    It "届かないときの文面に、場所を入れる" {
        $script:fakeState = ${pathStateUnreachable}
        Mock getPathState { @{ State = $script:fakeState; IsDirectory = $false; Message = "理由" } }

        { assertWorkspaceReachable "\\fileserver\共有\ws" } | Should -Throw "ワークスペースに接続できません：\\fileserver\共有\ws"
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
        @{ name = "候補のパスに [ と ] の両方があれば、置けない（Dir は空）"; kind = "path"; dir = "C:\Users\test\[共有]フォルダ"; expectedReason = "Brackets" }
        @{ name = "候補のパスに [ だけあっても、置けない"; kind = "path"; dir = "C:\Users\test\[共有フォルダ"; expectedReason = "Brackets" }
        @{ name = "候補のパスに ] だけあっても、置けない"; kind = "path"; dir = "C:\Users\test\共有]フォルダ"; expectedReason = "Brackets" }
        @{ name = "候補の長さが境目より短ければ、そのまま使う"; kind = "length"; offset = -1; expectedReason = "" }
        @{ name = "候補の長さが境目以上なら、置けない（Dir は空）"; kind = "length"; offset = 0; expectedReason = "TooLong" }
    ) {
        param ($name, $kind, $dir, $offset, $expectedReason)
        $workspace = if ($kind -eq "length") {
            newWorkspaceOfCandidateLength ($thresholdLen + $offset)
        } else {
            newTestWorkspace @{} $dir
        }

        $result = selectTmpDir $workspace
        if ($expectedReason) {
            $result.Dir | Should -BeNullOrEmpty
        } else {
            $result.Dir | Should -Be (getWorkspaceTmpDir $workspace)
        }
        $result.Reason | Should -Be $expectedReason
    }
}

Describe "getTmpDirUnavailableMessage" -Tag Unit {
    It "<reason> なら [ ]・長さに触れた文にする" -TestCases @(
        @{ reason = "Brackets"; expected = "\[ \]" }
        @{ reason = "TooLong"; expected = "長すぎる" }
    ) {
        param ($reason, $expected)
        getTmpDirUnavailableMessage $reason | Should -Match $expected
    }

    It "置けない理由を書く" {
        getTmpDirUnavailableMessage "Brackets" | Should -Match "置けません"
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
        $message | Should -Match "更新し直します"
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

Describe "getWorkspaceJobQueue" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "ローカルのワークスペースは既定の列"; dir = "C:\Users\test\Documents\tebunko_ws"; expected = "default" }
        @{ name = "UNC のワークスペースは専用の列"; dir = "\\fileserver\共有\tebunko_ws"; expected = "network" }
        @{ name = "記録の置き場所（ローカル）は既定の列"; dir = "C:\Users\test\Documents\tebunko_ws\office_pids\ab12cd34"; expected = "default" }
        @{ name = "記録の置き場所（UNC）は専用の列"; dir = "\\server\share\ws\office_pids\ab12cd34"; expected = "network" }
        @{ name = "\\?\UNC\ も専用の列"; dir = "\\?\UNC\server\share\ws\office_pids\ab12cd34"; expected = "network" }
    ) {
        getWorkspaceJobQueue $dir | Should -Be $expected
    }

    It "複数の場所のうち 1 つでもネットワークなら専用の列、すべてローカルなら既定の列、空も既定の列" {
        getWorkspaceJobQueue @("C:\Users\test\a", "\\server\share\b") | Should -Be "network"
        getWorkspaceJobQueue @("C:\Users\test\a", "D:\b") | Should -Be "default"
        getWorkspaceJobQueue @() | Should -Be "default"
    }
}

Describe "setCallerWorkspace" -Tag Unit {
    It "呼んだスクリプトのスコープに `$workspace を置き、呼び出しの外の `$workspace は変えない" {
        $before = $script:workspace
        $own = [Workspace]::new("C:\Users\test\Documents\tebunko_own")
        $seen = & {
            setCallerWorkspace $own
            $workspace.Dir
        }
        $seen | Should -Be "C:\Users\test\Documents\tebunko_own"
        [object]::ReferenceEquals($script:workspace, $before) | Should -Be $true
    }
}
