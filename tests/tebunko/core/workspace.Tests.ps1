# ワークスペースの中の場所（tebunko\core\workspace.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "Workspace" -Tag Unit {
    It "中の場所をワークスペースのフォルダから組み立てる（出力用のフォルダはプロセスごと）" {
        $target = [Workspace]::new("C:\Users\test\Documents\tebunko_ws")

        $target.Dir | Should -Be "C:\Users\test\Documents\tebunko_ws"
        $target.IndexDir | Should -Be "C:\Users\test\Documents\tebunko_ws\store\content"
        $target.StoreDir | Should -Be "C:\Users\test\Documents\tebunko_ws\store"
        $target.FormatFile | Should -Be "C:\Users\test\Documents\tebunko_ws\store\format.tsv"
        $target.OldIndexDir | Should -Be "C:\Users\test\Documents\tebunko_ws\content_index"
        $target.OldStatusFile | Should -Be "C:\Users\test\Documents\tebunko_ws\ingest_status.tsv"
        $target.LegacyIndexDir | Should -Be "C:\Users\test\Documents\tebunko_ws\index"
        $target.OldSystemIndexDir | Should -Be "C:\Users\test\Documents\tebunko_ws\system_index"
        $target.OldSystemIndexStateFile | Should -Be "C:\Users\test\Documents\tebunko_ws\system_index_state.tsv"
        $target.PublishDir | Should -Be "C:\Users\test\Documents\tebunko_ws\publish\$PID"
        $target.StatusFile | Should -Be "C:\Users\test\Documents\tebunko_ws\store\ingest_status.tsv"
        $target.JournalFile | Should -Be "C:\Users\test\Documents\tebunko_ws\store\journal\ingest.tsv"
        $target.PendingRemoveFile | Should -Be "C:\Users\test\Documents\tebunko_ws\store\pending_remove.tsv"
        $target.ResultFile | Should -Be "C:\Users\test\Documents\tebunko_ws\search_results.txt"
        $target.IndexingLogFile | Should -Be "C:\Users\test\Documents\tebunko_ws\indexing_log.txt"
        $target.LockFile | Should -Be "C:\Users\test\Documents\tebunko_ws\indexing.lock"
        $target.GuiErrorLogFile | Should -Be "C:\Users\test\Documents\tebunko_ws\gui_error_log.txt"
        $target.IndexingLogPreviousFile | Should -Be "C:\Users\test\Documents\tebunko_ws\indexing_log_previous.txt"
        $target.GuiErrorLogPreviousFile | Should -Be "C:\Users\test\Documents\tebunko_ws\gui_error_log_previous.txt"
        $target.TmpRoot | Should -Be "C:\Users\test\Documents\tebunko_ws\tmp"
    }

    It "鍵のファイルの場所は、ドライブの直下のワークスペースでも区切りが二重にならない（長いパスの形にしても開ける）" -TestCases @(
        @{ dir = "Y:\"; expected = "Y:\indexing.lock" }
        @{ dir = "Y:\ws"; expected = "Y:\ws\indexing.lock" }
    ) {
        param ($dir, $expected)
        ([Workspace]::new($dir)).LockFile | Should -Be $expected
    }

    It "関数は呼んだときの `$workspace の場所を使う（差し替えれば、読み込み直さずに別のワークスペースを使う）" {
        $workspace = newTestWorkspace @{} "$TestDrive\別"

        @((getIndexSummary).Missing) | Should -Be @("$TestDrive\別\store\content")
    }

    It "Entries() が前の版の index（LegacyIndexDir）と一時フォルダ（TmpRoot）を含む" {
        $target = [Workspace]::new("$TestDrive\entries")

        @($target.Entries()) | Should -Contain $target.LegacyIndexDir
        @($target.Entries()) | Should -Contain $target.StoreDir
        @($target.Entries()) | Should -Contain $target.OldIndexDir
        @($target.Entries()) | Should -Contain $target.TmpRoot
        @($target.Entries()) | Should -Contain $target.OfficePidRoot
        @($target.Entries()) | Should -Contain $target.LockFile
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
    It "<reason> なら [ ]・長さ・リンクのうち、その理由に触れた文にする" -TestCases @(
        @{ reason = "Brackets"; expected = "\[ \]" }
        @{ reason = "TooLong"; expected = "長すぎる" }
        @{ reason = "Link"; expected = "リンク" }
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
        @{ name = "index が無い"; hasDir = $false; hasMark = $false; content = "none"
           hasLegacy = $false; contentEmpty = $true }
        @{ name = "index はあるがしるしが無い"; hasDir = $true; hasMark = $false; content = "none"
           hasLegacy = $false; contentEmpty = $true }
        @{ name = "しるしあり・content_index 無し"; hasDir = $true; hasMark = $true; content = "none"
           hasLegacy = $true; contentEmpty = $true }
        @{ name = "しるしあり・content_index は空のフォルダだけ"; hasDir = $true; hasMark = $true; content = "emptyFolder"
           hasLegacy = $true; contentEmpty = $true }
        @{ name = "しるしあり・content_index にファイルがある"; hasDir = $true; hasMark = $true; content = "hasFile"
           hasLegacy = $true; contentEmpty = $false }
    ) {
        param ($name, $hasDir, $hasMark, $content, $hasLegacy, $contentEmpty)
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

        $state = getLegacyIndexState $dir
        $state.HasLegacyIndex | Should -Be $hasLegacy
        $state.ContentEmpty | Should -Be $contentEmpty
    }

    It "調べ終えたら、下のフォルダを掴んだままにしない（すぐに移動できる）" {
        # 見つけたところで列挙をやめても、列挙子が下のフォルダを開いたまま残らないこと（残るとインポートの上書きで移動できない）
        $dir = "$TestDrive\legacy_handle"
        [System.IO.Directory]::CreateDirectory("$dir\index\営業") | Out-Null
        newTsv "$dir\index\営業\元のフォルダ.txt" @("営業`tC:\元")
        newTsv "$dir\content_index\営業\見積\content_index.xlsx.001.tsv" @("a")
        newTsv "$dir\content_index\営業\見積\content_index.xlsx.002.tsv" @("b")

        $state = getLegacyIndexState $dir
        $state.ContentEmpty | Should -Be $false

        { [System.IO.Directory]::Move("$dir\content_index\営業", "$dir\moved_content") } | Should -Not -Throw
        { [System.IO.Directory]::Move("$dir\index\営業", "$dir\moved_index") } | Should -Not -Throw
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

Describe "removeOldSystemIndex" -Tag Io {
    It "前の版のシステムインデックスと状態ファイルがあれば消し、無くても Ok を返す" {
        $dir = "$TestDrive\old1"
        $ws = [Workspace]::new($dir)
        newTsv "$($ws.OldSystemIndexDir)\営業\システムインデックス.txt" @("x00000000")
        newTsv $ws.OldSystemIndexStateFile @("営業")
        newTsv "$dir\store\content\営業\a.tsv" @("a")

        $result = removeOldSystemIndex $dir
        $result.Ok | Should -Be $true
        $result.Reason | Should -Be ""
        Test-Path -LiteralPath $ws.OldSystemIndexDir | Should -Be $false
        Test-Path -LiteralPath $ws.OldSystemIndexStateFile | Should -Be $false
        Test-Path -LiteralPath "$dir\store\content\営業\a.tsv" | Should -Be $true

        (removeOldSystemIndex $dir).Ok | Should -Be $true
    }

    It "中にあるジャンクションの先は消さず、system_index ごと消える（system_index 自身がジャンクションでも先は残る）" -TestCases @(
        @{ name = "system_index の中"; nested = $true }
        @{ name = "system_index 自身"; nested = $false }
    ) {
        param ($name, $nested)
        $dir = "$TestDrive\old2_$nested"
        $outside = "$TestDrive\old2_outside_$nested"
        newTsv "$outside\営業\システムインデックス.txt" @("x00000000")
        $link = if ($nested) { "$dir\system_index\営業" } else { "$dir\system_index" }
        [System.IO.Directory]::CreateDirectory((Split-Path $link -Parent)) | Out-Null
        if (!(newTestJunction $link "$outside\営業")) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        $result = removeOldSystemIndex $dir
        $result.Ok | Should -Be $true
        Test-Path -LiteralPath "$dir\system_index" | Should -Be $false
        Test-Path -LiteralPath "$outside\営業\システムインデックス.txt" | Should -Be $true
    }

    It "開いたままで消せなくても例外にせず、Ok=`$false を理由付きで返す。閉じてもう一度で消える" {
        $dir = "$TestDrive\old3"
        $ws = [Workspace]::new($dir)
        $txt = "$($ws.OldSystemIndexDir)\営業\システムインデックス.txt"
        newTsv $txt @("x00000000")
        $stream = [System.IO.File]::Open($txt, "Open", "Read", "None")
        try {
            $result = removeOldSystemIndex $dir
            $result.Ok | Should -Be $false
            $result.Reason | Should -Not -BeNullOrEmpty
        } finally {
            $stream.Dispose()
        }

        (removeOldSystemIndex $dir).Ok | Should -Be $true
        Test-Path -LiteralPath $ws.OldSystemIndexDir | Should -Be $false
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
