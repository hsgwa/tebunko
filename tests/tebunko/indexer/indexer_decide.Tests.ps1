# 取り込むかどうかの判断（tebunko\indexer\indexer_decide.ps1）のテスト。
BeforeDiscovery {
    # -TestCases の表が使う一覧の状態（$stateDone など）
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    function newRow {
        param ([string]$state, [string]$updated = "2026/01/01 10:00:00", [string]$size = "1000", [string]$version = "2", [string]$relPath = "売上\a.xlsx")
        return [pscustomobject]@{ 相対パス = $relPath; 更新日時 = $updated; サイズ = $size; 状態 = $state; TSV数 = "3"; 抽出版 = $version }
    }
}

Describe "getExtractVersion" -Tag Unit {
    It "Excel・Word の新形式と .doc は 4、PowerPoint は 5、旧形式・バイナリ形式は 1（大文字の拡張子も同じ）" {
        # Excel の新形式は、ヘッダー・フッターを読む版で 4
        getExtractVersion "売上\a.xlsx" | Should -Be 4
        getExtractVersion "売上\a.XLSM" | Should -Be 4
        # Excel の旧形式・バイナリ形式は、図形・コメント・グラフ・SmartArt を読まない
        getExtractVersion "売上\a.xls" | Should -Be 1
        getExtractVersion "売上\a.xlsb" | Should -Be 1
        # Word の新形式は、埋め込みの中の文字を読む版で 4。.doc も Word が新形式に変換した後に埋め込みが残るため 4
        foreach ($ext in @(".docx", ".docm", ".doc")) {
            getExtractVersion "売上\a$ext" | Should -Be 4
        }
        # PowerPoint は、スライドマスター・ノートマスターを読む版で 5（.ppt も .pptx に変換して読むため上げる）
        foreach ($ext in @(".pptx", ".PPTM", ".ppt", ".PPT")) {
            getExtractVersion "売上\a$ext" | Should -Be 5
        }
        getExtractVersion "売上\a.txt" | Should -Be 1
    }
}

Describe "getIngestDecision（抽出版）" -Tag Unit {
    It "抽出版が空は 1 とみなす" {
        (getIngestDecision (newRow ${stateDone} -version "") "2026/01/01 10:00:00" "1000" $true).Reason | Should -Be "outdated"
        (getIngestDecision (newRow ${stateDone} -version "" -relPath "売上\a.docx") "2026/01/01 10:00:00" "1000" $true).Reason | Should -Be "outdated"
        # 版が上がっていない形式（Excel の旧形式）は、抽出版が空でも取り込み直さない
        (getIngestDecision (newRow ${stateDone} -version "" -relPath "売上\a.xls") "2026/01/01 10:00:00" "1000" $true).Reason | Should -Be "done"
    }

    It "ヘッダー・フッター・埋め込みをまだ読まなかった版（抽出版 3）の .xlsx・.xlsm・.docx は取り込み直す" {
        (getIngestDecision (newRow ${stateDone} -version "3") "2026/01/01 10:00:00" "1000" $true).Reason | Should -Be "outdated"
        (getIngestDecision (newRow ${stateDone} -version "3" -relPath "売上\a.xlsm") "2026/01/01 10:00:00" "1000" $true).Reason | Should -Be "outdated"
        (getIngestDecision (newRow ${stateDone} -version "4") "2026/01/01 10:00:00" "1000" $true).Reason | Should -Be "done"
        (getIngestDecision (newRow ${stateDone} -version "3" -relPath "報告\a.docx") "2026/01/01 10:00:00" "1000" $true).Reason | Should -Be "outdated"
        (getIngestDecision (newRow ${stateDone} -version "4" -relPath "報告\a.docx") "2026/01/01 10:00:00" "1000" $true).Reason | Should -Be "done"
    }

    It "スライドマスター・ノートマスターをまだ読まなかった版の .pptx・.pptm・.ppt は取り込み直し、Excel・Word は取り込み直さない" {
        foreach ($case in @(@("a.pptx", "4"), @("a.pptm", "3"), @("a.ppt", "2"), @("a.pptx", ""))) {
            (getIngestDecision (newRow ${stateDone} -version $case[1] -relPath "報告\$($case[0])") "2026/01/01 10:00:00" "1000" $true).Reason | Should -Be "outdated"
        }
        foreach ($name in @("a.pptx", "a.pptm", "a.ppt")) {
            (getIngestDecision (newRow ${stateDone} -version "5" -relPath "報告\$name") "2026/01/01 10:00:00" "1000" $true).Reason | Should -Be "done"
        }
        (getIngestDecision (newRow ${stateDone} -version "4" -relPath "報告\a.xlsx") "2026/01/01 10:00:00" "1000" $true).Reason | Should -Be "done"
        (getIngestDecision (newRow ${stateDone} -version "4" -relPath "報告\a.docx") "2026/01/01 10:00:00" "1000" $true).Reason | Should -Be "done"
    }

    It "グラフの項目名をまだ読まなかった版（抽出版 2）の .xlsx・.docx は取り込み直す（outdated）" {
        (getIngestDecision (newRow ${stateDone} -version "2") "2026/01/01 10:00:00" "1000" $true).Reason | Should -Be "outdated"
        (getIngestDecision (newRow ${stateDone} -version "2" -relPath "報告\a.docx") "2026/01/01 10:00:00" "1000" $true).Reason | Should -Be "outdated"
        (getIngestDecision (newRow ${stateDone} -version "2" -relPath "報告\a.xls") "2026/01/01 10:00:00" "1000" $true).Reason | Should -Be "done"
    }
}

Describe "getIngestDecision" -Tag Unit {
    # 一覧の行（状態・抽出版。更新日時 2026/01/01 10:00:00・サイズ 1000）と、今のファイル（更新日時・サイズ）・インデックスの有無 → 取り込むか・理由。
    # state が $null の行は、一覧に無いファイル
    It "<name>" -TestCases @(
        @{ name = "一覧に無ければ取り込む（new）"; state = $null; version = "4"; updated = "2026/01/01 10:00:00"; size = "1000"; indexed = $true; ingest = $true; reason = "new" }
        @{ name = "更新日時が変わっていれば取り込む（updated）"; state = $stateDone; version = "4"; updated = "2026/02/02 09:00:00"; size = "1000"; indexed = $true; ingest = $true; reason = "updated" }
        @{ name = "サイズが変わっていれば取り込む（updated）"; state = $stateDone; version = "4"; updated = "2026/01/01 10:00:00"; size = "2000"; indexed = $true; ingest = $true; reason = "updated" }
        @{ name = "取り込み済みで更新も無ければ取り込まない（done）"; state = $stateDone; version = "4"; updated = "2026/01/01 10:00:00"; size = "1000"; indexed = $true; ingest = $false; reason = "done" }
        @{ name = "取り込み済みでもインデックスが無ければ取り込み直す（lost）"; state = $stateDone; version = "4"; updated = "2026/01/01 10:00:00"; size = "1000"; indexed = $false; ingest = $true; reason = "lost" }
        @{ name = "前の抽出版で取り込んだ「済」は、更新が無くても取り込み直す（outdated）"; state = $stateDone; version = "1"; updated = "2026/01/01 10:00:00"; size = "1000"; indexed = $true; ingest = $true; reason = "outdated" }
        @{ name = "前回失敗し、更新も無ければ取り込まない（failed。再取り込みするかは呼び出し元が決める）"; state = $stateFailed; version = "4"; updated = "2026/01/01 10:00:00"; size = "1000"; indexed = $true; ingest = $false; reason = "failed" }
        @{ name = "前回失敗でも、更新されていれば取り込む（updated）"; state = $stateFailed; version = "4"; updated = "2026/03/03 08:00:00"; size = "1000"; indexed = $true; ingest = $true; reason = "updated" }
        @{ name = "前回失敗は、インデックスの有無を見ない（failed のまま）"; state = $stateFailed; version = "4"; updated = "2026/01/01 10:00:00"; size = "1000"; indexed = $false; ingest = $false; reason = "failed" }
        @{ name = "前回失敗は、抽出版が古くても failed のまま（再取り込みするかは呼び出し元が決める）"; state = $stateFailed; version = ""; updated = "2026/01/01 10:00:00"; size = "1000"; indexed = $true; ingest = $false; reason = "failed" }
        @{ name = "前回「未取り込み」で終わっていれば取り込む（pending）"; state = $stateNew; version = "4"; updated = "2026/01/01 10:00:00"; size = "1000"; indexed = $true; ingest = $true; reason = "pending" }
    ) {
        param ($name, $state, $version, $updated, $size, $indexed, $ingest, $reason)
        $row = if ($null -eq $state) { $null } else { newRow $state -version $version }
        $d = getIngestDecision $row $updated $size $indexed
        $d.Ingest | Should -Be $ingest
        $d.Reason | Should -Be $reason
    }
}

Describe "getIngestDecision（利用者のPowerPointのために取り込めなかった行）" -Tag Unit {
    It "「未取り込み」でエラー（理由）が書いてある行は、更新が無くても取り込む（pending）。「失敗」と違って、再取り込みを選ばなくてよい" {
        $row = newStatusRow "報告\a.ppt" "2026/01/01 10:00:00" "1000" $stateNew "" "" ${postponedReasonRunning}
        $d = getIngestDecision $row "2026/01/01 10:00:00" "1000" $true
        $d.Ingest | Should -Be $true
        $d.Reason | Should -Be "pending"
    }
}

Describe "getIngestApp・getRerouteApp" -Tag Unit {
    It "Excel はすべて Excel、旧形式の Word・PowerPoint はそのアプリ、新形式・テキスト・PDF は Office が要らない（空）" {
        $expected = @(
            @("営業\a.xlsx", "Excel"), @("a.xlsm", "Excel"), @("a.xls", "Excel"), @("a.xlsb", "Excel"), @("A.XLSX", "Excel"),
            @("a.doc", "Word"), @("B.DOC", "Word"),
            @("a.ppt", "PowerPoint"), @("B.PPT", "PowerPoint"),
            @("a.docx", ""), @("a.docm", ""), @("a.pptx", ""), @("a.pptm", ""), @("大文字.DOCX", ""),
            # テキストの拡張子は、どの分岐にも当たらないため Office が要らない
            @("a.txt", ""), @("a.log", ""),
            # PDF も Office を使わずに読む
            @("a.pdf", ""), @("営業\B.PDF", ""),
            # 表（取り込みの対象の拡張子）に無い拡張子は Office とみなさない
            @("a.xlsxyz", ""), @("a.xlsa", ""), @("a.docz", ""), @("a.pptz", ""), @("a.xlsx.bak", ""), @("", "")
        )
        foreach ($case in $expected) {
            getIngestApp $case[0] | Should -Be $case[1]
        }
    }

    It "Office を使わずに読めなかったファイルの回し直し先は、PowerPoint のファイルは PowerPoint、それ以外は Word" {
        getRerouteApp "a.pptx" | Should -Be "PowerPoint"
        getRerouteApp "A.PPTM" | Should -Be "PowerPoint"
        getRerouteApp "a.ppt" | Should -Be "PowerPoint"
        # 表に無い拡張子は PowerPoint ではないため、Word
        getRerouteApp "a.pptz" | Should -Be "Word"
        getRerouteApp "a.docx" | Should -Be "Word"
        getRerouteApp "a.docm" | Should -Be "Word"
    }
}

Describe "selectIngestWorker" -Tag Unit {
    BeforeDiscovery {
        function newWorker([int]$busy = 0, [string[]]$owns = @(), [bool]$abandoned = $false) {
            return @{ Busy = $busy; Owns = $owns; Abandoned = $abandoned }
        }
    }

    It "<name>" -TestCases @(
        @{ name = "Office のファイルは、そのアプリの持ち主へ（ほかが空いていても）"
           workers = @((newWorker 0), (newWorker 1 @("Excel")), (newWorker 0)); app = "Excel"; kind = "Worker"; index = 1 }
        @{ name = "持ち主が満杯なら、ほかが空いていても待つ（そのアプリは持ち主だけが扱う）"
           workers = @((newWorker 0), (newWorker 2 @("Excel"))); app = "Excel"; kind = "Wait"; index = -1 }
        @{ name = "持ち主がいなければ、Office を持たないワーカーのうち一番空いているものが持ち主になる"
           workers = @((newWorker 1 @("Word")), (newWorker 1), (newWorker 0)); app = "Excel"; kind = "Worker"; index = 2 }
        @{ name = "持ち主がいなければ、Office を持たないワーカーが全部満杯なら待つ"
           workers = @((newWorker 0 @("Word")), (newWorker 2), (newWorker 2)); app = "Excel"; kind = "Wait"; index = -1 }
        @{ name = "Office を持たないワーカーが 1 つも無いときだけ、空いているワーカーが 2 つ目のアプリを持つ"
           workers = @((newWorker 2 @("Word")), (newWorker 1 @("Excel"))); app = "PowerPoint"; kind = "Worker"; index = 1 }
        @{ name = "Office が要らないファイルは、Office を持たないワーカーを先に（持ち主が空いていても）"
           workers = @((newWorker 0 @("Excel")), (newWorker 1), (newWorker 0)); app = ""; kind = "Worker"; index = 2 }
        @{ name = "Office が要らないファイルは、Office を持たないワーカーが満杯なら、空いている持ち主へ"
           workers = @((newWorker 2), (newWorker 1 @("Excel")), (newWorker 2 @("Word"))); app = ""; kind = "Worker"; index = 1 }
        @{ name = "全部が満杯なら待つ（Office が要らないファイル）"
           workers = @((newWorker 2), (newWorker 2 @("Excel"))); app = ""; kind = "Wait"; index = -1 }
        @{ name = "持ち主が諦めたワーカーなら、そのアプリのファイルは後回し"
           workers = @((newWorker 0 @("PowerPoint") $true), (newWorker 0)); app = "PowerPoint"; kind = "Abandoned"; index = 0 }
        @{ name = "諦めたワーカーには、Office が要らないファイルも渡さない"
           workers = @((newWorker 0 @("PowerPoint") $true), (newWorker 1)); app = ""; kind = "Worker"; index = 1 }
        @{ name = "諦めていないアプリのファイルは、諦めたワーカーがいても持ち主へ"
           workers = @((newWorker 0 @("PowerPoint") $true), (newWorker 1 @("Excel"))); app = "Excel"; kind = "Worker"; index = 1 }
        @{ name = "ワーカーが全部諦めたなら、後回し"
           workers = @((newWorker 0 @("PowerPoint") $true)); app = ""; kind = "Abandoned"; index = -1 }
        @{ name = "回し直し（容量を大きくして呼ぶ）は、持ち主が満杯でも渡す"
           workers = @((newWorker 5 @("Word")), (newWorker 0)); app = "Word"; kind = "Worker"; index = 0; capacity = 1000 }
    ) {
        $cap = if ($capacity) { $capacity } else { 2 }
        $pick = selectIngestWorker $workers $app $cap
        $pick.Kind | Should -Be $kind
        $pick.Index | Should -Be $index
    }
}

Describe "splitCloudItems" -Tag Unit {
    # items: 取り込み対象（Failed が $false）と、前回失敗（Failed が $true）が 1 件ずつ
    It "<name>" -TestCases @(
        @{ name = "ダウンロードしないなら、すべて残す"; retry = $false; include = $false; included = @(); skipped = @("新規.xlsx", "失敗.xlsx") }
        @{ name = "ダウンロードしないなら、失敗分を再取り込みにしても残す"; retry = $true; include = $false; included = @(); skipped = @("新規.xlsx", "失敗.xlsx") }
        @{ name = "ダウンロードするなら、取り込み対象だけ取り込む（失敗分は再取り込みを選んでいなければ残す）"; retry = $false; include = $true; included = @("新規.xlsx"); skipped = @("失敗.xlsx") }
        @{ name = "ダウンロードして、失敗分も再取り込みするなら、すべて取り込む"; retry = $true; include = $true; included = @("新規.xlsx", "失敗.xlsx"); skipped = @() }
    ) {
        param ($name, $retry, $include, $included, $skipped)
        $items = @(
            @{ RelPath = "新規.xlsx"; Failed = $false }
            @{ RelPath = "失敗.xlsx"; Failed = $true }
        )
        $result = splitCloudItems $items $retry $include
        @($result.Included | ForEach-Object { $_.RelPath }) | Should -Be $included
        @($result.Skipped | ForEach-Object { $_.RelPath }) | Should -Be $skipped
    }

    It "空の一覧と `$null でも、空の結果を返す" {
        (splitCloudItems @() $true $true).Included.Count | Should -Be 0
        (splitCloudItems $null $true $true).Skipped.Count | Should -Be 0
    }

    It "要素の無い List[object]（createTargetList が返す形）でも、空の結果を返す" {
        $list = New-Object System.Collections.Generic.List[object]
        $result = splitCloudItems $list $true $true
        $result.Included.Count | Should -Be 0
        $result.Skipped.Count | Should -Be 0
    }
}

Describe "convertToPendingRemoveLines / convertFromPendingRemoveLines" -Tag Unit {
    It "記録の行は、組み立てて読み戻すと同じ（重なりは 1 行にまとめ、空の項目は書かない）" {
        $lines = convertToPendingRemoveLines @(
            @{ Folder = "営業\2024"; Name = "A社.xlsx" },
            @{ Folder = "営業\2024"; Name = "A社.xlsx" },
            @{ Folder = "経理"; Name = "b.docx" },
            @{ Folder = ""; Name = "x.xlsx" },
            @{ Folder = "経理"; Name = "" },
            $null
        )

        @($lines).Count | Should -Be 2
        $back = @(convertFromPendingRemoveLines $lines)
        $back.Count | Should -Be 2
        ($back | Where-Object { $_.Folder -eq "営業\2024" }).Name | Should -Be "A社.xlsx"
        ($back | Where-Object { $_.Folder -eq "経理" }).Name | Should -Be "b.docx"
    }

    It "壊れた行（<name>）は飛ばす" -TestCases @(
        @{ name = "タブが無い"; line = "営業A社.xlsx" },
        @{ name = "項目が 3 つ"; line = "営業`tA社.xlsx`t余り" },
        @{ name = "フォルダが空"; line = "`tA社.xlsx" },
        @{ name = "名前が空"; line = "営業`t" },
        @{ name = "空の行"; line = "" }
    ) {
        param ($line)
        @(convertFromPendingRemoveLines @($line, "営業`tok.xlsx")).Count | Should -Be 1
    }

    It "行が無ければ空" {
        @(convertToPendingRemoveLines @()).Count | Should -Be 0
        @(convertFromPendingRemoveLines @()).Count | Should -Be 0
    }
}

Describe "getFullyIngestedIndexNames" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "名前の有効なフォルダをすべて回した名前だけ返す"; ran = @("C:\a", "C:\b", "C:\c"); folders = @(
            [pscustomobject]@{ Name = "営業"; Path = "C:\a"; Enabled = $true }, [pscustomobject]@{ Name = "営業"; Path = "C:\b"; Enabled = $true },
            [pscustomobject]@{ Name = "経理"; Path = "C:\c"; Enabled = $true }); expected = @("営業", "経理") }
        @{ name = "同じ名前のフォルダの 1 つだけを回した名前は返さない"; ran = @("C:\a"); folders = @(
            [pscustomobject]@{ Name = "営業"; Path = "C:\a"; Enabled = $true }, [pscustomobject]@{ Name = "営業"; Path = "C:\b"; Enabled = $true }); expected = @() }
        @{ name = "チェックなしのフォルダは数えない（有効なものをすべて回せばよい）。有効なフォルダが無い名前は返さない"; ran = @("C:\a"); folders = @(
            [pscustomobject]@{ Name = "営業"; Path = "C:\a"; Enabled = $true }, [pscustomobject]@{ Name = "営業"; Path = "C:\b"; Enabled = $false },
            [pscustomobject]@{ Name = "総務"; Path = "C:\z"; Enabled = $false }); expected = @("営業") }
        @{ name = "パスの大文字・小文字は区別しない"; ran = @("c:\A"); folders = @(
            [pscustomobject]@{ Name = "営業"; Path = "C:\a"; Enabled = $true }); expected = @("営業") }
        @{ name = "回したフォルダが無ければ空"; ran = @(); folders = @([pscustomobject]@{ Name = "営業"; Path = "C:\a"; Enabled = $true }); expected = @() }
    ) {
        param ($name, $ran, $folders, $expected)
        @(getFullyIngestedIndexNames $ran $folders) | Should -Be $expected
    }
}
