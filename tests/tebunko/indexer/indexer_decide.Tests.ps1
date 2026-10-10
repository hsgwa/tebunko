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
    It "Excel・Word・PowerPoint の新形式と .doc は 4、.ppt・旧形式・バイナリ形式は 1・2（大文字の拡張子も同じ）" {
        # Excel の新形式は、ヘッダー・フッターを読む版で 4
        getExtractVersion "売上\a.xlsx" | Should -Be 4
        getExtractVersion "売上\a.XLSM" | Should -Be 4
        # Excel の旧形式・バイナリ形式は、図形・コメント・グラフ・SmartArt を読まない
        getExtractVersion "売上\a.xls" | Should -Be 1
        getExtractVersion "売上\a.xlsb" | Should -Be 1
        # Word・PowerPoint の新形式は、埋め込みの中の文字を読む版で 4
        foreach ($ext in @(".docx", ".docm", ".pptx", ".PPTM")) {
            getExtractVersion "売上\a$ext" | Should -Be 4
        }
        # .doc は Word が新形式に変換した後も埋め込みが残るため 4。.ppt は変換後に埋め込みが残らないため上げていない
        getExtractVersion "売上\a.doc" | Should -Be 4
        getExtractVersion "売上\a.PPT" | Should -Be 2
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
        (getIngestDecision (newRow ${stateDone} -version "4" -relPath "報告\a.pptx") "2026/01/01 10:00:00" "1000" $true).Reason | Should -Be "done"
    }

    It "グラフの項目名をまだ読まなかった版（抽出版 2）の .xlsx・.docx は取り込み直す（outdated）" {
        (getIngestDecision (newRow ${stateDone} -version "2") "2026/01/01 10:00:00" "1000" $true).Reason | Should -Be "outdated"
        (getIngestDecision (newRow ${stateDone} -version "2" -relPath "報告\a.docx") "2026/01/01 10:00:00" "1000" $true).Reason | Should -Be "outdated"
        # .ppt は版を上げていないため、抽出版 2 でも取り込み直さない
        (getIngestDecision (newRow ${stateDone} -version "2" -relPath "報告\a.ppt") "2026/01/01 10:00:00" "1000" $true).Reason | Should -Be "done"
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

Describe "getIngestLane・getOfficeLane" -Tag Unit {
    It "Excel はすべて Excel のレーン、旧形式の Word・PowerPoint は Office のレーン、新形式は読み取りのレーン" {
        $expected = @(
            @("営業\a.xlsx", ${laneExcel}), @("a.xlsm", ${laneExcel}), @("a.xls", ${laneExcel}), @("a.xlsb", ${laneExcel}), @("A.XLSX", ${laneExcel}),
            @("a.doc", ${laneWord}), @("B.DOC", ${laneWord}),
            @("a.ppt", ${lanePowerPoint}), @("B.PPT", ${lanePowerPoint}),
            @("a.docx", ${laneReader}), @("a.docm", ${laneReader}), @("a.pptx", ${laneReader}), @("a.pptm", ${laneReader}), @("大文字.DOCX", ${laneReader}),
            # テキストの拡張子は、どの分岐にも当たらないため既定の読み取りのレーンになる
            @("a.txt", ${laneReader}), @("a.log", ${laneReader})
        )
        foreach ($case in $expected) {
            getIngestLane $case[0] | Should -Be $case[1]
        }
    }

    It "読み取りのレーンから回し直すときは、PowerPoint のファイルは PowerPoint、それ以外は Word のレーン" {
        getOfficeLane "a.pptx" | Should -Be ${lanePowerPoint}
        getOfficeLane "A.PPTM" | Should -Be ${lanePowerPoint}
        getOfficeLane "a.docx" | Should -Be ${laneWord}
        getOfficeLane "a.docm" | Should -Be ${laneWord}
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
