# 取り込むかどうかの判断（tebunko_grep\indexer\indexer_decide.ps1）のテスト。
. "$PSScriptRoot\..\..\helpers\load.ps1"

function newRow {
    param ([string]$state, [string]$updated = "2026/01/01 10:00:00", [string]$size = "1000")
    return [pscustomobject]@{ 相対パス = "売上\a.xlsx"; 更新日時 = $updated; サイズ = $size; 状態 = $state; TSV数 = "3" }
}

Describe "getIngestDecision" -Tag Unit {
    It "一覧に無ければ取り込む（new）" {
        $d = getIngestDecision $null "2026/01/01 10:00:00" "1000" $true
        $d.Ingest | Should Be $true
        $d.Reason | Should Be "new"
    }

    It "更新日時が変わっていれば取り込む（updated）" {
        $d = getIngestDecision (newRow ${stateDone}) "2026/02/02 09:00:00" "1000" $true
        $d.Ingest | Should Be $true
        $d.Reason | Should Be "updated"
    }

    It "サイズが変わっていれば取り込む（updated）" {
        $d = getIngestDecision (newRow ${stateDone}) "2026/01/01 10:00:00" "2000" $true
        $d.Ingest | Should Be $true
        $d.Reason | Should Be "updated"
    }

    It "取り込み済みで更新も無ければ取り込まない（done）" {
        $d = getIngestDecision (newRow ${stateDone}) "2026/01/01 10:00:00" "1000" $true
        $d.Ingest | Should Be $false
        $d.Reason | Should Be "done"
    }

    It "取り込み済みでもインデックスが無ければ取り込み直す（lost）" {
        $d = getIngestDecision (newRow ${stateDone}) "2026/01/01 10:00:00" "1000" $false
        $d.Ingest | Should Be $true
        $d.Reason | Should Be "lost"
    }

    It "前回失敗し、更新も無ければ取り込まない（failed。再取り込みするかは呼び出し元が決める）" {
        $d = getIngestDecision (newRow ${stateFailed}) "2026/01/01 10:00:00" "1000" $true
        $d.Ingest | Should Be $false
        $d.Reason | Should Be "failed"
    }

    It "前回失敗でも、更新されていれば取り込む（updated）" {
        $d = getIngestDecision (newRow ${stateFailed}) "2026/03/03 08:00:00" "1000" $true
        $d.Ingest | Should Be $true
        $d.Reason | Should Be "updated"
    }

    It "前回「未取り込み」で終わっていれば取り込む（pending）" {
        $d = getIngestDecision (newRow ${stateNew}) "2026/01/01 10:00:00" "1000" $true
        $d.Ingest | Should Be $true
        $d.Reason | Should Be "pending"
    }

    It "前回失敗は、インデックスの有無を見ない（failed のまま）" {
        $d = getIngestDecision (newRow ${stateFailed}) "2026/01/01 10:00:00" "1000" $false
        $d.Ingest | Should Be $false
        $d.Reason | Should Be "failed"
    }
}