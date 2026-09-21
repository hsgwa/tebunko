# 変換するかどうかの判断（windox_grep\convert\convert_decide.ps1）のテスト。
. "$PSScriptRoot\..\..\helpers\load.ps1"

function newRow {
    param ([string]$state, [string]$updated = "2026/01/01 10:00:00", [string]$size = "1000")
    return [pscustomobject]@{ 相対パス = "売上\a.xlsx"; 更新日時 = $updated; サイズ = $size; 状態 = $state; TSV数 = "3" }
}

Describe "getConvertDecision" -Tag Unit {
    It "一覧に無ければ変換する（new）" {
        $d = getConvertDecision $null "2026/01/01 10:00:00" "1000" $true
        $d.Convert | Should Be $true
        $d.Reason | Should Be "new"
    }

    It "更新日時が変わっていれば変換する（updated）" {
        $d = getConvertDecision (newRow ${stateDone}) "2026/02/02 09:00:00" "1000" $true
        $d.Convert | Should Be $true
        $d.Reason | Should Be "updated"
    }

    It "サイズが変わっていれば変換する（updated）" {
        $d = getConvertDecision (newRow ${stateDone}) "2026/01/01 10:00:00" "2000" $true
        $d.Convert | Should Be $true
        $d.Reason | Should Be "updated"
    }

    It "変換済みで更新も無ければ変換しない（done）" {
        $d = getConvertDecision (newRow ${stateDone}) "2026/01/01 10:00:00" "1000" $true
        $d.Convert | Should Be $false
        $d.Reason | Should Be "done"
    }

    It "変換済みでも変換結果が無ければ変換し直す（lost）" {
        $d = getConvertDecision (newRow ${stateDone}) "2026/01/01 10:00:00" "1000" $false
        $d.Convert | Should Be $true
        $d.Reason | Should Be "lost"
    }

    It "前回失敗し、更新も無ければ変換しない（failed。再変換するかは呼び出し元が決める）" {
        $d = getConvertDecision (newRow ${stateFailed}) "2026/01/01 10:00:00" "1000" $true
        $d.Convert | Should Be $false
        $d.Reason | Should Be "failed"
    }

    It "前回失敗でも、更新されていれば変換する（updated）" {
        $d = getConvertDecision (newRow ${stateFailed}) "2026/03/03 08:00:00" "1000" $true
        $d.Convert | Should Be $true
        $d.Reason | Should Be "updated"
    }

    It "前回「未変換」で終わっていれば変換する（pending）" {
        $d = getConvertDecision (newRow ${stateNew}) "2026/01/01 10:00:00" "1000" $true
        $d.Convert | Should Be $true
        $d.Reason | Should Be "pending"
    }

    It "前回失敗は、変換結果の有無を見ない（failed のまま）" {
        $d = getConvertDecision (newRow ${stateFailed}) "2026/01/01 10:00:00" "1000" $false
        $d.Convert | Should Be $false
        $d.Reason | Should Be "failed"
    }
}