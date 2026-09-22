# 変換するかどうかの判断（tebunko_grep\convert\convert_decide.ps1）のテスト。
. "$PSScriptRoot\..\..\helpers\load.ps1"

function newRow {
    param ([string]$state, [string]$updated = "2026/01/01 10:00:00", [string]$size = "1000", [string]$version = "2", [string]$relPath = "売上\a.xlsx")
    return [pscustomobject]@{ 相対パス = $relPath; 更新日時 = $updated; サイズ = $size; 状態 = $state; TSV数 = "3"; 抽出版 = $version }
}

Describe "getExtractVersion" -Tag Unit {
    It "図形・コメントを読む形式は 2、それ以外は 1（大文字の拡張子も同じ）" {
        getExtractVersion "売上\a.xlsx" | Should Be 2
        getExtractVersion "売上\a.XLSM" | Should Be 2
        # Excel の旧形式・バイナリ形式は、図形・コメントを読まない
        getExtractVersion "売上\a.xls" | Should Be 1
        getExtractVersion "売上\a.xlsb" | Should Be 1
        # Word・PowerPoint は旧形式も新形式に変換してから読むため、どれも 2
        foreach ($ext in @(".docx", ".docm", ".doc", ".pptx", ".pptm", ".PPT")) {
            getExtractVersion "売上\a$ext" | Should Be 2
        }
        getExtractVersion "売上\a.txt" | Should Be 1
    }
}

Describe "getConvertDecision（抽出版）" -Tag Unit {
    It "前の抽出版で変換した「済」は、更新が無くても変換し直す（outdated）" {
        $d = getConvertDecision (newRow ${stateDone} -version "1") "2026/01/01 10:00:00" "1000" $true
        $d.Convert | Should Be $true
        $d.Reason | Should Be "outdated"
    }

    It "抽出版が空（以前の形式の変換一覧）は 1 とみなす" {
        (getConvertDecision (newRow ${stateDone} -version "") "2026/01/01 10:00:00" "1000" $true).Reason | Should Be "outdated"
        (getConvertDecision (newRow ${stateDone} -version "" -relPath "売上\a.docx") "2026/01/01 10:00:00" "1000" $true).Reason | Should Be "outdated"
        # 版が上がっていない形式（Excel の旧形式）は、抽出版が空でも変換し直さない
        (getConvertDecision (newRow ${stateDone} -version "" -relPath "売上\a.xls") "2026/01/01 10:00:00" "1000" $true).Reason | Should Be "done"
    }

    It "前回失敗は、抽出版が古くても failed のまま（再変換するかは呼び出し元が決める）" {
        (getConvertDecision (newRow ${stateFailed} -version "") "2026/01/01 10:00:00" "1000" $true).Reason | Should Be "failed"
    }
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