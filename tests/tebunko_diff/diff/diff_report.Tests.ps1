# 比較結果ファイルの文字（tebunko_diff\diff\diff_report.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\tebunko_diff\lib.ps1"

Describe "getFileReport" -Tag Unit {
    $left = [ordered]@{ "明細" = [string[]]@("品名`t金額", "サーバー構築`t120,000", "送料`t500"); "表紙" = [string[]]@("御見積書") }
    $right = [ordered]@{ "明細" = [string[]]@("品名`t金額", "サーバー構築`t144,000", "保守費用`t240,000"); "表紙" = [string[]]@("御見積書") }
    $diff = compareOfficeUnits "Excel" $left $right
    $lines = getFileReport "C:\共有\営業部\見積_v1.xlsx" "C:\共有\営業部\見積_v2.xlsx" $diff $null ([datetime]"2026-09-22 10:15:00")

    It "先頭に比較元・比較先・日時・設定・要約" {
        $lines[0] | Should Be "比較元: C:\共有\営業部\見積_v1.xlsx"
        $lines[2] | Should Be "日時:   2026/09/22 10:15:00"
        $lines[3] | Should Be "設定:   図形も比較・コメントも比較"
        $lines[4] | Should Be "要約:   シート 2 のうち 1 に変更（追加 1・削除 1・変更 1）"
    }

    It "違いのある場所ごとに、変更の行だけを並べる" {
        $body = $lines[6..($lines.Count - 1)] -join "`n"
        $body | Should Match "■ 明細"
        $body | Should Not Match "■ 表紙"
        $body | Should Match "変更`t2 → 2`tB2：120,000 → 144,000"
        $body | Should Match "削除`t3 → \(2\)`t送料`t500"
        $body | Should Match "追加`t\(3\) → 3`t保守費用`t240,000"
    }
}

Describe "getFolderReport" -Tag Unit {
    It "ファイルの一覧の後に、変更のあるファイルの中身を続ける" {
        $entries = getFolderEntries @(@{ RelPath = "a.docx"; Size = 1; Time = [datetime]::Now }, @{ RelPath = "b.docx"; Size = 1; Time = [datetime]::Now }) @(@{ RelPath = "a.docx"; Size = 2; Time = [datetime]::Now }, @{ RelPath = "c.docx"; Size = 1; Time = [datetime]::Now })
        $diff = compareOfficeUnits "Word" ([ordered]@{ "ページ001" = [string[]]@("x") }) ([ordered]@{ "ページ001" = [string[]]@("x", "y") })
        setEntryResult $entries[0] $diff
        $lines = getFolderReport "C:\左" "C:\右" $entries @{ "a.docx" = $diff } $null $true ([datetime]"2026-09-22")
        $lines[3] | Should Be "設定:   図形も比較・コメントも比較・サブフォルダも比較"
        $lines[4] | Should Be "要約:   ファイル 3（変更 1・追加 1・削除 1）"
        $lines[6] | Should Be "変更`ta.docx`t追加 1"
        $lines[7] | Should Be "削除`tb.docx`t"
        ($lines -join "`n") | Should Match "==== a.docx（場所 1 のうち 1 に変更（追加 1））"
    }
}

Describe "getOptionsText" -Tag Unit {
    It "設定の文言" {
        getOptionsText @{ IncludeShapes = $false; IncludeComments = $false; IgnoreWhitespace = $true; CaseSensitive = $false } | Should Be "空白の違いを無視・大文字と小文字を区別しない"
        getOptionsText @{ IncludeShapes = $false; IncludeComments = $false } | Should Be "（なし）"
    }
}


Describe "比較結果ファイル（細かい場合）" -Tag Unit {
    It "PowerPoint: スライドの見出しの行と、要約" {
        $l = [ordered]@{ "スライド001" = [string[]]@("表紙") }
        $r = [ordered]@{ "スライド001" = [string[]]@("表紙"); "スライド002" = [string[]]@("追加") }
        $diff = compareOfficeUnits "PowerPoint" $l $r
        $lines = getFileReportLines $diff
        ($lines -join "`n") | Should Match "□ スライド 2 追加"
        (getFileSummaryText $diff).Title | Should Be "スライド 1 → 2"
        $same = compareOfficeUnits "PowerPoint" $l $l
        (getFileSummaryText $same).Title | Should Be "スライド 1 → 1（違いはありません）"
        getFileSummaryParts $same | Should Be "スライド 1 → 1（違いはありません）"
    }

    It "注意のある場所は注意を出し、先頭の追加の位置は「先頭」" {
        $place = newTooLargePlace "S" "S" "S" ([string[]]@("a")) ([string[]]@("b"))
        $diff = [FileDiff]::new()
        $diff.Places = @($place)
        (getFileReportLines $diff)[1] | Should Match "行ごとには比べていません"
        $inserted = compareOfficeUnits "Word" ([ordered]@{ "ページ001" = [string[]]@("b") }) ([ordered]@{ "ページ001" = [string[]]@("a", "b") })
        ((getFileReportLines $inserted) -join "`n") | Should Match "\(先頭\) → p.1 ¶1"
    }

    It "状態の名前と、比較できないファイルの原因" {
        getStatusLabel "same" | Should Be "同じ"
        getStatusLabel "similar" | Should Be "中身は同じ"
        getStatusLabel "failed" | Should Be "比較できない"
        getStatusLabel "running" | Should Be "比較中"
        getStatusLabel "pending" | Should Be "比較待ち"
        getKindLabel "same" | Should Be "同じ"
        $entries = getFolderEntries @(@{ RelPath = "a.xlsx"; Size = 1; Time = [datetime]::Now }) @(@{ RelPath = "a.xlsx"; Size = 2; Time = [datetime]::Now })
        $entries[0].Status = "failed"
        $entries[0].Error = "比較元を読み取れませんでした。（パスワード）"
        $lines = getFolderReport "C:\左" "C:\右" $entries @{} $null $false ([datetime]"2026-09-22")
        $lines[6] | Should Be "比較できない`ta.xlsx`t比較元を読み取れませんでした。（パスワード）"
    }
}
