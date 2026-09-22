# ［2 比較］タブの文言と可否（tebunko_diff\ui\diff_view.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\tebunko_diff\lib.ps1"
. "${scriptsDir}\tebunko_diff\ui\diff_view.ps1"

function newItem {
    param ([string]$path, [bool]$exists = $true, [bool]$isFolder = $false)

    return @{ Path = $path; Exists = $exists; IsFolder = $isFolder }
}

Describe "getModeView" -Tag Unit {
    It "［ファイル］ではフォルダ用のオプションを押せず、理由をツールチップに出す" {
        $view = getModeView "file"
        $view.FolderOptions | Should Be $false
        $view.PickLabel | Should Be "ファイル…"
        $view.FolderOptionsTip | Should Be "フォルダを比べるときだけ使えます"
    }

    It "［フォルダ］ではフォルダ用のオプションを使える" {
        $view = getModeView "folder"
        $view.FolderOptions | Should Be $true
        $view.Placeholder | Should Be "フォルダを選ぶか、ここへドロップ"
    }
}

Describe "getCompareBlockReason" -Tag Unit {
    It "押せるとき" {
        getCompareBlockReason "file" (newItem "C:\a\見積.xlsx") (newItem "C:\b\見積.xls") | Should Be ""
        getCompareBlockReason "folder" (newItem "C:\a" $true $true) (newItem "C:\b" $true $true) | Should Be ""
    }

    It "空・見つからない" {
        getCompareBlockReason "file" (newItem "") (newItem "") | Should Be "比較元と比較先のファイルを選んでください。"
        getCompareBlockReason "file" (newItem "C:\a.xlsx") (newItem "") | Should Be "比較先のファイルを選んでください。"
        getCompareBlockReason "folder" (newItem "C:\a" $false $true) (newItem "C:\b" $true $true) | Should Be "比較元のフォルダが見つかりません。"
    }

    It "トグルと違うもの" {
        getCompareBlockReason "file" (newItem "C:\a" $true $true) (newItem "C:\b.xlsx") | Should Match "^比較元はフォルダです"
        getCompareBlockReason "folder" (newItem "C:\a" $true $true) (newItem "C:\b.xlsx") | Should Match "^比較先はフォルダではありません"
    }

    It "Office でない・種類が違う・同じもの" {
        getCompareBlockReason "file" (newItem "C:\a.txt") (newItem "C:\b.xlsx") | Should Be "Excel・Word・PowerPoint のファイルを選んでください。"
        getCompareBlockReason "file" (newItem "C:\a.xlsx") (newItem "C:\b.docx") | Should Be "種類の違うファイル（Excel と Word）は比べられません。"
        getCompareBlockReason "file" (newItem "C:\A.xlsx") (newItem "c:\a.xlsx") | Should Be "同じファイルが選ばれています。"
    }

    It "一方のフォルダがもう一方の中にあるとき（サブフォルダも比べるときだけ）" {
        getCompareBlockReason "folder" (newItem "C:\営業" $true $true) (newItem "C:\営業\2025" $true $true) | Should Be "比較先のフォルダが比較元のフォルダの中にあります。"
        getCompareBlockReason "folder" (newItem "C:\営業" $true $true) (newItem "C:\営業\2025" $true $true) $false | Should Be ""
    }
}

Describe "getDropAction" -Tag Unit {
    It "2 つまとめてドロップすると、名前の順に比較元・比較先へ入れ、トグルを合わせる" {
        $action = getDropAction @(@{ Path = "C:\b"; IsFolder = $true }, @{ Path = "C:\a"; IsFolder = $true }) "" "file"
        $action.Mode | Should Be "folder"
        $action.Left | Should Be "C:\a"
        $action.Right | Should Be "C:\b"
    }

    It "ファイルとフォルダを混ぜてドロップしたら入れない" {
        getDropAction @(@{ Path = "C:\a.xlsx"; IsFolder = $false }, @{ Path = "C:\b"; IsFolder = $true }) "" "file" | Should BeNullOrEmpty
    }

    It "1 つはドロップした側へ入れ、トグルと違うものならトグルを切り替える" {
        $action = getDropAction @(@{ Path = "C:\営業"; IsFolder = $true }) "right" "file"
        $action.Mode | Should Be "folder"
        $action.Left | Should BeNullOrEmpty
        $action.Right | Should Be "C:\営業"
    }
}

Describe "要約" -Tag Unit {
    It "ファイル同士（Excel）" {
        $diff = compareOfficeUnits "Excel" ([ordered]@{ A = [string[]]@("a", "b"); B = [string[]]@("x") }) ([ordered]@{ A = [string[]]@("a", "c", "d"); B = [string[]]@("x") })
        $text = getFileSummaryText $diff
        $text.Title | Should Be "シート 2 のうち 1 に変更"
        # 1 文字だけの行は似ている度合いが 0 のため、変更に組まず削除と追加にする
        $text.Counts | Should Be "追加 2・削除 1"
    }

    It "違いが無いとき" {
        $diff = compareOfficeUnits "Excel" ([ordered]@{ A = [string[]]@("a") }) ([ordered]@{ A = [string[]]@("a") })
        (getFileSummaryText $diff).Title | Should Be "違いはありません。（シート 1）"
    }

    It "フォルダ同士" {
        getFolderSummaryText ([ordered]@{ Total = 10; change = 2; insert = 1; delete = 0; failed = 1; similar = 1; same = 5; pending = 0; running = 0 }) |
            Should Be "ファイル 10（変更 2・追加 1・比較できない 1・中身は同じ 1・同じ 5）"
        getFolderProgressText ([ordered]@{ pending = 3; running = 1 }) 9 | Should Be "中身を比べています 5 / 9"
        getFolderProgressText ([ordered]@{ pending = 0; running = 0 }) 9 | Should Be ""
    }

    It "場所の見出し" {
        $place = [PlaceDiff]::new()
        $place.Name = "明細"
        $place.Status = "change"
        $place.Changes = 3
        $place.Inserts = 1
        getPlaceTabText $place | Should Be "明細  4"
        $place.Status = "insert"
        getPlaceTabText $place | Should Be "明細（追加）"
    }
}

Describe "findNextChangeRow" -Tag Unit {
    function newRows {
        param ([string]$kinds)

        return @($kinds.ToCharArray() | ForEach-Object {
            $row = [DiffRow]::new()
            $row.Kind = switch ($_) { "c" { "change" } "i" { "insert" } default { "same" } }
            $row
        })
    }

    It "続いている変更は 1 つの塊とし、その先頭へ移る" {
        $rows = newRows "sscisssc"
        findNextChangeRow $rows -1 1 | Should Be 2
        findNextChangeRow $rows 2 1 | Should Be 7
        findNextChangeRow $rows 7 1 | Should Be -1
        findNextChangeRow $rows 7 -1 | Should Be 2
    }

    It "最初に開く場所は、違いのある最初の場所" {
        $a = [PlaceDiff]::new(); $a.Status = "same"
        $b = [PlaceDiff]::new(); $b.Status = "change"
        getDefaultPlaceIndex @($a, $b) | Should Be 1
        getDefaultPlaceIndex @($a) | Should Be 0
    }
}

Describe "getDiffKind" -Tag Unit {
    It "拡張子から種類" {
        getDiffKind "a.XLSM" | Should Be "Excel"
        getDiffKind "a.doc" | Should Be "Word"
        getDiffKind "a.pptx" | Should Be "PowerPoint"
        getDiffKind "a.pdf" | Should Be ""
        getExtractFailureText "right" "パスワード" | Should Be "比較先を読み取れませんでした。（パスワード）"
    }
}
