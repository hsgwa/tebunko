# 【一時】終了コード 5 が、ダイアログのあとの PowerShell の終わりの片づけで決まるかを確かめる（調べが済んだら消す）。
# 条件ごとに画面を 6 回ずつ起動して閉じる。落ちた条件は、場面の名前（条件の名前-回）と終了の記録で分かる。
# 環境変数 TEBUNKO_BISECT が立っているときだけ条件を作る（ふだんの Gui の実行では何もしない）。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"
}

Describe "終了コードの切り分け" -Tag Gui {
    BeforeAll {
        $script:toolIndex = newGuiTool (Join-Path $TestDrive "index")
        newGuiSampleIndex $script:toolIndex (Join-Path $TestDrive "index_root")
    }

    It "<Name>" -ForEach (@(
        if ($env:TEBUNKO_BISECT) {
            foreach ($n in 1..6) {
                @{ Name = "E2基準-$n"; Variant = "exiting" }
                @{ Name = "F2Exit-$n"; Variant = "envexit" }
                @{ Name = "S1Dispatcher-$n"; Variant = "exiting,dispatcher" }
                @{ Name = "S2GC-$n"; Variant = "exiting,gc" }
                @{ Name = "S3exit0-$n"; Variant = "exiting,exit0" }
            }
        }
    ) | Where-Object { !$env:TEBUNKO_BISECT_ONLY -or $_.Name -eq $env:TEBUNKO_BISECT_ONLY }) {
        $tool = $script:toolIndex
        $env:TEBUNKO_TEST_EXIT_VARIANT = $Variant
        try {
            $S = startGui $tool $Name
        } finally {
            $env:TEBUNKO_TEST_EXIT_VARIANT = $null
        }
        invokeGuiScene $S {
            if ($false) {
                setGuiStep $S "検索"
                setGuiText $S (waitGuiById $S $S.Window "WordBox") "単価"
                clickGui $S $S.Window "SearchButton" "［検索］"
                waitGui $S "件数の表示（該当 2 件）" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "SummaryText")) -like "一致 2 件（*" } | Out-Null
                setGuiStep $S "行を選ぶ"
                clickGui $S $S.Window "ExpandAllButton" "［すべて開く］"
                $row = waitGui $S "結果の行" ${guiDefaultTimeout} { getGuiHitRows (findGui $S.Window -Id "ResultGrid") | Select-Object -Last 1 }
                selectGui $row
                waitGui $S "プレビュー（りんご）" ${guiDefaultTimeout} {
                    @(findAllGui (findGui $S.Window -Id "PreviewScroll") -Type Text | Where-Object { $_.Current.Name -eq "りんご" }).Count -gt 0
                } | Out-Null
            }
            if ($true) {
                setGuiStep $S "バージョン情報"
                clickGui $S $S.Window "AboutLink" "バージョン情報"
                $about = waitGuiWindow $S "「バージョン情報」のダイアログ" -Id "VersionText"
                clickGui $S $about "CloseButton" "［OK］"
                waitGuiWindowClosed $S $about "「バージョン情報」"
            }
            if ($false) {
                setGuiStep $S "［インデックス管理］で［＋ フォルダを追加］→［キャンセル］"
                selectGuiTab $S "IndexTab" "NewIndexButton"
                clickGui $S $S.Window "NewIndexButton" "［＋ フォルダを追加］"
                $dialog = waitGuiWindow $S "インデックスの追加のダイアログ" -Id "FolderBox"
                clickGui $S $dialog "CancelButton" "［キャンセル］"
                waitGuiWindowClosed $S $dialog "追加のダイアログ"
            }
            closeGui $S
        }
    }
}
