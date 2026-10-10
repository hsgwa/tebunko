# 【一時】Dispatcher を止める直しで、終了コード 5 が出なくなるかを確かめる（確かめが済んだら消す）。
# 条件ごとに画面を 50 回ずつ起動して閉じる。落ちた条件は、場面の名前（条件の名前-回）と終了の記録で分かる。
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
            foreach ($n in 1..50) {
                @{ Name = "V1開閉とバージョン情報-$n"; Op = "about"; Variant = "exiting" }
                @{ Name = "C2Exit-$n"; Op = "about"; Variant = "exiting,envexit" }
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
        $doOp = $Op
        invokeGuiScene $S {
            if ($doOp -eq "about") {
                setGuiStep $S "バージョン情報"
                clickGui $S $S.Window "AboutLink" "バージョン情報"
                $about = waitGuiWindow $S "「バージョン情報」のダイアログ" -Id "VersionText"
                clickGui $S $about "CloseButton" "［OK］"
                waitGuiWindowClosed $S $about "「バージョン情報」"
            } else {
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
