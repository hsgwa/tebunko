# 【一時】Dispatcher を止める直しで、終了コード 5 が出なくなるかを確かめる（確かめが済んだら消す）。
# 条件ごとに画面を 12 回ずつ起動して閉じる。落ちた条件は、場面の名前（条件の名前-回）と終了の記録で分かる。
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
            foreach ($n in 1..12) {
                @{ Name = "V1開閉とバージョン情報-$n"; Search = $false }
                @{ Name = "V2検索と行とバージョン情報-$n"; Search = $true }
            }
        }
    ) | Where-Object { !$env:TEBUNKO_BISECT_ONLY -or $_.Name -eq $env:TEBUNKO_BISECT_ONLY }) {
        $tool = $script:toolIndex
        $S = startGui $tool $Name
        $doSearch = $Search
        invokeGuiScene $S {
            if ($doSearch) {
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
            setGuiStep $S "バージョン情報"
            clickGui $S $S.Window "AboutLink" "バージョン情報"
            $about = waitGuiWindow $S "「バージョン情報」のダイアログ" -Id "VersionText"
            clickGui $S $about "CloseButton" "［OK］"
            waitGuiWindowClosed $S $about "「バージョン情報」"
            closeGui $S
        }
    }
}
