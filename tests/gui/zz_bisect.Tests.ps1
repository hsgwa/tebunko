# 【一時】終了コード 5 が、どの操作・どの終わり方で出るかを切り分ける（調べが済んだら消す）。
# 条件ごとに画面を 3 回ずつ起動して閉じる。落ちた条件は、場面の名前（条件の名前-回）と終了の記録で分かる。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"
}

Describe "終了コードの切り分け" -Tag Gui {
    BeforeAll {
        $script:toolEmpty = newGuiTool (Join-Path $TestDrive "empty")
        $script:toolIndex = newGuiTool (Join-Path $TestDrive "index")
        newGuiSampleIndex $script:toolIndex (Join-Path $TestDrive "index_root")
    }

    It "<Name>" -ForEach (@(
        foreach ($n in 1..3) {
            @{ Name = "A空-$n"; Index = $false; Search = $false; Preview = $false; About = $false; Idle = 0; Variant = "" }
            @{ Name = "B開閉-$n"; Index = $true; Search = $false; Preview = $false; About = $false; Idle = 0; Variant = "" }
            @{ Name = "C検索-$n"; Index = $true; Search = $true; Preview = $false; About = $false; Idle = 0; Variant = "" }
            @{ Name = "D行選択-$n"; Index = $true; Search = $true; Preview = $true; About = $false; Idle = 0; Variant = "" }
            @{ Name = "E全部-$n"; Index = $true; Search = $true; Preview = $true; About = $true; Idle = 0; Variant = "" }
            @{ Name = "F全部Exit-$n"; Index = $true; Search = $true; Preview = $true; About = $true; Idle = 0; Variant = "envexit" }
            @{ Name = "G全部Exiting-$n"; Index = $true; Search = $true; Preview = $true; About = $true; Idle = 0; Variant = "exiting" }
            @{ Name = "H待つ-$n"; Index = $true; Search = $false; Preview = $false; About = $false; Idle = 5; Variant = "" }
        }
    ) | Where-Object { !$env:TEBUNKO_BISECT_ONLY -or $_.Name -eq $env:TEBUNKO_BISECT_ONLY }) {
        $tool = if ($Index) { $script:toolIndex } else { $script:toolEmpty }
        $env:TEBUNKO_TEST_EXIT_VARIANT = $Variant
        try {
            $S = startGui $tool $Name
        } finally {
            $env:TEBUNKO_TEST_EXIT_VARIANT = $null
        }
        invokeGuiScene $S {
            if ($Search) {
                setGuiStep $S "検索"
                setGuiText $S (waitGuiById $S $S.Window "WordBox") "単価"
                clickGui $S $S.Window "SearchButton" "［検索］"
                waitGui $S "件数の表示（該当 2 件）" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "SummaryText")) -like "一致 2 件（*" } | Out-Null
            }
            if ($Preview) {
                setGuiStep $S "行を選ぶ"
                clickGui $S $S.Window "ExpandAllButton" "［すべて開く］"
                $row = waitGui $S "結果の行" ${guiDefaultTimeout} { getGuiHitRows (findGui $S.Window -Id "ResultGrid") | Select-Object -Last 1 }
                selectGui $row
                waitGui $S "プレビュー（りんご）" ${guiDefaultTimeout} {
                    @(findAllGui (findGui $S.Window -Id "PreviewScroll") -Type Text | Where-Object { $_.Current.Name -eq "りんご" }).Count -gt 0
                } | Out-Null
            }
            if ($About) {
                setGuiStep $S "バージョン情報"
                clickGui $S $S.Window "AboutLink" "バージョン情報"
                $about = waitGuiWindow $S "「バージョン情報」のダイアログ" -Id "VersionText"
                clickGui $S $about "CloseButton" "［OK］"
                waitGuiWindowClosed $S $about "「バージョン情報」"
            }
            if ($Idle -gt 0) { Start-Sleep -Seconds $Idle }
            closeGui $S
        }
    }
}
