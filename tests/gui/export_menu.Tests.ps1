# 画面のスモークテスト: 検索結果の［出力 ▾］のメニュー（共通の関数は gui_helpers.ps1）。
# ［出力 ▾］を押すと、Excel・CSV・TSV の 3 項目が開き、既定の種類（Excel）に印が付く。ファイルは出力しない。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"
}

Describe "検索結果の［出力 ▾］のメニュー" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:tool = newGuiTool $TestDrive
        newGuiSampleIndex $script:tool $TestDrive
    }

    It "［出力 ▾］を押すと 3 つの種類が並び、Excel に印が付いている" {
        $S = startGui $script:tool "ExportMenu1"
        $script:states = $null
        $script:names = $null
        invokeGuiScene $S {
            setGuiText $S (findGui $S.Window -Id "WordBox") "単価"
            waitGuiEnabled $S (findGui $S.Window -Id "SearchButton") "［検索］"
            clickGui $S $S.Window "SearchButton" "［検索］"
            waitGui $S "該当 2 件" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "SummaryText")) -like "一致 2 件（*" } | Out-Null

            clickGui $S $S.Window "ExportMenuButton" "［出力］の［▾］"
            $items = waitGui $S "メニューが開く" ${guiDefaultTimeout} {
                $found = @(getGuiOtherWindows $S | ForEach-Object { findAllGui $_ -Type MenuItem })
                if ($found.Count -ge 3) { , $found }
            }
            $script:names = @($items | ForEach-Object { $_.Current.Name })
            $script:states = [ordered]@{}
            foreach ($item in $items) {
                $script:states[$item.Current.Name] = getGuiToggleState $item
            }

            # メニューを閉じる（Esc）。閉じないと、画面を閉じる手順が進まない
            foreach ($window in @(getGuiOtherWindows $S)) {
                pressGuiKey $window 0x1B
            }
            waitGui $S "メニューが閉じる" ${guiDefaultTimeout} { @(getGuiOtherWindows $S).Count -eq 0 } | Out-Null
            closeGui $S
        }
        ($script:names -join "/") | Should -Be "Excel ブック (.xlsx)/CSV (.csv)/TSV (.tsv)"
        $script:states["Excel ブック (.xlsx)"] | Should -Be "On"
        $script:states["CSV (.csv)"] | Should -Be "Off"
        $script:states["TSV (.tsv)"] | Should -Be "Off"
    }

    It "利用者の環境に触っていない" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}
