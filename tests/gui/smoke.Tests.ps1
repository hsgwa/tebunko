# 画面のスモークテスト S1: 起動・検索・閉じる（本物の画面を UI オートメーションで操作する。共通の関数は gui_helpers.ps1）。
# 画面遷移の一覧（docs\design\testing\gui-smoke.md「画面のスモークテスト」）の #1・#2・#6・#7・#8・#9 を確かめる。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"
}

Describe "S1 起動・検索・閉じる" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:tool = newGuiTool $TestDrive
        newGuiSampleIndex $script:tool $TestDrive
    }

    It "起動して［2 検索］が選ばれ、3 つのタブ・検索・プレビュー・「tebunko について」・多重起動・閉じるが動く" {
        $S = startGui $script:tool "S1"
        invokeGuiScene $S {
            # 起動・インデックスがあれば［2 検索］が選ばれる（#1・#2）
            setGuiStep $S "起動時のタブ"
            getGuiSelectedTab $S | Should -Be "SearchTab"

            # 3 つの画面を選ぶ。選んだ画面の中の部品が UI オートメーションに出る（ContentHost の中身）（#7）
            foreach ($tab in @(
                @{ Id = "IndexTab"; Content = "IndexGrid" }, @{ Id = "SettingsTab"; Content = "ChangeWorkspaceButton" },
                @{ Id = "SearchTab"; Content = "WordBox" })) {
                setGuiStep $S "タブ $($tab.Id) を選ぶ"
                selectGuiTab $S $tab.Id $tab.Content
            }

            # 検索して、結果の行を選ぶとプレビューが出る（#23）
            setGuiStep $S "検索ワードを入れて［検索］"
            setGuiText $S (waitGuiById $S $S.Window "WordBox") "単価"
            $sw = [Diagnostics.Stopwatch]::StartNew()
            clickGui $S $S.Window "SearchButton" "［検索］"
            setGuiStep $S "検索の結果（該当 2 件）"
            waitGui $S "件数の表示（該当 2 件）" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "SummaryText")) -like "該当 2 件*" } | Out-Null
            $S.Timing["検索"] = [Math]::Round($sw.Elapsed.TotalSeconds, 1)
            clickGui $S $S.Window "ExpandAllButton" "［すべて展開］"
            setGuiStep $S "結果の行を選んでプレビュー"
            $row = waitGui $S "結果の行" ${guiDefaultTimeout} { getGuiHitRows (findGui $S.Window -Id "ResultGrid") | Select-Object -Last 1 }
            selectGui $row
            waitGui $S "プレビュー（りんご）" ${guiDefaultTimeout} {
                @(findAllGui (findGui $S.Window -Id "PreviewScroll") -Type Text | Where-Object { $_.Current.Name -eq "りんご" }).Count -gt 0
            } | Out-Null
            (getGuiText (findGui $S.Window -Id "DetailTitle")) | Should -BeLike "*見積.xlsx*"

            # 左の欄の「tebunko について」（#8）
            setGuiStep $S "「tebunko について」を開く"
            clickGui $S $S.Window "AboutLink" "tebunko について"
            $about = waitGuiWindow $S "「tebunko について」のダイアログ" -Id "VersionText"
            (getGuiText (findGui $about -Id "VersionText")) | Should -BeLike "バージョン *"
            clickGui $S $about "CloseButton" "［OK］"
            waitGuiWindowClosed $S $about "「tebunko について」"

            # 同じフォルダのツールをもう一度起動すると、2 つ目はすぐ終わり、1 つ目は残る（#6）
            setGuiStep $S "多重起動"
            $second = startGuiProcess $script:tool
            $S.Extra += $second
            waitGui $S "2 つ目の起動が終わる" ${guiDefaultTimeout} { $second.HasExited } | Out-Null
            $second.ExitCode | Should -Be 0
            $S.Process.HasExited | Should -BeFalse

            # 閉じると終了コード 0（#9）
            closeGui $S
            Write-Host ("S1 の秒数: " + (($S.Timing.GetEnumerator() | ForEach-Object { "$($_.Key) $($_.Value)" }) -join "・"))
        }
    }

    It "利用者の環境（作業ツリーの設定・work\content_index・LOCALAPPDATA・既定のワークスペース・Office のプロセス）に触っていない" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}

Describe "S8 単一 .ps1 版: 起動・検索・閉じる" -Tag Gui {
    BeforeAll {
        $script:envBeforeSingle = getGuiEnvSnapshot
        $script:singleTool = newGuiSingleScriptTool $TestDrive
        newGuiSampleIndex $script:singleTool $TestDrive "single"
    }

    It "起動して［2 検索］が選ばれ、検索で当たり、閉じると終了コード 0" {
        $S = startGui $script:singleTool "S8"
        invokeGuiScene $S {
            setGuiStep $S "起動時のタブ"
            getGuiSelectedTab $S | Should -Be "SearchTab"

            setGuiStep $S "検索ワードを入れて［検索］"
            setGuiText $S (waitGuiById $S $S.Window "WordBox") "単価"
            clickGui $S $S.Window "SearchButton" "［検索］"
            setGuiStep $S "検索の結果（該当 2 件）"
            waitGui $S "件数の表示（該当 2 件）" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "SummaryText")) -like "該当 2 件*" } | Out-Null

            closeGui $S
        }
    }

    It "利用者の環境に触っていない" {
        compareGuiEnvSnapshot $script:envBeforeSingle (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}

Describe "S1 同梱のフォント" -Tag Gui {
    It "配布する形（scripts\ の写し）に同梱のフォントがあり、Font.Body がそこから作られる" -Skip:($env:TEBUNKO_GUI_SINGLE -eq "1") {
        $tool = newGuiTool $TestDrive
        $fonts = "$($tool.Dir)\scripts\shared\fonts"
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
        . "$($tool.Dir)\scripts\shared\ui\app_host.ps1"
        $family = newAppFontFamily $fonts
        $family.BaseUri.LocalPath | Should -Be "$fonts\"
        @([System.Windows.Media.Fonts]::GetFontFamilies($family.BaseUri) | ForEach-Object { $_.FamilyNames.Values }) | Should -Contain "Rethink Sans"
    }
}
