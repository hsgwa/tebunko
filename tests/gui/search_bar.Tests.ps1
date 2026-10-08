# 画面のスモークテスト: 検索条件の行（共通の関数は gui_helpers.ps1）。
# ［ファイル内の対象 ▾］は、既定から変えて「・N 件変更」が付いても幅が変わらない（176 で固定）。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"
}

Describe "検索条件の［ファイル内の対象 ▾］の幅" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:tool = newGuiTool $TestDrive
        newGuiSampleIndex $script:tool $TestDrive
        $script:measure = {
            param ($S)
            $button = waitGuiById $S $S.Window "ScopeButton"
            $rect = $button.Current.BoundingRectangle
            return @{ Name = $button.Current.Name; Width = $rect.Width }
        }
    }

    It "既定でも、図形・コメントを外して「・2 件変更」が付いても、同じ幅（176）になる" {
        $S = startGui $script:tool "SearchBar1"
        $script:defaultScope = $null
        invokeGuiScene $S {
            $script:defaultScope = & $script:measure $S
            closeGui $S
        }
        # 設定ファイルで図形・コメントを外した状態で起動し直す（メニューは UI オートメーションで開けないため）
        $config = [ordered]@{ workspaceFolder = $script:tool.Work; includeShapes = $false; includeComments = $false }
        writeGuiConfig $script:tool.Dir $config
        $S = startGui $script:tool "SearchBar2"
        $script:changedScope = $null
        invokeGuiScene $S {
            $script:changedScope = & $script:measure $S
            closeGui $S
        }
        $script:defaultScope.Width | Should -Be 176
        $script:changedScope.Width | Should -Be 176
    }

    It "利用者の環境に触っていない" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}
