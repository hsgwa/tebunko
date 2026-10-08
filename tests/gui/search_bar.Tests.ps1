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

Describe "検索条件の行の折り返し" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:tool = newGuiTool $TestDrive
        newGuiSampleIndex $script:tool $TestDrive
        $script:rect = {
            param ($S, $id)
            return (findGui $S.Window -Id $id).Current.BoundingRectangle
        }
    }

    It "広い幅では 1 行（高速検索の印は右端）、最小の幅では高速検索の印だけが 2 行目の左端に落ちる。項目の間は 8" {
        $S = startGui $script:tool "SearchFlow"
        $script:wide = $null
        $script:narrow = $null
        invokeGuiScene $S {
            resizeGuiWindow $S 1280 700
            $script:wide = waitGui $S "広い幅の並び" ${guiDefaultTimeout} {
                $fast = & $script:rect $S "FastSearchText"
                $regex = & $script:rect $S "RegexCheck"
                if ($fast.Width -gt 0 -and $fast.Left -gt $regex.Right) { @{ Fast = $fast; Regex = $regex; Scope = (& $script:rect $S "ScopeButton"); Excel = (& $script:rect $S "KindChipExcel"); Word = (& $script:rect $S "KindChipWord") } }
            }
            resizeGuiWindow $S 1024 700
            $script:narrow = waitGui $S "最小の幅の並び" ${guiDefaultTimeout} {
                $fast = & $script:rect $S "FastSearchText"
                $scope = & $script:rect $S "ScopeButton"
                if ($fast.Width -gt 0 -and $fast.Top -gt $scope.Bottom) { @{ Fast = $fast; Scope = $scope; Excel = (& $script:rect $S "KindChipExcel"); Regex = (& $script:rect $S "RegexCheck") } }
            }
            closeGui $S
        }
        # 広い幅: 高速検索の印は検索条件より右で、ファイル内の対象と同じ行
        $script:wide.Fast.Top | Should -BeLessThan ($script:wide.Scope.Bottom)
        # 項目の間は 8（チップの間）
        ($script:wide.Word.Left - $script:wide.Excel.Right) | Should -BeGreaterThan 7.5
        ($script:wide.Word.Left - $script:wide.Excel.Right) | Should -BeLessThan 8.5
        # 最小の幅: ファイル内の対象・検索条件は 1 行目に残り、高速検索の印は 2 行目の左端（種類の左）に落ちる
        $script:narrow.Regex.Top | Should -BeLessThan ($script:narrow.Scope.Bottom)
        $script:narrow.Fast.Left | Should -BeLessThan $script:narrow.Excel.Left
        # 行の間は 8（1 行目の下端から 2 行目の印の上端までは、印の上下の余白を含めて 8 以上）
        ($script:narrow.Fast.Top - $script:narrow.Scope.Bottom) | Should -BeGreaterThan 7
    }

    It "利用者の環境に触っていない" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}
