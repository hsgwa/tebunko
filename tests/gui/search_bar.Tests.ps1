# 画面のスモークテスト: 検索条件の行（共通の関数は gui_helpers.ps1）。
# ［ファイル内の対象 ▾］は、既定から変えて「・N 件変更」が付いても幅が変わらない（176 で固定）。
# 高速検索の印は、使用可・使用不可、ⓘ の出し入れでも幅が変わらない（170 で固定）。
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

    # 窓の最小は 1024 なので、この幅は画面が小さくても必ず作れる
    It "最小の幅では、高速検索の印だけが 2 行目の左端に落ちる。項目の間は 8" {
        $S = startGui $script:tool "SearchFlowMin"
        $script:narrow = $null
        invokeGuiScene $S {
            resizeGuiWindow $S 1024 700
            $script:narrow = waitGui $S "最小の幅の並び" ${guiDefaultTimeout} {
                $fast = & $script:rect $S "FastSearchText"
                $scope = & $script:rect $S "ScopeButton"
                if ($fast.Width -gt 0 -and $fast.Top -gt $scope.Bottom) { @{ Fast = $fast; Scope = $scope; Label = (findGui $S.Window -Name "種類").Current.BoundingRectangle; Excel = (& $script:rect $S "KindChipExcel"); Word = (& $script:rect $S "KindChipWord"); Regex = (& $script:rect $S "RegexCheck") } }
            }
            # 正規表現の吹き出しを出しても、右端が窓の中に収まる
            toggleGui (waitGuiById $S $S.Window "RegexCheck")
            setGuiText $S (findGui $S.Window -Id "WordBox") "("
            $script:narrowBalloon = waitGui $S "最小の幅で吹き出しが出る" ${guiDefaultTimeout} {
                $text = findGui $S.Window -Id "RegexBalloonText"
                if ($text -and !$text.Current.IsOffscreen -and (getGuiText $text) -eq "正規表現が正しくありません") { @{ Text = $text.Current.BoundingRectangle; Window = $S.Window.Current.BoundingRectangle } }
            }
            closeGui $S
        }
        $script:narrowBalloon.Text.Right | Should -BeLessOrEqual $script:narrowBalloon.Window.Right -Because "最小の幅でも、吹き出しが窓の右端からはみ出さない"
        # ファイル内の対象・検索条件は 1 行目に残る
        $script:narrow.Regex.Top | Should -BeLessThan ($script:narrow.Scope.Bottom)
        # 高速検索の印は 2 行目の左端。印の字は欄（幅 170）の左に寄り、1 行目の左端の「種類」から 40 以内にある
        # （左寄せは 25 前後。右寄せのままなら、字は欄の右の端に寄って 46〜62 になる。しきい 40 は、その間（左寄せに 15、右寄せに 6 の余裕）に置いた）
        $diff = $script:narrow.Fast.Left - $script:narrow.Label.Left
        $diff | Should -BeGreaterThan 0
        $diff | Should -BeLessThan 40
        # 行の間は 8（1 行目の下端から 2 行目の印の上端までは、印の上下の余白を含めて 8 以上）
        ($script:narrow.Fast.Top - $script:narrow.Scope.Bottom) | Should -BeGreaterThan 7
        # 項目の間は 8（チップの間）
        ($script:narrow.Word.Left - $script:narrow.Excel.Right) | Should -BeGreaterThan 7.5
        ($script:narrow.Word.Left - $script:narrow.Excel.Right) | Should -BeLessThan 8.5
    }

    # 1 行に収まる幅（1100 以上）を作るため、窓を 1280×700 にする。画面の作業領域がそれより小さいと窓が切られて作れないので、その画面では飛ばす。
    # 収まる・収まらないの境目は、getConditionFlow の単体テストで確かめてある
    It "広い幅では 1 行になり、高速検索の印は右端に付く" {
        $wideWidth = 1280
        $wideHeight = 700
        Add-Type -AssemblyName System.Windows.Forms
        $area = [Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        if ($area.Width -lt $wideWidth -or $area.Height -lt $wideHeight) {
            Set-ItResult -Skipped -Because "画面の作業領域が $($area.Width)×$($area.Height) で、広い幅の窓を作れない"
            return
        }
        $S = startGui $script:tool "SearchFlowWide"
        $script:wide = $null
        invokeGuiScene $S {
            resizeGuiWindow $S $wideWidth $wideHeight
            $actual = $S.Window.Current.BoundingRectangle
            if ($actual.Width -lt $wideWidth - 20) { throw "窓を頼んだ幅（$wideWidth）にできなかった（実際は $($actual.Width)。作業領域は $($area.Width)×$($area.Height)）" }
            $script:wide = waitGui $S "広い幅の並び" ${guiDefaultTimeout} {
                $fast = & $script:rect $S "FastSearchText"
                $regex = & $script:rect $S "RegexCheck"
                if ($fast.Width -gt 0 -and $fast.Left -gt $regex.Right) { @{ Fast = $fast; Scope = (& $script:rect $S "ScopeButton") } }
            }
            closeGui $S
        }
        # 高速検索の印は検索条件より右で、ファイル内の対象と同じ行
        $script:wide.Fast.Top | Should -BeLessThan ($script:wide.Scope.Bottom)
    }

    # 正規表現の吹き出しを出したまま窓を広げると、［正規表現］だけが右へ動く。吹き出しはそれについていく（左端が［正規表現］の左端にそろったまま、右端が窓の中）
    It "正規表現の吹き出しを出したまま窓を広げても、［正規表現］の下にそろう" {
        $wideWidth = 1280
        $wideHeight = 700
        Add-Type -AssemblyName System.Windows.Forms
        $area = [Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        if ($area.Width -lt $wideWidth -or $area.Height -lt $wideHeight) {
            Set-ItResult -Skipped -Because "画面の作業領域が $($area.Width)×$($area.Height) で、広い幅の窓を作れない"
            return
        }
        $S = startGui $script:tool "SearchBalloonWide"
        $script:balloon = $null
        invokeGuiScene $S {
            toggleGui (waitGuiById $S $S.Window "RegexCheck")
            setGuiText $S (findGui $S.Window -Id "WordBox") "("
            waitGui $S "吹き出しが出る" ${guiDefaultTimeout} {
                $text = findGui $S.Window -Id "RegexBalloonText"
                $text -and !$text.Current.IsOffscreen -and (getGuiText $text) -eq "正規表現が正しくありません"
            } | Out-Null
            # いったん最小の幅にして、高速検索の印が 2 行目に落ちるのを待つ
            resizeGuiWindow $S 1024 700
            $before = waitGui $S "最小の幅の並び" ${guiDefaultTimeout} {
                $fast = & $script:rect $S "FastSearchText"
                $scope = & $script:rect $S "ScopeButton"
                if ($fast.Width -gt 0 -and $fast.Top -gt $scope.Bottom) { & $script:rect $S "RegexCheck" }
            }
            resizeGuiWindow $S $wideWidth $wideHeight
            # 幅が広がると高速検索の印が 1 行目に戻り、［正規表現］は右へ動く
            waitGui $S "高速検索の印が 1 行目に戻る" ${guiDefaultTimeout} {
                $fast = & $script:rect $S "FastSearchText"
                $scope = & $script:rect $S "ScopeButton"
                $fast.Width -gt 0 -and $fast.Top -le $scope.Bottom
            } | Out-Null
            (& $script:rect $S "RegexCheck").Left | Should -BeGreaterThan ($before.Left + 10) -Because "広げると［正規表現］は右へ動く"
            # 文字の左は、吹き出しの左端から枠・余白・アイコンの分（29 前後）だけ右
            $script:balloon = waitGui $S "吹き出しが［正規表現］の下にそろう" ${guiDefaultTimeout} {
                $regex = & $script:rect $S "RegexCheck"
                $text = & $script:rect $S "RegexBalloonText"
                $gap = $text.Left - $regex.Left
                if ($gap -ge 24 -and $gap -le 34) { @{ Gap = $gap; Right = $text.Right; Window = $S.Window.Current.BoundingRectangle.Right } }
            }
            closeGui $S
        }
        $script:balloon.Right | Should -BeLessOrEqual $script:balloon.Window -Because "吹き出しが窓の右端からはみ出さない"
    }

    It "利用者の環境に触っていない" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}
