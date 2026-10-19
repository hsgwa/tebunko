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

    # 窓の最小は 1024 なので、この幅は画面が小さくても必ず作れる
    It "最小の幅でも、検索条件は 1 行に収まる。項目の間は 8" {
        $S = startGui $script:tool "SearchFlowMin"
        $script:narrow = $null
        invokeGuiScene $S {
            resizeGuiWindow $S 1024 700
            $script:narrow = waitGui $S "最小の幅の並び" ${guiDefaultTimeout} {
                $scope = & $script:rect $S "ScopeButton"
                if ($scope.Width -gt 0) { @{ Window = $S.Window.Current.BoundingRectangle; Scope = $scope; Label = (findGui $S.Window -Name "種類").Current.BoundingRectangle; Excel = (& $script:rect $S "KindChipExcel"); Word = (& $script:rect $S "KindChipWord"); Regex = (& $script:rect $S "RegexCheck") } }
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
        # ファイル内の対象・検索条件も同じ 1 行に残る
        $script:narrow.Regex.Top | Should -BeLessThan ($script:narrow.Scope.Bottom)
        $script:narrow.Regex.Right | Should -BeLessOrEqual $script:narrow.Window.Right
        # 項目の間は 8（チップの間）
        ($script:narrow.Word.Left - $script:narrow.Excel.Right) | Should -BeGreaterThan 7.5
        ($script:narrow.Word.Left - $script:narrow.Excel.Right) | Should -BeLessThan 8.5
    }

    # 吹き出しを出し入れ（式を直して消し、また誤らせて出す）しても、位置は変わらない。最小の幅（吹き出しの右端が窓の右端に近い）で確かめる
    It "最小の幅で、正規表現の吹き出しを出し入れしても、［正規表現］の下の同じ位置に出る" {
        $S = startGui $script:tool "SearchBalloonBadge"
        $script:pos = @()
        invokeGuiScene $S {
            resizeGuiWindow $S 1024 700
            toggleGui (waitGuiById $S $S.Window "RegexCheck")
            $measure = {
                $regex = & $script:rect $S "RegexCheck"
                $text = & $script:rect $S "RegexBalloonText"
                @{ Gap = $text.Left - $regex.Left; Right = $text.Right; Window = $S.Window.Current.BoundingRectangle.Right; Below = $text.Top - $regex.Bottom }
            }
            foreach ($word in @("(", "(a", "ab")) {
                setGuiText $S (findGui $S.Window -Id "WordBox") $word
                if ($word -eq "ab") {
                    # 正規表現として正しい語。吹き出しは消える
                    waitGui $S "吹き出しが消える" ${guiDefaultTimeout} {
                        $text = findGui $S.Window -Id "RegexBalloonText"
                        !$text -or $text.Current.IsOffscreen
                    } | Out-Null
                    setGuiText $S (findGui $S.Window -Id "WordBox") "("
                }
                $script:pos += waitGui $S "吹き出しが出る（$word）" ${guiDefaultTimeout} {
                    $text = findGui $S.Window -Id "RegexBalloonText"
                    if ($text -and !$text.Current.IsOffscreen -and (getGuiText $text) -eq "正規表現が正しくありません") { & $measure }
                }
            }
            closeGui $S
        }
        foreach ($p in $script:pos) {
            # 最小の幅では、右端が窓に収まるまで左へ寄る（左端が［正規表現］より左になる）ので、左端の差は決めず、出し入れで変わらないことを見る
            $p.Below | Should -BeGreaterThan 0 -Because "吹き出しは［正規表現］の下に出る"
            $p.Right | Should -BeLessOrEqual $p.Window
        }
        ($script:pos | ForEach-Object { [math]::Round($_.Gap) } | Select-Object -Unique).Count | Should -Be 1 -Because "出し入れのあとも左端の位置が変わらない"
    }

    # 1 行に収まる幅（1100 以上）を作るため、窓を 1280×700 にする。画面の作業領域がそれより小さいと窓が切られて作れないので、その画面では飛ばす。
    # GitHub のホストランナー（windows-latest）の画面は 1024×768 前後で、ここで飛ばされる（CI の gui.yml では画面の大きさを変えていない）
    # 収まる・収まらないの境目は、getConditionFlow の単体テストで確かめてある
    It "広い幅では 1 行になり、検索条件は右端に寄る" {
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
                $regex = & $script:rect $S "RegexCheck"
                $scope = & $script:rect $S "ScopeButton"
                if ($regex.Width -gt 0 -and $regex.Left -gt $scope.Right) { @{ Regex = $regex; Scope = $scope } }
            }
            closeGui $S
        }
        # 検索条件はファイル内の対象より右で、同じ行
        $script:wide.Regex.Top | Should -BeLessThan ($script:wide.Scope.Bottom)
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
            # いったん最小の幅にして、並びが決まるのを待つ
            resizeGuiWindow $S 1024 700
            $before = waitGui $S "最小の幅の並び" ${guiDefaultTimeout} {
                $regex = & $script:rect $S "RegexCheck"
                if ($regex.Width -gt 0) { $regex }
            }
            resizeGuiWindow $S $wideWidth $wideHeight
            # 幅が広がると、［正規表現］は右へ動く
            waitGui $S "［正規表現］が右へ動く" ${guiDefaultTimeout} {
                (& $script:rect $S "RegexCheck").Left -gt ($before.Left + 10)
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
