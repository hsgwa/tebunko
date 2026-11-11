# 画面のスモークテスト: インデックスが多いとき（200 件）に、一覧・詳細欄の切り替え・検索対象のツリーが決めた秒数以内に応答する。
# 秒数は GitHub Actions の Windows の実行機で、通常の 5 倍ほどの余裕を見て決めた上限（遅くなったことに気づくための目安で、速さの計測ではない）。
# 流すのは Gui タグ（gui.yml）。共通の関数は gui_helpers.ps1。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"
}

Describe "S2b インデックスが 200 件ある" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:indexCount = 200
        $script:limitSeconds = 15
        $script:tool = newGuiTool $TestDrive
        $folders = foreach ($i in 1..$script:indexCount) {
            @{ name = ("インデックス{0:000}" -f $i); path = (Join-Path $TestDrive ("元のフォルダ\{0:000}" -f $i)); enabled = $true }
        }
        $config = readGuiConfig $script:tool
        $config | Add-Member -NotePropertyName targetFolders -NotePropertyValue @($folders) -Force
        writeGuiConfig $script:tool.Dir $config
    }

    It "一覧・詳細欄の切り替え・検索対象のツリーが上限の秒数以内に応答する" {
        $S = startGui $script:tool "S2b"
        invokeGuiScene $S {
            setGuiStep $S "一覧が出る"
            $sw = [Diagnostics.Stopwatch]::StartNew()
            $row = waitGui $S "一覧の先頭の行" ${guiDefaultTimeout} { @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")) | Select-Object -First 1 }
            $S.Timing["一覧"] = [Math]::Round($sw.Elapsed.TotalSeconds, 1)
            $sw.Elapsed.TotalSeconds | Should -BeLessThan $script:limitSeconds
            waitGui $S "登録のフォルダ数" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "IndexFooterFolders")) -like "*$($script:indexCount)*" } | Out-Null

            setGuiStep $S "詳細欄の切り替え"
            selectGui $row
            waitGui $S "詳細の見出し" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "IndexDetailTitle")) -like "インデックス*- 詳細" } | Out-Null
            $rows = @(getGuiGridRows (findGui $S.Window -Id "IndexGrid"))
            $rows.Count | Should -BeGreaterThan 1
            $other = $rows[$rows.Count - 1]
            $name = (getGuiRowTexts $other) | Where-Object { $_ -like "インデックス*" } | Select-Object -First 1
            $sw.Restart()
            selectGui $other
            waitGui $S "別の行の詳細の見出し" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "IndexDetailTitle")) -eq "$name - 詳細" } | Out-Null
            $S.Timing["詳細の切り替え"] = [Math]::Round($sw.Elapsed.TotalSeconds, 1)
            $sw.Elapsed.TotalSeconds | Should -BeLessThan $script:limitSeconds

            setGuiStep $S "検索対象のツリー"
            $sw.Restart()
            selectGuiTab $S "SearchTab" "IndexTree"
            waitGui $S "検索対象のツリー" ${guiDefaultTimeout} { findGui $S.Window -Id "IndexTree" } | Out-Null
            $S.Timing["ツリー"] = [Math]::Round($sw.Elapsed.TotalSeconds, 1)
            $sw.Elapsed.TotalSeconds | Should -BeLessThan $script:limitSeconds

            closeGui $S
            Write-Host ("S2b の秒数: " + (($S.Timing.GetEnumerator() | ForEach-Object { "$($_.Key) $($_.Value)" }) -join "・"))
        }
    }

    It "利用者の環境に触っていない（Office のプロセスの数も変わらない）" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}

# 測る件数。既定は 500 件だけ（50・200 件との比べは TEBUNKO_GUI_STALL_COUNTS="50,200,500" で流す）
$stallCounts = @(($(if ($env:TEBUNKO_GUI_STALL_COUNTS) { $env:TEBUNKO_GUI_STALL_COUNTS } else { "500" }) -split ",") | ForEach-Object { @{ Count = [int]$_ } })

Describe "S2c インデックスが <count> 件あるときの、起動後の止まり" -Tag Gui -ForEach $stallCounts {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:indexCount = $count
        # 起動後に画面が応答しなかった一番長い時間の上限（秒）。インデックスの読み込みを画面のスレッドでしていた版は、件数に比例して長く止まった
        $script:stallLimit = $(if ($env:TEBUNKO_GUI_STALL_LIMIT) { [double]$env:TEBUNKO_GUI_STALL_LIMIT } else { 8.0 })
        $script:tool = newGuiTool $TestDrive
        $utf8 = New-Object Text.UTF8Encoding($true)
        $folders = foreach ($i in 1..$script:indexCount) {
            $name = "インデックス{0:000}" -f $i
            # 本文インデックス（1 つのインデックスの下に、ファイルの 1 つ分の TSV）を作り、検索対象のツリーに並ぶようにする
            $dir = "$($script:tool.Work)\store\content\$name\見積.xlsx"
            [void][IO.Directory]::CreateDirectory($dir)
            [IO.File]::WriteAllText("$dir\Sheet1.tsv", "`t見積`r`n", $utf8)
            @{ name = $name; path = (Join-Path $TestDrive ("元のフォルダ\{0:000}" -f $i)); enabled = $true }
        }
        $config = readGuiConfig $script:tool
        $config | Add-Member -NotePropertyName targetFolders -NotePropertyValue @($folders) -Force
        writeGuiConfig $script:tool.Dir $config
    }

    It "窓が開いてから一覧・ツリーがそろうまで、画面が上限の秒数を超えて応答しなくなることはない" {
        $S = startGui $script:tool "S2c-$($script:indexCount)"
        invokeGuiScene $S {
            setGuiStep $S "一覧と検索対象のツリーがそろうまでの止まりを測る"
            $stall = measureGuiStall $S {
                (getGuiText (findGui $S.Window -Id "IndexFooterFolders")) -like "*$($script:indexCount)*"
            }
            setGuiStep $S "検索対象のツリーにも入ったか"
            selectGuiTab $S "SearchTab" "IndexTree"
            waitGui $S "検索対象の数" ${guiDefaultTimeout} {
                (getGuiText (findGui $S.Window -Id "TargetCountText")) -eq "検索対象 $($script:indexCount) / $($script:indexCount)"
            } | Out-Null
            $S.Timing["一番長い止まり"] = $stall.Longest
            $S.Timing["そろうまで"] = $stall.Total

            closeGui $S
            Write-Host ("S2c（{0} 件）: 一番長い止まり {1} 秒・そろうまで {2} 秒（問い合わせ {3} 回）" -f $script:indexCount, $stall.Longest, $stall.Total, $stall.Queries)
            $stall.Longest | Should -BeLessThan $script:stallLimit
        } -NoWait
    }

    It "利用者の環境に触っていない（Office のプロセスの数も変わらない）" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}

# S13 は、インデックスが多く（既定 500 件）、ヒットしたファイルも多い（9,500 件）ときの、操作ごとの画面の止まり。
# 操作 1 回のあとに画面が応答しなかった一番長い時間を測る（measureGuiActionStall）。S2c は起動の止まりだけで、起動後の操作は測れない。
# 上限（stallLimits）の決め方は docs/design/testing/perf-check.md の「上限」。TEBUNKO_GUI_STALL_LIMIT があれば全場面をその値にする（手元の遅い PC 用）
Describe "S13 インデックスが <count> 件・ヒット約 1 万件のときの、操作ごとの止まり" -Tag Gui -ForEach $stallCounts {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:indexCount = $count
        $script:perIndex = [Math]::Floor(9500 / $script:indexCount)
        $script:hitTotal = $script:perIndex * $script:indexCount
        # 場面ごとの上限（秒。仮の値: 手元の測りの 2 倍ほど。CI の 5 回の測りで決め直す）。A: ツリーのチェック・B: 前面復帰・D: 絞り込みの 1 文字・E: インデックス一覧の行の選択
        $script:stallLimits = @{ A = 1.0; B = 10.0; D = 25.0; E = 5.0 }
        if ($env:TEBUNKO_GUI_STALL_LIMIT) { foreach ($key in @($script:stallLimits.Keys)) { $script:stallLimits[$key] = [double]$env:TEBUNKO_GUI_STALL_LIMIT } }
        $script:tool = newGuiTool $TestDrive
        # インデックス N 件に、同じ語を 1 行だけ含むファイルを perIndex 件ずつ。ファイル名は「見積<番号>」（数字と漢字だけ）
        $tsvRoot = Join-Path $TestDrive "tsv"
        foreach ($i in 1..$script:indexCount) {
            foreach ($j in 1..$script:perIndex) {
                $file = "見積{0:000}{1:000}" -f $i, $j
                newTsv "$tsvRoot\インデックス$("{0:000}" -f $i)\$file.xlsx\$(toIndexFileName "Sheet1")" @("品名`t数", "試験の語`t1")
            }
        }
        [void](newContentIndexFiles $tsvRoot "$($script:tool.Work)\store\content")
        $folders = foreach ($i in 1..$script:indexCount) {
            @{ name = ("インデックス{0:000}" -f $i); path = (Join-Path $TestDrive ("元のフォルダ\{0:000}" -f $i)); enabled = $true }
        }
        $config = readGuiConfig $script:tool
        $config | Add-Member -NotePropertyName targetFolders -NotePropertyValue @($folders) -Force
        $script:baseConfig = $config
    }

    BeforeEach {
        # 操作のたびに、検索対象の外し方などが設定に残らないよう、最初の設定に戻す
        writeGuiConfig $script:tool.Dir $script:baseConfig
    }

    It "E インデックス一覧の行を選んで、詳細が替わるまで" {
        $S = startGui $script:tool "S13E-$($script:indexCount)"
        invokeGuiScene $S {
            setGuiStep $S "一覧の先頭の行を選ぶ"
            selectGuiTab $S "IndexTab" "IndexGrid"
            $first = waitGui $S "一覧の先頭の行" ${guiDefaultTimeout} { @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")) | Select-Object -First 1 }
            waitGui $S "登録のフォルダ数" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "IndexFooterFolders")) -like "*$($script:indexCount)*" } | Out-Null
            selectGui $first
            waitGui $S "詳細の見出し" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "IndexDetailTitle")) -like "インデックス*- 詳細" } | Out-Null
            # 一覧は仮想化されるので、取れる最後の行は一覧の本当の最後とは限らない（S2b と同じ）。取れた最後の行を選ぶ
            $last = @(getGuiGridRows (findGui $S.Window -Id "IndexGrid"))[-1]
            $name = (getGuiRowTexts $last) | Where-Object { $_ -like "インデックス*" } | Select-Object -First 1
            setGuiStep $S "最後の行を選ぶ止まりを測る"
            $stall = measureGuiActionStall $S -Argument $last -Action {
                param ($element)
                $element.GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern).Select()
            } -Done { (getGuiText (findGui $S.Window -Id "IndexDetailTitle")) -eq "$name - 詳細" }
            closeGui $S
            Write-Host ("S13 E（{0} 件）: 一番長い止まり {1} 秒（{2} 秒・問い合わせ {3} 回）" -f $script:indexCount, $stall.Longest, $stall.Total, $stall.Queries)
            $stall.Longest | Should -BeLessThan $script:stallLimits.E
        }
    }

    It "A 検索対象のツリーのチェック 1 回" {
        $S = startGui $script:tool "S13A-$($script:indexCount)"
        invokeGuiScene $S {
            selectGuiTab $S "SearchTab" "IndexTree"
            waitGui $S "検索対象の数" ${guiDefaultTimeout} {
                (getGuiText (findGui $S.Window -Id "TargetCountText")) -eq "検索対象 $($script:indexCount) / $($script:indexCount)"
            } | Out-Null
            clickGui $S $S.Window "UncheckAllIndexButton" "［解除］"
            waitGui $S "検索対象が 0 になる" ${guiDefaultTimeout} {
                (getGuiText (findGui $S.Window -Id "TargetCountText")) -eq "検索対象 0 / $($script:indexCount)"
            } | Out-Null
            $root = waitGui $S "ツリーの根" ${guiDefaultTimeout} { findGui (findGui $S.Window -Id "IndexTree") -Type TreeItem }
            # 根を選んで Space を送る（部品の TogglePattern は Click を起こさず、検索対象に反映されない）。キーは窓へ送る（WM_KEYDOWN・WM_KEYUP）
            selectGui $root
            $root.SetFocus()
            setGuiStep $S "根を 1 つチェックする止まりを測る"
            $stall = measureGuiActionStall $S -Argument ([int]$S.Window.Current.NativeWindowHandle) -Action {
                param ($handle)
                [void][TebunkoGuiNative]::PostMessage([IntPtr]$handle, 0x0100, [IntPtr]0x20, [IntPtr]0)
                [void][TebunkoGuiNative]::PostMessage([IntPtr]$handle, 0x0101, [IntPtr]0x20, [IntPtr]0)
            } -Done { (getGuiText (findGui $S.Window -Id "TargetCountText")) -eq "検索対象 1 / $($script:indexCount)" }
            closeGui $S
            Write-Host ("S13 A（{0} 件）: 一番長い止まり {1} 秒（{2} 秒・問い合わせ {3} 回）" -f $script:indexCount, $stall.Longest, $stall.Total, $stall.Queries)
            $stall.Longest | Should -BeLessThan $script:stallLimits.A
        }
    }

    It "B 窓を最小化して戻した（前面に出た）あと" {
        $S = startGui $script:tool "S13B-$($script:indexCount)"
        invokeGuiScene $S {
            waitGui $S "登録のフォルダ数" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "IndexFooterFolders")) -like "*$($script:indexCount)*" } | Out-Null
            setGuiStep $S "最小化して戻す止まりを測る"
            $stall = measureGuiActionStall $S -Argument $S.Window -Settle 1.5 -Action {
                param ($element)
                $pattern = $element.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern)
                $pattern.SetWindowVisualState([Windows.Automation.WindowVisualState]::Minimized)
                Start-Sleep -Milliseconds 500
                $pattern.SetWindowVisualState([Windows.Automation.WindowVisualState]::Normal)
            } -Done {
                param ($actionDone)
                $actionDone -and [string]$S.Window.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern).Current.WindowVisualState -eq "Normal"
            }
            closeGui $S
            Write-Host ("S13 B（{0} 件）: 一番長い止まり {1} 秒（{2} 秒・問い合わせ {3} 回）" -f $script:indexCount, $stall.Longest, $stall.Total, $stall.Queries)
            $stall.Longest | Should -BeLessThan $script:stallLimits.B
        }
    }

    It "D 結果の絞り込み欄に 1 文字打った（ヒット約 1 万件）" {
        $S = startGui $script:tool "S13D-$($script:indexCount)"
        invokeGuiScene $S {
            selectGuiTab $S "SearchTab" "WordBox"
            setGuiText $S (findGui $S.Window -Id "WordBox") "試験の語"
            waitGuiEnabled $S (findGui $S.Window -Id "SearchButton") "検索"
            clickGui $S $S.Window "SearchButton" "検索"
            $hits = $script:hitTotal.ToString("N0")
            $searchClock = [Diagnostics.Stopwatch]::StartNew()
            try {
                waitGui $S "一致 $hits 件" 300 { (getGuiText (findGui $S.Window -Id "SummaryText")) -like "一致 $hits 件（$hits ファイル）*" } | Out-Null
            } catch {
                throw ("{0}（要約欄の最後の文: 「{1}」）" -f $_.Exception.Message, (getGuiText (findGui $S.Window -Id "SummaryText")))
            }
            Write-Host ("S13 D（{0} 件）: 検索が終わるまで {1} 秒" -f $script:indexCount, [Math]::Round($searchClock.Elapsed.TotalSeconds, 1))
            $box = findGui $S.Window -Id "FilterBox"
            setGuiStep $S "絞り込みに 1 文字打つ止まりを測る"
            $stall = measureGuiActionStall $S -Argument $box -Action {
                param ($element)
                $element.GetCurrentPattern([Windows.Automation.ValuePattern]::Pattern).SetValue("＃")
            } -Done { (getGuiText (findGui $S.Window -Id "SummaryText")) -eq "一致 $hits 件中 0 件を表示" }
            closeGui $S
            Write-Host ("S13 D（{0} 件）: 一番長い止まり {1} 秒（{2} 秒・問い合わせ {3} 回）" -f $script:indexCount, $stall.Longest, $stall.Total, $stall.Queries)
            $stall.Longest | Should -BeLessThan $script:stallLimits.D
        }
    }

    It "利用者の環境に触っていない（Office のプロセスの数も変わらない）" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}

# S13 の C は、更新中の毎秒の反映（元のフォルダを実際に取り込む）。ほかの操作のデータ（本文インデックス）と混ざらないよう、別の Describe・別のワークスペースで流す
Describe "S13C インデックスが <count> 件あるときの、更新中の 20 秒の止まり" -Tag Gui -ForEach $stallCounts {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:indexCount = $count
        # インデックスごとの元のファイルの組の数（3 ファイルずつ）。測る 20 秒より前に更新が終わらない数にする
        $script:copiesPerIndex = 1
        $script:measureSeconds = 20
        $script:stallLimitC = $(if ($env:TEBUNKO_GUI_STALL_LIMIT) { [double]$env:TEBUNKO_GUI_STALL_LIMIT } else { 1.0 })
        $dir = Join-Path $TestDrive "c"
        [void][IO.Directory]::CreateDirectory($dir)
        $script:tool = newGuiTool $dir @{ ingestThreads = 1 }
        $folders = foreach ($i in 1..$script:indexCount) {
            $path = Join-Path $dir ("元のフォルダ\{0:000}" -f $i)
            newGuiSourceFolder $path -Copies $script:copiesPerIndex
            @{ name = ("インデックス{0:000}" -f $i); path = $path; enabled = $true }
        }
        $config = readGuiConfig $script:tool
        $config | Add-Member -NotePropertyName targetFolders -NotePropertyValue @($folders) -Force
        writeGuiConfig $script:tool.Dir $config
    }

    It "C ［すべて更新］のあと 20 秒の間、画面が上限の秒数を超えて応答しなくなることはない" {
        $S = startGui $script:tool "S13C-$($script:indexCount)"
        invokeGuiScene $S {
            waitGui $S "登録のフォルダ数" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "IndexFooterFolders")) -like "*$($script:indexCount)*" } | Out-Null
            startGuiIndexing $S
            waitGui $S "取り込み中（［更新中…］）" ${guiDefaultTimeout} { testGuiIndexing $S } | Out-Null
            setGuiStep $S "更新中の 20 秒の止まりを測る"
            $clock = [Diagnostics.Stopwatch]::StartNew()
            $stall = measureGuiActionStall $S -Done { $clock.Elapsed.TotalSeconds -ge $script:measureSeconds }
            # 測った区間が更新中のままでなければ、毎秒の反映が無かったことになる。ファイルの数（copiesPerIndex）を増やす
            if (!(testGuiIndexing $S) -or !(testGuiVisible (findGui $S.Window -Id "IndexingProgressText"))) {
                throw "測っている間に更新が終わってしまった。tests\gui\index_many.Tests.ps1 の copiesPerIndex（インデックスごとのファイルの数）を増やす"
            }
            setGuiStep $S "取り込みを止めて閉じる"
            closeGuiWindowAsync $S $S.Window
            $confirm = waitGuiWindow $S "閉じる確認" -Id "HeadingText" -Text "中止して閉じますか"
            clickGuiByName $S $confirm "中止して閉じる"
            waitGui $S "取り込みを止めて画面が終了する" ${guiIndexTimeout} -AllowExited { $S.Process.HasExited } | Out-Null
            Write-Host ("S13 C（{0} 件）: 一番長い止まり {1} 秒（{2} 秒・問い合わせ {3} 回）" -f $script:indexCount, $stall.Longest, $stall.Total, $stall.Queries)
            $stall.Longest | Should -BeLessThan $script:stallLimitC
        }
    }

    It "利用者の環境に触っていない（Office のプロセスの数も変わらない）" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}
