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
