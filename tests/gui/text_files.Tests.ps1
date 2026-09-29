# 画面でのテキストファイルの取り込み・検索・表示・プレビュー・開く（共通の関数は gui_helpers.ps1）。
# feat「.txt などテキストファイルも検索できるようにする」の実機の確かめ（行番号の表示・前後のプレビュー・
# 長い行を切って固まらないこと・実在するファイルを開けること）を、UI オートメーションで実機の画面を使って確かめる。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"
}

Describe "S5 テキストファイルの検索と表示" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:source = Join-Path $TestDrive "元のフォルダ\資料"
        [void][System.IO.Directory]::CreateDirectory($script:source)
        [System.IO.File]::WriteAllText("$($script:source)\メモ.txt", "1行目`r`n検索語のある行`r`n3行目`r`n", (New-Object System.Text.UTF8Encoding($false)))
        # 1 行が長い .json（1,000 文字を超え、切って表示する対象になる長さ）
        $filler = "a" * 2000
        [System.IO.File]::WriteAllText("$($script:source)\big.json", ('{"k":"' + $filler + 'ロングヒット' + $filler + '"}'), (New-Object System.Text.UTF8Encoding($false)))
        $script:tool = newGuiTool $TestDrive @{ targetFolders = @(@{ name = "資料"; path = $script:source; enabled = $true }) }
    }

    It "取り込み・検索結果の表示（行番号）・プレビュー・実在するファイルを開くが実機で動く" {
        $S = startGui $script:tool "S5"
        invokeGuiScene $S {
            $hitRows = { getGuiHitRows (findGui $S.Window -Id "ResultGrid") }
            $search = {
                param ($word)
                setGuiText $S (findGui $S.Window -Id "WordBox") $word
                waitGuiEnabled $S (findGui $S.Window -Id "SearchButton") "［検索］"
                clickGui $S $S.Window "SearchButton" "［検索］"
            }

            setGuiStep $S "インデックス作成"
            startGuiIndexing $S
            waitGui $S "取り込みが終わる" ${guiIndexTimeout} { !(testGuiIndexing $S) } | Out-Null

            setGuiStep $S "［2 検索］へ"
            selectGuiTab $S "SearchTab"

            setGuiStep $S "検索語で検索"
            & $search "検索語のある行"
            waitGui $S "該当 1 件" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "SummaryText")) -like "該当 1 件*" } | Out-Null
            clickGui $S $S.Window "ExpandAllButton" "［すべて展開］"
            $row = waitGui $S "結果の行が出る" ${guiDefaultTimeout} { @(& $hitRows) | Select-Object -First 1 }

            setGuiStep $S "場所の表記（N 行目）を見る"
            ((getGuiRowTexts $row) -join "|") | Should -Match "2 行目"

            setGuiStep $S "選択行のプレビュー（前後の行）を見る"
            selectGui $row
            waitGui $S "プレビューに前後の行が出る" ${guiDefaultTimeout} {
                $texts = (@(findAllGui $S.Window -Type Text) | ForEach-Object { $_.Current.Name }) -join "|"
                $texts -match "1行目" -and $texts -match "3行目"
            } | Out-Null

            setGuiStep $S "実在する元のファイルを開く"
            # ［開く］は利用者の既定のアプリ（メモ帳・サクラエディタ等。環境で変わる）で実際に開く。名前・題名では
            # ほかの利用者のアプリまで止めかねないため、押す前後の PID の差だけで新しく起動したものを見つける
            # （既存のプロセスに開かれた場合は閉じない。閉じなくても $TestDrive の片付けは困らない）
            $pidsBefore = @(Get-Process -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
            waitGuiEnabled $S (findGui $S.Window -Id "OpenButton") "［開く］"
            clickGui $S $S.Window "OpenButton" "［開く］"
            waitGui $S "ステータスが「開きました」になる" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "StatusText")) -like "開きました*" } | Out-Null
            foreach ($p in @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $pidsBefore -notcontains $_.Id })) {
                $S.Extra += $p
            }

            setGuiStep $S "1 行が長いファイルの語を検索し、選んでも固まらない"
            & $search "ロングヒット"
            waitGui $S "該当 1 件" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "SummaryText")) -like "該当 1 件*" } | Out-Null
            clickGui $S $S.Window "ExpandAllButton" "［すべて展開］"
            $bigRow = waitGui $S "結果の行が出る" ${guiDefaultTimeout} { @(& $hitRows) | Select-Object -First 1 }
            selectGui $bigRow
            # 選択後、既定の待ち時間内に［開く］が押せる状態になれば、画面のスレッドが固まっていないと分かる
            waitGuiEnabled $S (findGui $S.Window -Id "OpenButton") "［開く］"

            closeGui $S
        }
    }

    It "利用者の環境に触っていない" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}
