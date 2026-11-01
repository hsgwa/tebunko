# 画面のスモークテスト: 取り込んだファイルが数万件あるインデックスで、［ファイルの一覧］のダイアログを開いてから件数が出るまでと、
# 表を末尾までスクロールしたときの応答が、決めた秒数以内であること（表は見えている行だけ作る＝仮想化）。
# 秒数は GitHub Actions の Windows の実行機の実測（開くまで 12 秒ほど）に、混んだときの余裕を見て決めた上限（遅くなったことに気づくための目安で、速さの計測ではない）。
# 流すのは Gui タグ（gui.yml）。共通の関数は gui_helpers.ps1。index_many.Tests.ps1 と同じ形。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"
}

Describe "S2d 取り込んだファイルが 30,000 件ある" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:fileCount = 30000
        $script:openLimitSeconds = 60
        $script:scrollLimitSeconds = 30
        $script:tool = newGuiTool $TestDrive
        $config = readGuiConfig $script:tool
        $config | Add-Member -NotePropertyName targetFolders -NotePropertyValue @(@{ name = "営業"; path = (Join-Path $TestDrive "元のフォルダ\営業"); enabled = $true }) -Force
        writeGuiConfig $script:tool.Dir $config
        # 取り込み一覧（取り込み済みの 30,000 行）を、起動の前に書く
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add("相対パス`t更新日時`tサイズ`t状態`tTSV数`t取り込み日時`tエラー`t抽出版")
        foreach ($i in 1..$script:fileCount) {
            $lines.Add(("営業\資料{0:00000}.docx`t2026-01-01 00:00:00`t100`t済`t1`t2026-01-02 00:00:00`t`t" -f $i))
        }
        [void][IO.Directory]::CreateDirectory("$($script:tool.Work)\store")
        [IO.File]::WriteAllText("$($script:tool.Work)\store\ingest_status.tsv", ($lines -join "`r`n") + "`r`n", (New-Object Text.UTF8Encoding($true)))
    }

    It "ダイアログが上限の秒数以内に開いて件数が出て、末尾までスクロールしても応答する" {
        $S = startGui $script:tool "S2d"
        invokeGuiScene $S {
            setGuiStep $S "行を選ぶ"
            $row = waitGui $S "一覧の先頭の行" ${guiDefaultTimeout} { @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")) | Select-Object -First 1 }
            selectGui $row
            waitGui $S "詳細の見出し" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "IndexDetailTitle")) -eq "営業 - 詳細" } | Out-Null

            setGuiStep $S "［ファイルの一覧］を開く"
            $sw = [Diagnostics.Stopwatch]::StartNew()
            clickGui $S $S.Window "IndexFilesButton" "詳細の［ファイルの一覧］"
            $dialog = waitGuiWindow $S "取り込んだファイルの一覧のダイアログ" -Id "IndexFilesGrid"
            waitGui $S "見出しに件数が出る" ${guiDefaultTimeout} { (getGuiText (findGui $dialog -Id "IndexFilesHeading")) -like "ファイル 30,000 件*" } | Out-Null
            $S.Timing["開くまで"] = [Math]::Round($sw.Elapsed.TotalSeconds, 1)
            $sw.Elapsed.TotalSeconds | Should -BeLessThan $script:openLimitSeconds
            # 見えている行だけが作られている（30,000 行は作らない）
            $rows = @(getGuiGridRows (findGui $dialog -Id "IndexFilesGrid"))
            $rows.Count | Should -BeGreaterThan 0
            $rows.Count | Should -BeLessThan 500

            setGuiStep $S "末尾までスクロールする"
            $sw.Restart()
            $scroll = (findGui $dialog -Id "IndexFilesGrid").GetCurrentPattern([System.Windows.Automation.ScrollPattern]::Pattern)
            $scroll.SetScrollPercent([System.Windows.Automation.ScrollPattern]::NoScroll, 100)
            waitGui $S "末尾の行が出る" ${guiDefaultTimeout} {
                @(getGuiGridRows (findGui $dialog -Id "IndexFilesGrid") | Where-Object { ((getGuiRowTexts $_) -join " ") -like "*資料30000.docx*" }).Count -gt 0
            } | Out-Null
            $S.Timing["末尾までスクロール"] = [Math]::Round($sw.Elapsed.TotalSeconds, 1)
            $sw.Elapsed.TotalSeconds | Should -BeLessThan $script:scrollLimitSeconds

            clickGui $S $dialog "IndexFilesCloseButton" "［閉じる］"
            waitGuiWindowClosed $S $dialog "ファイルの一覧"
            closeGui $S
            Write-Host ("S2d の秒数: " + (($S.Timing.GetEnumerator() | ForEach-Object { "$($_.Key) $($_.Value)" }) -join "・"))
        }
    }

    It "利用者の環境に触っていない（Office のプロセスの数も変わらない）" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}
