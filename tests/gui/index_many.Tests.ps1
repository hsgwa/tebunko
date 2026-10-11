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
