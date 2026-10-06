# 画面のスモークテスト S9: ナビで画面を移ったときの読み直しと、修飾キーなしの F5（共通の関数は gui_helpers.ps1）。
# 計画の「ナビで［1 インデックス管理］へ移ると状態が読み直される」を確かめる。［9 プロセス停止］へ移ると一覧が読まれることは S7（process.Tests.ps1）が確かめる。
# キーは、Ctrl を押した形を送れない（ハンドラが本物のキーボードの状態を読むため）。修飾キーなしの F5 だけ、WM_KEYDOWN のメッセージで送る
# （Ctrl 付きの振り分けは tests\tebunko\ui\shell\nav.Tests.ps1 で確かめる）。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"
}

Describe "S9 ナビの切り替えと F5" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:tool = newGuiTool $TestDrive
    }

    It "検索から［1 インデックス管理］へ移ると取り込み一覧の状態が読み直され、F5 でも読み直される" {
        # 起動したあとに、取り込み一覧（取り込んでいないファイルが残っている形）を書く。起動時の読み込みでは読めない
        $writeStatus = {
            param ($count)
            $lines = @("相対パス`t更新日時`tサイズ`t状態`tTSV数`t取り込み日時`tエラー`t抽出版")
            for ($i = 1; $i -le $count; $i++) { $lines += "営業\資料$i.docx`t2026-01-01 00:00:00`t100`t未取り込み`t0`t`t`t" }
            [IO.File]::WriteAllText("$($script:tool.Work)\ingest_status.tsv", ($lines -join "`r`n") + "`r`n", (New-Object Text.UTF8Encoding($true)))
        }
        $S = startGui $script:tool "S9"
        invokeGuiScene $S {
            setGuiStep $S "［2 検索］へ移る"
            selectGuiTab $S "SearchTab" "GoIndexTabButton"

            setGuiStep $S "取り込み一覧を書き、［1 インデックス管理］へ移る"
            & $writeStatus 1
            selectGuiTab $S "IndexTab" "IndexingStateText"
            waitGui $S "状態が読み直される（残り 1 件）" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "IndexingStateText")) -like "*残り 1 件*" } | Out-Null

            setGuiStep $S "取り込み一覧を書き換えて F5"
            & $writeStatus 2
            activateGuiWindow $S
            pressGuiKey $S.Window 0x74
            waitGui $S "F5 で読み直される（残り 2 件）" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "IndexingStateText")) -like "*残り 2 件*" } | Out-Null
            closeGui $S
        }
    }

    It "利用者の環境に触っていない（Office のプロセスの数も変わらない）" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}
