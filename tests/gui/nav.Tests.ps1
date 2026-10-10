# 画面のスモークテスト S9: ナビで画面を切り替えたときの読み直しと、修飾キーなしの F5（共通の関数は gui_helpers.ps1）。
# 計画の「ナビで［インデックス管理］へ移ると状態が読み直される」を確かめる。
# キーは、Ctrl を押した形を送れない（ハンドラが本物のキーボードの状態を読むため）。修飾キーなしの F5 だけ、WM_KEYDOWN のメッセージで送る
# （Ctrl 付きの振り分けは tests\tebunko\ui\shell\nav.Tests.ps1 で確かめる）。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\gui_helpers.ps1"
}

Describe "S9 ナビの画面の切り替えと F5" -Tag Gui {
    BeforeAll {
        $script:envBefore = getGuiEnvSnapshot
        $script:tool = newGuiTool $TestDrive
    }

    It "検索から［インデックス管理］へ移ると取り込み一覧の状態が読み直され、F5 でも読み直される" {
        # 起動したあとに、取り込み一覧（取り込んでいないファイルが残っている形）を書く。起動時の読み込みでは読めない
        $writeStatus = {
            param ($count)
            $lines = @("相対パス`t更新日時`tサイズ`t状態`tTSV数`t取り込み日時`tエラー`t抽出版")
            for ($i = 1; $i -le $count; $i++) { $lines += "営業\資料$i.docx`t2026-01-01 00:00:00`t100`t未取り込み`t0`t`t`t" }
            [IO.File]::WriteAllText("$($script:tool.Work)\ingest_status.tsv", ($lines -join "`r`n") + "`r`n", (New-Object Text.UTF8Encoding($true)))
        }
        $S = startGui $script:tool "S9"
        invokeGuiScene $S {
            # 窓を先に前面にする（あとで前面になると Activated で状態が読み直され、F5・ナビの確かめと区別できなくなる）
            activateGuiWindow $S
            # 起動時の読み込み（取り込み一覧・本文インデックスのファイル）が終わるまで待つ。終わる前に取り込み一覧を書くと、
            # 起動時の読み込みがそれを読んでしまい、ナビの切り替えでの読み直しと区別できなくなる
            setGuiStep $S "［インデックス管理］で起動時の読み込みの終わりを待つ"
            selectGuiTab $S "IndexTab" "IndexingButton"
            waitGui $S "起動時の読み込みが終わる（集約ファイルが無い）" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "IndexSummaryText")) -like "まだインデックスがありません*" } | Out-Null
            waitGui $S "起動時の読み込みが終わる（取り込みの状態に残りが無い）" ${guiDefaultTimeout} { (getGuiIndexingBannerText $S) -notlike "*残り*" } | Out-Null

            setGuiStep $S "［検索］へ移る"
            selectGuiTab $S "SearchTab" "GoIndexTabButton"

            setGuiStep $S "取り込み一覧を書き、［インデックス管理］へ戻る"
            & $writeStatus 1
            selectGuiTab $S "IndexTab" "IndexingButton"
            waitGui $S "状態が読み直される（残り 1 件）" ${guiDefaultTimeout} { (getGuiIndexingBannerText $S) -like "*残り 1 件*" } | Out-Null

            setGuiStep $S "取り込み一覧と集約ファイルを書き換えて F5"
            waitGui $S "集約ファイルはまだ無い" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "IndexSummaryText")) -like "まだインデックスがありません*" } | Out-Null
            & $writeStatus 2
            $contentIndexDir = "$($script:tool.Work)\content_index"
            [IO.Directory]::CreateDirectory($contentIndexDir) | Out-Null
            [IO.File]::WriteAllText("$contentIndexDir\content_index.xlsx.tsv", "x`r`n", (New-Object Text.UTF8Encoding($true)))
            pressGuiKey $S.Window 0x74
            # F5 だけが読み直すもの（本文インデックスのファイルの件数）でも待ち、Activated と区別する
            waitGui $S "F5 で読み直される（残り 2 件・集約ファイル）" ${guiDefaultTimeout} {
                (getGuiIndexingBannerText $S) -like "*残り 2 件*" -and
                    (getGuiText (findGui $S.Window -Id "IndexSummaryText")) -like "*集約ファイル*"
            } | Out-Null
            closeGui $S
        }
    }

    It "利用者の環境に触っていない（Office のプロセスの数も変わらない）" {
        compareGuiEnvSnapshot $script:envBefore (getGuiEnvSnapshot) | Should -BeNullOrEmpty
    }
}
