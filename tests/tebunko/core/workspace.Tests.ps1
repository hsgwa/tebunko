# ワークスペースの中の場所（tebunko\core\workspace.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"

Describe "Workspace" -Tag Unit {
    It "中の場所をワークスペースのフォルダから組み立てる（出力用のフォルダはプロセスごと）" {
        $target = [Workspace]::new("C:\Users\test\Documents\tebunko_ws")

        $target.Dir | Should Be "C:\Users\test\Documents\tebunko_ws"
        $target.IndexDir | Should Be "C:\Users\test\Documents\tebunko_ws\index"
        $target.SystemIndexDir | Should Be "C:\Users\test\Documents\tebunko_ws\system_index"
        $target.SystemIndexStateFile | Should Be "C:\Users\test\Documents\tebunko_ws\システムインデックスの状態.tsv"
        $target.PublishDir | Should Be "C:\Users\test\Documents\tebunko_ws\取り込み出力\$PID"
        $target.StatusFile | Should Be "C:\Users\test\Documents\tebunko_ws\取り込み一覧.tsv"
        $target.IngestingFile | Should Be "C:\Users\test\Documents\tebunko_ws\取り込み中.txt"
        $target.ResultFile | Should Be "C:\Users\test\Documents\tebunko_ws\検索結果.txt"
        $target.IndexingLogFile | Should Be "C:\Users\test\Documents\tebunko_ws\インデックス作成ログ.txt"
        $target.GuiErrorLogFile | Should Be "C:\Users\test\Documents\tebunko_ws\画面エラー.txt"
    }

    It "関数は呼んだときの `$workspace の場所を使う（差し替えれば、読み込み直さずに別のワークスペースを使う）" {
        $workspace = newTestWorkspace @{} "$TestDrive\別"

        @((getIndexSummary).Missing) | Should Be @("$TestDrive\別\index")
    }
}
