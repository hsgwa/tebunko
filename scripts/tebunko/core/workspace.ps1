# ワークスペース（インデックス・取り込み一覧・ログを置くフォルダ）の中の、tebunko が使うファイルの場所。
# 場所はフォルダ（Dir）から組み立てるだけで、設定は読まない。どのフォルダを使うかは設定の workspaceFolder で決まる（settings.ps1 の getWorkDir）。
# 今のワークスペースは paths.ps1 の ${workspace} に置く。別のスレッドへは Dir（文字列）を渡し、そこで作り直す。
# クラスのメソッドからはスクリプトの関数・変数が見えないため、ここでは .NET の型だけを使う

class Workspace {
    [string]$Dir
    [string]$IndexDir
    # システムインデックス（本文インデックスの 2-gram を書いた txt。index と同じ相対パスの構成。Windows Search に索引させる）と、その状態
    [string]$SystemIndexDir
    [string]$SystemIndexStateFile
    # 取り込んだTSVをインデックスに入れる直前に集めるフォルダ（publishIndexFiles）。
    # フォルダごと入れ替えるため、インデックスと同じドライブ（ワークスペースの中）に置く。
    # 検索対象に入らないよう index の外にする。インデックス作成を同時に複数実行しても混ざらないよう、プロセスごとに分ける
    [string]$PublishDir
    # 取り込み一覧・出力（自動生成）
    [string]$StatusFile
    [string]$IngestingFile  # 取り込み中のファイル。強制終了で残っていれば、そのファイルの取り込み中に止まった
    [string]$ResultFile
    # インデックス作成の記録。画面とインデクサの受け渡しはメモリ上で行う（indexer_state.ps1 の newIndexerChannel）
    [string]$IndexingLogFile  # インデクサの表示内容の記録（実行ごとに上書き）
    [string]$GuiErrorLogFile  # 画面で起きた予期しないエラーの記録（追記。原因を後から追えるようにする）

    Workspace([string]$dir) {
        $this.Dir = $dir
        $this.IndexDir = "$dir\index"
        $this.SystemIndexDir = "$dir\system_index"
        $this.SystemIndexStateFile = "$dir\システムインデックスの状態.tsv"
        $this.PublishDir = "$dir\取り込み出力\$([System.Diagnostics.Process]::GetCurrentProcess().Id)"
        $this.StatusFile = "$dir\取り込み一覧.tsv"
        $this.IngestingFile = "$dir\取り込み中.txt"
        $this.ResultFile = "$dir\検索結果.txt"
        $this.IndexingLogFile = "$dir\インデックス作成ログ.txt"
        $this.GuiErrorLogFile = "$dir\画面エラー.txt"
    }
}
