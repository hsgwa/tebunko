# tebunko_diff が使うファイルの場所と、そこに書く値の定義。

# 比較の作業フォルダ（抽出した TSV を置く）。画面のプロセスごとに分け、閉じたら消す。
# 本文の平文の写しを work に残さないため、TEMP に置く
${diffTempRoot} = Join-Path ([System.IO.Path]::GetTempPath()) "tebunko\diff"

# 比較結果ファイル（［結果をファイルに出力］）
${diffWorkDir}    = "${workDir}\tebunko_diff"
${diffResultFile} = "${diffWorkDir}\比較結果.txt"

# 画面と抽出プロセス（differ.ps1）の受け渡し（作業フォルダの中の名前）
${diffRequestFileName}  = "抽出要求.tsv"    # 画面が書く。1 行 1 ファイル: 番号 <TAB> 側（left / right）<TAB> 元のファイルのパス
${diffResultFileName}   = "抽出結果.tsv"    # 抽出プロセスが 1 件ごとに追記する: 番号 <TAB> 側 <TAB> 状態 <TAB> TSV数 <TAB> エラー
${diffProgressFileName} = "進捗.txt"        # 抽出プロセスが書く 1 行: 済んだ数 <TAB> 全体の数 <TAB> 今のファイルのパス
${diffPriorityFileName} = "優先.tsv"        # 画面が書く。次に抽出してほしい 番号 <TAB> 側（1 行 1 件。上から）
${diffStopFileName}     = "中止要求"        # 画面が作ると、抽出プロセスはファイルの切れ目で止まる
${diffErrorFileName}    = "比較エラー.txt"  # 抽出プロセスを続けられないエラーのメッセージ
${diffOrderFileName}    = "順番.txt"        # 抽出した TSV の名前を、抽出した順に並べたもの（場所の順番）

${diffStateDone}   = "済"
${diffStateFailed} = "失敗"
${diffSides}       = @("left", "right")

# 1 ファイルの抽出の制限時間（分）。インデクサと同じ
${diffFileTimeoutMinutes} = 10

# 比べるファイルの数の上限（フォルダ同士。両方合わせて）
${diffMaxFiles} = 10000
