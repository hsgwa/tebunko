# tebunko_grep が使うファイルの場所と、そこに書く値の定義。

# ツールの ID（多重起動の防止・ミューテックスの名前に使う）
${appId} = "tebunko_grep"

# Excelは [ ] を含むパスに保存できないため TEMP を使う。
# インデックス作成を同時に複数実行しても互いのTSVを削除・移動しないよう、プロセスごとに分ける
${tmpDir}    = Join-Path ([System.IO.Path]::GetTempPath()) "tebunko_grep\${PID}"
${indexDir}  = "${workDir}\index"

# 取り込んだTSVをインデックスに入れる直前に集めるフォルダ（publishIndexFiles）。
# フォルダごと入れ替えるため、インデックスと同じドライブ（work の中）に置く。
# 検索対象に入らないよう work\index の外にする
${publishDir} = "${workDir}\取り込み出力\${PID}"

# インデックスのフォルダに置く、インデックス名とクロール対象フォルダの対応（インデクサが作成する）。
# インデックスのフォルダごと別の場所・PCへコピーしても、検索結果から元のファイルの場所が分かるようにする。
# 拡張子を .tsv にすると検索対象になるため .txt にする
${sourceFolderFileName} = "元のフォルダ.txt"

# 取り込み一覧・出力（自動生成）
${statusFile} = "${workDir}\取り込み一覧.tsv"
${ingestingFile} = "${workDir}\取り込み中.txt"  # 取り込み中のファイル。強制終了で残っていれば、そのファイルの取り込み中に止まった
${resultFile} = "${workDir}\検索結果.txt"

# 画面（gui.ps1）とインデクサ（indexer.ps1）の受け渡し
${stopRequestFile}  = "${workDir}\インデックス作成中止要求"    # 画面が作成すると、インデクサはファイルの切れ目で中止する
${indexingErrorFile} = "${workDir}\インデックス作成エラー.txt"  # インデクサを続けられないエラーのメッセージ（正常終了時は削除）
${indexingLogFile}   = "${workDir}\インデックス作成ログ.txt"    # インデクサの表示内容の記録（実行ごとに上書き）
${guiErrorLogFile}  = "${workDir}\画面エラー.txt"  # 画面で起きた予期しないエラーの記録（追記。原因を後から追えるようにする）
# インデックス作成の進み具合（インデクサが1行だけ書き、画面が読む）。
# 画面が取り込み一覧（数万行になる）を毎秒読み直すと、その間ずっと画面が固まるため、進み具合はこの1行から読む
${indexingProgressFile} = "${workDir}\インデックス作成進捗.txt"
# 取り込み対象を数え終えたときにインデクサが書く、インデックスごとの件数（画面が読んで確認のダイアログに出す）
${ingestPlanFile} = "${workDir}\取り込み予定.tsv"
# 画面が作成すると、インデクサは確認待ちから先へ進む（中身で、前回失敗したファイルも再取り込みするかを伝える）
${indexingStartRequestFile} = "${workDir}\インデックス作成開始要求"

# インデックス作成の進み具合の段階（インデックス作成進捗.txt の1列目）
${indexingPhaseCrawl}   = "クロール"  # 取り込み対象のファイルを探している（件数はまだ分からない）
${indexingPhaseConfirm} = "確認"      # 取り込み対象を数え終え、画面で取り込むかどうかを選ぶのを待っている
${indexingPhaseIngest}  = "取り込み"  # 1ファイルずつ取り込んでいる
${indexingPhaseFinish}  = "仕上げ"    # 後片付け（Officeアプリの終了・取り込み一覧の書き直し）

# 取り込み予定（取り込み予定.tsv）の列と、インデックスごとの区分
${ingestPlanColumns} = @("インデックス名", "元のフォルダ", "区分", "ファイル数", "取り込み対象", "新規", "更新あり", "前回未完了", "インデックスなし", "前回失敗")
${planKindIngest}    = "取り込み"      # チェックが付いていて元のフォルダも見つかった（数えた結果を出す）
${planKindUnchecked} = "チェックなし"  # ［作成］のチェックが外れているため数えていない
${planKindMissing}   = "フォルダなし"  # 元のフォルダが見つからないため数えていない

# インデックス作成開始要求の中身（前回失敗したファイルも再取り込みするかどうか）
${retryFailedMark} = "失敗分も再取り込み"

# 取り込み一覧の列と状態
${statusColumns}   = @("相対パス", "更新日時", "サイズ", "状態", "TSV数", "取り込み日時", "エラー", "抽出版")
${legacyStatusColumnCount} = 7  # 抽出版の列が無い以前の形式の列数
${statusFolderKey} = "クロール対象フォルダ"
${stateNew}    = "未取り込み"
${stateDone}   = "済"
${stateFailed} = "失敗"
