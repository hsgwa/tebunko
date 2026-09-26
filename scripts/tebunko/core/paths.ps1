# tebunko が使うファイルの場所と、そこに書く値の定義。

# ツールの ID（多重起動の防止・ミューテックスの名前に使う）
${appId} = "tebunko"

# Excelは [ ] を含むパスに保存できないため TEMP を使う。
# インデックス作成を同時に複数実行しても互いのTSVを削除・移動しないよう、プロセスごとに分ける
${tmpDir}    = Join-Path ([System.IO.Path]::GetTempPath()) "tebunko\${PID}"

# work（インデックス・取り込み一覧・ログ・取り込みの出力）の置き場所。
# setting.config の workspaceFolder で変えられる。空なら既定（設定ファイルと同じフォルダの work。settings.ps1 の getWorkDir）
${workDir}   = getWorkDir
${indexDir}  = "${workDir}\index"
# システムインデックス（本文インデックスの 2-gram を書いた txt。index と同じ相対パスの構成。Windows Search に索引させる）と、その状態
${systemIndexDir} = "${workDir}\system_index"
${systemIndexStateFile} = "${workDir}\システムインデックスの状態.tsv"

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

# インデックス作成の記録。画面とインデクサの受け渡しはメモリ上で行う（indexer_state.ps1 の newIndexerChannel）
${indexingLogFile}   = "${workDir}\インデックス作成ログ.txt"    # インデクサの表示内容の記録（実行ごとに上書き）
${guiErrorLogFile}  = "${workDir}\画面エラー.txt"  # 画面で起きた予期しないエラーの記録（追記。原因を後から追えるようにする）

# インデックス作成の進み具合の段階（writeIndexingProgress の phase）
${indexingPhaseCrawl}   = "クロール"  # 取り込み対象のファイルを探している（件数はまだ分からない）
${indexingPhaseConfirm} = "確認"      # 取り込み対象を数え終え、画面で取り込むかどうかを選ぶのを待っている
${indexingPhaseIngest}  = "取り込み"  # 1ファイルずつ取り込んでいる
${indexingPhaseFinish}  = "仕上げ"    # 後片付け（Officeアプリの終了・取り込み一覧の書き直し）

# 取り込み予定（newIngestPlanRow）の列と、インデックスごとの区分
${ingestPlanColumns} = @("インデックス名", "元のフォルダ", "区分", "ファイル数", "取り込み対象", "新規", "更新あり", "前回未完了", "インデックスなし", "前回失敗")
${planKindIngest}    = "取り込み"      # チェックが付いていて元のフォルダも見つかった（数えた結果を出す）
${planKindUnchecked} = "チェックなし"  # ［作成］のチェックが外れているため数えていない
${planKindMissing}   = "フォルダなし"  # 元のフォルダが見つからないため数えていない

# 取り込み一覧の列と状態
${statusColumns}   = @("相対パス", "更新日時", "サイズ", "状態", "TSV数", "取り込み日時", "エラー", "抽出版")
${legacyStatusColumnCount} = 7  # 抽出版の列が無い以前の形式の列数
${statusFolderKey} = "クロール対象フォルダ"
${stateNew}    = "未取り込み"
${stateDone}   = "済"
${stateFailed} = "失敗"
