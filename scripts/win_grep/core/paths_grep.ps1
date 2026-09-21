# win_grep が使うファイルの場所と、そこに書く値の定義。

# ツールの ID（多重起動の防止・ミューテックスの名前に使う）
${appId} = "win_grep"

# Excelは [ ] を含むパスに保存できないため TEMP を使う。
# 変換を同時に複数実行しても互いのTSVを削除・移動しないよう、プロセスごとに分ける
${tmpDir}    = Join-Path ([System.IO.Path]::GetTempPath()) "win_grep\${PID}"
${indexDir}  = "${workDir}\index"

# 変換したTSVをインデックスに入れる直前に集めるフォルダ（publishIndexFiles）。
# フォルダごと入れ替えるため、インデックスと同じドライブ（work の中）に置く。
# 検索対象に入らないよう work\index の外にする
${publishDir} = "${workDir}\変換出力\${PID}"

# インデックスのフォルダに置く、インデックス名と変換対象フォルダの対応（変換処理が作成する）。
# インデックスのフォルダごと別の場所・PCへコピーしても、検索結果から元のファイルの場所が分かるようにする。
# 拡張子を .tsv にすると検索対象になるため .txt にする
${sourceFolderFileName} = "元のフォルダ.txt"

# 変換一覧・出力（自動生成）
${statusFile} = "${workDir}\変換一覧.tsv"
${convertingFile} = "${workDir}\変換中.txt"  # 変換中のファイル。強制終了で残っていれば、そのファイルの変換中に止まった
${resultFile} = "${workDir}\検索結果.txt"

# 画面（config_gui.ps1）と変換処理（office_to_tsv.ps1）の受け渡し
${stopRequestFile}  = "${workDir}\変換中止要求"    # 画面が作成すると、変換処理はファイルの切れ目で中止する
${convertErrorFile} = "${workDir}\変換エラー.txt"  # 変換処理を続けられないエラーのメッセージ（正常終了時は削除）
${convertLogFile}   = "${workDir}\変換ログ.txt"    # 変換処理の表示内容の記録（実行ごとに上書き）
${guiErrorLogFile}  = "${workDir}\画面エラー.txt"  # 画面で起きた予期しないエラーの記録（追記。原因を後から追えるようにする）
# 変換の進み具合（変換が1行だけ書き、画面が読む）。
# 画面が変換一覧（数万行になる）を毎秒読み直すと、その間ずっと画面が固まるため、進み具合はこの1行から読む
${convertProgressFile} = "${workDir}\変換進捗.txt"
# 変換対象を数え終えたときに変換側が書く、インデックスごとの件数（画面が読んで確認のダイアログに出す）
${convertPlanFile} = "${workDir}\変換予定.tsv"
# 画面が作成すると、変換処理は確認待ちから先へ進む（中身で、前回失敗したファイルも再変換するかを伝える）
${convertStartRequestFile} = "${workDir}\変換開始要求"

# 変換の進み具合の段階（変換進捗.txt の1列目）
${convertPhaseScan}    = "準備"    # 変換対象のファイルを探している（件数はまだ分からない）
${convertPhaseConfirm} = "確認"    # 変換対象を数え終え、画面で変換するかどうかを選ぶのを待っている
${convertPhaseRun}     = "変換"    # 1ファイルずつ変換している
${convertPhaseFinish}  = "仕上げ"  # 後片付け（Officeアプリの終了・変換一覧の書き直し）

# 変換予定（変換予定.tsv）の列と、インデックスごとの区分
${convertPlanColumns} = @("インデックス名", "元のフォルダ", "区分", "ファイル数", "変換対象", "新規", "更新あり", "前回未完了", "変換結果なし", "前回失敗")
${planKindConvert}   = "変換"          # チェックが付いていて元のフォルダも見つかった（数えた結果を出す）
${planKindUnchecked} = "チェックなし"  # ［変換］のチェックが外れているため数えていない
${planKindMissing}   = "フォルダなし"  # 元のフォルダが見つからないため数えていない

# 変換開始要求の中身（前回失敗したファイルも再変換するかどうか）
${retryFailedMark} = "失敗分も再変換"

# 変換一覧の列と状態
${statusColumns}   = @("相対パス", "更新日時", "サイズ", "状態", "TSV数", "変換日時", "エラー")
${statusFolderKey} = "変換対象フォルダ"
${stateNew}    = "未変換"
${stateDone}   = "済"
${stateFailed} = "失敗"

# 検索結果から元のファイルを開くときの開き方（設定 openMode の値）
${openModeNormal}   = "normal"    # そのまま開く（編集する）
${openModeReadOnly} = "readOnly"  # 読み取り専用で開く（誤って上書きしない）
${openModeNew}      = "new"       # 新規（元のファイルを基にした無題の文書）で開く。元のファイルを占有しない
${openModes}        = @(${openModeNormal}, ${openModeReadOnly}, ${openModeNew})
