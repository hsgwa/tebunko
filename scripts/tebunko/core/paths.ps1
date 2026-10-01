# tebunko が使うファイルの場所と、そこに書く値の定義。

# 取り込みの作業フォルダ（本体はワークスペースの tmp\ の下。index_migrate.ps1 の initTmpDir が
# インデックス作成の始めに selectTmpDir で決めて入れる）。決めるまでの既定値は空
${tmpDir} = ""

# ワークスペースの tmp\ が使えないとき（パスに [ ] がある・長すぎる）の代わりの場所（前の版までの置き場所）。
# Excelは [ ] を含むパスに保存できず、Office は長すぎるパスも開けないため、%TEMP% に逃がす。
# インデックス作成を同時に複数実行しても互いのTSVを削除・移動しないよう、プロセスごとに分ける（selectTmpDir が <PID> を足す）
${legacyTmpParent} = Join-Path ([System.IO.Path]::GetTempPath()) "tebunko"

# Excelで開けるパスの長さの目安（古い版の上限）。作業フォルダの候補がこれ以上ならワークスペースの tmp\ を諦める
$excelMaxPath = 218
# 取り込みのスレッドが作業フォルダの下に作る、最も長いファイル名の分（"\w999\converted.pptx"）。
# selectTmpDir で、この分を足しても $excelMaxPath を超えないかを見る
${tmpNameReserve} = "\w999\converted.pptx".Length

# ワークスペース（インデックス・取り込み一覧・ログ・取り込みの出力の置き場所。中の場所は workspace.ps1 の Workspace）。
# setting.config の workspaceFolder で変えられる。空なら既定（settings.ps1 の getWorkDir）
${workspace} = [Workspace]::new((getWorkDir))

# インデックスのフォルダに置く、インデックス名とクロール対象フォルダの対応（インデクサが作成する）。
# インデックスのフォルダごと別の場所・PCへコピーしても、検索結果から元のファイルの場所が分かるようにする。
# 拡張子を .tsv にすると検索対象になるため .txt にする
${sourceFolderFileName} = "元のフォルダ.txt"

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
${statusFolderKey} = "クロール対象フォルダ"
${stateNew}    = "未取り込み"
${stateDone}   = "済"
${stateFailed} = "失敗"
