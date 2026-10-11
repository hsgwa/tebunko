# tebunko が使うファイルの場所と、そこに書く値の定義。

# 取り込みの作業フォルダ（本体はワークスペースの tmp\ の下。index_migrate.ps1 の initTmpDir が
# インデックス作成の始めに selectTmpDir で決めて入れる）。決めるまでの既定値は空。
# 置けなかったとき（パスに [ ] がある・長すぎる）も空のまま（どのファイルも中間 TSV などをこのフォルダに作るため、
# テキストファイルを含めすべての取り込みをスキップする。%TEMP% には逃がさない）
${tmpDir} = ""
# ${tmpDir} を置けなかった理由（selectTmpDir の Reason）。置けたときは空
${tmpDirReason} = ""

# 前の版（%TEMP%\tebunko\<PID> に一時ファイルを置いていた版）が残した作業フォルダの片付け専用。
# 今の版はここには書き込まない（removeStaleTmpDirs が、強制終了などで残った前の版のフォルダを消すためだけに使う）
${legacyTmpParent} = Join-Path ([System.IO.Path]::GetTempPath()) "tebunko"

# Excelで開けるパスの長さの目安（古い版の上限）。作業フォルダの候補がこれ以上ならワークスペースの tmp\ を諦める
$excelMaxPath = 218
# 取り込みのスレッドが作業フォルダの下に作る、最も長いファイル名の分（"\w999\converted.pptx"）。
# selectTmpDir で、この分を足しても $excelMaxPath を超えないかを見る
${tmpNameReserve} = "\w999\converted.pptx".Length

# ワークスペース（インデックス・取り込み一覧・ログ・取り込みの出力の置き場所。中の場所は workspace.ps1 の Workspace）。
# setting.config の workspaceFolder で変えられる。空なら既定（settings.ps1 の getWorkDir）。
# 読み込んだとき（スクリプトの読み込み時）には決めず、起動口（startGui・invokeIndexerMain）が
# 壊れた設定ファイルの退避の後に呼ぶ（単一 .ps1 版は setting.config を読む前に読み込みだけ先に済ませるため）
function initWorkspace {
    ${script:workspace} = [Workspace]::new((getWorkDir))
}

# インデックスのフォルダに置く、インデックス名とクロール対象フォルダの対応（インデクサが作成する）。
# インデックスのフォルダごと別の場所・PCへコピーしても、検索結果から元のファイルの場所が分かるようにする。
# 拡張子を .tsv にすると検索対象になるため .txt にする
${sourceFolderFileName} = "source_folder.txt"

# 前の版(ワークスペースのファイル名を英語化する前)の元のフォルダの記録。
# getLegacyIndexState が旧版の index\ フォルダを見分けるためだけに使う。読み取り専用。
${legacySourceFolderFileName} = "元のフォルダ.txt"

# インデックス作成の進み具合の段階（writeIndexingProgress の phase）
${indexingPhaseCrawl}   = "クロール"  # 取り込み対象のファイルを探している（件数はまだ分からない）
${indexingPhaseConfirm} = "確認"      # 取り込み対象を数え終え、画面で取り込むかどうかを選ぶのを待っている
${indexingPhaseIngest}  = "取り込み"  # 1ファイルずつ取り込んでいる
${indexingPhaseFinish}  = "仕上げ"    # 後片付け（Officeアプリの終了・取り込み一覧の書き直し）

# 取り込み予定（newIngestPlanRow）の列と、インデックスごとの区分
${ingestPlanColumns} = @("インデックス名", "元のフォルダ", "区分", "ファイル数", "取り込み対象", "新規", "更新あり", "前回未完了", "インデックスなし", "前回失敗", "クラウド", "クラウド失敗", "クラウド容量", "クラウド失敗容量")
${planKindIngest}    = "取り込み"      # チェックが付いていて元のフォルダも見つかった（数えた結果を出す）
${planKindUnchecked} = "チェックなし"  # チェックが外れているため数えていない
${planKindMissing}   = "フォルダなし"  # 元のフォルダが見つからないため数えていない
${planKindDropped}   = "削除予定"      # 設定から外れたインデックス（確認で［更新を開始］が押されたときだけ消す）

# 取り込み一覧の列と状態
${statusColumns}   = @("相対パス", "更新日時", "サイズ", "状態", "TSV数", "取り込み日時", "エラー", "抽出版")
${statusFolderKey} = "クロール対象フォルダ"
${stateNew}    = "未取り込み"
${stateDone}   = "済"
${stateFailed} = "失敗"
