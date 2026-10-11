# 部品ごとの関数（インデックス作成）

扱うこと: 取り込み一覧・状態ファイル・画面との受け渡しの口（`indexer_state.ps1`）、取り込み直すかの判断（`indexer_decide.ps1`）、インデックス名とインデックスの管理（`index_name.ps1`・`index_store.ps1`）の関数一覧。扱わないこと: 設定・ファイル操作（[部品ごとの関数（設定・ファイル）](settings.md)）、TSV・本文インデックス（[部品ごとの関数（TSV と本文インデックス）](tsv.md)）。先に読むページ: [部品から関数一覧を引く](index.md)。

## 取り込み一覧・状態ファイル・画面との受け渡しの口（`tebunko/indexer/indexer_state.ps1`）

| 関数 | 入力 | 出力 | 概要 | 詳細 | 使用元 |
|---|---|---|---|---|---|
| `newStatusRow` | 相対パス, 更新日時, サイズ, 状態, TSV数, 取り込み日時, エラー, 抽出版（2 番目以降は省略可） | 行オブジェクト | 取り込み一覧の 1 行を作る | [取り込み一覧](../indexing/ingest-list.md) | インデックス作成 |
| `toStatusLine` | row | string | 取り込み一覧の 1 行を文字列にする（エラーのタブ・改行はスペース）。`writeStatusFile` は速さのため同じ整形をその場に展開しているため、変えるときは両方を直す（テストで同じ結果になることを確かめている） | 同上 | addStatusRow |
| `describeIngestError` | exception | string | 取り込みの例外から、取り込み一覧のエラー列・画面に表示する失敗の原因を作る | [失敗の原因](../indexing/office-apps.md#失敗の原因describeingesterror) | インデックス作成 |
| `readStatusFile` | path（既定 `$workspace.StatusFile`） | `@{Folders; Rows}` | 取り込み一覧を読む。Folders はクロール対象フォルダ `@{Path; Name}`（Name = インデックス名）、Rows は相対パス（`インデックス名\フォルダからの相対パス`。大文字・小文字を区別しない）→ 行。同じ相対パスは後の行を優先し、列数の合わない行は無視。インデックス作成中に読んでも書き込みを妨げない共有モードで開く。無ければ空 | [取り込み一覧](../indexing/ingest-list.md) | インデックス作成・画面 |
| `writeStatusFile` | folders（`@{Path; Name}` の配列）, rows, path（既定 `$workspace.StatusFile`） | – | 取り込み一覧を書き出す（先頭にクロール対象フォルダの行、1 ファイル 1 行。一時ファイルに書いてから置き換える） | 同上 | インデックス作成 |
| `addStatusRow` | row, path（既定 `$workspace.StatusFile`） | – | 取り込み一覧の末尾に 1 行追記する | 同上 | インデックス作成 |
| `readStatusLines` | path（既定 `$workspace.StatusFile`） | 行の配列 | 取り込み一覧を共有を許して 1 行ずつ読む | – | renameStatusIndexName, removeStatusIndexName |
| `renameStatusIndexName` / `removeStatusIndexName` | oldName, newName, path / name, path | – | 取り込み一覧のインデックス名を書き換える / その記録を取り除く。行の順序と内容はそのまま保つ（`readStatusLines` で読み、`writeTextLinesAtomic` で置き換える） | [追加・編集のダイアログ](../gui/index-tab.md#追加編集のダイアログ) | renameIndex, removeIndex |
| `readIngestingFiles` | path（既定 `$workspace.IngestingFile`） | `@{RelPath; Count}` の配列 | 取り込み中のファイルの記録（1 行に 1 ファイル。相対パスと、続けて取り込みを始めて終わらなかった回数）を読む。無ければ空。壊れた行は読み飛ばす | [取り込み一覧](../indexing/ingest-list.md#強制終了時間切れからの再開) | インデックス作成 |
| `writeIngestingFiles` | entries（`@{RelPath; Count}` の配列）, path（既定 `$workspace.IngestingFile`） | – | 取り込み中のファイルを 1 行に 1 つ `<回数><TAB><相対パス>` で記録する。無ければ記録を消す | 同上 | インデックス作成 |
| `removeIngestingFile` | path（既定 `$workspace.IngestingFile`） | – | 取り込み中のファイルの記録を削除する（無くてもエラーにしない） | 同上 | インデックス作成 |
| `newIndexerChannel` | retryFailed, confirmTargets, workers（既定 -1）, onlyNames（既定は空）, includeCloud（既定 `$false`） | 受け渡しの口（`[hashtable]::Synchronized`） | 画面とインデクサの受け渡しの口を作る（`RetryFailed`・`ConfirmTargets`・`Workers`・`OnlyNames`（更新するインデックス名。空ならチェックの付いたものすべて）・`OnlySkipped`（`OnlyNames` のうち更新できなかった名前と理由）・`Progress`・`Stop`・`Plan`・`Answer`・`Answered`・`Error`・`ExitCode`・`Notice`・`Postponed`・`IncludeCloud`（クラウドにだけあるファイルもダウンロードして取り込むか）・`CloudSkipped`（ダウンロードせずに残した件数）・`OfficePids`）。`Workers` は -1 で設定・コア数から決める、0 で司令のスレッドで取り込む（テスト） | [画面とインデクサの受け渡し](../structure/threads.md#画面とインデクサの受け渡し) | 画面・indexer.ps1 |
| `writeIndexingProgress` | phase, processed, remaining, failed, detail, channel（既定はいま動いているインデックス作成の口） | – | インデックス作成の進み具合を受け渡しの口の `Progress` に入れる（口が無ければ何もしない） | [インデックス作成のメインフロー](../indexing/flow.md) | インデックス作成（1 ファイルにつき 1 回） |
| `readIndexingProgress` | channel | `@{Phase; Processed; Remaining; Failed; Detail}` / `$null` | インデックス作成の進み具合を受け渡しの口から読む（まだ無ければ `$null`）。画面が 1 秒ごとに呼ぶ（数万行の取り込み一覧を読み直さない） | [インデックス作成の実行](../gui/indexing-run.md#更新の進み具合) | 画面 |
| `requestIndexingStop` | channel | – | 中止を求める（`Stop` を立て、確認を待っていれば取りやめの返事にする） | 同上 | 画面 |
| `answerIndexingPlan` | channel, answer（`@{RetryFailed; IncludeCloud}` / `$null`） | – | 確認のダイアログの返事をインデクサに伝える。`$null` は取りやめ（`Stop` も立てる） | [インデックス作成の実行](../gui/indexing-run.md#インデックス更新の確認ダイアログ) | 画面 |
| `testIndexerRunning` | dir（既定 `$workspace.Dir`） | bool | この `work` でインデックス作成が動いているか（インデクサのミューテックスを取れるかで調べ、取れたらすぐ放す）。画面を使わずに起動したものも分かる | [画面とインデクサの受け渡し](../structure/threads.md#画面とインデクサの受け渡し) | 画面（［設定］） |
| `writeIndexerLog` | text, color | – | インデックス作成の表示内容をログ（`indexing_log.txt`）に書く。画面を使わずに実行したときはコンソールにも出す（color はそのときの色）。取り込みのスレッドでは 1 ファイル分を貯め、司令がまとめて書く | [インデックス作成のメインフロー](../indexing/flow.md) | インデックス作成 |
| `writeZipSizeLimitLog` | exception | – | .docx・.pptx・.xlsx を直接読んでサイズの上限を超えたときの詳細（部品名・大きさ・部品ごとか合計か）をインデックス作成のログにだけ書く（画面・取り込み一覧には出さない）。渡した例外が `ZipSizeLimitException`（`shared/office/office_reader.ps1`）でなければ何もしない | [失敗の原因](../indexing/office-apps.md#失敗の原因describeingesterror) | インデックス作成（extract_office.ps1, indexer_run.ps1） |
| `newIngestPlanRow` | name, path, kind, total, targets, new, updated, pending, lost, failed, cloud, cloudFailed, cloudBytes, cloudFailedBytes | 取り込み予定の 1 行（`[pscustomobject]`） | インデックス 1 件分の取り込み対象の件数を作る（`$ingestPlanColumns` と同じ列） | [取り込み対象の決定](../indexing/target-decision.md#取り込み予定画面の確認に出す件数) | インデックス作成 |
| `getIndexingState` | since, path | [元のファイルの特定・画面](search.md#元のファイルの特定画面) を参照 | 取り込み一覧の状態ごとの件数など | [続き](search.md#元のファイルの特定画面) | 画面 |

### StatusLedger（`indexer_state.ps1`。司令のスレッドで作る）

インデクサの司令（`invokeIndexerBody`）が、コンストラクタで受け取った `Workspace` の `StatusFile`・`IngestingFile` を読み書きする。メソッドは、上の表の同じ処理をする関数をそのまま呼ぶ（列と状態の定義は `core/paths.ps1` のまま動かさない）。

| メソッド／プロパティ | 入力 | 出力 | 概要 |
|---|---|---|---|
| `ReadStatus` / `WriteStatus` / `AddRow` | – / folders, rows / row | `readStatusFile` と同じ / – / – | `readStatusFile`・`writeStatusFile`・`addStatusRow` を呼ぶ |
| `ReadIngestingFiles` / `WriteIngestingFiles` / `RemoveIngestingFile` | – / entries / – | `readIngestingFiles` と同じ / – / – | 取り込み中のファイルの記録を読み書き・削除する |
| `RenameIndexName` / `RemoveIndexName` | oldName, newName / name | – | `renameStatusIndexName` / `removeStatusIndexName` を呼ぶ |
| `Failures`（プロパティ）・`AddFailure` | relPath, message | – | 今回の取り込みで失敗したファイル（`@{RelPath; Message}` の並び）を集める |
| `DroppedRows`（プロパティ）・`AddDropped` | relPath | – | 取り込みの直前に元のファイルが無くなった相対パスを集める |

### PendingPublish（新規 `tebunko/indexer/pending_publish.ps1`。司令のスレッドで作る）

フォルダごとの取り込み中の数とまだ渡していない数を数え、本文インデックスに書き出してよいフォルダを決めて返すだけ（書き出し `publishIndexFolders` は司令が呼ぶ）。

| メソッド | 入力 | 出力 | 概要 |
|---|---|---|---|
| `Add` | relPath, removed | – | 取り込んだ・無くなった元のファイルのフォルダを書き出し待ちにする（`removed` なら、そのファイル名も記録する） |
| `MarkFolder` | folder | – | 前回のインデックス作成で本文インデックスに入れていないフォルダを、無くなったファイルの記録無しで書き出し待ちにする |
| `AddPending` | folder | – | 取り込み対象を数えるとき、そのファイルの分だけ「まだ渡していない数」を 1 つ足す |
| `Dispatch` | folder | – | ファイルを 1 つ渡す（まだ渡していない数を 1 減らし、取り込み中の数を 1 増やす） |
| `Skip` | folder | – | ファイルを 1 つ、渡さずに済ませる（元のファイルが無くなった等。まだ渡していない数だけ減らす） |
| `Complete` | folder | – | 取り込み中のファイルが 1 つ終わる（成功・失敗・後回しのどれでも。取り込み中の数を 1 減らす） |
| `TakeFlushable` | – | フォルダ → 無くなったファイル名の集まり（`Dictionary[string,object]`） | 取り込み中・まだ渡していない数がどちらも 0 のフォルダを、書き出し待ちから取り出して返す |
| `TakeAll` | – | 同上 | 取り込み中・まだ渡していない数を見ずに、書き出し待ちのフォルダをすべて取り出す（インデックス作成の終わり） |

### IndexingReporter（新規 `tebunko/indexer/indexing_reporter.ps1`。司令のスレッドで作る）

受け渡しの口（`newIndexerChannel`）を持ち、進み具合を書く・画面の確認を待つ。段階（`${indexingPhase*}`）はメソッドの中で読まず、呼び出し元（司令）から引数で受け取る。ログ（`writeIndexerLog`・`$script:indexerLog`）は部品のまま、このクラスには入れない。

| メソッド | 入力 | 出力 | 概要 |
|---|---|---|---|
| `Progress` | phase, processed, remaining, failed, detail | – | `writeIndexingProgress` を呼ぶ |
| `WaitForApproval` | phase, plan, targetCount, failedCount, timeoutMinutes | `@{RetryFailed}` / `$null` | 取り込む内容を画面に渡し、開始の返事を待つ（旧 `indexer_plan.ps1` の `waitForIndexingApproval`） |

## 取り込み直すかの判断（`tebunko/indexer/indexer_decide.ps1`）

| 関数 | 入力 | 出力 | 概要 | 詳細 | 使用元 |
|---|---|---|---|---|---|
| `getExtractVersion` | path | int | ファイルの形式（拡張子）の今の抽出版（読み取る内容の版。`.xlsx` `.xlsm` はグラフの項目名を読まず、グラフ・SmartArt・ヘッダー・フッターを読む 4、`.docx` `.docm` `.pptx` `.pptm` はグラフの項目名を読まない 3、旧形式（`.doc` `.ppt`）など図形・コメントを読む形式は 2、ほかは 1） | [取り込み対象の決定](../indexing/target-decision.md#取り込み対象の決定createtargetlist) | testExtractOutdated、インデックス作成（取り込み一覧の抽出版の列） |
| `testExtractOutdated` | row | bool | 取り込み一覧の行が、今の抽出版より前の版で取り込んだものか（抽出版の列が空なら 1） | 同上 | getIngestDecision |
| `getIngestDecision` | old（前回の行 / `$null`）, updated, size, indexComplete | `@{Ingest; Reason}` | 取り込むかどうかと理由（`done` / `failed` / `new` / `updated` / `pending` / `lost` / `outdated`）。更新日時・サイズが同じでも、TSV が欠けていれば `lost`、前の抽出版なら `outdated` で取り込み直す | 同上 | createTargetList |
| `getIngestLane` | relPath | string | 取り込むレーン（`.xls*` は `Excel`、`.doc` は `Word`、`.ppt` は `PowerPoint`、それ以外（`.docx`・`.docm`・`.pptx`・`.pptm`、テキストの拡張子）は既定の `Reader`） | [取り込みの並列化](../indexing/parallel.md) | インデックス作成（司令） |
| `getOfficeLane` | relPath | string | 読み取りのレーンで Office が要ると分かったファイルの回し先（`.ppt*` は `PowerPoint`、ほかは `Word`） | 同上 | インデックス作成（司令） |

## インデックス名とインデックスの管理（`tebunko/index/index_name.ps1`・`index_store.ps1`）

| 関数 | 入力 | 出力 | 概要 | 詳細 | 使用元 |
|---|---|---|---|---|---|
| `newIndexName` | folderPath, usedNames（HashSet・配列・文字列・`$null`） | string | インデックス名（フォルダ名・ドライブ名・共有名。重複すれば `名前(2)`…） | [取り込み一覧](../indexing/ingest-list.md) | assignIndexNames |
| `assignIndexNames` | targetFolders, previousFolders（readStatusFile の Folders） | `@{Path; Enabled; Name}` の配列 | 設定の名前（getTargetFolders の Name）をそのまま使う。名前が無ければ、前回の取り込み一覧の同じフォルダの名前、それも無ければフォルダ名から重複しない名前を作る | 同上 | インデックス作成 |
| `splitIndexRelPath` | relPath | `@{Name; Rest}` | `work\content_index` からの相対パスを、先頭のインデックス名と残りに分ける | 同上 | インデックス作成, resolveSourcePath |
| `testSourceNameRecordable` | name | bool | 元のフォルダを設定に記録できる名前か（空・前後に空白があれば偽。`setIndexSourceFolder` の入口と、確認ダイアログの文言で使う） | [元のファイルを開く](../gui/open-file.md) | 判断 |
| `testIndexName` | name, usedNames | string（使えれば空） | インデックス名として使えるか調べ、使えない理由を返す（空・前後の空白・255 文字超・使えない文字・末尾の `.`・Windows の予約語・ほかと重複） | [追加・編集のダイアログ](../gui/index-tab.md#追加編集のダイアログ) | 画面 |
| `getIndexNameMap` | path（既定 `$workspace.StatusFile`） | Dictionary（インデックス名 → フォルダパス） | 取り込み一覧のインデックス名からクロール対象フォルダを引く表。クロール対象フォルダの行は先頭にあるため、見出し行まで読んで打ち切る | 同上 | resolveSourcePath, 画面 |
| `getIndexStats` | rows（readStatusFile の Rows） | 名前 → `@{Total; Done; Pending; Failed; LastIngested}` | 取り込み一覧の行をインデックス名ごとに集計する（一覧の「ファイル」「最終取り込み」） | [一覧の列](../gui/index-tab.md#一覧の列) | 画面（getIndexingState 経由） |
| `renameIndex` | oldName, newName, dir（既定 `$workspace.IndexDir`）, statusPath, settingsPath（既定 `$settingsFile`） | – | インデックス名を変える。`work\content_index\<旧名>` を改名し、取り込み一覧の記録（`renameStatusIndexName`）も書き換えるため、**インデックスは作り直さない**。移動先が既にあれば例外。旧名・新名の下の `searchExcludes` も消す（`removeSearchExcludesUnder`。付け替えず、外したフォルダは検索対象に戻る） | 同上 | 画面（［編集…］） |
| `removeIndex` | name, dir（既定 `$workspace.IndexDir`）, statusPath, settingsPath（既定 `$settingsFile`） | – | インデックスを削除する。`work\content_index\<名前>` を中身ごと削除し、取り込み一覧からもその記録を取り除く（`removeStatusIndexName`）。そのインデックスの下の `searchExcludes` も消す（`removeSearchExcludesUnder`） | 同上 | 画面（［削除］） |
| `removeIndexes` | names, dir, statusPath, settingsPath（`removeIndex` と同じ） | `@{Name; Ok; Reason}` の配列（名前ごと） | 選んだインデックスをまとめて削除する。`removeIndex` を 1 つずつ呼び、1 つ失敗しても残りを続ける（`Reason` は失敗の理由。空の名前は失敗にする）。画面は別スレッドの仕事の中で呼ぶ | 同上 | 画面（［削除］） |
| `getSearchIndexes` | dir（既定 `$workspace.IndexDir`）, statusPath, settingsPath | `@{Name; Path; SourcePath}` の配列 | インデックスの一覧（`work\content_index` 直下のフォルダ 1 つがインデックス 1 つ）。並びは［インデックス管理］の一覧と同じで、一覧に無いもの（コピーしたインデックスなど）は名前順で後ろ。`SourcePath` は元のフォルダ（分からなければ空） | [インデックスの一覧](../search/index.md#インデックスの一覧getsearchindexes) | 画面（検索対象のツリー） |
| `getIndexTreeData` | dir, paths（展開・チェックを外したフォルダ） | `@{State; Message; Root; Sources; ...}` | 検索対象ツリーの材料（ワークスペースの状態・インデックスの一覧・展開したフォルダの子）を 1 回で集める。ワークスペースがネットワークにあるときの裏の仕事で呼ぶ（届かなければ State と Message で知らせる） | [検索対象のツリー](../gui/search-tree.md) | 画面（検索対象のツリー） |
| `getIndexFolderChildren` | dir | `@{HasFiles; Folders; Error}` | インデックスのフォルダの子（本のフォルダを除いた名前順）。読めなければ Error に文面を入れる | 同上 | 画面（ツリーの展開） |

## インデックスのエクスポート・インポート（`tebunko/index/index_archive_rules.ps1`・`index_archive.ps1`）

詳細は [インデックスのエクスポート・インポート](../index-data/format.md#インデックスのエクスポートインポート)。

| 関数 | 入力 | 出力 | 概要 | 詳細 | 使用元 |
|---|---|---|---|---|---|
| `exportIndex` | name, destPath, ws（既定 `$workspace`）, settingsPath（既定 `$settingsFile`） | `@{Name; Path; Files; Bytes}` | 1 つのインデックスを 1 つの zip に書き出す。インデックス作成のロックを取り、入れる前の TSV が残っていれば例外。`<destPath>.tmp` に書いてから置き換える | 同上 | 画面（［エクスポート…］） |
| `exportIndexToFolder` | name, folder, ws, settingsPath | `exportIndex` と同じ | 書き出し先のフォルダを確かめ（無ければ `書き出し先のフォルダが見つかりません：<フォルダ>` の例外）、そのフォルダにある zip の名前から `getExportFileName` で重ならない名前を決め、`exportIndex` で書き出す | 同上 | 画面（［エクスポート…］の別スレッド） |
| `exportIndexes` | names, destination（フォルダ）, ws, settingsPath | `@{Name; Ok; Reason; Path}` の配列（名前ごと） | 選んだインデックスを、書き出し先のフォルダにインデックス 1 つにつき zip 1 つで書き出す。`exportIndexToFolder` を 1 つずつ呼び、1 つ失敗しても残りを続ける（`Reason` は失敗の理由、`Path` は書き出した zip。失敗なら空）。zip の形・ファイル名は 1 つずつの書き出しと同じ | 同上 | 画面（［エクスポート…］の別スレッド） |
| `readIndexArchiveInfo` | zipPath | `@{IndexName; SourceFolder; Files; Bytes; FormatVersion; AppVersion}` | zip の目録を読んで確かめる（インポートはしない）。インポートの確認ダイアログの既定値に使う | 同上 | 画面（［インポート…］） |
| `getImportArchiveInfo` | zipPath | `readIndexArchiveInfo` の結果に `SourceExists`（元のフォルダがあるか。ローカルのドライブでなければ有無を調べず `$null`）を足したもの | インポートの確認ダイアログ用。目録の元のフォルダが共有フォルダでも、接続を待たない | 同上 | 画面（［インポート…］の別スレッド） |
| `importIndex` | zipPath, collisionMode（`Rename`/`Overwrite`/`Cancel`）, name, sourceFolder, ws, settingsPath | `@{Name; SourcePath; Enabled; Files; Bytes; Warnings}`（`Cancel` は `$null`） | zip から 1 つのインデックスをインポートする。前の版のワークスペースの片付け・目録の確かめ・設定への登録・`content_index\<名前>\` の入れ替え・取り込み一覧の書き直しを行う。途中で失敗したら逆の操作で戻す | 同上 | 画面（［インポート…］） |
| `getWorkspaceFreeSpace` | root | long（調べられなければ `$null`） | ドライブのルートの空き容量。UNC など `DriveInfo` にできないパスは例外にせず `$null` | 同上 | importIndex |
| `testImportFreeSpace` | totalBytes, root, getFreeSpace | string（足りていれば空） | 空き容量が「合計 + 1GB」に足りなければ理由を返す。調べられない（`$null`・例外）ときは確かめない | 同上 | importIndexCore |
| `expandImportArchive` / `registerImportedIndexInSettings` / `swapInImportedIndexDir` / `rewriteStatusForImport` | 同上（`importIndexCore` の手順 6〜9） | – | インポートの手順を 1 つずつ。失敗したら自分で戻す（`restoreImportedSettings`・`restoreSwappedIndexDir`）。取り込み一覧は `readStatusFile` → `writeStatusFile` で書く | 同上 | importIndexCore |
| `getImportIndexName` | suggestedName, usedNames, collisionMode | string（`$null` は Cancel） | 同じ名前のインデックスがあるときの扱いから、インポートで使う名前を決める（`Rename` は `newIndexName` と同じ決まり） | 同上 | importIndex, 画面 |
| `getExportFileName` | indexName, usedNames（保存先フォルダの既存のファイル名）, now | string | エクスポートの既定のファイル名（`<インデックス名>_インデックス_<yyyyMMdd>.zip`。重なれば `(2)`…） | 同上 | 画面（［エクスポート…］） |
| `testIndexArchiveManifest` | manifest, entryNames, manifestBytes | string（受け付けられれば空） | zip の目録・エントリーが受け付けられるかを調べる（形式の版・zip slip・重なり・大きさ） | 同上 | openIndexArchive |
