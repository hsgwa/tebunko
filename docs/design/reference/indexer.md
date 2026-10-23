# 部品ごとの関数（インデックス作成）

扱うこと: 取り込み一覧の読み書きとインデックス作成のログ（`status_file.ps1`・`diagnostics.ps1`）、取り込み中のファイル・画面との受け渡しの口（`indexer_state.ps1`）、取り込み直すかの判断（`indexer_decide.ps1`）、インデックス名とインデックスの管理（`index_name.ps1`・`index_store.ps1`）の関数一覧。扱わないこと: 設定・ファイル操作（[部品ごとの関数（設定・ファイル）](settings.md)）、TSV・本文インデックス（[部品ごとの関数（TSV と本文インデックス）](tsv.md)）。先に読むページ: [部品から関数一覧を引く](index.md)。

## 取り込み一覧の読み書き・インデックス作成のログ（`tebunko/core/status_file.ps1`・`diagnostics.ps1`）

| 関数 | 入力 | 出力 | 概要 | 詳細 | 使用元 |
|---|---|---|---|---|---|
| `newStatusRow` | 相対パス, 更新日時, サイズ, 状態, TSV数, 取り込み日時, エラー, 抽出版（2 番目以降は省略可） | 行オブジェクト | 取り込み一覧の 1 行を作る | [取り込み一覧](../indexing/ingest-list.md) | インデックス作成 |
| `toStatusLine` | row | string | 取り込み一覧の 1 行を文字列にする（エラーのタブ・改行はスペース）。`writeStatusFile` は速さのため同じ整形をその場に展開しているため、変えるときは両方を直す（テストで同じ結果になることを確かめている） | [取り込み一覧](../indexing/ingest-list.md) | addStatusRow |
| `readStatusFile` | path（既定 `$workspace.StatusFile`） | `@{Folders; Rows; Rejected; LegacyFolders}` | 取り込み一覧を読む。名前・相対パスが使えない行（`testStoredIndexName`・`testIndexRelPath` と同じ決め。予約名は通す。名前だけの行は使えない。数万行を読むため 1 行ごとの関数呼び出しはせず、その場で確かめる）は読み飛ばし、数を Rejected で返す。最後が `.` や空白の名前など、今の決まりに合わないが外へ出ない前の版のクロール対象フォルダは LegacyFolders に分ける（設定から外れたインデックスとして確認に出す）。Folders はクロール対象フォルダ `@{Path; Name}`（Name = インデックス名）、Rows は相対パス（`インデックス名\フォルダからの相対パス`。大文字・小文字を区別しない）→ 行。同じ相対パスは後の行を優先し、列数の合わない行は無視。インデックス作成中に読んでも書き込みを妨げない共有モードで開く。無ければ空 | [取り込み一覧](../indexing/ingest-list.md) | インデックス作成・画面 |
| `writeStatusFile` | folders（`@{Path; Name}` の配列）, rows, path（既定 `$workspace.StatusFile`） | – | 取り込み一覧を書き出す（先頭にクロール対象フォルダの行、1 ファイル 1 行。一時ファイルに書いてから置き換える） | [取り込み一覧](../indexing/ingest-list.md) | インデックス作成 |
| `getIndexNameMap` | path（既定 `$workspace.StatusFile`） | Dictionary（インデックス名 → フォルダパス） | 取り込み一覧のインデックス名からクロール対象フォルダを引く表。クロール対象フォルダの行は先頭にあるため、見出し行まで読んで打ち切る | [取り込み一覧](../indexing/ingest-list.md) | resolveSourcePath, 画面 |
| `writeIndexerLog` | text, color | – | インデックス作成の表示内容をログ（`indexing_log.txt`）に書く。画面を使わずに実行したときはコンソールにも出す（color はそのときの色）。取り込みのスレッドでは 1 ファイル分を貯め、司令がまとめて書く | [インデックス作成のメインフロー](../indexing/flow.md) | インデックス作成 |

## 取り込み中のファイル・画面との受け渡しの口（`tebunko/indexer/indexer_state.ps1`）

| 関数 | 入力 | 出力 | 概要 | 詳細 | 使用元 |
|---|---|---|---|---|---|
| `describeIngestError` | exception | string | 取り込みの例外から、取り込み一覧のエラー列・画面に表示する失敗の原因を作る | [失敗の原因](../indexing/office-apps.md#失敗の原因describeingesterror) | インデックス作成 |
| `addStatusRow` | row, path（既定 `$workspace.StatusFile`） | – | 取り込み一覧の末尾に 1 行追記する | [取り込み一覧](../indexing/ingest-list.md) | インデックス作成 |
| `readStatusLines` | path（既定 `$workspace.StatusFile`） | 行の配列 | 取り込み一覧を共有を許して 1 行ずつ読む | – | renameStatusIndexName, removeStatusIndexName |
| `renameStatusIndexName` / `removeStatusIndexName` | oldName, newName, path / name, path | – | 取り込み一覧のインデックス名を書き換える / その記録を取り除く。行の順序と内容はそのまま保つ（`readStatusLines` で読み、`writeTextLinesAtomic` で置き換える） | [名前とフォルダパスの編集](../gui/index-tab.md#名前とフォルダパスの編集) | IndexCatalog.RenameIndex, RemoveIndex |
| `readStatusRelPaths` | path（既定 `$workspace.StatusFile`） | 相対パスの配列 | 取り込み一覧のデータ行の相対パスを、ファイルの順に返す（同じ相対パスの追記も別の行として数える）。ジャーナルの「この回の行」の数え方に使う。無ければ空 | [ジャーナルの形](../indexing/ingest-list.md#ジャーナルの形) | StatusLedger.ReadRelPaths |
| `convertToIngestJournalLine` | kind, fields | string | ジャーナルの 1 記録（版・回・書・始・持・未・出・終）を 1 行の文字列にする。列のタブ・改行はスペースにする | 同上 | インデックス作成 |
| `newIngestJournalHead` | version, stamp, rowCount, carried, unpublished | string の配列 | 新しい回のジャーナルの頭の行（版・回・持ち越した仕事の「持」・書き出せていないフォルダの「未」の順）を組み立てる（判断層） | 同上 | インデックス作成 |
| `convertFromIngestJournal` | lines, statusRelPaths, 上限 | `@{Valid; Finished; Aligned; Folders; Resume}` | 前回のジャーナルの行と取り込み一覧の相対パスから、始めの復旧の内容（先に書き出すフォルダ・やり直す仕事）を決める。ファイルには触らない（判断層） | 同上 | インデックス作成 |
| `recoverIngestJournal` | ledger, journal, pending, channel, reporter | `@{Resume; Unpublished}` | 始めの復旧。前回のジャーナル（と前の版の `ingesting.txt`）から、やり直す仕事を決め、書き出していないフォルダを先に書き出す。ジャーナルが読めないときは今ある書き出し前のフォルダを書き出す。書き出せなかった・中止で飛ばしたフォルダは Unpublished に入れ、次の回のジャーナルへ「未」で持ち越す | 同上 | インデックス作成 |
| `writeStatusWithJournal` | journal, ledger, folders, rows, keptDropped, previous, [-Finish] | – | 取り込み一覧を書き直す。前に「書」を足し、-Finish のときは書き直したあとに「終」を足す。最後にジャーナルを閉じる | 同上 | インデックス作成 |
| `getLegacyIngestingPath` / `readLegacyIngesting` / `removeLegacyIngesting` | – | パス / `@{RelPath; Count}` の配列 / – | 前の版の `ingesting.txt` を始めに 1 度だけ読んで、ジャーナルの復旧に足し、読んだあとに消す（`ingest_journal.ps1`） | 同上 | インデックス作成 |
| `newIndexerChannel` | retryFailed, confirmTargets, workers（既定 -1）, onlyNames（既定は空）, includeCloud（既定 `$false`） | 受け渡しの口（`[hashtable]::Synchronized`） | 画面とインデクサの受け渡しの口を作る（`RetryFailed`・`ConfirmTargets`・`Workers`・`OnlyNames`（更新するインデックス名。空ならチェックの付いたものすべて）・`OnlySkipped`（`OnlyNames` のうち更新できなかった名前と理由）・`Progress`・`Stop`・`Plan`・`Answer`・`Answered`・`Error`・`ExitCode`・`Notice`・`Postponed`・`IncludeCloud`（クラウドにだけあるファイルもダウンロードして取り込むか）・`CloudSkipped`（ダウンロードせずに残した件数）・`OfficePids`）。`Workers` は -1 で設定・コア数から決める、0 で司令のスレッドで取り込む（テスト） | [画面とインデクサの受け渡し](../structure/threads.md#画面とインデクサの受け渡し) | 画面・indexer.ps1 |
| `writeIndexingProgress` | phase, processed, remaining, failed, detail, channel（既定はいま動いているインデックス作成の口） | – | インデックス作成の進み具合を受け渡しの口の `Progress` に入れる（口が無ければ何もしない） | [インデックス作成のメインフロー](../indexing/flow.md) | インデックス作成（1 ファイルにつき 1 回） |
| `readIndexingProgress` | channel | `@{Phase; Processed; Remaining; Failed; Detail}` / `$null` | インデックス作成の進み具合を受け渡しの口から読む（まだ無ければ `$null`）。画面が 1 秒ごとに呼ぶ（数万行の取り込み一覧を読み直さない） | [インデックス作成の実行](../gui/indexing-run.md#更新の進み具合) | 画面 |
| `requestIndexingStop` | channel | – | 中止を求める（`Stop` を立て、確認を待っていれば取りやめの返事にする） | 同上 | 画面 |
| `answerIndexingPlan` | channel, answer（`@{RetryFailed; IncludeCloud}` / `$null`） | – | 確認のダイアログの返事をインデクサに伝える。`$null` は取りやめ（`Stop` も立てる） | [インデックス作成の実行](../gui/indexing-run.md#インデックス更新の確認ダイアログ) | 画面 |
| `enterIndexerLock` | dir（既定 `$workspace.Dir`）, attempts（既定 3。画面のスレッドでは 1） | `@{Mutex; Stream; Acquired}` | インデックス作成の鍵（ミューテックス → 鍵のファイルの順）を取る。取れなければ `Acquired` が `$false` でどちらも持たない。鍵のファイルで例外になったときはミューテックスを放して投げ直す。呼ぶのは `invokeIndexer`・`exportIndex`・`importIndex` | [インデックス作成のメインフロー](../indexing/flow.md) | インデックス作成・エクスポート・インポート |
| `exitIndexerLock` | lock（`enterIndexerLock` の戻り値） | なし | 鍵を、取った順の逆（鍵のファイル → ミューテックス）に放す。2 回呼んでも例外にならない | [インデックス作成のメインフロー](../indexing/flow.md) | 同上（`finally`） |
| `testIndexerRunning` | dir（既定 `$workspace.Dir`） | `@{Running; Error}` | この `work` でインデックス作成が動いているか（ミューテックスと鍵のファイルが取れるかで調べ、取れてもすぐ放す。待たずに 1 回）。画面を使わずに起動したもの・ほかのセッション・ほかの PC も分かる。フォルダが無ければミューテックスだけで決める。鍵のファイルを調べて例外になったときは `Running` を `$true`、`Error` に理由を入れる（始められないとして扱う） [画面とインデクサの受け渡し](../structure/threads.md#画面とインデクサの受け渡し) | 画面（［設定］） |
| `writeZipSizeLimitLog` | exception | – | .docx・.pptx・.xlsx を直接読んでサイズの上限を超えたときの詳細（部品名・大きさ・部品ごとか合計か）をインデックス作成のログにだけ書く（画面・取り込み一覧には出さない）。渡した例外が `ZipSizeLimitException`（`shared/office/office_reader.ps1`）でなければ何もしない | [失敗の原因](../indexing/office-apps.md#失敗の原因describeingesterror) | インデックス作成（extract_office.ps1, indexer_run.ps1） |
| `newIngestPlanRow` | name, path, kind, total, targets, new, updated, pending, lost, failed, cloud, cloudFailed, cloudBytes, cloudFailedBytes | 取り込み予定の 1 行（`[pscustomobject]`） | インデックス 1 件分の取り込み対象の件数を作る（`$ingestPlanColumns` と同じ列） | [取り込み対象の決定](../indexing/target-decision.md#取り込み予定画面の確認に出す件数) | インデックス作成 |
| `getIndexingState` | since, path | [元のファイルの特定・画面](search.md#元のファイルの特定画面) を参照 | 取り込み一覧の状態ごとの件数など | [続き](search.md#元のファイルの特定画面) | 画面 |

### StatusLedger（`indexer_state.ps1`。司令のスレッドで作る）

インデクサの司令（`invokeIndexerBody`）が、コンストラクタで受け取った `Workspace` の `StatusFile` を読み書きする。メソッドは、上の表の同じ処理をする関数をそのまま呼ぶ（列と状態の定義は `core/paths.ps1` のまま動かさない）。

| メソッド／プロパティ | 入力 | 出力 | 概要 |
|---|---|---|---|
| `ReadStatus` / `WriteStatus` / `AddRow` | – / folders, rows / row | `readStatusFile` と同じ / – / – | `readStatusFile`・`writeStatusFile`・`addStatusRow` を呼ぶ |
| `ReadRelPaths` | – | `readStatusRelPaths` と同じ | 取り込み一覧の相対パスの並びを読む |
| `RenameIndexName` / `RemoveIndexName` | oldName, newName / name | – | `renameStatusIndexName` / `removeStatusIndexName` を呼ぶ |
| `Failures`（プロパティ）・`AddFailure` | relPath, message | – | 今回の取り込みで失敗したファイル（`@{RelPath; Message}` の並び）を集める |
| `DroppedRows`（プロパティ）・`AddDropped` | relPath | – | 取り込みの直前に元のファイルが無くなった相対パスを集める |

### PendingPublish（新規 `tebunko/indexer/pending_publish.ps1`。司令のスレッドで作る）

フォルダごとの取り込み中の数とまだ渡していない数を数え、本文インデックスに書き出してよいフォルダを決めて返すだけ（書き出し `publishIndexFolders` は司令が呼ぶ）。

| メソッド | 入力 | 出力 | 概要 |
|---|---|---|---|
| `Add` | relPath, removed | – | 取り込んだ・無くなった元のファイルのフォルダを書き出し待ちにする（`removed` なら、そのファイル名も記録する） |
| `MarkFolder` | folder | – | 前回のインデックス作成で本文インデックスに入れていないフォルダを、無くなったファイルの記録無しで書き出し待ちにする |
| `Restore` | folder, names | – | 前の回に外し終えていなかった、無くなった元のファイル名を、書き出し待ちと `Unpublished` の両方に戻す |
| `MarkPublished` | folders, taken | – | 本文インデックスのファイルを書き終えたフォルダの、そのとき書き出した名前（`taken`。`TakeAll` などの戻り値）だけを、`Unpublished`（外し終えていない削除。`TakeAll` などでは消えない）から外す。同じ回に書き出しを飛ばした名前は残る |
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

### IngestDispatcher（新規 `tebunko/indexer/ingest_dispatcher.ps1`。司令のスレッドで作る）

取り込みのワーカー（1 種類。どれも STA）・先読み・取り込み中の管理だけを持つ。取り込み一覧・本文インデックスの書き出し待ち・進み具合は触らず、司令が戻り値を見て書く。取り込みのスレッドとは hashtable と `BlockingCollection` で受け渡す（クラスのインスタンスは渡さない）。

| メソッド／プロパティ | 入力 | 出力 | 概要 |
|---|---|---|---|
| コンストラクタ | targets, apps, bookFolders, folderByName, carried, options | – | 取り込み対象（apps はファイルごとに最初に要る Office のアプリ。要らなければ空）と設定を受け取る。必要な設定（`LookAhead`・`WorkerScript`・`PostponedReasonInUse`・`PostponedReasonRunning`）が足りなければ、その名前を書いて失敗する。省略できる `TestPath`（`{ param($path, $kind) }`。`$kind` は `File` か `Folder`）で、元のファイル・フォルダの有無の調べ方を差し替えられる |
| `Start` | workerCount, settings | – | ワーカーごとの列を持つ入れ物を作る（スレッドは、そのワーカーに初めて渡すときに始める。持ち主のアプリ・渡してある数・諦めたかの状態もワーカーごとに持つ）。0 のときは作らず、司令のスレッドで取り込む |
| `BeginPass` | – | – | 渡す繰り返しの 1 回分を始める（先読みの範囲を数え直す） |
| `DispatchNext` | – | `@{Kind=Abandoned／Vanished／Dispatched; ...}` / `$null` | 並びの先から、`selectIngestWorker` で渡し先のワーカーを決め、空きがあるファイルを 1 つ選ぶ（渡し先は `Worker`。ワーカーを使うときは `Task.Office` に Office が要るかも入れる） |
| `Submit` | step | – | 選んだファイルをワーカーの列に渡す |
| `SetInlineResult` | result | – | スレッドを使わないとき、司令が自分のスレッドで取り込んだ結果を `InlineResult` に置く（次の `Receive` が返して空にする） |
| `Receive` | – | `@{Kind=Completed／Postponed／RerouteAbandoned; ...}` / `$null` | 結果を 1 つ受け取る |
| `TakeAbandonedWorker` | – | `@{Worker; Entries}` / `$null` | 諦めたワーカーを 1 つ取り出す（取り込み中・列のファイルといっしょに返し、取り込み中から外す） |
| `StopIfRequested` | requested（bool） | bool | 中止を求められ、渡していないファイルがあるとき、今回初めて止めた場合だけ `$true` を返し、以後は渡さない（残りの件数は `Remaining`） |
| `GetInflightCount` | relPath | 件数 | 取り込み中のファイルを続けて取り込みを始めた回数（今の分を含む。取り込み中でなければ 0。ジャーナルの「始」に書く） |
| `Stop` | – | – | 列を閉じ、諦めていないワーカーが終わるのを待ってから片づける。諦めたワーカーは待たず、止める依頼だけを出す |
| `LastResult`（プロパティ） | – | 結果 / `$null` | 直前の `Receive` が受け取った結果（回し直し・中止も含む。諦めたワーカーの結果は捨てるので `$null`）。司令が種類ごとの時間を足すのに使う |
| `Pool`・`UsesThreads`・`Stopped`・`FolderLost`・`Remaining`・`Undispatched`（プロパティ） | – | – | プール・スレッドを使うか・中止・フォルダ消失・残り・まだ渡していないファイル |

### IngestTiming（新規 `tebunko/indexer/ingest_timing.ps1`。司令のスレッドで作る）

段階ごとの時間と、取り込みの種類（拡張子）ごとの時間を足し合わせるだけ（ファイルにも画面にも触らない）。書くのは司令（`writeStageTimingLog`）。時刻は Stopwatch だけ。

| メソッド | 入力 | 出力 | 概要 |
|---|---|---|---|
| `Start` / `Mark` / `Lap` | – / – / stage | – | 全体の計りを始める / 区切りの印を今に置く（その間はどの段にも足さない） / 前の印から今までを stage に足して印を置く |
| `AddStage` | stage, ms | – | 段階に足す（同じ段階は足し合わせる） |
| `Finish` | [ms] | – | 全体の時間を決める（画面の返事を待った `Wait` は引く） |
| `AddFile` | kind, outcome（Ok／Failed／Reroute／Postponed）, copyMs, extractMs, publishMs | – | 取り込みの結果 1 件の時間を種類に足す |
| `AddPostponed` | kind | – | 結果を受け取らずに後回しにした分の件数だけを足す（時間は足さない） |
| `GetOutcome` / `GetKind`（static） | result / relPath | string | 結果を Ok・Reroute・Postponed・Failed に分ける / 拡張子（小文字） |
| `ToHashtable` | – | `@{ Stages; Kinds }` | 受け渡しの口の `StageTimings`。段は秒、種類ごとは件数・秒・成功の 1 件の平均 ms。その他は全体から段の合計を引いた残り |
| `FormatLogLines` | – | string の配列 | インデックス作成ログの行（段階の 1 行と、抽出の長い順の種類ごとの行） |

司令の側の関数は `indexer_run.ps1` にある。`addStageTiming`（Stopwatch の経過を段に足す）・`addIngestResultTiming`（取り込みの結果の `Timing`〔`invokeTimedIngest`・`publishTsvTimed` が測ったコピー・抽出・置き場への移し〕を種類に足す）・`addPostponedTiming`・`addReceivedTiming`（受け取った結果の時間を足す）・`postponeAbandonedEntries`・`postponeUndispatched`・`dropVanishedFile`（取り込みの司令の分け出し）・`writeStageTimingLog`（回の終わりに受け渡しの口の `StageTimings` に入れ、ログに書く）。`flushPendingPublish` は段階（既定は仕上げのまとめ。起動の中・取り込みの中は別の段）を受け取り、中身は `publishFlushable`。測る側の `getIngestStageTimings`（`tools/perf/ingest_common.ps1`）は口の `StageTimings` を読む（無い版は `$null`）。

### IngestPlanner（`tebunko/indexer/indexer_plan.ps1`。司令のスレッドで作る）

クロール対象フォルダへの名前の割り当てと、クロールの結果から取り込む順番を決める。取り込み直すかの判断は `indexer_decide.ps1` の関数を `createTargetList` 経由で呼ぶ。計画の種類・状態の値はコンストラクタの `kinds`（`Unchecked`・`Missing`・`Dropped`・`StateFailed`。足りなければ例外）で、名前の保存は、`SaveNames` に司令が渡すスクリプトブロック（中身は `IndexCatalog.StoreAssignedNames`）で行う。メソッドは次の順に呼ぶ: `AssignNames` → `SaveNames` → `SelectOnly`（選んだときだけ）→ `Crawl` → `KeepForCancel`（取りやめるとき）か `ApplyFailed` → `ApplyInterrupted` → `FolderByName`。`AssignNames` の前に `Crawl` を呼ぶと例外にする。結果は `Folders`・`Previous`・`Dropped`・`Only`・`Rows`・`Targets`・`Failed`・`Cloud`（クラウドにだけあるファイル）・`Plan`・`RemovedPaths`・`Carried` で読む。

| メソッド | 入力 | 出力 | 概要 |
|---|---|---|---|
| `AssignNames` | targetFolders, status（`@{Folders; Rows; Rejected; LegacyFolders}`） | – | インデックス名を割り当て、前回の一覧と外れたインデックスを覚える（前回の名前に `LegacyFolders` も足して、外れたインデックスを調べる） |
| `SaveNames` | targetFolders, save（割り当てた名前を受け取って保存し、保存した一覧を返すスクリプトブロック） | – | 名前の無いフォルダがあったときだけ `save` を呼び、`Folders` を保存した一覧にする（保存の間に設定から外れた・名前を付けられなかったフォルダは落とす。司令は `IndexCatalog.StoreAssignedNames` を呼ぶ `save` を渡す） |
| `SelectOnly` | onlyNames | 更新できない名前（`@{Name; Reason}` の並び） | 選んだインデックスだけを更新する場合の当てはめ |
| `Extensions`（プロパティ） | – | string[] | `Crawl` が `createTargetList` に渡す、取り込む拡張子の一覧。司令が設定から `getTargetExtensionList` で作って `Crawl` の前に入れる（`$null` なら固定の一覧のすべて） |
| `Crawl` | indexCounts, onFolder | – | フォルダごとに取り込み対象・チェックなし・フォルダなしに分け、`Rows`・`Targets`・`Failed`・`Plan` を組み立てる。`onFolder` は検索の直前に呼ぶ（進み具合を出す。例外は包まずにそのまま伝える） |
| `KeepForCancel` | – | 取り込み一覧に書く行 | 取りやめるとき、取り込み対象の行（とクラウドにだけあって前回未完了の行）を前回の記録のままにする |
| `ApplyFailed` | retry | – | 前回失敗した行を、再取り込みする（`Targets` に足す）かスキップするか |
| `ApplyInterrupted` | entries, interruptLimit | – | 前回強制終了したファイルを最後に回し（`Carried`）、上限に達したものは失敗にする |
| `FolderByName` | – | インデックス名 → フォルダ | 取り込みのワーカーへ渡すための対応 |

## 取り込み直すかの判断（`tebunko/indexer/indexer_decide.ps1`）

| 関数 | 入力 | 出力 | 概要 | 詳細 | 使用元 |
|---|---|---|---|---|---|
| `getExtractVersion` | path | int | ファイルの形式（拡張子）の今の抽出版（読み取る内容の版。`.xlsx` `.xlsm` はグラフの項目名を読まず、グラフ・SmartArt・ヘッダー・フッターを読む 4、`.docx` `.docm` `.pptx` `.pptm` はグラフの項目名を読まない 3、旧形式（`.doc` `.ppt`）など図形・コメントを読む形式は 2、ほかは 1） | [取り込み対象の決定](../indexing/target-decision.md#取り込み対象の決定createtargetlist) | testExtractOutdated、インデックス作成（取り込み一覧の抽出版の列） |
| `testExtractOutdated` | row | bool | 取り込み一覧の行が、今の抽出版より前の版で取り込んだものか（抽出版の列が空なら 1） | 同上 | getIngestDecision |
| `getIngestDecision` | old（前回の行 / `$null`）, updated, size, indexComplete | `@{Ingest; Reason}` | 取り込むかどうかと理由（`done` / `failed` / `new` / `updated` / `pending` / `lost` / `outdated`）。更新日時・サイズが同じでも、TSV が欠けていれば `lost`、前の抽出版なら `outdated` で取り込み直す | 同上 | createTargetList |
| `getIngestApp` | relPath | string | 取り込みに最初に要る Office のアプリ（表（`$officeFileKinds`）の Excel の拡張子は `Excel`、`.doc` は `Word`、`.ppt` は `PowerPoint`。それ以外（`.docx`・`.docm`・`.pptx`・`.pptm`、テキストの拡張子、`.pdf`）は Office が要らず空） | [取り込みの並列化](../indexing/parallel.md) | インデックス作成（司令） |
| `getRerouteApp` | relPath | string | Office を使わずに読めず、Office が要ると分かったファイルの回し先のアプリ（`getOfficeFileKind` が `PowerPoint`（`.pptx` `.pptm` `.ppt`）なら `PowerPoint`、ほかは `Word`） | 同上 | インデックス作成（司令） |
| `selectIngestWorker` | workers（ワーカーごとの `@{Busy; Owns; Abandoned}`）, app, capacity | `@{Kind=Worker／Wait／Abandoned; Index}` | 取り込むファイルの渡し先を決める。Office が要るファイルは、そのアプリの持ち主へ（満杯なら `Wait`、持ち主が諦めていれば `Abandoned`）。持ち主がいなければ、持ち主のいないワーカーのうちいちばん空いているもの。Office が要らないファイルは、持ち主のいないワーカーのうちいちばん空いているもの（無ければ、諦めていないいちばん空いているもの）。すべて諦めていれば `Abandoned` | 同上 | `IngestDispatcher` |

## インデックス名とインデックスの管理（`tebunko/index/index_name.ps1`・`index_store.ps1`）

| 関数 | 入力 | 出力 | 概要 | 詳細 | 使用元 |
|---|---|---|---|---|---|
| `newIndexName` | folderPath, usedNames（HashSet・配列・文字列・`$null`） | string | インデックス名（フォルダ名・ドライブ名・共有名。重複すれば `名前(2)`…） | [取り込み一覧](../indexing/ingest-list.md) | assignIndexNames |
| `assignIndexNames` | targetFolders, previousFolders（readStatusFile の Folders） | `@{Path; Enabled; Name}` の配列 | 設定の名前（getTargetFolders の Name）をそのまま使う（`testStoredIndexName` を通らない名前は無いものとする。予約名の `aux.backup` などは、既にある名前なのでそのまま使う）。名前が無ければ、前回の取り込み一覧の同じフォルダの名前、それも無ければフォルダ名から重複しない名前を作る | 同上 | インデックス作成 |
| `testPathSegments` | segments, lenient（既定 `$false`） | string（問題の文。使えれば空） | `\` で区切った各部分が、空・`.`・`..`・長すぎる・使えない文字でないか。lenient のときは、前の版の名前（最後の `.`・空白など）も通し、パスにすると外へ出るものだけ止める | 同上 | testIndexRelPath, removeDroppedFolders, `index_archive_rules.ps1` の相対パスの確かめ |
| `testSourceBookName` | book | bool | インデックスの「ファイル名」（Book）が開いてよい名前か。フォルダの区切り・`:`・無効な文字・制御文字・向きを変える文字を含まない 1 つの名前で、終わりの `.` と空白を落とした拡張子が取り込み対象（Office・テキスト・PDF）にあるときだけ真 | [元のファイルを開く](../gui/open-file.md) | openSource, openWithShell, readContentIndexPlaces（検索のスレッドにも読み込む） |
| `testIndexRelPath` | relPath, allowNameOnly（既定 `$false`）, checkedNames | string（問題の文。使えれば空） | `インデックス名\フォルダからの相対パス` が置き場の外へ出ないか。先頭の名前は `testStoredIndexName`（予約名は通す）で確かめる。取り込み一覧・取り込み中の記録の行は名前だけを通さない | 同上 | IngestJournal（読み込み） |
| `splitIndexRelPath` | relPath | `@{Name; Rest}` | `work\store\content` からの相対パスを、先頭のインデックス名と残りに分ける | 同上 | インデックス作成, resolveSourcePath |
| `testIndexName` | name, usedNames | string（使えれば空） | 新しく付けるインデックス名として使えるか調べ、使えない理由を返す（フォルダ名の決め `getFolderNameProblemText`（空・前後の空白・255 文字超・使えない文字・末尾の `.`・Windows の予約語）に、ほかと重複を足す）。`newIndexName` が作る名前は必ず通る。`IndexCatalog.RenameIndex` は新しい名前を、これで確かめる | [名前とフォルダパスの編集](../gui/index-tab.md#名前とフォルダパスの編集) | 画面 |
| `testStoredIndexName` | name | string（使えれば空） | 保存してあるインデックス名（設定・取り込み一覧・状態のフォルダから読んだもの）として使えるか調べる。`testIndexName` との違いは、Windows の予約名（`aux.backup`・`CON.txt` など）を拒まないこと（既にある実在のフォルダの名前を空にせず使い続けるため）。置き場の外へ出る名前・最後の `.` や前後の空白・使えない文字は拒む。`IndexCatalog.RenameIndex` は元の名前を、`RemoveIndex` は消す前に名前を、これで確かめる（通らない名前は何も動かさず例外にする） | 同上 | testIndexRelPath, assignIndexNames, `IndexCatalog` |
| `testStatusRowFailed` | row | bool | 取り込み一覧の行を失敗に数えるか。状態が `失敗` の行と、`未取り込み` でエラーの列に理由がある行（利用者の PowerPoint のために取り込めなかったファイル。次の更新で取り込むため状態は `未取り込み` のまま） | [取り込み一覧](../indexing/ingest-list.md) | `getIndexStats`・`getIndexingState` |
| `getWorkspaceIndexFolders` | ws | `@{Name; Path}` の配列 | ワークスペースに残っているインデックスを名前と前回の元のフォルダの組で集める（取り込み一覧のフォルダの行は名前と Path、インデックスのフォルダ（`store\content\`・古い `content_index\`）の直下にだけあって行の無いものは Path を空にする）。設定から外れたまま残してあるものを含む | [名前とフォルダパスの編集](../gui/index-tab.md#名前とフォルダパスの編集) | `getImportUsedIndexNames`・画面（`refreshLeftoverIndexes`） |
| `getIndexStats` | rows（readStatusFile の Rows） | 名前 → `@{Total; Done; Pending; Failed; LastIngested}` | 取り込み一覧の行をインデックス名ごとに集計する（一覧の「ファイル」「最終取り込み」。`未取り込み` でエラーがある行は `Failed` に数え、`Pending` に数えない） | [一覧の列](../gui/index-tab.md#一覧の列) | 画面（getIndexingState 経由） |
| `getSearchIndexes` | dir（既定 `$workspace.IndexDir`）, statusPath, settingsPath | `@{Name; Path; SourcePath}` の配列 | インデックスの一覧（`work\store\content` 直下のフォルダ 1 つがインデックス 1 つ）。並びは［インデックス管理］の一覧と同じで、一覧に無いもの（コピーしたインデックスなど）は名前順で後ろ。`SourcePath` は元のフォルダ（分からなければ空） | [インデックスの一覧](../search/index.md#インデックスの一覧getsearchindexes) | 画面（検索対象のツリー） |
| `getIndexTreeData` | dir, paths（展開・チェックを外したフォルダ） | `@{State; Message; Root; Sources; ...}` | 検索対象ツリーの材料（ワークスペースの状態・インデックスの一覧・展開したフォルダの子）を 1 回で集める。ワークスペースがネットワークにあるときの裏の仕事で呼ぶ（届かなければ State と Message で知らせる） | [検索対象のツリー](../gui/search-tree.md) | 画面（検索対象のツリー） |
| `getIndexFolderChildren` | dir | `@{HasFiles; Folders; Error}` | インデックスのフォルダの子（本のフォルダを除いた名前順）。読めなければ Error に文面を入れる | 同上 | 画面（ツリーの展開） |

### IndexCatalog（`tebunko/index/index_catalog.ps1`。画面のスレッドで作る。別のスレッドの仕事の中では `newIndexCatalog` で作り直す）

インデックスの改名・削除・名前の割り当て・一覧と元のフォルダの記録の保存。インデックス 1 つの情報が置かれている保存先（設定の `targetFolders`・取り込み一覧・インデックスのフォルダと `source_folder.txt`・設定の `searchExcludes`・`store\gram\`）のうち、改名・削除が書き換えるものの手順と順番、途中で失敗したときの扱いだけを持つ（一覧 `targetFolders` と元のフォルダの記録の書き直しは、画面が `SaveTargets`・`WriteSourceFile` で呼ぶ。改名のあとの記録は `RenameIndexWithSource` が書く）。保存先ごとの読み書きは既存の関数と、作る側が渡す `StatusLedger`（取り込み一覧）に任せる。場所はコンストラクタで受け取る `Workspace` と設定ファイルの場所で決める。作り方は `newIndexCatalog`（ワークスペースの場所の文字列から、`StatusLedger` も作って渡す）。以前の `index_store.ps1` の `renameIndex`・`removeIndex`・`removeIndexes` の手順をそのまま移したもの。

| メソッド | 入力 | 出力 | 概要 | 詳細 | 使用元 |
|---|---|---|---|---|---|
| `RenameIndex` | oldName, newName | – | インデックス名を変える（旧 `renameIndex`）。`work\store\content\<旧名>` を改名し、取り込み一覧の記録（`StatusLedger.RenameIndexName`）も書き換えるため、**インデックスは作り直さない**。大文字・小文字だけの変更は一時名を経由する。移動先が既にあれば例外。`store\gram\` の古い名前の分と、旧名・新名の下の `searchExcludes`（`removeSearchExcludesUnder`。付け替えず、外したフォルダは検索対象に戻る）も消す。この 2 つの整理は、失敗しても止めない | [名前とフォルダパスの編集](../gui/index-tab.md#名前とフォルダパスの編集) | 画面（詳細の名前の欄） |
| `RenameIndexWithSource` | oldName, newName, folders（`@{Name; Path}` の配列） | `@{SourceFolderError}` | `RenameIndex` のあと、元のフォルダの記録（`source_folder.txt`。`WriteSourceFile`）を書き直す。記録だけが失敗しても改名は成功にし、書けなかった理由を `SourceFolderError` で返す（成功なら空）。改名の失敗は例外のまま出す | 同上 | 画面（詳細の名前の欄。裏の仕事） |
| `RemoveIndex` | name | – | インデックスを削除する（旧 `removeIndex`）。`work\store\content\<名前>` を中身ごと削除し（消せなければ例外にして、取り込み一覧の記録は残す）、取り込み一覧からもその記録を取り除く（`StatusLedger.RemoveIndexName`）。`store\gram\` のそのインデックスの分と、そのインデックスの下の `searchExcludes` も消す | 同上 | 画面（［削除］） |
| `RemoveIndexes` | names | `@{Name; Ok; Reason}` の配列（名前ごと） | 選んだインデックスをまとめて削除する（旧 `removeIndexes`）。`RemoveIndex` を 1 つずつ呼び、1 つ失敗しても残りを続ける（`Reason` は失敗の理由。空の名前は失敗にする）。画面は別スレッドの仕事の中で呼ぶ | 同上 | 画面（［削除］） |
| `newIndexCatalog` | workspaceDir, settingsPath（既定 `$settingsFile`） | `IndexCatalog` | ワークスペースの場所（文字列）から `IndexCatalog` を作る。取り込み一覧を書く `StatusLedger` も作って渡す。別のスレッドへはインスタンスを渡さず、場所を渡してこれで作り直す | – | 画面・裏の仕事 |
| `AssertReachable` | – | – | ワークスペースに届かないときは例外にする（`assertWorkspaceReachable`）。削除・改名の裏の仕事の先頭で呼び、届かないのに設定だけ変えてしまうことを防ぐ | [ネットワーク上のワークスペース](../gui/network.md) | 画面（裏の仕事） |
| `AssignAndSaveNames` | folders | 保存した一覧 | 名前の決まっていないインデックスに名前を割り当てて設定に保存し、保存した一覧を返す。取り込み一覧の読み込みは排他の外で済ませる | 同上 | 画面（`loadTargets`） |
| `StoreAssignedNames` | assigned（`assignIndexNames` の結果） | 保存した一覧 | 割り当てた名前を設定に保存する（`saveAssignedIndexNames`） | – | インデクサ（司令。`IngestPlanner.SaveNames` に渡す） |
| `SaveTargets` | folders（`@{Name; Path; Enabled}` の配列） | – | インデックスの一覧を設定の `targetFolders` に書く（`writeTargetFolders`） | – | 画面（`saveTargets`） |
| `WriteSourceFile` | folders（`@{Name; Path}` の配列） | – | インデックスのフォルダの `source_folder.txt` を一覧に合わせて書き直す（`writeSourceFolderFile`。インデックスが無いフォルダには書かない） | – | 画面（`applyIndexEdit`） |

## インデックスのエクスポート・インポート（`tebunko/index/index_archive_rules.ps1`・`index_archive.ps1`）

詳細は [インデックスのエクスポート・インポート](../index-data/format.md#インデックスのエクスポートインポート)。

| 関数 | 入力 | 出力 | 概要 | 詳細 | 使用元 |
|---|---|---|---|---|---|
| `getIndexTsvFilesNoFollow` | indexDir | FileInfo の配列 | インデックスのフォルダの下の `.tsv` を、リンクのフォルダへ降りずに集める（`getFilesNoFollow`）。置き場の中にリンクがあっても、先のファイルは zip に入れず、親を指すリンク（輪）でも止まらない。最後まで読めなかったときは `インデックスのフォルダを最後まで読めませんでした` の例外 | 同上 | `exportIndex`・`getIndexExportSize` |
| `exportIndex` | name, destPath, ws（既定 `$workspace`）, settingsPath（既定 `$settingsFile`） | `@{Name; Path; Files; Bytes}` | 1 つのインデックスを 1 つの zip に書き出す。インデックス作成のロックを取り、入れる前の TSV が残っていれば例外。`<destPath>.tmp` に書いてから置き換える | 同上 | 画面（［エクスポート…］） |
| `exportIndexToFolder` | name, folder, ws, settingsPath | `exportIndex` と同じ | 書き出し先のフォルダを確かめ（無ければ `書き出し先のフォルダが見つかりません：<フォルダ>` の例外）、そのフォルダにある zip の名前から `getExportFileName` で重ならない名前を決め、`exportIndex` で書き出す | 同上 | 画面（［エクスポート…］の別スレッド） |
| `exportIndexes` | names, destination（フォルダ）, ws, settingsPath | `@{Name; Ok; Reason; Path}` の配列（名前ごと） | 選んだインデックスを、書き出し先のフォルダにインデックス 1 つにつき zip 1 つで書き出す。`exportIndexToFolder` を 1 つずつ呼び、1 つ失敗しても残りを続ける（`Reason` は失敗の理由、`Path` は書き出した zip。失敗なら空）。zip の形・ファイル名は 1 つずつの書き出しと同じ | 同上 | 画面（［エクスポート…］の別スレッド） |
| `readIndexArchiveInfo` | zipPath | `@{IndexName; SourceFolder; Files; Bytes; FormatVersion; AppVersion}` | zip の目録を読んで確かめる（インポートはしない）。インポートの確認ダイアログの既定値に使う | 同上 | 画面（［インポート…］） |
| `getImportArchiveInfo` | zipPath | `readIndexArchiveInfo` の結果に `SourceExists`（元のフォルダがあるか。ローカルのドライブでなければ有無を調べず `$null`）を足したもの | インポートの確認ダイアログ用。目録の元のフォルダが共有フォルダでも、接続を待たない | 同上 | 画面（［インポート…］の別スレッド） |
| `importIndex` | zipPath, collisionMode（`Rename`/`Overwrite`/`Cancel`）, name, sourceFolder, ws, settingsPath | `@{Name; SourcePath; Enabled; Files; Bytes; Warnings}`（`Cancel` は `$null`） | zip から 1 つのインデックスをインポートする。目録の確かめ（版 1 の zip は断る）・展開・前の版のワークスペースの片付け・設定への登録・`store\content\<名前>\` の入れ替え・取り込み一覧の書き直しを行う。途中で失敗したら逆の操作で戻す | 同上 | 画面（［インポート…］） |
| `readArchiveEntryBytesLimited` | entry, maxBytes | byte[]（maxBytes を超えたら `$null`） | zip のエントリーを、実際に読んだバイト数を数えながら読む。中央ディレクトリの大きさの申告に頼らない（目録の読み取りに使う） | 同上 | openIndexArchive |
| `getWorkspaceFreeSpace` | root | long（調べられなければ `$null`） | ドライブのルートの空き容量。UNC など `DriveInfo` にできないパスは例外にせず `$null` | 同上 | importIndex |
| `testImportFreeSpace` | totalBytes, root, getFreeSpace | string（足りていれば空） | 空き容量が「合計 + 1GB」に足りなければ理由を返す。調べられない（`$null`・例外）ときは確かめない | 同上 | importIndexCore |
| `expandImportArchive` / `registerImportedIndexInSettings` / `swapInImportedIndexDir` / `rewriteStatusForImport` | 同上（`importIndexCore` の手順 5・7〜9） | – | インポートの手順を 1 つずつ。失敗したら自分で戻す（`restoreImportedSettings`・`restoreSwappedIndexDir`）。取り込み一覧は `readStatusFile` → `writeStatusFile` で書く | 同上 | importIndexCore |
| `getImportIndexName` | suggestedName, usedNames, collisionMode | string（`$null` は Cancel） | 同じ名前のインデックスがあるときの扱いから、インポートで使う名前を決める（`Rename` は `newIndexName` と同じ決まり） | 同上 | importIndex, 画面 |
| `getExportFileName` | indexName, usedNames（保存先フォルダの既存のファイル名）, now | string | エクスポートの既定のファイル名（`<インデックス名>_インデックス_<yyyyMMdd>.zip`。重なれば `(2)`…） | 同上 | 画面（［エクスポート…］） |
| `testIndexArchiveManifest` | manifest, entryNames, manifestBytes | string（受け付けられれば空） | zip の目録・エントリーが受け付けられるかを調べる（形式の版・zip slip・重なり・大きさ・エントリーの数と合計の上限） | 同上 | openIndexArchive |
