# ファイルの読み書きの範囲

## 書き込み・削除する場所

書き込みは次の場所に限られる。`tests/meta/safety.Tests.ps1` の「書き込み先が限られていること」が、これらの定義と、ドライブ直下・システムフォルダを指す書き込み先が無いことを確かめる。

| 場所 | 内容 | 定義 |
|---|---|---|
| `work/` 配下（既定は `%USERPROFILE%\Documents\tebunko_ws`。利用者が画面で選んだフォルダ（ワークスペース）にも置ける） | 本文インデックス（`work/content_index/`）、取り込み一覧・インデックス作成ログ・制御用ファイル、検索結果 | `tebunko/core/paths.ps1`（`$workspace`）、`tebunko/core/workspace.ps1`（`Workspace`） |
| `work/office_pids/<PC の鍵>/<PID>.txt` | インデックス作成が起動した Office の PID の記録（起動時の確認で、記録のあるものだけを終了するため）。Office を終了したとき・起動時の確認のときに消す | `tebunko/core/workspace.ps1`（`OfficePidRoot`）、`shared/office/office_process.ps1`（`addOfficeRecord`） |
| `work/tmp/<PC の鍵>/<PID>` 配下（ワークスペースのパスに `[` `]` があるか長すぎるときは作らない。どのファイルも中間 TSV などをこの作業領域に作るため、テキストファイルを含めすべての取り込みをスキップする） | 取り込みの作業領域（原本のコピー・中間 TSV）。インデックス作成の完了時・開始時に空にする | `tebunko/indexer/index_migrate.ps1` の `initTmpDir` |
| `setting.config`（ツールを置いたフォルダの直下。書き込めないときは既定のワークスペース `%USERPROFILE%\Documents\tebunko_ws` の直下） | 画面が保存する設定（クロール対象フォルダ・検索対象インデックス・`work` の置き場所など）。書き込めない場所（Program Files・読み取り専用の共有フォルダ）に置いたときだけ、ワークスペースの既定の場所に置く（[データの置き場所とパスの決め方](../design/structure/data.md)）。保存のときの一時ファイル `setting.config.tmp` と、壊れた設定の退避 `setting.config.broken-<日時>` も同じフォルダに置く | `shared/core/data_dir.ps1`、`tebunko/core/settings.ps1` の `getSettingsFilePath` |
| `startup_error.txt`（ツールを置いたフォルダの直下。書き込めなければ作らない） | 起動そのものに失敗したとき（画面が開く前）の記録。起動の前はワークスペースが決まらないため、ツールのフォルダに置く（[起動に失敗したときの知らせ](disclosure.md#起動に失敗したときの知らせtebunkobat)） | `tebunko.bat`、`tebunko/ui/startup_error.ps1` の `writeStartupErrorFile` |
| 利用者が指定した出力先 | ［結果をファイルに出力］の保存先（既定は `work\search_results.txt`） | 画面のダイアログで利用者が指定 |
| 利用者が指定したエクスポート先 | インデックスのエクスポート（`exportIndex`）の保存先の zip と、書き終えてから置き換えるまでの間だけ残る `<保存先>.tmp`（強制終了すると残ることがある） | 画面のダイアログで利用者が指定 |

単一 .ps1 版（試験版）は、`tebunko.bat` の代わりに `.ps1` ファイル自身を実行する形のため、上の表の「ツールを置いたフォルダ」は `.ps1` ファイル自身がある場所を指す。`setting.config` は `.ps1` のある場所に置き、書き込めないときは既定のワークスペースの直下に置く（`getSettingsFilePath`。zip 版と同じ決め方）。ワークスペースの既定の置き場所（`%USERPROFILE%\Documents\tebunko_ws`）は、別のフォルダに置いた zip 版と同じ場所になる（[単一 PowerShell のビルド](../design/structure/single-script.md)「実行時の違い」）。

削除（`Remove-Item`・`[System.IO.File]::Delete` など）の対象はすべてワークスペース配下（取り込みの作業領域 `work/tmp/` を含む）、すなわち**本ツールが自分で作ったファイル**である（前の版（`%TEMP%\tebunko\<PID>` に一時ファイルを置いていた版）が残した作業フォルダの片付けだけ例外。`removeStaleTmpDirs`）。クロール対象フォルダ内のファイルを削除する処理は無い。

［設定］でワークスペースを変えるときは、前のワークスペースの中の本ツールが作ったファイル・フォルダ（`Workspace.Entries`：`content_index`・前の版の `index`・`system_index`・取り込み一覧・状態・ログ・`publish`）だけを新しいワークスペースへ移す（`moveWorkspace`）。別のドライブへは写してから元を削除する。選んだフォルダにすでにインデックスなどがあり、利用者が［消して、最初からやり直す］を選んだときは、そのフォルダの本ツールのファイル・フォルダ（同じ `Workspace.Entries`）だけを削除する（`removeWorkspaceEntries`）。利用者がワークスペースに置いたほかのファイルは移さず、削除しない。

前の版のワークスペース（`index\`。しるしあり）が見つかり、`content_index\` が空のまま取り込み直しを始めるときは、インデクサが `system_index\`（本ツールが自分で作ったフォルダ）を削除し、`system_index_state.tsv`（本ツールが自分で作ったファイル）の中身を空にする（`clearLegacySystemIndex`）。前の版の `index\` そのものは読まず、削除しない（利用者が消す）。

インデックスのインポート（`importIndex`）の書き込みは、ワークスペースの中（`publish\<PID>\import\`・`content_index\<名前>\`・取り込み一覧・`setting.config`）だけである。zip から展開したファイルは、確かめた目録のパスからだけファイル名を組み立てる（zip のエントリー名は使わない）。インポートを始めるときは、エクスポートと同じくインデックス作成のロック（`newAppMutex "indexer"`）を取り、ほかの操作と同時に走らせない。

`content_index\` の下のフォルダ・ファイルには、Windows Search の索引の対象から外すため「内容のインデックスを作成しない」属性（`NotContentIndexed`）を付ける（`setNotContentIndexed`。[入れ替えと書き出し](../design/index-data/publish.md#windows-search-の対象から外すnotcontentindexed)）。インデックス作成のとき（`content_index` を作った直後・入れた直後）と、インデックスをインポートした直後（入れたフォルダ全体）と、画面がインデックス一覧を保存したとき（`source_folder.txt` 1 件）に付ける。ほかの属性（読み取り専用など）は変えず、内容も書き換えない。

確認コマンド（[検査項目と結果](checks.md#検査項目と結果) の `scan` を使う）:

```powershell
# 削除・書き込みの呼び出し箇所をすべて列挙し、対象のパスが work / tmpDir 由来であることを目で確かめる
scan 'Remove-Item','WriteAllText','WriteAllLines','StreamWriter','\.SaveAs','::Move','Move-Item'
# NotContentIndexed 属性を書く箇所（fs.ps1 の setNotContentIndexed だけ）を確かめる
scan 'NotContentIndexed'
```

### 取り込みの作業フォルダに置くもの

取り込みの作業フォルダ（既定 `work\tmp\<PC の鍵>\<PID>\w<番号>\`）には、原本のコピー・Excel のシートごとの一時保存・旧形式から変換した一時ファイル・公開前の中間 TSV を置く。置き場所はワークスペースの下である。

| 内容 | 今の置き場所 | 定義 |
|---|---|---|
| 原本のコピー（Excel・Word・PowerPoint が開く対象） | `work\tmp\<PC の鍵>\<PID>\w<番号>\<元のファイル名>` | `extract_office.ps1` の `copyFileShared` 呼び出し |
| Excel のシートごとの一時保存（`sheet<番号>.tmp`） | 同上の下 | `extract_office.ps1` の `extractWorkbook` |
| 旧形式・不明な形式から変換した一時ファイル（`converted.docx`・`converted.pptx`、リネームした `source.doc`・`source.ppt`） | 同上の下 | `extract_office.ps1` の `extractDocument` |
| 公開前の中間 TSV（本文・シート・図形などの TSV。テキストファイルの `doc_body.tsv` を含む） | 同上の下 | `index_migrate.ps1` の `publishTsv`、`extract_text.ps1` の `extractTextFile` |

ワークスペースのパスに `[` `]` を含む・長すぎて置けないときは、作業フォルダを作らず、取り込みをすべてスキップする（`selectTmpDir`・`initTmpDir`。[データの置き場所とパスの決め方](../design/structure/data.md)）。

### ワークスペースの外に、まだ書くもの

上の作業フォルダをワークスペースの下に寄せた後も、次のものはワークスペースの外（または、そもそもディスクに書かない）のままである。

| 種類 | 置く場所 | 中身 | 外に置く理由 |
|---|---|---|---|
| 利用者ごとの設定データ | `setting.config`（ツールのフォルダに書き込めるときはその直下。書き込めないときだけ、既定のワークスペース `%USERPROFILE%\Documents\tebunko_ws` の直下） | 画面が保存する設定 | ツールを置いたフォルダ（`Program Files`・読み取り専用の共有フォルダ）に書けないときの代わりの場所。前の版の `%LOCALAPPDATA%\tebunko` には置かない（ツールのフォルダ自体が `%LOCALAPPDATA%\Programs\tebunko` にあるインストーラー版では、書き込めるのでその直下に置く）。ワークスペースより先に（設定を読む前に）決まる必要があるため、ワークスペースの既定の場所に固定する |
| 起動失敗の記録 | ツールのフォルダの `startup_error.txt`（書けなければ作らない） | 起動に失敗した理由・日時・実行環境の情報 | 画面が開く前（ワークスペースも設定も読めないことがある）に書くため、ツールのフォルダにする |
| Office アプリの PID | メモリ上の `OfficePids`（画面を閉じるとき用）と、ワークスペースの `office_pids\<PC の鍵>\<PID>.txt`（次の起動時の確認用） | 起動した Office の PID → プロセス名・起動時刻・起動した側の PID | 画面を閉じるときは、応答の無い Office だけを PID で止める。強制終了などで残ったものは、次の起動時に記録のあるものだけを確認して止めるため、ファイルに残す |
| 書き込み中の一時ファイル（`<保存先>.tmp`。置き換えたら消える） | 書く先のすぐ隣（`setting.config.tmp`・`ingest_status.tsv.tmp` などの TSV の隣・利用者が指定したエクスポート/インポート先の隣 など） | 書き込み中の内容 | 途中で強制終了してもファイルが壊れないよう、一時ファイルに書いてから置き換える（`writeTextLinesAtomic` など）。置き換え（`File.Replace`）は同じフォルダ内でしか使えない |
| 書き込みテストの確認用 | ツールのフォルダ直下 `.tebunko_write_test_<GUID>.tmp` | 空（`DeleteOnClose` でファイルを閉じた瞬間に消える） | 起動のたびに、ツールのフォルダに書き込めるかを確かめるため（`testWritableFolder`） |
| Office・PowerShell が自分で書く一時ファイル | `%TEMP%` 配下（本ツールは関与しない） | Office のロックファイル（`~$<ファイル名>`）・自動回復用のファイルなど、PowerShell・.NET ランタイムが内部で使う一時ファイル | tebunko が選べる場所ではなく、Office・PowerShell・.NET 自身の既定の動きのため |
| 開発用の出力 | リポジトリ直下 `work/test/`・`work/release/`・`work/site/`・`work/cache/`（配布物には含めない） | テスト結果・カバレッジ・配布 zip・設計書のサイト・サイトのビルドキャッシュ | 開発・CI の道具（`tests/run.ps1`・`tools/`）が使う出力で、配布した tebunko 自体は書かない |

## 取り込み対象のファイルは書き換えない

本ツールは、クロール対象フォルダのファイルを**直接開かない**。作業フォルダへコピーし、そのコピーだけを開く（`tebunko/indexer/extract_office.ps1:103`・`261` の `copyFileShared`）。コピー元は読み取り専用で開く（`shared/core/fs.ps1:121`。`FileAccess::Read` で開き、ほかのアプリの読み書き・削除を妨げない）。

この設計は、インデックス作成中も利用者が原本を上書き保存・移動できるように（ファイルロックを避けるために）入れたものだが、結果として原本への書き込み経路そのものが存在しない。

```mermaid
flowchart LR
    src[("クロール対象フォルダ<br>原本")]
    copy["work/tmp/#lt;PC の鍵#gt;/#lt;PID#gt;<br>コピー"]
    office["Excel / Word / PowerPoint<br>読み取り専用・マクロ無効"]
    tsv["work/tmp/#lt;PC の鍵#gt;/#lt;PID#gt;<br>中間 TSV"]
    idx[("work/content_index/<br>本文インデックス")]

    src -- "コピー（読み取りのみ）" --> copy
    copy --> office
    office -- "SaveAs（保存先はコピー側）" --> tsv
    tsv --> idx
```

Office の `SaveAs` は 3 か所あるが、保存先は常に作業フォルダ内のパス（`$tmpPath` / `$destPath`）である（`extract_office.ps1:156`・`210`・`234`）。原本のパスを `SaveAs` に渡す経路は無い。

テキストファイル（`.txt` 等）は Office を使わないため、コピーも作らない。原本を読み取り専用の共有（`copyFileShared` と同じ `FileShare.ReadWrite | Delete`）で直接開いて読み、閉じるだけである（`shared/core/text_file.ps1` の `readTextFile`）。書き込みの API には渡さない。

`tests/meta/safety.Tests.ps1` の「取り込み対象のファイルを書き換えないこと」が、原本のパスを書き込み・削除の API に渡さないこと、原本を読むのは `copyFileShared` の読み取りだけであること、`SaveAs` の保存先が作業フォルダだけであること、Word・PowerPoint を読み取り専用で開くことを確かめる（Excel を読み取り専用で開くことは「Office ファイルを安全に開くこと」が確かめる）。

## 読み取りのみであることの実証

静的な確認に加えて、実際に動かして確認できる。第三者が再現できる手順は次の 2 つ。

**手順 A: ハッシュの比較**

```powershell
$target = "C:\確認用\対象フォルダ"
$before = Get-ChildItem $target -Recurse -File | Get-FileHash
# ここで tebunko.bat を起動し、$target を対象にインデックスを作成する
$after  = Get-ChildItem $target -Recurse -File | Get-FileHash
Compare-Object $before $after -Property Path,Hash   # 何も出力されなければ、原本は 1 バイトも変わっていない
```

**手順 B: 書き込み権限を外して実行**

対象フォルダに読み取りの権限だけを与えたアカウントで、インデックス作成を最後まで実行する。書き込みを試みていれば失敗するため、完走することが「読み取りのみ」の裏付けになる。
