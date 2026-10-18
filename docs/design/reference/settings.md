# 部品ごとの関数（設定・ファイル）

扱うこと: 設定ファイル（`tebunko/core/settings.ps1`）、検索対象でチェックを外したフォルダ（`search_excludes.ps1`）、ワークスペース（`workspace.ps1`）とその切り替え（`workspace_mover.ps1`）、ファイル・フォルダ操作（`shared/core/fs.ps1`・`folder.ps1`・`office_file_kind_view.ps1`）の関数一覧（入力・出力・概要・使用元）。扱わないこと: 取り込み一覧・インデックスの管理（[部品ごとの関数（インデックス作成）](indexer.md)）。先に読むページ: [部品から関数一覧を引く](index.md)。

## 設定ファイル（`tebunko/core/settings.ps1`・`source_name.ps1`）

| 関数 | 入力 | 出力 | 概要 | 詳細 | 使用元 |
|---|---|---|---|---|---|
| `newSettings` | – | ordered hashtable | 設定の既定値（`targetFolders` `indexSources` `searchExcludes` `useRegex` `caseSensitive` `fileKinds` `includeShapes` `includeComments` `openMode` `workspaceFolder` `searchLimit` `ingestThreads`） | [設定ファイル（setting.config）](../structure/settings-file.md) | readSettings |
| `readSettings` | path（既定 `$settingsFile`） | ordered hashtable | 設定を読む。記載の無いキーは既定値。ファイルが無ければ既定値（ファイルは作らない）。数値のキーは文字列でも数値にして読む（読めなければ既定値）。JSON として読めなければ `FormatException`（ファイルは動かさない）。ロック・共有違反は `IOException` のまま | 同上 | 設定の各関数 |
| `toSettingBool` | value, default | bool | 設定ファイルの真偽値を読む。文字列の `"true"` / `"false"` も読み、読めなければ default | 同上 | readSettings, readSearchExcludes（`search_excludes.ps1`） |
| `writeSettings` | settings, path（既定 `$settingsFile`） | – | 設定を JSON（UTF-8 BOM なし）で保存。一時ファイルに書いてから置き換える（`writeTextLinesAtomic`） | 同上 | updateSettings |
| `invokeSettingsLocked` | path, action, timeout（既定 5000 ミリ秒） | action の出力 | 設定ファイルごとの名前付きミューテックス（`Local\tebunko_settings_<getFolderKey の鍵>`）の中で action を実行する。設定の「読む → 変える → 書く」の一続きを囲む。引数の順は path, action, timeout（`-action` は名前で渡す） | [設定ファイル（setting.config）](../structure/settings-file.md) | updateSettings など |
| `updateSettings` | key, value, path（既定 `$settingsFile`） | – | ファイルを読み直し、key の値だけ変えて保存する（`invokeSettingsLocked` の中） | 同上 | 設定の各関数 |
| `repairBrokenSettings` | path（既定 `$settingsFile`） | 退避したパス（何もしなければ空文字） | 設定ファイルが JSON として読めないときだけ、`setting.config.broken-<yyyyMMdd-HHmmss>`（重なれば `-2`・`-3`…）に中身のまま移してパスを返す。移せなければ例外。`invokeSettingsLocked` の中で行う。起動口が起動の時に 1 回だけ呼ぶ | 同上 | gui.ps1, indexer.ps1 |
| `getSettingsRecoveryMessage` | brokenPath | 文字列 | 退避して既定の設定で起動したときの知らせの文言（画面・インデクサで共通） | 同上 | gui.ps1, invokeIndexer |
| `getTargetFolders` | path（既定 `$settingsFile`） | `@{Name; Path; Enabled}` の配列 | クロール対象フォルダ（`targetFolders`。記載順）。Name はインデックス名、Path は今の置き場所（分けて持つ）。`enabled` が `false` はチェックなし（Enabled = `$false`）。同じフォルダ・同じ名前は最初のものだけ（書き方が違うだけで同じフォルダも `testSameFolder` で同一とみなす。重複した名前は空にして割り当て直す） | [クロール対象フォルダと取り込み対象](../indexing/crawl.md#クロール対象フォルダgettargetfolders) | インデックス作成・画面 |
| `writeTargetFolders` | folders（`@{Name; Path; Enabled}` の配列）, path（既定 `$settingsFile`） | – | クロール対象フォルダを `targetFolders` に保存（名前も保存する） | 同上 | インデックス作成・画面 |
| `mergeAssignedIndexNames` / `saveAssignedIndexNames` | current, assigned / assigned, path | 一覧 / 一覧 | 割り当てたインデックス名（assigned）を、読み直した今の一覧（current）の名前が空の項目にだけ、パスで突き合わせて足す（同じ名前をほかの項目が使っていれば付けない）/ 排他の中で読み直して足して保存し、保存した一覧を返す | [設定ファイル（setting.config）](../structure/settings-file.md) | インデクサ・画面（loadTargets） |
| `readIndexSources` / `writeIndexSources` | path（既定 `$settingsFile`） / sources, path | `@{Name; Path}` の配列 / – | インデックス作成の対象にしないインデックスの元のフォルダ（`indexSources`）を読み書きする | [設定ファイル（setting.config）](../structure/settings-file.md) | getSourceFolderMap, 画面 |
| `setIndexSourceFolder` | name, folder, path（既定 `$settingsFile`） | – | インデックス名に対する元のフォルダを記録する。クロール対象フォルダにある名前ならそのフォルダの Path を書き換え、無ければ `indexSources` に記録する | [元のファイルが見つからないとき（元のフォルダを設定する）](../gui/open-file.md#元のファイルが見つからないとき元のフォルダを設定する) | 画面 |
| `testSourceNameRecordable` | name | bool | 元のフォルダを設定に記録できる名前か（空・前後に空白があれば偽。`setIndexSourceFolder` の入口と、確認ダイアログの文言で使う。`tebunko/core/source_name.ps1`。判断層） | [元のファイルを開く](../gui/open-file.md) | 判断 |
| `readVersionFile` | path（配布物の `VERSION.txt` のパス） | `@{Tag; Sha}` / `$null` | 版とコミットの記録を読む。無い・読めない・2行でない・形が違えば `$null`（画面は「開発版」と表示する。`about_view.ps1` の `getAboutView`） | [画面](../gui/index.md) | 画面 |
| `resolveGitRef` | ref, dirs | string | リポジトリのリファレンス（`refs/heads/main` など。形が違えば空）が指すコミットの SHA を、`refs\` の下のファイル → `packed-refs` の順に読む（外部のコマンドは起動しない） | [画面](../gui/screens/window.md) | readDevVersion |
| `readDevVersion` | rootDir | `@{Tag; AtHead; Sha}` / `$null` | 配布物でない（`VERSION.txt` が無い）ときの版の手がかりを、作業ツリーの `.git` のファイルだけから読む（worktree の `gitdir:`・`commondir` も読む）。HEAD に付いたリリースのタグ（`v<数>.<数>.<数>`）があればそれ、無ければ版の並びで最も新しいタグ。どちらも分からなければ SHA だけ、SHA も無ければ `$null` | [画面](../gui/screens/window.md) | getToolDiagnosticHeader, 画面の起動（裏の仕事） |
| `readSearchOption` / `writeSearchOption` | path / option, path | `@{UseRegex; CaseSensitive; IncludeShapes; IncludeComments}` / – | 画面の検索条件（種類は `readFileKinds` / `writeFileKinds`）。[元のファイルの特定・画面](search.md#元のファイルの特定画面) を参照 | [続き](search.md#元のファイルの特定画面) | 画面 |
| `getSearchLimit` | value | int | 検索結果の件数の上限として使う値を返す。0 は制限なし（`${searchLimitNone}`）でそのまま返す。数値にできない・負の数は既定の 10,000（`${searchLimitDefault}`）、範囲の外は近い端（1,000＝`${searchLimitMin}`～100,000＝`${searchLimitMax}`）に寄せる。ファイルは読まない | [設定ファイル（setting.config）](../structure/settings-file.md) | readSearchLimit |
| `readSearchLimit` | path（既定 `$settingsFile`） | int | 検索結果の件数の上限（`searchLimit`）を読み、`getSearchLimit` で寄せて返す。画面は起動のときに読む | [検索タブ](../gui/search-tab.md) | 画面（`startGui`） |
| `writeSearchLimit` | value、path（既定 `$settingsFile`） | – | 検索結果の件数の上限（`searchLimit`）を書く。0（制限なし）か 1,000〜100,000 だけ書き、ほかの値は例外にする。ほかのキーは変えない | [設定ファイル（setting.config）](../structure/settings-file.md) | 設定の画面（`commitSearchLimit`） |
| `readCloudFiles` / `writeCloudFiles` | path（既定 `$settingsFile`） / mode, path | string / – | クラウドにだけあるファイルの扱い（`cloudFiles`。`ask` / `download`）を読み書きする。無い・知らない値なら `ask` | [取り込み対象の決定](../indexing/target-decision.md#クラウドにだけあるファイル) | 画面・インデクサ |
| `readOpenMode` / `writeOpenMode` | path（既定 `$settingsFile`） / mode, path | string / – | 元のファイルの開き方（`openMode`）を読み書きする。無い・知らない値なら `normal` | [元のファイルを開く](../gui/open-file.md) | 画面 |
| `getDefaultWorkDir` | profileDir（既定は利用者のプロファイル） | string | 既定のワークスペース `<profileDir>\Documents\tebunko_ws`（OneDrive にリダイレクトされた「ドキュメント」は使わない）。`profileDir` を渡さず、環境変数 `TEBUNKO_DEFAULT_WORKSPACE`（絶対パス）があるときは、その場所（テスト・実機の確かめ用。絶対パスでなければ例外） | [データの置き場所とパスの決め方](../structure/data.md) | getWorkDir, writeWorkspaceFolder, 画面 |
| `getWorkDir` | path（既定 `$settingsFile`） | string | `work` の置き場所（`workspaceFolder`。空なら既定。相対パスは設定ファイルのフォルダから、`%変数%` は展開する） | 同上 | `paths.ps1`（`$workspace`） |
| `writeWorkspaceFolder` | folder, path（既定 `$settingsFile`） | – | `work` の置き場所を保存する。既定の場所なら空で保存する | 同上 | 画面 |

## ワークスペース（`tebunko/core/workspace.ps1`）

`Workspace` クラスは、ワークスペースのフォルダ（Dir）から、中のファイル・フォルダの場所を組み立てるだけで、設定は読まない。どのフォルダを使うかは設定の `workspaceFolder` で決まる（`getWorkDir`）。関数はその場所を使って、取り込みの作業フォルダの置き場所の判断・前の版のワークスペースの扱いを行う（ワークスペースの移動は下の `WorkspaceMover`）。画面の操作は [データの置き場所とパスの決め方](../structure/data.md)、［設定］タブは [［設定］タブ](../gui/settings-tab.md) を参照。

| 関数 | 入力 | 出力 | 概要 | 使用元 |
|---|---|---|---|---|
| `getMachineKey` | – | string（8 文字） | この PC を識別する短い鍵（`getFolderKey` の先頭 8 文字）。共有フォルダのワークスペースを複数の PC から使うとき、一時フォルダを PC ごとに分ける | getWorkspaceTmpDir |
| `assertWorkspaceReachable` | dir | – | 裏の仕事の先頭でワークスペースに届くかを確かめる。届かない・確かめられないときは「ワークスペースに接続できません：…」「ワークスペースを確かめられません：…」の例外にする。まだ無いだけなら何もしない（呼び出しは裏の仕事の中。画面のスレッドでは呼ばない） | `IndexCatalog.AssertReachable`（index_catalog.ps1） |
| `getWorkspaceJobQueue` | 場所（複数可） | string | 裏の列を選ぶ。渡した場所のどれかがネットワークなら `network`、それ以外は `default`。届かない共有の仕事が、ほかの仕事を待たせないため | index_store.ps1, index_tree.ps1 ほか |
| `getWorkspaceTmpDir` | workspace | string | 取り込みの作業フォルダの候補 `<TmpRoot>\<PC の鍵>\<PID>`。副作用は無い | selectTmpDir, index_migrate.ps1 |
| `selectTmpDir` | workspace | `@{Dir; Reason}` | 作業フォルダの置き場所を決める。候補のパスに `[` `]` があれば（Excel が保存できない）`Dir` を空にして `Reason = Brackets`、候補の長さに取り込みのスレッドが下に作る名前の分を足して `$excelMaxPath` 以上なら `Reason = TooLong`。どちらでもなければ候補のまま（`Reason` は空）。`%TEMP%` には逃がさない | paths.ps1, indexer_run.ps1, extract_office.ps1, index_migrate.ps1 |
| `getTmpDirUnavailableMessage` | reason | string | 作業フォルダを置けない理由（`Brackets`・`TooLong`・`Link`（作業領域の途中にリンクがある）・その他）を、利用者向けの 1 文にする | indexer_run.ps1, index_migrate.ps1 |
| `getLegacyIndexState` | dir | `@{HasLegacyIndex; ContentEmpty}` | 前の版のワークスペースの状態。`index\` の直下のフォルダに元のフォルダ.txt があれば前の版のしるし（`HasLegacyIndex`）。`store\content\` と前の版の `content_index\` がどちらも空か（`ContentEmpty`） | indexing_tab.ps1, workspace_jobs.ps1, index_archive.ps1, indexer_run.ps1, paths.ps1 |
| `getLegacyIndexMessage` | dir, hasLegacyIndex | string | 前の版のインデックスが見つかったときの知らせ。しるしが無ければ空 | workspace_jobs.ps1, indexer_run.ps1 |
| `removeOldSystemIndex` | dir | `@{Ok; Reason}` | 前の版が作った検索用のファイル（`system_index\` と `system_index_state.tsv`）を消す。ワークスペースの下で途中にリンクが無いことを確かめてから消し（`removeDirectoryUnder`・`removeFileUnder`）、無ければ何もしない。消せなくても例外にせず `Ok = $false`（次に呼べば続きから） | startup.ps1（画面の起動）, indexer_run.ps1, index_archive.ps1 |
| `newWorkspace` | dir | `Workspace` | `Workspace` のインスタンスを作る（クラスの本体では型の名前を書けないため、`WorkspaceMover`・`IndexCatalog` がこの関数を通して作る） | workspace_mover.ps1 |

## ワークスペースの切り替え（`tebunko/core/workspace_mover.ps1`。`WorkspaceMover`）

作るのは `[WorkspaceMover]::new(設定ファイルの場所)`。`ListEntries` などの確かめは画面のスレッドで使う。画面（`applyWorkspace`）は、選んだ直後のフォルダを確かめて利用者に選ばせたあと、`ChangeWorkspace` を画面の裏の仕事（`startWorkspaceMove`）で呼ぶ。終わったら、画面のスレッドで結果を出して画面を切り替える（結果の判断は `getWorkspaceMoveOutcome`）。`ChangeWorkspace` は成功したときだけ設定を書き、失敗したら戻すので、裏で失敗しても設定は変わらない。クラスの考え方は [クラスと関数の使い分け](../structure/classes.md)。

| メソッド | 入力 | 出力 | 概要 | 使用元 |
|---|---|---|---|---|
| `ListEntries` | dir | string[] | tebunko のファイル・フォルダ（`Workspace.Entries` のうち、あるもの）のフルパス。利用者のほかのファイルは含めない（旧 `getWorkspaceEntries`） | settings.ps1 |
| `Conflicts` | from, to | string[] | from の中身を to へ移すとき、to に同じ名前が既にあるものの名前（旧 `getWorkspaceMoveConflicts`） | `MoveEntries` |
| `ChangeWorkspace` | mode（`use`・`move`・`reset`）, from, to | `@{Ok; Stage; Message; Targets; Count}` | ワークスペースを from から to に変えて設定に保存する。`use` は to のインデックスを使う（`UseTargets`。from の中身は移さない）、`move` は中身を移す、`reset` は to の中身を消してから移す（`RemoveEntries`・`MoveEntries`）。移したあとに `searchExcludes` を付け替える（`moveSearchExcludes`）。失敗したときは設定を変えず、`Stage`（`remove`・`move`）と `Message` を返す | settings.ps1 |
| `MoveEntries` | from, to | int | from の中身を to へ移し、移した数を返す。to に同じ名前があれば何も移さずに例外にする。途中で移せなければ、移した分を from へ戻してから例外にする（旧 `moveWorkspace`。1 つずつ移す `MoveEntry` と、別のドライブへ写す `CopyTree` は内部） | `ChangeWorkspace` |
| `RemoveEntries` | dir | int | tebunko のファイル・フォルダを削除し、削除した数を返す。「消して、最初からやり直す」で使う（旧 `removeWorkspaceEntries`） | `ChangeWorkspace` |
| `UseTargets` | dir | int | dir の取り込み一覧（新しい `store\ingest_status.tsv` と前の版の直下の `ingest_status.tsv` の両方。新しい方が先）にあるクロール対象フォルダを、インデックスの一覧（`targetFolders`）にし、その数を返す。取り込み一覧が無ければ一覧は変えない（旧 `useWorkspaceTargets`） | `ChangeWorkspace` |

## ［新規で開く］のコピー（`tebunko/core/open_copy.ps1`）

| 関数 | 入力 | 出力 | 概要 | 使用元 |
|---|---|---|---|---|
| `getOpenCopyDir` | workspace | string | コピーを置くフォルダ（`work\tmp\<PC の鍵>\open_copy`）。強制終了の片付けは数字の名前のフォルダだけを消すので、ここは消えない | `openFoundSource`, 起動時の片付け |
| `testOpenCopyDir` | copyDir, tmpRoot | bool | 作る・消してよい置き場所か。絶対パスで、`tmpRoot\<PC の鍵>\open_copy`（大文字・小文字は区別しない）と一致し、`open_copy`・`<PC の鍵>`・`tmpRoot` がリンク（ジャンクション・シンボリックリンク。`testPathIsLink`）でないときだけ真。空・相対パス・ドライブの根・別の場所・リンクは偽 | `copySourceForOpen`, `clearOldOpenCopies` |
| `getOpenCopyName` | book, now | string | コピーのファイル名 `<元の名前>（コピー yyyyMMdd-HHmm）.<拡張子>`（元のファイル名の末尾の `.`・空白を除く）。名前が空なら空 | `copySourceForOpen`, `getOpenCopyBlocker` |
| `getOpenCopyBlocker` | book, copyDir, now | string | コピーを作らずに断る理由。`NoName`（名前が空）・`TooLong`（パスが長すぎる）・作れるなら空（名前に `[` `]` があっても断らない） | `copySourceForOpen`, `openFoundSource` |
| `copySourceForOpen` | sourcePath, book, copyDir, tmpRoot, now | `@{Path; Reason}` | `copyDir\<12 桁の英数字>\<getOpenCopyName>` に、元のファイルを共有で読んでコピーし、`Path` にそのパスを入れる（`Reason` は空）。作れないときは `Path` を空にし、`Reason` に理由を入れる（`Place`＝置き場所が不正・`NoName`／`TooLong`＝`getOpenCopyBlocker` の結果・`Failed`＝作る・書くのに失敗） | `openFoundSource`（裏の仕事） |
| `clearOldOpenCopies` | copyDir, tmpRoot, now, keepDays | int | `open_copy` の直下のフォルダのうち、最後に使ってから（`getOpenCopyLatestTime`）keepDays（既定 1 日）より古いものだけを、ファイルを 1 つずつ消してから空のフォルダとして消す。リンクは辿らず、開かれていて消せないものは残す。消したフォルダの数を返す | 起動時の裏の仕事 |

## 検索対象でチェックを外したフォルダ（`tebunko/core/search_excludes.ps1`）

設定の `searchExcludes` の読み書きと、フォルダとその下の比べ方。`IndexCatalog`（改名・削除）、`WorkspaceMover`（切り替え）、検索対象のツリーが同じ部品を通す。

| 関数 | 入力 | 出力 | 概要 | 使用元 |
|---|---|---|---|---|
| `testFolderOrUnder` | path, folder | bool | path が folder そのもの、または folder の下かを返す。`searchExcludes` の削除・付け替え・保存でのフォルダとその下の比べ方はここだけ。大文字・小文字を区別しない序数の比べ（`OrdinalIgnoreCase`）で、folder の末尾の `\` は外し、`<folder>\` で始まるものだけを「下」にする（`Sales` の下に `Sales2` は入らない）。folder が空ならどれも含まない。既知の制限: `\\?\` 付きと無し、`/` 区切り、`..`、8.3 形式の短い名前は同じ場所として扱わない（書き方は呼び出す側がそろえる） | `removeSearchExcludesUnder`, `moveSearchExcludes`, 検索対象のツリー（`saveSearchExcludes`） |
| `readSearchExcludes` / `writeSearchExcludes` | path（既定 `$settingsFile`） / excludes, path | `@{Path; Subfolders}` の配列 / – | 画面の検索対象のツリーでチェックを外したフォルダ（`searchExcludes`）を読み書きする。無ければ空（すべて検索） | 画面（検索対象のツリー） |
| `removeSearchExcludesUnder` | folder, path（既定 `$settingsFile`） | – | フォルダとその下の `searchExcludes` を消す（`Sales` を消しても `Sales2` は消さない）。設定の書き込みは `invokeSettingsLocked` の中で行い、消せなくても例外にしない（インデックスの改名・削除を止めない） | `IndexCatalog.RenameIndex` / `RemoveIndex` |
| `moveSearchExcludes` | from, to, path（既定 `$settingsFile`） | int | 検索対象ツリーでチェックを外したフォルダ（`searchExcludes`）のうち、from の `store\content` の下のものを to の下に付け替えて保存し、付け替えた数を返す | `WorkspaceMover.ChangeWorkspace` |

## ファイルとフォルダ（`shared/core/fs.ps1`・`folder.ps1`・`shared/office/office_file_kind_view.ps1`）

| 関数 | 入力 | 出力 | 概要 | 詳細 | 使用元 |
|---|---|---|---|---|---|
| `getOfficeFileKind` | name | string（`Excel` / `Word` / `PowerPoint`。表に無ければ空） | ファイル名・パスの拡張子から Office の種類を返す。拡張子と種類の表（`$officeFileKinds`。`$officeExtensions` はこの表の拡張子）だけを見る（表に無い `.xlsxyz`・`.xlsa` などは空。大文字・小文字は区別しない）。どのツールからも使うため `shared/` に置く | – | getAppName, getContentIndexFileKind, getIngestApp, getRerouteApp, getOfficeOpenKind |
| `getOfficeExtensionsOfKind` | kind | string[] | 種類（`Excel` / `Word` / `PowerPoint`）の拡張子を、表（`$officeFileKinds`）の並びで返す。表に無い種類は空 | – | getFileKindExtensions |
| `readListFile` | path | string[] | 行ファイルの空行以外の行（Trim しない）。無ければ空配列。読めなければ例外（空の一覧と取り違えない） | – | 取り込み中のファイル・元のフォルダの記録 など |
| `writeListFile` | path, lines | – | 行ファイルを UTF-8（BOM 付き）で保存 | – | 同上 |
| `writeTextLinesAtomic` | path, lines, encoding（既定は BOM 付き UTF-8） | – | 一時ファイルに書いてから置き換える（置き換えられなければ少し待って 5 回まで試す） | – | writeStatusFile, renameStatusIndexName など |
| `formatFileTime` | time | string | 取り込み一覧に記録する日時（`yyyy/MM/dd HH:mm:ss`）。更新の有無はこの文字列で比べる | [取り込み一覧](../indexing/ingest-list.md) | インデックス作成 |
| `copyFileShared` | sourcePath, destPath | – | 元のファイルを読み取りだけで開いてコピーする（ほかのアプリの読み書き・削除を妨げない）。インデックス作成は、このコピーを開く | [取り込み対象のファイルは書き換えない](../../safety/file-access.md#取り込み対象のファイルは書き換えない) | インデックス作成 |
| `invokeWithSourceCopy` | sourcePath, copyPath, readBlock, extraFiles | readBlock の出力 | 元のファイルを `copyFileShared` で作業領域へコピーし、コピーを読む処理（readBlock）を実行して、終わったらコピー（と extraFiles）を必ず消す。読む処理が成功したのにコピーを消せなかったときは、`作業領域のコピーを消せませんでした: <元のメッセージ>` を投げて取り込みの失敗にする（テキスト・PDF・Excel・Word・PowerPoint で共通） | [取り込み対象のファイルは書き換えない](../../safety/file-access.md#取り込み対象のファイルは書き換えない) | インデックス作成 |
| `getFolderKey` | dir | string（16 進 64 文字） | フォルダのパスを小文字にした SHA-256。名前付きミューテックス・イベントの名前に使う。FIPS モードの Windows でも動くよう、FIPS 準拠の実装（`SHA256CryptoServiceProvider`）を使う | – | newAppMutex, 画面（多重起動の防止） |
| `testWritableFolder` | dir | bool | フォルダにファイルを作れるか（試しに作ったファイルは閉じると消える）。無いフォルダは `$false` | [データの置き場所とパスの決め方](../structure/data.md) | getDataDir, 画面（置き場所の変更） |
| `getDataDir` | root（既定 `$rootDir`）, fallbackDir（必須） | string | 設定ファイルを置くフォルダ。root に書き込めれば root、書き込めなければ fallbackDir（`%LOCALAPPDATA%\tebunko` には置かない。`shared/core/data_dir.ps1`） | 同上 | `settings.ps1`（`getSettingsFilePath`） |
| `getSettingsFilePath` | root, defaultWorkDir | string | `setting.config` のパス。root に書き込めれば root 直下、書き込めなければ既定のワークスペース直下 | 同上 | `settings.ps1`（`$settingsFile`） |
| `testSettingsFileName` | name | bool | 設定ファイル（`setting.config`・`.tmp`・`.broken-<日時>[-<番号>]`）の名前か。既定のワークスペースが空かの判断と、中身として数える対象から外すのに使う | [設定ファイル](../structure/settings-file.md) | `testDefaultWorkspace`, 画面（`getCountedEntryPaths`） |
| `getSettingsFileView` | settingsFile, rootDir | `@{Path; Note}` | 設定タブに出す設定ファイルの場所と、ツールのフォルダに置いているか（書き込めなくて既定のワークスペースに置いているか）の注記。判断層（`settings_view.ps1`） | [設定タブ](../gui/settings-tab.md) | 画面（設定タブ） |
| `getCountedEntryPaths` | entryPaths, isDefaultWorkspace | string[] | ワークスペースに選んだフォルダの中身のうち「空でない」の数に入れるもの。選んだフォルダが既定のワークスペースなら、設定ファイルとその付属ファイル（`testSettingsFileName`）を外す。判断層（`settings_view.ps1`。カバレッジの分母に入る） | [設定タブ](../gui/settings-tab.md) | 画面（設定タブ） |
| `newAppMutex` | name（`gui` / `indexer`）, dir（既定 `$rootDir`）, app（ツールの名前。呼ぶ側が渡す。空なら例外） | `@{Mutex; Acquired}` | 同じツール（配置フォルダ）の処理を二重に動かさないための名前付きミューテックス（`Local\<app>_<name>_<getFolderKey の鍵>`。tebunko では `Local\tebunko_<name>_<鍵>`）。`Acquired` が `$false` なら、ほかで実行中。プロセスが終われば解放されるため、強制終了しても残らない | [インデックス作成のメインフロー](../indexing/flow.md) | インデックス作成（起動時。dir に `$workspace.Dir` を渡し、同じ `work` を使うものを 1 つにする）。画面は `getFolderKey` の鍵で同じ規則の名前を自前で作る（ウィンドウを前面に出すイベントの名前と共通の鍵を使うため） |
| `invokeWithNamedMutex` | mutexName, timeoutMilliseconds, action | action の出力 | 名前付きミューテックスを取って action を実行し、終わったら手放す。同じスレッドの入れ子は通す。時間を過ぎても取れなければ「排他の待ちが時間切れになりました」の例外。前の持ち主が手放さずに終わっていた（abandoned）ときは取れたものとして続ける | [プロセスとスレッド](../structure/threads.md) | invokeSettingsLocked |
| `normalizeFolderPath` | path | string | フォルダパスを 1 つの書き方にそろえる（前後の空白・`"` の除去、環境変数の展開、`/` → `\`、`\\?\` の除去、重なった `\` ・ `.` ・ `..` の解決、相対パスは `$rootDir` から、末尾の `\` の除去。ドライブ直下は `D:\` のまま）。解釈できない場合は書かれたとおり | [クロール対象フォルダと取り込み対象](../indexing/crawl.md#クロール対象フォルダgettargetfolders) | getTargetFolders, 画面 |
| `getPathUnderFolder` | path, folder | string / `$null` | path が folder 自身か下なら folder からの相対パス（folder 自身は空）、下でなければ `$null`（大文字・小文字と末尾の `\` を無視）。パスの文字数で切り出さないために使う | 同上 | インデックス作成・検索 |
| `getDriveTargets` | – | Dictionary（`Z:` → 割り当て先） | ネットワークドライブの割り当てを CIM（`Win32_LogicalDisk` の DriveType=4）で調べる（同じプロセスで 1 回だけ。`subst` は解決しない） | 同上 | getFolderPathAliases |
| `getFolderPathAliases` | path, drives（既定 `getDriveTargets`） | string[] | 同じ場所を指す別の書き方を path 自身を先頭にして返す（`Z:\見積` ⇔ `\\server\share\見積`） | 同上 | testSameFolder, 画面 |
| `testSameFolder` | a, b, drives（既定 `getDriveTargets`） | bool | 2 つのパスが同じフォルダを指すか（書き方の違い・ドライブの割り当てをたどる） | 同上 | getTargetFolders, getSourceFolderMap, 画面 |
| `testFolderUnder` | path, folder, drives（既定 `getDriveTargets`） | bool | path が folder 自身か folder の下のフォルダか（`testSameFolder` と同じく書き方の違いをたどる）。クロール対象フォルダが入れ子にならないかの確認に使う | 同上 | 画面（インデックスの追加・編集） |
| `getFolderLeafName` | folderPath | string | フォルダ名（ドライブ直下はドライブ名、UNC は共有名） | [取り込み一覧](../indexing/ingest-list.md) | newIndexName |
| `joinSourcePath` | folder, rest, name | string | フォルダ・相対フォルダ・ファイル名をつなぐ（ドライブ直下でも `\` を重ねない） | – | getFolderPathAliases, 元のファイルの特定（`source_map.ps1`）, 画面 |
| `getExistingAncestorFolder` | folder | string | folder が今もあればそのまま、無ければその上の今もあるフォルダ。どこにも無ければ空 | [フォルダ選択ダイアログ（［参照…］）](../gui/common.md#フォルダ選択ダイアログ参照) | 画面（フォルダ選択） |

## 診断情報とログの世代（`shared/core/diagnostics_view.ps1`・`diagnostics.ps1`・`tebunko/core/diagnostics.ps1`）

ログの先頭に書く版と環境の見出しと、ログの入れ替え。仕様は[ログに書く版と環境の情報](../../safety/disclosure.md#ログに書く版と環境の情報)と[増え方と片付け](../structure/io-files.md#増え方と片付け)。

| 関数 | 入力 | 出力 | 概要 | 使用元 |
|---|---|---|---|---|
| `getVersionLabel` | versionInfo（`readVersionFile` の結果） | `@{Version; Commit}` | 版の表示。タグとコミットの先頭 7 桁。`$null` なら `開発版`（Commit は空）。`readDevVersion` の結果（devInfo）があれば `開発版（v0.5.0・d79f898）` / `開発版（v0.5.0 より後・d79f898）` / `開発版（d79f898）`。判断層 | `getAboutView`, `formatDiagnosticHeader` |
| `getOfficeKind` | exePath, programFiles64, programFilesX86 | `@{Kind; Bits}` | Office の実行ファイルのパスから、Click-to-Run か MSI か、32・64 ビットかを決める。判断層 | `findOfficeApps` |
| `hideUserProfile` | text, profileDir | string | 文字列の中のプロファイルのフォルダを `%USERPROFILE%` に置き換える（大文字小文字は区別しない。前方だけ一致する別のフォルダは置き換えない）。判断層。今は呼び出し元が無い（見出しにプロファイルのパスを入れないため） | – |
| `formatDiagnosticHeader` | toolName, versionInfo, envInfo, now | string[]（見出しの行） | 日時・版・OS・PowerShell・Office・表示言語の見出しを組み立てる。調べられない項目は `不明`。判断層 | `getToolDiagnosticHeader` |
| `getEnvironmentInfo` | – | hashtable | OS・PowerShell・Office・表示言語を調べる。項目ごとに調べ、失敗した項目は空にする（Office を調べている途中で失敗したら `$null`。見出しでは `不明`）。レジストリ・WMI・COM は使わず、Office は起動しない | `getToolDiagnosticHeader` |
| `findOfficeApps` | programFiles64, programFilesX86（省略すると Windows のあるドライブの決まった名前）, os64 | `@{Name; Version; Kind; Bits}` の配列 | `Program Files` の決まった場所の EXCEL・WINWORD・POWERPNT を探し、ファイルの版を読む。32 ビットの Windows では `Program Files` の Office を 32 ビットとする。例外は投げたまま（呼び出し側が `不明` にする） | `getEnvironmentInfo` |
| `getPreviousLogPath` | path | string | `<名前>_previous<拡張子>`（同じフォルダ） | `rotateLogFile` |
| `rotateLogFile` | path, maxBytes | bool | ファイルが maxBytes 以上なら `_previous` へ移す（あれば `File.Replace` で置き換える。今のログが移せないときも、今のログは失わない）。maxBytes が 0 なら、あれば必ず移す。移したら `$true`、ファイルが無い・小さい・移せないときは `$false`（例外にしない） | `invokeIndexer`, `writeErrorLog` |
| `getToolDiagnosticHeader` | – | string[] | このツールの版（`VERSION.txt`。無ければ `readDevVersion` のタグ・コミットを添えた開発版）と環境の見出し | `invokeIndexer`, `getToolGuiErrorLogHeader` |
