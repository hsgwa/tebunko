# 部品ごとの関数（設定・ファイル）

扱うこと: 設定ファイル（`tebunko/core/settings.ps1`）、ワークスペース（`tebunko/core/workspace.ps1`）とファイル・フォルダ操作（`shared/core/fs.ps1`・`folder.ps1`・`office_files.ps1`）の関数一覧（入力・出力・概要・使用元）。扱わないこと: 取り込み一覧・インデックスの管理（[部品ごとの関数（インデックス作成）](indexer.md)）。先に読むページ: [部品から関数一覧を引く](index.md)。

## 設定ファイル（`tebunko/core/settings.ps1`）

| 関数 | 入力 | 出力 | 概要 | 詳細 | 使用元 |
|---|---|---|---|---|---|
| `newSettings` | – | ordered hashtable | 設定の既定値（`targetFolders` `indexSources` `searchExcludes` `useRegex` `caseSensitive` `fileKinds` `includeShapes` `includeComments` `openMode` `workspaceFolder` `ingestThreads`） | [設定ファイル（setting.config）](../structure/settings-file.md) | readSettings |
| `readSettings` | path（既定 `$settingsFile`） | ordered hashtable | 設定を読む。記載の無いキーは既定値。ファイルが無ければ既定値（ファイルは作らない）。数値のキーは文字列でも数値にして読む（読めなければ既定値）。JSON として読めなければ `FormatException`（ファイルは動かさない）。ロック・共有違反は `IOException` のまま | 同上 | 設定の各関数 |
| `toSettingBool` | value, default | bool | 設定ファイルの真偽値を読む。文字列の `"true"` / `"false"` も読み、読めなければ default | 同上 | readSettings, readSearchExcludes |
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
| `readSearchExcludes` / `writeSearchExcludes` | path（既定 `$settingsFile`） / excludes, path | `@{Path; Subfolders}` の配列 / – | 画面の検索対象のツリーでチェックを外したフォルダ（`searchExcludes`）を読み書きする。無ければ空（すべて検索） | [インデックスの一覧](../search/index.md#インデックスの一覧getsearchindexes) | 画面（検索対象のツリー） |
| `removeSearchExcludesUnder` | folder, path（既定 `$settingsFile`） | – | フォルダとその下の `searchExcludes` を、大文字・小文字を区別せずに消す（`Sales` を消しても `Sales2` は消さない）。設定の書き込みは `invokeSettingsLocked` の中で行い、消せなくても例外にしない（インデックスの改名・削除を止めない） | 同上 | `renameIndex` / `removeIndex` |
| `readVersionFile` | path（配布物の `VERSION.txt` のパス） | `@{Tag; Sha}` / `$null` | 版とコミットの記録を読む。無い・読めない・2行でない・形が違えば `$null`（画面は「開発版」と表示する。`about_view.ps1` の `getAboutView`） | [画面](../gui/index.md) | 画面 |
| `readSearchOption` / `writeSearchOption` | path / option, path | `@{UseRegex; CaseSensitive; IncludeShapes; IncludeComments}` / – | 画面の検索条件（種類は `readFileKinds` / `writeFileKinds`）。[元のファイルの特定・画面](search.md#元のファイルの特定画面) を参照 | [続き](search.md#元のファイルの特定画面) | 画面 |
| `readOpenMode` / `writeOpenMode` | path（既定 `$settingsFile`） / mode, path | string / – | 元のファイルの開き方（`openMode`）を読み書きする。無い・知らない値なら `normal` | [元のファイルを開く](../gui/open-file.md) | 画面 |
| `getDefaultWorkDir` | profileDir（既定は利用者のプロファイル） | string | 既定のワークスペース `<profileDir>\Documents\tebunko_ws`（OneDrive にリダイレクトされた「ドキュメント」は使わない）。`profileDir` を渡さず、環境変数 `TEBUNKO_DEFAULT_WORKSPACE`（絶対パス）があるときは、その場所（テスト・実機の確かめ用。絶対パスでなければ例外） | [データの置き場所とパスの決め方](../structure/data.md) | getWorkDir, writeWorkspaceFolder, 画面 |
| `getWorkDir` | path（既定 `$settingsFile`） | string | `work` の置き場所（`workspaceFolder`。空なら既定。相対パスは設定ファイルのフォルダから、`%変数%` は展開する） | 同上 | `paths.ps1`（`$workspace`） |
| `writeWorkspaceFolder` | folder, path（既定 `$settingsFile`） | – | `work` の置き場所を保存する。既定の場所なら空で保存する | 同上 | 画面 |

## ワークスペース（`tebunko/core/workspace.ps1`）

`Workspace` クラスは、ワークスペースのフォルダ（Dir）から、中のファイル・フォルダの場所を組み立てるだけで、設定は読まない。どのフォルダを使うかは設定の `workspaceFolder` で決まる（`getWorkDir`）。関数はその場所を使って、取り込みの作業フォルダの置き場所の判断・前の版のワークスペースの扱い・ワークスペースの移動を行う。画面の操作は [データの置き場所とパスの決め方](../structure/data.md)、［設定］タブは [［設定］タブ](../gui/settings-tab.md) を参照。

| 関数 | 入力 | 出力 | 概要 | 使用元 |
|---|---|---|---|---|
| `getMachineKey` | – | string（8 文字） | この PC を識別する短い鍵（`getFolderKey` の先頭 8 文字）。共有フォルダのワークスペースを複数の PC から使うとき、一時フォルダを PC ごとに分ける | getWorkspaceTmpDir |
| `assertWorkspaceReachable` | dir | – | 裏の仕事の先頭でワークスペースに届くかを確かめる。届かない・確かめられないときは「ワークスペースに接続できません：…」「ワークスペースを確かめられません：…」の例外にする。まだ無いだけなら何もしない（呼び出しは裏の仕事の中。画面のスレッドでは呼ばない） | index_edit.ps1 |
| `getWorkspaceJobQueue` | 場所（複数可） | string | 裏の列を選ぶ。渡した場所のどれかがネットワークなら `network`、それ以外は `default`。届かない共有の仕事が、ほかの仕事を待たせないため | index_edit.ps1, index_tree.ps1 ほか |
| `getWorkspaceTmpDir` | workspace | string | 取り込みの作業フォルダの候補 `<TmpRoot>\<PC の鍵>\<PID>`。副作用は無い | selectTmpDir, index_migrate.ps1 |
| `selectTmpDir` | workspace | `@{Dir; Reason}` | 作業フォルダの置き場所を決める。候補のパスに `[` `]` があれば（Excel が保存できない）`Dir` を空にして `Reason = Brackets`、候補の長さに取り込みのスレッドが下に作る名前の分を足して `$excelMaxPath` 以上なら `Reason = TooLong`。どちらでもなければ候補のまま（`Reason` は空）。`%TEMP%` には逃がさない | paths.ps1, indexer_run.ps1, extract_office.ps1, index_migrate.ps1 |
| `getTmpDirUnavailableMessage` | reason | string | 作業フォルダを置けない理由（`Brackets`・`TooLong`・その他）を、利用者向けの 1 文にする | indexer_run.ps1, index_migrate.ps1 |
| `getLegacyIndexState` | dir | `@{HasLegacyIndex; ContentEmpty; HasLegacySystemIndex}` | 前の版のワークスペースの状態。`index\` の直下のフォルダに元のフォルダ.txt があれば前の版のしるし（`HasLegacyIndex`）。`content_index\` が空か（`ContentEmpty`）。空のときだけ、`system_index\` に前の名前の txt があるかも調べる | gui_main.ps1, indexing_tab.ps1, settings_tab.ps1, index_archive.ps1, indexer_run.ps1, paths.ps1 |
| `testLegacyCleanupNeeded` | state | bool | 取り込み直しを始めるときに、前の版のシステムインデックスを片付けるか。`content_index\` が空で、前の版のしるしか前の名前の txt があるとき | index_archive.ps1, indexer_run.ps1 |
| `getLegacyIndexMessage` | dir, hasLegacyIndex | string | 前の版のインデックスが見つかったときの知らせ。しるしが無ければ空 | gui_main.ps1, settings_tab.ps1, indexer_run.ps1 |
| `clearLegacySystemIndex` | dir | `@{Ok; Reason}` | 前の版のシステムインデックスを片付ける。状態ファイルの中身を排他の中で空にしてから、`system_index\` を消す。どちらかに失敗したら `Ok = $false`（次に呼べば続きから） | indexer_run.ps1, index_archive.ps1 |
| `getWorkspaceEntries` | dir | string[] | tebunko のファイル・フォルダ（`Workspace.Entries` のうち、あるもの）のフルパス。利用者のほかのファイルは含めない | settings_tab.ps1, settings_view.ps1 |
| `getWorkspaceMoveConflicts` | from, to | string[] | from の中身を to へ移すとき、to に同じ名前が既にあるものの名前 | moveWorkspace |
| `moveWorkspaceEntry` | source, dest | – | ファイル・フォルダを 1 つ移す。同じドライブならそのまま移し、別のドライブのフォルダは写してから元を消す。写している途中で失敗したら、写した分を消して例外にする | moveWorkspace |
| `copyDirectoryTree` | source, dest（`\\?\` 付き） | – | フォルダを中身ごと写す | moveWorkspaceEntry |
| `removeWorkspaceEntries` | dir | int | tebunko のファイル・フォルダを削除し、削除した数を返す。「消して、最初からやり直す」で使う | settings_tab.ps1 |
| `useWorkspaceTargets` | dir, path（既定 `$settingsFile`） | int | dir の取り込み一覧にあるクロール対象フォルダを、インデックスの一覧（`targetFolders`）にし、その数を返す。取り込み一覧が無ければ一覧は変えない | settings_tab.ps1 |
| `moveSearchExcludes` | from, to, path（既定 `$settingsFile`） | int | 検索対象ツリーでチェックを外したフォルダ（`searchExcludes`）のうち、from の `content_index` の下のものを to の下に付け替えて保存し、付け替えた数を返す | settings_tab.ps1 |
| `moveWorkspace` | from, to | int | from の中身を to へ移し、移した数を返す。to に同じ名前があれば何も移さずに例外にする。途中で移せなければ、移した分を from へ戻してから例外にする | settings_tab.ps1, settings_view.ps1 |

## ファイルとフォルダ（`shared/core/fs.ps1`・`folder.ps1`・`shared/office/office_files.ps1`）

| 関数 | 入力 | 出力 | 概要 | 詳細 | 使用元 |
|---|---|---|---|---|---|
| `readListFile` | path | string[] | 行ファイルの空行以外の行（Trim しない）。無ければ空配列。読めなければ例外（空の一覧と取り違えない） | – | 取り込み中のファイル・元のフォルダの記録 など |
| `writeListFile` | path, lines | – | 行ファイルを UTF-8（BOM 付き）で保存 | – | 同上 |
| `writeTextLinesAtomic` | path, lines, encoding（既定は BOM 付き UTF-8） | – | 一時ファイルに書いてから置き換える（置き換えられなければ少し待って 5 回まで試す） | – | writeStatusFile, renameStatusIndexName など |
| `formatFileTime` | time | string | 取り込み一覧に記録する日時（`yyyy/MM/dd HH:mm:ss`）。更新の有無はこの文字列で比べる | [取り込み一覧](../indexing/ingest-list.md) | インデックス作成 |
| `copyFileShared` | sourcePath, destPath | – | 元のファイルを読み取りだけで開いてコピーする（ほかのアプリの読み書き・削除を妨げない）。インデックス作成は、このコピーを開く | [取り込み対象のファイルは書き換えない](../../safety/file-access.md#取り込み対象のファイルは書き換えない) | インデックス作成 |
| `setNotContentIndexed` | path, Recurse（スイッチ） | `@{Ok; Changed; Reason}` | フォルダ（またはファイル）に「内容のインデックスを作成しない」属性を付ける。Recurse 無しは path 自身と直下のファイルだけ、あれば下のすべてに付ける。既に付いているものは Changed に数えない。Recurse でフォルダを渡したときは、根フォルダに既に付いていれば（前回すべて付け終えたとみなし）下はたどらず `GetAttributes` 1 回で終え、付いていなければ下をすべて付け終えて成功したときにだけ最後に根へ付ける（途中で失敗した回は根に付けない）。1 件ずつの失敗・列挙そのものの失敗も例外にせず Ok=`$false`・Reason を返す | [Windows Search の対象から外す（`NotContentIndexed`）](../index-data/publish.md#windows-search-の対象から外すnotcontentindexed) | インデックス作成（`indexer_run.ps1`・`index_store.ps1`・`pack_store.ps1`・`source_map.ps1`） |
| `testNotContentIndexed` | path | bool | フォルダ（またはファイル）に「内容のインデックスを作成しない」属性が付いているか調べる。無ければ `$false` | 同上 | `newWorkerTmpDir`（親の作業フォルダに付いていれば、スレッドの作業フォルダにも同じ属性を付ける） |
| `getFolderKey` | dir | string（16 進 64 文字） | フォルダのパスを小文字にした SHA-256。名前付きミューテックス・イベントの名前に使う。FIPS モードの Windows でも動くよう、FIPS 準拠の実装（`SHA256CryptoServiceProvider`）を使う | – | newAppMutex, 画面（多重起動の防止） |
| `testWritableFolder` | dir | bool | フォルダにファイルを作れるか（試しに作ったファイルは閉じると消える）。無いフォルダは `$false` | [データの置き場所とパスの決め方](../structure/data.md) | getDataDir, 画面（置き場所の変更） |
| `getDataDir` | root（既定 `$rootDir`）, fallbackDir（必須） | string | 設定ファイルを置くフォルダ。root に書き込めれば root、書き込めなければ fallbackDir（`%LOCALAPPDATA%\tebunko` には置かない。`shared/core/data_dir.ps1`） | 同上 | `settings.ps1`（`getSettingsFilePath`） |
| `getSettingsFilePath` | root, defaultWorkDir | string | `setting.config` のパス。root に書き込めれば root 直下、書き込めなければ既定のワークスペース直下 | 同上 | `settings.ps1`（`$settingsFile`） |
| `testSettingsFileName` | name | bool | 設定ファイル（`setting.config`・`.tmp`・`.broken-<日時>[-<番号>]`）の名前か。既定のワークスペースが空かの判断と、中身として数える対象から外すのに使う | [設定ファイル](../structure/settings-file.md) | `testDefaultWorkspace`, 画面（`getCountedEntryPaths`） |
| `getSettingsFileView` | settingsFile, rootDir | `@{Path; Note}` | 設定タブに出す設定ファイルの場所と、ツールのフォルダに置いているか（書き込めなくて既定のワークスペースに置いているか）の注記。判断層（`settings_view.ps1`） | [設定タブ](../gui/settings-tab.md) | 画面（設定タブ） |
| `getCountedEntryPaths` | entryPaths, isDefaultWorkspace | string[] | ワークスペースに選んだフォルダの中身のうち「空でない」の数に入れるもの。選んだフォルダが既定のワークスペースなら、設定ファイルとその付属ファイル（`testSettingsFileName`）を外す。判断層（`settings_view.ps1`。カバレッジの分母に入る） | [設定タブ](../gui/settings-tab.md) | 画面（設定タブ） |
| `newAppMutex` | name（`gui` / `indexer`）, dir（既定 `$rootDir`） | `@{Mutex; Acquired}` | 同じツール（配置フォルダ）の処理を二重に動かさないための名前付きミューテックス（`Local\tebunko_<name>_<getFolderKey の鍵>`）。`Acquired` が `$false` なら、ほかで実行中。プロセスが終われば解放されるため、強制終了しても残らない | [インデックス作成のメインフロー](../indexing/flow.md) | インデックス作成（起動時。dir に `$workspace.Dir` を渡し、同じ `work` を使うものを 1 つにする）。画面は `getFolderKey` の鍵で同じ規則の名前を自前で作る（ウィンドウを前面に出すイベントの名前と共通の鍵を使うため） |
| `invokeWithNamedMutex` | mutexName, timeoutMilliseconds, action | action の出力 | 名前付きミューテックスを取って action を実行し、終わったら手放す。同じスレッドの入れ子は通す。時間を過ぎても取れなければ「排他の待ちが時間切れになりました」の例外。前の持ち主が手放さずに終わっていた（abandoned）ときは取れたものとして続ける | [プロセスとスレッド](../structure/threads.md) | invokeSettingsLocked |
| `normalizeFolderPath` | path | string | フォルダパスを 1 つの書き方にそろえる（前後の空白・`"` の除去、環境変数の展開、`/` → `\`、`\\?\` の除去、重なった `\` ・ `.` ・ `..` の解決、相対パスは `$rootDir` から、末尾の `\` の除去。ドライブ直下は `D:\` のまま）。解釈できない場合は書かれたとおり | [クロール対象フォルダと取り込み対象](../indexing/crawl.md#クロール対象フォルダgettargetfolders) | getTargetFolders, 画面 |
| `getPathUnderFolder` | path, folder | string / `$null` | path が folder 自身か下なら folder からの相対パス（folder 自身は空）、下でなければ `$null`（大文字・小文字と末尾の `\` を無視）。パスの文字数で切り出さないために使う | 同上 | インデックス作成・検索 |
| `getDriveTargets` | – | Dictionary（`Z:` → 割り当て先） | ネットワークドライブの割り当てを CIM（`Win32_LogicalDisk` の DriveType=4）で調べる（同じプロセスで 1 回だけ。`subst` は解決しない） | 同上 | getFolderPathAliases |
| `getFolderPathAliases` | path, drives（既定 `getDriveTargets`） | string[] | 同じ場所を指す別の書き方を path 自身を先頭にして返す（`Z:\見積` ⇔ `\\server\share\見積`） | 同上 | testSameFolder, 画面 |
| `testSameFolder` | a, b, drives（既定 `getDriveTargets`） | bool | 2 つのパスが同じフォルダを指すか（書き方の違い・ドライブの割り当てをたどる） | 同上 | getTargetFolders, getSourceFolderMap, 画面 |
| `testFolderUnder` | path, folder, drives（既定 `getDriveTargets`） | bool | path が folder 自身か folder の下のフォルダか（`testSameFolder` と同じく書き方の違いをたどる）。クロール対象フォルダが入れ子にならないかの確認に使う | 同上 | 画面（インデックスの追加・編集） |
| `getFolderLeafName` | folderPath | string | フォルダ名（ドライブ直下はドライブ名、UNC は共有名） | [取り込み一覧](../indexing/ingest-list.md) | newIndexName |
| `getExistingAncestorFolder` | folder | string | folder が今もあればそのまま、無ければその上の今もあるフォルダ。どこにも無ければ空 | [フォルダ選択ダイアログ（［参照…］）](../gui/common.md#フォルダ選択ダイアログ参照) | 画面（フォルダ選択） |
