# 部品ごとの関数（設定・ファイル）

扱うこと: 設定ファイル（`tebunko/core/settings.ps1`）とファイル・フォルダ操作（`shared/core/fs.ps1`・`folder.ps1`・`office_files.ps1`）の関数一覧（入力・出力・概要・使用元）。扱わないこと: 取り込み一覧・インデックスの管理（[部品ごとの関数（インデックス作成）](indexer.md)）。先に読むページ: [部品から関数一覧を引く](index.md)。

## 設定ファイル（`tebunko/core/settings.ps1`）

| 関数 | 入力 | 出力 | 概要 | 詳細 | 使用元 |
|---|---|---|---|---|---|
| `newSettings` | – | ordered hashtable | 設定の既定値（`targetFolders` `indexSources` `searchExcludes` `useRegex` `caseSensitive` `fileFilter` `includeShapes` `includeComments` `openMode` `workspaceFolder` `ingestThreads`） | [設定ファイル（setting.config）](../structure/settings-file.md) | readSettings |
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
| `readSearchOption` / `writeSearchOption` | path / option, path | `@{UseRegex; CaseSensitive; FileFilter; IncludeShapes; IncludeComments}` / – | 画面の検索条件。[元のファイルの特定・画面](#元のファイルの特定画面) を参照 | [続き](#元のファイルの特定画面) | 画面 |
| `readOpenMode` / `writeOpenMode` | path（既定 `$settingsFile`） / mode, path | string / – | 元のファイルの開き方（`openMode`）を読み書きする。無い・知らない値なら `normal` | [元のファイルを開く](../gui/open-file.md) | 画面 |
| `getDefaultWorkDir` | profileDir（既定は利用者のプロファイル） | string | 既定のワークスペース `<profileDir>\Documents\tebunko_ws`（OneDrive にリダイレクトされた「ドキュメント」は使わない） | [データの置き場所とパスの決め方](../structure/data.md) | getWorkDir, writeWorkspaceFolder, 画面 |
| `getWorkDir` | path（既定 `$settingsFile`） | string | `work` の置き場所（`workspaceFolder`。空なら既定。相対パスは設定ファイルのフォルダから、`%変数%` は展開する） | 同上 | `paths.ps1`（`$workspace`） |
| `writeWorkspaceFolder` | folder, path（既定 `$settingsFile`） | – | `work` の置き場所を保存する。既定の場所なら空で保存する | 同上 | 画面 |

## ファイルとフォルダ（`shared/core/fs.ps1`・`folder.ps1`・`shared/office/office_files.ps1`）

| 関数 | 入力 | 出力 | 概要 | 詳細 | 使用元 |
|---|---|---|---|---|---|
| `readListFile` | path | string[] | 行ファイルの空行以外の行（Trim しない）。無ければ空配列。読めなければ例外（空の一覧と取り違えない） | – | 取り込み中のファイル・元のフォルダの記録 など |
| `writeListFile` | path, lines | – | 行ファイルを UTF-8（BOM 付き）で保存 | – | 同上 |
| `writeTextLinesAtomic` | path, lines, encoding（既定は BOM 付き UTF-8） | – | 一時ファイルに書いてから置き換える（置き換えられなければ少し待って 5 回まで試す） | – | writeStatusFile, renameStatusIndexName など |
| `formatFileTime` | time | string | 取り込み一覧に記録する日時（`yyyy/MM/dd HH:mm:ss`）。更新の有無はこの文字列で比べる | [取り込み一覧](../indexing/ingest-list.md) | インデックス作成 |
| `copyFileShared` | sourcePath, destPath | – | 元のファイルを読み取りだけで開いてコピーする（ほかのアプリの読み書き・削除を妨げない）。インデックス作成は、このコピーを開く | [取り込み対象のファイルは書き換えない](../../safety/file-access.md#取り込み対象のファイルは書き換えない) | インデックス作成 |
| `getFolderKey` | dir | string（16 進 64 文字） | フォルダのパスを小文字にした SHA-256。名前付きミューテックス・イベントの名前に使う。FIPS モードの Windows でも動くよう、FIPS 準拠の実装（`SHA256CryptoServiceProvider`）を使う | – | newAppMutex, 画面（多重起動の防止） |
| `testWritableFolder` | dir | bool | フォルダにファイルを作れるか（試しに作ったファイルは閉じると消える）。無いフォルダは `$false` | [データの置き場所とパスの決め方](../structure/data.md) | getDataDir, 画面（置き場所の変更） |
| `getDataDir` | root（既定 `$rootDir`）, fallbackBase（既定 `%LOCALAPPDATA%`） | string | 設定ファイルを置くフォルダ。root に書き込めれば root、書き込めなければ `<fallbackBase>\tebunko\<getFolderKey の先頭 16 文字>`（`shared/core/data_dir.ps1`） | 同上 | `data_dir.ps1`（`$dataDir`） |
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
