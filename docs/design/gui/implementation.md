# 画面の実装

## 実装構成

| ファイル | 内容 |
|---|---|
| `tebunko.bat` | 起動用バッチ。`conhost.exe` を通して `powershell -NoProfile -STA -ExecutionPolicy RemoteSigned -WindowStyle Hidden -Command "..."` を `start` で起動する。`-Command` の中で `scripts` 配下の Mark-of-the-Web を消してから、`& 'scripts\tebunko\gui.ps1'` で画面を開く（PowerShell の起動は 1 回。`conhost.exe` を通すのは、既定のターミナルが Windows Terminal でも窓を隠すため）。ASCII・CRLF で書き、説明コメントは英語にする（cmd はコードページ・改行に敏感なため）。詳細は [画面の共通仕様](common.md#配布と実行ポリシーmark-of-the-web) |
| `scripts/tebunko/gui.ps1` | 画面の起動口。起動中の表示（`splash.xaml`。スクリプトを読み込む前に出し、区切りごとに `stepSplash` で描画を進め、画面の `ContentRendered` で閉じる）、Mark-of-the-Web の除去、多重起動の防止（[待たせ方の方針（画面を固まらせない）](#待たせ方の方針画面を固まらせない)）、ウィンドウとタブの読み込みと `x:Name` の対応表（`$ui`）の作成、各部品の読み込み（**順に意味がある**）、ウィンドウ全体のイベント、`ShowDialog`、描き終わってからの一覧の読み込み（`loadStartupData`。済むまで `Activated`・タブの切り替えでは読み直さない）。画面を開いている間使うスレッド（検索の司令 `SearchService`・画面の裏の仕事 `BackgroundQueue`）の用意と、閉じるときの順番（インデックス作成を止めてから閉じる。[閉じる](state-flow.md#閉じる)）。起動時に `GCSettings.LatencyMode` を `SustainedLowLatency` にする。判定・検索ロジックは持たない |
| `scripts/tebunko/xaml/tebunko.xaml` | 画面の枠（ウィンドウとタブ）。タブの中身は別ファイルにし、`gui.ps1` が読み込んで入れる |
| `scripts/tebunko/xaml/splash.xaml` | 起動中の表示（[ウィンドウ](index.md#ウィンドウ)）。早く出すため `theme.xaml` を読み込まず、色は `theme.xaml` と同じ値を直接書く |
| `scripts/tebunko/xaml/tab_index.xaml`<br>`tab_search.xaml`<br>`tab_settings.xaml`<br>`tab_kill.xaml` | 各タブの中身。イベントは書かず、`x:Name` だけを付ける。ルート要素には `StaticResource` を使う属性を置かない（自分の `Resources` より先に解決されるため）。`x:Name` が `gui.ps1` の一覧と食い違っていないことは `tests/meta/structure.Tests.ps1` で確かめる |
| `scripts/tebunko/xaml/dialog_index_edit.xaml` | インデックスの追加・編集のダイアログ（[追加・編集のダイアログ](index-tab.md#追加編集のダイアログ)）。追加と編集で同じ定義を使い、表題・説明・注意書きを `ui/index_tab.ps1` で変える |
| `scripts/tebunko/xaml/dialog_indexing_confirm.xaml` | インデックス作成の確認ダイアログ（[インデックス作成の確認ダイアログ](index-tab.md#インデックス作成の確認ダイアログ)）。インデックスごとの取り込み対象の件数（受け渡しの口の `Plan`）を一覧にする。`ui/indexing_tab.ps1` が開く |
| `scripts/shared/xaml/dialog_confirm.xaml` | 確認ダイアログ（[確認ダイアログ](common.md#確認ダイアログshowconfirm)）。見出し・結果の一覧（`ItemsControl` に `ConfirmFact` をバインド）・補足だけを定義し、ボタンは場面ごとに違うため `shared/ui/shell.ps1` の `showConfirm` が組み立てて `ButtonPanel` / `ChoicePanel` に入れる |
| `scripts/shared/xaml/theme.xaml` | 画面の見た目（色・文字・コントロールの形）の共通定義 |
| `scripts/shared/ui/types.ps1`<br>`scripts/tebunko/ui/types.ps1` | 画面で使う型（PowerShell class。[実行時コンパイル（csc.exe）を使わない](#実行時コンパイルcscexeを使わない)）。継承元（`NotifyBase`）と共通の型（`ConfirmFact`）は shared に置き、先に読み込む（型の解決は読み込む順に依存するため、`tests/meta/structure.Tests.ps1` で確かめる）。主な型は次のとおり。<br>・`FileGroup`：結果の表の、元のファイル 1 つ分の見出し（[ファイルごとにまとめた表示](search-tab.md#ファイルごとにまとめた表示)）。ヒットは生のまま `Hits` に持ち、表の行は開いたとき・絞り込み・並べ替え・出力のときに作る<br>・`HitRow`：結果の 1 行。件数が多いので**生成時は生データのみ**を持ち、表示用（強調セグメント・DisplayLine・「セル」列）は**画面に出た行だけ** `Prepare()` で作る（`ResultGrid` の `LoadingRow` から呼び、`INotifyPropertyChanged` で反映）。選択行のプレビュー（[選択行のプレビュー](search-tab.md#選択行のプレビュー)）の表は `BuildPreview`（`readPackContext` で読んだ行から `PreviewTable` を作る）で組み立てる<br>・`PreviewColumn`：列の幅。列見出しと各行のセルで共有し、見出しの `Thumb` のドラッグで `SetWidth` を呼ぶと列全体に反映する<br>・`IndexNode`：検索対象のツリー（[検索対象のツリー（No.13）](search-tab.md#検索対象のツリーno13)）。展開時のフォルダ読み込み・3 状態のチェック・検索範囲 `SearchTarget`／除外 `SearchExclude` の組み立てを持つ。チェックは OneWay バインドとし、クリック（`Toggle`）・展開（`LoadChildren`）はイベントで駆動する（PS class はプロパティのセッターにロジックを書けないため） |
| `scripts/shared/ui/app_host.ps1` | 画面の土台。XAML の読み込み（`loadXaml`。`ParserContext.BaseUri` にそのファイルの場所を渡し、`theme.xaml` への相対参照を解決する）、ウィンドウの読み込みとアイコン（`loadWindow`）、色の取得（`themeBrush`）、予期しないエラーの記録（`writeErrorLog`） |
| `scripts/shared/ui/shell.ps1` | 画面の共通部品。ステータス表示（`setStatus`）、メッセージ（`showMessage`）、確認ダイアログ（`showConfirm`。結果の行は `factGone` / `factKept` / `factNext` で `ConfirmFact` を作り、選択肢のボタンは `newChoiceContent` で 2 行のボタンにする。ボタンの `Click` と `ContentRendered` は `GetNewClosure()` で値を取り込んだスクリプトブロックにし、スクリプトスコープの変数に置かない）、例外を拾う `safe`、タイマー（`newTimer`）、画面の裏の仕事（`startJob`。`gui.ps1` が用意する `BackgroundQueue` に渡す。[待たせ方の方針（画面を固まらせない）](#待たせ方の方針画面を固まらせない)） |
| `scripts/shared/ui/folder_dialog.ps1` | Windows 標準のフォルダ選択（`selectFolder`。[フォルダ選択ダイアログ（［参照…］）](index-tab.md#フォルダ選択ダイアログ参照)）。WinForms 内部の `IFileDialog` をリフレクションで呼んで開き、開けなければ `OpenFileDialog` でフォルダの中に入って選んでもらう（[実行時コンパイル（csc.exe）を使わない](#実行時コンパイルcscexeを使わない)）。フォルダのドラッグ＆ドロップの判定（`getDroppedFolders`・`onFolderDragOver`）も置く |
| `scripts/tebunko/ui/index_tab.ps1`<br>`indexing_tab.ps1` | ［1 インデックス管理］タブ（[［1 インデックス管理］タブ](index-tab.md)）。インデックスの一覧・追加・編集・削除と、インデックス作成の開始（`IndexingSession`）・中止・進み具合の表示。インデックス一覧のほかでの変更は、`getTargetsKey`（保存されている一覧を比べるための文字列）をウィンドウがアクティブになったときに比べて検出する（[画面での読み書き](../architecture/settings-file.md#画面での読み書き)） |
| `scripts/tebunko/ui/search_tab.ps1`<br>`result_list.ps1`<br>`preview.ps1`<br>`open_source.ps1`<br>`index_tree.ps1` | ［2 検索］タブ（[［2 検索］タブ](search-tab.md)）。検索の実行（`search_tab.ps1`）、結果の表（ファイルごとの見出しと行・絞り込み・並べ替え。`result_list.ps1`）、選択行のプレビュー（`preview.ps1`）、元のファイルを開く処理（Excel COM。[元のファイルを開く](search-tab.md#元のファイルを開く)。`open_source.ps1`）、検索対象インデックスのツリー（`index_tree.ps1`） |
| `scripts/tebunko/ui/settings_tab.ps1` | ［8 設定］タブ（[［8 設定］タブ](settings-tab.md)）。ワークスペースを変えたら、その場で切り替える |
| `scripts/tebunko/ui/process_tab.ps1` | ［9 プロセス停止］タブ（[［9 プロセス停止］タブ](process-tab.md)）。定期的な一覧の取り直しとタブ見出しの ⚠ の判定は `startJob` で行う（ボタンでの取り直しは画面のスレッドで行う） |
| `scripts/tebunko/ui/index_view.ps1`<br>`indexing_view.ps1`<br>`search_view.ps1`<br>`preview_view.ps1`<br>`settings_view.ps1` | 画面の**判断層**。「何を出すか」を決める部分を `$ui` に触らない関数として切り出したもの（入力は素の値、出力は素の値）。色は意味（`info` / `ok` / `warn` / `ng` / `gray`）で返し、実際の色はタブ側で対応表から引く。Pester でテストする（[画面の単体テスト](../testing/index.md#画面の単体テスト)） |
| `scripts/tebunko/lib.ps1` | 画面とインデックス作成（`tebunko/indexer.ps1`）で使う部品の読み込み口（下表の関数。実体は `core/`・`index/`・`indexer/`・`search/` と `shared/`）。検索の司令・画面の裏の仕事のスレッドも、始めたときに 1 回だけこれを読み込む |

`tebunko/lib.ps1` の関数のうち、画面が使うものは次のとおり（引数・戻り値・内容は [共通モジュール](../architecture/modules.md#関数一覧) の関数一覧にある）：`isValidRegex`・`getIndexPackFiles`・`testIndexExists`・`getIndexSummary`・`newSearchRegex`・`searchPackIndex`・`toSearchResultLines`・`writeSearchResult`・`readPackContext`・`getSourceLocation`・`resolveSourcePath`・`findMovedSource`・`setIndexSourceFolder`・`getTargetFolders`・`writeTargetFolders`・`normalizeFolderPath`・`readSearchOption`・`writeSearchOption`・`readOpenMode`・`writeOpenMode`・`getSearchIndexes`・`newSearchService`・`newSearchRequest`・`newIndexingSession`・`newIndexerChannel`・`readIndexingProgress`・`requestIndexingStop`・`answerIndexingPlan`・`testIndexerRunning`・`readSearchExcludes`・`writeSearchExcludes`・`getIndexingState`・`getOfficeProcesses`・`stopOfficeProcesses`。

### 待たせ方の方針（画面を固まらせない）

**画面が「応答なし」になることは、処理が長いこと自体よりも利用者を不安にさせる**（壊れたと思って強制終了され、インデックス作成が中断して Office のプロセスが残る）。そのため、次を守る。

| 方針 | 具体 |
|---|---|
| **画面のスレッドで 0.2 秒を超える可能性のある処理をしない** | ファイル数・行数に比例する処理（検索、取り込み一覧の集計、インデックスの削除、TSV の数え上げ）は別のスレッド（検索の司令・画面の裏の仕事。[スレッドの一覧](../architecture/threads.md#スレッドの一覧)）で行い、終わったら画面のスレッドで結果を反映する。終わるまでは、その操作のボタンを無効にして何をしているかをステータスに出す |
| **繰り返し読むものは小さくする** | 1 秒ごとに読むものは、受け渡しの口の進み具合（`readIndexingProgress`。メモリ上の値）にする。数万行の取り込み一覧を毎秒読み直すと、その間ずっと画面が止まる（5 万行で 1 回約 1.4 秒）。書く側（インデクサ）が数えて進み具合に入れる |
| **無言の時間を作らない** | 待ちが数秒を超えうる場所では「何をしているか」を必ず出す（`クロールしています…` ＋ 今見ているフォルダ、`インデックス作成を終えています…` ＋ 今の後片付け、`検索対象のファイルを確認しています…（N 件）`、`インデックス […] の TSV を削除しています…`）。件数が増えるのが見えれば、止まっていないと分かる |
| **長い処理は止められるようにする** | インデックス作成は［中止］（ファイルの切れ目で止まる・続きから再開できる）、検索は［中止］（それまでの結果は残す） |
| **終わったことを知らせる** | インデックス作成が終わったとき、進み具合とステータスで知らせる（タスクバーのボタンを光らせる `FlashWindowEx` は P/Invoke が要るため使わない。[実行時コンパイル（csc.exe）を使わない](#実行時コンパイルcscexeを使わない)） |

数万件規模での実測（この方針の根拠。Windows PowerShell 5.1・ローカルディスク。1 万 = 元ファイル 1 万件（取り込み一覧 1 万行・TSV 2 万件）、5 万 = 取り込み一覧 5 万行）:

| 処理 | 1 万 | 5 万 | 扱い |
|---|---|---|---|
| 取り込み一覧の読み込み（`readStatusFile`） | 0.32 秒 | 1.4 秒 | 1 行ごとに関数を呼ばない。毎秒は読まず、受け渡しの口の進み具合を読む |
| 取り込み一覧の書き出し（`writeStatusFile`） | 0.08 秒 | 0.22 秒 | 1 行ごとに関数・パイプラインを使わない（インデックス作成の開始・終了時の待ちに効く） |
| 状態の集計（`getIndexingState`） | 0.91 秒 | 5.3 秒 | 別スレッドで行う（`refreshIndexingState`）。日時の解析は必要なときだけ行う |
| 検索対象の数え上げ（集約ファイルにする前の TSV の列挙） | 2.2 秒（TSV 2 万件） | – | .NET で列挙する。検索スレッドで行い、途中の件数を表示する。今は集約ファイル（フォルダ・拡張子ごとに 1 つ）だけを列挙する（`getIndexPackFiles`） |
| インデックスの削除 | 6.7 秒（1 万フォルダ） | – | 別スレッドで行う（削除中は操作を止めてステータスに出す） |

> PowerShell 5.1 では**関数呼び出しが 1 回あたり約 50 マイクロ秒**かかる。5 万行で 1 行 1 回呼ぶと、それだけで約 2.7 秒になる。数万件を回すところでは、関数呼び出し・パイプライン（`ForEach-Object` / `Where-Object`）・`$obj.$名前` の動的アクセスを避ける（普通の場所では読みやすさを優先してよい）。

別スレッドとタイマーの扱い：

- スレッドの分け方・寿命・閉じる順番は [プロセスとスレッド](../architecture/threads.md) にまとめる。画面のプロセスの中で、画面・検索・画面の裏の仕事・インデックス作成が別々のスレッドで動く。
- 検索は検索の司令のスレッド（`SearchService`。画面を開いている間 1 つ）が、要求（`newSearchRequest`）を 1 つずつ実行する。lib.ps1 の読み込みと照合のプールの用意は、スレッドを始めたときに 1 回だけ行う。結果は要求の中のキュー（`ConcurrentQueue`）に入れ、画面の `DispatcherTimer`（0.1 秒間隔）で取り出して表に加える（`pumpSearch`）。1 回に取り出す量は件数ではなく時間で区切る（件数で区切ると、ヒットが多いときに画面が止まる）。中止は要求の `Stop` で伝える。新しい要求を渡すと、前の要求の `Stop` を立てる。
- 集約ファイルの件数取得・取り込み一覧の集計・インデックスの削除・選択行のプレビューの読み込み・［9 プロセス停止］の定期的な一覧の取り直しとタブ見出しの ⚠ の判定は `startJob`（`shell.ps1`）で実行する。`startJob` は画面の裏の仕事のスレッド（`BackgroundQueue`。2 つ。長い仕事の間もプレビューが待たないため）に渡す。各スレッドは最初の仕事の前に lib.ps1 を 1 回だけ読み込む。終わったかは `DispatcherTimer`（`jobTimer`。仕事がある間だけ 0.05 秒間隔）で確かめ、画面のスレッドで `onDone` を呼ぶ。同じ集計が重ならないよう、実行中なら「もう一度」の印だけを立てて、終わってからやり直す。
- インデックス作成は `newIndexingSession` で画面のプロセスの中のスレッド（MTA・BelowNormal）で実行し、`DispatcherTimer`（1 秒間隔）で `IsRunning` と受け渡しの口の進み具合（`readIndexingProgress`）を確認する（[インデックス作成の進み具合](index-tab.md#インデックス作成の進み具合)）。終わったら受け渡しの口の終了コード（`GetExitCode`）で表示を分け（[中止・終了・ログ](index-tab.md#中止終了ログ)）、`Close` でスレッドを片づける。
- 進み具合の段階が `確認` になったら、インデックス作成の確認ダイアログを開く（[インデックス作成の確認ダイアログ](index-tab.md#インデックス作成の確認ダイアログ)）。**ダイアログを開いている間も `DispatcherTimer` は動く**（`ShowDialog` は入れ子のメッセージループのため）。開く前に「開いた」ことにして、二重に開かないようにする。
- 多重起動は、ツールの配置フォルダごとの名前付き Mutex で防ぐ。2 つ目のプロセスは、同じくフォルダごとの名前付きイベント（`Local\tebunko_gui_activate_<フォルダのハッシュ>`）を合図して終了し、すでに開いている画面が `DispatcherTimer`（0.3 秒間隔）でそれを確認して `Window.Activate()`（＋最小化なら復帰）で前面に出す。別アプリが前面のときはフォーカスを奪えないことがある（OS の制限。P/Invoke の `SetForegroundWindow` は使わない。[実行時コンパイル（csc.exe）を使わない](#実行時コンパイルcscexeを使わない)）。

注意（PowerShell 5.1）：`System.Collections.Generic.List[object]` を `@()` で配列にする・`[object[]]` 引数に渡すと例外（Argument types do not match）になる。結果のリストは `List[psobject]` にするか `.ToArray()` で返す。また PS class のメソッド内では、クラスのプロパティと同じ名前（大文字小文字を区別しない）のローカル変数は使えない（`$this.` の付け忘れとみなされる）ため、別名にする。

### 実行時コンパイル（csc.exe）を使わない

資産管理・EDR は「`powershell.exe` が `csc.exe` を起動して `%TEMP%` の一時 DLL を読み込む」動き（実行時コンパイル。MITRE ATT&CK T1027.004）を、良性でも拾うことがある。tebunko はこれを一切出さない設計にする（`tests/meta/safety.Tests.ps1` の「実行時にコードをコンパイルしない」が確かめる。[危険とされる処理の検査結果](../../safety/checks.md#検査項目と結果)）。

- **画面で使う型は PowerShell class**（`HitRow`・`FileGroup`・`Segment`・`PreviewColumn`／`PreviewCell`／`PreviewRow`／`PreviewTable`・`ProcRow`・`FailRow`・`PlanRow`・`FolderItem`・`SearchTarget`・`SearchExclude`・`IndexNode`・`ConfirmFact`）。PS class はエンジンがメモリ内で用意し、`csc.exe`・一時 DLL を出さない。`INotifyPropertyChanged` は `NotifyBase` を継承して実装する。
- **検索・集約ファイルの読み取りは .NET を直接呼ぶ**（`searchPackFiles`／`readPackPlaces`／`readPackContext`。`StreamReader`＋`[regex]`）。
- **Win32 API（P/Invoke）と、インターフェース定義が要る COM を使わない**。代わりに次を使う。

  | 目的 | 使うもの | 使わないもの（実行時コンパイルが要る） |
  |---|---|---|
  | フォルダ選択 | Windows 標準のエクスプローラー形式のダイアログ（`IFileOpenDialog`）。インターフェースの定義は WinForms が内部に持つもの（`FileDialogNative+IFileDialog`）をリフレクションで使う（[フォルダ選択ダイアログ（［参照…］）](index-tab.md#フォルダ選択ダイアログ参照)） | `IFileOpenDialog` のインターフェースを自分で定義すること |
  | ドライブの列挙・ネットワークドライブの割り当て先 | `System.IO.DriveInfo`・CIM（`Win32_LogicalDisk`） | `QueryDosDevice` |
  | 多重起動時の前面化 | `Window.Activate` | `SetForegroundWindow`／`AttachThreadInput` |
  | 完了の通知 | 進み具合とステータスの表示 | `FlashWindowEx` |
  | タスクバー・タイトルバーのアイコン | `Window.Icon` | `SetCurrentProcessExplicitAppUserModelID` |

- 結果、起動〜検索を通じて `Add-Type`（実行時コンパイル）は 0 回、`csc.exe` の起動・一時 DLL の生成も 0（実起動で確認）。読み込むのは Microsoft 署名済みの GAC アセンブリ（WPF・WinForms）だけ。
- 代償と対処：PS class はメソッド呼び出しが C# より遅い（実測: 検索コアは C# の約 2.6 倍だが 50 万行で約 0.6 秒。結果行 `HitRow` の一括生成は 1 万件で約 10 秒と遅いため、**可視行だけ作る遅延描画**（`HitRow.Prepare`＋`LoadingRow`。軽量生成は 1 万件で約 0.56 秒）で回避する）。PS class はプロパティのセッターにロジックを書けないため、列幅・チェック・選択は `Set*` メソッドやイベントで変える。`subst` ドライブは実体のパスに解決しない。タスクバーのボタンは PowerShell と同じグループにまとめられる（アイコン自体は tebunko のものが出る。[表示・アクセシビリティ](common.md#表示アクセシビリティ)）。

## 今後の検討

- Word・PowerPoint の該当ページ・スライドへの移動（今は Excel だけ該当セルを選んで開く。Word・PowerPoint の COM 操作を足す必要がある）

## 決めたこと

画面を作ったときに、選択肢から決めたこと。

| # | 事項 | 選択肢 | 決定 |
|---|---|---|---|
| U1 | 入口の名前 | `0_設定.bat` / `tebunko.bat` など | **`tebunko.bat`**。画面がすべての操作の入口で `.bat` もこれ 1 本だけのため、番号を付けず、ツールの名前にする |
| U2 | 正規表現の既定 | オン / オフ（文字どおり） | **オフ** |
| U3 | 検索結果の件数の上限 | 1,000 / 10,000 / 上限なし | **10,000 件** |
| U4 | インデックス作成中の検索 | 許可（注意を表示） / 禁止 | **許可**（注意を表示） |
| U5 | 複数ワードの一括検索 | 設ける / 設けない | **設けない** |
| U6 | 元のファイルを開くときの Excel | 起動中の Excel を使う / 常に新しく起動する / 既定のアプリで開くだけ | **起動中の Excel を使う**（無ければ起動）。該当セルを選択できる |
| U7 | ファイル出力の対象 | 絞り込み後の表示分 / 全件 | **絞り込み後の表示分** |

## 既知の制約

確度: ◎ = 動作確認済み、○ = コード上明らか、△ = 推定。

| # | 内容 | 確度 |
|---|---|---|
| 1 | 元のファイルのパスは、取り込み時に記録したクロール対象フォルダから組み立てる。取り込み後に元のフォルダを移動・名前変更した場合は、開くときにフォルダを選び直す必要がある（選んだフォルダは置き換えとして記録し、同じフォルダの下は次から聞かない。[元のファイルが見つからないとき（元のフォルダを設定する）](search-tab.md#元のファイルが見つからないとき元のフォルダを設定する)） | ○ |
| 2 | `元のフォルダ.txt` の無いインデックス（インデックス作成を 1 回実行すると作られる）を別の場所にコピーした場合は、元の場所が分からないため、開くときにフォルダを選ぶ必要がある | ○ |
| 3 | バックグラウンドの判定は `MainWindowHandle` による。画面に表示中でも、すべてのウィンドウを閉じた直後などはバックグラウンドと判定される（[プロセスの表](process-tab.md#プロセスの表)） | △ |
| 4 | フォルダ選択は Windows 標準のダイアログのため、フォルダの中のファイル（Office ファイルがあるか）は一覧に出ない（[フォルダ選択ダイアログ（［参照…］）](index-tab.md#フォルダ選択ダイアログ参照)）。WinForms の内部の型を使うため、Windows PowerShell 5.1（.NET Framework）が前提。内部の型が使えないときは、`OpenFileDialog` でフォルダの中に入って［開く］を押す形になる | ○ |
| 5 | 集約ファイルが非常に多い場合、件数を数え終わるまで状態表示が `確認中…` になる | △ |

---

> **16. テスト**（[画面の単体テスト](../testing/index.md#画面の単体テスト) 単体テスト・[画面の確認](../testing/index.md#画面の確認) 画面の確認）→ [テスト](../testing/index.md)
