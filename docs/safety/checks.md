# 危険とされる処理の検査結果

## 検査項目と結果

検査対象は `scripts/` 配下のすべての `.ps1` である。下の確認コマンドは、説明のコメント行（「〜は使わない」などの記述）を除いて検索する。「該当 0 件」は、コマンドが 1 件も出力しないことを意味する。同じ検査を `tests/meta/safety.Tests.ps1` の「危険な処理を使っていないこと」が行い、CI（`test.yml` の `test`。main へのマージに必須）で PR ごとに実行する。

確認コマンドの準備（リポジトリ直下、または配布 zip を展開したフォルダで実行する）:

```powershell
function scan { param([string[]]$Pattern)
    Get-ChildItem .\scripts -Recurse -Filter *.ps1 | Select-String -Pattern $Pattern |
        Where-Object { $_.Line -notmatch '^\s*#' } }
```

| # | 危険とされる処理 | 結果 | 確認コマンド |
|---|---|---|---|
| 1 | 文字列を式として実行（`Invoke-Expression` / 行頭を含む `iex` / `[ScriptBlock]::Create` / `NewScriptBlock` / `InvokeScript` / 型を変数に入れた `::Create`。`scan` は大文字小文字を区別せず、行頭が `#` の行だけを除く。ブロックコメントの終わり `#>` の後ろのコード・同じ行の `#>` が複数のときの間のコードは、`findBannedCode`（テストと道具）が見る） | 該当 0 件（配布する単一 .ps1 の末尾 1 行だけが例外。[下記](#文字列を式として実行1の例外)） | `scan 'Invoke-Expression','^iex(?![-\w])','[^-\w]iex(?![-\w])','ScriptBlock\]\s*::\s*Create','NewScriptBlock','InvokeScript','\$\{?[\w:]+\}?\s*::\s*Create'` |
| 2 | 難読化されたコマンド（Base64・`-EncodedCommand`） | 該当 0 件 | `scan 'FromBase64String','EncodedCommand'` |
| 3 | ネットワーク通信 | 該当 0 件 | `scan 'Invoke-WebRequest','Invoke-RestMethod','WebClient','HttpClient','Net\.Sockets','Start-BitsTransfer','DownloadFile','DownloadString','System\.Net\.'` |
| 4 | Windows API の直接呼び出し（P/Invoke） | 該当 0 件 | `scan 'DllImport','GetDelegateForFunctionPointer'` |
| 5 | C# ソースの実行時コンパイル | 該当 0 件 | `scan 'Add-Type'` で出る行がすべて `-AssemblyName` であること（下表の 3 行） |
| 6 | レジストリの読み書き | 該当 0 件 | 表の下の「レジストリの確認コマンド」 |
| 7 | 権限・サービス・自動起動の変更 | 該当 0 件 | `scan 'Set-Acl','icacls','schtasks','New-Service','Start-Service','\bsc\.exe','RunAs'` |
| 8 | 実行ポリシーの恒久変更・`Bypass` | 該当 0 件 | `scan 'Set-ExecutionPolicy','ExecutionPolicy\s+Bypass'` と `Select-String -Path .\tebunko.bat -Pattern 'Bypass','Set-ExecutionPolicy'` |
| 9 | 資格情報の入力要求・保存 | 該当 0 件 | `scan 'Get-Credential','ConvertTo-SecureString','PSCredential'` |
| 10 | リモート実行 | 該当 0 件 | `scan 'Invoke-Command','New-PSSession','Enter-PSSession','WinRM'` |
| 11 | 外部プロセスの起動 | 決まった起動だけ（[外部プロセスの起動](#外部プロセスの起動)） | `scan 'Start-Process'` |
| 12 | プロセスの強制終了 | 1 か所のみ（[前回残った Office の終了（起動時の確認）](disclosure.md#前回残った-office-の終了起動時の確認)）。止める前に照合（`isSameOfficeProcess`）を通す。`taskkill`・`CloseMainWindow`・WMI の `Terminate` などの別の止め方は 0 件 | `scan 'Stop-Process'`、`scan 'taskkill','CloseMainWindow','Terminate','spps'` |
| 13 | 壊れたハッシュ（MD5・SHA-1）、FIPS 準拠でないハッシュの実装 | 該当 0 件 | `scan 'MD5','SHA1','RIPEMD','SHA256Managed','SHA384Managed','SHA512Managed','HashAlgorithm\]::Create'` |
| 14 | 内部の型（`NonPublic`）のリフレクションでの呼び出し | 1 か所のみ（下） | `scan 'NonPublic','Reflection\.BindingFlags'` |
| 15 | プロセスの強制終了（`.Kill()`） | 3 か所のみ（Office の終了の `quitApp`・制限時間の見張り・画面を閉じるときの `KillOffice`）。どれも、同じ関数の中で、止める前に照合（`isSameOfficeProcess`、または `compareOfficeProcess` の結果が `Same` のときの枝）を通し、控えたときと同じプロセス（名前・起動時刻）だけを止める。制限時間の見張りと `KillOffice` は、さらに、取り込み中に見える窓があった（利用者が使った）PowerPoint を止めない（`quitApp` は Office を終わらせる通常の手順のため、この印を見ない）。12 の `Stop-Process` も同じ | `scan '\.Kill\('`（`tests/meta/safety.Tests.ps1` が、場所と照合を構文で確かめる） |

**レジストリの確認コマンド（6）**：.NET の型（`Microsoft.Win32.Registry`）・PowerShell のプロバイダー・COM（`WScript.Shell`）・WMI（`StdRegProv`）・`reg.exe` の書き方も見つける。パターンは `tools/script_rules.ps1` の `Registry` と同じ語の集まり（`|` を使わず、1 語ずつ別の文字列にして並べた）。

```powershell
scan 'HKLM','HKCU','HKEY_','HKCR','\bHKU\b','HKCC','Set-ItemProperty','New-ItemProperty','Remove-ItemProperty','Clear-ItemProperty','Rename-ItemProperty','Copy-ItemProperty','Move-ItemProperty','Registry::','Microsoft\.Win32\.Registry','Registry\]::','RegistryKey','RegistryHive','RegistryView','RegistryValueKind','RegistryRights','OpenSubKey','OpenBaseKey','OpenRemoteBaseKey','CreateSubKey','DeleteSubKey','PSProvider\s+\W?Registry','\bRegRead\b','\bRegWrite\b','\bRegDelete\b','StdRegProv','\breg(\.exe)?\s+query\b','\breg(\.exe)?\s+add\b','\breg(\.exe)?\s+delete\b','\breg(\.exe)?\s+copy\b','\breg(\.exe)?\s+save\b','\breg(\.exe)?\s+load\b','\breg(\.exe)?\s+unload\b','\breg(\.exe)?\s+restore\b','\breg(\.exe)?\s+import\b','\breg(\.exe)?\s+export\b','\breg(\.exe)?\s+compare\b','\breg(\.exe)?\s+flags\b'
```

クロール対象が共有フォルダにある場合のファイルの読み取りは、Windows のファイル共有（SMB）を通る。これは OS のファイル API によるもので、本ツールが通信処理を持つわけではない。

`Add-Type` は 3 か所あるが、いずれも Windows 標準アセンブリの読み込み（`-AssemblyName`）であり、コードのコンパイルではない。

| 場所 | 読み込むアセンブリ | 用途 |
|---|---|---|
| `tebunko/gui.ps1:8` | `PresentationFramework` / `PresentationCore` / `WindowsBase` / `System.Windows.Forms` | 画面（WPF）の表示、フォルダ選択（WinForms） |
| `shared/office/office_reader.ps1:15-16` | `System.IO.Compression` / `System.IO.Compression.FileSystem` | `.xlsx` / `.docx` / `.pptx`（ZIP）の読み取り |

実行時コンパイル（`csc.exe` の起動）を使わないのは設計方針である。検索処理・画面の型定義などは C# を使わず PowerShell と .NET の直接呼び出しで書いている（[画面の実装](../design/gui/implementation.md#実行時コンパイルcscexeを使わない)）。実行時に一時フォルダへコンパイル結果を書き出す動作が無いため、EDR やアプリケーション制御（AppLocker・WDAC）との衝突も避けられる。`tests/meta/safety.Tests.ps1` の「実行時にコードをコンパイルしない」が、`Add-Type` は標準アセンブリの読み込み（`-AssemblyName`）だけであることを確かめる。

**内部の型のリフレクション（14）**：フォルダ選択（`shared/ui/folder_dialog.ps1` の `showFolderPicker`）だけが、WinForms が内部に持つ `IFileOpenDialog` の定義（`FileDialogNative+IFileDialog`）をリフレクション（`BindingFlags` の `NonPublic`）で呼ぶ。Windows 標準のエクスプローラー形式のダイアログを、実行時コンパイルなしで開くためである（[フォルダ選択ダイアログ（［参照…］）](../design/gui/common.md#フォルダ選択ダイアログ参照)）。呼ぶのは .NET Framework に署名付きで入っている WinForms のダイアログの処理だけで、メモリの書き換えや AMSI などの検査の仕組みには触れない。開いている間に `csc.exe` などの子プロセスと一時フォルダへの DLL の書き出しが無いことを実機で確かめた。なお、Windows PowerShell 5.1 は `Add-Type`・`BindingFlags`・`NonPublic`・`GetMethod`・`IntPtr`・`InteropServices` などの語を含むスクリプトを、ログの設定に関わらずイベントログ（Microsoft-Windows-PowerShell/Operational の 4104、警告）に自動で記録する。本ツールはこの変更の前から `Add-Type` などで記録の対象である。`tests/meta/safety.Tests.ps1` が、`NonPublic` を使うのがこの 1 か所だけであることを確かめる。

## 文字列を式として実行（1）の例外

`scripts/` の中は上の確認コマンドのとおり 0 件である。配布する単一 .ps1（`tools/new_single_script.ps1` が作る 1 本）だけは、末尾の 1 行に `. ([scriptblock]::Create(${tebunkoSource})) @PSBoundParameters` がある。

- **理由**: Windows PowerShell 5.1 は、BOM の無い UTF-8 のファイルを CP932 で読み、日本語の文字の最後のバイトが次の記号を飲み込んで構文が壊れる。日本語の本体をヒアストリング（データ）に入れ、ScriptBlock にして実行することで、BOM が落ちても動かす（[単一 .ps1 の包み](../design/structure/single-script.md)）。
- **範囲**: 作った単一 .ps1 の最後の文の 1 行だけ。本体は、包む前に禁止の語の検査（`tools/script_rules.ps1`）を通っている。
- **検査**: 包んだ結果の全体を、構文木で確かめる。ScriptBlock の Create・`Invoke-Expression`・`iex` の呼び出しが 1 件だけで、ファイルの最後の文であり、その文の全文が決めた 1 行と一字一句同じであること（同じ文に別の式を足したら止まる）。その文を除いた全体（頭・尾・本体の文字列）に、禁止の語（`tools/script_rules.ps1` の一覧の全分類）が無いこと。1 つでも外れると、包む道具が何も書き出さずに止まる（`tests/tools/single_script_guard.Tests.ps1` が、足した Create・`iex`・引数の違い・2 行目・本体の中の語を止めることを表で確かめる）。
- **実機で確かめた結果**: Windows 11 の Microsoft Defender ウイルス対策（既定の設定。リアルタイム保護・振る舞い監視・AMSI が有効）で、この形を含む起動を動かして確かめた（下の「[ウイルス対策ソフト・EDR での確かめ](#ウイルス対策ソフトedr-での確かめ)」）。単一 .ps1 版の起動とインデックス作成で、Defender の検知・隔離は出なかった。PowerShell 自身が残す 4104（警告レベル）の自動の記録は、この行の `[scriptblock]::Create` を含めて出る（件数は下の表の下のとおり）。Microsoft Defender for Endpoint（EDR 本体）では未確認。検知されたら、方式を決め直す。

## Office ファイルを開くときの設定

インデックス作成のために Excel・Word・PowerPoint を COM で操作するが、アプリを起動したときに次の設定を行う（`shared/office/office_app.ps1:79-96`）。`tests/meta/safety.Tests.ps1` の「Office ファイルを安全に開くこと」がこれを確かめる。**利用者の PowerPoint には接続しない**（1 つのセッションに 1 つのプロセスしか持てず、接続すると下の設定を利用者のアプリまで書き換えてしまうため）。起動する前に自分のセッションに `POWERPNT` が無いかを確かめ、あれば設定を行わずに例外にする（[Office アプリ（Excel・Word・PowerPoint）の管理](../design/indexing/office-apps.md#office-アプリexcelwordpowerpointの管理)「PowerPoint」）。

| 設定 | 値 | 意味 |
|---|---|---|
| `AutomationSecurity` | `3`（`msoAutomationSecurityForceDisable`） | **マクロを強制的に無効にする**。マクロ有効ファイル（`.xlsm` / `.docm` / `.pptm`）でもマクロは実行されない（Excel・Word・PowerPoint とも） |
| `AutomationSecurity`（検索結果から元の Excel ファイルを開くとき） | `2`（`msoAutomationSecurityByUI`） | 利用者の Excel のマクロの設定（セキュリティセンター）に従う。設定を緩める値（`1`）にはしない。開いた後は元の値に戻す（`tebunko/ui/open_source.ps1` の `openInExcel`）。上の `3` はインデクサ（取り込み）の設定で、利用者が編集する用途の「開く」とは分けている |
| `EnableEvents` | `$false` | `Workbook_Open` などのイベントマクロを発火させない（Excel） |
| `AskToUpdateLinks` | `$false`、`Workbooks.Open` の `UpdateLinks` = `0` | **外部リンクを更新しない**。他ブックや外部データソースへのアクセスが発生しない（Excel） |
| `Visible` | `$false` | 画面に出さずに処理する（Excel・Word。PowerPoint はアプリを隠せないため、`Open` の `WithWindow` = False でウィンドウ無しで開く）。例外は、利用者のファイルが入った Excel・Word を利用者に渡すときだけ `$true` にする（`restoreHandedOverApp`。`handOverApp` と `retryKeptApps` が呼ぶ） |
| `DisplayAlerts` | 無効 | ダイアログで処理が止まらないようにする |
| `Open` の `ReadOnly` | 真 | 読み取り専用で開く（`tebunko/indexer/extract_office.ps1:122`（Excel）・`202`（Word）・`228`（PowerPoint）） |

`Workbooks.Open` / `Workbooks.Add` / `Documents.Open` / `Presentations.Open` を呼ぶファイルは、`AutomationSecurity` を `2` にしていることを `tests/meta/safety.Tests.ps1` が確かめる（インデクサの `office_app.ps1` と `extract_office.ps1` は上の `3`）。もらったインデックスのマクロを持てる形式の開き方は [もらったインデックスのマクロを持てる形式](../design/gui/open-file.md#もらったインデックスのマクロを持てる形式) にある。

パスワード付きファイルは、開くときに固定文字列 `"dummy"` をパスワードとして渡す（同じ 3 行）。これはパスワードを破るための処理ではなく、**パスワード入力ダイアログを出さずに確実に失敗させる**ための指定である。パスワード付きファイルは取り込まれず、`work\store\ingest_status.tsv` に失敗として記録される。本ツールがパスワードを入力・保存・送信することはない。

Excel は、セルの値をテキストに書き出すために、すべてのブックを上の設定で開く。Word・PowerPoint で開くのは、ZIP 形式でない旧形式（`.doc` / `.ppt`）などを新形式に変換するときだけである。新形式（`.xlsx` / `.docx` / `.pptx`）の図形・コメント・本文・SmartArt・グラフの文字は、Office を使わずに ZIP の中の XML を直接読む（`shared/office/office_reader.ps1`）。いずれも作業フォルダのコピーを読む（[取り込み対象のファイルは書き換えない](file-access.md#取り込み対象のファイルは書き換えない)）。

### 細工した Office ファイルから身を守ること

ZIP（`.xlsx` / `.docx` / `.pptx`）を直接読む処理は、細工したファイルで時間・メモリを使い切らせる攻撃（ZIP 爆弾・DTD の実体参照を入れ子にして膨張させる攻撃 等）を想定し、次の 2 つを行う。`tests/meta/safety.Tests.ps1` の「細工したOfficeファイル（.docx・.pptx・.xlsx）から身を守ること」がこれを確かめる。

- **ZIP の部品のサイズの上限**: ZIP 内の 1 部品を読む `readZipEntry`（`shared/office/office_reader.ps1`）が ZIP の中身を読み、展開後の大きさに**部品 1 つにつき 100MB**・**1 ファイルで読む合計 300MB** の上限を設ける。超えたら、部品名・大きさ・部品ごとか合計かの区別を出さない簡潔な文言（`ファイルサイズが大きすぎるため更新できません。`。画面・取り込み一覧向け。区別を出すと悪用のヒントになるため）の `ZipSizeLimitException` で例外にする。原因を調べられるよう、部品名・大きさ・部品ごとか合計かは `work/indexing_log.txt` にだけ書く（`writeZipSizeLimitLog`。`tebunko/indexer/indexer_state.ps1`）。ZIP のヘッダー（中央ディレクトリ）が実際より小さい大きさを偽って書いてあるものも、読む量を `Length + 1` バイトまでに絞り、`Length` を超えて読めたら例外にすることで検知する（大きな配列は作らず、小さな入れ物を使い回して読む）。**例外が 1 つある**: シートのヘッダー・フッターを読む `readXlsxSheetHeaderFooter` は、シートの部品が大きい（`sheetData`）ため、文字列にせず流れ（`XmlReader`）で読む。`readZipEntry` は通らないので、代わりに別の上限（1000MB。`$script:zipSheetStreamMaxBytes`。流れ読みはメモリを使わず、かかるのは時間だけなので、部品の上限より大きくしてある）を、**1 ファイルの流れ読みが実際に読んだ合計**（`$script:zipSheetStreamReadBytes`。`readXlsxObjectUnits` が呼ぶたびに 0 から数え直す）に設ける。展開後の大きさの申告（中央ディレクトリ）が残り（上限 − 既に読んだ合計）を超える部品は開かない。申告を偽ったもの（実際は大きいのに小さく申告）は、`XmlReader` に渡す流れを、読んだバイト数を数えて残りを超えたら止める包み（`LimitedReadStream`）にして、読む途中に打ち切る。読み終えたら読んだ量を合計に足す。`workbook.xml` の複数の `<sheet>` が同じシートの部品を指すときは、流れ読みを最初の 1 回だけにして結果を使い回す（同じ部品を何度も読ませて長引かせられないようにする）。この合計は `readZipEntry` の合計（300MB）には数えない。超えたときは、そのシートのヘッダー・フッターだけを読めなかった扱いにし（セルの値・図形・コメントは出す）、理由（流れ読みの上限）をインデックス作成のログに書く。検索結果を .xlsx に書き出す `shared/office/xlsx_writer.ps1` の `entry.Open()` は書き込み専用（読まない）なので、この検査の対象から外している。この例外が増えないよう、`entry.Open()`・`ReadToEnd`・`ExtractToFile`・`ExtractToDirectory`・`Expand-Archive`・`CopyTo` を使う場所は `tests/meta/safety.Tests.ps1` が機械的に確かめる。上限の値は `setting.config` に出しておらず、変えるには `shared/office/office_reader.ps1` の `$script:zipPartMaxBytes`・`$script:zipTotalMaxBytes`・`$script:zipSheetStreamMaxBytes` を直す（[既知の問題](../design/indexing/known-issues.md#共通)）。
- **埋め込んだファイルの読み取り**: Word・PowerPoint に埋め込んだ Office のファイルは、`readZipEntryBytes`（`office_reader.ps1`）が中身を読む。申告の大きさ（`Length`）の分だけの配列を、部品の上限（100MB）と合計の上限（300MB）を確かめたあとに作り、`Length` を超えて読めたら例外にする。埋め込んだファイルを開いた先の部品（Excel のシートを含む）も、同じ `readZipEntry`・`readZipEntryBytes` で読み、部品ごと・合計の上限に数える（埋め込みのシートには、親の Excel のシートのような流れ読みの緩い上限は使わない）。シートの XML を読む `MaxCharactersInDocument` は部品の上限（100MB）にする。さらに、共有文字列を多数のセルが参照すると小さな埋め込みから巨大な行ができるため、埋め込みから出す行の文字数を、親の 1 ファイルの全埋め込みで合計して数え（`addEmbeddedOutputChars`）、上限（`embeddedOutputMaxChars`、32M 文字）を超えたらその埋め込みを読まない（`ZipSizeLimitException` の `LimitKind` が `Output`）。さらに中に埋め込んだものは読まない（深さは 1 段まで）。
- **DTD（`<!DOCTYPE>`）の処理を禁止**: `XmlDocument` を作るところ（`LoadXml` を直接呼ばず、すべて `newXmlDocument` を通す）は `XmlReaderSettings.DtdProcessing` を `Prohibit` にし、DTD 宣言を含む XML を例外にする。

いずれも、見つからなかった部品があるときと同じ扱い（ファイル全体を失敗にするか、1 つの図形・グラフ・SmartArt だけを空にして続けるか）は、読んだ関数の呼び出し元の決まりに従う（[Word・PowerPoint のテキスト読み取り](../design/indexing/office-apps.md#wordpowerpoint-のテキスト読み取りscriptssharedofficeoffice_readerps1)・[Excel の図形・コメントの読み取り](../design/indexing/excel.md#excel-の図形コメントの読み取りreadxlsxobjectunits)）。

## 取り込み対象の拡張子

Office に加えてテキストファイルと PDF ファイルも取り込むが、対象の拡張子は `shared/office/office_file_kind_view.ps1` の `$officeExtensions`（10 個）と `shared/core/text_file.ps1` の `$textExtensions`（75 個。既定のアプリで開く `$textOpenExtensions` 60 個 + メモ帳で開く `$textNotepadExtensions` 15 個）と `shared/core/pdf_file.ps1` の `$pdfExtensions`（`.pdf` の 1 個）を合わせた、**固定の 86 個の一覧**である（[テキストファイルの読み取り](../design/indexing/text.md)・[PDF ファイルの読み取り](../design/indexing/pdf.md)）。設定ファイルで拡張子を足す仕組みは無い（設定で選べるのは、この固定の 75 個の中で、取り込まない・メモ帳で開く、の 2 方向だけ。設定を読む `readTextExtensionSettings` は 75 個に無いものを捨て、実行・登録になる `$textNotepadExtensions` は設定にかかわらずメモ帳で開く。`tests/meta/safety.Tests.ps1` が確かめる）。

危険な種類のファイル（実行される・既定のアプリがスクリプトを実行する `.exe` `.com` `.scr` `.hta` `.wsf` `.wsh` `.jse` `.vbe` `.lnk` `.url` `.msi` `.cpl` `.inf` 等）を、実行される種類を並べて弾く形にすると、一覧から漏れて対象に紛れ込む恐れがある。そのため `tests/meta/safety.Tests.ps1` は「対象の拡張子が固定であること」として、`$textExtensions` がちょうど 75 個（内 `$textNotepadExtensions` が 15 個、`$textOpenExtensions` が 60 個）であること、`$pdfExtensions` が `.pdf` だけであること、`$officeExtensions` と `$textExtensions` と `$pdfExtensions` を合わせた一覧がこの 86 個ちょうどであること、上に挙げた危険な種類の拡張子がどちらの一覧にも含まれないことを確かめる。対象の拡張子を足す変更は必ずこのテストを落とし、この一覧の見直しを促す。

`$textNotepadExtensions`（`.bat` `.cmd` `.ps1` `.vbs` `.js` `.reg` `.sh` `.py` `.rb` `.pl` `.php` `.lua` `.mjs` `.psm1` `.psd1`）は、既定のアプリで開くと実行・登録になる（スクリプト・解釈器で動かす言語を含む）ため取り込みの対象ではあるが、検索結果から開くときは既定のアプリではなくエディタ（既定はメモ帳）で開く（下の「外部プロセスの起動」）。取り込むかどうかはインデックスごとに、開き方は全体で選べるが、選べるのは固定の 75 個の中だけで、実行・登録になる 15 個は設定にかかわらずエディタで開く。`.hs` `.kts` `.exs` は、Windows に既定の関連付けが無く、実行される既定の登録が一般的でないため既定のアプリのままとする。

開く前には、インデックスの「ファイル名」を `testSourceBookName` で確かめる（取り込み対象の拡張子であること、フォルダの区切り・`..`・`:`・制御文字・表示の向きを変える文字を含まないこと。末尾のピリオド・空白は Windows が無視するので外して見る）。書き換えたインデックスに `evil.hta` などがあっても、検索結果に出さず、開かない。

### PDF ファイルを読むときの上限

PDF は PowerShell だけで読み、外の部品・Office・外部プロセスを使わない。細工した PDF で固まる・メモリを使い切ることを防ぐため、`shared/core/pdf_file.ps1` は次の上限を持ち、超えたら取り込みの失敗にする（`tests/shared/core/pdf_file.Tests.ps1` が、上限を小さく差し替えて確かめる）。

| 上限 | 値 |
|---|---|
| 元のファイルの大きさ（コピーの前に、メタデータで確かめる） | 100MB |
| 取り出したストリームの大きさの合計（取り出す前の大きさと、圧縮を解いた後の大きさ。圧縮の有無にかかわらず、同じストリームを何度参照されても、そのたびに数える。圧縮爆弾の対策） | 100MB |
| オブジェクトの数 | 1,000,000 |
| 入れ子の深さ（配列・辞書・ページツリー、読み込み中の間接参照の重なり） | 64 |
| 文字の対応表（`ToUnicode`）の項目の合計（同じ表を使うフォントは 1 度だけ数える） | 1,000,000 |
| たどる相互参照（`Prev`・`XRefStm`）の数 | 1,000 |
| 配列の要素・辞書の項目の数（1 つあたり）、読んだ相互参照の行の総数 | 1,000,000（オブジェクトの数と同じ） |
| ストリームの終わりを探して読み進めた文字数の合計（`/Length` が合わないストリームの細工への対策） | ファイルの大きさの 4 倍 |
| ページ数 | 10,000 |

参照の循環（`Prev` のつながりを含む）とファイルの外を指す位置は例外にする。`Prev` は呼び出しの入れ子にせず、順に読む。暗号化された PDF は、空の利用者パスワードで開ける AES-256（V5 の R5・R6）だけ復号して読む。パスワードが要るもの・未対応の方式（RC4 40・128、AES-128 など）は、そのファイルだけ失敗にする（パスワードを探さない）。壊れた暗号化の辞書（短い `/U` `/UE`）は決まった長さの確かめで断り、AES-256 R6 の鍵の計算の繰り返しは 300 回で止める。鍵・途中の値は例外の文にも記録にも出さない。PDF の中のスクリプト・リンク・添付ファイル・外部の参照は、実行も取得もしない（文字の命令だけを読む）。

`.xml` は、既定のアプリ（Windows の「XML エディター」）が先頭の `<?mso-application progid=…?>` を見て Excel・Word で開くことがあるが、マクロは既定で無効（上の「Office ファイルを開くときの設定」）であるため、開くこと自体の害は小さい。

## 外部プロセスの起動

`tests/meta/safety.Tests.ps1` の「外部プロセスの起動は explorer.exe・固定のメモ帳（エラー記録）と、openWithEditor が … 選んだエディタ … だけ」が確かめる。

| 関数 | 使う API | 用途 |
|---|---|---|
| `applyFailedFileState`（`tebunko/ui/index/index_detail.ps1`。フォルダを開く 1 か所。ファイルを選んで見せるときは `showInExplorer` を使う）・`applyIndexSourceFolderState`（`tebunko/ui/index/index_list.ps1`）・`showInExplorer`（`tebunko/ui/open_source.ps1`。元のファイルの場所と、出力した検索結果のファイルを、エクスプローラーで選んで見せる） | `Start-Process`（`explorer.exe`） | 一覧・検索結果から元のファイルの場所を開く（利用者の操作時のみ）。フォルダと確かめたものだけを渡す（ファイルをフォルダとして渡すと、そのファイルが既定のアプリで実行されるため） |
| `openAboutLink`（`tebunko/ui/about_dialog.ps1`） | `Start-Process`（`explorer.exe`。アドレスを 1 つ渡す） | 「バージョン情報」の GitHub・リリースのリンク。開くのは `getAboutLinks` に載せた 2 つのアドレス（`https://github.com/hsgwa/tebunko` と、その `/releases`）と一字一句同じものだけ（`testAboutLinkUrl`）。利用者がリンクを押したときのみ |
| `openWithEditor`（`tebunko/ui/open_source.ps1`） | `Start-Process`（`getTextEditorChoice` が選んだプログラム。既定は固定のパス `%SystemRoot%\System32\notepad.exe`） | 検索結果から、既定のアプリで開くと実行・登録になる拡張子（`$textNotepadExtensions`）と、設定でエディタにした拡張子のテキストファイルを開く（利用者の操作時のみ）。エディタを差し替えられるのは、利用者が自分の設定ファイル `setting.config` の `textEditorPath` に書いたときだけ（下の「エディタの差し替え」） |
| `openWithShell`（`tebunko/ui/open_source.ps1`） | `ProcessStartInfo`・`UseShellExecute`・`Process::Start`・`Invoke-Item` | 元のファイルを既定のアプリで開く。先頭で `testSourceBookName` により名前を確かめ、通らなければ何も起動しない |
| `showErrorDialog`（`shared/ui/shell.ps1`） | `Start-Process`（固定のパスの `notepad.exe`） | エラーの記録のファイルをメモ帳で開く（利用者が押したときのみ） |

### エディタの差し替え（`textEditorPath`）

「エディタで開く」のエディタは、既定は固定のパスのメモ帳。利用者が自分の `setting.config` に `textEditorPath` を書くと、そのプログラムで開く。守りは次のとおり（`tests/meta/safety.Tests.ps1` の「外部プロセスの起動は …」「openWithEditor は …」と、`text_open_view.Tests.ps1` の `testTextEditorPath` の表が確かめる）。

- **指定できるのは利用者本人だけ**: 値は利用者の `setting.config`（ワークスペースの外。ツールのフォルダか既定のワークスペース）にだけある。インデックス・取り込み一覧・共有のワークスペースからは読まない。この版では設定の画面から変えられない。
- **パスの検証（`testTextEditorPath`）**: ドライブ文字から始まる絶対パスで、拡張子が `.exe`。UNC・相対パス・環境変数（`%`）・引用符・引数付き・`..`・`[` `]` `*` `?`（`Start-Process` がワイルドカードとして扱う）・ファイル名に使えない文字は不可。8.3 の短い名前（`~` を含む。`POWERS~1.EXE` のように表に当たらずに別のプログラムを指せる）と、ドライブ文字の後ろの `:`（別のデータストリーム `a.exe:powershell.exe`）も不可。通らなければ、既定のアプリには戻さずメモ帳で開き、`設定のエディタ（textEditorPath）が使えないため、メモ帳で開きます。` と知らせる。
- **スクリプトや命令を動かすプログラムは拒む（`testTextEditorDenied`）**: 実行・登録になる拡張子（`$textNotepadExtensions`）のファイルは、必ずエディタに渡される。エディタが `powershell.exe` や `python.exe` のようなプログラムだと、渡したスクリプトがそのまま動いてしまうため、ファイル名（大文字・小文字を区別せず、フォルダの付いた形・拡張子の有無も同じに見る）が表（`$textEditorDenyNames`）にあるものは、エディタにできない。表は `powershell`・`powershell_ise`・`pwsh`・`cmd`・`wscript`・`cscript`・`mshta`・`conhost`・`wt`・`bash`・`wsl`・`sh`・`python`・`pythonw`・`py`・`pyw`・`node`・`ruby`・`perl`・`php`・`lua`・`java`・`javaw`・`regedit`・`reg`・`msbuild`・`rundll32`・`regsvr32`・`msiexec`・`cmstp`・`installutil`・`regasm`・`regsvcs`・`forfiles`・`pcalua`・`explorer`。拒んだときは、既定のアプリには戻さずメモ帳で開き、`設定のエディタ（textEditorPath）は、スクリプトを動かすプログラムのため使えません。メモ帳で開きます。` と知らせる（保存済みの設定にあっても、開くたびに確かめるので同じ）。
- **この表の限界**: 表は漏れうる。ここに無い名前でも、別のプログラムでスクリプトを動かせるもの（名前を変えたコピー・ほかの言語の処理系・スクリプトを動かす機能のあるエディタなど）は防げない。エディタに何を指定するかは利用者本人が自分の設定ファイルに書くもので、この表は、よく知られたものを誤って指定したときの事故を防ぐ守りにとどまる。エディタがスクリプトを動かせるプログラムだと、実行・登録になる拡張子のファイルは実行され得る。
- **ファイルの有無**: `openWithEditor` が `Test-Path -LiteralPath -PathType Leaf` で確かめる。無ければメモ帳で開き、同じ知らせを出す。
- **起動の形**: `Start-Process -FilePath`（エディタ）`-ArgumentList`（引用した 1 つのパス）。シェル・関連付け（`openWithShell`・`Invoke-Item`）は通さない。実行・登録になる拡張子（`$textNotepadExtensions`）は、設定にかかわらず常にエディタで開き、既定のアプリには戻さない。
- **開く対象**: 開くファイルは、これまでどおり取り込み対象の名前（`testSourceBookName`）を通ったものだけ。


上の表にない関数が、シェルで開く API を書くと `tests/meta/safety.Tests.ps1` の「検索結果から開くとき、シェルで開く処理が決まった所にだけあること」が落ちる（別名・型名の書き方違いも構文木で見つける）。足すときは、表と守りを同じ変更で直す。

インデックス作成は画面のプロセスの中のスレッドで動かすため、インデックス作成のために `powershell.exe` を起動することはない（[プロセス](../design/structure/threads.md#プロセス)）。

元のファイルを既定のアプリで開く操作（`tebunko/ui/open_source.ps1` の `openWithShell`）では、`Start-Process` ではなく `ProcessStartInfo` を使う（`Start-Process` は `[` `]` を含むパスをワイルドカードとして解釈するため）。起動対象は、利用者が選んだ行のファイルと、それに関連付けられたアプリケーションである。検索結果の出力を開く操作（`openExportedFile`）は、このツールが今書いた .xlsx・.csv・.tsv だけを開く。.xlsx は `Invoke-Item -LiteralPath`（既定のアプリ）、.csv・.tsv は固定のパスのメモ帳（`openWithEditor` の `-notepadOnly`。設定のエディタ `textEditorPath` は使わない。値を加工しないので、表計算アプリが数式として解釈する開き方を避ける）。`ProcessStartInfo`・`[System.Diagnostics.Process]::Start` を使うのは `openWithShell` の 1 か所だけであることを `tests/meta/safety.Tests.ps1` が確かめる。

`tebunko.bat` は、上の起動とは別に、**起動そのものに失敗したとき（画面が開く前）だけ** `notepad.exe` を開く（[起動に失敗したときの知らせ](disclosure.md#起動に失敗したときの知らせtebunkobat)）。制限言語モード・実行ポリシーでスクリプトの読み込みが止まる場面では WPF・WinForms のメッセージボックスを出せないため、失敗の理由と記録のファイルの名前をメモ帳で示す。`tests/meta/safety.Tests.ps1` の「tebunko.bat が起動する外部のプログラムは conhost.exe・powershell.exe・notepad.exe だけ」が、`tebunko.bat` が起動する外部のプログラムをこの 3 つに限る。

起動するプロセスは、この表のほかに次の 2 つがある。どちらも利用者の操作や取り込みに伴うものである。

- **Office（`EXCEL.EXE`・`WINWORD.EXE`・`POWERPNT.EXE`）**: 旧形式（`.xls` `.doc` `.ppt`）やパスワード付きのファイルを COM で開くときだけ起動する。画面のプロセスが COM で頼み、Windows（DCOM）が起動するため、プロセスの親は `powershell.exe` ではなく `svchost.exe`（DcomLaunch）になる（[Office ファイルを開くときの設定](#office-ファイルを開くときの設定)）。`Start-Process` は使わない。
- **`conhost.exe` と `powershell.exe`**: `tebunko.bat` が 1 回だけ起動する（[起動口](disclosure.md#起動に失敗したときの知らせtebunkobat)）。`-Command` に渡す文字列は `tebunko.bat` の中に固定で書いた内容だけで、利用者の入力や外からの値を混ぜない（ツールのフォルダだけを環境変数で渡し、コードとして解釈しない）。`-EncodedCommand`・`-ExecutionPolicy Bypass`・`Invoke-Expression` の形は使わない。

## ウイルス対策ソフト・EDR での確かめ

非表示の PowerShell（`-WindowStyle Hidden`）での起動と、上の外部プロセスが、ウイルス対策ソフトや EDR に検知されないかを、実機で確かめた。

- **環境**: Windows 11、Microsoft Defender ウイルス対策（既定の設定。リアルタイム保護・振る舞い監視・AMSI が有効）。
- **Microsoft Defender for Endpoint（EDR 本体）は未確認。** 同じエンジンと AMSI を使うため手がかりにはなるが、MDE の検知の規則は別である。MDE の機械で確かめるときは、下の場面を同じ手順で流し、「前後で増えない」を比べる。
- **見たもの**: `Get-MpThreatDetection` の件数、イベントログ `Microsoft-Windows-Windows Defender/Operational` の 1006〜1008・1015・1116〜1119、`Microsoft-Windows-PowerShell/Operational` の 4104（警告レベル）。

- **前後の数え方**: 場面ごとの前後で、次を管理者でない PowerShell で流して件数を比べる。MDE の機械では、これに加えて MDE のポータルのアラートも見る。第二段（Azure の MDE で流す）は別の作業で行う。

```
# 数える前に、ログが読めることを確かめる（読めないと下の 2 本は黙って 0 になる）。エラーが出たら管理者の PowerShell で流す
(Get-WinEvent -ListLog 'Microsoft-Windows-Windows Defender/Operational','Microsoft-Windows-PowerShell/Operational').RecordCount
@(Get-MpThreatDetection).Count
@(Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-Windows Defender/Operational'; Id=1006,1007,1008,1015,1116,1117,1118,1119} -ErrorAction SilentlyContinue).Count
@(Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-PowerShell/Operational'; Id=4104; Level=3} -ErrorAction SilentlyContinue).Count
```

| 場面 | 操作 | 検知の件数（前 → 後） | Defender のイベント（前 → 後） |
|---|---|---|---|
| 1 | Mark-of-the-Web を付けた展開済みのフォルダの `tebunko.bat` から起動（画面が開くまで 15 秒待つ） | 0 → 0 | 0 → 0 |
| 2 | インデックス作成（`.docx` `.pptx` `.xlsx` のほか、旧形式の `.xls` `.doc` `.ppt` を COM で 5 ファイル。失敗 0 件） | 0 → 0 | 0 → 0 |
| 3 | 結果からフォルダを開く（`explorer.exe`）、`.ps1` をメモ帳で開く（`notepad.exe`） | 0 → 0 | 0 → 0 |
| 4 | 単一 .ps1 版の起動とインデックス作成（`[scriptblock]::Create`・`[char]0xNNNN` の組み立てを含む） | 0 → 0 | 0 → 0 |
| 5 | ［参照…］のフォルダ選択（`NonPublic` のリフレクション）を開く、キャンセルする、フォルダを選ぶ | 0 → 0 | 0 → 0 |

- **4104（警告レベル）**: 場面 1 では、起動のたびに画面のプロセスが読み込んだスクリプト 28 件が記録された。画面を開くテストの一式（場面 2・5 を含む）でも、`scripts` のスクリプトの読み込みで 21 ファイル分が記録された。場面 4（単一 .ps1 版）は、起動から 20 秒で 350 件が警告レベルで記録された。うち 1 件が末尾の `[scriptblock]::Create` の行（2,277 文字）で、残りは本体のスクリプトの塊の記録である。Windows PowerShell 5.1 が `Add-Type`・`BindingFlags`・`NonPublic` などの語や、スクリプトブロックの生成を含むスクリプトを自動で記録するもので、検知や隔離ではない。`-EncodedCommand` の形の記録は出なかった。MDE の機械では、この記録を EDR が取り込んで警告にするかを比べる。
- **結論**: 今の起動のしかたは変えない。検知は出なかった。

### 組織が許可するとき

EDR で検知されたときは、次の順で許可を求める。

1. **発行元（署名）で許可する。** 配布物にコード署名をして、その発行元（証明書）で許可するのが最も確実である（コード署名は今後対応する予定）。ファイルのハッシュでの除外は、版が変わるたびに作り直しになるため勧めない。
2. **署名が使えないときだけ、フォルダで許可する。** `tebunko.bat` と `scripts` のあるフォルダ（展開先）だけを、スクリプトの実行を許すフォルダにする。利用者が書き換えられない場所（管理者だけが書き込める場所）に置いたときに限る。利用者が書き込めるフォルダをパスで許可すると、そこに置いた別のスクリプトも通してしまう。ユーザー全体や一時フォルダは許可しない。
3. **許可の対象は 1 つの起動の形だけに絞る。** `conhost.exe` から `powershell.exe -NoProfile -STA -ExecutionPolicy RemoteSigned -WindowStyle Hidden -Command …` を起動する形（`tebunko.bat` の 1 行）である。`-ExecutionPolicy Bypass` など、ほかの形の許可は要らない。