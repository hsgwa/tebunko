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
| 1 | 文字列を式として実行（`Invoke-Expression` / `iex` / `[ScriptBlock]::Create`） | 該当 0 件 | `scan 'Invoke-Expression','[^-\w]iex[ (]','ScriptBlock\]::Create'` |
| 2 | 難読化されたコマンド（Base64・`-EncodedCommand`） | 該当 0 件 | `scan 'FromBase64String','EncodedCommand'` |
| 3 | ネットワーク通信 | 該当 0 件 | `scan 'Invoke-WebRequest','Invoke-RestMethod','WebClient','HttpClient','Net\.Sockets','Start-BitsTransfer','DownloadFile','DownloadString','System\.Net\.'` |
| 4 | Windows API の直接呼び出し（P/Invoke） | 該当 0 件 | `scan 'DllImport','GetDelegateForFunctionPointer'` |
| 5 | C# ソースの実行時コンパイル | 該当 0 件 | `scan 'Add-Type'` で出る行がすべて `-AssemblyName` であること（下表の 3 行） |
| 6 | レジストリの読み書き | 該当 0 件 | `scan 'HKLM','HKCU','HKEY_','Set-ItemProperty','New-ItemProperty','Registry::'` |
| 7 | 権限・サービス・自動起動の変更 | 該当 0 件 | `scan 'Set-Acl','icacls','schtasks','New-Service','Start-Service','\bsc\.exe','RunAs'` |
| 8 | 実行ポリシーの恒久変更・`Bypass` | 該当 0 件 | `scan 'Set-ExecutionPolicy','ExecutionPolicy\s+Bypass'` と `Select-String -Path .\tebunko.bat -Pattern 'Bypass','Set-ExecutionPolicy'` |
| 9 | 資格情報の入力要求・保存 | 該当 0 件 | `scan 'Get-Credential','ConvertTo-SecureString','PSCredential'` |
| 10 | リモート実行 | 該当 0 件 | `scan 'Invoke-Command','New-PSSession','Enter-PSSession','WinRM'` |
| 11 | 外部プロセスの起動 | 3 か所のみ（[外部プロセスの起動（3 か所）](#外部プロセスの起動3-か所)） | `scan 'Start-Process'` |
| 12 | プロセスの強制終了 | 1 か所のみ（[Office プロセスの強制終了（［9 プロセス停止］タブ）](disclosure.md#office-プロセスの強制終了9-プロセス停止タブ)） | `scan 'Stop-Process'` |
| 13 | 壊れたハッシュ（MD5・SHA-1）、FIPS 準拠でないハッシュの実装 | 該当 0 件 | `scan 'MD5','SHA1','RIPEMD','SHA256Managed','SHA384Managed','SHA512Managed','HashAlgorithm\]::Create'` |
| 14 | 内部の型（`NonPublic`）のリフレクションでの呼び出し | 1 か所のみ（下） | `scan 'NonPublic','Reflection\.BindingFlags'` |

クロール対象が共有フォルダにある場合のファイルの読み取りは、Windows のファイル共有（SMB）を通る。これは OS のファイル API によるもので、本ツールが通信処理を持つわけではない。

`Add-Type` は 3 か所あるが、いずれも Windows 標準アセンブリの読み込み（`-AssemblyName`）であり、コードのコンパイルではない。

| 場所 | 読み込むアセンブリ | 用途 |
|---|---|---|
| `tebunko/gui.ps1:8` | `PresentationFramework` / `PresentationCore` / `WindowsBase` / `System.Windows.Forms` | 画面（WPF）の表示、フォルダ選択（WinForms） |
| `shared/office/office_reader.ps1:12-13` | `System.IO.Compression` / `System.IO.Compression.FileSystem` | `.xlsx` / `.docx` / `.pptx`（ZIP）の読み取り |

実行時コンパイル（`csc.exe` の起動）を使わないのは設計方針である。検索処理・画面の型定義などは C# を使わず PowerShell と .NET の直接呼び出しで書いている（[画面の実装](../design/gui/implementation.md#実行時コンパイルcscexeを使わない)）。実行時に一時フォルダへコンパイル結果を書き出す動作が無いため、EDR やアプリケーション制御（AppLocker・WDAC）との衝突も避けられる。`tests/meta/safety.Tests.ps1` の「実行時にコードをコンパイルしない」が、`Add-Type` は標準アセンブリの読み込み（`-AssemblyName`）だけであることを確かめる。

**内部の型のリフレクション（14）**：フォルダ選択（`shared/ui/folder_dialog.ps1` の `showFolderPicker`）だけが、WinForms が内部に持つ `IFileOpenDialog` の定義（`FileDialogNative+IFileDialog`）をリフレクション（`BindingFlags` の `NonPublic`）で呼ぶ。Windows 標準のエクスプローラー形式のダイアログを、実行時コンパイルなしで開くためである（[フォルダ選択ダイアログ（［参照…］）](../design/gui/common.md#フォルダ選択ダイアログ参照)）。呼ぶのは .NET Framework に署名付きで入っている WinForms のダイアログの処理だけで、メモリの書き換えや AMSI などの検査の仕組みには触れない。開いている間に `csc.exe` などの子プロセスと一時フォルダへの DLL の書き出しが無いことを実機で確かめた。なお、Windows PowerShell 5.1 は `Add-Type`・`BindingFlags`・`NonPublic`・`GetMethod`・`IntPtr`・`InteropServices` などの語を含むスクリプトを、ログの設定に関わらずイベントログ（Microsoft-Windows-PowerShell/Operational の 4104、警告）に自動で記録する。本ツールはこの変更の前から `Add-Type` などで記録の対象である。`tests/meta/safety.Tests.ps1` が、`NonPublic` を使うのがこの 1 か所だけであることを確かめる。

## Office ファイルを開くときの設定

インデックス作成のために Excel・Word・PowerPoint を COM で操作するが、アプリを起動したときに次の設定を行う（`shared/office/office_app.ps1:79-96`）。`tests/meta/safety.Tests.ps1` の「Office ファイルを安全に開くこと」がこれを確かめる。**利用者の PowerPoint には接続しない**（1 つのセッションに 1 つのプロセスしか持てず、接続すると下の設定を利用者のアプリまで書き換えてしまうため）。起動する前に自分のセッションに `POWERPNT` が無いかを確かめ、あれば設定を行わずに例外にする（[Office アプリ（Excel・Word・PowerPoint）の管理](../design/indexing/office-apps.md#office-アプリexcelwordpowerpointの管理)「PowerPoint」）。

| 設定 | 値 | 意味 |
|---|---|---|
| `AutomationSecurity` | `3`（`msoAutomationSecurityForceDisable`） | **マクロを強制的に無効にする**。マクロ有効ファイル（`.xlsm` / `.docm` / `.pptm`）でもマクロは実行されない（Excel・Word・PowerPoint とも） |
| `EnableEvents` | `$false` | `Workbook_Open` などのイベントマクロを発火させない（Excel） |
| `AskToUpdateLinks` | `$false`、`Workbooks.Open` の `UpdateLinks` = `0` | **外部リンクを更新しない**。他ブックや外部データソースへのアクセスが発生しない（Excel） |
| `Visible` | `$false` | 画面に出さずに処理する（Excel・Word。PowerPoint はアプリを隠せないため、`Open` の `WithWindow` = False でウィンドウ無しで開く） |
| `DisplayAlerts` | 無効 | ダイアログで処理が止まらないようにする |
| `Open` の `ReadOnly` | 真 | 読み取り専用で開く（`tebunko/indexer/extract_office.ps1:122`（Excel）・`202`（Word）・`228`（PowerPoint）） |

パスワード付きファイルは、開くときに固定文字列 `"dummy"` をパスワードとして渡す（同じ 3 行）。これはパスワードを破るための処理ではなく、**パスワード入力ダイアログを出さずに確実に失敗させる**ための指定である。パスワード付きファイルは取り込まれず、`work\取り込み一覧.tsv` に失敗として記録される。本ツールがパスワードを入力・保存・送信することはない。

Excel は、セルの値をテキストに書き出すために、すべてのブックを上の設定で開く。Word・PowerPoint で開くのは、ZIP 形式でない旧形式（`.doc` / `.ppt`）などを新形式に変換するときだけである。新形式（`.xlsx` / `.docx` / `.pptx`）の図形・コメント・本文・SmartArt・グラフの文字は、Office を使わずに ZIP の中の XML を直接読む（`shared/office/office_reader.ps1`）。いずれも作業フォルダのコピーを読む（[取り込み対象のファイルは書き換えない](file-access.md#取り込み対象のファイルは書き換えない)）。

### 細工した Office ファイルから身を守ること

ZIP（`.xlsx` / `.docx` / `.pptx`）を直接読む処理は、細工したファイルで時間・メモリを使い切らせる攻撃（ZIP 爆弾・DTD の実体参照を入れ子にして膨張させる攻撃 等）を想定し、次の 2 つを行う。`tests/meta/safety.Tests.ps1` の「細工したOfficeファイル（.docx・.pptx・.xlsx）から身を守ること」がこれを確かめる。

- **ZIP の部品のサイズの上限**: ZIP 内の 1 部品を読む `readZipEntry`（`shared/office/office_reader.ps1`）だけが ZIP の中身を読み、展開後の大きさに**部品 1 つにつき 100MB**・**1 ファイルで読む合計 300MB** の上限を設ける。超えたら、部品名・大きさ・部品ごとか合計かの区別を出さない簡潔な文言（`ファイルサイズが大きすぎるため取り込めません。`）で例外にする。ZIP のヘッダー（中央ディレクトリ）が実際より小さい大きさを偽って書いてあるものも、`Length + 1` バイトの入れ物を作って読み切ることで検知する。上限の値は `setting.config` に出しておらず、変えるには `shared/office/office_reader.ps1` の `$script:zipPartMaxBytes`・`$script:zipTotalMaxBytes` を直す（[既知の問題](../design/indexing/known-issues.md#共通)）。
- **DTD（`<!DOCTYPE>`）の処理を禁止**: `XmlDocument` を作るところ（`LoadXml` を直接呼ばず、すべて `newXmlDocument` を通す）は `XmlReaderSettings.DtdProcessing` を `Prohibit` にし、DTD 宣言を含む XML を例外にする。

いずれも、見つからなかった部品があるときと同じ扱い（ファイル全体を失敗にするか、1 つの図形・グラフ・SmartArt だけを空にして続けるか）は、読んだ関数の呼び出し元の決まりに従う（[Word・PowerPoint のテキスト読み取り](../design/indexing/office-apps.md#wordpowerpoint-のテキスト読み取りscriptssharedofficeoffice_readerps1)・[Excel の図形・コメントの読み取り](../design/indexing/excel.md#excel-の図形コメントの読み取りreadxlsxobjectunits)）。

## 取り込み対象の拡張子

Office に加えてテキストファイルも取り込むが、対象の拡張子は `shared/office/office_files.ps1` の `$officeExtensions`（10 個）と `shared/core/text_file.ps1` の `$textExtensions`（`.txt` `.csv` `.tsv` `.md` `.log` `.json` `.xml` の 7 個）を合わせた、**固定の 17 個の一覧**である（[テキストファイルの読み取り](../design/indexing/text.md)）。設定ファイルで拡張子を足す仕組みは無い。

危険な種類のファイル（実行される・既定のアプリがスクリプトを実行する `.bat` `.js` `.vbs` `.ps1` `.hta` 等）を、実行される種類を並べて弾く形にすると、一覧から漏れて対象に紛れ込む恐れがある。そのため `tests/meta/safety.Tests.ps1` は「対象の拡張子が固定であること」として、`$textExtensions` がちょうど 7 個であること、`$officeExtensions` と `$textExtensions` を合わせた一覧がこの 17 個ちょうどであることを確かめる。対象の拡張子を足す変更は必ずこのテストを落とし、この一覧の見直しを促す。

`.xml` は、既定のアプリ（Windows の「XML エディター」）が先頭の `<?mso-application progid=…?>` を見て Excel・Word で開くことがあるが、マクロは既定で無効（上の「Office ファイルを開くときの設定」）であるため、開くこと自体の害は小さい。

## 外部プロセスの起動（3 か所）

`tests/meta/safety.Tests.ps1` の「外部プロセスの起動は explorer.exe だけ」が確かめる。

| 場所 | 起動するもの | 用途 |
|---|---|---|
| `tebunko/ui/index_tab.ps1:618`・`623`、`tebunko/ui/open_source.ps1:311` | `explorer.exe` | 一覧・検索結果から元のファイルの場所を開く（利用者の操作時のみ） |

インデックス作成は画面のプロセスの中のスレッドで動かすため、インデックス作成のために `powershell.exe` を起動することはない（[プロセス](../design/structure/threads.md#プロセス)）。

元のファイルを開く操作（`tebunko/ui/open_source.ps1:125`）では、`Start-Process` ではなく `ProcessStartInfo` を使う（`Start-Process` は `[` `]` を含むパスをワイルドカードとして解釈するため）。起動対象は、利用者が選んだ行のファイルと、それに関連付けられたアプリケーションである。

`tebunko.bat` は、上の 3 か所とは別に、**起動そのものに失敗したとき（画面が開く前）だけ** `notepad.exe` を開く（[起動に失敗したときの知らせ](disclosure.md#起動に失敗したときの知らせtebunkobat)）。制限言語モード・実行ポリシーでスクリプトの読み込みが止まる場面では WPF・WinForms のメッセージボックスを出せないため、失敗の理由と記録のファイルの名前をメモ帳で示す。`tests/meta/safety.Tests.ps1` の「tebunko.bat が起動する外部のプログラムは conhost.exe・powershell.exe・notepad.exe だけ」が、`tebunko.bat` が起動する外部のプログラムをこの 3 つに限る。
