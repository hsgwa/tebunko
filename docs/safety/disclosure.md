# 開示事項

次の 7 点は審査で論点になるため、本ツールの側から先に開示する。[前回残った Office の終了（起動時の確認）](#前回残った-office-の終了起動時の確認)・[Mark-of-the-Web の解除](#mark-of-the-web-の解除tebunkobat)・[検索結果から開くと Excel が前面に出る](#検索結果から開くと-excel-が前面に出る) は利用者の操作に紐づく動作、[インデックスが元文書の本文を保持する（情報の集約）](#インデックスが元文書の本文を保持する情報の集約)・[原本の一時コピーと、その回収](#原本の一時コピーとその回収)・[透過暗号化の製品が平文で見せたファイルの扱い](#透過暗号化の製品が平文で見せたファイルの扱い) はツールの性質上避けられない情報の扱い、[インストーラー版](#インストーラー版) はインストーラー版だけに当てはまる動作である。

## 前回残った Office の終了（起動時の確認）

インデックス作成が異常終了すると、画面の無い Excel・Word・PowerPoint のプロセスが残り、メモリを占有し続けることがある。これを止めるために、次に画面を起動したとき、利用者に確認したうえで `Stop-Process -Force` を実行する（`shared/office/office_process.ps1` の `stopOfficeProcesses`。`safety.Tests.ps1` がこの 1 か所だけであることを確かめる）。

- **影響**: 強制終了するため、対象プロセスで**未保存の内容は失われる**。
- **制御**: 確認で利用者が［終了する］を選んだときだけ実行される。確認で出した PID のうち、実行の直前にも同じ条件を満たすものだけを止める。自動では実行されない。監視・常駐もしない。［今回は終了しない］を選べば何も止めず、次の起動でもう一度確認する。
- **対象の限定**: 対象は、インデックス作成が起動したときにワークスペースへ記録した PID だけである。記録のある Excel（`EXCEL`）・Word（`WINWORD`）・PowerPoint（`POWERPNT`）で、窓を持たず、起動した側がもう動いていないものに限る。利用者が自分で開いた Office には記録が無いため、確認にも出ず、止めない。
- **照らし直し**: 止める直前に、PID のプロセスの名前と起動時刻が記録と同じかをもう一度確かめる。違うとき（PID が別のプロセスに使われたとき）は止めず、そのことを知らせる。

詳細は [前回残った Office の確認](../design/gui/leftover-office.md) を参照。

## Mark-of-the-Web の解除（`tebunko.bat`）

`tebunko.bat` は PowerShell を 1 回起動し、その中で配布フォルダの `scripts` 配下の Mark-of-the-Web を解除してから、画面のスクリプトを実行する。

```bat
set "PS1=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS1%" (
    start "" notepad.exe "%~dp0scripts\tebunko\startup\no_powershell.txt"
    exit /b 1
)
rem Step 1: the notepad opener (kept as a single assignment so a test can
rem replace it) and the tool's folder paths, used by every later step.
set "PSCMD=$opener='notepad.exe';$root='%~dp0';$startup=Join-Path $root 'scripts\tebunko\startup';$gui=Join-Path $root 'scripts\tebunko\gui.ps1';"
rem Step 2: clear Mark-of-the-Web from the scripts folder (see the comment above).
set "PSCMD=%PSCMD%Get-ChildItem -LiteralPath (Join-Path $root 'scripts') -Recurse -File -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue; "
rem Step 3: run the GUI script, catching a failure from before its own window opens.
set "PSCMD=%PSCMD%try { & $gui } catch { $err = $_; "
rem Step 4: in the catch, pick the reason (missing gui.ps1, Constrained Language
rem Mode, execution policy, or something else) and its text file.
set "PSCMD=%PSCMD%if (-not (Test-Path -LiteralPath $gui)) { $reason = Join-Path $startup 'missing_files.txt' } elseif ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') { $reason = Join-Path $startup 'constrained_language.txt' } elseif ($err.FullyQualifiedErrorId -like 'UnauthorizedAccess*') { $reason = Join-Path $startup 'execution_policy.txt' } else { $reason = Join-Path $startup 'failed.txt' }; $reasonLines = @(if (Test-Path -LiteralPath $reason) { Get-Content -LiteralPath $reason -Encoding UTF8 } else { @('tebunko could not start.') }); "
rem Step 5: build the record's detail lines (date, tool path, language mode,
rem PowerShell version, execution policy, and the error message).
set "PSCMD=%PSCMD%$detailLines = @(('==== {0} startup ====' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')), ('Tool: {0}' -f $root), ('LanguageMode: {0}' -f $ExecutionContext.SessionState.LanguageMode), ('PSVersion: {0}' -f $PSVersionTable.PSVersion), (Get-ExecutionPolicy -List | Out-String), ('{0}' -f $err.Exception.Message)); $allLines = $reasonLines + '' + $detailLines; "
rem Step 6: write the record in the tool folder (no record if it cannot be written; then the reason is shown), then show it.
set "PSCMD=%PSCMD%$openTarget = $reason; $candidate = Join-Path $root 'startup_error.txt'; try { Set-Content -LiteralPath $candidate -Value $allLines -Encoding UTF8 -ErrorAction Stop; $openTarget = $candidate } catch { }; & $opener $openTarget }"
start "" conhost.exe "%PS1%" -NoProfile -STA -ExecutionPolicy RemoteSigned -WindowStyle Hidden -Command "%PSCMD%"
```

`-Command` の中身は 1 行に書かず、意味のまとまりごとに `set "PSCMD=..."` を積み重ねて組み立てる（`tebunko.bat` が長い 1 行になって読みにくくならないようにするため）。`delayed expansion` は使わず、各 `set` は前の行までの `%PSCMD%` を展開してから文字列を足すだけなので、素直な変数展開で足りる。`tests/meta/launcher.Tests.ps1` が、組み立てた結果が想定した文字列とちょうど一致することを確かめる。（`try`/`catch` の中身は、次の「起動に失敗したときの知らせ」で説明する。）

- **理由**: zip で配布したものを展開すると全ファイルに Mark-of-the-Web が付き、実行ポリシー `RemoteSigned` では未署名スクリプトが実行できないため。
- **範囲**: `%~dp0scripts`（`tebunko.bat` があるフォルダの `scripts`）配下のみ。実行するスクリプトはすべてここにある。インデックス（`work`）には触らない。システム全体やユーザーのドキュメントには影響しない。画面（`gui.ps1`）も起動時に同じ `scripts` 配下に対して同じ解除を行う（ショートカットから直接起動したときのため）。
- **実行ポリシー**: `Bypass` は使わず `RemoteSigned` で起動する。恒久的な設定変更（`Set-ExecutionPolicy`）も行わない。`-ExecutionPolicy` はこのプロセスだけに効く。理由は次の「`RemoteSigned` で起動する理由」を参照。
- **`conhost.exe` を通す理由**: 既定のターミナルが Windows Terminal の場合でも `-WindowStyle Hidden` で PowerShell の窓を隠すため（[画面の共通仕様](../design/gui/common.md#配布と実行ポリシーmark-of-the-web)）。
- **恒久的な解消策**: 配布物のカタログ（[配布物の完全性（カタログ・ハッシュ一覧・来歴の署名）](scans.md#配布物の完全性カタログハッシュ一覧来歴の署名)）にコードサイニング証明書で署名すれば、この `Unblock-File` は不要になり、`AllSigned` でも動作する。

## 起動に失敗したときの知らせ（`tebunko.bat`）

制限の強い環境（実行ポリシー・制限言語モード）ほど、`gui.ps1` の画面が開く前に起動が止まり、利用者から見て「ダブルクリックしても何も起きない」ことになりやすい。`tebunko.bat` は、`& $gui`（画面の起動）を `try`/`catch` で包み、失敗したときは理由と記録のファイルの名前を Notepad で示す。インラインの `-Command` は実行ポリシーの対象外で、制限言語モードでも動くため、この知らせの処理自体は止まらない。

- **PowerShell が無い**: `%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe`（PATH からは探さない。`installer\tebunko.cs` と同じ）が無ければ、PowerShell を起動せずに Notepad で理由を示す。この場面だけは記録が残らない（PowerShell 自体が使えないため）。
- **理由の見分け方**: `catch` は、`gui.ps1` が無ければファイル不足、その場の言語モード（`$ExecutionContext.SessionState.LanguageMode`）が `FullLanguage` でなければ制限言語モード、`FullyQualifiedErrorId` が `UnauthorizedAccess` なら実行ポリシー、それ以外は起動できなかったと分ける。`gui.ps1` の有無を先に見るのは、無いのに「フォルダの許可を情報システム部門に相談」と案内しないため。許可されていない制限言語モードでは `gui.ps1` の最初の読み込み（`paths.ps1` の `New-Object`）で失敗し、ツールのフォルダを AppLocker・WDAC で許可している PC では `gui.ps1` は `FullLanguage` で動くため、この `catch` には来ない（画面が開く）。
- **理由の文言**: 日本語の文言は ASCII で書く決まりの `tebunko.bat` に書けないため、`scripts\tebunko\startup\*.txt`（BOM 付き UTF-8・CRLF）に分けて置き、`Get-Content -Encoding UTF8` で読む。
- **記録**: 理由の文言と、詳しい情報（日時・エラーの内容・言語モード・`Get-ExecutionPolicy -List`・PowerShell の版・ツールの場所）を、ツールのフォルダの `startup_error.txt` に書く（`%LOCALAPPDATA%` や `%TEMP%` には書かない）。起動の前はワークスペースを決められない（設定を読むには `RemoteSigned` で止まる場面がある）ため、ツールのフォルダにする。ツールのフォルダに書けないとき（`C:\Program Files` に入れた場合など）は記録を残さず、理由の文言のファイルを Notepad で示すだけにする。ツールの場所（利用者名を含むパス）が記録に入るため、公開の Issue に貼る前に確かめるよう[トラブルシューティング](../guide/troubleshooting.md)に書いている。
- **`gui.ps1` の `trap` との関係**: `gui.ps1` の `trap` は、`writeErrorLog` がまだ使えない（`app_host.ps1` を読み込む前）ときに起きた失敗を、同じ場所に記録してからメッセージボックスを出そうとする。制限言語モード・WPF が読めない場面ではメッセージボックスも出せないため、`trap` は元の例外を投げ直し、上の `catch` が受けて記録のファイルを上書きする（1 か所に記録が 1 件残る）。メッセージボックスを出せたとき（`trap` が `exit 1` で終わったとき）は、`catch` には来ない（二重に知らせない）。インストーラー版 `tebunko.exe` から起動したときは、`trap` の記録に加えて、`tebunko.exe` がエラー出力（元の例外の内容）をメッセージで出す。
- **AppLocker の実行ファイルの規則で `powershell.exe` 自体が止められた場合**: `start` は起動の失敗を待たないため、`tebunko.bat` は気づけず、何も知らせない（Windows 自身の知らせと、イベントログ（AppLocker）だけが残る）。`start` を起動の失敗を待つ形にすると、画面を閉じるまで `tebunko.bat` の窓が残ってしまうため、この場面は扱わない。

### `RemoteSigned` で起動する理由

Windows のクライアント版（Windows 10 / 11）は、実行ポリシーの既定が `Restricted` で、`.ps1` を 1 つも実行できない。本ツールはスクリプトでできているため、何らかの実行ポリシーを指定しないと起動できない。そこで、起動口（zip 版の `tebunko.bat`・インストーラー版の `tebunko.exe`）は、`powershell.exe` の `-ExecutionPolicy RemoteSigned` で**そのプロセスだけ**の実行ポリシーを指定する。

- **設定を変えない**: `-ExecutionPolicy` の指定は、起動したプロセスとその子プロセスだけに効く。PC や利用者の実行ポリシーは書き換えず（`Set-ExecutionPolicy` を使わない）、管理者権限も要らない。本ツールを終えると元のとおりである。
- **`RemoteSigned` を選ぶ理由**: 署名の無い手元のスクリプトを実行できる実行ポリシーのうち、いちばん厳しいため。手元のスクリプトは実行し、インターネットから来た（Mark-of-the-Web が付いた）スクリプトは署名が無ければ止める。本ツールのスクリプトは、上の `Unblock-File` で自分のフォルダの `scripts` 配下だけ印を消して実行する。それ以外のダウンロードしたスクリプトは、本ツールのプロセスの中でも今までどおり止まる。
- **組織が実行ポリシーを決めている場合**: グループポリシーで実行ポリシーを決めている PC（`Get-ExecutionPolicy -List` の `MachinePolicy` か `UserPolicy` に値がある）では、そちらが優先され、`-ExecutionPolicy` の指定は効かない。グループポリシーが `AllSigned` なら、本ツールは署名するまで起動できない。これは組織の設定を守るための Windows の仕様であり、本ツールから回避はしない。

ほかの実行ポリシーを使わない理由は次のとおり。

| 実行ポリシー | 使わない理由 |
|---|---|
| `Bypass` | 何も止めず、警告も出さない。未署名のスクリプトを `Bypass` とウィンドウを隠す指定（`-WindowStyle Hidden`）で動かす形は、マルウェアがよく使う形と同じで、EDR やウイルス対策ソフトに怪しまれやすい |
| `Unrestricted` | インターネットから来た未署名のスクリプトも、確認を出したうえで実行できてしまう。確認はウィンドウを隠した起動では見えず、処理が止まる |
| `AllSigned` | すべてのスクリプトに署名が要る。本ツールはまだ署名していない（署名すれば `AllSigned` でも動く。上の「恒久的な解消策」） |

## インデックスが元文書の本文を保持する（情報の集約）

本ツールは全文検索のために、**元の文書の本文を平文の TSV として本文インデックス（インデックスのうち `work\content_index\` の TSV）に保持する**。これはツールの目的そのものであり、避けられない。審査では次の点が論点になる。

| 論点 | 内容 |
|---|---|
| アクセス権の非継承 | インデックスは**作成した人の権限で作られる**。元のファイルに設定されたアクセス権は引き継がれない。インデックスを置いた場所を読める人は、元のファイルを読む権限が無くても本文を読める |
| システムインデックス | 本文インデックスとは別に、高速検索のために `work\system_index\` の txt にも、本文から作った隣り合う 2 文字の組が入る。この txt は Windows Search に索引させる。本文インデックス（`content_index`）は、索引の対象から自動で外す。仕様は [システムインデックス](../design/index-data/system-index.md) にある |
| 検索結果ファイル | ［結果をファイルに出力］で書き出す `work\search_results.txt` にも本文（該当行）が入る |
| 持ち出し経路 | インデックスのフォルダはコピーして別の PC でも使えるほか、［インデックス管理］の［エクスポート…］で 1 つの zip に書き出し、別の PC・ワークスペースで［インポート…］して使える（[インデックスのエクスポート・インポート](../design/index-data/format.md#インデックスのエクスポートインポート)）。利用者の明示操作だが、設計上の持ち出し経路である。エクスポートした zip には、元のファイルの本文の文字・ファイル名・フォルダの構成・元のフォルダのパスが入るため、**元のファイルと同じ扱いで渡す・置く**（メールに添付する・共有フォルダに置くときは、元のファイルを渡すときと同じ判断をする） |
| 残留 | `work/` を消さない限りインデックスは残る。元のファイルを削除しても、インデックス側の本文は次のインデックス作成まで残る。前の版のワークスペース（`index\`）は自動で消さず、利用者が消す（取り込み直した後、`system_index\` は本ツールが消すが、`index\` はそのまま残る） |

運用での対策（導入時に決めることを推奨）:

- ワークスペース（既定は `%USERPROFILE%\Documents\tebunko_ws`。［設定］で変えられる）のアクセス権を、**クロール対象フォルダと同等以上に制限する**。
- 権限の異なる利用者が読める共有フォルダに、インデックスをそのまま置かない（ワークスペースを各利用者のローカルに置く）。
- 持ち出し（別 PC へのコピー）の可否を運用ルールで決める。
- ディスク暗号化（BitLocker）を前提にする。

本ツール側では、インデックスの場所をワークスペース配下に限定し、アクセス権の変更（`Set-Acl` 等）を一切行わない。したがって**インデックスの保護はワークスペースのアクセス権で決まる**。

## 原本の一時コピーと、その回収

インデックス作成中は、原本のコピーを作業領域（`work/tmp/<PC の鍵>/<PID>`）に作る（[取り込み対象のファイルは書き換えない](file-access.md#取り込み対象のファイルは書き換えない)）。作業領域はワークスペース（本文インデックスと同じ置き場所）の下に**限る**。機密文書のコピーがシステムの `%TEMP%` へ出ることは無い。ワークスペースのパスに `[` `]` が含まれる・長すぎるなど、作業フォルダをワークスペースの下に置けないときは、取り込みをすべてスキップする（どのファイルも中間 TSV などをこの作業領域に作るため、テキストファイルを含めすべての取り込みが対象になる。取り込み一覧に失敗として残る。[エラーメッセージ](../design/indexing/errors.md)）。

- **通常終了時**: ファイルごとに、読み終えたコピーを削除する。インデックス作成の完了時に作業フォルダごと空にする（`removeTmpDir`）。
- **異常終了時**: 強制終了・停電などでコピーが残ることがある。この場合、**次回のインデックス作成開始時に、終了済みのプロセスの作業フォルダをまとめて削除する**（`tebunko/indexer/index_migrate.ps1` の `removeStaleTmpDirs`。インデックス作成の開始時に `initTmpDir` が呼ぶ。`safety.Tests.ps1` の「異常終了で残った作業フォルダを次回起動時に回収する」が確かめる）。前の版（`%TEMP%\tebunko\<PID>` に一時ファイルを置いていた版）が残したものがあれば、同時に片付ける。
- **残る期間**: 異常終了から次回のインデックス作成開始までの間は残る。気になる場合は `work/tmp` を手で削除してよい（動作に影響しない）。

## 透過暗号化の製品が平文で見せたファイルの扱い

社内のファイル暗号化製品（透過型。IRM・秘密度ラベルの暗号化とは別の仕組み）が導入されている環境では、その製品が Office を「許可されたアプリ」として登録していることがある。その場合、本ツールが作業領域に作ったコピーを Office が開くと、製品が復号した**平文**を Office に渡す（[暗号化されたファイルの判定](../design/indexing/office-apps.md#暗号化されたファイルの判定office_protectionps1office_protection_viewps1)の「形式の分からないバイナリ」の予備）。

- **本文はインデックスに平文で入る**。これは「[インデックスが元文書の本文を保持する（情報の集約）](#インデックスが元文書の本文を保持する情報の集約)」と同じ性質で、透過暗号化を導入した目的（暗号化されていない経路での漏えい防止）と、本文インデックスに平文で保持することは緊張関係にある。
- **本ツールは製品を判別しない**。製品に許可されたアプリかどうかを問い合わせる手段を持たず、結果（読めた・読めなかった）でしか分からない。
- **読めなかった場合は理由付きで失敗にする**（`暗号化されているか壊れているため取り込めません。`）。中身の分からない暗号文をインデックスに書き込むことはない。
- **運用での対策**: 透過暗号化を導入している環境では、対象フォルダを取り込み対象にするかどうかを、[インデックスが元文書の本文を保持する（情報の集約）](#インデックスが元文書の本文を保持する情報の集約)の対策と合わせて判断する。

## 検索結果から開くと Excel が前面に出る

検索結果から元の Excel ファイルを開くときは、Excel を可視化して前面に出す（`tebunko/ui/open_source.ps1:162`・`213`）。利用者が開くよう操作したときの動作である。インデクサ側の Office は常に不可視で動作する（[Office ファイルを開くときの設定](checks.md#office-ファイルを開くときの設定)。`safety.Tests.ps1` が、`Visible = $true` は `open_source.ps1` にしか無いことを確かめる）。

## インストーラー版

配布物は 2 つある。中身のスクリプト（`scripts/`）は同じで、入れ方と起動口だけが違う。

| 配布物 | 向いている環境 | 実行ファイル |
|---|---|---|
| `tebunko-<タグ>.zip` | 実行ファイルを入れたくない・入れられない環境 | 含まない（起動口は `tebunko.bat`） |
| `tebunko-setup-<タグ>.exe` | 手軽にインストール・アンインストールしたい環境 | インストーラー自身と、入れる `tebunko.exe`・`uninstall\unins000.exe` |

インストーラーは [Inno Setup](https://jrsoftware.org/isinfo.php) 7 で作る（`installer/tebunko.iss`・`tools/new_installer.ps1`）。

- **入れる場所と権限**: 既定は管理者権限なしで `%LOCALAPPDATA%\Programs\tebunko` に入れる。管理者は、最初の画面で「すべてのユーザー」を選ぶと `C:\Program Files\tebunko` に入れられる。
- **入れるもの**: `tebunko.exe`・`scripts\`・`LICENSE`・`VERSION.txt`（版とコミットの記録。画面の「バージョン情報」に出す）と、Inno Setup のアンインストーラー（`uninstall\unins000.exe`・`uninstall\unins000.dat`。名前は Inno Setup が決め、変える設定が無いため、tebunko のものと分かるよう `uninstall` フォルダに置く）だけ。スタートメニューにショートカットを作る（デスクトップは選んだときだけ）。
- **レジストリ**: インストーラーが、Windows の「インストールされているアプリ」に出すためのアンインストールの情報（利用者ごとなら `HKCU`、すべてのユーザーなら `HKLM` の `Software\Microsoft\Windows\CurrentVersion\Uninstall\{E2FA5AC9-5A36-41AF-B5E2-B989E2878D3C}_is1`）を書く。これ以外のレジストリ、PATH、ファイルの関連付け、サービス、自動起動には触らない（`installer.Tests.ps1` が `[Registry]` の節が無いことなどを確かめる）。ツール本体（スクリプト）はレジストリを読み書きしない（[検査項目と結果](checks.md#検査項目と結果)）。
- **起動口 `tebunko.exe`**: 本リポジトリの `installer/tebunko.cs`（約 100 行）を、リリースのときに Windows 標準の `csc.exe`（.NET Framework）でビルドする。していることは `tebunko.bat` と同じで、Windows の PowerShell 5.1 で `gui.ps1` を `-ExecutionPolicy RemoteSigned` で起動するだけである（[Mark-of-the-Web の解除](#mark-of-the-web-の解除tebunkobat) の「`RemoteSigned` で起動する理由」）。インストーラーが書いたファイルには Mark-of-the-Web が付かないため、`Unblock-File` はしない。窓のアプリとしてビルドし、PowerShell も窓を作らずに起動する（`conhost.exe` は要らない）。PowerShell が 0 以外で終わりエラー出力があれば、その内容をメッセージで出す（実行ポリシーや制限言語モードで `gui.ps1` を読み込めないときに、何も出ずに終わらないようにするため）。エラー出力を受け取るため、画面を閉じるまで `tebunko.exe` のプロセスも残る（窓は無い）。
- **更新**: 新しい版のインストーラーを実行する。前の版の `scripts\` を消してから入れるため、前の版にだけあったスクリプトは残らない。`setting.config` は残る。
- **起動中の更新・削除**: `tebunko.exe` が起動中はミューテックス `tebunko-installed-app` を持ち、インストーラーとアンインストーラーはこれを見て、tebunko を閉じるよう求める（Inno Setup の `AppMutex`）。
- **アンインストール**: 「設定」→「アプリ」から行う。入れたファイル・ショートカット・アンインストールの情報と、ツールのフォルダの `setting.config` を消す。**インデックス（ワークスペース）は消さない。** ワークスペースは利用者が選んだフォルダのこともあり、中身を確かめずに消すのは危ないためである。インデックスには文書の文字がそのまま入っているため（[インデックスが元文書の本文を保持する（情報の集約）](#インデックスが元文書の本文を保持する情報の集約)）、アンインストールの最後に、残っている場所（既定は `%USERPROFILE%\Documents\tebunko_ws`）を知らせる。
- **無人のインストール**: `/VERYSILENT /CURRENTUSER`（アンインストールは `uninstall\unins000.exe /VERYSILENT`）で、画面を出さずに入れる・消すことができる。
- **署名**: インストーラーと `tebunko.exe` にはまだコード署名が無い。そのため、ダウンロードしたインストーラーを実行すると SmartScreen の「Windows によって PC が保護されました」が出る（［詳細情報］→［実行］で進める）。改ざんの確認には、ビルドの来歴の署名と SHA256 を使う（[配布物の完全性（カタログ・ハッシュ一覧・来歴の署名）](scans.md#配布物の完全性カタログハッシュ一覧来歴の署名)）。
- **再現性**: `csc.exe`（.NET Framework）と Inno Setup はビルドの日時を埋め込むため、同じソースからビルドしてもハッシュは一致しない。zip の中身と同じスクリプトが入っていることは、インストールしたフォルダでカタログ（[配布物の完全性（カタログ・ハッシュ一覧・来歴の署名）](scans.md#配布物の完全性カタログハッシュ一覧来歴の署名)）と比べれば分かる。カタログにある `tebunko.bat` はインストーラー版に無いため、`Status` は `ValidationFailed` になる。`tebunko.bat` を除いて違うファイルが無い（次の結果が空になる）ことを確かめる。

```powershell
$r = Test-FileCatalog -Path .\scripts -CatalogFilePath <ダウンロードした tebunko.cat> -Detailed
@($r.CatalogItems.Keys + $r.PathItems.Keys | Sort-Object -Unique | Where-Object { $_ -ne 'tebunko.bat' -and $r.CatalogItems[$_] -ne $r.PathItems[$_] })
```
- **第三者の部品**: インストーラーとアンインストーラーの本体は Inno Setup のものである（[Inno Setup License](https://jrsoftware.org/files/is/license.txt)）。インストールとアンインストールのときだけ動き、tebunko の実行中には使わない。ツール本体の第三者の部品は 0 のまま（[供給網（サプライチェーン）とライセンス](supply-chain.md)）。

## 単一 .ps1 版（試験版）

zip・インストーラーに加えて、展開せずに 1 本の `.ps1`（`tebunko-<タグ>.ps1`）だけで動く試験版を並べて配る（[単一 PowerShell のビルド](../design/structure/single-script.md)）。中身のスクリプトは同じ `scripts/` から機械的に結合したものである。

- **`tebunko.bat` に相当する起動口が無い**: 自分自身の Mark-of-the-Web を解除する動き（`Unblock-File`）は持たない。実行ポリシーも指定せず、右クリック［PowerShell で実行］や、呼び出す側が指定したポリシーのまま動く（[単一 PowerShell のビルド](../design/structure/single-script.md)「実行時の違い」）。
- **起動失敗の知らせ方**: 画面が開く前の失敗は、zip 版と同じ `reportStartupFailure`／`writeStartupErrorFile`（ツールのフォルダの `startup_error.txt`。書けなければ記録を残さない）に記録する。`tebunko.bat` の `catch` に相当する外側の受け皿が無いため、単一 .ps1 の起動口（`gui.ps1` の本体）の `try`／`catch` が直接この関数を呼ぶ。ただし `lib.ps1` の読み込みは `try` の外にあるため、その失敗は `reportStartupFailure` に届かず、PowerShell の窓に出る。
- **署名・改ざんの確認**: `tebunko.cat` は対象にしない（1 本のファイルのため、[配布物の完全性（カタログ・ハッシュ一覧・来歴の署名）](scans.md#配布物の完全性カタログハッシュ一覧来歴の署名)の SHA256SUMS.txt と来歴の署名だけで確かめる）。
- **試験版という扱い**: 利用者の確かめが済むまでは「試験版」とし、zip 版・インストーラー版と並べて配る。
