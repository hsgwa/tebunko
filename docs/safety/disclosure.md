# 開示事項

次の 6 点は審査で論点になるため、本ツールの側から先に開示する。4.1・4.2・4.5 は利用者の操作に紐づく動作、4.3・4.4 はツールの性質上避けられない情報の扱い、4.6 はインストーラー版だけに当てはまる動作である。

## Office プロセスの強制終了（［9 プロセス停止］タブ）

インデックス作成が異常終了すると、画面の無い Excel・Word・PowerPoint のプロセスが残り、メモリを占有し続けることがある。これを止めるために、画面の［9 プロセス停止］タブから `Stop-Process -Force` を実行する（`shared/office/office_process.ps1:43`。`safety.Tests.ps1` がこの 1 か所だけであることを確かめる）。

- **影響**: 強制終了するため、対象プロセスで**未保存の内容は失われる**。
- **制御**: 実行されるのは利用者が一覧で選んだ PID のみである（`stopOfficeProcesses` は PID の配列を受け取る）。自動では実行されない。監視・常駐もしない。
- **判断材料の提示**: 一覧には、ウィンドウを持たないプロセス（インデクサが残したものと判断できる）を区別して表示する（`getOfficeProcesses` の `Background` 列）。利用者が自分で開いている Excel を誤って選ばないようにするための表示である。
- **対象の限定**: 対象は Excel（`EXCEL`）・Word（`WINWORD`）・PowerPoint（`POWERPNT`）のみ。他のプロセスは一覧に出ないため、選ぶことができない。

詳細は [［9 プロセス停止］タブ](../design/gui/process-tab.md) を参照。

## Mark-of-the-Web の解除（`tebunko.bat`）

`tebunko.bat` は PowerShell を 1 回起動し、その中で配布フォルダの `scripts` 配下の Mark-of-the-Web を解除してから、画面のスクリプトを実行する。

```bat
start "" conhost.exe powershell -NoProfile -STA -ExecutionPolicy RemoteSigned -WindowStyle Hidden -Command "Get-ChildItem -LiteralPath '%~dp0scripts' -Recurse -File | Unblock-File -ErrorAction SilentlyContinue; & '%~dp0scripts\tebunko\gui.ps1'"
```

- **理由**: zip で配布したものを展開すると全ファイルに Mark-of-the-Web が付き、実行ポリシー `RemoteSigned` では未署名スクリプトが実行できないため。
- **範囲**: `%~dp0scripts`（`tebunko.bat` があるフォルダの `scripts`）配下のみ。実行するスクリプトはすべてここにある。インデックス（`work`）には触らない。システム全体やユーザーのドキュメントには影響しない。画面（`gui.ps1`）も起動時に同じ `scripts` 配下に対して同じ解除を行う（ショートカットから直接起動したときのため）。
- **実行ポリシー**: `Bypass` は使わず `RemoteSigned` で起動する。恒久的な設定変更（`Set-ExecutionPolicy`）も行わない。`-ExecutionPolicy` はこのプロセスだけに効く。理由は次の「`RemoteSigned` で起動する理由」を参照。
- **`conhost.exe` を通す理由**: 既定のターミナルが Windows Terminal の場合でも `-WindowStyle Hidden` で PowerShell の窓を隠すため（[画面の共通仕様](../design/gui/common.md#配布と実行ポリシーmark-of-the-web)）。
- **恒久的な解消策**: 配布物のカタログ（[配布物の完全性（カタログ・ハッシュ一覧・来歴の署名）](scans.md#配布物の完全性カタログハッシュ一覧来歴の署名)）にコードサイニング証明書で署名すれば、この `Unblock-File` は不要になり、`AllSigned` でも動作する。

### `RemoteSigned` で起動する理由

Windows のクライアント版（Windows 10 / 11）は、実行ポリシーの既定が `Restricted` で、`.ps1` を 1 つも実行できない。本ツールはスクリプトでできているため、何らかの実行ポリシーを指定しないと起動できない。そこで、`tebunko.bat` と画面（インデクサの起動）は、`powershell.exe` の `-ExecutionPolicy RemoteSigned` で**そのプロセスだけ**の実行ポリシーを指定する。

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

本ツールは全文検索のために、**元の文書の本文を平文の TSV としてインデックス（`work\index\`）に保持する**。これはツールの目的そのものであり、避けられない。審査では次の点が論点になる。

| 論点 | 内容 |
|---|---|
| アクセス権の非継承 | インデックスは**作成した人の権限で作られる**。元のファイルに設定されたアクセス権は引き継がれない。インデックスを置いた場所を読める人は、元のファイルを読む権限が無くても本文を読める |
| 検索結果ファイル | ［結果をファイルに出力］で書き出す `work\検索結果.txt` にも本文（該当行）が入る |
| 持ち出し経路 | インデックスのフォルダは別の PC へコピーして使える（README の「インデックスを別の PC で使う」）。利用者の明示操作だが、設計上の持ち出し経路である |
| 残留 | `work/` を消さない限りインデックスは残る。元のファイルを削除しても、インデックス側の本文は次のインデックス作成まで残る |

運用での対策（導入時に決めることを推奨）:

- ツールを置いたフォルダのアクセス権を、**クロール対象フォルダと同等以上に制限する**。
- 権限の異なる利用者が読める共有フォルダに、インデックスをそのまま置かない（`work/index/` を各利用者のローカルに置く）。
- 持ち出し（別 PC へのコピー）の可否を運用ルールで決める。
- ディスク暗号化（BitLocker）を前提にする。

本ツール側では、インデックスの場所を `work/` 配下に限定し、アクセス権の変更（`Set-Acl` 等）を一切行わない。したがって**インデックスの保護は設置フォルダのアクセス権で決まる**。

## 原本の一時コピーと、その回収

インデックス作成中は、原本のコピーを `%TEMP%\tebunko\<PID>` に作る（[取り込み対象のファイルは書き換えない](file-access.md#取り込み対象のファイルは書き換えない)）。機密文書のコピーが一時的に `%TEMP%` に存在することになるため、扱いを明記する。

- **通常終了時**: ファイルごとに、読み終えたコピーを削除する。インデックス作成の完了時に作業フォルダごと空にする（`removeTmpDir`）。
- **異常終了時**: 強制終了・停電などでコピーが残ることがある。この場合、**次回のインデックス作成開始時に、終了済みのプロセスの作業フォルダをまとめて削除する**（`tebunko/indexer/index_migrate.ps1:29` の `removeStaleTmpDirs`。`tebunko/indexer.ps1:97` で起動時に呼ぶ。`safety.Tests.ps1` の「異常終了で残った作業フォルダを次回起動時に回収する」が確かめる）。
- **残る期間**: 異常終了から次回のインデックス作成開始までの間は残る。気になる場合は `%TEMP%\tebunko` を手で削除してよい（動作に影響しない）。

## 検索結果から開くと Excel が前面に出る

検索結果から元の Excel ファイルを開くときは、Excel を可視化して前面に出す（`tebunko/ui/open_source.ps1:162`・`213`）。利用者が開くよう操作したときの動作である。インデクサ側の Office は常に不可視で動作する（2.2。`safety.Tests.ps1` が、`Visible = $true` は `open_source.ps1` にしか無いことを確かめる）。

## インストーラー版

配布物は 2 つある。中身のスクリプト（`scripts/`）は同じで、入れ方と起動口だけが違う。

| 配布物 | 向いている環境 | 実行ファイル |
|---|---|---|
| `tebunko-<タグ>.zip` | 実行ファイルを入れたくない・入れられない環境 | 含まない（起動口は `tebunko.bat`） |
| `tebunko-setup-<タグ>.exe` | 手軽にインストール・アンインストールしたい環境 | インストーラー自身と、入れる `tebunko.exe`・`uninstall\unins000.exe` |

インストーラーは [Inno Setup](https://jrsoftware.org/isinfo.php) 7 で作る（`installer/tebunko.iss`・`tools/new_installer.ps1`）。

- **入れる場所と権限**: 既定は管理者権限なしで `%LOCALAPPDATA%\Programs\tebunko` に入れる。管理者は、最初の画面で「すべてのユーザー」を選ぶと `C:\Program Files\tebunko` に入れられる。
- **入れるもの**: `tebunko.exe`・`scripts\`・`LICENSE` と、Inno Setup のアンインストーラー（`uninstall\unins000.exe`・`uninstall\unins000.dat`。名前は Inno Setup が決め、変える設定が無いため、tebunko のものと分かるよう `uninstall` フォルダに置く）だけ。スタートメニューにショートカットを作る（デスクトップは選んだときだけ）。
- **レジストリ**: インストーラーが、Windows の「インストールされているアプリ」に出すためのアンインストールの情報（利用者ごとなら `HKCU`、すべてのユーザーなら `HKLM` の `Software\Microsoft\Windows\CurrentVersion\Uninstall\{E2FA5AC9-5A36-41AF-B5E2-B989E2878D3C}_is1`）を書く。これ以外のレジストリ、PATH、ファイルの関連付け、サービス、自動起動には触らない（`installer.Tests.ps1` が `[Registry]` の節が無いことなどを確かめる）。ツール本体（スクリプト）はレジストリを読み書きしない（[検査項目と結果](checks.md#検査項目と結果)）。
- **起動口 `tebunko.exe`**: 本リポジトリの `installer/tebunko.cs`（約 100 行）を、リリースのときに Windows 標準の `csc.exe`（.NET Framework）でビルドする。していることは `tebunko.bat` と同じで、Windows の PowerShell 5.1 で `gui.ps1` を `-ExecutionPolicy RemoteSigned` で起動するだけである（4.2 の「`RemoteSigned` で起動する理由」）。インストーラーが書いたファイルには Mark-of-the-Web が付かないため、`Unblock-File` はしない。窓のアプリとしてビルドし、PowerShell も窓を作らずに起動する（`conhost.exe` は要らない）。PowerShell が 0 以外で終わりエラー出力があれば、その内容をメッセージで出す（実行ポリシーや制限言語モードで `gui.ps1` を読み込めないときに、何も出ずに終わらないようにするため）。エラー出力を受け取るため、画面を閉じるまで `tebunko.exe` のプロセスも残る（窓は無い）。
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
