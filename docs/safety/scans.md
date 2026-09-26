# 第三者のツールによる検査結果

本ツールの作者以外が作ったツールでの検査結果を示す。5.1〜5.4 は誰でも同じコマンドで再現できる。

## 配布物の完全性（カタログ・ハッシュ一覧・来歴の署名）

GitHub Release の配布 zip（`tebunko-<タグ>.zip`）は、`v` で始まるタグを push したときに `.github/workflows/release.yml` が作る。作る前に `test.yml` と同じ検査・テストを通す。展開したときに何を起動すればよいかが分かるよう、zip にはツール本体と README・ライセンスだけを入れ、確認用のファイルは zip と並べてリリースに載せる（`tools/new_release_package.ps1`）。`docs/`・`tests/`・`work/`・`setting.config` はどちらにも入れない。SECURITY は README とリリースの説明からリンクする。

| ファイル | 置き場所 | 内容 |
|---|---|---|
| `tebunko.bat`・`scripts/` | zip の中 | ツール本体 |
| `README.md` | zip の中 | 使い方。相対リンクと画像は、その版の GitHub の URL に書き換えて入れる（`docs/` や画像は zip に入れないため） |
| `LICENSE` | zip の中 | ライセンス（MIT。写しに許諾表示を含めるため同梱する） |
| `tebunko-setup-<タグ>.exe` | リリース（zip の横） | インストーラー版（4.6。`tools/new_installer.ps1` が作る）。中身のスクリプトは zip と同じ |
| `tebunko.cat`・`SHA256SUMS.txt` | リリース（zip の横） | 改ざんの確認用（`tools/new_release_files.ps1` が作る） |
| `sbom.cdx.json` | リリース（zip の横） | 部品表 |

**カタログ（証明書は不要）**: Windows PowerShell 5.1 標準の `New-FileCatalog` で、`scripts/` 配下すべてと `tebunko.bat` の SHA256 を 1 つのカタログにまとめてある。受け取った側は、リリースから `tebunko.cat` もダウンロードし、zip を展開したフォルダで次を実行すれば、配布時点から 1 バイトも変わっていないことを確かめられる。

```powershell
Test-FileCatalog -Path .\scripts, .\tebunko.bat -CatalogFilePath <ダウンロードした tebunko.cat> -Detailed
```

| 項目 | 値 |
|---|---|
| 対象ファイル数 | 63（`scripts/` 配下すべてと `tebunko.bat`） |
| 検証結果 | `Status: Valid`（改ざんなし） |
| 署名 | `NotSigned`（コードサイニング証明書を導入すれば、発行者の保証も付く） |
| ハッシュ方式 | SHA256（`-CatalogVersion 2`） |

`SHA256SUMS.txt` は、同じファイルと `sbom.cdx.json`・`LICENSE` の SHA256 をテキストで並べたもので、目視・比較に使える（リリースにも単独で載せている）。手元でカタログとハッシュ一覧を作り直すときは `.\tools\new_release_files.ps1`（`work\release\` に出力）を実行する。

**来歴の署名**: 配布 zip には、ビルドの来歴（どのコミットから、どのワークフローで作ったか）の署名を付ける（GitHub の Artifact Attestations。Sigstore の証明書で署名し、コードサイニング証明書は要らない）。受け取った側は [GitHub CLI](https://cli.github.com/) で、zip が本リポジトリのワークフローで作られ、その後変わっていないことを確かめられる。

```powershell
gh attestation verify .\tebunko-v0.1.0.zip -R hsgwa/tebunko
```

署名の bundle（`tebunko-<タグ>.zip.sigstore.json`）もリリースに載せている。GitHub に問い合わせずに確かめるときは `--bundle .\tebunko-v0.1.0.zip.sigstore.json` を付ける。zip 自体の SHA256 はリリースの説明に書いてある（5.6 の照会に使える）。インストーラー（`tebunko-setup-<タグ>.exe`）にも同じく来歴の署名を付け、bundle（`tebunko-setup-<タグ>.exe.sigstore.json`）と SHA256 をリリースに載せる。確かめ方は zip と同じ（`gh attestation verify .\tebunko-setup-<タグ>.exe -R hsgwa/tebunko`）。

## 静的解析: PSScriptAnalyzer（Microsoft）

PowerShell スクリプトの静的解析ツール。Microsoft が公開しているルールで、コードインジェクション・平文パスワードなどを検出する。

```powershell
Install-Module PSScriptAnalyzer -RequiredVersion 1.25.0 -Scope CurrentUser   # 未導入の場合

# 安全性にかかわるルールでの検査（TypeNotFound のほかに 0 件であることを確認する）
Invoke-ScriptAnalyzer -Path .\scripts -Recurse -Settings .\tests\meta\PSScriptAnalyzer.security.psd1

# 全ルールでの検査（書き方のルールを含む）
Invoke-ScriptAnalyzer -Path .\scripts -Recurse
```

実施結果（PSScriptAnalyzer 1.25.0）:

| 検査 | 結果 |
|---|---|
| 安全性にかかわる 14 ルール（`tests/meta/PSScriptAnalyzer.security.psd1`） | **指摘 0 件**（`TypeNotFound` を除く。下の注を参照） |
| 全ルールのうち `Error` 重大度 | **指摘 0 件**（`TypeNotFound` を除く） |
| 全ルール | 262 件（`Warning` 219 件 / `Information` 43 件。内訳は下表） |

CI（`test.yml` の `test`）は、版を 1.25.0 に固定した PSScriptAnalyzer で `Error` 重大度の指摘が 0 件であることを確かめ、1 件でもあれば失敗する（版を固定するのは、新しい版でルールが増えてもコードを変えずに CI が落ちないようにするため）。`tests/meta/safety.Tests.ps1` の「サードパーティの静的解析」は、PSScriptAnalyzer が入っている環境で、安全性の 14 ルール・`Error` 重大度・制限言語モードの指摘が 0 件であることを確かめる（入っていない環境では飛ばす）。

有効にしている安全性のルールは次の 14 件である。設定ファイルからルールを削って「0 件」にする抜け道を防ぐため、この 14 件が設定に含まれていることも `tests/meta/safety.Tests.ps1` で検査している。

| 分類 | ルール |
|---|---|
| 文字列を式として実行しない | `PSAvoidUsingInvokeExpression` |
| パスワード・資格情報の平文の扱い | `PSAvoidUsingPlainTextForPassword`・`PSAvoidUsingConvertToSecureStringWithPlainText`・`PSUsePSCredentialType`・`PSAvoidUsingUsernameAndPasswordParams` |
| 通信先の固定・暗号化されない通信 | `PSAvoidUsingComputerNameHardcoded`・`PSAvoidUsingAllowUnencryptedAuthentication` |
| 危険・非推奨な API | `PSAvoidUsingBrokenHashAlgorithms`・`PSAvoidUsingWMICmdlet` |
| 標準コマンドの上書き・別名の定義 | `PSAvoidOverwritingBuiltInCmdlets`・`PSAvoidGlobalAliases`・`PSReservedCmdletChar`・`PSReservedParams` |
| 文字コード | `PSUseBOMForUnicodeEncodedFile` |

全ルールで出る 262 件の内訳は次のとおりで、**すべて書き方（可読性・保守性）のルール**である。安全性の判断には関わらない。

| 件数 | ルール | 内容と本ツールでの理由 |
|---|---|---|
| 59 | `PSAvoidUsingWriteHost` | `Write-Host` でコンソールへ出力している。画面を使わずに `indexer.ps1` を実行したときに、インデクサの進み具合と失敗の原因を利用者に見せるための出力（`writeIndexerLog`）であり、パイプラインへ返す値ではない |
| 51 | `PSUseDeclaredVarsMoreThanAssignments` | `-ErrorVariable` で受け取る変数、`${...}` 記法で別ファイルから参照する変数など、別の構文で使う変数 |
| 46 | `PSReviewUnusedParameter` | 画面のイベントハンドラーの引数（`sender`・`e`）など、WPF の呼び出し規約で受け取るだけの引数 |
| 37 | `PSAvoidUsingPositionalParameters` | COM の呼び出し（`Workbooks.Open` 等）は引数の順序が API 側で決まっており、名前付きにできない |
| 36 | `PSAvoidUsingEmptyCatchBlock` | 権限の無いプロセスの起動時刻が取れない場合など、失敗しても処理を続けるべき箇所。理由は各 `catch` にコメントで記載している |
| 27 | `PSAvoidAssignmentToAutomaticVariable` | WPF のイベントハンドラーの引数名 `sender` に対する指摘。PowerShell の自動変数と同名だが、ハンドラー内の引数であり影響しない |
| 6 | `TypeNotFound` | `tebunko/ui/types.ps1` が、`shared/ui/types.ps1` の型（`NotifyBase`）を継承しているため。1 ファイルだけを解析すると継承元が見つからないという指摘で、実際の読み込み順では解決する（`tests/meta/structure.Tests.ps1` の「型の読み込み」で実際に読み込んで確認している）。どのルールを指定しても出る |

なお `tests/testdata/make_testdata.ps1` には `PSAvoidUsingPlainTextForPassword` の指摘が 7 件出る。パスワード付きの Office ファイルを**テストデータとして作る**スクリプトであり、パスワードは固定のテスト用文字列である。配布 zip には含まれない。

## マルウェア検査: Microsoft Defender（外部への送信なし）

ローカルのウイルス対策エンジンでフォルダを検査する。ファイルを外部へ送信しないため、社内限りのコードでも実行できる。

```powershell
Start-MpScan -ScanType CustomScan -ScanPath (Get-Location).Path
Get-MpThreatDetection | Where-Object { $_.InitialDetectionTime -gt (Get-Date).AddMinutes(-10) }
Get-MpComputerStatus | Select-Object AMProductVersion, AntivirusSignatureVersion, AntivirusSignatureLastUpdated
```

実施結果（2026/09/21 実施）:

| 項目 | 値 |
|---|---|
| 検出 | **0 件** |
| エンジン | Microsoft Defender Antivirus 4.18.26080.4 |
| 定義ファイル | 1.459.309.0（2026/09/21 更新） |

さらに、Windows PowerShell 5.1 は実行するスクリプトブロックを **AMSI**（Antimalware Scan Interface）に渡す。Defender を有効にした PC で本ツールを動かしている限り、実行のたびに第三者のエンジンが内容を検査していることになる。

## 実行環境の制約との適合

アプリケーション制御を強制している環境で動くかどうかは、安全性の間接的な指標になる（制限された環境で動くコードは、危険な機能を使っていない）。

| 指標 | 結果 | 確認コマンド |
|---|---|---|
| 制限言語モード（Constrained Language Mode）で使えない書き方 | **指摘 0 件**（`TypeNotFound` を除く） | `Invoke-ScriptAnalyzer -Path .\scripts -Recurse -IncludeRule PSUseConstrainedLanguageMode` |
| Windows PowerShell 5.1 の構文互換 | **指摘 0 件**（`TypeNotFound` を除く） | `Invoke-ScriptAnalyzer -Path .\scripts -Recurse -IncludeRule PSUseCompatibleSyntax -Settings @{ Rules = @{ PSUseCompatibleSyntax = @{ Enable = $true; TargetVersions = @('5.1') } } }` |
| AppLocker・WDAC との相性 | 実行時コンパイル（`csc.exe`）を行わないため、一時 DLL の生成でブロックされることがない（[検査項目と結果](checks.md#検査項目と結果)） | – |

ただし実際に制限言語モードを強制した環境では、COM オブジェクトの生成が制限されるため、**実動確認は別に必要**である（本ツールは Excel の COM を使う）。

攻撃面の縮小（ASR）ルールとの関係は次のとおり。

| ASR ルール | 本ツールへの影響 |
|---|---|
| Office アプリが子プロセスを作るのをブロック | **該当しない**。本ツールは逆方向（PowerShell が Office を起動する）であり、Office 側から子プロセスを作らない |
| 潜在的に難読化されたスクリプトの実行をブロック | 該当しない想定（難読化を行わないため）。導入先で有効な場合は、まず監査モードで 1 回通し実行して確認することを推奨。**このルールが通ることは、Microsoft のヒューリスティックが難読化と判定しないことの証拠にもなる** |
| Office マクロからの Win32 API 呼び出しをブロック | 該当しない。マクロを強制無効で開く（[Office ファイルを開くときの設定](checks.md#office-ファイルを開くときの設定)） |

## 実行時の観測（動作の記録による確認）

「コードを読んだ結果」よりも「動かして記録した結果」のほうが強い証拠になる。導入先の環境で 1 回実施すれば、[危険とされる処理の検査結果](checks.md)・[ファイルの読み書きの範囲](file-access.md)の主張を実測で裏付けられる。

| 手段 | 何が分かるか | 手順 |
|---|---|---|
| Process Monitor（Sysinternals）または Sysmon | ファイル・レジストリ・ネットワーク・プロセス生成の**全アクセス**。「`work` と `%TEMP%` 以外に書いていない」「通信していない」を実測で示せる | Procmon でプロセス名 `powershell.exe` / `EXCEL.EXE` を絞り込み、インデックス作成を 1 回実行して保存する |
| 送信の全遮断で完走 | 通信が不要であること | Windows ファイアウォールで送信を全ブロック、または Windows Sandbox（ネットワーク無効）でインデックス作成を完走させる |
| PowerShell のログ | 実行された全コマンド | グループポリシーでスクリプトブロックログ（イベント ID 4104）・モジュールログ・トランスクリプションを有効にして 1 回実行し、記録を提出する |
| 標準ユーザーで完走 | 管理者権限が不要であること | 管理者権限のないアカウントで起動して一通り操作する |
| 対象フォルダを読み取り専用にして完走 | 原本を書き換えないこと（3.3 手順 B） | 読み取りの権限だけを与えたアカウントで実行する |

参考: 通信状況はインデックス作成中に次のコマンドでも確認できる。

```powershell
Get-NetTCPConnection -OwningProcess (Get-Process powershell).Id -ErrorAction SilentlyContinue
```

## 複数エンジンでの検査: VirusTotal（外部へファイルを送信する）

約 70 社のエンジンで一度に判定でき、結果が URL として残る。ただし**アップロードしたファイルはセキュリティ各社と共有される**ため、社外に出せないコードを含む場合は実施しない。その場合は次のいずれかで代える。

- 5.3 のローカルスキャン（送信なし）で代える。
- ファイルのハッシュだけを照会する（ハッシュは内容を復元できない。ただし過去に誰かが同じファイルを提出していない限り「未登録」となる）。配布 zip の SHA256 はリリースの説明にも書いてある。

```powershell
Get-FileHash .\tebunko-v0.1.0.zip -Algorithm SHA256    # この値を VirusTotal の検索欄で照会する
```

## 同梱の機械検査と品質の指標

[危険とされる処理の検査結果](checks.md)・[ファイルの読み書きの範囲](file-access.md)の主張と 5.2・5.4 の結果は、1 つのコマンドで検査できる（リポジトリを clone して実行する。配布 zip には `tests/` を入れていない）。

```powershell
.\tests\run.ps1 -Tag Meta     # 安全性・構成の検査だけ
.\tests\run.ps1 -Ci           # 既定のテストとカバレッジ（CI と同じ）
```

`tests/meta/safety.Tests.ps1`（34 件）は次を検査する。将来の変更でいずれかの前提が崩れれば失敗し、CI の必須チェック `test` が通らないため main にマージできない。

- [危険とされる処理の検査結果](checks.md)の禁止する処理が 0 件であること、許す処理（`Add-Type`・`Start-Process`・`Stop-Process`）が限定されていること
- Office をマクロ無効・イベント無効・外部リンク更新なし・不可視・読み取り専用で開くこと
- 原本のパスを書き込み・削除の API に渡さないこと、`SaveAs` の保存先が作業フォルダだけであること、原本を読むのは `copyFileShared` の読み取りだけであること
- 書き込み先が `work` 配下と `%TEMP%` 配下に限られること、異常終了で残った作業フォルダを次回起動時に回収すること（[原本の一時コピーと、その回収](disclosure.md#原本の一時コピーとその回収)）
- PSScriptAnalyzer の安全性ルール・`Error` 重大度・制限言語モードの指摘が 0 件であること（未導入の環境では飛ばす）
- 安全性の説明（`docs/safety/`）・[SECURITY.md](../../.github/SECURITY.ja.md)・`tools\new_release_files.ps1`・[sbom.cdx.json](../../sbom.cdx.json)・[LICENSE](../../LICENSE) がそろっており、SBOM が第三者の部品を含まないこと、LICENSE が MIT の条文と著作権表示を含み SBOM の記載と一致すること

テストとコードの指標（安全性の間接的な根拠）:

| 指標 | 値 |
|---|---|
| テスト件数 | 1,189 件（既定のタグ Unit・Io・Meta。失敗 0 件。うち安全性・構成の検査（タグ `Meta`）が 122 件） |
| コードカバレッジ | 97.0%（6,473/6,671 コマンド。手元の実測。CI では本物の Windows Search を使うテストが保留になるため少し低くなる）。対象は画面の起動口・タブ・ダイアログを除くスクリプト（`tests/run.ps1` の `CodeCoverage` の条件）。Office の COM を使うインデクサは、COM の入口（`getApp`）を偽のオブジェクトに差し替えて検証している。検索結果の一覧・プレビューなど計測の対象に入る画面の部品は、画面のコントロールを偽のオブジェクトに差し替えて検証している。下限は `tests/coverage.baseline`（90.0%）で、下回ると CI が失敗する |
| 規模 | `scripts/` 配下 48 ファイル・11,306 行（空行を除く）・389 関数 |
| 第三者依存 | 0 件 |

規模と依存の数を示す意味は、**監査にかかる手間を見積もれる**ことである。第三者依存が無いため、読む範囲はこの 11,306 行で閉じる。
