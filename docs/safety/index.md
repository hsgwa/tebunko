# 安全性の要約

この区分は、本ツールの導入を審査する担当者（情報システム部門・セキュリティ部門など）と、OSS としての安全性を評価する人に向けて、**本ツールが何をするか・何をしないか**と、**その根拠をどう自分の手で確かめられるか**を示す。

ここでの主張はすべて、このリポジトリのコード・設定に基づく。各主張には、**それを機械的に守らせている仕組み**（テスト・CI）と、**審査する側が自分で再現するコマンド**を併記する。この説明を信用しなくても、同じ結果を自分で確かめられる。

脆弱性・不審な挙動の連絡先は [SECURITY.md](../../.github/SECURITY.ja.md)（GitHub の非公開の報告窓口）。

| 観点 | 本ツールの挙動 | 守らせている仕組み | 詳細 |
|---|---|---|---|
| 構成物 | **zip 版**: Windows PowerShell スクリプト（`scripts/**/*.ps1`）、画面定義（`*.xaml`）、起動用 `tebunko.bat`、アイコン `tebunko.ico` のみ。実行可能バイナリ（`.exe` / `.dll`）を同梱しない。**インストーラー版**: 同じスクリプトに、起動用の `tebunko.exe`（本リポジトリのソースからビルド）と、Inno Setup のインストーラー・アンインストーラーが加わる | zip の中身は `tools/new_release_package.ps1`、インストーラーの中身は `installer/tebunko.iss` が決める（`installer.Tests.ps1`） | [配布物の完全性（カタログ・ハッシュ一覧・来歴の署名）](scans.md#配布物の完全性カタログハッシュ一覧来歴の署名)・[インストーラー版](disclosure.md#インストーラー版) |
| 第三者ライブラリ | **使用しない**。実行時の依存は Windows 標準の .NET アセンブリと Microsoft Office のみ | `safety.Tests.ps1`（SBOM に第三者の部品が無いこと） | [供給網（サプライチェーン）とライセンス](supply-chain.md) |
| ネットワーク通信 | **行わない**。通信用の API を使っていない | `safety.Tests.ps1`「ネットワーク通信を行わない」 | [検査項目と結果](checks.md#検査項目と結果) |
| 動的コード実行・難読化 | **行わない**。`Invoke-Expression`、文字列からのスクリプト生成、Base64 のコマンドを使っていない | `safety.Tests.ps1`・PSScriptAnalyzer | [検査項目と結果](checks.md#検査項目と結果) |
| 実行時コンパイル・P/Invoke | **行わない**。C# の `Add-Type` コンパイル（`csc.exe` の起動）と Windows API の直接呼び出しを使っていない | `safety.Tests.ps1`・`structure.Tests.ps1` | [検査項目と結果](checks.md#検査項目と結果) |
| 管理者権限・常駐 | **不要・しない**。昇格要求・サービス登録・タスクスケジューラ登録・自動起動の登録を行わない。画面を閉じれば終了する。インストーラー版も既定は管理者権限なしで入れる | `safety.Tests.ps1`・`installer.Tests.ps1` | [検査項目と結果](checks.md#検査項目と結果) |
| レジストリ | **読み書きしない**。インストーラー版では、インストーラーがアンインストールの情報（`HKCU` または `HKLM` の `...NINSTALL`）だけを書き、アンインストールで消す | `SAFETY.TESTS.PS1`・`INSTALLER.TESTS.PS1` | [検査項目と結果](checks.md#検査項目と結果)・[インストーラー版](disclosure.md#インストーラー版) |
| Windows Search | **読み取りだけ**。高速検索（[検索](../design/search/index.md) 4.4）で、OLE DB（`Search.CollatorDSO`）に SELECT だけを送る（`windows_search.ps1` の 1 か所。管理者権限は要らない）。Windows Search の設定（索引の対象など）もファイルのアクセス権も変えない。システムインデックスを作るスレッドの優先度を下げる（BelowNormal。プロセス全体の優先度は変えない） | `safety.Tests.ps1`「Windows Search への問い合わせは windows_search.ps1 だけで行い、SELECT だけを送る」 | [検査項目と結果](checks.md#検査項目と結果) |
| 実行ポリシー | `Bypass` を使わず `RemoteSigned` で起動する。恒久的な変更（`Set-ExecutionPolicy`）をしない | `safety.Tests.ps1` | [Mark-of-the-Web の解除](disclosure.md#mark-of-the-web-の解除tebunkobat) |
| 取り込み対象のファイル | **読み取りのみ**。原本を開かず、作業フォルダへコピーしたものを読む | `safety.Tests.ps1` | [取り込み対象のファイルは書き換えない](file-access.md#取り込み対象のファイルは書き換えない) |
| マクロ | **実行させない**。Office のマクロを強制無効にしてから開く | `safety.Tests.ps1` | [Office ファイルを開くときの設定](checks.md#office-ファイルを開くときの設定) |
| 資格情報 | 保存・送信・入力要求のいずれも行わない。パスワード付きファイルは解除を試みず、失敗として記録する | `safety.Tests.ps1`・PSScriptAnalyzer | [Office ファイルを開くときの設定](checks.md#office-ファイルを開くときの設定) |
| 書き込み先 | `work/` 配下（既定は `%USERPROFILE%\Documents\tebunko_ws`。利用者が画面で選んだフォルダ（ワークスペース）にも置ける）、`%TEMP%\tebunko\<PID>` 配下、`setting.config`（ツールのフォルダに書き込めなければ `%LOCALAPPDATA%\tebunko\<鍵>`）、利用者が指定した検索結果の出力先のみ | `safety.Tests.ps1` | [書き込み・削除する場所](file-access.md#書き込み削除する場所) |
| 配布物の完全性 | 配布 zip と並べて、カタログ（`tebunko.cat`）とハッシュ一覧（`SHA256SUMS.txt`）をリリースに載せ、zip とインストーラーにはビルドの来歴の署名を付ける | `release.yml` | [配布物の完全性（カタログ・ハッシュ一覧・来歴の署名）](scans.md#配布物の完全性カタログハッシュ一覧来歴の署名) |
| 開発の過程 | main へは PR からだけ入れ、テスト・CodeQL などの必須チェックを通す。コミットには `Signed-off-by`（DCO）を付ける。GitHub Actions は版を固定し、Dependabot が更新する | ブランチ保護・CI | [供給網（サプライチェーン）とライセンス](supply-chain.md) |
| 規模（監査の目安） | `scripts/` 配下 48 ファイル・11,306 行（空行を除く）・389 関数。第三者依存が無いため、監査対象はこの範囲で閉じる | – | [同梱の機械検査と品質の指標](scans.md#同梱の機械検査と品質の指標) |

第三者が作ったツールでの検査結果は [第三者のツールによる検査結果](scans.md)のとおり。PSScriptAnalyzer（Microsoft）の安全性にかかわるルールは**指摘 0 件**、Microsoft Defender のスキャンは**検出 0 件**である。

利用者の明示操作により影響が出る機能と、ツールの性質上避けられない情報リスクは [開示事項](disclosure.md)で開示している（Office プロセスの強制終了、Mark-of-the-Web の解除、**インデックスが元文書の本文を保持すること**、一時コピー、Excel が前面に出ること）。

## この区分のページ

- [危険とされる処理の検査結果](checks.md)
- [ファイルの読み書きの範囲](file-access.md)
- [開示事項](disclosure.md)
- [第三者のツールによる検査結果](scans.md)
- [供給網（サプライチェーン）とライセンス](supply-chain.md)
