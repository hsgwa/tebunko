# 開発に参加する

[English](CONTRIBUTING.md) | 日本語

tebunko への不具合の報告・要望・修正の提案を歓迎します。このページは、変更を提案するときの手順と決まりをまとめたものです。

- 使い方の質問は [サポート](SUPPORT.ja.md) を見てください。
- 安全性に関わる問題は、Issue ではなく [セキュリティ](SECURITY.ja.md) の手順で非公開で連絡してください。
- 参加する人は [行動規範](CODE_OF_CONDUCT.ja.md) に従ってください。
- Issue・Pull Request は日本語でも英語でもかまいません。

## 変更の流れ

変更は **Issue → ブランチ → Pull Request → main** の順で入れます。

1. **Issue を立てる。** 不具合は「不具合」、機能の追加・変更は「機能の要望」のテンプレートを使います。タイトルは下の「コミットと Pull Request のタイトル」の形にします（テンプレートが `fix: ` `feat: ` を最初から入れます）。誤字の修正のような小さな変更は、Issue なしで Pull Request を出してかまいません。
2. **main から作業用のブランチを作る。**
3. **変更し、テストを通す**（下の「テスト」）。
4. **Pull Request を出す。** 出す前に最新の main を merge で取り込みます（main を取り込んでいないブランチはマージできません。push 済みのブランチを rebase して force push しないでください）。本文はテンプレートに沿って書き、`Closes #<Issue の番号>` で Issue とつなぎます。前の版と互換が無くなる変更では、移行の手順を書きます。ラベルはメンテナが付けます。
5. **CI が通り、レビューが済んだら squash merge します。** Pull Request のタイトルが main のコミットのタイトルになるので、下の「コミットと Pull Request のタイトル」の形で、変更の内容が分かる 1 行にしてください。

1 つの Pull Request には 1 つの変更だけを入れます。関係のない修正は別の Pull Request にしてください。

### コミットと Pull Request のタイトル

コミットメッセージの 1 行目、Pull Request のタイトル、Issue のタイトルは [Conventional Commits](https://www.conventionalcommits.org/ja/v1.0.0/) の形にします。型は英語、説明は日本語で書きます（英語で書いてもかまいません）。

```
<型>(<範囲>)!: <説明>

feat: Excel の図形の文字を検索できるようにする
fix(gui): 検索結果の件数が 0 のままになるのを直す
feat!: インデックスの形式を変える
```

- 範囲（`(gui)` など）は省略できます。書くときは空白を含まない 1 語にします。
- `!` は、設定ファイルやインデックスが前の版のまま使えなくなる変更に付けます。Pull Request の本文に移行の手順を書きます。
- 型の後ろは半角のコロンと空白 1 つです。

| 型 | 使うとき |
|---|---|
| `feat` | 機能の追加・変更 |
| `fix` | 不具合の修正 |
| `docs` | 文書だけの変更 |
| `refactor` | 動きを変えない書き直し |
| `perf` | 速さの改善 |
| `test` | テストだけの追加・修正 |
| `style` | 書式だけの変更（空白・改行など） |
| `build` | 配布物の作り方・依存の更新 |
| `ci` | CI・git のフック・開発用の道具 |
| `chore` | 上のどれにも当たらないもの |
| `revert` | 前の変更の取り消し |

形は機械的に確かめます。コミットは `commit-msg` フック（下の「開発の準備」）、Pull Request のタイトルは CI が確かめ、形が違うとマージできません。Issue のタイトルが違う形のときは、直し方が自動でコメントされます。git が自動で作るメッセージ（`Merge ...` `Revert "..."` `fixup! ...`）は調べません。

### Signed-off-by

コミットは `git commit -s` で作り、`Signed-off-by: <名前> <メールアドレス>` を付けてください。[DCO（Developer Certificate of Origin）](https://developercertificate.org/)に同意し、その変更をこのリポジトリのライセンスで出す権利があることを表します。メールアドレスはコミットの作者のものと同じにしてください。

付いていないコミットは `commit-msg` フックが止めます。Pull Request では CI が調べ、付いていないコミットがあるとマージできません。付け忘れたまま push したときは、`git rebase --signoff origin/main` で付け直して `git push --force-with-lease` してください。

## 開発の準備

必要なもの:

- Windows と Windows PowerShell 5.1（Windows に最初から入っています）
- Pester 3.4（テストの実行に使います。Windows に最初から入っている版です）
- Microsoft Excel・Word・PowerPoint（インデックス作成を実際に試すとき。自動テストには不要です）

clone したら、コミット前の検査を有効にします（1 回だけ）。

```powershell
.\tools\install_hooks.ps1
```

以後、コミットのたびに次の検査が動きます。引っかかったときは `--no-verify` で飛ばさず、内容を直してください。

- 個人情報（実名・メールアドレス・利用者名を含むパスなど）と、スクリプトの文字コードの検査（`tools\check_commit.ps1 -Staged`）
- 速いテスト（`tests\run.ps1 -Tag Unit,Meta -Quiet`）
- コミットメッセージの 1 行目の形（`tools\check_commit_message.ps1`。上の「コミットと Pull Request のタイトル」）
- `Signed-off-by` が付いていること（`tools\check_signoff.ps1`。上の「Signed-off-by」）

## テスト

```powershell
.\tests\run.ps1                 # 既定のテスト（Unit・Io・Meta）
.\tests\run.ps1 -Tag Office     # Excel・Word・PowerPoint が要るテスト
```

- **機能を足したり動きを変えたりしたら、同じ Pull Request でテストを書いてください。** 画面に出す文言や可否の判定（判断層）は特に厚くします。
- **カバレッジは下げません。** CI（`.\tests\run.ps1 -Ci`）は、カバレッジが `tests/coverage.baseline` の値を下回ると失敗します。下回ったらテストを足して戻し、上がったら同じ Pull Request で下限を上げてください。画面層は計測の対象外です。
- 画面を変えたときは、`tebunko.bat` で実際に起動して確かめてください。

テストの分け方と CI の中身は [docs/00_共通_3_テスト.md](../docs/00_共通_3_テスト.md) にあります。

## コードの決まり

`tests/meta/` のテストと、CI の静的解析（PSScriptAnalyzer。重大度 Error の指摘が 0 件であること）が機械的に確かめます。

- **スクリプト・XAML は BOM 付き UTF-8・CRLF** で保存します（Windows PowerShell 5.1 の前提）。
- **`scripts/shared/` は個々のツールを知らない**ようにします。ツール同士も互いを読み込みません。
- **判断層**（`*_view.ps1`・`index_name.ps1`・`search_query.ps1`・`text.ps1`）は画面（WPF）に触れません。画面に出す文言や可否の判定はここに置き、テストを書きます。
- 足したファイルは、読み込み口（`shared/shared.ps1`・`tebunko/lib.ps1`・`gui.ps1`・`indexer.ps1`）から読み込みます。
- **ネットワーク通信・動的なコード実行・レジストリの変更・第三者のライブラリは使いません。** 利用者が導入を審査するときの前提です（[docs/04_安全性.md](../docs/04_安全性.md)）。

ソースの構成は [docs/00_index.md](../docs/00_index.md) にあります。

## 個人情報を含めない

このリポジトリは公開しています。**履歴に一度でも入れたものは公開されたのと同じ**なので、コミットする前に確かめてください。

- テストデータや文書の例には、架空の名前（`山田` `佐藤`）、`example.com` のメールアドレス、架空の組織（`(株)山田商事`・`C:\共有\営業部`）を使います。
- 利用者名を含む絶対パス（`C:\Users\<利用者名>\...`）を書きません。例示が必要なら `C:\Users\test\...` とします。
- Office ファイルには、中身を見ただけでは分からない場所に作成者名などが埋め込まれます。テストデータを作り直したら、[tests/testdata/README.md](../tests/testdata/README.md) の手順で取り除いてください。
- コミットの作者のメールアドレスは GitHub の noreply アドレスにしてください。

## ドキュメント

- 設計書は `docs/` にあります。動きを変えたら、該当する設計書も同じ Pull Request で直してください。
- 図は Mermaid か draw.io（`.drawio.png`）で描きます。罫線文字のアスキーアートは使いません。
- 設計書は [Web サイト](https://hsgwa.github.io/tebunko/)（MkDocs）にもなります。設計書を足したときは `tools/mkdocs/mkdocs.yml` の `nav` にも足してください。設計書を変えた Pull Request には、サイトのプレビューの URL が自動でコメントされます。手元での確かめ方は [docs/00_共通_3_テスト.md](../docs/00_共通_3_テスト.md) の「CI」にあります。

## リリース

リリースはメンテナが行います。各版の変更点は [GitHub Release](https://github.com/hsgwa/tebunko/releases) にあります。

## ライセンス

Pull Request で提出した変更（コード・ドキュメント・テストデータを含む）は、本リポジトリと同じ [MIT License](../LICENSE) で公開されます。提出するのは、自分で作ったものか、MIT License で公開してよい権利を持つものに限ってください。第三者のコード・画像・文書をそのまま持ち込むときは、事前に Issue で相談してください（本ツールは第三者の部品を含まないことを前提にしています。[sbom.cdx.json](../sbom.cdx.json)）。
