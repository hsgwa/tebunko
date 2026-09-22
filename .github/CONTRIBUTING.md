# 開発に参加する

tebunko への不具合の報告・要望・修正の提案を歓迎します。このページは、変更を提案するときの手順と決まりをまとめたものです。

- 使い方の質問は [SUPPORT.md](SUPPORT.md) を見てください。
- 安全性に関わる問題は、Issue ではなく [SECURITY.md](SECURITY.md) の手順で非公開で連絡してください。
- 参加する人は [行動規範](CODE_OF_CONDUCT.md) に従ってください。

## 変更の流れ

変更は **Issue → ブランチ → Pull Request → main** の順で入れます。

1. **Issue を立てる。** 不具合は「不具合」、機能の追加・変更は「機能の要望」のテンプレートを使います。タイトルは下の「コミットと Pull Request のタイトル」の形にします（テンプレートが `fix: ` `feat: ` を最初から入れます）。誤字の修正のような小さな変更は、Issue なしで Pull Request を出してかまいません。
2. **main から作業用のブランチを作る。**
3. **変更し、テストを通す**（下の「テスト」）。
4. **Pull Request を出す。** 本文はテンプレートに沿って書き、`Closes #<Issue の番号>` で Issue とつなぎます。ラベルを 1 つ付けます（`enhancement` / `bug` / `documentation` / `dependencies`）。前の版と互換が無くなる変更には `breaking` も付け、移行の手順を書きます。
5. **CI が通り、レビューが済んだら squash merge します。** Pull Request のタイトルが main のコミットのタイトルになるので、下の「コミットと Pull Request のタイトル」の形で、変更の内容が分かる 1 行にしてください。

1 つの Pull Request には 1 つの変更だけを入れます。関係のない修正は別の Pull Request にしてください。

### コミットと Pull Request のタイトル

コミットメッセージの 1 行目、Pull Request のタイトル、Issue のタイトルは [Conventional Commits](https://www.conventionalcommits.org/ja/v1.0.0/) の形にします。型は英語、説明は日本語で書きます。

```
<型>(<範囲>)!: <説明>

feat: Excel の図形の文字を検索できるようにする
fix(gui): 検索結果の件数が 0 のままになるのを直す
feat!: インデックスの形式を変える
```

- 範囲（`(gui)` など）は省略できます。書くときは空白を含まない 1 語にします。
- `!` は、設定ファイルやインデックスが前の版のまま使えなくなる変更に付けます。Pull Request には `breaking` ラベルも付けます（上げる桁は下の「リリース」）。
- 型の後ろは半角のコロンと空白 1 つです。

| 型 | 使うとき | Pull Request のラベル |
|---|---|---|
| `feat` | 機能の追加・変更 | `enhancement` |
| `fix` | 不具合の修正 | `bug` |
| `docs` | 文書だけの変更 | `documentation` |
| `refactor` | 動きを変えない書き直し | 内容に近いもの |
| `perf` | 速さの改善 | `enhancement` |
| `test` | テストだけの追加・修正 | 内容に近いもの |
| `style` | 書式だけの変更（空白・改行など） | 内容に近いもの |
| `build` | 配布物の作り方・依存の更新 | `dependencies`（依存の更新のとき） |
| `ci` | CI・git のフック・開発用の道具 | 内容に近いもの |
| `chore` | 上のどれにも当たらないもの | 内容に近いもの |
| `revert` | 前の変更の取り消し | 取り消す変更と同じもの |

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

画面を変えたときは、`tebunko.bat` で実際に起動して確かめてください。テストの分け方と CI の中身は [docs/00_共通_3_テスト.md](../docs/00_共通_3_テスト.md) にあります。

## コードの決まり

`tests/meta/` のテストが機械的に確かめます。

- **スクリプト・XAML は BOM 付き UTF-8・CRLF** で保存します（Windows PowerShell 5.1 の前提）。
- **`scripts/shared/` は個々のツールを知らない**ようにします。ツール同士も互いを読み込みません。
- **判断層**（`*_view.ps1`・`index_name.ps1`・`search_query.ps1`・`text.ps1`）は画面（WPF）に触れません。画面に出す文言や可否の判定はここに置き、テストを書きます。
- 足したファイルは、読み込み口（`shared/shared.ps1`・`tebunko_grep/lib.ps1`・`tebunko_diff/lib.ps1`・`tebunko/gui.ps1`・`indexer.ps1`・`differ.ps1`）から読み込みます。
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

管理者が `v<メジャー>.<マイナー>.<パッチ>`（SemVer）のタグを main に付けて push すると、テストを通したうえで配布 zip が GitHub Release に公開されます。

今は 0.x の版です。前の版以降にマージされた Pull Request のラベルで、上げる桁を決めます。

- **マイナー（0.3.1 → 0.4.0）** … `breaking`（設定ファイル `setting.config` やインデックスの形式が変わり、前の版のものがそのまま使えなくなる）を含むとき
- **パッチ（0.3.1 → 0.3.2）** … 機能の追加（`enhancement`）か不具合の修正（`bug`）だけのとき
- 文書（`documentation`）・依存の更新（`dependencies`）だけのときはリリースしません

1.0.0 以降は、`breaking` でメジャー、`enhancement` でマイナー、`bug` でパッチを上げます。

リリースノートは Pull Request のラベルで分類されます。

## ライセンス

Pull Request で提出した変更（コード・ドキュメント・テストデータを含む）は、本リポジトリと同じ [MIT License](../LICENSE) で公開されます。提出するのは、自分で作ったものか、MIT License で公開してよい権利を持つものに限ってください。第三者のコード・画像・文書をそのまま持ち込むときは、事前に Issue で相談してください（本ツールは第三者の部品を含まないことを前提にしています。[sbom.cdx.json](../sbom.cdx.json)）。
