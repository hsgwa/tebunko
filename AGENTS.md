# AGENTS.md

## 言語

利用者への返答・報告は日本語で書く（コードのコメント・ドキュメント・コミットメッセージも日本語）。

ただし `.github/` の CONTRIBUTING・SECURITY・SUPPORT・CODE_OF_CONDUCT は、英語版（`<名前>.md`）を正とし、日本語版（`<名前>.ja.md`）を並べて置く。GitHub が案内に使うのは英語版の名前のファイルのため。内容を変えるときは、両方を同じ PR で直す。日本語の文書（README・`docs/` など）からは日本語版へリンクする。README は内容をよく変えるため日本語だけにする。

## GitHub の運用

変更は **ブランチ → PR → main** の順で入れる。GitHub の操作は `gh` で行う。貢献の始め方（Issue の立て方・開発の準備）は [.github/CONTRIBUTING.ja.md](.github/CONTRIBUTING.ja.md) にある。

**ラベル付け・マージ・リリースはメンテナが行う。** エージェントはこれらを試みない。

- **main へは PR 経由でだけ入れる。** main への直接 push はブランチ保護（ruleset）で禁止し、必須チェック（`test.yml` の `test`・`title.yml` の `pr-title`・`docs.yml` の `docs`・`codeql.yml` の `analyze`）が通らないとマージできない。PR のブランチが最新の main を取り込んでいないときもマージできない。
- **1 つの PR には 1 つの目的だけを入れる。** 目的と関係のない修正は別の PR にする。
- **PR 本文は `.github/pull_request_template.md` に沿って書く。** Issue があれば `Closes #<番号>` でつなぐ（マージすると Issue が自動で閉じる）。無ければ目的を PR 本文に書く。
- **前の版と互換が無くなる PR は、タイトルの型に `!` を付ける**（下の Conventional Commits）。設定ファイル（`setting.config`）・インデックスの形式、起動の仕方、配布物のファイル構成が変わり、前の版のものがそのまま使えなくなるときがこれに当たる。PR 本文に移行の手順を書く。
- **PR を出す前に最新の main を取り込む。** 取り込みは merge で行い、push 済みのブランチを rebase して force push しない。

  ```
  git fetch origin
  git merge origin/main
  ```

- **マージは squash merge で行う**（GitHub の設定で squash だけを許している）。PR 1 つが main のコミット 1 つになり、**PR のタイトルがそのコミットのタイトルになる。** タイトルは、コミットメッセージと同じく変更の内容が分かる日本語の 1 行にする。
- **コミットメッセージの 1 行目・PR のタイトル・Issue のタイトルは Conventional Commits の形 `<型>(<範囲>)!: <説明>` にする。** 型は英語、説明は日本語で書く（例 `feat: Excel の図形の文字を検索できるようにする`）。範囲と `!`（前の版と互換が無くなる変更）は省略できる。

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

  機械的に確かめる: コミットは `commit-msg` フック、PR のタイトルは CI（`.github/workflows/title.yml` の `pr-title`。失敗するとマージできない）、Issue のタイトルは同じワークフローが形の違うものに直し方をコメントする。判定は `tools/check_commit_message.ps1` にまとめてある。git が自動で作るメッセージ（`Merge ...` `Revert "..."` `fixup! ...`）は調べない。
- **コミットには `git commit -s` で `Signed-off-by: <名前> <メールアドレス>` を付ける。** [DCO](https://developercertificate.org/)（その変更を出す権利があること）に同意したことを表す。メールアドレスはコミットの作者のもの（下の「個人情報を書かない」の noreply）と同じにする。

  機械的に確かめる: コミットは `commit-msg` フック、PR のコミットは CI（`test.yml` の `test`。失敗するとマージできない）。判定は `tools/check_signoff.ps1`。マージコミットと bot（Dependabot など）のコミットは調べない。決まりを作る前（#53 より前）のコミットには付いていないが、書き換えない。付け忘れたまま PR のブランチに push したときに限り、`git rebase --signoff origin/main` で付け直して `git push --force-with-lease` してよい。

## 除外設定（読ませない・検索させない）

Claude Code に `.claudeignore` は無い。除外は次の 2 か所で行う。

- **`.gitignore`** … ファイル検索（Glob・Grep）はここを見る。リポジトリに入れないものと、検索でも見せたくないものをここに書く。
  - `work/`（生成されるインデックス・ログ）、`sample*/`（手元の動作確認用データ）、`.claude/worktrees/`（作業用ツリー。別ツリーの同じコードが検索に二重に出るのを防ぐ）、`setting.config`（利用者ごとの設定）
- **`.claude/settings.json` の `permissions`** … ファイルを開くこと自体を止める。`Read(<パターン>)` は gitignore と同じ書き方で、`deny` は読み書きとも不可、`ask` は読む前に確認する。
  - `deny`: `sample*/`（大きいバイナリ）、`.env` / `*.pem` / `id_rsa*`（秘密情報の保険）
  - `ask`: `work/index/`（生成された大量の TSV。調べるときだけ読む）
  - 個人だけの設定は `.claude/settings.local.json`（git 管理外）に書く。

## 個人情報を書かない

本リポジトリは GitHub で公開する。次のものをコミットに含めない。作業ツリーだけでなく、**履歴に一度でも入れたら公開されたのと同じ**なので、コミットする前に確かめる。

- **実名・メールアドレス・アカウント ID。** テストデータの人名は `山田` `佐藤` などの架空名、メールアドレスは `example.com` のドメイン、ファイルの作成者名は `test` を使う。
- **利用者名を含む絶対パス**（`C:\Users\<利用者名>\...`）。スクリプトでは絶対パスを直接書かず `$PSScriptRoot` などから組み立てる。設定例のように書かざるを得ない場合は `C:\Users\test\...` とする。
- **実在の組織が分かる文字列**（社名・顧客名・サーバー名）。設計書や画面図の例は `(株)山田商事` `C:\共有\営業部` のような架空のものにする。

コミットの作者は GitHub の noreply アドレスにする。

```
git config user.email <ID>+<アカウント名>@users.noreply.github.com
```

**Office ファイルは、中身を見ただけでは分からない場所に作成者名・保存先の絶対パス・Microsoft アカウント ID が埋め込まれる。** `make_testdata.ps1` で作り直したテストデータをコミットする前に必ず取り除く。埋め込まれる場所の一覧は [tests/testdata/README.md](tests/testdata/README.md) の「生成」にある。

## ドキュメントの図

`docs/` 配下の設計書などに図を描くときは、以下のどちらかを使う。罫線文字や記号による AA（アスキーアート）の図は使わない。

- **Mermaid（基本）**: フロー図・シーケンス図・状態遷移図など、Mermaid で表せる図は Markdown 内に ```` ```mermaid ```` ブロックで書く。
  - ラベル内の `<` `>` は `#lt;` `#gt;` で書く。改行は `<br>`。
- **draw.io（`.drawio.png`）**: レイアウトの自由度が必要など Mermaid で表しにくい図は、図データを埋め込んだ PNG（`<図名>.drawio.png`）として `docs/images/` に置き、Markdown から相対パスで参照する。
  - 例: `![全体構成](images/全体構成.drawio.png)`
  - draw.io で再編集できるよう、「図のコピーを含める」を有効にして書き出す。

入出力の書式例・メッセージ例・ファイル名規則などの**図ではないテキスト**は、従来どおりコードブロックで書いてよい。

`docs/` は MkDocs で Web サイトにして GitHub Pages に公開する（`tools/mkdocs/mkdocs.yml`・`.github/workflows/docs.yml`）。設計書を足したら `tools/mkdocs/mkdocs.yml` の `nav` にも足す。リンク先のファイルや見出しが無いと CI（`docs`）が失敗するので、見出しを変えたらリンクも直す。`docs/` の外（`README.md`・`AGENTS.md`・`.github/` など）を含む全 Markdown のリンクは `tests/meta/links.Tests.ps1`（`tools/check_markdown_links.ps1`）が確かめ、切れているとコミット前のフックと CI（`test`）が失敗する。

## ソースの分け方

`scripts/` は **文脈**（`shared/` = どのツールからも使う、`tebunko_grep/` = このツール固有）と **層**（判断層・状態層・画面層）で分ける。詳細は [docs/00_index.md 1.4](docs/00_index.md) と [docs/00_共通_2_共通モジュール.md 5.0](docs/00_共通_2_共通モジュール.md)。

守ること（`tests/meta/` が機械的に確かめる）:

- `shared/` はツールを知らない。ツール同士も互いを読み込まない。
- 判断層（`*_view.ps1`・`index_name.ps1`・`search_query.ps1`・`text.ps1`）は `$ui` / `$window` / WPF の型に触らない。画面に出す文言や可否の判定はここに置き、テストを書く。
- 足したファイルは、必ず読み込み口（`shared/shared.ps1`・`tebunko_grep/lib.ps1`・`gui.ps1`・`indexer.ps1`）から読み込む。
- スクリプト・XAML は BOM 付き UTF-8・CRLF で保存する。

テストは `.\tests\run.ps1`（タグ `Unit` / `Io` / `Meta` / `Office` / `Slow`）。コミット前に通す。

## テストカバレッジの方針

全体の目標値は決めない。**下げないこと**と、**テストが効く層を厚くすること**を守る。数値の見方と CI での扱いは [docs/00_共通_3_テスト.md](docs/00_共通_3_テスト.md) の 6.3「タグと実行」と「CI」にある。

- **下限は `tests/coverage.baseline`。** `.\tests\run.ps1 -Ci` がこれを下回ると失敗し、CI の必須チェック `test` が通らないためマージできない。
  - 下回ったら、テストを足して戻す。下限の値を下げて通さない。
  - カバレッジが上がったら、同じ PR で下限を実測の値（小数点以下 1 桁）まで上げる。
- **判断層は 95% 以上を保つ。** 画面に出す文言や可否の判定を足したら、同じ PR でテストを書く。
- **状態層・I/O もテストを書く。** 特にインデックスや設定ファイルの読み書き・形式の移行のように、壊れると利用者のデータが使えなくなる処理は、層にかかわらずテストを書く。
- **画面層は計測の対象外にする。** 対象外のファイルは `tests/run.ps1` の `CodeCoverage` の条件で外す。画面層のファイルを足したら、この条件から外れているか（分母に入っていないか）を確かめる。
- **数字を上げるためだけのテストは書かない。** 実行するだけで結果を確かめないテストや、実装の細部に依存して壊れやすいテストは足さない。

## コミット前の検査と CI

clone したら `.\tools\install_hooks.ps1` を 1 回実行する（`core.hooksPath` を `tools/hooks` にする）。以後コミットのたびに次が動く。

- `tools/check_commit.ps1 -Staged` … 上の「個人情報を書かない」と文字コードの決まりを、ステージした内容で機械的に確かめる。
- `tests/run.ps1 -Tag Unit,Meta -Quiet` … 速いテスト。
- `tools/check_commit_message.ps1`（`commit-msg` フック）… コミットメッセージの 1 行目が上の「GitHub の運用」の形であること。
- `tools/check_signoff.ps1`（`commit-msg` フック）… 作者の `Signed-off-by` が付いていること。

CI（`.github/workflows/test.yml`）は全ファイル・全履歴の検査、既定のテスト、PSScriptAnalyzer を行い、カバレッジを Codecov に送る。`v` で始まるタグを push すると `release.yml` が配布 zip を作って GitHub Release に載せる。詳細は [docs/00_共通_3_テスト.md](docs/00_共通_3_テスト.md) の「CI」。

検査に引っかかったら、`--no-verify` で飛ばさずに内容を直す。例示用として通す名前・ドメインを増やすときは `tools/check_commit.ps1` の先頭で行う。