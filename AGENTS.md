# AGENTS.md

## 言語

利用者への返答・報告は日本語で書く（コードのコメント・ドキュメント・コミットメッセージも日本語）。

## 作業場所

本リポジトリ配下での作業は git worktree（EnterWorktree 等）で隔離したツリー上で行い、作業ツリーを直接編集しない。

ブランチ（worktree）を作るときは、先に GitHub の main を取得し、その最新から作る。手元の `master` や古いブランチを起点にしない。

```
git fetch origin
git worktree add -b worktree-<名前> .claude/worktrees/<名前> origin/main
```

一つのセッションを続けて別の機能に取りかかるときは、今の worktree で続けるか、別の worktree を新しく作るかを利用者に確認してから始める。別の worktree にする場合も、上と同じく最新の `origin/main` から作る。

作業内容をメインブランチへマージしたら、その worktree は削除する。マージ済みで不要になった worktree を残さない。

squash merge では手元のコミットが main に入らないため、`git branch -d` はブランチを消せない。PR がマージ済み（`MERGED`）であることを確かめてから `-D` で消す。

```
gh pr view worktree-<名前> --json state --jq .state
git worktree remove .claude/worktrees/<名前>
git branch -D worktree-<名前>
```

未コミットの変更が残っている worktree は削除しない。コミットするか破棄するかを利用者に確認してから削除する。

## GitHub の運用

変更は **Issue → ブランチ → PR → main** の順で入れる。GitHub の操作は `gh` で行う。

- **Issue から始める。** 機能追加・不具合修正は、先に Issue を立てる（テンプレートは `.github/ISSUE_TEMPLATE/`）。誤字直しのような小さな変更は Issue なしで PR を出してよい。
- **main へは PR 経由でだけ入れる。** main への直接 push はブランチ保護（ruleset）で禁止し、必須チェック（`test.yml` の `test`・`title.yml` の `pr-title`・`docs.yml` の `docs`・`codeql.yml` の `analyze`）が通らないとマージできない。PR のブランチが最新の main を取り込んでいないときもマージできない。
- **1 つの PR には 1 つの機能だけを入れる。** 関係のない修正は別の PR にする。
- **PR 本文は `.github/pull_request_template.md` に沿って書き、`Closes #<番号>` で Issue とつなぐ。** マージすると Issue が自動で閉じる。
- **PR にはラベルを 1 つ付ける**（`enhancement` / `bug` / `documentation` / `dependencies`）。リリースノートはこのラベルで分類される（`.github/release.yml`）。
- **PR を出す前に最新の main を取り込む。** 取り込みは merge で行い、push 済みのブランチを rebase して force push しない。

  ```
  git fetch origin
  git merge origin/main
  ```

- **マージは squash merge で行う**（GitHub の設定で squash だけを許している）。PR 1 つが main のコミット 1 つになり、**PR のタイトルがそのコミットのタイトルになる。** タイトルは、コミットメッセージと同じく変更の内容が分かる日本語の 1 行にする。
- **コミットメッセージの 1 行目・PR のタイトル・Issue のタイトルは Conventional Commits の形 `<型>(<範囲>)!: <説明>` にする。** 型は英語、説明は日本語で書く（例 `feat: Excel の図形の文字を検索できるようにする`）。範囲と `!`（前の版と互換が無くなる変更。リリースではメジャーを上げる）は省略できる。

  | 型 | 使うとき | PR のラベル |
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

  機械的に確かめる: コミットは `commit-msg` フック、PR のタイトルは CI（`.github/workflows/title.yml` の `pr-title`。失敗するとマージできない）、Issue のタイトルは同じワークフローが形の違うものに直し方をコメントする。判定は `tools/check_commit_message.ps1` にまとめてある。git が自動で作るメッセージ（`Merge ...` `Revert "..."` `fixup! ...`）は調べない。
- マージしたブランチは GitHub が自動で消す。手元の worktree とブランチは上の「作業場所」の手順で消す。
- **GitHub Actions の更新は Dependabot が PR を出す**（`.github/dependabot.yml`）。CI が通れば、内容を見てマージする。

### リリース

`v<メジャー>.<マイナー>.<パッチ>`（SemVer）のタグを main のコミットに付けて push する。上げる桁は次で決める。

- **メジャー** … 設定ファイル（`setting.config`）やインデックスの形式が変わり、前の版のものがそのまま使えなくなるとき。
- **マイナー** … 前の版と互換のまま機能を足したとき。
- **パッチ** … 不具合の修正だけのとき。

```
git fetch origin
git tag v1.2.0 origin/main
git push origin v1.2.0
```

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

`docs/` は MkDocs で Web サイトにして GitHub Pages に公開する（`tools/mkdocs/mkdocs.yml`・`.github/workflows/docs.yml`）。設計書を足したら `tools/mkdocs/mkdocs.yml` の `nav` にも足す。リンク先のファイルや見出しが無いと CI（`docs`）が失敗するので、見出しを変えたらリンクも直す。

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

CI（`.github/workflows/test.yml`）は全ファイル・全履歴の検査、既定のテスト、PSScriptAnalyzer を行い、カバレッジを Codecov に送る。`v` で始まるタグを push すると `release.yml` が配布 zip を作って GitHub Release に載せる。詳細は [docs/00_共通_3_テスト.md](docs/00_共通_3_テスト.md) の「CI」。

検査に引っかかったら、`--no-verify` で飛ばさずに内容を直す。例示用として通す名前・ドメインを増やすときは `tools/check_commit.ps1` の先頭で行う。