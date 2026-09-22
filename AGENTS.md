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
- **main へは PR 経由でだけ入れる。** main への直接 push はブランチ保護で禁止し、CI（`test.yml` の `test`）が通らないとマージできない。
- **1 つの PR には 1 つの機能だけを入れる。** 関係のない修正は別の PR にする。
- **PR 本文は `.github/pull_request_template.md` に沿って書き、`Closes #<番号>` で Issue とつなぐ。** マージすると Issue が自動で閉じる。
- **PR にはラベルを 1 つ付ける**（`enhancement` / `bug` / `documentation` / `dependencies`）。リリースノートはこのラベルで分類される（`.github/release.yml`）。
- **前の版と互換が無くなる PR には、さらに `breaking` を付ける。** 設定ファイル（`setting.config`）・インデックスの形式、起動の仕方、配布物のファイル構成が変わり、前の版のものがそのまま使えなくなるときがこれに当たる。PR 本文に移行の手順を書く。
- **PR を出す前に最新の main を取り込む。** 取り込みは merge で行い、push 済みのブランチを rebase して force push しない。

  ```
  git fetch origin
  git merge origin/main
  ```

- **マージは squash merge で行う**（GitHub の設定で squash だけを許している）。PR 1 つが main のコミット 1 つになり、**PR のタイトルがそのコミットのタイトルになる。** タイトルは、コミットメッセージと同じく変更の内容が分かる日本語の 1 行にする。
- マージしたブランチは GitHub が自動で消す。手元の worktree とブランチは上の「作業場所」の手順で消す。
- **GitHub Actions の更新は Dependabot が PR を出す**（`.github/dependabot.yml`）。CI が通れば、内容を見てマージする。

### マージ

PR のマージはエージェントが行ってよい。次をすべて満たしたら、利用者の確認を待たずにマージする。

1. 最新の main を取り込んである（上の `git merge origin/main`）。
2. CI がすべて通っている（`gh pr checks <番号>`）。失敗・実行中のものが 1 つでもあればマージしない。
3. PR テンプレートの「確認したこと」をすべて済ませた。画面を変えた場合は、実際に起動して見た。
4. 差分を読み直し、PR の目的と関係のない変更・個人情報が入っていない。
5. ラベルが付いている。

```
gh pr merge <番号> --squash
```

- **`breaking` の PR は、利用者の了承を得てからマージする。** 移行の手順を添えて了承を求める。
- `--admin` でブランチ保護を飛ばさない。
- マージしたら、上の「作業場所」の手順で worktree とブランチを消す。続けて下の「リリース」の要否を判断する。

### リリース

`v<メジャー>.<マイナー>.<パッチ>`（SemVer）のタグを main のコミットに付けて push すると、`release.yml` が配布 zip を GitHub Release に載せる。リリースもエージェントが行ってよい。

**今は 0.x（`v0.<マイナー>.<パッチ>`）を続ける。** 1.0.0 に上げるのは利用者が決める。0.x の間、上げる桁は前のタグ以降にマージされた PR のラベルで次のように決める。上から順に見て、最初に当てはまったものにする。

| マージされた PR のラベル | 0.x の間 | 1.0.0 以降 |
|---|---|---|
| `breaking` が 1 つでもある | マイナー（0.3.1 → 0.4.0） | メジャー（1.3.1 → 2.0.0） |
| `enhancement` がある | パッチ（0.3.1 → 0.3.2） | マイナー（1.3.1 → 1.4.0） |
| `bug` がある | パッチ（0.3.1 → 0.3.2） | パッチ（1.3.1 → 1.3.2） |
| `documentation` / `dependencies` だけ | リリースしない | リリースしない |

リリースする時機:

- **`bug` の PR をマージしたら、すぐにパッチを出す。** 不具合で困っている利用者に早く届けるため。
- **`enhancement` は、利用者に頼まれた一連の作業の PR をすべてマージしてから、まとめて 1 回出す。** PR ごとに版を分けない。
- **`breaking` を含むリリースは、利用者の了承を得てから出す。** 互換を壊す変更はなるべく 1 回のリリースにまとめる。

手順:

```
git fetch origin --tags
$prev = git describe --tags --abbrev=0 origin/main        # 前の版
$since = git log -1 --format=%cI $prev                     # その日時
gh pr list --state merged --base main --search "merged:>$since" --json number,title,labels
```

一覧のラベルで次の版を決め、タグを付けて push する。push したら `gh run watch` で `release` の成功を確かめ、利用者に版と含めた PR の番号を報告する。

```
git tag v0.1.0 origin/main
git push origin v0.1.0
```

- タグは一度 push したら付け直さない。間違えたら、直した版を次の番号で出す。

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

`docs/` は MkDocs で Web サイトにして GitHub Pages に公開する（`mkdocs.yml`・`.github/workflows/docs.yml`）。設計書を足したら `mkdocs.yml` の `nav` にも足す。リンク先のファイルや見出しが無いと CI（`docs`）が失敗するので、見出しを変えたらリンクも直す。

## ソースの分け方

`scripts/` は **文脈**（`shared/` = どのツールからも使う、`tebunko_grep/` = このツール固有）と **層**（判断層・状態層・画面層）で分ける。詳細は [docs/00_index.md 1.4](docs/00_index.md) と [docs/00_共通_2_共通モジュール.md 5.0](docs/00_共通_2_共通モジュール.md)。

守ること（`tests/meta/` が機械的に確かめる）:

- `shared/` はツールを知らない。ツール同士も互いを読み込まない。
- 判断層（`*_view.ps1`・`index_name.ps1`・`search_query.ps1`・`text.ps1`）は `$ui` / `$window` / WPF の型に触らない。画面に出す文言や可否の判定はここに置き、テストを書く。
- 足したファイルは、必ず読み込み口（`shared/shared.ps1`・`tebunko_grep/lib.ps1`・`gui.ps1`・`indexer.ps1`）から読み込む。
- スクリプト・XAML は BOM 付き UTF-8・CRLF で保存する。

テストは `.\tests\run.ps1`（タグ `Unit` / `Io` / `Meta` / `Office` / `Slow`）。コミット前に通す。

## コミット前の検査と CI

clone したら `.\tools\install_hooks.ps1` を 1 回実行する（`core.hooksPath` を `tools/hooks` にする）。以後コミットのたびに次が動く。

- `tools/check_commit.ps1 -Staged` … 上の「個人情報を書かない」と文字コードの決まりを、ステージした内容で機械的に確かめる。
- `tests/run.ps1 -Tag Unit,Meta -Quiet` … 速いテスト。

CI（`.github/workflows/test.yml`）は全ファイル・全履歴の検査、既定のテスト、PSScriptAnalyzer を行い、カバレッジを Codecov に送る。`v` で始まるタグを push すると `release.yml` が配布 zip を作って GitHub Release に載せる。詳細は [docs/00_共通_3_テスト.md](docs/00_共通_3_テスト.md) の「CI」。

検査に引っかかったら、`--no-verify` で飛ばさずに内容を直す。例示用として通す名前・ドメインを増やすときは `tools/check_commit.ps1` の先頭で行う。