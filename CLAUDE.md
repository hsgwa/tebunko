# CLAUDE.md

## 言語

利用者への返答・報告は日本語で書く（コードのコメント・ドキュメント・コミットメッセージも日本語）。

## 作業場所

本リポジトリ配下での作業は git worktree（EnterWorktree 等）で隔離したツリー上で行い、作業ツリーを直接編集しない。

作業内容をメインブランチへマージしたら、その worktree は削除する。マージ済みで不要になった worktree を残さない。

```
git worktree remove .claude/worktrees/<名前>
git branch -d worktree-<名前>
```

未コミットの変更が残っている worktree は削除しない。コミットするか破棄するかを利用者に確認してから削除する。

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

## ソースの分け方

`scripts/` は **文脈**（`shared/` = どのツールからも使う、`windox_grep/` = このツール固有）と **層**（判断層・状態層・画面層）で分ける。詳細は [docs/00_index.md 1.4](docs/00_index.md) と [docs/00_共通_2_共通モジュール.md 5.0](docs/00_共通_2_共通モジュール.md)。

守ること（`tests/meta/` が機械的に確かめる）:

- `shared/` はツールを知らない。ツール同士も互いを読み込まない。
- 判断層（`*_view.ps1`・`index_name.ps1`・`search_query.ps1`・`text.ps1`）は `$ui` / `$window` / WPF の型に触らない。画面に出す文言や可否の判定はここに置き、テストを書く。
- 足したファイルは、必ず読み込み口（`shared/shared.ps1`・`windox_grep/lib.ps1`・`gui.ps1`・`convert.ps1`）から読み込む。
- スクリプト・XAML は BOM 付き UTF-8・CRLF で保存する。

テストは `.\tests\run.ps1`（タグ `Unit` / `Io` / `Meta` / `Office` / `Slow`）。コミット前に通す。