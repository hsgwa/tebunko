# テストの実行と CI

テストは `scripts/` と同じ構成に並べる。どのモジュールのテストがどこにあるか、たどらなくても分かるようにするため。

| フォルダ | 内容 |
|---|---|
| `tests/helpers/` | 共通の準備（`load.ps1`。`$here`・`${scriptsDir}`・`${testDataDir}` を決めて `tebunko/lib.ps1` を読み込む）とテスト用の TSV 作成（`tsv.ps1`） |
| `tests/shared/core/` | `fs`・`text`・`folder`・`worker_pool`・`data_dir` |
| `tests/shared/office/` | `office_process`・`office_reader`・`office_app` |
| `tests/shared/ui/` | 画面の型（`types`） |
| `tests/tebunko/core/` | `settings`（`setting.config`）・`workspace` |
| `tests/tebunko/index/` | `index_name`・`index_store`・`pack_format`・`system_index` |
| `tests/tebunko/indexer/` | `indexer_state`・`indexer_decide`・`indexer_plan`・`extract_office`・`index_migrate`・`indexing_session`、起動口の通しのテスト（`indexer`） |
| `tests/tebunko/search/` | `search_query`・`search_run`・`pack_search`・`search_service`・`source_map`・高速検索（`search_gram`・`fast_search`・`windows_search`） |
| `tests/tebunko/ui/` | 画面の判断層（`index_view`・`indexing_view`・`search_view`・`preview_view`・`settings_view`）と、`$ui` を偽物にした画面の部品（`result_list`・`open_source`・`preview`・`index_tree`）・型（`types`） |
| `tests/tools/` | 開発用の道具（`check_commit_message`・`check_signoff`・`check_markdown_links`・`measure_perf`・`run_commit_tests`） |
| `tests/meta/` | 構成を守るテスト（`structure`・`encoding`・`layers`・`links`・`runner`・`classes`）と安全性の検査（`safety`・`installer`） |
| `tests/testdata/` | 手動の結合テスト用のデータ（[結合テスト（手動）](index.md#結合テスト手動)）と、その生成（`make_testdata.ps1`）・個人情報の除去（`scrub_personal`） |

入力と期待値だけが違うテストは、`-TestCases` の 1 つの `It` にまとめる（例: `tests/tebunko/ui/types.Tests.ps1` の `HitRow.Contains`）。表は `It` の中にそのまま書き、計算で作らない。キーには `input`・`args`・`_`・`Matches` など PowerShell の自動変数の名前を使わない。各行に `name` を持たせ、`It "<name>"` で失敗した行が分かるようにする。表の中で変数（`$stateDone` など）を使うときは、`BeforeDiscovery` で用意する（表はテストを探す段階で作られ、`BeforeAll` より先に評価されるため）。

Pester 5 はテストを「探す段階」と「流す段階」に分けて動かす。探す段階では `Describe` の中身がそのまま実行され、`BeforeAll` の中身は流す段階まで実行されない。そのため次のように書く。

- ファイルの先頭の読み込み（`. "$PSScriptRoot\..\helpers\load.ps1"`・`Add-Type`）と、テスト用の関数・偽物のオブジェクトは、いちばん外の `BeforeAll { }` に書く。
- `Describe` の準備（`$TestDrive` へのファイル作成・変数・`Mock`）は、その `Describe` の `BeforeAll` / `BeforeEach` に書く。探す段階では `$TestDrive` が空のため、`Describe` の中に直接書くとドライブ直下を指してしまう。
- `Mock` はそれを書いたブロック（`It`・`Describe`）の中だけで効く。呼ばれたかは `Should -Invoke` で確かめる。

`tests/meta/` は、直し忘れると気づきにくい決まりごとを機械的に確かめる。

| テスト | 確かめること |
|---|---|
| `structure` | `$rootDir` などのパス定義、全 `.ps1` の構文、`ui/` のスクリプトで `$PSScriptRoot` を使わない、`gui.ps1` が指すインデクサのファイルがある、XAML が XML として読める・`gui.ps1` が使う `x:Name` がある、型を読み込む順で継承が解決できる |
| `encoding` | 全 `.ps1`・`.xaml` が BOM 付き UTF-8 で、改行が CRLF。`.github/codecov.yml` が ASCII の文字だけ |
| `layers` | `shared/` にツールの名前が出てこない、ツール同士が互いを読み込まない、起動口からたどれない `.ps1` が無い、判断層（`text.ps1`・`index_name.ps1`・`search_query.ps1`・`indexer_decide.ps1`・`*_view.ps1`）に画面への依存が無い |
| `links` | git で管理している全 `.md` の相対リンク（画像・参照リンクの定義・HTML の `href`/`src` を含む）の先のファイルがあり（大文字・小文字も区別する）、`.md` のアンカーの見出しがある（`tools/check_markdown_links.ps1`。外部の URL は調べない） |
| `runner` | `tests/run.ps1` が、実行したテストが 0 件なら失敗にすること、`powershell.exe -File` で渡したカンマ区切りのタグを分けて受け取ること |
| `safety` | 危険な処理を使っていない、Office をマクロ無効・読み取り専用で開く、原本を書き換えない、書き込み先が `work`・`%TEMP%` だけ、PSScriptAnalyzer の指摘が 0 件、審査用の資料がそろっている（[単体テスト](index.md#単体テスト)、[安全性の要約](../../safety/index.md)） |

## タグと実行

`Describe` にタグを付け、実行するものを選べるようにする。

| タグ | 内容 | 外部ソフト | 既定で実行 |
|---|---|---|---|
| `Unit` | ファイルに触らないもの（判断層、画面の部品を偽物にしたもの） | 不要 | する |
| `Io` | ファイルの読み書き（`$TestDrive` の中で完結する） | 不要 | する |
| `Meta` | 構成を守るテスト・安全性の検査 | 不要（PSScriptAnalyzer があれば静的解析も行う） | する |
| `Office` | Excel・Word・PowerPoint の COM を実際に動かすもの（今は該当するテストが無い。COM は `Mock` で確かめる） | 必要 | しない（`-All` で実行） |
| `Slow` | 時間のかかるもの（今は該当するテストが無い） | 不要 | しない（`-All` で実行） |
| `Manual` | 手で確かめるもの（今は該当するテストが無い） | – | しない（`-All` でも実行しない） |

実行は `tests/run.ps1` から行う。

```
.\tests\run.ps1              既定（Unit・Io・Meta。Office・Slow・Manual は外す）
.\tests\run.ps1 -Tag Unit    速い確認だけ
.\tests\run.ps1 -All         Office・Slow も含める（Office が必要）
.\tests\run.ps1 -Ci          結果の XML（work\test\results.xml）とカバレッジ（work\test\coverage.xml）を出し、カバレッジの下限を確かめる
.\tests\run.ps1 -Path .\tests\shared\core   指定したフォルダ・ファイルのテストだけ
.\tests\run.ps1 -Quiet       失敗したテストだけを表示する
```

いずれも失敗したテストの数を終了コードにする（フックと CI が見る）。1 件も実行しなかったときも失敗にする（終了コード 1）。タグの打ち間違いで、何も確かめないまま通るのを防ぐため。

`-Tag`・`-ExcludeTag` はカンマ区切りの文字列でも受け取る。`powershell.exe -File` で呼ぶと `-Tag Unit,Meta` は配列にならず 1 つの文字列で渡るため（pre-commit フックがこの呼び方）。

**カバレッジ**

- 対象は `scripts/` の `.ps1` のうち、画面層の `gui.ps1`・`*_tab.ps1`・`shell.ps1`・`app_host.ps1`・`*_dialog.ps1` を除いたもの（`tests/run.ps1` の `CodeCoverage` の条件）。除いたものは自動テストの対象外で、手で確かめる。画面層のファイルを足したら、この条件から外れているか（分母に入っていないか）を確かめる。
- 値は Pester のコマンド単位（実行されたコマンドの数 ÷ 全コマンドの数。小数点以下 1 桁）。
- **下限は `tests/coverage.baseline`（90.0）。** `-Ci` はこれを下回ると失敗し、CI の必須チェック `test` が通らない。下回ったらテストを足して戻し、下限は下げない。実測が上がっても下限は上げない（変えるのはメンテナだけ）。
- 全体の目標値は決めない。判断層は 95% 以上を保ち、数字を上げるためだけのテスト（結果を確かめないもの）は書かない（AGENTS.md「テストカバレッジの方針」）。
- カバレッジはブレークポイントを使わない計測（Pester の Profiler。`CodeCoverage.UseBreakpoints = $false`）で取る。ブレークポイントを使う計測より速い。
- `-Ci` はカバレッジを Cobertura 形式の XML（`work\test\coverage.xml`）にも書き出す。Pester が書き出す XML には利用者名を含む絶対パスが入るため使わず、`tests/run.ps1` が結果（コマンド単位）から行単位に組み立てる。1 行に実行されたコマンドが 1 つでもあれば、その行は通ったものとする。そのため、行単位の値（Codecov の表示）はコマンド単位の値と少し違う。ファイル名はリポジトリからの相対パスで書き、利用者名を含む絶対パスは入れない。

## CI

GitHub Actions のワークフローは次のとおり。使うアクションは、どのワークフローでもコミットのハッシュで固定する（版はハッシュの後ろのコメントに書き、Dependabot が更新する）。

| ワークフロー | ジョブ（チェック名） | 動く時 | 内容 | main の必須チェック |
|---|---|---|---|---|
| `test.yml` | `test` | PR、main への push、`release.yml` からの呼び出し | 個人情報・文字コードの検査、`Signed-off-by` の検査（PR のみ）、テスト（`run.ps1 -Ci`）、Codecov への送信、PSScriptAnalyzer | ○ |
| `title.yml` | `pr-title`・`issue-title` | PR・Issue の作成と編集（PR は push でも） | タイトルが Conventional Commits の形か | ○（`pr-title`） |
| `docs.yml` | `docs`（ほかに `changes`・`publish`） | PR、main で設計書が変わったとき | 設計書のサイトを作り、GitHub Pages に公開する | ○ |
| `codeql.yml` | `analyze` | PR、main への push、毎週 1 回 | ワークフローの静的解析 | ○ |
| `scorecard.yml` | `analysis` | main への push、ブランチ保護の変更、毎週 1 回 | OpenSSF Scorecard の採点 | – |
| `release.yml` | `test`・`release` | `v` で始まるタグの push | テストのうえ、配布 zip とインストーラーを GitHub Release に載せる | – |
| `perf.yml` | `perf` | 手動（`workflow_dispatch`） | pack の作成・検索の速さとリソースの推移を測る | – |

**`test.yml`**

pull request と main への push のたびに windows ランナーで実行する。作業ブランチへの push だけでは動かない（PR のブランチで同じテストが 2 回走らないようにするため）。PR を出す前に CI で確かめたいときは、下書き（draft）の PR を出す。

- Windows PowerShell 5.1 はランナーに最初から入っている（`shell: powershell` を明示する。`pwsh`（PowerShell 7）では COM と文字コードの扱いが変わる）
- ランナーには Windows に最初から入っている Pester 3.4 もある。`tests/run.ps1` は `Import-Module Pester -RequiredVersion 5.9.0` で版を指定し、CI は 5.9.0 が無ければ入れる（版は `test.yml` の `PESTER_VERSION` と `tests/run.ps1` の 2 か所で同じにする）
- スクリプトの改行はランナーの `core.autocrlf` に左右されないよう、`.gitattributes` で `.ps1`・`.xaml`・`.bat` を CRLF に固定している
- ランナーに Office は入っていないため、タグ `Office` のテストは既定で外れる。COM を使うインデックス作成の確認は手元で行う（[結合テスト（手動）](index.md#結合テスト手動)）
- あわせて PSScriptAnalyzer の `Error` を 0 件に保つ（`TypeNotFound` は除く）。PSScriptAnalyzer の版は `test.yml` で固定する（新しい版でルールが増えて、コードを変えずに CI が落ちるのを防ぐ。Dependabot の対象外のため、上げるときは版を書き換えた PR で指摘が 0 件のままか確かめる）。安全性の検査（`tests/meta/safety.Tests.ps1`）はタグ `Meta` のため既定で実行される
- テストの前に `tools/check_commit.ps1` で、追跡している全ファイルと全コミットの作者を検査する（下の「公開してはいけない内容の検査」）。作者を全履歴で調べるため `fetch-depth: 0` で取得する
- PR では、PR のコミットに作者の `Signed-off-by` があるかを `tools/check_signoff.ps1 -Range HEAD^1..HEAD^2` で確かめる。pull_request では PR をマージした形のコミットを取り出すため、1 つ目の親が main、2 つ目の親が PR の先頭になる。その間だけを調べるので、main から取り込んだコミットと、決まりを作る前の履歴（`Signed-off-by` が無い）は調べない。マージコミットと bot（作者が `...[bot]@users.noreply.github.com`）のコミットも調べない
- 同じブランチに続けて push したときは、古い実行を取り消す（`concurrency`）
- テスト結果（`work/test/`）は成否にかかわらず成果物として保存する
- テストのあと、`work\test\coverage.xml` を [Codecov](https://codecov.io/gh/hsgwa/tebunko) に送る（`codecov/codecov-action`）。README のバッジ、PR へのコメント（ファイルごとの増減）、行ごとの表示は Codecov が出す
  - 送るのは相対パスと行ごとの実行回数だけで、ソースは送らない（Codecov はソースを GitHub から読む）
  - トークンはリポジトリの Secrets の `CODECOV_TOKEN` に置く。fork からの PR ではトークンが無くても送れる
  - 送れなくても CI は失敗にしない。Codecov の判定（`.github/codecov.yml`）も参考表示だけにする。下限の確認は `tests/coverage.baseline` が受け持つ
  - `.github/codecov.yml` には ASCII の文字だけを書き、コメントも書かない。Windows で動く Codecov の CLI がこのファイルを cp1252 として読み、日本語があると `UnicodeDecodeError` で止まるため。`tests/meta/encoding.Tests.ps1` が確かめる
  - タグの push（`release.yml` から呼ばれたとき）では送らない

**`title.yml`**

PR と Issue のタイトルを `tools/check_commit_message.ps1 -Title` で確かめる。squash merge では PR のタイトルが main のコミットのタイトルになるため、main の履歴の形はここで決まる。

- PR（ジョブ `pr-title`）… 形が違えば失敗にする。ブランチ保護の必須のチェックにしてあり、失敗するとマージできない。必須のチェックは head のコミットごとに要るため、タイトルの編集だけでなく push でも動かす
- Issue（ジョブ `issue-title`）… 作成は止められないため、形が違えば `.github/title_comment.md` の直し方を 1 回だけコメントする（1 行目の目印が付いたコメントが既にあれば書かない）
- タイトルは誰でも書ける信頼できない入力のため、式で `run` に埋め込まず環境変数で渡す
- 起動の速い ubuntu のランナーで `pwsh`（PowerShell 7）を使う。そのため `tools/check_commit_message.ps1` は 5.1 と 7 の両方で動くように書く
- Dependabot の PR も同じ形にする（`.github/dependabot.yml` の `commit-message`。`ci(deps): bump ...` `build(deps): bump ...`）

**`codeql.yml`・`scorecard.yml`（サプライチェーンの安全性）**

CodeQL（`analyze`）は main の必須チェックで、指摘があるとマージできない。Scorecard は採点するだけで、マージは止めない。

- `codeql.yml` … GitHub Actions のワークフローを CodeQL で解析し、Security タブ（Code scanning）に載せる。信頼できない入力を `run` に埋め込む書き方などを見つける。PowerShell は CodeQL の対象外のため、スクリプトは上の PSScriptAnalyzer が受け持つ
- `scorecard.yml` … [OpenSSF Scorecard](https://scorecard.dev/viewer/?uri=github.com/hsgwa/tebunko) で採点し、結果を Code scanning と README のバッジに載せる
  - Scorecard はリポジトリを Linux で展開する。名前が UTF-8 で 255 バイト（日本語で 85 字）を超えるファイルがあると展開できず、ほとんどの項目が採点できなくなる。そのため `tools/check_commit.ps1` がこの長さを超える名前を止める

**`release.yml`（配布物の公開）**

`v` で始まるタグを push すると動き、`test.yml` と同じ検査・テストを通したうえで、`tools/new_release_package.ps1` で配布 zip（`tebunko-<タグ>.zip`）を、`tools/new_installer.ps1` でインストーラー（`tebunko-setup-<タグ>.exe`）を作って GitHub Release に載せる。

- zip にはツール本体（`tebunko.bat`・`scripts/`）と `README.md`・`LICENSE` だけを入れる。README の相対リンクと画像は、その版の GitHub の URL に書き換える。カタログ（`tebunko.cat`）・ハッシュ一覧（`SHA256SUMS.txt`）・部品表（`sbom.cdx.json`）は zip と並べてリリースに載せ（[安全性の要約](../../safety/index.md) の [配布物の完全性（カタログ・ハッシュ一覧・来歴の署名）](../../safety/scans.md#配布物の完全性カタログハッシュ一覧来歴の署名)）、SECURITY は README とリリースの説明からリンクする。zip 自体の SHA256 はリリースの説明に書く（同 [複数エンジンでの検査: VirusTotal（外部へファイルを送信する）](../../safety/scans.md#複数エンジンでの検査-virustotal外部へファイルを送信する) の VirusTotal での照会用）
- インストーラーは Inno Setup 7 で作る（6.7.1 は、Program Files に入れたものを消すとアンインストーラーが残ったため 7.1.0 にした）。Inno Setup は版を固定して公式のリリースから取り、SHA256 を確かめてから、持ち運び版（レジストリに書かない）でランナーの一時フォルダに入れる。版を上げるときは `release.yml` の URL と SHA256 を一緒に直す（Dependabot の対象外）。インストーラーの SHA256 もリリースの説明に書く
- zip とインストーラーのビルドの来歴を Sigstore で署名して GitHub に登録し、署名の bundle（`tebunko-<タグ>.zip.sigstore.json`・`tebunko-setup-<タグ>.exe.sigstore.json`）もリリースに載せる（同 [配布物の完全性（カタログ・ハッシュ一覧・来歴の署名）](../../safety/scans.md#配布物の完全性カタログハッシュ一覧来歴の署名)）
- リリースノートは GitHub が PR から作り、`.github/release.yml` で PR のラベルごとに分ける
- 版の付け方とリリースの時機は AGENTS.md の「リリース」に従う。タグは main のコミットに付ける

```
git tag v0.1.0 origin/main
git push origin v0.1.0
```

**`docs.yml`（設計書のサイト）**

設計書（`docs/`）は [MkDocs](https://www.mkdocs.org/)（テーマは Material for MkDocs）で Web サイトにし、main で `docs/`・`tools/mkdocs/` が変わるたびに [GitHub Pages](https://hsgwa.github.io/tebunko/) に公開する。

- 設定は `tools/mkdocs/mkdocs.yml`。サイトはどれもリポジトリの直下で `mkdocs build --strict -f tools/mkdocs/mkdocs.yml` を実行して作る。リンク先のファイルや見出し（アンカー）が無いと失敗する。`docs/` の外の Markdown と、`docs/` から外へのリンクはここでは調べないため、`tests/meta/` の `links`（[テストの実行と CI](ci.md)）が調べる
- サイトを作るジョブ（チェックの名前は `docs`）は main の必須チェックにしている。必須のチェックはワークフローが動かないと待ったままになるため、PR では変えたファイルで絞らず、どの PR でもサイトを作る
- GitHub Pages は `gh-pages` ブランチの内容をそのまま公開する。main のサイトはその直下に置く
- `docs/`・`tools/mkdocs/` を変えた PR では（変えたかどうかは `changes` のジョブが PR のファイルの一覧で調べる）、サイトのプレビューを `gh-pages` の `pr-preview/pr-<PR の番号>/` に置き、URL を PR にコメントする（`rossjrw/pr-preview-action`）。push のたびに更新し、PR を閉じたら消す。fork からの PR ではプレビューを作らず、サイトを作れることだけを確かめる
- サイトを作るジョブは PR のコード（`tools/mkdocs/hooks.py`）を動かすため、読み取りの権限だけで動かす。`gh-pages` への書き込みと PR へのコメントは、作ったサイトを受け取るだけの別のジョブ（`publish`）で行う
- アンカーは GitHub と同じ形（日本語を残し、記号を除き、英字を小文字にする）で作る（`tools/mkdocs/mkdocs.yml` の `toc.slugify`）。`docs/` の中のリンクは GitHub で読めるように書けば、サイトでもそのまま飛べる
- GitHub で読むときとの違いは `tools/mkdocs/hooks.py` で埋める。先頭ページは `docs/index.md`（`tools/mkdocs/overrides/home.html` で組み立てる）。`docs/` の外（`../.github/SECURITY.md` など）へのリンクは GitHub 上のファイルへのリンクに書き換える
- 使うパッケージは `tools/mkdocs/requirements.txt` に版とハッシュで固定し、`pip install --require-hashes` で入れる。更新は Dependabot が PR を出す。サイトを作るときだけに使い、配布物には含めない
- サイトを見る人のブラウザーが第三者のサーバーへ通信しないようにする。フォントは Google Fonts を使わず OS のもので表示する（`theme.font: false`）。Mermaid は Material テーマが既定では unpkg から読み込むため、`extra_javascript` で版を固定して先に読み込ませ、`privacy` プラグインがビルドのときに取り込んでサイトに同梱する（取り込んだファイルは `work/cache/privacy/` に置く）。Mermaid の版は、Material テーマが読み込む版（`mermaid@11`）に合わせる
- 手元で見るときは次のとおり（生成物は `work/site/`）

```
pip install --require-hashes -r tools/mkdocs/requirements.txt
mkdocs serve -f tools/mkdocs/mkdocs.yml
```

**`perf.yml`（性能とリソースの計測）**

手動で起動し、決まった量のデータで、インデックス作成と検索の速さ、およびその間のリソース（メモリ・CPU・スレッド・ハンドル・GC）を測る。大きな変更の前後で数字を比べるときなどに使う。必須のチェックではない。

```
gh workflow run perf.yml -f ref=<測る ref> -f scale=0.1
```

- 入力: `ref`（測る ref。作業中のブランチも測れる）、`scale`（データの量。`1` でブック 5.4 万・TSV 16 万・約 1GB）、`count`（語ごとに続けて検索する回数。既定 20）
- 計測スクリプト（`tools/measure_perf.ps1` と `tools/perf/`）はワークフローの ref から、計測対象のコードは入力の ref からチェックアウトする。計測スクリプトが呼ぶ関数（`publishIndexFolders`・`getIndexPackFiles`・`searchPackIndex`）が無い ref（pack 形式より前の版）では、エラーメッセージを出して止まる
- データは [tebunko-perfdata](https://github.com/hsgwa/tebunko-perfdata) の `new_index.ps1` で毎回生成する（取り込みの一時置き場と同じ形の TSV）。使う版は `perf.yml` の `PERFDATA_SHA` でコミットに固定する。検索する語は同じリポジトリの `words.tsv`（0 件・まれ（3 件）・大量・正規表現）
- 測るもの（それぞれ別のプロセスで動かし、リソースが混ざらないようにする）
  - インデックス作成 … インデクサと同じ `publishIndexFolders`（pack を書く・TSV を消す・システムインデックスを作る）。Office からの取り込みは、ランナーに Office が無いので測らない。画面ではインデックス作成のスレッドの優先度を下げている（BelowNormal）が、計測ではほかに動くものが無いので、優先度は下げずに測る
  - 検索 … 語ごとに新しいプロセスを起動し、同じプロセスで `count` 回続けて検索する（上限 1 万件）。画面と同じく、検索の司令のスレッド（`newSearchService`）を 1 つ作り、要求（`newSearchRequest`）を順に送る。`lib.ps1` の読み込みと照合のスレッドの用意は司令のスレッドが始めに 1 回だけ行うので、1 回目の時間に含まれる。検索のキャッシュも画面と同じくプロセスで 1 つを持ち続ける。語ごとに、検索時間の最小・中央値・平均・最大と、その内訳（列挙・照合）を出す。1 回目は画面を開き直した直後の検索に当たり、たいてい最大の値になる。高速検索は、ランナーに Windows Search が無いので使わない
    - 司令のスレッドが無い版（#102 より前）は、その版の画面と同じく、検索のたびに新しい Runspace で `lib.ps1` を読み込んで測る（内訳に読み込みの時間も出す）。どちらで測ったかは結果の `SearchMode`（`service` / `runspace`）に出るので、#102 の前後を同じワークフローで比べられる
  - リソース … 別のスレッドで 0.2 秒ごと（および段階の切り替わり）に、プロセスのワーキングセット・プライベート・マネージドヒープ、CPU 時間、スレッド・ハンドルの数、GC の回数と、PC 全体の CPU・空きメモリを記録する。検索では、検索 1 回ごとのメモリも記録し、10 回あたりの増え方（最小二乗の傾き）を出す（検索を続けたときにメモリが増え続けないかを見るため）
- OS のファイルキャッシュは空にしない。pack は作成の直後なので OS のキャッシュに載っており、PC を起動した直後の検索（ディスクから読む）とは違う
- 結果は、ジョブの Summary（表と Mermaid のグラフ）と、artifact `perf-<実行の番号>`（保存期間は 90 日）に出力する。パスやファイル名は出力しない
  - `summary.md` … Summary と同じ内容
  - `result.json` … すべての数字。形式の版（`Schema`）と実行の情報（実行の番号・ref・コミット・scale・日時・ランナーの CPU）を含む
  - `metrics.csv` … 1 行 1 指標の縦長の形（`run_id, date, ref, sha, scale, metric, word, stat, value, unit`）。実行をまたいで数字を貯め、Grafana などで見るときに使う（貯める仕組みはまだ無い）
  - `searches.csv`（検索 1 回ごと）・`resource-index.csv`・`resource-search.csv`（リソースの記録）
- ランナー（4 コア）は手元の PC と条件が違う。Defender のリアルタイム保護の状態は、結果の「環境」に出力する。比べるときは、同じワークフローで測った数字どうしで比べる。同じ条件でも、pack の作成の時間は 5 割ほどぶれることがある
- 手元でも同じ計測スクリプトで測れる（`.\tools\measure_perf.ps1 -Index <TSV のフォルダ> -Work <作業フォルダ> -Words <words.tsv>`）。渡した TSV は pack に変換され、元の TSV は削除される

## コミット前の検査（pre-commit フック）

clone したら 1 回だけ次を実行する。`core.hooksPath` を `tools/hooks` にし、コミットのたびに `tools/hooks/pre-commit` と `tools/hooks/commit-msg` が動く（git worktree で作ったツリーでも、そのツリーのフックが動く）。

```
.\tools\install_hooks.ps1              有効にする
.\tools\install_hooks.ps1 -Uninstall   やめる
```

| フック | 順 | 内容 | 所要時間の目安 |
|---|---|---|---|
| `pre-commit` | 1 | `tools/check_commit.ps1 -Staged`：ステージした内容と、これから作るコミットの作者を検査する | 数秒 |
| `pre-commit` | 2 | `tools/run_commit_tests.ps1`：ステージした変更に対応するテストだけを、既定のタグ（`Unit`・`Io`・`Meta`）で流す。失敗したテストだけを表示する。選び方は下の表。全テストと静的解析は CI で行う | 変更による（文書だけなら 20 秒ほど、全部流すと 1 分半ほど） |
| `commit-msg` | 1 | `tools/check_commit_message.ps1 -Path <メッセージのファイル>`：1 行目が Conventional Commits の形（`feat: ...` など。AGENTS.md「GitHub の運用」）か。`#` で始まる行と空行は飛ばし、git が自動で作るメッセージ（`Merge ...` `Revert "..."` `fixup! ...` など）は調べない | – |
| `commit-msg` | 2 | `tools/check_signoff.ps1 -Path <メッセージのファイル>`：作者（`git var GIT_AUTHOR_IDENT`）のメールアドレスの `Signed-off-by:` 行があるか（`git commit -s` で付く。大文字・小文字は区別しない）。`#` で始まる行と、`git commit -v` の切り取り線より後ろ（差分）は数えない。マージコミット（1 行目が `Merge `）と bot の作者は調べない | – |

`run_commit_tests.ps1` は、ステージした変更（名前の変更は、消したファイルと足したファイルに分ける）から流すテストを選ぶ。`-List -Files <パス>` で、選んだテストを流さずに出せる。

| 変更したファイル | 流すテスト |
|---|---|
| `scripts/<パス>.ps1` | `tests/<パス>.Tests.ps1` と `meta/structure`・`meta/layers`（`scripts/tebunko/indexer.ps1` は `tests/tebunko/indexer/indexer.Tests.ps1`） |
| `scripts/` の `.xaml` | `meta/structure` |
| `tests/` の `*.Tests.ps1` | そのテスト |
| `tools/<名前>.ps1` | `tests/tools/<名前>.Tests.ps1`（無ければ流さない）。`check_markdown_links.ps1` は `meta/links` も |
| `tests/testdata/scrub_personal.ps1` | `tests/testdata/scrub_personal.Tests.ps1` |
| `.md`・`docs/` の中 | `meta/links` |
| 対応するテストが無い・消した `scripts/` の `.ps1`、`tests/` のほかのファイル（`run.ps1`・`helpers/`・`testdata/` など） | 速いテスト（`Unit`・`Meta`）を全部 |
| それ以外（`.github/`・画像・設定など） | 流さない |

文字コードは `check_commit.ps1 -Staged` が確かめるため、`meta/encoding` はフックでは流さない。

引っかかったら内容を直す。`git commit --no-verify` で飛ばしても、CI で同じ検査が走る（コミットメッセージは、squash merge で main のコミットになる PR のタイトルを CI が確かめる）。フックは git for Windows の sh が読むため LF で保存する（`.gitattributes` で固定）。

## 公開してはいけない内容の検査（`tools/check_commit.ps1`）

リポジトリは公開しているため、履歴に一度でも入れたら公開されたのと同じになる（AGENTS.md「個人情報を書かない」）。これを機械的に止める。手元では `-Staged` でステージした内容を、CI では追跡している全ファイルと全コミットを調べる。

| 対象 | 見つけたら失敗にするもの | 例示用として通すもの |
|---|---|---|
| 絶対パス | `C:\Users\<名前>`（`/` 区切り・`\\` のエスケープも） | `test`（`test_` のように長さを揃えたものも）・`a`・`Public`・`Default`・`%USERNAME%` `<利用者名>` のような置き換え用の表記 |
| メールアドレス | すべて | `example.com` などの例示用のドメイン、GitHub の noreply |
| Office ファイル | 作成者・最終更新者（`docProps/core.xml`・ODF の `meta.xml`・PDF の `/Author`）、コメント・変更履歴の作成者に付くアカウント ID（`userId`） | `test`、空、`0` だけの ID |
| 手元で実行したとき | Windows のユーザー名そのもの（CI では調べない） | — |
| `.ps1`・`.xaml` | BOM 付き UTF-8 でない、改行が CRLF でない | — |
| ファイル名・フォルダ名 | UTF-8 で 255 バイトを超えるもの（上の Scorecard の制約） | — |
| コミットの作者・コミッター | GitHub の noreply 以外のメールアドレス | `noreply@github.com`（GitHub 上でマージしたときのコミッター） |

Office ファイル（ZIP）は展開して中の XML・バイナリを 1 つずつ調べる。旧形式（`.doc` `.xls` `.ppt`）・PDF などはバイト列を UTF-8 と UTF-16LE の両方で読んで調べる。中身は作業ツリーではなく git に記録される内容を読む（改行だけは、git が LF にして記録するため作業ツリーのファイルで調べる）。例示用として通す名前・ドメインを増やすときは `tools/check_commit.ps1` の先頭で行う。
