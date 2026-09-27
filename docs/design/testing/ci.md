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
| `tests/tools/` | 開発用の道具（`check_commit_message`・`check_signoff`・`check_release_tag`・`check_markdown_links`・`measure_perf`・`run_commit_tests`） |
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
| `Office` | Excel・Word・PowerPoint の COM を実際に動かすもの（`indexer.Tests.ps1` の「利用者のPowerPointが起動している場合」。ほかは `Mock` で確かめる） | 必要 | しない（`-All` で実行） |
| `Slow` | 時間のかかるもの（検索・本文インデックスの作成・取り込みの速さの回帰テスト。`tests/tools/perf_*.Tests.ps1`。[`perf-check.yml`](#ci)） | 検索の側は tebunko-perfdata（データを作るスクリプトのリポジトリ）が要る | しない（`-All` または `-Tag Slow -ExcludeTag Manual` で実行） |
| `Manual` | 手で確かめるもの（今は該当するテストが無い） | – | しない（`-All` でも実行しない） |

実行は `tests/run.ps1` から行う。

```
.\tests\run.ps1              既定（Unit・Io・Meta。Office・Slow・Manual は外す）
.\tests\run.ps1 -Tag Unit    速い確認だけ
.\tests\run.ps1 -All         Office・Slow も含める（Office と、tebunko-perfdata が必要）
.\tests\run.ps1 -Tag Slow -ExcludeTag Manual   Slow だけ（手元で 3〜4 分ずつ。下の「`perf-check.yml`」）
.\tests\run.ps1 -Ci          結果の XML（work\test\results.xml）とカバレッジ（work\test\coverage.xml）を出し、カバレッジの下限を確かめる
.\tests\run.ps1 -Path .\tests\shared\core   指定したフォルダ・ファイルのテストだけ
.\tests\run.ps1 -Quiet       失敗したテストだけを表示する
```

いずれも失敗したテストの数を終了コードにする（フックと CI が見る）。1 件も実行しなかったときも失敗にする（終了コード 1）。タグの打ち間違いで、何も確かめないまま通るのを防ぐため。

`-Tag Slow` だけでは、既定の除外（`Office`・`Slow`・`Manual`）が残って 0 件になる。`-ExcludeTag` を渡すと既定の除外が置き換わるので、`-Tag Slow -ExcludeTag Manual` とする。

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
| `release.yml` | `guard`・`test`・`release` | `v` で始まるタグの push | テストのうえ、配布 zip とインストーラーを GitHub Release に載せる | – |
| `perf.yml` | `perf` | 手動（`workflow_dispatch`） | Office からの取り込み（.docx・.pptx）・本文インデックスの作成・検索の速さとリソースの推移を測る | – |
| `perf-check.yml` | `search`・`ingest` | PR にラベル `perf-check` を付けたとき（付けたあとの push でも）、main への push（速さに効くファイルが変わったとき）、手動 | 検索・本文インデックスの作成・取り込み（.docx・.pptx）の速さを上限と比べる（回帰テスト） | –（流した PR で落ちていればマージしない） |

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

`v` で始まるタグを push すると動く。最初の `guard` が `tools/check_release_tag.ps1` で、タグが `v<メジャー>.<マイナー>.<パッチ>` の形（大文字の `V`・全角の数字・0 始まり・`-rc1` などの接尾辞は不可）で、指すコミットが `origin/main` の履歴にあることを確かめる（満たさないタグからは、テストにも公開にも進まない）。続けて `test.yml` と同じ検査・テストを通したうえで、`tools/new_release_package.ps1` で配布 zip（`tebunko-<タグ>.zip`）を、`tools/new_installer.ps1` でインストーラー（`tebunko-setup-<タグ>.exe`）を作って GitHub Release に載せる。

- zip にはツール本体（`tebunko.bat`・`scripts/`）と `README.md`・`LICENSE`・`VERSION.txt`（版とコミットの記録）だけを入れる。README の相対リンクと画像は、その版の GitHub の URL に書き換える。カタログ（`tebunko.cat`）・ハッシュ一覧（`SHA256SUMS.txt`）・部品表（`sbom.cdx.json`。配布物を作るたびに zip の中身から作る）は zip と並べてリリースに載せ（[安全性の要約](../../safety/index.md) の [配布物の完全性（カタログ・ハッシュ一覧・来歴の署名）](../../safety/scans.md#配布物の完全性カタログハッシュ一覧来歴の署名)）、SECURITY は README とリリースの説明からリンクする。zip 自体の SHA256 はリリースの説明に書く（同 [複数エンジンでの検査: VirusTotal（外部へファイルを送信する）](../../safety/scans.md#複数エンジンでの検査-virustotal外部へファイルを送信する) の VirusTotal での照会用）
- 配布物を作った後、来歴に署名する前に、`tools/check_release_package.ps1` が zip の中身（ファイルの過不足・dot-source の先・構文・カタログ・ハッシュ一覧・部品表・`VERSION.txt`）を確かめる。通らなければ、署名も公開も行わない。公開する zip は変更の無い作業ツリーで作る（CI の checkout はそうなっている）
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

手動で起動し、決まった量のデータで、Office からの取り込み、インデックス作成、検索の速さ、およびその間のリソース（メモリ・CPU・スレッド・ハンドル・GC）を測る。大きな変更の前後で数字を比べるときなどに使う。必須のチェックではない。

```
gh workflow run perf.yml -f ref=<測る ref> -f scale=0.1 -f ingest=200
```

- 入力: `ref`（測る ref。作業中のブランチも測れる）、`scale`（データの量。`1` でブック 5.4 万・TSV 16 万・約 1GB）、`count`（語ごとに続けて検索する回数。既定 20）、`ingest`（取り込みを測る .docx・.pptx のそれぞれの数。`0` `200` `1000`。既定 `0` は取り込みを測らない。検索だけを測る起動が長くならないようにする）
- 計測スクリプト（`tools/measure_perf.ps1` と `tools/perf/`）はワークフローの ref から、計測対象のコードは入力の ref からチェックアウトする。計測スクリプトが呼ぶ関数（`publishIndexFolders`・`getIndexPackFiles`・`searchPackIndex`）が無い ref（本文インデックスの形式より前の版）では、エラーメッセージを出して止まる
- データは [tebunko-perfdata](https://github.com/hsgwa/tebunko-perfdata) の `new_index.ps1` で毎回生成する（取り込みの一時置き場と同じ形の TSV）。使う版は `perf.yml` の `PERFDATA_SHA` でコミットに固定する。検索する語は同じリポジトリの `words.tsv`（0 件・まれ（3 件）・大量・正規表現）
- 測るもの（それぞれ別のプロセスで動かし、リソースが混ざらないようにする）
  - Office からの取り込み … `ingest` が `0` でないときだけ。`tools/perf/new_ingest_data.ps1` でテストデータ（`tests/testdata/office`）の複製を作り（.docx・.pptx を `ingest` 個ずつ。50 ファイルごとにフォルダを分ける）、`measure_perf.ps1 -Office` で測る。ランナーには Office が無いので、Office を使わずに読むファイル（.docx・.pptx。読み取りのスレッド）だけを測り、Excel・Word・PowerPoint を使う取り込み（.xlsx・.doc・.ppt）は手元の Windows で測る（下の「取り込みの計測」）
  - インデックス作成 … インデクサと同じ `publishIndexFolders`（本文インデックスを書く・TSV を消す・システムインデックスを作る）。Office からの取り込みは、下の取り込みの計測で別に測る。画面ではインデックス作成のスレッドの優先度を下げている（BelowNormal）が、計測ではほかに動くものが無いので、優先度は下げずに測る
  - 検索 … 語ごとに新しいプロセスを起動し、同じプロセスで `count` 回続けて検索する（上限 1 万件）。画面と同じく、検索の司令のスレッド（`newSearchService`）を 1 つ作り、要求（`newSearchRequest`）を順に送る。`lib.ps1` の読み込みと照合のスレッドの用意は司令のスレッドが始めに 1 回だけ行うので、1 回目の時間に含まれる。検索のキャッシュも画面と同じくプロセスで 1 つを持ち続ける。語ごとに、検索時間の最小・中央値・平均・最大と、その内訳（列挙・照合）を出す。1 回目は画面を開き直した直後の検索に当たり、たいてい最大の値になる。高速検索は、ランナーに Windows Search が無いので使わない
    - 司令のスレッドが無い版（#102 より前）は、その版の画面と同じく、検索のたびに新しい Runspace で `lib.ps1` を読み込んで測る（内訳に読み込みの時間も出す）。どちらで測ったかは結果の `SearchMode`（`service` / `runspace`）に出るので、#102 の前後を同じワークフローで比べられる
  - リソース … 別のスレッドで 0.2 秒ごと（および段階の切り替わり）に、プロセスのワーキングセット・プライベート・マネージドヒープ、CPU 時間、スレッド・ハンドルの数、GC の回数と、PC 全体の CPU・空きメモリを記録する。検索では、検索 1 回ごとのメモリも記録し、10 回あたりの増え方（最小二乗の傾き）を出す（検索を続けたときにメモリが増え続けないかを見るため）
- OS のファイルキャッシュは空にしない。本文インデックスは作成の直後なので OS のキャッシュに載っており、PC を起動した直後の検索（ディスクから読む）とは違う
- 結果は、ジョブの Summary（表と Mermaid のグラフ）と、artifact `perf-<実行の番号>`（保存期間は 90 日）に出力する。パスやファイル名は出力しない
  - `summary.md` … Summary と同じ内容（「Office からの取り込み」の節を含む）
  - `result.json` … すべての数字。形式の版（`Schema`）と実行の情報（実行の番号・ref・コミット・scale・日時・ランナーの CPU）を含む
  - `metrics.csv` … 1 行 1 指標の縦長の形（`run_id, date, ref, sha, scale, metric, word, stat, value, unit`）。実行をまたいで数字を貯め、Grafana などで見るときに使う（貯める仕組みはまだ無い）
  - `searches.csv`（検索 1 回ごと）・`resource-ingest.csv`・`resource-index.csv`・`resource-search.csv`（リソースの記録）
- ランナー（4 コア）は手元の PC と条件が違う。Defender のリアルタイム保護の状態は、結果の「環境」に出力する。比べるときは、同じワークフローで測った数字どうしで比べる。同じ条件でも、本文インデックスの作成の時間は 5 割ほどぶれることがある
- 手元でも同じ計測スクリプトで測れる（`.\tools\measure_perf.ps1 -Index <TSV のフォルダ> -Work <作業フォルダ> -Words <words.tsv>`）。渡した TSV は本文インデックスに変換され、元の TSV は削除される

**取り込みの計測（手元の Windows。Excel・Word・PowerPoint が入った機械で）**

```
.\tools\perf\new_ingest_data.ps1 -Dest C:\perf\office -Docx 50 -Pptx 50
.\tools\measure_perf.ps1 -Office C:\perf\office -Work C:\perf\work -Repeat 1
```

- データ（`new_ingest_data.ps1`）… .docx・.pptx（`-Docx`・`-Pptx`。既定 200）、.doc・.ppt（`-Doc`・`-Ppt`。既定 0）はテストデータの複製、.xlsx（`-Xlsx`。既定 0）は tebunko-perfdata の `new_books.ps1` で作ったブック（`-Books`）の先頭の数冊。50 ファイルごとにフォルダを分ける（フォルダごとの本文インデックスの書き出しも一緒に測るため）。同じ引数からは同じ構成・同じ中身になる。Excel は .xlsx、Word は .doc、PowerPoint は .ppt を読むので、使う Office はデータにある種類で決まる
- 測り方（`tools/perf/measure_ingest.ps1`）… `-Repeat` 回（既定 3）、1 回ごとに新しいプロセス・空のワークスペースで流す（Office の起動を含む、初めての取り込みの時間）。測る tebunko の `scripts` を作業フォルダに写して `setting.config` を書くので、利用者の設定・既定のワークスペース・リポジトリの `setting.config` には触らない。取り込みは起動口 `indexer.ps1 -Channel` で動かし、記録のスレッドが受け渡しの口の `Progress.Phase` を読んで、段階（クロール・確認・取り込み・仕上げ）ごとの時間とリソースを出す。1 ファイルあたりの ms は、取り込みの段階の秒 ÷ ファイル数
- 成功・失敗は、ワークスペースの `取り込み一覧.tsv` を計測の側で読んで数える（同じ相対パスは最後の行の状態）。成功 + 失敗がファイル数と合わなければ、終了コードが 0 でなければ、取り込みの段階が読めなかったとき・知らない段階の名前が来たときは、計測を失敗にする。失敗したファイルがあれば `summary.md` に書く
- リソースの表が測るのは、計測の PowerShell のプロセスだけ。EXCEL・WINWORD・POWERPNT のプロセスは含まない（「PC の CPU」には含む）
- Office を使う取り込み（.xlsx・.doc・.ppt）の時間に固定の上限を付けた合否のテストは無い。Office の時間は機械・Defender・Office の版で大きく揺れるので、比べるのは同じ機械・同じ日・同じ引数で続けて測った数字どうしにする（Office の版は結果の実行の情報に出す）。比べ方は下の「Office を使う形式の比べ方」。Office を使わずに読む .docx・.pptx は、`perf-check.yml` が固定の上限と比べる

**計測の口（取り込みの計測が頼るもの）**

取り込みのコードを書き直すときは、この口を変えない。変えるときは、先に計測（`measure_ingest.ps1`）を直し、変える前の版を測り直してから比べる。

| 口 | 中身 |
|---|---|
| 起動口 | `scripts\tebunko\indexer.ps1 -Channel <口>`。終了コード 0 = 完了 |
| 読み込み口 | `scripts\tebunko\lib.ps1`（`resolveTebunkoLib` が頼っている）の `newIndexerChannel`（引数 `retryFailed`・`confirmTargets`・`workers`） |
| 受け渡しの口 | `Progress.Phase` と、その値 `クロール`・`確認`・`取り込み`・`仕上げ` |
| 設定 | `setting.config` がツールのフォルダにあること。キー `targetFolders`（`@{ name; path; enabled }` の配列）・`workspaceFolder`・`ingestThreads` |
| 取り込み一覧 | ワークスペース直下の `取り込み一覧.tsv`。見出しの `相対パス`・`状態`。状態の値 `済`・`失敗` |
| 取り込んだ結果 | 最後の回のワークスペース（`<作業フォルダ>\ingest\ws`。`measure_ingest.ps1` は次の回の始めまで消さない）の下に、本文インデックスのファイル `content_index.*.tsv`（[配置・命名規則](../indexer/index-format.md#配置命名規則)の形。サブフォルダの下にもできる）があること。取り込みの回帰テスト（`perf_ingest.Tests.ps1`）が、数と合計の大きさを数える |

**`perf-check.yml`（速さの回帰テスト）**

検索・本文インデックスの作成・取り込みの速さが落ちたことを、PR の段階で見つける。`measure_perf.ps1` で測り、上限と比べて合否を出す（`perf.yml` は数字を見るだけで、合否は出さない）。必須のチェックにはしないが、流した PR で `search`・`ingest` が落ちていればマージしない。

| 何を | ランナー（`perf-check.yml`） | Windows 機の手元 |
|---|---|---|
| 検索（4 語） | 固定の上限と比べる（`search`） | `.\tests\run.ps1 -Tag Slow -ExcludeTag Manual -Path tests\tools\perf_search.Tests.ps1` |
| 本文インデックスの作成（段階 `pack の作成`） | 固定の上限と比べる（`search`） | 同上 |
| 取り込み（.docx・.pptx。読み取りのスレッド） | 固定の上限と比べる（`ingest`） | `.\tests\run.ps1 -Tag Slow -ExcludeTag Manual -Path tests\tools\perf_ingest.Tests.ps1` |
| 取り込み（.xlsx・.doc・.ppt。Excel・Word・PowerPoint） | 測らない（ランナーに Office が無い） | main と PR のブランチを同じ日に続けて測り、比で判断する（下の「Office を使う形式の比べ方」） |

- 動く時
  - PR にラベル `perf-check` を付けたとき。付けたあとの push・reopen でも流し直す。ラベルが無い PR では、ジョブが「スキップ」になり、失敗にはならない
  - main への push のうち、速さに効くファイル（`scripts/**`・`tools/measure_perf.ps1`・`tools/perf/**`・`tests/tools/perf_*.Tests.ps1`・`tests/testdata/office/**`・`.github/workflows/perf-check.yml`）が変わったとき。ラベルを付け忘れた回帰も、マージの後には見つかる
  - 手動（`workflow_dispatch`）。main への push と手動は、ラベルを見ずに流す
- ラベル `perf-check` は、リポジトリに作ってある。付けるのはコンサルタント（メンテナの代わり）。リリースノートの分類（`.github/release.yml`）には入れない（ほかの分類のラベルと一緒に付くので、その分類に入る）
- ジョブは `search`（検索と本文インデックスの作成）と `ingest`（取り込み）の 2 つを、別のランナーで並べて流す。取り込みの後に同じジョブで検索すると遅く出る回があるため、分ける。どちらも `windows-latest`（4 コア）、`timeout-minutes: 30`。ランナーでかかる時間は `search` が 3〜4 分、`ingest` が 4〜5 分
- 各ジョブの `if` は、`github.event_name != 'pull_request'`（main への push・手動）、または `labeled` でラベル名が `perf-check`（付けたとき）、または `labeled` 以外（`synchronize`・`reopened`）で PR に `perf-check` が付いているとき。ほかのラベルを付けたときは流し直さない
- 流れている実行の取り消し（`concurrency`）は、同じ PR に push を足したときと、`perf-check` を付け直したときだけ。グループは `perf-check-<PR の番号（無ければ ref）>` で、`perf-check` 以外のラベルを付けて起動した実行（ジョブはスキップになる）は、末尾に `run_id` を付けた別のグループに入れる。自分の PR に必ず付ける分類のラベル（`enhancement` など）を付けても、流れている `search`・`ingest` は取り消されない
- 結果は PR の Checks の `search`・`ingest` の合否と、ジョブの Summary（`summary.md` の表）、artifact（`perf-check-search-<実行の番号>-<試行の番号>`・`perf-check-ingest-...`。保存期間は 90 日）で見る。PR にコメントは書かない
- フォークからの PR でも `pull_request` のまま動かす（`pull_request_target` は使わない）。読み取りだけのトークンで動き、秘密の値は使わない（tebunko-perfdata は公開）。フォークの作者はラベルを付けられないので、流すかどうかはメンテナの側が決める。初めての貢献者の PR は、GitHub の設定どおり実行の承認が要る
- アクションはハッシュで固定する。`PERFDATA_SHA`（tebunko-perfdata のコミット）は `perf.yml` と同じ値にする（`PESTER_VERSION` と同じく、2 か所を手で同じにする）。`run` の中は ASCII だけで書く

**回帰テストの中身**

| テスト | データ | 測り方 | 比べる値 |
|---|---|---|---|
| `tests/tools/perf_search.Tests.ps1`（検索と本文インデックスの作成） | tebunko-perfdata の `new_index.ps1 -Scale 0.1`（種は既定の 1。ブック 5,361・TSV 16,078 → 本文インデックスのファイル 521・118MB）。語は同じリポジトリの `words.tsv` | `measure_perf.ps1 -Index ... -Words ... -Count 20`（`perf.yml` の既定と同じ手順・同じ条件） | 語ごとの検索の中央値（`Search[].TotalMs.Median`）・本文インデックスの作成の秒（`Index.Pack.Seconds`。1 回の値） |
| `tests/tools/perf_ingest.Tests.ps1`（取り込み） | `new_ingest_data.ps1 -Docx 50 -Pptx 50`（テストデータの複製。計 100 ファイル） | `measure_perf.ps1 -Office ... -Threads 2 -Repeat 3`（1 回ごとに新しいプロセス・空のワークスペース） | 1 ファイルあたりの ms の中央値（`Ingest.PerFileMs.Median`）・全体の秒の中央値（`Ingest.Seconds.Median`） |

比べる処理は `tools/perf/perf_common.ps1` の `getSearchPerfProblems`・`getIngestPerfProblems` で、合わないものの一覧を返す（空なら合格）。失敗のメッセージには、対象の名前・値・上限の数字だけを書く。上限との境目・欠けた語・件数などは、各テストファイルの `Unit` で確かめる。

- 「速く終わっても、何もしていない」誤りを通さないため、時間のほかに次も確かめる
  - 検索: `searches.csv` の 20 回すべてで、件数が `words.tsv` の「件数」と合うこと（数ならその件数、`10000+` なら 1 万件で打ち切り）、照合した本文インデックスのファイルの数（`Packs`）が本文インデックスの作成の数（521）と同じで 0 より大きいこと。件数は scale 0.1・種 1 のときの値。検索の流れ（`Run.SearchMode`）が `service` であること（`newSearchService` が無いと、`measure_search.ps1` は黙って `runspace` で測るため）
  - 取り込み: `Total`・`Done` が 100、`Failed` が 0。最後の回のワークスペースに本文インデックスがあり、合計の大きさが 0 より大きいこと（取り込み一覧の状態だけが `済` になり、中身を書かずに終わる誤りを通さないため。この形に頼ることは、下の「計測の口」の表にある）
- 結果（`summary.md`・`result.json` など。数字だけでパスは入らない）は `work\test\perf-search\`・`work\test\perf-ingest\`（git 管理外）に残る。`summary.md` の見出しに、tebunko-perfdata のコミットが入る（取れなければ「不明」）
- tebunko-perfdata の場所は、環境変数 `TEBUNKO_PERFDATA`。無ければリポジトリと並んだ `tebunko-perfdata`（git worktree のときは、本体のチェックアウトと並んだもの）。どちらにも無いときは、`git clone` の取り方を示して失敗にする。手元の clone は、`PERFDATA_SHA` と `new_index.ps1`・`words.tsv` が同じであること（違うとデータが変わり、ランナーの数字と比べられない）
- `-All` は `Slow` も流すので、tebunko-perfdata が要る

**上限**

上限は、ランナー（`windows-latest`）で、`perf-check.yml` と同じ構成（同じジョブ）で 5 回測った比べる値の最大に余裕を掛け、切り上げて決める。各テストファイルの先頭の表（`searchLimitsMs`・`packLimitSeconds`・`perFileLimitMs`・`secondsLimit`）に置く。

| 対象 | 比べる値 | 5 回の実測（最小〜最大） | 余裕 | 上限 |
|---|---|---|---|---|
| 検索 0 件 | 中央値 | 298〜337 ms | 1.5 倍（50 ms 単位） | 550 ms |
| 検索 まれ | 中央値 | 353〜500 ms | 1.5 倍（50 ms 単位） | 750 ms |
| 検索 大量 | 中央値 | 1,290〜1,634 ms | 1.5 倍（50 ms 単位） | 2,500 ms |
| 検索 正規表現 | 中央値 | 1,458〜1,517 ms | 1.5 倍（50 ms 単位） | 2,300 ms |
| 本文インデックスの作成（段階 `pack の作成`） | 1 回の秒 | 39.8〜51.6 秒 | 2 倍（5 秒単位） | 105 秒 |
| 取り込み 1 ファイルあたり | 中央値 | 723〜793 ms | 1.5 倍（50 ms 単位） | 1,200 ms |
| 取り込み 全体 | 中央値 | 73.2〜80.5 秒 | 1.5 倍（10 秒単位） | 130 秒 |

- 余裕: 検索の同じ条件の 2 回の差は 7% ほど、取り込みの 3 回の差は 2% ほど。ランナーの機械の違いを見込んで 1.5 倍を取りつつ、流れが崩れたときに出る数倍の遅れは確実に止める。本文インデックスの作成は 1 回の値で、5 割ほどぶれることがあるため 2 倍にする
- `perf.yml` の数字は、構成が違う（取り込みの後に同じジョブで検索すると遅く出た回がある）ので、上限の元にしない。インクリメンタルサーチの目標（7,000 件規模で 0.1 秒前後）も上限にしない。今の main はこのデータで 0.3 秒ほどなので、上限にすると今の main で落ちる
- 変え方: 速くした PR では、同じ決め方で下げてよい。上げるのは、ランナーが遅くなったなど、速さを落としていない理由があるときだけにする。理由と数字を PR 本文に書き、メンテナの了承を得る
- 上限はランナー（4 コア）に合わせたもの。手元の PC で超えたときは、同じ PC で main を測って比べ、回帰かどうかを見る

**誤って落ちたとき**

- ジョブを 1 回だけ再実行する。続けて落ちたら、`perf-check.yml` を main で `workflow_dispatch` で流して比べる
- main も同じくらい遅ければ、ランナーの問題として扱う。上限を上げるかは、メンテナが決める
- main が通って PR だけが落ちれば、回帰として直す

**Office を使う形式の比べ方（手元の Windows。Excel・Word・PowerPoint が入った機械で）**

- 対象は .xlsx（Excel）・.doc（Word）・.ppt（PowerPoint）の取り込み。ランナーに Office が無く、Office の版・機械・Defender で大きく揺れるので、固定の上限は置かない
- データは、.xlsx 100・.docx 200・.pptx 200・.doc 20・.ppt 20（`new_ingest_data.ps1 -Xlsx 100 -Docx 200 -Pptx 200 -Doc 20 -Ppt 20 -Books <new_books.ps1 で作ったブックのフォルダ>`）。測るのは `-Threads 2 -Repeat 3`
- 同じ PC・同じ日に、main → PR のブランチ → main の順に、`measure_perf.ps1 -Office <データ> -Work <作業フォルダ> -Tool <測る版のフォルダ>` で測る
- 判断: PR のブランチの 1 ファイルあたりの中央値（`Ingest.PerFileMs.Median`）が、前後の main の 2 回のうち大きいほうの 1.2 倍以下で、失敗（`Ingest.Failed`）が 0 件なら合格とする。数字は PR 本文に書く
  - 同じ日・同じ PC で、ほかの作業が動いていないときの 2 回の差は 2% ほどだが、ほかの作業（ほかのテスト・ビルド）と重なると 2 割ほどぶれることがある（同じコードの main を続けて測って 1.02 倍と 1.17 倍、同じコードの PR のブランチと main で 1.24 倍の回があった）。1.2 倍の境目で落ちたときは、ほかの作業が無い時間に測り直してから判断する。測っている間は、ほかのテストや Office の作業を動かさない
- 流す時: `perf-check` を付けた PR のうち、取り込みとインデックスの書き出し（`scripts/shared/office/`・`scripts/tebunko/indexer/`・`scripts/tebunko/index/`・`scripts/tebunko/core/`）に触るもの

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
