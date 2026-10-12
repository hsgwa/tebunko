# 速さの回帰テストと上限の決め方

扱うこと: perf-check.yml による速さの回帰テスト、回帰テストの中身、上限の決め方、誤って落ちたときの扱い、Office を使う形式の比べ方。扱わないこと: 手動の性能計測そのもの（[性能とリソースの計測のしかた](perf.md)）。先に読むページ: [CI](ci.md)。

## perf-check.yml（速さの回帰テスト）

検索・本文インデックスの作成・選択行のプレビュー・取り込みの速さが落ちたことを、PR の段階で見つける。`measure_perf.ps1` で測り、上限と比べて合否を出す（`perf.yml` は数字を見るだけで、合否は出さない）。必須のチェックにはしないが、流した PR で `search`・`ingest` が落ちていればマージしない。

```mermaid
flowchart TD
    T["PR にラベル perf-check／main への push／手動"] --> S["search ジョブ<br>検索・本文インデックスの作成・選択行のプレビュー"]
    T --> I["ingest ジョブ<br>取り込み（.docx・.pptx）"]
    S --> CMP1{"上限と比べる<br>getSearchPerfProblems<br>getPreviewPerfProblems"}
    I --> CMP2{"上限と比べる<br>getIngestPerfProblems"}
    CMP1 -- 超えた --> NG1["Checks を失敗にする<br>（必須ではないがマージしない）"]
    CMP2 -- 超えた --> NG2["Checks を失敗にする"]
    CMP1 -- 収まる --> OK["Summary に結果を出す"]
    CMP2 -- 収まる --> OK
```

| 何を | ランナー（`perf-check.yml`） | Windows 機の手元 |
|---|---|---|
| 検索（4 語） | 固定の上限と比べる（`search`） | `.\tests\run.ps1 -Tag Slow -ExcludeTag Manual -Path tests\tools\perf_search.Tests.ps1` |
| 本文インデックスの作成（段階 `pack の作成`） | 固定の上限と比べる（`search`） | 同上 |
| 選択行のプレビュー（読み込みと、残るメモリ） | 固定の上限と比べる（`search`。検索の測りのあとに別のプロセスで流す） | `.\tests\run.ps1 -Tag Slow -ExcludeTag Manual -Path tests\tools\perf_preview.Tests.ps1` |
| プレビューの表の描画（`-Render`。部品の数・スクロールの時間・残るメモリを含む）、読み足し（`-ReadMore`）、作ったセルの数、表を作る時間の 1 回目・2 回目以降 | 値を出すだけで、上限は置かない（選択行のプレビューと同じテストで測る。`powershell.exe` の子プロセスは既定で STA なので、画面を出さずに Measure・Arrange できる） | 上と同じ。小さなインデックスでの動作確認は `.\tests\run.ps1 -Tag Gui -Path tests\tools\perf_preview.Tests.ps1` |
| 取り込み（.docx・.pptx。読み取りのスレッド） | 固定の上限と比べる（`ingest`） | `.\tests\run.ps1 -Tag Slow -ExcludeTag Manual -Path tests\tools\perf_ingest.Tests.ps1` |
| 取り込み（.xlsx・.doc・.ppt。Excel・Word・PowerPoint） | 測らない（ランナーに Office が無い） | main と PR のブランチを同じ日に続けて測り、比で判断する（下の「Office を使う形式の比べ方」） |

- 動く時
  - PR にラベル `perf-check` を付けたとき。付けたあとの push・reopen でも流し直す。ラベルが無い PR では、ジョブが「スキップ」になり、失敗にはならない
  - main への push のうち、速さに効くファイル（`scripts/**`・`tools/measure_perf.ps1`・`tools/perf/**`・`tests/tools/perf_*.Tests.ps1`・`tests/testdata/office/**`・`.github/workflows/perf-check.yml`）が変わったとき。ラベルを付け忘れた回帰も、マージの後には見つかる
  - 手動（`workflow_dispatch`）。main への push と手動は、ラベルを見ずに流す
- ラベル `perf-check` は、リポジトリに作ってある。付けるのはメンテナ。リリースノートの分類（`.github/release.yml`）には入れない（ほかの分類のラベルと一緒に付くので、その分類に入る）
- ジョブは `search`（検索・本文インデックスの作成・選択行のプレビュー。プレビューは検索の測りのあとに別のプロセス・別のデータで流し、どちらかが落ちても、もう一方の結果は残す）と `ingest`（取り込み）の 2 つを、別のランナーで並べて流す。取り込みの後に同じジョブで検索すると遅く出る回があるため、分ける。どちらも `windows-latest`（4 コア）、`timeout-minutes: 30`。ランナーでかかる時間は `search` が 3〜4 分、`ingest` が 4〜5 分
- 各ジョブの `if` は、`github.event_name != 'pull_request'`（main への push・手動）、または `labeled` でラベル名が `perf-check`（付けたとき）、または `labeled` 以外（`synchronize`・`reopened`）で PR に `perf-check` が付いているとき。ほかのラベルを付けたときは流し直さない
- 流れている実行の取り消し（`concurrency`）は、同じ PR に push を足したときと、`perf-check` を付け直したときだけ。グループは `perf-check-<PR の番号（無ければ ref）>` で、`perf-check` 以外のラベルを付けて起動した実行（ジョブはスキップになる）は、末尾に `run_id` を付けた別のグループに入れる。自分の PR に必ず付ける分類のラベル（`enhancement` など）を付けても、流れている `search`・`ingest` は取り消されない
- 結果は PR の Checks の `search`・`ingest` の合否と、ジョブの Summary（`summary.md` の表）、artifact（`perf-check-search-<実行の番号>-<試行の番号>`・`perf-check-preview-...`・`perf-check-ingest-...`。保存期間は 90 日）で見る。このワークフロー自身の `pr-comment` ジョブ（[CI](ci.md) の「結果を PR のコメントに書く（`pr-comment` ジョブ）」）が、`search`・`ingest` のうち悪いほうの結果と実行へのリンクを PR のコメントに出すが、数字（Summary の表）はそこには書き写さない
- フォークからの PR でも `pull_request` のまま動かす（`pull_request_target` は使わない）。読み取りだけのトークンで動き、秘密の値は使わない（tebunko-perfdata は公開）。フォークの作者はラベルを付けられないので、流すかどうかはメンテナの側が決める。初めての貢献者の PR は、GitHub の設定どおり実行の承認が要る
- アクションはハッシュで固定する。`PERFDATA_SHA`（tebunko-perfdata のコミット）は `perf.yml` と同じ値にする（`PESTER_VERSION` と同じく、2 か所を手で同じにする）。`run` の中は ASCII だけで書く

## 回帰テストの中身

| テスト | データ | 測り方 | 比べる値 |
|---|---|---|---|
| `tests/tools/perf_search.Tests.ps1`（検索と本文インデックスの作成） | tebunko-perfdata の `new_index.ps1 -Scale 0.1`（種は既定の 1。ブック 5,361・TSV 16,078 → 本文インデックスのファイル 521・118MB）。語は同じリポジトリの `words.tsv` | `measure_perf.ps1 -Index ... -Words ... -Count 20`（`perf.yml` の既定と同じ手順・同じ条件） | 語ごとの検索の中央値（`Search[].TotalMs.Median`）・本文インデックスの作成の秒（`Index.Pack.Seconds`。1 回の値） |
| `tests/tools/perf_ingest.Tests.ps1`（取り込み） | `new_ingest_data.ps1 -Docx 50 -Pptx 50`（テストデータの複製。計 100 ファイル） | `measure_perf.ps1 -Office ... -Threads 2 -Repeat 3`（1 回ごとに新しいプロセス・空のワークスペース） | 1 ファイルあたりの ms の中央値（`Ingest.PerFileMs.Median`）・全体の秒の中央値（`Ingest.Seconds.Median`） |
| `tests/tools/perf_preview.Tests.ps1`（選択行のプレビュー） | 検索と同じ `new_index.ps1 -Scale 0.1` に、テストの作業用の場所（`$TestDrive`）に作る本文インデックスのファイルを 2 つ足したもの。1 つは大きなファイル（約 10 MB・約 1 万行。架空の文で、ヒットは先頭の近くと終わりの近くに 10 行ずつ）、もう 1 つは横に広いファイル（2,000 列・1,300 行。語は 1,500 列目にだけ置く）。語はそれぞれのファイルに当たる 1 語（大きい本文・横に広い）を `words.tsv` に足す | `measure_perf.ps1 -Index ... -Words ... -Count 1 -Preview 20 -PreviewFullRead -Render -ReadMore`（語ごとに、ヒットの先頭から 20 件を、検索のキャッシュ無し・ありで 1 件ずつプレビューする。`tools/perf/measure_preview.ps1` が別のプロセスで測る。`-PreviewFullRead` は、前の読み方（毎回ファイルの全体を読む）でも同じ選択を測って並べる。`-Render` は表を画面を出さずに描画して部品の数・時間を測り、`-ReadMore` は下に 200 行ずつ読み足す時間を測る） | 語ごとの、キャッシュ無しの読み込みの中央値（`Preview.Words[].ReadUncachedMs.Median`）・プレビューのあと、GC のあとにマネージドヒープに残る増えの合計（`RetainedManagedMB`）。上限と比べるのは、perfdata の語と大きい本文の 4 語（`readLimitsMs` の語）で、横に広い語は値を出すだけで上限と比べない。前の読み方の値（`FullRead`）と、キャッシュ無しの選択で全体を読んだ回数（`ReadTextCallsUncached`）も出すだけ。`-Render` を付けた測りでは、計測が描画のために持つ表 1 枚（全部の列の文字を持つ）が残るメモリの値に入り、前の読み方（`-Render` を付けない）の値には入らない。同じ条件の比べではないので、横に広い語の差はそのまま読まない |

比べる処理は `tools/perf/perf_common.ps1` の `getSearchPerfProblems`・`getIngestPerfProblems`・`getPreviewPerfProblems` で、合わないものの一覧を返す（空なら合格）。失敗のメッセージには、対象の名前・値・上限の数字だけを書く。上限との境目・欠けた語・件数などは、各テストファイルの `Unit` で確かめる。

- 「速く終わっても、何もしていない」誤りを通さないため、時間のほかに次も確かめる
  - プレビュー: 測った回数が、ヒットの数と 20 の小さいほうの 2 倍（キャッシュ無し・あり）であること、作った表のセルが 0 より多いこと。前の読み方の測定は、全部読む関数（`readContentIndexText`）がキャッシュ無しの選択のたびに呼ばれたことを数えて確かめ（時間の大小では決めない）、呼ばれていなければ失敗にする
  - 検索: `searches.csv` の 20 回すべてで、件数が `words.tsv` の「件数」と合うこと（数ならその件数、`10000+` なら 1 万件で打ち切り）、照合した本文インデックスのファイルの数（`Packs`）が本文インデックスの作成の数（521）と同じで 0 より大きいこと。件数は scale 0.1・種 1 のときの値。検索の流れ（`Run.SearchMode`）が `service` であること（`newSearchService` が無いと、`measure_search.ps1` は黙って `runspace` で測るため）
  - 取り込み: `Total`・`Done` が 100、`Failed` が 0。最後の回のワークスペースに本文インデックスがあり、合計の大きさが 0 より大きいこと（取り込み一覧の状態だけが `済` になり、中身を書かずに終わる誤りを通さないため。この形に頼ることは、下の「計測の口」の表にある）
- 結果（`summary.md`・`result.json` など。数字だけでパスは入らない）は `work\test\perf-search\`・`work\test\perf-preview\`・`work\test\perf-ingest\`（git 管理外）に残る。`summary.md` の見出しに、tebunko-perfdata のコミットが入る（取れなければ「不明」）
- tebunko-perfdata の場所は、環境変数 `TEBUNKO_PERFDATA`。無ければリポジトリと並んだ `tebunko-perfdata`（git worktree のときは、本体のチェックアウトと並んだもの）。どちらにも無いときは、`git clone` の取り方を示して失敗にする。手元の clone は、`PERFDATA_SHA` と `new_index.ps1`・`words.tsv` が同じであること（違うとデータが変わり、ランナーの数字と比べられない）
- `-All` は `Slow` も流すので、tebunko-perfdata が要る

## 上限

上限は、ランナー（`windows-latest`）で、`perf-check.yml` と同じ構成（同じジョブ）で 5 回測った比べる値の最大に余裕を掛け、切り上げて決める。各テストファイルの先頭の表（`searchLimitsMs`・`packLimitSeconds`・`readLimitsMs`・`retainedLimitMB`・`perFileLimitMs`・`secondsLimit`）に置く。

| 対象 | 比べる値 | 5 回の実測（最小〜最大） | 余裕 | 上限 |
|---|---|---|---|---|
| 検索 0 件 | 中央値 | 298〜337 ms | 1.5 倍（50 ms 単位） | 550 ms |
| 検索 まれ | 中央値 | 353〜500 ms | 1.5 倍（50 ms 単位） | 750 ms |
| 検索 大量 | 中央値 | 1,290〜1,634 ms | 1.5 倍（50 ms 単位） | 2,500 ms |
| 検索 正規表現 | 中央値 | 1,458〜1,517 ms | 1.5 倍（50 ms 単位） | 2,300 ms |
| 本文インデックスの作成（段階 `pack の作成`） | 1 回の秒 | 39.8〜51.6 秒 | 2 倍（5 秒単位） | 105 秒 |
| プレビュー 読み込み（キャッシュ無し。語ごと） | 中央値 | ランナーで 2 回測った値の最大（この行だけ 2 回） | 1.5 倍（10 ms 単位） | `perf_preview.Tests.ps1` の `readLimitsMs` |
| プレビューのあとに残るメモリ（合計） | GC のあとのマネージドヒープの増え（1 回あたりではなく、語ごとの合計） | ランナーで 2 回測った値の最大（この行だけ 2 回。語のうち一番大きい値） | 2 倍（0.1 MB 単位） | `perf_preview.Tests.ps1` の `retainedLimitMB` |
| 取り込み 1 ファイルあたり | 中央値 | 723〜793 ms | 1.5 倍（50 ms 単位） | 1,200 ms |
| 取り込み 全体 | 中央値 | 73.2〜80.5 秒 | 1.5 倍（10 秒単位） | 130 秒 |

- 余裕: 検索の同じ条件の 2 回の差は 7% ほど、取り込みの 3 回の差は 2% ほど。ランナーの機械の違いを見込んで 1.5 倍を取りつつ、流れが崩れたときに出る数倍の遅れは確実に止める。本文インデックスの作成は 1 回の値で、5 割ほどぶれることがあるため 2 倍にする。プレビューの残るメモリも 2 倍にする（GC のあとの増えは、GC の時機で負の値になることもあるほどぶれるため）
- `perf.yml` の数字は、構成が違う（取り込みの後に同じジョブで検索すると遅く出た回がある）ので、上限の元にしない。インクリメンタルサーチの目標（7,000 件規模で 0.1 秒前後）も上限にしない。今の main はこのデータで 0.3 秒ほどなので、上限にすると今の main で落ちる
- 変え方: 速くした PR では、同じ決め方で下げてよい。上げるのは、ランナーが遅くなったなど、速さを落としていない理由があるときだけにする。理由と数字を PR 本文に書き、メンテナの了承を得る
- 上限はランナー（4 コア）に合わせたもの。手元の PC で超えたときは、同じ PC で main を測って比べ、回帰かどうかを見る

## 誤って落ちたとき

- ジョブを 1 回だけ再実行する。続けて落ちたら、`perf-check.yml` を main で `workflow_dispatch` で流して比べる
- main も同じくらい遅ければ、ランナーの問題として扱う。上限を上げるかは、メンテナが決める
- main が通って PR だけが落ちれば、回帰として直す

## 計測の子が失敗したとき

`tools/measure_perf.ps1` は、計測ごとに子のプロセス（`tools/perf/measure_*.ps1`）を起こす。子が 0 以外で終わると、例外の文に次を入れる（出すのは CI のログと例外の文だけで、ファイルや artifact には残さない）。

- 終了コード（10 進と 16 進。異常終了のときの NTSTATUS が読める）
- 子の標準出力・標準エラーの終わりの 20 行（出力が無ければ「何も出力しませんでした」）。プレビューの子は、段階の切り替えで `段階: <名前>`（読み込み・開始・検索・プレビュー（最初の読み）・読み足し・描画・記録の終わり・記録の停止の終わり・結果の書き出し・書き出しの終わり）を 1 行ずつ出すので、どこまで進んだかが分かる。名前だけで、値は出さない
- CI（`GITHUB_ACTIONS` が `true`）のときだけ、子が動いていた間の Application ログのうち、`powershell.exe` に関する Application Error・.NET Runtime・Windows Error Reporting（5 件まで、1 件 400 文字まで）。手元の機械のログは読まない。ログを読めなくても、上の 2 つは出る

## Office を使う形式の比べ方（手元の Windows。Excel・Word・PowerPoint が入った機械で）

- 対象は .xlsx（Excel）・.doc（Word）・.ppt（PowerPoint）の取り込み。ランナーに Office が無く、Office の版・機械・Defender で大きく揺れるので、固定の上限は置かない
- データは、.xlsx 100・.docx 200・.pptx 200・.doc 20・.ppt 20（`new_ingest_data.ps1 -Xlsx 100 -Docx 200 -Pptx 200 -Doc 20 -Ppt 20 -Books <new_books.ps1 で作ったブックのフォルダ>`）。測るのは `-Threads 2 -Repeat 3`
- 同じ PC・同じ日に、main → PR のブランチ → main の順に、`measure_perf.ps1 -Office <データ> -Work <作業フォルダ> -Tool <測る版のフォルダ>` で測る
- 判断: PR のブランチの 1 ファイルあたりの中央値（`Ingest.PerFileMs.Median`）が、前後の main の 2 回のうち大きいほうの 1.2 倍以下で、失敗（`Ingest.Failed`）が 0 件なら合格とする。数字は PR 本文に書く
  - 同じ日・同じ PC で、ほかの作業が動いていないときの 2 回の差は 2% ほどだが、ほかの作業（ほかのテスト・ビルド）と重なると 2 割ほどぶれることがある（同じコードの main を続けて測って 1.02 倍と 1.17 倍、同じコードの PR のブランチと main で 1.24 倍の回があった）。1.2 倍の境目で落ちたときは、ほかの作業が無い時間に測り直してから判断する。測っている間は、ほかのテストや Office の作業を動かさない
- 流す時: `perf-check` を付けた PR のうち、取り込みとインデックスの書き出し（`scripts/shared/office/`・`scripts/tebunko/indexer/`・`scripts/tebunko/index/`・`scripts/tebunko/core/`）に触るもの

