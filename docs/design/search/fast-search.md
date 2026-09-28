# 高速検索（Windows Search）

扱うこと: Windows Search に検索語を含みうるフォルダを先に絞らせる仕組み、システムインデックス（2-gram）の使い道、使える条件と使えないときの扱い、実測。扱わないこと: 本文インデックスそのものの照合の速さ（[検索を速くする仕組み](speed.md)）、システムインデックスの作り方（[システムインデックス](../index-data/system-index.md)）。先に読むページ: [検索](index.md)、[検索を速くする仕組み](speed.md)。

インデックスが大きいと、すべての本文インデックスを読んで照合するのに時間がかかる（1GB で初回 38 秒・2 回目以降 16〜18 秒。[検索を速くする仕組み](speed.md)）。そこで、Windows Search に「検索語を含みうるフォルダ」を先に絞らせ、その中の本文インデックスだけを [検索を速くする仕組み](speed.md) の方式で照合する（`tebunko/search/fast_search.ps1` の `getFastSearchPackFiles`）。照合と結果の組み立ては [検索を速くする仕組み](speed.md) と同じため、**結果（行・行番号・順番・上限で切る位置）は、すべての本文インデックスを照合したときと同じ**になる。

**システムインデックス**: インデックス作成が、`work/content_index` の中のフォルダごとに、直下の本文インデックスの中身（メタ情報の行を除く）から、文字の 2-gram を英数字の語にした txt を `work/system_index` の同じ相対パスに作る（[システムインデックス](../index-data/system-index.md)）。Windows Search は語単位でしか一致を取らないため、本文をそのまま索引させると語の途中からの一致（`ニター` など）を落とす。2-gram の語なら、ワードを含む本文の txt には、ワードのすべての 2-gram が必ず入っている。

| 項目 | 仕様 |
|---|---|
| 使える条件（画面の「高速検索：使用可」） | Windows Search を開けて `work/system_index` が索引の対象であり、［正規表現を使う］がオフで、ワードに 2 文字以上の部分がある（`testFastSearchUsable`・`getFastSearchView`）。使えないときは [検索を速くする仕組み](speed.md) のとおりすべてを照合する |
| 使えない理由の区別 | `getWindowsSearchState`（`windows_search.ps1`）が、`NoFolder`（`system_index` が無い）・`NoConnection`（Windows Search を開けない・問い合わせの失敗。時間切れを含む）・`NotInScope`（`system_index` が 0 件で、ワークスペースの中も 0 件）・`NotYet`（`system_index` は 0 件だが、ワークスペースの中のほかのものは索引されている）・`Ok` のどれかを返す。管理者の権限なしに Windows Search の対象の一覧を読む方法が無いため、`NotInScope` と `NotYet` はワークスペース（`SCOPE`）の中の様子で見分ける（作ったばかりのワークスペースでは、対象でも `NotInScope` になることがある）。`testWindowsSearch` は `Ok` かどうかだけを返す |
| 反映の進み具合 | `getSystemIndexProgress`（`fast_search.ps1`）が `@{ Folders; Waiting; ContentIndexed }` を返す。`Folders` は txt があるフォルダと状態ファイルの「反映待ち」のフォルダを合わせた数（分けた txt は 1 フォルダ）、`Waiting` は反映待ちのうち Windows Search でまだ反映済みでないフォルダの数（検索と同じ判定。`getReflectedSystemIndexEntries`）、`ContentIndexed` は本文インデックス（`content_index` の TSV）も索引の対象か（`testTsvIndexedByWindowsSearch`）。状態ファイルは書き換えず（反映済みの行を消すのは検索）、読めない・問い合わせに失敗したときは `$null`。検索では呼ばない（画面が確かめるときだけ） |
| 画面の表示と確かめ直す時機 | `getFastSearchView` が短い表示（理由・`使用可（反映 N%）`・`確認中…`。ワードの理由（正規表現・1 文字）を Windows Search の理由より先に出す）、`getFastSearchDetail` が押したときの詳しい画面（状態・進み具合・理由ごとの直し方・確かめた時刻）を返す。確かめ直すのは、画面を開いたとき・ワークスペースを変えたとき・インデックス作成が終わったとき・表示を押したとき、準備中（`NotYet` か反映待ちがある。`testFastSearchPreparing`）の間の 5 分おき。検索のたびの確かめは理由（`FastReason`）だけを差し替える。1 回の確かめは Windows Search への接続を 1 つだけ開いて共有する |
| 語の作り方 | ワードを空白で区切り、2 文字以上の部分の隣り合う 2 文字を小文字にし、UTF-16LE の 4 バイトを 16 進にした語（`x` ＋ 8 桁）にする（`getSearchGrams`）。最大 16 個（多いときは均等に間引く。間引いても候補が増えるだけ） |
| 候補 | `CONTAINS(System.Search.Contents, '"x…" AND "x…"')` で、検索対象のフォルダの中の txt を探す。語の範囲で分けた txt（`system_index_1.txt` …）は語ごとに問い合わせ、すべての語がどれかで見つかったフォルダを候補にする |
| 反映の判定 | txt が Windows Search に反映済み ⇔ `System.Search.GatherTime` が空でない（本文を読み終えた）かつ `System.DateModified` が txt の更新日時（UTC）を秒で切り捨てた値と同じ（Windows Search は秒未満を切り捨てて持つ。`testSystemIndexReflected`） |
| 照合するフォルダ | 候補、反映されていないフォルダ（状態ファイルの「反映待ち」）、対象外のフォルダ、対応済みでないインデックスの全体。これらのフォルダの本文インデックスを `getIndexPackFiles` で集め、`searchPackIndex` で照合する |
| すべてを照合する場合 | 使える条件を満たさない、状態ファイルを読めない、Windows Search への問い合わせに失敗した（時間切れ 10 秒を含む）、インデックス全体（ツリーの一番上より上）・別の場所のインデックス・無いフォルダを検索対象にした |
| 状態の整理 | 反映済みになった「反映待ち」の行は、検索のたびに状態ファイルから消す（書けなければ次の機会に回す） |

```mermaid
flowchart TD
    S(["検索"]) --> U{"高速検索が使える？"}
    U -- いいえ --> ALL["getIndexPackFiles<br>すべての本文インデックス"]
    U -- はい --> Q["Windows Search に問い合わせる<br>候補のフォルダ・反映の判定"]
    Q -- 失敗 --> ALL
    Q --> T["照合するフォルダ<br>候補＋反映待ち＋対象外＋未対応"]
    T --> F["getIndexPackFiles<br>そのフォルダの本文インデックス"]
    ALL --> R["searchPackIndex（4.3 と同じ）"]
    F --> R
```

**`work/content_index` を Windows Search の対象から外す（推奨）**: 本文インデックス（`.tsv`）も索引の対象になっていると、Windows Search はその処理に時間を取られ、システムインデックスが索引されるまで（高速検索が効くまで）1 日以上かかることがある。外しても検索の結果は変わらない（本文インデックスは [検索を速くする仕組み](speed.md) の方式で読む）。ツールは Windows Search の設定を変えないため、利用者が［インデックスのオプション］で `work/content_index` を外す（README の「高速検索」）。

実測（合成データ: ブック 5.4 万・TSV 16 万・1GB、物理メモリ 8GB の PC。結果はすべて一致。本文インデックスにする前の版で、TSV を照合したとき。本文インデックスでは、すべてを照合しても初回 38 秒（[検索を速くする仕組み](speed.md)））:

| 検索 | すべてを照合 | 高速検索 |
|---|---|---|
| 0 件 | 650 秒 | 0.49 秒 |
| まれな語（3 件） | 602 秒 | 29.9 秒 |

よく出る語はほとんどのフォルダが候補になるため、速くならない（上限の 1 万件で打ち切るのは [検索を速くする仕組み](speed.md) と同じ）。
