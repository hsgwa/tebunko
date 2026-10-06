# ［1 インデックス管理］タブ

細かい仕様は [［1 インデックス管理］タブ](../index-tab.md) にある。

## 目的

検索の対象にするフォルダを登録し、インデックスを作る。インデックスの有無・作成中・中断・失敗の状態を、この画面で見分けられる。

## 部品と並び

- 上に「更新の帯」。更新中（青）・中断（黄）・終わった（緑）を 1 枚で出す。更新中は進み具合と［中止］、中断は［続きから再開］を右に置く。帯が無いときは場所を取らない。
- その下に「インデックス一覧」。チェックを付けた行だけを更新する。列は チェック・インデックス名・パス・ファイル数・最終更新・ステータス・高速検索・操作。表の下に `登録済み: N フォルダ` と `合計: N ファイル`。
- 一覧の右上に［すべて更新］［インポート…］［＋ フォルダを追加］。インデックスが 1 件も無いときは、表の中の案内にも同じ動きの［＋ フォルダを追加］を置く。行の右端の［⋯］から開くメニューに［編集…］［エクスポート…］［削除］。一覧へのドラッグ＆ドロップでも追加できる。
- 下に「インデックスの状態」（選んだ行の詳細。選んでいなければ、すべてのインデックスの合計）。
- ナビの［インデックス管理］の横に、更新中は `● 66%`、中断中は `中断`、失敗があるだけなら `失敗`。失敗があると、失敗したファイルの一覧と「失敗の原因」の列が出る。

## 状態と遷移

```mermaid
stateDiagram-v2
    state "index-tab/empty" as empty
    state "index-tab/add" as add
    state "index-tab/normal" as normal
    state "index-tab/edit" as edit
    state "index-tab/delete-confirm" as del
    state "index-tab/start-confirm" as start
    state "index-tab/running" as running
    state "index-tab/stop-confirm" as stop
    state "index-tab/interrupted" as interrupted
    state "index-tab/done" as done
    state "index-tab/failed" as failed
    empty --> add: ［＋ フォルダを追加］（11）
    add --> normal: ［OK］
    normal --> edit: ［編集…］（14）
    edit --> normal: ［OK］
    normal --> del: ［削除］（15）
    del --> normal: ［削除する］
    normal --> start: 開始（17）
    start --> running: ［更新を開始］
    running --> stop: ［中止］（19）
    stop --> running: ［キャンセル］
    stop --> interrupted: ［中止する］
    interrupted --> start: 続きから再開
    running --> done: 失敗なし（18）
    running --> failed: 失敗あり（18）
```

- 追加のダイアログで入力が足りないと、ダイアログの中に注意が出る（`index-tab/add-error`。12）。
- 一覧の［作成］のチェックを外すと、その行は更新しない（`index-tab/unchecked`。16）。
- 作成中は［＋ フォルダを追加］と行のメニューの［編集…］［削除］が押せない（20）。作成が終わると、また押せる。
- 前回の作成が中断したまま起動すると、この画面が選ばれて `index-tab/interrupted` になる（4）。
状態の判定と画面の更新は[状態の判定と操作の流れ](../state-flow.md)、メッセージの一覧は[メッセージ一覧](../messages.md)にあり、ここには書き写さない。
## 表示する文言

| 状態 | 文言 |
|---|---|
| 一覧が空 | 「フォルダを追加してください」「検索したいフォルダを追加すると、ここに表示されます。」と［＋ フォルダを追加］、［すべて更新］の下に「フォルダを追加すると更新できます」。詳細は「まだインデックスがありません。」「まだインデックスがありません。」「まずインデックスを追加して、［作成］にチェックを付けてください。」 |
| 追加 | 「インデックスの追加」。欄は「元のフォルダ（Office ファイルのある場所）」「インデックス名」。注意は「元のフォルダを指定してください。」 |
| 編集 | 「インデックスの編集」。名前と場所だけを変えられる旨の注記 |
| 削除の確認 | 「「営業」のインデックスを削除しますか？」「元のファイルは削除されません。」、［削除する］［キャンセル］ |
| 更新の確認 | 「インデックス更新の確認」「更新対象: 1 フォルダ / 3 ファイル（最新のフォルダは更新しません）」、［更新を開始］［キャンセル］ |
| 更新中 | 「インデックスを更新しています（18 / 300 件）」「残り 1 分未満」、［中止］ |
| 中止の確認 | 「インデックスの更新を中止しますか？」、［中止する］［キャンセル］ |
| 中断 | 「更新を中断しました（残り 265 件）」、［続きから再開］ |
| 完了 | 「更新が終わりました（成功 3 件 / 失敗 0 件）」、［ログを開く］ |
| 失敗 | 「更新に失敗したファイル 1 件」と「失敗の原因」の列 |

## 写真

### index-tab/empty（インデックスが無い。遷移 3）

![インデックスが無い](../../../images/screens/index-tab/empty.png)

### index-tab/normal（インデックスがあり、更新済み。遷移 7）

![インデックスがあり、更新済み](../../../images/screens/index-tab/normal.png)

### index-tab/interrupted（前回の作成が中断している。遷移 4）

![前回の作成が中断している](../../../images/screens/index-tab/interrupted.png)

### index-tab/add（追加のダイアログ。遷移 11）

![追加のダイアログ](../../../images/screens/index-tab/add.png)

### index-tab/add-error（追加のダイアログの注意。遷移 12）

![追加のダイアログの注意](../../../images/screens/index-tab/add-error.png)

### index-tab/edit（編集のダイアログ。遷移 14）

![編集のダイアログ](../../../images/screens/index-tab/edit.png)

### index-tab/delete-confirm（削除の確認。遷移 15）

![削除の確認](../../../images/screens/index-tab/delete-confirm.png)

### index-tab/unchecked（［作成］のチェックを外した行がある。遷移 16）

![チェックを外した行がある](../../../images/screens/index-tab/unchecked.png)

### index-tab/start-confirm（作成の確認ダイアログ。遷移 17）

![作成の確認ダイアログ](../../../images/screens/index-tab/start-confirm.png)

### index-tab/running（作成中。遷移 17・20）

![作成中](../../../images/screens/index-tab/running.png)

### index-tab/stop-confirm（中止の確認。遷移 19）

![中止の確認](../../../images/screens/index-tab/stop-confirm.png)

### index-tab/done（作成が終わった。遷移 18）

![作成が終わった](../../../images/screens/index-tab/done.png)

### index-tab/failed（失敗したファイルの一覧がある。遷移 18）

![失敗したファイルの一覧がある](../../../images/screens/index-tab/failed.png)
