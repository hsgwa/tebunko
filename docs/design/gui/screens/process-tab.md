# ［9 プロセス停止］タブ

部品と並び・状態と遷移・表示する文言・改善の候補は、写真を載せた後の別のアイテムでまとめる。細かい仕様は [［9 プロセス停止］タブ](../process-tab.md) にある。写真の中の Excel のプロセスは、`PING.EXE` を `EXCEL.EXE` の名前で写した偽のプロセス（[画面のスモークテスト](../../testing/gui-smoke.md#画面のスモークテスト)と同じ）。

## 写真

### process-tab/empty（Office のプロセスが無い。遷移 36）

![Office のプロセスが無い](../../../images/screens/process-tab/empty.png)

### process-tab/list（偽のプロセスがある。遷移 36）

![偽のプロセスがある](../../../images/screens/process-tab/list.png)

### process-tab/stop-all-confirm（［すべて終了］の確認。遷移 37）

![［すべて終了］の確認](../../../images/screens/process-tab/stop-all-confirm.png)

### process-tab/stop-background-confirm（［バックグラウンドのみ終了］の確認。遷移 37）

![［バックグラウンドのみ終了］の確認](../../../images/screens/process-tab/stop-background-confirm.png)

### process-tab/stop-selected-confirm（選んで終了するときの確認。遷移 38）

![選んで終了するときの確認](../../../images/screens/process-tab/stop-selected-confirm.png)
