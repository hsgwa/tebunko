# サポート

## 困ったときは

1. **[README](../README.md) の「使い方」と「制約」を見る。** 検索できない文字・取り込めないファイルの一覧があります。
2. **ログを見る。** インデックス作成の経過と、取り込みに失敗したファイルの理由は `work\インデックス作成ログ.txt` に、インデックス作成を続けられなかったときのエラーは `work\インデックス作成エラー.txt` に出ます。エラーの意味は [docs/01_インデックス作成_8_エラーメッセージ一覧.md](../docs/01_インデックス作成_8_エラーメッセージ一覧.md) にあります。
3. **設計書を見る。** 画面・インデックス作成・検索の細かい動きは [docs/](../docs/00_index.md) にあります。
4. **既存の Issue を探す。** 同じ現象がすでに報告されていないか、[Issue](https://github.com/hsgwa/tebunko/issues?q=is%3Aissue) を検索してください。

## 報告する

| 内容 | 連絡先 |
|---|---|
| 思ったとおりに動かない | [Issue（不具合）](https://github.com/hsgwa/tebunko/issues/new?template=bug.yml) |
| 機能の追加・変更の要望 | [Issue（機能の要望）](https://github.com/hsgwa/tebunko/issues/new?template=feature.yml) |
| 安全性に関わる問題 | [非公開の報告](https://github.com/hsgwa/tebunko/security/advisories/new)（[SECURITY.md](../SECURITY.md)） |

Issue に画面の写しやログを貼るときは、利用者名を含むパス（`C:\Users\<利用者名>\...`）や社名・顧客名を消してから貼ってください。Office ファイルそのものは、機密情報を含むおそれがあるため添付しないでください。

## サポートの範囲

tebunko は個人が開発・保守しているオープンソースソフトウェアです。回答や修正は可能な範囲で行いますが、期限の約束はできません。最新の版で直っていることがあるので、まず [Releases](https://github.com/hsgwa/tebunko/releases/latest) の最新版で試してください。
