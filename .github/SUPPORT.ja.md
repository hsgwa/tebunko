# サポート

[English](SUPPORT.md) | 日本語

## 困ったときは

1. **使い方の「[トラブルシューティング](../docs/guide/troubleshooting.md)」と「[制限事項](../docs/guide/limitations.md)」を確認する。** 主な症状と対処方法、検索対象外の内容・取り込めないファイルの一覧があります。
2. **ログを見る。** インデックス作成の経過と、取り込みに失敗したファイルの理由は、ワークスペース（既定は `%USERPROFILE%\Documents\tebunko_ws`）の `インデックス作成ログ.txt` に出ます。インデックス作成を続けられなかったときのエラーは、画面に出るほか、同じログにも出ます。エラーの意味は [docs/design/indexer/errors.md](../docs/design/indexer/errors.md) にあります。
3. **設計書を見る。** 画面・インデックス作成・検索の細かい動きは [docs/](../docs/design/index.md) にあります。
4. **既存の Issue を探す。** 同じ現象がすでに報告されていないか、[Issue](https://github.com/hsgwa/tebunko/issues?q=is%3Aissue) を検索してください。

## 報告する

| 内容 | 連絡先 |
|---|---|
| 思ったとおりに動かない | [Issue（不具合）](https://github.com/hsgwa/tebunko/issues/new?template=bug.yml) |
| 機能の追加・変更の要望 | [Issue（機能の要望）](https://github.com/hsgwa/tebunko/issues/new?template=feature.yml) |
| 安全性に関わる問題 | [非公開の報告](https://github.com/hsgwa/tebunko/security/advisories/new)（[セキュリティ](SECURITY.ja.md)） |

Issue に画面の写しやログを貼るときは、利用者名を含むパス（`C:\Users\<利用者名>\...`）や社名・顧客名を消してから貼ってください。Office ファイルそのものは、機密情報を含むおそれがあるため添付しないでください。

Issue・Pull Request・脆弱性の報告は、日本語でも英語でも受け付けます。返答は日本語か英語で行います。

## サポートの範囲

tebunko は個人が開発・保守しているオープンソースソフトウェアです。回答や修正は可能な範囲で行いますが、期限の約束はできません。最新の版で直っていることがあるので、まず [Releases](https://github.com/hsgwa/tebunko/releases/latest) の最新版で試してください。
