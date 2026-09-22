<!-- title-check -->
タイトルを `<型>: <説明>` の形（[Conventional Commits](https://www.conventionalcommits.org/ja/v1.0.0/)）にしてください。PR を出すときのタイトルと、main のコミットのタイトルもこの形になります。

例: `feat: Excel の図形の文字を検索できるようにする` / `fix(gui): 検索結果の件数が 0 のままになるのを直す`

| 型 | 使うとき |
|---|---|
| `feat` | 機能の追加・変更 |
| `fix` | 不具合の修正 |
| `docs` | 文書だけの変更 |
| `refactor` | 動きを変えない書き直し |
| `perf` | 速さの改善 |
| `test` | テストだけの追加・修正 |
| `style` | 書式だけの変更（空白・改行など） |
| `build` | 配布物の作り方・依存の更新 |
| `ci` | CI・git のフック・開発用の道具 |
| `chore` | 上のどれにも当たらないもの |
| `revert` | 前の変更の取り消し |

型の後ろは半角のコロンと空白 1 つです。前の版と互換が無くなる変更は `feat!:` のように `!` を付けます。詳しくは [CONTRIBUTING.ja.md](https://github.com/hsgwa/tebunko/blob/main/.github/CONTRIBUTING.ja.md#コミットと-pull-request-のタイトル) を見てください。
