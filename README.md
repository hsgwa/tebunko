<div align="center">

<img src="docs/images/logo.svg" alt="" width="96" height="96">

# tebunko

**フォルダに溜まった Excel・Word・PowerPoint を、中身の文字で一瞬で探す。**

インストール不要・管理者権限不要・ネットワーク通信なし。zip を展開すれば、職場の PC でもすぐ使えます。

[![最新版](https://img.shields.io/github/v/release/hsgwa/tebunko?label=%E6%9C%80%E6%96%B0%E7%89%88&color=2563EB)](https://github.com/hsgwa/tebunko/releases/latest)
[![test](https://github.com/hsgwa/tebunko/actions/workflows/test.yml/badge.svg?branch=main)](https://github.com/hsgwa/tebunko/actions/workflows/test.yml)
[![codecov](https://codecov.io/gh/hsgwa/tebunko/branch/main/graph/badge.svg)](https://codecov.io/gh/hsgwa/tebunko)
[![OpenSSF Scorecard](https://api.securityscorecards.dev/projects/github.com/hsgwa/tebunko/badge)](https://scorecard.dev/viewer/?uri=github.com/hsgwa/tebunko)
[![Windows](https://img.shields.io/badge/Windows-PowerShell%205.1-0078D4?logo=windows)](#はじめかた)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

[**ダウンロード**](https://github.com/hsgwa/tebunko/releases/latest) ・ [使い方](#使い方) ・ [ドキュメント](https://hsgwa.github.io/tebunko/) ・ [サポート](.github/SUPPORT.md)

<img src="docs/images/screenshot_search.png" alt="tebunko_grep の検索画面。「(株)山田商事」で検索し、Excel・Word・PowerPoint の一致した箇所が一覧に出ている" width="880">

</div>

## なぜ tebunko か

「あの取引先の名前が出てくる見積書は、どれだったか」。共有フォルダに何千もの Office ファイルがあると、1 つずつ開いて探すしかありません。Windows の検索では、Excel のどのシートのどのセルか、PowerPoint の何枚目か、までは分かりません。

**tebunko_grep** は、フォルダの Office ファイルをあらかじめ読み込んでおき（インデックス）、ワードを入れると一致した箇所を数秒で一覧にします。結果をダブルクリックすれば、そのファイルが開きます。Excel なら、そのセルを選んだ状態で開きます。

**tebunko**（手文庫）は、Windows で Office ファイルを扱う道具箱です。いまは検索ツール `tebunko_grep` が入っています。

## 目次

- [できること](#できること)
- [はじめかた](#はじめかた)
- [使い方](#使い方)
- [できないこと](#できないこと)
- [安全性](#安全性)
- [仕組み](#仕組み)
- [開発に参加する](#開発に参加する)
- [今後の予定](#今後の予定)
- [ライセンス](#ライセンス)

## できること

- **速い。** 検索は作っておいたインデックスに対して行います。インデックスが 70MB（約 1 万ファイル分）でも、1〜2 秒で結果が出ます。
- **どこに書いてあるかまで分かる。** 結果は「ファイル → シート・ページ・スライド → 行」の単位で出ます。一致した文字は黄色で示します。
- **本文以外も探せる。** セル・段落・表の文字に加えて、図形・テキストボックス・コメントの文字も検索します。Word・PowerPoint では SmartArt・グラフの文字も検索します。
- **2 回目からは差分だけ。** 更新されたファイルと新しいファイルだけを取り込みます。途中で止めても、次は続きから再開します。
- **共有フォルダでも使える。** ネットワークドライブや UNC パス（`\\server\share`）のフォルダも指定できます。
- **原本に触れない。** 取り込みはコピーを読み取り専用・マクロ無効で開いて行います。元のファイルは書き換えません。

| 種類 | 拡張子 | 結果に出る「場所」 |
|---|---|---|
| Excel | .xlsx / .xlsm / .xls / .xlsb | シート（行・セル） |
| Word | .docx / .docm / .doc | ページ（目安）・ヘッダー/フッター・脚注 |
| PowerPoint | .pptx / .pptm / .ppt | スライド・ノート・フッター |

## はじめかた

**必要なもの**

- Windows（PowerShell 5.1。Windows に最初から入っています）
- Microsoft Excel
- Word・PowerPoint（旧形式の .doc・.ppt を取り込むときだけ）

管理者権限と、実行ポリシーの変更はいりません。

**インストール**

1. [Releases](https://github.com/hsgwa/tebunko/releases/latest) から `tebunko_grep-<バージョン>.zip` をダウンロードし、好きな場所に展開します（共有フォルダでもかまいません）。
2. `tebunko_grep.bat` をダブルクリックします。

初回だけ「セキュリティの警告」が出ます。［実行］を押してください。出したくない場合は、先に `tebunko_grep.bat` のプロパティで［ブロックの解除］にチェックを入れます。

## 使い方

1. **フォルダを登録する。**［1 インデックス管理］タブで［追加…］を押し、Office ファイルのあるフォルダを選びます。
2. **取り込む。**［作成］にチェックを付けて［インデックス作成を開始］を押します。進み具合が画面に出ます。
3. **検索する。**［2 検索］タブでワードを入れて Enter を押します。結果はファイルごとの見出しにまとまります。見出しをクリックすると一致した行が開き、行をダブルクリックすると元のファイルが開きます。

よく使うオプション:

| オプション | 働き |
|---|---|
| ［正規表現を使う］ | オフのときは `(株)` や `1.5` も文字どおりに探します |
| ［大文字と小文字を区別］ | 英字の大文字・小文字を区別します |
| ［図形も検索］［コメントも検索］ | 既定はオン。本文だけを探したいときはオフにします |
| 対象ファイル | ファイル名で絞り込みます。例: `*.xlsx;見積;!*old*` |
| ［結果をファイルに出力］ | 結果を `work/検索結果.txt` に書き出します |

元のフォルダを移動したときは、［編集…］で新しい場所を指定するだけです。取り込み直す必要はありません。

画面ごとの詳しい説明は [ドキュメント](https://hsgwa.github.io/tebunko/) にあります。

## できないこと

- **検索できない文字:** 画像の中の文字、埋め込みオブジェクト、Excel の非表示シート・ヘッダー/フッター・グラフ・SmartArt、PowerPoint のスライドマスター
- **取り込めないファイル:** パスワード付きのファイル、1 ファイルの取り込みに 10 分以上かかるファイル
- **検索の単位:** 1 回に 1 ワード。結果は 10,000 件まで
- **Excel の数値:** 画面に表示されている形で検索します。表示形式が「標準」の 12 桁以上の数値は指数表記（`4.90123E+12`）になるため、元の番号では見つかりません
- **反映の時機:** 元のファイルの変更は、次のインデックス作成で反映されます

すべての制約は [設計書](docs/01_インデックス作成_7_出力TSVと既知の問題.md) にあります。

## 安全性

職場の PC に入れる前に審査できるよう、何をして何をしないかを開示し、テストで毎回確かめています。

- **入っているのは PowerShell スクリプトと画面定義だけ。** 実行ファイル（.exe・.dll）も、第三者のライブラリも含みません（[部品表](sbom.cdx.json)）。
- **しないこと:** ネットワーク通信、管理者権限の要求、レジストリの変更、常駐、実行時のコード生成。
- **検査:** 上の主張は `tests/meta/safety.Tests.ps1` と静的解析（PSScriptAnalyzer・CodeQL）が CI で確かめます。

> [!IMPORTANT]
> インデックス（`work/index/`）は文書の本文を平文で持ち、元のファイルのアクセス権を引き継ぎません。ツールを置くフォルダのアクセス権は、取り込むフォルダと同じか、それより厳しくしてください。

根拠・自分で確かめる手順・配布物の検証方法は [安全性の説明](docs/04_安全性.md) に、脆弱性の報告は [SECURITY.md](.github/SECURITY.md) にあります。

## 仕組み

```mermaid
flowchart LR
    A["Office ファイル"] -->|"読み取り専用でコピーを開く"| B["インデクサ<br>（Excel の COM・ZIP の XML）"]
    B -->|"場所ごとの TSV"| C[("インデックス<br>work/index/")]
    D["画面（WPF）"] -->|"ワード"| C
    C -->|"一致した行"| D
```

- **PowerShell 5.1 と WPF だけで作る。** Windows に最初から入っているものだけを使い、インストールも第三者のライブラリも不要にしました。
- **読める形式は直接読む。** .docx・.pptx の本文や、図形・コメントは ZIP の中の XML を直接読みます。Office のアプリを使うのは、Excel のセルの表示値の取り出しと、旧形式の変換だけです。
- **検索はテキストの一括照合。** 場所ごとの TSV を丸ごと読み、正規表現を全文に 1 回かけます。読んだ内容は使い回すので、2 回目からの検索はさらに速くなります。

設計の詳細は [設計書](https://hsgwa.github.io/tebunko/)（[docs/](docs/00_index.md)）にあります。

## 開発に参加する

不具合の報告・要望・修正の提案を歓迎します。参加する人は [行動規範](.github/CODE_OF_CONDUCT.md) に従ってください。

| したいこと | 場所 |
|---|---|
| 不具合の報告・機能の要望 | [Issue](https://github.com/hsgwa/tebunko/issues/new/choose) |
| 使い方の質問 | [SUPPORT.md](.github/SUPPORT.md) |
| 脆弱性の報告（非公開） | [SECURITY.md](.github/SECURITY.md) |
| 変更の提案 | [CONTRIBUTING.md](.github/CONTRIBUTING.md) |

テストは次で実行します（Pester 3.4）。Pull Request ごとに CI がテスト・カバレッジの下限・静的解析・個人情報の検査を行います。

```powershell
.\tests\run.ps1
```

版は SemVer（`v0.<マイナー>.<パッチ>`）で付け、[Releases](https://github.com/hsgwa/tebunko/releases) に変更点とともに公開します。

## 今後の予定

- 比較ツール `tebunko_diff`（Office ファイルの差分を見る）を加える予定です。

## ライセンス

[MIT](LICENSE)

Microsoft、Excel、Word、PowerPoint、Windows、PowerShell は、米国 Microsoft Corporation の米国およびその他の国における登録商標または商標です。tebunko は Microsoft Corporation とは関係のない個人のプロジェクトであり、Microsoft Corporation が承認・支援したものではありません。
