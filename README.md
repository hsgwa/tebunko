<div align="center">

<img src="docs/images/logo.svg" alt="" width="96" height="96">

# tebunko

**Excel・Word・PowerPoint のファイルを、中身の文字でまとめて検索する Windows 用ツール**

インストール不要・管理者権限不要・ネットワーク通信なし。何千ファイルあっても数秒で見つかります。

[![最新版](https://img.shields.io/github/v/release/hsgwa/tebunko?label=%E6%9C%80%E6%96%B0%E7%89%88&color=2563EB)](https://github.com/hsgwa/tebunko/releases/latest)
[![test](https://github.com/hsgwa/tebunko/actions/workflows/test.yml/badge.svg?branch=main)](https://github.com/hsgwa/tebunko/actions/workflows/test.yml)
[![codecov](https://codecov.io/gh/hsgwa/tebunko/branch/main/graph/badge.svg)](https://codecov.io/gh/hsgwa/tebunko)
[![OpenSSF Scorecard](https://api.securityscorecards.dev/projects/github.com/hsgwa/tebunko/badge)](https://scorecard.dev/viewer/?uri=github.com/hsgwa/tebunko)
[![Windows](https://img.shields.io/badge/Windows-PowerShell%205.1-0078D4?logo=windows)](#動作環境)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

[**ダウンロード**](https://github.com/hsgwa/tebunko/releases/latest) ・ [使い方](#使い方) ・ [安全性](#安全性) ・ [ドキュメント](https://hsgwa.github.io/tebunko/) ・ [サポート](SUPPORT.md)

<img src="docs/images/screenshot_search.png" alt="tebunko_grep の検索画面。「(株)山田商事」で検索し、Excel・Word・PowerPoint の一致した箇所が一覧に出ている" width="880">

</div>

**tebunko**（手文庫）は、Windows で Office ファイルを扱う道具箱です。いまは検索ツール `tebunko_grep` が入っています。比較ツール `tebunko_diff` を加える予定です。

## 特長

`tebunko_grep` は、フォルダ内の Office ファイルを「場所」（シート・ページ・スライド）ごとのテキストに変換してインデックスを作ります。検索はそのインデックスに対して行うので、ファイルを 1 つずつ開くより桁違いに速く結果が出ます。

- **画面だけで操作できる。** インデックスの作成・変換・検索を 1 つの画面で行います。
- **一致した場所をすぐ開ける。** 結果をダブルクリックすると元のファイルを開きます。Excel は該当のセルを選んだ状態で開きます。
- **2 回目からは差分だけ変換する。** 変換するのは更新されたファイルと新しいファイルだけです。中断しても続きから再開します。
- **インストールしない。** PowerShell スクリプトだけで動きます。管理者権限も、第三者のライブラリも使いません。
- **共有フォルダにも使える。** ネットワークドライブや UNC パスのフォルダも指定できます。インデックスは別の PC へコピーして検索できます。
- **原本を書き換えない。** 変換は作業フォルダへのコピーを、読み取り専用・マクロ無効で開いて行います。

| 種類 | 拡張子 | 検索結果の「場所」 |
|---|---|---|
| Excel | .xlsx / .xlsm / .xls / .xlsb | シート（行・セル） |
| Word | .docx / .docm / .doc | ページ（目安）・ヘッダー/フッター・脚注 |
| PowerPoint | .pptx / .pptm / .ppt | スライド・ノート・フッター |

## はじめる

### 動作環境

- Windows（Windows PowerShell 5.1。Windows に最初から入っています）
- Microsoft Excel
- Word・PowerPoint は、旧形式（.doc / .ppt）のファイルを変換するときだけ使います

実行ポリシーの変更と管理者権限はいりません。

### インストール

1. [Releases](https://github.com/hsgwa/tebunko/releases/latest) から `tebunko_grep-<バージョン>.zip` をダウンロードし、好きな場所に展開します（PC の中でも共有フォルダでもかまいません）。
2. `tebunko_grep.bat` をダブルクリックします。

初回だけ「セキュリティの警告」が出ます。［実行］を押してください。2 回目からは出ません。警告を出したくない場合は、事前に `tebunko_grep.bat` のプロパティで［ブロックの解除］にチェックを入れます。

> [!TIP]
> ダウンロードした zip が改ざんされていないことは、zip に同梱のカタログで確かめられます（[安全性](#安全性)）。

## 使い方

1. **インデックスを作る。**［1 インデックス管理］タブで［新規作成…］を押し、Office ファイルのあるフォルダを指定します。
2. **変換する。** インデックスにチェックを付け、［変換を開始］を押します。進み具合が画面に出ます。［中止］で止められます。
3. **検索する。**［2 検索］タブでワードを入れて Enter を押します。一致した箇所が黄色で表示されます。

検索のオプション:

- ［正規表現を使う］: オンにしない限り、`(株)` や `1.5` も文字どおりに検索します。
- ［大文字と小文字を区別］
- 「対象ファイル」: ファイル名で絞り込みます。例: `*.xlsx;見積;!*old*`（.xlsx か名前に「見積」を含み、「old」を含まないファイル）
- ［結果をファイルに出力］: `work/検索結果.txt` に書き出します。

元のフォルダを移動した場合は、［編集…］で新しい場所を指定します。変換をやり直す必要はありません。変換後に残った Office のプロセスは、［9 プロセス停止］タブで終了できます。

### インデックスを別の PC で使う

`work/index/<インデックス名>/` のフォルダを、別の PC の `work/index/` へコピーします。別の PC で `tebunko_grep.bat` を開くと、そのまま検索できます。元のファイルが見つからないときは、フォルダを選ぶよう聞かれます。選んだ場所は記録され、次からは聞かれません。

## 制約

インデックスに入るのは、セル・段落・表の文字だけです。

- **検索できない文字:** 図形（Excel）、SmartArt・グラフ、埋め込みオブジェクト、コメント、画像内の文字、Excel の非表示シート・ヘッダー/フッター、PowerPoint のスライドマスター
- **変換できないファイル:** Office 以外のファイル、パスワード付きのファイル、1 ファイルの変換に 10 分以上かかるファイル
- **検索:** 1 回に 1 ワードです。一致は 1 行（Excel の 1 行、Word・PowerPoint の 1 段落）の単位です。結果は 10,000 件までです。
- **Excel の数値** は表示されている形で検索します。表示形式が「標準」の 12 桁以上の数値は、指数表記（`4.90123E+12`）になるため元の番号では見つかりません。
- インデックスは変換した時点の内容です。元のファイルの変更は、次の変換で反映されます。

詳細は [docs/](docs/00_index.md) にあります。同じ内容を [Web サイト](https://hsgwa.github.io/tebunko/) でも読めます。

## 安全性

職場の PC に入れる前に審査できるよう、**何をするか・何をしないか**を開示し、その主張をテストで機械的に確かめています。

| 観点 | tebunko の挙動 |
|---|---|
| 構成物 | PowerShell スクリプトと画面定義だけ。実行ファイル（.exe・.dll）を含みません |
| 第三者のライブラリ | 使いません（[部品表](sbom.cdx.json) の第三者の部品は 0 件） |
| ネットワーク通信 | しません |
| 動的なコード実行・実行時コンパイル | しません |
| 管理者権限・レジストリの変更・常駐 | いずれもしません |
| 変換するファイル | 読み取りだけ。作業フォルダへのコピーを、読み取り専用・マクロ無効で開きます |

> [!IMPORTANT]
> インデックス（`work/index/`）は文書の本文を平文で持ち、元のファイルのアクセス権を引き継ぎません。ツールを置くフォルダのアクセス権は、変換対象のフォルダと同じか、それより厳しくしてください。

- **根拠と確認手順:** [docs/04_安全性.md](docs/04_安全性.md)（審査する側が、同じ結果を自分で確かめられる手順つき）
- **機械的な検査:** 上の主張は `tests/meta/safety.Tests.ps1` が CI で毎回確かめます。静的解析（PSScriptAnalyzer）の安全性のルールでの指摘は 0 件です。
- **配布物の完全性:** Releases の zip にはカタログ（`tebunko.cat`）とハッシュ一覧（`SHA256SUMS.txt`）が入っています。`Test-FileCatalog` で確かめられます（証明書は不要）。zip 自体の SHA256 はリリースの説明に載せています。
- **脆弱性の報告:** [SECURITY.md](SECURITY.md)（非公開で報告できます）

## 開発に参加する

不具合の報告・要望・修正の提案を歓迎します。手順と決まりは [CONTRIBUTING.md](CONTRIBUTING.md) に、使い方の質問の窓口は [SUPPORT.md](SUPPORT.md) にあります。参加する人は [行動規範](CODE_OF_CONDUCT.md) に従ってください。

## ライセンス

[MIT](LICENSE)
