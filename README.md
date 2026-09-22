<div align="center">

<img src="docs/images/logo.svg" alt="" width="96" height="96">

# tebunko

**Excel・Word・PowerPoint のファイルを、中身の文字でまとめて検索・比較する Windows 用ツール**

インストール不要・管理者権限不要・ネットワーク通信なし。何千ファイルあっても数秒で見つかります。

[![最新版](https://img.shields.io/github/v/release/hsgwa/tebunko?label=%E6%9C%80%E6%96%B0%E7%89%88&color=2563EB)](https://github.com/hsgwa/tebunko/releases/latest)
[![test](https://github.com/hsgwa/tebunko/actions/workflows/test.yml/badge.svg?branch=main)](https://github.com/hsgwa/tebunko/actions/workflows/test.yml)
[![codecov](https://codecov.io/gh/hsgwa/tebunko/branch/main/graph/badge.svg)](https://codecov.io/gh/hsgwa/tebunko)
[![OpenSSF Scorecard](https://api.securityscorecards.dev/projects/github.com/hsgwa/tebunko/badge)](https://scorecard.dev/viewer/?uri=github.com/hsgwa/tebunko)
[![Windows](https://img.shields.io/badge/Windows-PowerShell%205.1-0078D4?logo=windows)](#動作環境)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

[**ダウンロード**](https://github.com/hsgwa/tebunko/releases/latest) ・ [使い方](#使い方) ・ [安全性](#安全性) ・ [ドキュメント](https://hsgwa.github.io/tebunko/) ・ [サポート](.github/SUPPORT.md)

<img src="docs/images/screenshot_search.png" alt="tebunko の検索画面。「(株)山田商事」で検索し、Excel・Word・PowerPoint の一致した箇所が一覧に出ている" width="880">

</div>

**tebunko**（手文庫）は、Windows で Office ファイルを扱う道具箱です。検索（`tebunko_grep`）と比較（`tebunko_diff`）が入っていて、`tebunko.bat` で開く 1 つの画面から、タブを切り替えて使います。

## 特長

検索は、フォルダ内の Office ファイルを「場所」（シート・ページ・スライド）ごとのテキストに書き出してインデックスを作ります。検索はそのインデックスに対して行うので、ファイルを 1 つずつ開くより桁違いに速く結果が出ます。

- **画面だけで操作できる。** インデックスの登録・作成・検索を 1 つの画面で行います。
- **一致した場所をすぐ開ける。** 結果をダブルクリックすると元のファイルを開きます。Excel は該当のセルを選んだ状態で開きます。
- **2 回目からは差分だけ取り込む。** 取り込むのは更新されたファイルと新しいファイルだけです。中断しても続きから再開します。
- **インストールしない。** PowerShell スクリプトだけで動きます。管理者権限も、第三者のライブラリも使いません。
- **共有フォルダにも使える。** ネットワークドライブや UNC パスのフォルダも指定できます。インデックスは別の PC へコピーして検索できます。
- **2 つのファイル・フォルダを比べられる。** Excel・Word・PowerPoint のファイル同士、フォルダ同士を比べ、違うところを並べて表示します（[比較する](#比較する)）。
- **原本を書き換えない。** インデックス作成も比較も、作業フォルダへのコピーを、読み取り専用・マクロ無効で開いて行います。

| 種類 | 拡張子 | 検索結果の「場所」 |
|---|---|---|
| Excel | .xlsx / .xlsm / .xls / .xlsb | シート（行・セル） |
| Word | .docx / .docm / .doc | ページ（目安）・ヘッダー/フッター・脚注 |
| PowerPoint | .pptx / .pptm / .ppt | スライド・ノート・フッター |

## はじめる

### 動作環境

- Windows（Windows PowerShell 5.1。Windows に最初から入っています）
- Microsoft Excel
- Word・PowerPoint は、旧形式（.doc / .ppt）のファイルを取り込む・比べるときだけ使います

実行ポリシーの変更と管理者権限はいりません。

### インストール

1. [Releases](https://github.com/hsgwa/tebunko/releases/latest) から `tebunko-<バージョン>.zip` をダウンロードし、好きな場所に展開します（PC の中でも共有フォルダでもかまいません）。展開すると `tebunko` フォルダができます。
2. `tebunko.bat` をダブルクリックします。

初回だけ「セキュリティの警告」が出ます。［実行］を押してください。2 回目からは出ません。警告を出したくない場合は、事前に `tebunko.bat` のプロパティで［ブロックの解除］にチェックを入れます。

#### 以前の版（tebunko_grep）から移る

以前の版の `tebunko_grep` フォルダにある `work\` フォルダと `setting.config` を、新しい `tebunko` フォルダへ移します。作ったインデックスと設定がそのまま使えます。以前の版の画面は閉じてから移してください。

`tebunko_grep.bat` も同梱しています。以前の版で作ったショートカットからでも、同じ画面が［1 検索］タブで開きます。

> [!TIP]
> ダウンロードした zip が改ざんされていないことは、zip に同梱のカタログで確かめられます（[安全性](#安全性)）。

## 使い方

画面の上の段のタブで、［1 検索］［2 比較］［9 プロセス停止］を切り替えます（Ctrl+1 / Ctrl+2 / Ctrl+9）。

### 検索する

［1 検索］タブの中に、［インデックス管理］と［検索］の 2 つのタブがあります。

1. **インデックスを追加する。**［インデックス管理］タブで［追加…］を押し、Office ファイルのあるフォルダを指定します。
2. **取り込む。** インデックスにチェックを付け、［インデックス作成を開始］を押します。進み具合が画面に出ます。［中止］で止められます。
3. **検索する。**［検索］タブでワードを入れて Enter を押します。一致した箇所が黄色で表示されます。どのタブからでも Ctrl+F で［検索］タブへ移れます。

検索のオプション:

- ［正規表現を使う］: オンにしない限り、`(株)` や `1.5` も文字どおりに検索します。
- ［大文字と小文字を区別］
- ［図形も検索］［コメントも検索］: 図形・テキストボックス・SmartArt・グラフの文字、コメントも検索します（既定はオン）。本文（セルの値・段落）だけを探したいときはオフにします。PowerPoint のテキストボックス・図形の文字はスライドの本文として扱います。
- 「対象ファイル」: ファイル名で絞り込みます。例: `*.xlsx;見積;!*old*`（.xlsx か名前に「見積」を含み、「old」を含まないファイル）
- ［結果をファイルに出力］: `work/検索結果.txt` に書き出します。

元のフォルダを移動した場合は、［編集…］で新しい場所を指定します。インデックス作成をやり直す必要はありません。取り込み後に残った Office のプロセスは、［9 プロセス停止］タブで終了できます。

### インデックスを別の PC で使う

`work/index/<インデックス名>/` のフォルダを、別の PC の `work/index/` へコピーします。別の PC で `tebunko.bat` を開くと、そのまま検索できます。元のファイルが見つからないときは、フォルダを選ぶよう聞かれます。選んだ場所は記録され、次からは聞かれません。

### 比較する

［2 比較］タブで、2 つのファイル、または 2 つのフォルダを比べます。インデックスは使いません。画面の左半分が比較元、右半分が比較先です。

1. **比べるものを選ぶ。** 入力欄の先頭の［ファイル｜フォルダ］で、ファイル同士を比べるか、フォルダ同士を比べるかを選びます。
2. **比較元と比較先を指定する。**［ファイル…］（［フォルダ］のときは［フォルダ…］）で選ぶか、エクスプローラーからドラッグ&ドロップします。
3. **比べる。**［比較］を押します。

- **ファイル同士:** 違うところを左右に並べて表示します。Excel はシートの行とセル、Word は段落、PowerPoint はスライドとノートの単位で比べます。
- **フォルダ同士:** 上に比較元と比較先のフォルダを左右のツリーで、下に選んだファイルの差分を表示します。ツリーの中央の印は、`≠` 変更、`+` 追加（比較先だけ）、`−` 削除（比較元だけ）、`≈` 中身は同じ（抽出した文字が同じ）、`=` 同じ、`!` 比較できない、です。
- F8 で次の変更へ、Shift+F8 で前の変更へ移ります。F5 で同じ条件で比べ直します。
- 差分の行をダブルクリックすると、元のファイルをその位置で開きます。
- ［結果をファイルに出力］: `work/tebunko_diff/比較結果.txt` に書き出します。

詳細は [docs/05_比較.md](docs/05_比較.md) にあります。

## 制約

インデックスに入るのは、セル・段落・表の文字だけです。

- **検索できない文字:** 図形（Excel）、SmartArt・グラフ、埋め込みオブジェクト、コメント、画像内の文字、Excel の非表示シート・ヘッダー/フッター、PowerPoint のスライドマスター
- **取り込めないファイル:** Office 以外のファイル、パスワード付きのファイル、1 ファイルの取り込みに 10 分以上かかるファイル
- **検索:** 1 回に 1 ワードです。一致は 1 行（Excel の 1 行、Word・PowerPoint の 1 段落）の単位です。結果は 10,000 件までです。
- **Excel の数値** は表示されている形で検索します。表示形式が「標準」の 12 桁以上の数値は、指数表記（`4.90123E+12`）になるため元の番号では見つかりません。
- インデックスは取り込んだ時点の内容です。元のファイルの変更は、次のインデックス作成で反映されます。

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
| 取り込む・比べるファイル | 読み取りだけ。作業フォルダへのコピーを、読み取り専用・マクロ無効で開きます |

> [!IMPORTANT]
> インデックス（`work/index/`）は文書の本文を平文で持ち、元のファイルのアクセス権を引き継ぎません。ツールを置くフォルダのアクセス権は、取り込み対象のフォルダと同じか、それより厳しくしてください。比較で書き出した文字は `%TEMP%\tebunko\diff\` に置き、画面を閉じると消します。

- **根拠と確認手順:** [docs/04_安全性.md](docs/04_安全性.md)（審査する側が、同じ結果を自分で確かめられる手順つき）
- **機械的な検査:** 上の主張は `tests/meta/safety.Tests.ps1` が CI で毎回確かめます。静的解析（PSScriptAnalyzer）の安全性のルールでの指摘は 0 件です。
- **配布物の完全性:** Releases の zip にはカタログ（`tebunko.cat`）とハッシュ一覧（`SHA256SUMS.txt`）が入っています。`Test-FileCatalog` で確かめられます（証明書は不要）。zip 自体の SHA256 はリリースの説明に載せています。
- **脆弱性の報告:** [SECURITY.md](https://github.com/hsgwa/tebunko/blob/main/.github/SECURITY.md)（非公開で報告できます）

## 開発に参加する

不具合の報告・要望・修正の提案を歓迎します。手順と決まりは [CONTRIBUTING.md](.github/CONTRIBUTING.md) に、使い方の質問の窓口は [SUPPORT.md](.github/SUPPORT.md) にあります。参加する人は [行動規範](.github/CODE_OF_CONDUCT.md) に従ってください。

## ライセンス

[MIT](LICENSE)

## 商標

Microsoft、Excel、Word、PowerPoint、Windows、PowerShell は、米国 Microsoft Corporation の米国およびその他の国における登録商標または商標です。tebunko は Microsoft Corporation とは関係のない個人のプロジェクトであり、Microsoft Corporation が承認・支援したものではありません。
