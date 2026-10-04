# Figma と WPF の構造の対応（案、2026-10-03）

目的: Figma の部品の分け方・名前・レイアウトを、そのまま WPF の XAML のファイル・x:Name・パネルにする。
画面を領域ごとのファイルに分け、領域ごとに実装・テストできるようにする。

## 1. 対応の決まり

| Figma | WPF |
|---|---|
| Variables（色・角丸） | `theme.xaml` の Brush・CornerRadius のキー（`Bg.Window`・`Ink.Body`・`Accent` など。名前を同じにする。`Bg/Window` のようにスラッシュで階層にしてよい） |
| Text styles | `theme.xaml` の TextBlock の Style（`Heading`・`Label`・`Meta`・`Hint`・`Cell` など） |
| 基本部品のコンポーネント（ボタン・入力欄・チップ・バッジ・バナー） | `theme.xaml` の Style／ControlTemplate（`Primary`・`Link`・`ToolButton`・`TextBox.Base` …）。variant の名前 = Style のキー |
| 領域のコンポーネント（`xaml/<画面>/<領域>`） | XAML のファイル 1 つ（`scripts/tebunko/xaml/<画面>/<領域>.xaml`）＋画面層 `ui/<画面>/<領域>.ps1`＋判断層 `ui/<画面>/<領域>_view.ps1` |
| 画面のフレーム（106 の各画面） | 領域のインスタンスを並べただけのもの（状態の見本）。新しい描き込みはしない |
| コードから触る層 | 層の名前 = x:Name（PascalCase。例 `WordBox`・`ResultGrid`）。飾りの層は小文字の名前 |
| Auto layout 縦／横 | `StackPanel`（Orientation）。Fill の子があるときは `Grid`（Fill = `*`・Hug = `Auto`） |
| 固定の幅・高さ | Width／Height。Padding・gap は Padding／Margin（数値をそのまま） |
| 重なり（Absolute） | `Grid` の同じセル |

領域のコンポーネントの説明欄（description）に、`型: UserControl 相当（XamlReader で読み込んで ContentControl に差す）`、使う x:Name の一覧、出す元のデータ（判断層の関数名の案）を書く。

## 2. 分け方（ToBe）

```
shell                      窓（tebunko.xaml）: 左の欄（nav＋下のペイン）・中身の差し込み口・ステータスバー
  shell/nav                左のナビ（検索・インデックス管理・設定・Office の終了・バージョン情報）
  shell/status_bar         下のステータスバー
search                     検索の画面（view_search.xaml）
  search/search_bar        検索ワード・［検索］・種類のチップ・［探す範囲 ▾］・語の扱いのトグル・高速検索のバッジ
  search/target_tree       検索対象のツリー（左の欄の下のペインに出す）
  search/result_list       件数の行（すべて開く／折りたたむ・絞り込み・保存）と結果の表（グループ見出し・行）
  search/preview           プレビュー（見出し・開く／フォルダを開く・格子・注記）
index                      インデックス管理の画面
  index/index_list         一覧（行の操作ボタン・⋯ メニュー・高速検索の列）
  index/index_detail       詳細欄（基本設定・高速検索の反映の進み具合）
settings                   設定の画面
office                     Office の終了の画面
dialog/*                   ダイアログ 1 つ = 1 ファイル（dialog_confirm・dialog_about・dialog_index_edit・dialog_indexing_confirm・dialog_import …）
parts                      theme.xaml に入れる基本部品（button・text_box・chip・badge・banner・tree_item・grid_row など）
```

## 3. 今の WPF からの移り方（Windows で実装、別の起票）

- 今の `tab_*.xaml` を `<画面>/<領域>.xaml` に分ける。読み込みは今のタブと同じ仕組み（`loadXaml` して `ContentControl.Content` に差し、ファイルごとの Names で `$ui` を作る）を領域に広げる。
- `*_tab.ps1`・`*_view.ps1` も領域ごとに分ける（判断層のテストは領域ごと）。
- 新しい色・形は `theme.xaml` にキーを足してから使う（Figma の Variables と同時に足す）。
