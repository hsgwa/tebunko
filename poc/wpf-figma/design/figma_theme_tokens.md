# theme のキー（Figma を正にした値）

Figma のコレクション「theme」の Variables と、Text styles の一覧。`theme.xaml` には、Figma の `/` を `.` に替えたキーで入れる（`Bg/Window` → `Bg.Window`）。

- 「今の theme」の欄が空のものは、新しく足すキー。
- 値の差が 1〜2 の色（`#5E6369`・`#0A5CAD`・`#4D4D54`・`#E8A121`）は、近いキーにまとめた（下の「まとめた色」）。

## 色（SolidColorBrush）

| キー | 値 | 今の theme | 用途 |
|---|---|---|---|
| Bg.Window | #F5F7FA | #F6F7F9 | ファイルの見出しの行・プレビューの上の帯・表の見出し |
| Bg.Surface | #FFFFFF | #FFFFFF | 本体の面・入力欄・チェックボックス |
| Bg.Subtle | #F9FAFA | #FAFBFC | 左の欄・結果の上の帯・列の見出し |
| Bg.Stripe | #FAFBFC | | 表の 1 行おきの背景 |
| Bg.Hover | #F3F3F4 | #F3F5F7 | ステータスバー・プレビューの行番号の列・小さなバッジ |
| Bg.Pane | #F0F2F4 | | 検索対象の欄（左の欄の下）の背景 |
| Bg.Tag | #F1F3F4 | | 色の無いバッジ・タグ |
| Bg.Button | #F2F2F5 | | 「参照…」などの小さなボタンの背景 |
| Bg.Summary | #F7FAFC | | 更新の確認の表の合計の行 |
| Bg.Section | #E8EBF0 | | 設定の節の見出しの線 |
| Bg.TitleBar | #F0F0F0 | | 窓のタイトルバー・最小化などのボタン |
| Border.Soft | #E0E2E5 | #E9ECF0 | 区切り線・表の罫線・進み具合の下地 |
| Border.Normal | #D9DEE3 | #D7DBE0 | 詳細の入力欄の枠 |
| Border.Input | #D1D1D1 | | 検索ワード・絞り込み・ファイル内の対象の枠 |
| Border.Strong | #C9CED4 | | 「フォルダを探す」の枠 |
| Border.Separator | #D5D9DE | | 左の欄のナビの下の線 |
| Border.Splitter | #D0D4D9 | | 左の欄の幅を変える線 |
| Border.Grip | #A9AFB6 | | 幅を変える線のつまみ |
| Border.Divider | #E5E8ED | | 詳細の 2 列の間の線 |
| Border.Row | #EDF0F2 | | 詳細の表の行の線 |
| Border.Dialog | #D1D6E0 | | ダイアログの枠・中の線 |
| Border.Check | #9EA3AB | | 大文字・小文字／正規表現の切り替えの枠 |
| Border.Window | #999999 | | タイトルバーの下の線 |
| Ink.Strong | #202124 | #1F2937 | 本文の強い文字・ナビの文字・アイコン（更新など） |
| Ink.Value | #212126 | | 入力欄の値・詳細の見出し |
| Ink.Body | #5F6368 | #414B5A | ステータスバー・補足の文字・アイコンの線 |
| Ink.Muted | #6B737D | #78828F | 項目の名前（インデックス名・パス など） |
| Ink.Subtle | #80868B | | 絞り込みの空のときの文字・検索のアイコン |
| Ink.Placeholder | #9AA0A6 | | 「フォルダを探す」の空のときの文字・版の補足 |
| Ink.Faint | #99A1AB | #A3ABB5 | ヘルプのアイコン・区切りの「\|」 |
| Ink.Note | #8C949E | | ダイアログの「※」の注意書き・情報のアイコン |
| Ink.OnAccent | #FFFFFF | | 主ボタンの上の文字・チェックの印 |
| Button.Text | #4D4D4D | | 枠だけのボタン（保存・参照…）の文字 |
| Button.Icon | #666666 | | 枠だけのボタンのアイコン |
| Accent | #0078D4 | #2563EB | 選んでいるナビの印・チェックボックス・リンク・主ボタン |
| Accent.Hover | #0B5CAD | #1D4ED8 | 種別のチップの文字 |
| Accent.Soft | #E5F1FB | #EFF4FF | 種別のチップの背景 |
| Select.Soft | #E1F2FF | #DBEAFE | 選んでいるナビ・選んでいる行・インデックスの件数のバッジ |
| Hit | #FFF176 | #FDE68A | 結果の中の一致した文字のマーカー |
| Hit.Cell | #FFF3CD | | プレビューの一致したセル |
| Ok | #218A21 | #15803D | 高速検索の表示・「最新」「可」のバッジの文字・進み具合 |
| Ok.Strong | #1E8E3E | | ダイアログの「見つかりました」とチェックの印 |
| Ok.Soft | #E0F7E0 | | 「最新」「可」のバッジの背景・高速検索の表示の背景 |
| Warn | #BA7D00 | #A16207 | 「要更新」「反映中」のバッジの文字 |
| Warn.Dot | #E8A400 | | 検索対象の欄の「要更新」の点 |
| Warn.Soft | #FFF5E0 | | 「要更新」「反映中」のバッジの背景 |
| Warn.Note | #FFF7E0 | | ダイアログの注意の枠の背景 |
| Warn.Line | #F0C36D | | ダイアログの注意の枠の線 |
| Warn.Strong | #6B4E00 | | ダイアログの注意の枠の文字 |
| Danger.Text | #D13438 | #C2410C | 「エラー」のバッジの文字・失敗のアイコン |
| Danger.Dot | #D93025 | | 検索対象の欄の「エラー」の点 |
| Danger.Soft | #FFE6E6 | #FEF4EF | 「エラー」のバッジの背景 |
| Banner.Info.Ink | #0B3D66 | | 情報のバナーの文 |
| Banner.Warn.Ink | #5C3D00 | | 注意のバナーの文 |
| Banner.Ok.Ink | #124D12 | | 完了のバナーの文 |
| Banner.Danger.Ink | #7A1C1F | | 失敗のバナーの文 |
| File.Excel | #107C41 | | Excel のアイコン |
| File.Word | #185ABD | | Word のアイコン |
| File.PowerPoint | #C43E1C | | PowerPoint のアイコン |
| File.Folder | #E8A020 | | フォルダのアイコン・進み具合のバー |

今の theme にあって Figma のコレクションに無いキー: `Bg.Pressed`（#E8EBEF）・`Accent.Ring`（#BFD3FE）・`Danger.Line`（#E9B8A3）。「08 部品」の領域に当たる色が無かった。

### まとめた色

| Figma の値 | まとめたキー | 場所 |
|---|---|---|
| #5E6369 | Ink.Body | 大文字・小文字／正規表現の切り替えの文字 |
| #0A5CAD | Accent.Hover | 106 の 1 か所 |
| #4D4D54 | Button.Text | 「参照…」の文字 |
| #E8A121 | File.Folder | 106 の 2 か所 |

- 「08 部品」の外（「01 画面」〜「07 プロトタイプ」の画面だけ）にある色（#333840・#F0F4F9・#B3261E など）は、まだキーにしていない。画面を領域のインスタンスに置き換えたあとで、残ったものをキーにする。

## 文字（TextBlock の Style）

Figma のフォントはすべて Rethink Sans。行の高さは、書いていないものが自動（WPF では `LineHeight` を書かない）。

### WPF で使うフォント（2026-10-04 から試し中）

Rethink Sans は Windows に無いので、WPF では Windows 10・11 に標準で入っているフォントにそろえる。配布物にフォントを入れない。

- `FontFamily="Segoe UI, Yu Gothic UI"` をウィンドウに 1 か所だけ書く（英数字は Segoe UI、日本語は Yu Gothic UI）。個々の Style・部品にフォント名を書かない。
- どちらのフォントにも Medium が無い。表の太さは次のとおりに読み替える。

  | 表の太さ | WPF の `FontWeight` |
  |---|---|
  | Regular | Normal |
  | Medium | SemiBold |
  | SemiBold | SemiBold |
  | Bold | Bold |

- 文字の描き方（`TextOptions.TextFormattingMode`）は `Display` と `Ideal` の両方を撮って見比べてから決める（未定）。
- 字の幅が Figma と変わるので、文字が切れる・折り返す所は、幅を固定せず中身に合わせる。直せない所は見比べの結果に書く。
- Figma 側のフォントをどうそろえるかは未定（Figma では Segoe UI・Yu Gothic UI を選べない）。

| キー | 大きさ | 太さ | 行の高さ・字間 | 今の theme | 用途 |
|---|---|---|---|---|---|
| Micro | 10 | Regular | | 11・Ink.Faint | ステータスバー・プレビューの上の補足・場所 |
| ColumnHeader | 10 | Bold | | | 結果の列の見出し（場所・種別・該当行） |
| Meta | 11 | Regular | | 12・Ink.Muted | 補足の文字・リンク・木の項目の補足 |
| Meta.Tall | 11 | Regular | 16 | | バージョン情報の値 |
| Meta.Strong | 11 | Bold | | | 件数（「14 件（5 ファイル）」）・畳んだファイルの補足 |
| Meta.Key | 11 | SemiBold | 16 | | バージョン情報の項目の名前 |
| Note | 11 | Regular | 160% | | ダイアログの「※」の注意書き |
| Chip | 11 | Medium | | | 種別のチップ・バッジ |
| Link | 11 | SemiBold | | | バージョン情報のリンク |
| Cell | 12 | Regular | | 12・Ink.Body | 表のセル・木の項目 |
| Cell.Key | 12 | Medium | | 12・Ink.Strong | インデックス名 |
| Label | 12 | SemiBold | | 13 | 欄の見出し・枠だけのボタン |
| Label.Strong | 12 | Bold | | | 詳細の見出し（「営業部 - 詳細」） |
| Brand | 12 | SemiBold | 字間 0.2 | | タイトルバーの「tebunko」 |
| Body | 13 | Regular | | | 検索ワード・結果の本文 |
| Body.Strong | 13 | Bold | | | 一致した文字・開いたファイルの名前 |
| Nav | 13 | Medium | | | ナビの項目 |
| Nav.Tall | 13 | Medium | 18 | | ファイル内の対象の項目（1 か所） |
| Heading | 13 | SemiBold | | 13 | 選んでいるナビ・畳んだファイルの名前 |
| Focal | 14 | SemiBold | | 14 | 設定の節の見出し |
| PageTitle | 16 | Bold | | | 画面の見出し（インデックス管理・Office の終了） |
| Title | 18 | SemiBold | | | 設定の画面の見出し |
| AppTitle | 20 | SemiBold | 28 | | バージョン情報のアプリ名 |

- `Hint`（今の theme では Meta を BasedOn）と `DialogHeading`（16 SemiBold）に当たる文字は、「08 部品」に無かった。`Hint` は `Meta` の BasedOn のまま残す。
- 区切りの「│」の 3 か所は、Figma では Noto Sans JP（Rethink Sans に無い字）なので、Style に結びつけていない。WPF ではウィンドウのフォントのままでよい。
