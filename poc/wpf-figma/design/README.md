# PoC 段 1 の見比べの材料（2026-10-03 版。マージしない）

PoC「Figma の UI を WPF で再現して見た目を確かめる」の段 1 で、WPF を合わせる正。前の見比べ材料の 26 枚は古いので、こちらに置き換える。

## 中身

| パス | 中身 |
|---|---|
| `png/` | 段 1 の 26 枚（前と同じノード。ファイル名はノード ID、`names.tsv` に画面の記号）。ウィンドウは 1280×820 に切り取り済み |
| `figma_wpf_map.md` | 領域の部品 ↔ XAML のファイル・x:Name の対応、部品（parts）、ファイルの種類のアイコン |
| `figma_theme_tokens.md` | 色（Variables「theme」57 色）・Text styles 23 の値。theme.xaml のキー名と同じ |
| `figma_wpf_structure.md` | Figma の構造 = WPF の構造（領域 1 つ = XAML 1 ファイル）の決まり |
| `illust/` | 空の状態のイラスト 2 つ（Lucide を 64px・線 2px 相当にしたもの） |
| `illust/app_logo.svg` | アプリのロゴ（Figma の「logo 1」を書き出したもの） |
| `lucide/` | 画面で使う Lucide の公式 SVG（lucide-static 1.50.0、ISC ライセンス） |

## 前の材料から変わったこと（WPF で合わせ直すところ）

1. **アイコンはすべて Lucide の公式 SVG。** XAML では SVG の `d` をそのまま `Path.Data` にし、`Viewbox`（24×24 の Canvas）で表示の大きさに縮める。塗りなし、線の端と角は Round。
   - ⚡ → zap、✓ ✕ ⚠ とバナーの印 → check・x・circle-check・info・circle-alert・circle-x・triangle-alert。
   - 見出しの開閉の塗った ▸ ▾ → chevron-right・chevron-down（線、色 `Ink.Body`）。
   - 行のフォルダ → folder、⋯ → ellipsis、外部リンク → external-link、更新 → refresh-cw。
   - 「?」「⋯」「×」「!」「✓」を文字で出すのはやめる（前の計画の「Rethink Sans の文字で出す」は取り消し）。
2. **線の太さは「表示の大きさ × 2/24」で 1 つにそろえる**（10px → 0.83、12px → 1.0、14px → 1.17、16px → 1.33）。`Viewbox` で縮めるなら、Path の `StrokeThickness` は 2 のままでよい。例外: 空の状態のイラスト（64px・線 2px 相当）、チェックボックス。
3. **ファイルの種類のアイコン**（`figma_wpf_map.md` の「ファイルの種類のアイコン」）: Excel = file-spreadsheet（緑 `File/Excel`）、Word = file-text（青）、PowerPoint = presentation（赤）、テキスト（.txt・.md）= file（灰 `Ink/Body`）、フォルダ = folder（黄 `File/Folder`）。
4. **色と文字は Figma を正にする。** `figma_theme_tokens.md` の値で theme.xaml のキーを合わせる（足りないキーは足す）。
5. **チェックボックス・ラジオ・コンボの ▾・区切り線は、WPF の標準の部品（CheckBox・RadioButton・ComboBox・Separator）を theme で Figma の見た目に寄せる。** Figma のパスをそのまま描かない。
6. **初めて使うとき（H0）のイラストは Lucide の folder-search**（`illust/illust_first_run.svg`、色 `Illust.Line`）。
7. 領域の分け方（search_bar・target_tree・result_list・preview など）と x:Name は `figma_wpf_map.md` のとおり。PoC でも、できる範囲で領域ごとの XAML に分けておくと、本体に移すときにそのまま使える。
8. **レイアウトとリサイズ**（`figma_wpf_map.md` の「レイアウトとリサイズ」）: 窓は最小 1024×640。余白・間隔・揃え・Fill／Hug／Fixed・最小と最大・切る／折り返すを表のとおりにし、1024×640 と 1600×1000 でも崩れないようにする。
9. **ⓘ（info）は Lucide のままでは「i」が潰れるので、中だけ描き直す**（Figma も直した）。丸は Lucide のまま（直径 = 表示の大きさ × 20/24、線 = 表示の大きさ × 2/24）。中の「i」は、表示の大きさを s として、縦棒が上から 0.48s〜0.70s、点が上から 0.28s、どちらも左右の真ん中。線の太さは 1.5px（s = 16 のときは 1.75px）、端は Round。`Viewbox` で縮めると線が細るので、中の 2 本は表示の大きさのまま描く（点は直径 = 線の太さの塗った丸でよい）。`SnapsToDevicePixels` と `UseLayoutRounding` を付け、縦棒が 2 つの画素にまたがってぼけないようにする。s は 12・13・14・16 の 4 つ（ナビ・検索対象・プレビュー = 14、高速検索の表示 = 13、確認ダイアログの注 = 12、バナー = 16）。`circle-alert` など、ほかの丸のアイコンは変えない。
10. **アプリのアイコンは新しいロゴ**（`illust/app_logo.svg`。青・緑・橙の 3 段。256×256、塗りだけ・線なし）。前の「青い角丸の四角に file-search」と、起動画面の絵はやめる。
    - タイトルバー: 32×32 の枠はそのまま、地の色と角丸を無くし、ロゴを 22×22 で真ん中に置く。
    - 起動画面（SP）: 48×48。
    - バージョン情報（DV）: 上の大きいアイコン 48×48（薄い青の地と角丸を無くす）、タイトルの小さいアイコン 16×16。
    - 窓のアイコン（タスクバー・Alt+Tab に出るもの）も同じロゴにする（16・20・24・32・48・256 を入れた .ico）。
