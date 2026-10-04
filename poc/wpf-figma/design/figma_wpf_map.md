# Figma の領域と WPF の対応表

Figma のページ「08 部品」にある領域のコンポーネントと、WPF の実装（XAML のファイル・x:Name・theme のキー）の対応。

- 領域のコンポーネントは、ページ 106 の今の画面（108:230・108:510・108:776・109:1638・ダイアログ）を複製して作った。見た目は 106 と同じ。
- shell 以外のコンポーネントは、variant のセットにしてある。「既定」は複製の元にした状態。ほかの variant は、106・148 の画面で使っている状態で、`状態=<画面の記号>`（148 の画面は `状態=P-<記号>`）の名前にした。106・148 の画面の該当する部分は、すべてこの variant のインスタンスに置き換えてある。
- 表の ID は、セットの ID と既定の variant の ID（`セット / 既定`）。
- x:Name の欄で **太字** のものは新しく足した名前（今の XAML に無い）。それ以外は今の XAML の名前を引き継いだ。
- theme のキーは、Figma の色に一番近い今の `theme.xaml` のキー。Figma の色の値は theme と一致していない（下の「食い違い」）。

## 窓（shell）

| Figma のコンポーネント | XAML のファイル | x:Name | 使う theme のキー |
|---|---|---|---|
| `xaml/shell/shell` 340:1650（variant なし） | `scripts/tebunko/xaml/tebunko.xaml`（Window） | **NavHost**・**ContentHost**・**StatusBarHost** | Bg.Window・Bg.Surface・Border.Soft・Ink.Strong |
| `xaml/shell/nav` 353:30193 / 340:553 | `scripts/tebunko/xaml/shell/nav.xaml` | SearchTab・IndexTab・SettingsTab・KillTab・**IndexTabBadge**・**NavPaneHost**・**AboutLink** | Bg.Subtle・Bg.Hover・Select.Soft・Accent・Ink.Body・Ink.Muted・Border.Soft |
| `xaml/shell/status_bar` 350:10437 / 340:53 | `scripts/tebunko/xaml/shell/status_bar.xaml` | StatusText | Bg.Subtle・Border.Soft・Meta（Ink.Muted） |

- 今の WPF は `Tabs`（TabControl）と `IndexTabHeader`・`KillTabHeader`・`MoreButton`・`AboutMenuItem` で組んでいる。Figma は左の欄のナビなので、`SearchTab` などの名前はナビの項目に付けた。
- `NavPaneHost` は、検索の画面のとき `xaml/search/target_tree` を差す口。nav の検索の variant（検索・検索-空・P-H 系・P-E13）は、NavPaneHost の中に target_tree のインスタンスを持つ。
- 106 の左の欄（前の `Nav Pane`）も nav の variant に置き換えた。`Nav Pane` のセットは消した。
- variant:
  - nav: 既定・検索・検索-空・インデックス管理・設定・Office の終了（106）、P-H0・P-H・P-H-B・P-H-P・P-E13・P-X0・P-X-U・P-X-P・P-C・P-P（148）
  - status_bar: 既定・H0・H・H-R・X・X-N・X0（106）、P・P-H-E2・P-E14・P-H-saved・P-X-del・P-P-killed（148）

## 検索（search）

| Figma のコンポーネント | XAML のファイル | x:Name | 使う theme のキー |
|---|---|---|---|
| `xaml/search/target_tree` 366:313 / 340:564 | `scripts/tebunko/xaml/search/target_tree.xaml` | SearchTargetText・CheckAllIndexButton・UncheckAllIndexButton・**IndexTreeFilterBox**・IndexTree・IndexTreePlaceholder | Heading・Link（Accent）・TextBox.Base・Placeholder・TreeViewItem.Base・TreeChevron・Choice・Ok・Warn・Danger.Text |
| `xaml/search/search_bar` 348:9775 / 340:610 | `scripts/tebunko/xaml/search/search_bar.xaml` | WordBox・SearchButton・**FileKindChips**・**ScopeButton**・CaseCheck・RegexCheck・FastSearchText | Label・TextBox.Base・Primary・ToolButton・Accent.Soft・Accent・Meta・Ok |
| `xaml/search/result_list` 352:18716 / 340:723 | `scripts/tebunko/xaml/search/result_list.xaml` | SummaryText・ExpandAllButton・CollapseAllButton・FilterBox・FilterPlaceholder・ExportButton・ResultGrid | Focal・Link・TextBox.Base・Placeholder・ToolButton・Cell・Cell.Key・Cell.Meta・Hit・Select.Soft・Accent・Border.Soft・Bg.Subtle |
| `xaml/search/preview` 353:33742 / 340:954 | `scripts/tebunko/xaml/search/preview.xaml` | DetailTitle・OpenButton・OpenModeCombo・OpenFolderButton・PreviewHeader・PreviewRows・PreviewNote・PreviewPlaceholder | Heading・Primary・ToolButton・Cell・Cell.Number・Hit・Border.Soft・Bg.Subtle・Meta・Placeholder |

- variant:
  - target_tree: 既定・空・P-E13
  - search_bar: 既定・H0・H・H-B・H-R・H-範囲2・E13・thumb-H-R（106）、P-H・P-H-R・P-H-B・P-H-E・P-H-1・P-E12・P-E13・P-E14（148）
  - result_list: 既定・H・H-B・H-W・H-R・H-F・thumb・thumb-2・thumb-3（106）、P-H・P-H-E2・P-H-S・P-H-saved・P-H-row3・P-H-row4・P-E12・P-E14・P-E16・P-H-folded・P-H-open（148）、H-PP・H-TX
  - preview: 既定・H-W・thumb（106）、P-H・P-H-row3・P-H-row4（148）
- 108:230（H 検索）と 148:138 では、preview の［開く ▾］のメニューを開いた状態で見せている（意図どおり）。
- `ExportButton` は今の XAML ではプレビューの側にあるが、Figma では結果の上の「保存」にある。
- `PreviewNote`・`PreviewPlaceholder` は、「02 画面のバリエーション」の Preview（324:8867）の別の kind（列を省いた・読めない・選んでいない）にある。`xaml/search/preview` は kind=Excel-セル から作った。
- 今の XAML にある `ShapeCheck`・`CommentCheck`・`FileFilterBox`・`FileFilterPlaceholder`・`WordNotice`・`GoIndexTabButton`・`SearchProgress`・`DetailPanel` に当たる層は、この 4 つのコンポーネントに無い（Figma では種別のチップ・ファイル内の対象・絞り込みにまとめた）。

## インデックス管理（index）

| Figma のコンポーネント | XAML のファイル | x:Name | 使う theme のキー |
|---|---|---|---|
| `xaml/index/index_list` 352:29841 / 340:1100 | `scripts/tebunko/xaml/index/index_list.xaml` | IndexingButton・ImportIndexButton・NewIndexButton・IndexGrid | Heading・Meta・Button.Base・Primary・Cell・Cell.Key・Cell.Number・Choice・Ok・Warn・Danger.Text・Danger.Soft・Accent.Soft・Border.Soft |
| `xaml/index/index_detail` 352:39102 / 340:1203 | `scripts/tebunko/xaml/index/index_detail.xaml` | **IndexDetailTitle** | Heading・Label・Meta・Cell・Border.Soft・Bg.Surface |

- variant:
  - index_list: 既定・X・X-P・X-C・X-N・X0・X-⋯・X-エクスポート中・X-取り込み後・X-取り込み後2・X-高速検索-反映中・インデックス管理（高速検索が使えない）・thumb〜thumb-10（106）、P-X0・P-X-N・P-X・P-X-U・P-X-P・P-X-C・P-X-S・P-X-顧客・P-X-del・P-X-refolder・P-X-exporting・P-X-imported（148）
  - index_detail: 既定・X・X-C・X-N・X0・X-取り込み後・X-取り込み後2・X-高速検索-反映中・インデックス管理（高速検索が使えない）・thumb〜thumb-6（106）、P-X-edit（148）
- 今の XAML の `EditIndexButton`・`RemoveIndexButton`・`ExportIndexButton` は、Figma では行の「…」（mini-button-more）の中。`IndexingProgressPanel` などの取り込み中の表示・`FailedPanel` は、このコンポーネントに無い。

## 設定（settings）

| Figma のコンポーネント | XAML のファイル | x:Name | 使う theme のキー |
|---|---|---|---|
| `xaml/settings/settings` 352:44948 / 340:1241 | `scripts/tebunko/xaml/settings/settings.xaml` | WorkspaceText・ChangeWorkspaceButton・ResetWorkspaceButton・**MaxParallelCombo**・**PageSizeCombo** | Heading・Label・Meta・Card・Button.Base・Bg.Surface・Border.Soft |

- variant: 既定・C（106）、P-C-moved・P-C-reset（148）
- 今の XAML の `WorkspaceNote`・`SettingsFileText`・`SettingsFileNote` は Figma に無い。`MaxParallelCombo`・`PageSizeCombo` は Figma にあって WPF に無い。

## Office の終了（office）

| Figma のコンポーネント | XAML のファイル | x:Name | 使う theme のキー |
|---|---|---|---|
| `xaml/office/office` 374:50 / 340:1438 | `scripts/tebunko/xaml/office/office.xaml` | KillSelectedButton・KillAllButton・ProcessGrid・ProcessSummaryText | Heading・Meta・Button.Base・Danger・Cell・Choice・Border.Soft |

- variant: 既定、P（106 の 109:1638）、E60（381:76。止めるもの（動いている Office）が無いとき。106 の sample E60 117:1303 で使う）、P-P・P-P-killed（148）
- E60 は、表の行を隠して中央に `parts/illust_no_office` と 2 行（「動いている Office はありません」「Excel・Word・PowerPoint が起動していると、ここに表示されます。」）を出し、ボタンを 2 つとも無効にして、その下に「終了するものがありません」（**KillReasonText**）を出す。
- Office の終了の画面では、中央の部分（前は index_list・splitter・index_detail で組んでいた）を、この 1 つのインスタンスに置き換えた。
- `ProcessSummaryText` に当たる文は、Figma ではステータスバーにある（コンポーネントからは外した）。`KillBackgroundButton`・`RefreshProcessButton` に当たる層は無い。

## ダイアログ（dialog）

| Figma のコンポーネント | XAML のファイル | x:Name | 使う theme のキー |
|---|---|---|---|
| `xaml/dialog/about` 372:54 / 340:1489 | `scripts/tebunko/xaml/dialog_about.xaml` | CloseButton（AppIcon・VersionText・CommitText は app-info-block の中。層の名前は付けていない） | DialogHeading・Meta・Link・DialogButton |
| `xaml/dialog/indexing_confirm` 372:669 / 340:1544 | `scripts/tebunko/xaml/dialog_indexing_confirm.xaml` | StartButton・CancelButton・PlanGrid | DialogHeading・Meta・Cell・Primary・DialogButton・Warn |
| `xaml/dialog/export` 372:1008 / 340:1567 | `scripts/tebunko/xaml/dialog_export.xaml`（新しいファイル） | **ExportPathBox**・CancelButton・**ExportButton** | DialogHeading・Label・TextBox.Base・Warn・Primary・DialogButton |
| `xaml/dialog/import` 372:1112 / 340:1597 | `scripts/tebunko/xaml/dialog_import.xaml`（新しいファイル） | NameBox・FolderBox・CancelButton・**ImportButton** | DialogHeading・Label・TextBox.Base・Primary・DialogButton |
| `xaml/dialog/confirm` 373:50（既定は Type ごと。下の variant） | `scripts/shared/xaml/dialog_confirm.xaml` | ButtonPanel・HeadingText・FactsPanel・FactsList・FactDetail・ChoicePanel・HintText（層の名前はまだ付けていない） | DialogHeading・Meta・Hint・Primary・Danger・DialogButton |

- variant:
  - about: 既定、P-DV（148）
  - indexing_confirm: 既定、P-DC-start・P-DC-start-1（148）
  - export: 既定だけ（106・148 の DX）
  - import: 既定、DI-2（106・148）・DI-3・DI-4（106）、P-DI-1（148）
  - confirm は `Type`（normal・danger・choice・progress・error）と `状態` の 2 つの項目で表す。Type ごとの既定は normal 373:51・danger 373:66・choice 373:83・progress 373:100・error 373:115。
    - normal: DC-stop・DC-kill・DC-ws・DC-reset・DC-close・DC-src（106）、P-DC-stop・P-DC-ws・P-DC-reset・P-DC-kill・P-DC-close・P-DC-refolder（148）
    - danger: DC-del・DC-kill-vis（106）、P-DC-del（148）
    - choice: DC-ws-idx（106）
    - progress: DC-stopping・DC-close-wait・DC-move（106）
    - error: DC-err・E54・DI-E1・DI-E2・DI-E3（106。DI-E は インポートのエラー。148 には無い）
- 106 の前の `Confirm Dialog` のセットは、使っている所を無くして消した。
- about は、層の順がタイトルバーより中身が先（106 の元の画面のまま）。
- `dialog_index_edit.xaml`（OkButton・NameBox・FolderBox・BrowseButton など）に当たるダイアログの領域は、今回は作っていない（106 のどの画面が当たるかを確かめていない）。

## 部品（parts）

| Figma のコンポーネント | WPF での形 | 使う画面 | 使う theme のキー |
|---|---|---|---|
| `parts/illust_first_run` 346:9181 | Path／Canvas（theme.xaml の DrawingImage か Canvas） | インデックスが無いとき（H0 の検索・X0 のインデックス管理の Empty State） | Illust.Line |
| `parts/illust_no_office` 379:35945 | Path／Canvas（theme.xaml の DrawingImage か Canvas） | Office の終了で、止めるもの（動いている Office）が無いとき（E60） | Illust.Line |

- どちらも 64×64、Lucide のアイコン（lucide-static 1.50.0、ISC）を 24 から 64 に拡大したもの。線 2px・塗りなし・端と角は丸。illust_first_run は folder-search、illust_no_office は circle-check。
- `parts/illust_no_office` は、`xaml/office/office` の 状態=E60 の中で使う。
- `parts/illust_first_run` は、106 の thumb H0 の empty-state（111:1092）でも、16×16 に縮めたインスタンス（層の名前 empty-icon）として使う。

## ファイルの種類のアイコン

Lucide（lucide-static 1.50.0）。線の太さは画面の Lucide のアイコン全部で「表示の大きさ × 2/24」（10px → 0.83、12px → 1.0、14px → 1.17、16px → 1.33）。端と角は丸、塗りなし。

| 種類 | アイコン名（層の名前） | 色の変数 | 大きさ・線の太さ |
|---|---|---|---|
| Excel | file-spreadsheet（折りたたんだ見出しでは chart-column） | File/Excel #107C41 | 14px・1.17 |
| Word | file-text | File/Word #185ABD | 14px・1.17 |
| PowerPoint | presentation | File/PowerPoint #C43E1C | 14px・1.17 |
| テキスト（.txt・.md など） | file | Ink/Body #5F6368 | 14px・1.17 |
| フォルダ（インデックス一覧の行・ツリー） | folder | File/Folder #E8A020 | 14px・1.17 |

- 検索結果の見出し（`xaml/search/result_list` の file-header）で使う。種類のチップ（FileKindChips）はアイコンを持たない（チェックと文言だけ）。

## レイアウトとリサイズ（2026-10-03）

Figma の部品は Auto layout で組み、余白・間隔・揃え・最小と最大を持たせてある。WPF では、方向 = StackPanel の Orientation（Fill の子があれば Grid）、Fill = `*`、Hug = `Auto`、Fixed = 数値、padding = Padding、gap = 子の Margin、揃え = HorizontalAlignment／VerticalAlignment、「…」で切る = `TextTrimming="CharacterEllipsis"`、折り返す = `TextWrapping="Wrap"`（行の折り返しは WrapPanel）にする。部品ごとの注記（Figma の annotation）にも同じことを書いてある。見本は「08 部品」のセクション「リサイズの見本」（H・H0・H-row3・X・C・DC-ws を 1024×640 と 1600×1000 で）。

| 領域 | 方向 | 揃え | padding | gap | 子の大きさ | 最小／最大 | リサイズのときの動き | 切る／折り返す |
|---|---|---|---|---|---|---|---|---|
| 窓 | 縦 | 左上 | 0 | 0 | タイトルバー: 横 Fill・高さ 32。本体: Fill | 最小 1024×640、既定 1280×820 | 本体が伸びる | – |
| タイトルバー | 横 | 両端・上下の中央 | 左右 8 | 0 | 左: Hug、右の窓ボタン: Hug | 高さ 32 | 間が伸びる | – |
| 本体 | 横 | 左上 | 0 | 0 | 左の欄: 固定、主な領域: Fill | 左の欄: 既定 220・最小 180・最大 360 | 主な領域が伸びる。左の欄は右端の線（1px）とつまみでドラッグ（GridSplitter） | – |
| 左の欄 nav | 縦 | 左上 | 上 8 | 0 | 項目の一覧: Hug（下 8）、ツリー: Fill、下のヘルプ: Hug | 上の行のとおり | ツリーが縦に伸びる。ツリーの無い状態は空き（spacer）が伸びる | – |
| 検索対象ツリー | 縦 | 左上 | 元のまま | 元のまま | 見出し・絞り込み・ツリー: 横 Fill。ツリー: 縦 Fill。開閉の印: 12 固定 | – | 縦にスクロール | 名前は 1 行で「…」。状態の札がある行は札を名前の直後に置く |
| 主な領域 | 縦 | 左上 | 0 | 0 | 検索バー: Hug、一覧: Fill、境目: 5、プレビュー: 固定、ステータスバー: 24 | 一覧: 最小の高さ 120 | 一覧が縦に伸びる | – |
| 検索バー | 縦 | 左上 | 上 12・左右 16・下 10 | 8 | 1 行目: ワード欄 Fill・検索ボタン 100 固定（間 12）。2 行目: Fill | 高さ Hug | 2 行目は入りきらないと次の行へ（行の間 8）。幅 804 では高さ 86 → 114 | 折り返す |
| 結果一覧の件数の行 | 横 | 両端・上下の中央 | 上下 4 | 0 | 件数: Hug、すべて開く・すべて折りたたむ: Hug（間 12）、絞り込み・保存: Hug | 高さ Hug（既定 32） | 入りきらないと右の塊を次の行へ（行の間 4） | 折り返す |
| 結果一覧（ファイルの行・該当の行） | 横 | 左・上下の中央 | 元のまま | 該当の行 12、ファイルの見出し 8 | 場所 180 固定、種別 48 固定、該当の文: Fill | 最小の高さ 120 | 縦にスクロール | 場所は 1 行で「…」。該当の文とファイルの見出しは 1 行で切る |
| プレビュー | 縦 | 左上 | 0 | 0 | ツールバー: 横 Fill、表: Fill | 既定の高さ 180、最小 120 | 境目 5 のドラッグで高さを変える。表の列は固定の幅で、入りきらないと横・縦にスクロール | ファイル名は 1 行で「…」 |
| プレビューのツールバー | 横 | 両端・上下の中央 | 上下 8・左右 16 | 16 | ファイル名: Fill、ボタン: Hug（間 12） | – | ファイル名が縮む | 1 行で「…」 |
| ステータスバー | 横 | 両端・上下の中央 | 左右 12 | 0 | 文言: Hug | 高さ 24 | 横 Fill | 1 行 |
| 空の状態（H0 など） | 縦 | 上下左右の中央 | 0 | 16 | 中身の塊: Hug | 塊は最大幅 480 | 塊が残りの領域の中央を保つ | 説明は折り返す |
| インデックス管理の一覧 | 縦 | 左上 | 元のまま | 見出しの中 16 | 見出しの題と説明: Fill、ボタン: Hug、表: Fill | 一覧: 最小の高さ 225（見出し＋3 行）。表: 最小幅 991。「場所」の列: 最小 160 | 窓を高くすると一覧が伸びる。低くすると先に詳細欄が縮み、一覧は 225 で止まる。「場所」の列だけ伸びる。幅が足りないと横に、行が多いと縦にスクロール | 説明は折り返す。パスは 1 行で「…」 |
| インデックスの詳細 | 縦（一覧の下） | 左上 | 元のまま | 元のまま | 横 Fill、高さは既定 397 | 最小の高さ 240、最大 397。進み具合の線: 60〜100 | 境目 5px のドラッグで高さを変える。窓を低くすると一覧より先に 240 まで縮み、中は縦にスクロール。窓を高くしても 397 のまま。進み具合は線が伸び縮みする | 値の行は 1 行で「…」。進み具合の文は切らず、入らないときは線の下に折り返す（間 4） |
| 設定 | 縦 | 左上 | 見出し: 上下 16・左右 24。本文: 左右・下 24 | 節の間 24、行の間 12 | 行: Fill、入力欄: Fill、ボタン: Hug | – | 縦にスクロール | 説明は折り返す |
| Office の終了 | インデックス管理の一覧と同じ | 同 | 同 | 同 | 伸びる列だけ Fill（最小 160） | 表の最小幅 1047 | 横・縦にスクロール | 1 行で「…」 |
| ダイアログ | 縦 | 窓の上下左右の中央 | 題: 上下 10・右 12・左 16。本文: 上 16・左右 20・下 20 | 本文 12、ボタン 8 | 幅固定（520。種類により 620・440、バージョン情報は 460）、高さ Hug | 高さの最小は今の高さ | 暗幕は窓いっぱい。ダイアログは中央を保つ | 題は 1 行。本文は折り返す。パスは 1 行で「…」。ボタンは右寄せ |
| 重ねて出すメニュー | 重ね | – | – | – | Hug | – | 開くメニュー: ボタンの右下に付く。行のメニュー: 右上。範囲のメニュー: 左上 | – |

- **縮む順番は WPF のコードで決める。** インデックス管理は、窓が低いときに詳細欄（397 → 240）を先に縮め、一覧は 225 で止める。Grid の固定の高さの行は先に縮まないので、`SizeChanged` で詳細欄の行の高さを決めるか、一覧の行を `*`（MinHeight 225）、詳細欄の行を `Auto` にして中身の MaxHeight 397・MinHeight 240 で持たせる。
- 表の最小幅（インデックス管理 991・Office の終了 1047）は主な領域（最小 804）より広い。最小の窓では表が横にスクロールする。

## 食い違い（Figma にあって theme に無いもの）

- 色は Figma のコレクション「theme」の Variables にした（キーと値は `figma_theme_tokens.md`）。部品の色は、値が一致する Variable につないである。
- 色の値は今の `theme.xaml` と一致しない（Figma を正にする。差は `figma_theme_tokens.md` の「今の theme」の欄）。
- theme にキーが無い色: Excel の緑 `#107C41`・Word の青 `#185ABD`・PowerPoint の橙 `#C43E1C`・フォルダの黄 `#E8A020`・Ok の淡い背景 `#E0F7E0`・Warn の淡い背景 `#FFF5E0`・`#000000`。
- 文字: Figma の 10px・18px Bold と、Medium の太さは theme の Style に無い。
