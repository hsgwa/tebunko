# ステータスバー（status_bar）

## 1. 役割

主な領域のいちばん下の 1 行。いまの画面で最後に起きたこと（検索した・削除した・終了した）や、画面の数（登録済みのフォルダの数・合計のファイル数）を短く出す。左に文、右に合計の 2 か所を持つ。左の欄の下には付かない（`shell.md`）。

- Figma の部品: `xaml/shell/status_bar` 350:10437（既定 340:53）。
- variant（13）: 既定 340:53・H0 350:10398・H 350:10401・H-R 350:10420・X 350:10405・X-N 350:10409・X0 350:10413（106）、P 350:10417・P-H-E2 350:10423・P-E14 350:10426・P-H-saved 350:10428・P-X-del 350:10432・P-P-killed 350:10436（148）。
- 使う画面: すべての画面。見本は `../png/148_67.png`（H0）・`../png/108_230.png`（H）・`../png/135_1498.png`（H-R）・`../png/157_359.png`（H-E2）・`../png/157_1538.png`（E14）・`../png/158_527.png`（H-saved）。
- Figma の説明: 「x:Name: StatusText データ: 各画面の状態の文（toStatusText など）」。

## 2. 置き場所

- XAML: `scripts/tebunko/xaml/shell/status_bar.xaml`（ルートは `Border`。UserControl 相当）。
- 画面層: `scripts/tebunko/ui/shell/status_bar.ps1`（今の `shared/ui/shell.ps1` の `setStatus` をここに移すか、そのまま呼ぶかは 10 章）。判断層: `scripts/tebunko/ui/shell/status_bar_view.ps1`。
- 親: `tebunko.xaml` の `StatusBarHost`（主な領域の Row 1、高さ 24）。起動のときに 1 回だけ `loadXaml` して差し、x:Name を `$ui` に足す。

## 3. 部品の木

```xml
<Border x:Name="StatusBarRoot"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Height="24"
        Background="{StaticResource Bg.Hover}"
        BorderBrush="{StaticResource Border.Soft}" BorderThickness="0,1,0,0"
        Padding="12,0,12,0">
  <Grid VerticalAlignment="Center">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>     <!-- 左の文（StatusText） -->
      <ColumnDefinition Width="Auto"/>  <!-- 右の合計（StatusRightText） -->
    </Grid.ColumnDefinitions>
    <TextBlock x:Name="StatusText" Grid.Column="0"
               Style="{StaticResource Micro}" Foreground="{StaticResource Ink.Body}"
               HorizontalAlignment="Left" VerticalAlignment="Center"
               TextTrimming="CharacterEllipsis" TextWrapping="NoWrap"
               Text="{Status.Left}"/>
    <TextBlock x:Name="StatusRightText" Grid.Column="1" Margin="12,0,0,0"
               Style="{StaticResource Micro}" Foreground="{StaticResource Ink.Body}"
               HorizontalAlignment="Right" VerticalAlignment="Center"
               TextWrapping="NoWrap" Visibility="Collapsed"
               Text="{Status.Right}"/>
  </Grid>
</Border>
```

- `StatusRightText` は `figma_wpf_map.md` の x:Name の一覧に無い。Figma の X・P 系の variant に右の文があるので足す案（決めること）。
- 左と右の文の間の 12 は Figma に無い（Figma は両端寄せで、間の最小が決まっていない）。仮の値（決めること）。

## 4. 寸法と色

| 要素 | 大きさ | 余白 | 地の色 | 枠 | 文字の Style | 文字の色 | 角丸 |
|---|---|---|---|---|---|---|---|
| StatusBarRoot | 横 Fill × 24（1280 のとき 1060×24） | 左右 12 | 検索の画面: Bg.Hover（#F3F3F4）。インデックス管理・Office の終了: Bg.Subtle（#F9FAFA）（5 章） | 上 1px Border.Soft（#E0E2E5） | – | – | 0 |
| StatusText | Hug（入りきらないときは「…」） | – | – | – | Micro（10 Regular） | Ink.Body（#5F6368） | – |
| StatusRightText | Hug | – | – | – | Micro（10 Regular） | Ink.Body（#5F6368） | – |

- `figma_wpf_map.md` は使うキーを「Bg.Subtle・Border.Soft・Meta（Ink.Muted）」としているが、Figma の部品の値は文字が Micro（10）・Ink.Body、地は variant により Bg.Hover と Bg.Subtle。ここでは Figma の部品の値に合わせた（`figma_wpf_map.md` を直す案）。

## 5. 状態ごとの見え方

| 状態 | 入る条件（判断層の入力） | 地の色 | StatusText | StatusRightText |
|---|---|---|---|---|
| 既定（340:53）・H（350:10401） | 検索が終わった（`getSearchDoneStatus`。ワード `(株)山田商事`・15 件） | Bg.Hover | Status.SearchDone | 出さない |
| H-R（350:10420） | 同上（16 件） | Bg.Hover | Status.SearchDone | 出さない |
| P-E14（350:10426） | 同上（ワード `(株)山田商店`・0 件） | Bg.Hover | Status.SearchDone | 出さない |
| H0（350:10398） | 検索の画面で、インデックスが 1 つも無い | Bg.Hover | Status.NoIndex | 出さない |
| P-H-E2（350:10423） | 検索の画面で、まだ何も起きていない（文が空） | Bg.Hover | 空（Figma はゼロ幅の空白 U+200B） | 出さない |
| P-H-saved（350:10428） | 検索の結果を保存した（`../png/158_527.png`） | Bg.Hover | 空（Figma は文の要素ごと無い） | 出さない |
| X（350:10405） | インデックス管理の画面（4 フォルダ・2,530 ファイル） | Bg.Subtle | Status.Registered | Status.TotalFiles |
| X-N（350:10409） | 同上（1 フォルダ・0 ファイル。作ったばかり） | Bg.Subtle | Status.Registered | Status.TotalFiles |
| X0（350:10413） | 同上（0 フォルダ・0 ファイル） | Bg.Subtle | Status.Registered | Status.TotalFiles |
| P-X-del（350:10432） | インデックス管理で削除した（合計 2,285 ファイル） | Bg.Subtle | Status.Deleted | Status.TotalFiles |
| P（350:10417） | Office の終了の画面（3 件、うちバックグラウンド 2 件） | Bg.Subtle | Status.Office | 空（U+200B） |
| P-P-killed（350:10436） | Office の終了で終了した（2 件） | Bg.Subtle | Status.OfficeKilled | 空（U+200B） |

- 設定の画面の文は Figma に無い（未定）。
- 地の色が画面で違うのは Figma のとおり。どちらかに揃えるかは決めること。
- 文が空の状態でも、バーの高さ・地・上の線は残す（P-H-E2・P-H-saved）。
- P-H-saved で文を空にするのは Figma のとおり。保存したことを知らせる文を出すかは決めること（今は保存のときに文を出していない）。

## 6. 文言

| 文言の ID | 全文 | 使う状態 |
|---|---|---|
| Status.SearchDone | `検索しました：{ワード} {件数} 件`（「：」は全角。ワードと件数の間・件数と「件」の間は半角の空白。例 `検索しました：(株)山田商事 15 件`・`検索しました：(株)山田商店 0 件`） | 既定・H・H-R・P-E14 |
| Status.NoIndex | `インデックスが未作成です` | H0 |
| Status.Registered | `登録済み: {フォルダ数} フォルダ`（「:」は半角で、後ろに半角の空白。例 `登録済み: 4 フォルダ`） | X・X-N・X0 |
| Status.TotalFiles | `合計: {ファイル数} ファイル`（数は 3 桁ごとに「,」。例 `合計: 2,530 ファイル`・`合計: 0 ファイル`・`合計: 2,285 ファイル`） | X・X-N・X0・P-X-del |
| Status.Deleted | `削除しました` | P-X-del |
| Status.Office | `Office: {件数} 件（バックグラウンド {バックグラウンドの件数} 件）`（「:」は半角、括弧は全角。例 `Office: 3 件（バックグラウンド 2 件）`） | P |
| Status.OfficeKilled | `{件数} 件終了しました`（例 `2 件終了しました`） | P-P-killed |

- 件数の 3 桁ごとの「,」は、Figma の例が 3 桁以下の件数（14・16・0）では確かめられない。ファイル数（2,530）と同じく付ける（`ToString('N0')`。今の検索の文と同じ）。
- 今の文のうち、Figma に無いもの（検索の条件「条件：…」・打ち切り・中止・見つからない検索対象フォルダ・作成中のインデックスを検索している注意・セルのコピー・インデックスの作成の進み具合など）を、この形に合わせてどう書くかは未定（決めること）。キー操作のヒント（今の「Shift＋クリック・ドラッグで複数選べます」）は出さない。

## 7. 操作と振る舞い

| きっかけ | 何が起きるか | 呼ぶ関数 | 次の状態 |
|---|---|---|---|
| 検索が終わる | 左に検索の結果の文 | `setStatus (getSearchDoneStatus $word $count)` | H・H-R・P-E14 |
| 検索の画面を開いた・インデックスが無い | 左に「インデックスが未作成です」 | `setStatus (getNoIndexStatus)` | H0 |
| インデックス管理の画面を開いた・一覧が変わった | 左に登録済みの数、右に合計 | `setStatus` に `getIndexListStatus` の Left・Right | X・X-N・X0 |
| インデックスを削除した | 左に「削除しました」、右に合計（削除のあとの数） | `setStatus` に `getIndexDeletedStatus` の Left・Right | P-X-del |
| Office の終了の画面を開いた・一覧が変わった | 左に件数 | `setStatus (getOfficeStatus $count $background)` | P |
| Office を終了した | 左に終了した件数 | `setStatus (getOfficeKilledStatus $count)` | P-P-killed |
| 画面を切り替えた | その画面の最後の文に替える（画面ごとに覚えておく） | `applyStatusFor $screen` | 各画面の状態 |
| StatusText にカーソルを当てる | 文の全文をツールチップで出す（今と同じ。切れているときに読める） | – | – |

- バーはクリック・ホバーで見た目が変わらない。Tab で止まらない（`Focusable=False`）。
- 画面を切り替えたときに、前の画面の文を残すか替えるかは Figma に無い。上の表は、地の色が画面で違うことから、画面ごとに文を持つとした案（決めること）。

## 8. リサイズ

| 部分 | 動き |
|---|---|
| バー | 高さ 24 固定、横 Fill |
| StatusText | 1 行。入りきらないときは末尾を「…」で切る（`TextTrimming="CharacterEllipsis"`） |
| StatusRightText | 1 行。切らない（Auto の列）。左の文が先に縮む |

- 1024×640 のとき: バーの幅 804（左の欄 220 のとき）。1600×1000 のとき: 1380。
- 長い文（今の打ち切りの文など）は切れる。全文はツールチップで読む。

## 9. 判断層

`scripts/tebunko/ui/shell/status_bar_view.ps1`。文を組み立てる関数だけを置き、`$ui` に触らない。今の `toStatusText`（`preview_view.ps1`。文に入れる値を短くする）はそのまま使う。

| 関数 | 入力 | 出力 |
|---|---|---|
| `getSearchDoneStatus` | `[string]$word`・`[int]$count` | string |
| `getNoIndexStatus` | なし | string |
| `getIndexListStatus` | `[int]$folders`・`[long]$files` | hashtable `Left`（string）・`Right`（string） |
| `getIndexDeletedStatus` | `[long]$files`（削除のあとの合計） | hashtable `Left`・`Right` |
| `getOfficeStatus` | `[int]$count`・`[int]$background` | string |
| `getOfficeKilledStatus` | `[int]$count` | string |
| `getStatusBarBackground` | `[string]$screen` | string（theme のキー名 `Bg.Hover` / `Bg.Subtle`） |

```powershell
It "検索の文: <word> <count> 件 → <expected>" -TestCases @(
    @{ word = '(株)山田商事'; count = 15;   expected = '検索しました：(株)山田商事 15 件' }
    @{ word = '(株)山田商事'; count = 16;   expected = '検索しました：(株)山田商事 16 件' }
    @{ word = '(株)山田商店'; count = 0;    expected = '検索しました：(株)山田商店 0 件' }
    @{ word = '見積';         count = 1234; expected = '検索しました：見積 1,234 件' }
) {
    param ($word, $count, $expected)
    getSearchDoneStatus $word $count | Should -Be $expected
}

It "インデックスが無いときの文" {
    getNoIndexStatus | Should -Be 'インデックスが未作成です'
}

It "インデックス管理の文: <folders> フォルダ・<files> ファイル" -TestCases @(
    @{ folders = 4; files = 2530; left = '登録済み: 4 フォルダ'; right = '合計: 2,530 ファイル' }
    @{ folders = 1; files = 0;    left = '登録済み: 1 フォルダ'; right = '合計: 0 ファイル' }
    @{ folders = 0; files = 0;    left = '登録済み: 0 フォルダ'; right = '合計: 0 ファイル' }
) {
    param ($folders, $files, $left, $right)
    $status = getIndexListStatus $folders $files
    $status.Left | Should -Be $left
    $status.Right | Should -Be $right
}

It "削除のあとの文" {
    $status = getIndexDeletedStatus 2285
    $status.Left | Should -Be '削除しました'
    $status.Right | Should -Be '合計: 2,285 ファイル'
}

It "Office の文: <count> 件（バックグラウンド <background> 件）" -TestCases @(
    @{ count = 3; background = 2; expected = 'Office: 3 件（バックグラウンド 2 件）' }
    @{ count = 0; background = 0; expected = 'Office: 0 件（バックグラウンド 0 件）' }
) {
    param ($count, $background, $expected)
    getOfficeStatus $count $background | Should -Be $expected
}

It "Office を終了したあとの文" {
    getOfficeKilledStatus 2 | Should -Be '2 件終了しました'
}

It "画面 <screen> の地の色は <key>" -TestCases @(
    @{ screen = 'search';   key = 'Bg.Hover' }
    @{ screen = 'index';    key = 'Bg.Subtle' }
    @{ screen = 'office';   key = 'Bg.Subtle' }
) {
    param ($screen, $key)
    getStatusBarBackground $screen | Should -Be $key
}
```

- `Office: 0 件（バックグラウンド 0 件）` の形は Figma に無い（0 件のときの文は未定。表は今の形を延ばした仮の期待値）。
- 設定の画面の地の色は Figma に無い（未定のため表に入れない）。

## 10. 画面層

`scripts/tebunko/ui/shell/status_bar.ps1`

`setStatus` を 2 か所に書けるように広げる。今の呼び方（`setStatus "文"`）はそのまま使える。

```powershell
function setStatus {
    param (
        [string]$text,
        [string]$right = ""
    )

    $ui.StatusText.Text = $text
    $ui.StatusText.ToolTip = if ($text) { $text } else { $null }
    $ui.StatusRightText.Text = $right
    $ui.StatusRightText.Visibility = if ($right) { 'Visible' } else { 'Collapsed' }
}
```

| x:Name | 配線 |
|---|---|
| StatusBarRoot | 画面を切り替えたとき（`showScreen`）に `Background` を `$window.FindResource((getStatusBarBackground $screen))` にする |
| StatusText | `setStatus` が書く。ToolTip は文と同じ（空なら出さない） |
| StatusRightText | `setStatus` の 2 つ目の引数で書く。空なら `Collapsed` |

- `setStatus` は今 `shared/ui/shell.ps1` にあり、どのツールからも使う。`StatusRightText` はこのツールの XAML にしか無いので、2 つ目の引数を足した版は `tebunko/ui/shell/status_bar.ps1` に置き、`shared` の版は残すか、`StatusRightText` が無ければ書かないようにする（決めること）。
- `setStatus` は UI スレッドで呼ぶ。別のスレッドの処理（検索・削除・Office の一覧）の終わりは、今と同じくタイマーか `$window.Dispatcher.Invoke` で UI スレッドに戻ってから呼ぶ。

## 11. 受け入れ

見比べる画像: `../png/148_67.png`（H0）・`../png/108_230.png`（H）・`../png/135_1498.png`（H-R）・`../png/157_1538.png`（E14）・`../png/158_527.png`（H-saved）。窓 1280×820。インデックス管理・Office の終了は Figma の status_bar の variant（X・X-N・X0・P・P-X-del・P-P-killed）のスクリーンショットと見比べる。

- バーの高さが 24、上に 1px の線（#E0E2E5）。左の欄の下には無く、x=220 から右端まで。
- 検索の画面の地が #F3F3F4、インデックス管理・Office の終了の地が #F9FAFA。
- 文が左から 12 の所に 10px・#5F6368 で、上下の中央。
- インデックス管理では右から 12 の所に「合計: 2,530 ファイル」が右寄せで出る。
- 文が空でもバーの高さと線が変わらない。
- 文言が 6 章の全文と、全角・半角・空白まで同じ。

## 決めること

- 地の色を画面で変えるか（Figma は検索 Bg.Hover、インデックス管理・Office の終了 Bg.Subtle）、どちらかに揃えるか。設定の画面の地の色。
- 右の文の x:Name `StatusRightText` を足すこと（`figma_wpf_map.md` に無い）と、左と右の文の間の最小の間（仮に 12）。
- 画面を切り替えたとき、文を画面ごとに覚えて替えるか、最後の文を残すか。
- 検索の結果を保存したとき（P-H-saved）に文を出すか（Figma は空）。
- 設定の画面の文、Office が 0 件のときの文。
- 今の文のうち Figma に無いもの（検索の条件・打ち切り・中止・見つからない検索対象フォルダ・作成中のインデックスの注意・インデックスの作成の進み具合・取り込み・書き出し・セルのコピーなど）の新しい書き方。
- `setStatus` の置き場所（`shared/ui/shell.ps1` のままにして右の文を受けるか、`tebunko/ui/shell/status_bar.ps1` に移すか）。
- `figma_wpf_map.md` の使うキー（Bg.Subtle・Meta（Ink.Muted））を、Figma の部品の値（Bg.Hover／Bg.Subtle・Micro・Ink.Body）に直すこと。
