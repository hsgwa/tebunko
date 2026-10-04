# インデックス管理の一覧（index_list）

## 1. 役割

検索するフォルダとインデックスを並べて管理する、インデックス管理の画面の上半分。行ごとにインデックスの状態と高速検索の状態を出す。そこから更新・中止・再設定、行の［⋯］からエクスポート・削除ができる。追加とインポートは見出しのボタンで行う。行を選ぶと、下の詳細欄（`index_detail.md`）にそのインデックスが出る。名前とフォルダパスの設定は詳細欄の「基本設定」だけで行い、この一覧には置かない。

- Figma の部品: `xaml/index/index_list` 352:29841（既定 340:1100）。
- variant（106）: 既定・X 352:24953・X-P 352:25108・X-C 352:25261・X-N 352:25331・X0 352:25481・X-⋯ 352:27313・X-エクスポート中 352:27470・X-取り込み後 352:27624・X-取り込み後2 352:27778・X-高速検索-反映中 352:27932・インデックス管理（高速検索が使えない）352:25746・thumb〜thumb-10。
- variant（148）: P-X0 352:28082・P-X 352:28306・P-X-U 352:28460・P-X-S 352:28921・P-X-顧客 352:29187・P-X-del 352:29313・P-X-refolder 352:29529・P-X-exporting 352:29686・P-X-imported 352:29840。148 で使う名前は P-X-N・P-X-P・P-X-C もある。
- 行のメニュー（［⋯］）は部品の外にあり、画面の上に重ねて出す。
  - 106: X-⋯（221:6314 の row-menu）、作成を中止したインデックスの row-menu（222:6153）、Disabled Reason（239:8525）。
  - 148: ov ⋯menu（224:8249）。
- リサイズの見本: X 1024×640（408:33785）・X 1600×1000（408:33812）。
- `../png/` にこの領域の画像は無い。見比べは Figma の get_screenshot で行う（11 章）。

## 2. 置き場所

- XAML: `scripts/tebunko/xaml/index/index_list.xaml`。
- 画面層 `ui/index/index_list.ps1`、判断層 `ui/index/index_list_view.ps1`。
- 親は、インデックス管理の画面の主な領域（shell の `ContentHost` に差す画面の Grid）。
  - その Grid は 4 行: 一覧・境目・詳細欄・（ステータスバーは shell の側）。
  - 一覧は 0 行目に置く（8 章）。
- 読み込み: 起動のときに 1 回だけ `XamlReader::Load` で読む。インデックス管理の画面の Grid の `Grid.Row="0"` に入れ、`FindName` で x:Name を引く。画面を切り替えても作り直さない。

```xml
<!-- インデックス管理の画面（主な領域）。境目と詳細欄の扱いは 8 章 -->
<Grid x:Name="IndexPage">
  <Grid.RowDefinitions>
    <RowDefinition Height="*" MinHeight="225" />   <!-- index_list -->
    <RowDefinition Height="5" />                    <!-- IndexSplitter -->
    <RowDefinition Height="Auto" />                 <!-- index_detail（MinHeight 240・MaxHeight 397 は中身に持たせる） -->
  </Grid.RowDefinitions>
  <ContentControl x:Name="IndexListHost" Grid.Row="0" />
  <GridSplitter x:Name="IndexSplitter" Grid.Row="1" Height="5" HorizontalAlignment="Stretch"
                ResizeDirection="Rows" ResizeBehavior="PreviousAndNext" Background="{DynamicResource Bg.Subtle}"
                BorderBrush="{DynamicResource Border.Soft}" BorderThickness="0,1,0,1" />
  <ContentControl x:Name="IndexDetailHost" Grid.Row="2" />
</Grid>
```

- 境目（`IndexSplitter`）は部品の外にあり、リサイズの見本で読み取った（8 章）。
  - 地は Bg.Subtle（#F9FAFA）、線は Border.Soft（#E0E2E5）。
  - 中央のつまみは 32×2、角丸 1、Ink.Body（#5F6368）。
  - 線がどの辺に付くかは未定（「決めること」）。

## 3. 部品の木

```xml
<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
      x:Name="IndexListRoot" Background="{DynamicResource Bg.Surface}">
  <Grid.RowDefinitions>
    <RowDefinition Height="Auto" />  <!-- 見出し -->
    <RowDefinition Height="*" />     <!-- 表（または空のとき） -->
  </Grid.RowDefinitions>

  <!-- 見出し（header-section） -->
  <Border Grid.Row="0" Padding="16" BorderBrush="{DynamicResource Border.Soft}" BorderThickness="0,0,0,1">
    <Grid>
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*" />
        <ColumnDefinition Width="Auto" />
      </Grid.ColumnDefinitions>
      <StackPanel Grid.Column="0" Orientation="Vertical" Margin="0,0,16,0">
        <StackPanel Orientation="Horizontal">
          <TextBlock Style="{StaticResource PageTitle}" Foreground="{DynamicResource Ink.Strong}" Text="{index.list.title}" />
          <ContentControl x:Name="IndexListHelpIcon" Width="14" Height="14" Margin="6,0,0,0" VerticalAlignment="Center"
                          ToolTip="{index.list.help}" /> <!-- help-icon -->
        </StackPanel>
        <TextBlock Style="{StaticResource Cell}" Foreground="{DynamicResource Ink.Body}" Margin="0,4,0,0"
                   TextWrapping="Wrap" Text="{index.list.description}" />
      </StackPanel>
      <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
        <Button x:Name="ImportIndexButton" Style="{StaticResource Button.Base}" Height="32" Padding="14,0"
                Content="{index.list.import}" />
        <Button x:Name="NewIndexButton" Style="{StaticResource Primary}" Margin="8,0,0,0" Padding="16,7"
                Content="{index.list.new}" />
      </StackPanel>
    </Grid>
  </Border>

  <!-- 表。横は 991 より狭いとスクロール、縦は行が多いとスクロール -->
  <ScrollViewer x:Name="IndexGridScroll" Grid.Row="1" HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Disabled">
    <Grid MinWidth="991">
      <Grid.RowDefinitions>
        <RowDefinition Height="28" /> <!-- 列の見出し -->
        <RowDefinition Height="*" />  <!-- 行 -->
      </Grid.RowDefinitions>
      <!-- 列の見出し（grid-headers）。列の幅は行と同じ SharedSizeGroup ではなく固定値でそろえる -->
      <Border Grid.Row="0" Background="{DynamicResource Bg.Subtle}" BorderBrush="{DynamicResource Border.Soft}"
              BorderThickness="0,0,0,1" Padding="12,0">
        <Grid x:Name="IndexGridHeader">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="13" />                  <!-- チェック -->
            <ColumnDefinition Width="12" />
            <ColumnDefinition Width="120" />                 <!-- インデックス名 -->
            <ColumnDefinition Width="12" />
            <ColumnDefinition Width="*" MinWidth="160" />    <!-- パス -->
            <ColumnDefinition Width="12" />
            <ColumnDefinition Width="60" />                  <!-- ファイル数 -->
            <ColumnDefinition Width="12" />
            <ColumnDefinition Width="120" />                 <!-- 最終更新 -->
            <ColumnDefinition Width="12" />
            <ColumnDefinition Width="164" />                 <!-- ステータス -->
            <ColumnDefinition Width="12" />
            <ColumnDefinition Width="96" />                  <!-- 高速検索 -->
            <ColumnDefinition Width="12" />
            <ColumnDefinition Width="150" />                 <!-- 操作 -->
          </Grid.ColumnDefinitions>
          <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
            <TextBlock Style="{StaticResource IndexColumnHeader}" Text="{index.list.col.name}" />
            <ContentControl Width="8" Height="8" Margin="4,0,0,0" /> <!-- chevron-down -->
          </StackPanel>
          <TextBlock Grid.Column="4" Style="{StaticResource IndexColumnHeader}" VerticalAlignment="Center" Text="{index.list.col.path}" />
          <TextBlock Grid.Column="6" Style="{StaticResource IndexColumnHeader}" VerticalAlignment="Center" HorizontalAlignment="Right" Text="{index.list.col.files}" />
          <TextBlock Grid.Column="8" Style="{StaticResource IndexColumnHeader}" VerticalAlignment="Center" Text="{index.list.col.updated}" />
          <StackPanel Grid.Column="10" Orientation="Horizontal" VerticalAlignment="Center">
            <TextBlock Style="{StaticResource IndexColumnHeader}" Text="{index.list.col.status}" />
            <ContentControl x:Name="IndexStatusHelpIcon" Width="12" Height="12" Margin="4,0,0,0" ToolTip="{index.list.col.status.help}" />
          </StackPanel>
          <TextBlock Grid.Column="12" Style="{StaticResource IndexColumnHeader}" VerticalAlignment="Center" Text="{index.list.col.fast}" />
          <TextBlock Grid.Column="14" Style="{StaticResource IndexColumnHeader}" VerticalAlignment="Center" HorizontalAlignment="Right" Text="{index.list.col.actions}" />
        </Grid>
      </Border>

      <!-- 行。ItemsControl＋選択は ListBox で持つ（今の DataGrid の列の見出しは使わない） -->
      <ListBox x:Name="IndexGrid" Grid.Row="1" SelectionMode="Single" BorderThickness="0" Padding="0"
               ScrollViewer.VerticalScrollBarVisibility="Auto" ScrollViewer.HorizontalScrollBarVisibility="Disabled"
               AlternationCount="2" ItemContainerStyle="{StaticResource IndexRow}">
        <ListBox.ItemTemplate>
          <DataTemplate>
            <Grid Height="36">
              <Grid.ColumnDefinitions>
                <!-- 見出しと同じ 15 列 -->
                <ColumnDefinition Width="13" /><ColumnDefinition Width="12" />
                <ColumnDefinition Width="120" /><ColumnDefinition Width="12" />
                <ColumnDefinition Width="*" MinWidth="160" /><ColumnDefinition Width="12" />
                <ColumnDefinition Width="60" /><ColumnDefinition Width="12" />
                <ColumnDefinition Width="120" /><ColumnDefinition Width="12" />
                <ColumnDefinition Width="164" /><ColumnDefinition Width="12" />
                <ColumnDefinition Width="96" /><ColumnDefinition Width="12" />
                <ColumnDefinition Width="150" />
              </Grid.ColumnDefinitions>
              <CheckBox Grid.Column="0" Style="{StaticResource Choice}" Width="13" Height="13" VerticalAlignment="Center"
                        IsChecked="{Binding Enabled, Mode=TwoWay}" />
              <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                <ContentControl Width="14" Height="14" /> <!-- folder（File.Folder） -->
                <TextBlock Margin="4,0,0,0" Style="{StaticResource IndexRowName}" Foreground="{DynamicResource Ink.Strong}"
                           TextTrimming="CharacterEllipsis" Text="{Binding Name}" ToolTip="{Binding Name}" />
              </StackPanel>
              <TextBlock Grid.Column="4" Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Body}" VerticalAlignment="Center"
                         TextTrimming="CharacterEllipsis" Text="{Binding Path}" ToolTip="{Binding Path}" />
              <TextBlock Grid.Column="6" Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Strong}" VerticalAlignment="Center"
                         HorizontalAlignment="Right" Text="{Binding FilesText}" />
              <TextBlock Grid.Column="8" Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Body}" VerticalAlignment="Center"
                         Text="{Binding UpdatedText}" />
              <!-- ステータス: バッジ＋（更新中は進み具合のバー / 要更新で残りがあれば「残り N 件」）。縦に積み、間 4 -->
              <StackPanel Grid.Column="10" Orientation="Vertical" VerticalAlignment="Center">
                <Border Style="{StaticResource Badge}" Tag="{Binding StatusLevel}" HorizontalAlignment="Left" ToolTip="{Binding StatusToolTip}">
                  <TextBlock Style="{StaticResource Chip}" Text="{Binding StatusText}" />
                </Border>
                <Grid Width="120" Height="4" Margin="0,4,0,0" HorizontalAlignment="Left" Visibility="{Binding BarVisibility}">
                  <Border Background="{DynamicResource Border.Soft}" CornerRadius="2" />
                  <Border Background="{DynamicResource File.Folder}" CornerRadius="2" HorizontalAlignment="Left" Width="{Binding BarWidth}" />
                </Grid>
                <TextBlock Margin="0,4,0,0" Style="{StaticResource Meta}" Foreground="{DynamicResource Warn}"
                           Text="{Binding StatusSubText}" Visibility="{Binding SubTextVisibility}" />
              </StackPanel>
              <Border Grid.Column="12" Style="{StaticResource Badge}" Tag="{Binding FastLevel}" HorizontalAlignment="Left"
                      VerticalAlignment="Center" ToolTip="{Binding FastToolTip}">
                <TextBlock Style="{StaticResource Chip}" Text="{Binding FastText}" />
              </Border>
              <StackPanel Grid.Column="14" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
                <Button x:Name="IndexRowActionButton" Style="{StaticResource MiniButton}" Tag="{Binding ActionKind}"
                        IsEnabled="{Binding ActionEnabled}" ToolTip="{Binding ActionToolTip}">
                  <StackPanel Orientation="Horizontal">
                    <ContentControl Width="10" Height="10" /> <!-- refresh-cw / stop-circle -->
                    <TextBlock Margin="4,0,0,0" Text="{Binding ActionText}" />
                  </StackPanel>
                </Button>
                <Button x:Name="IndexRowMoreButton" Style="{StaticResource MiniButton}" Width="26" Padding="8,0" Margin="6,0,0,0">
                  <ContentControl Width="10" Height="14" /> <!-- ellipsis -->
                </Button>
              </StackPanel>
            </Grid>
          </DataTemplate>
        </ListBox.ItemTemplate>
      </ListBox>

      <!-- 空のとき（X0）。表の見出しも隠し、中央に Empty State -->
      <Grid x:Name="IndexEmptyState" Grid.Row="0" Grid.RowSpan="2" Background="{DynamicResource Bg.Surface}" Visibility="Collapsed">
        <StackPanel HorizontalAlignment="Center" VerticalAlignment="Center">
          <ContentControl x:Name="IndexEmptyIllust" /> <!-- parts/illust_first_run -->
          <TextBlock HorizontalAlignment="Center" Text="{index.list.empty.title}" />
          <TextBlock HorizontalAlignment="Center" Text="{index.list.empty.body}" />
          <Button x:Name="EmptyNewIndexButton" Style="{StaticResource Primary}" HorizontalAlignment="Center" Content="{index.list.new}" />
        </StackPanel>
      </Grid>
    </Grid>
  </ScrollViewer>
</Grid>
```

- 行の中の x:Name（`IndexRowActionButton`・`IndexRowMoreButton`）はテンプレートの中なので `FindName` では引けない。画面層は `ListBox` に `Button.Click` を `AddHandler` で付け、`$e.OriginalSource` から `Tag`・`DataContext` を引く（10 章）。
- 縦のスクロールは `ListBox` の中で行う。表の見出しの行（28）は縦に流れず固定。横は外の `ScrollViewer` で見出しと行が一緒に流れる。
- 新しく足す Style のキー:
  - `IndexColumnHeader`・`IndexRowName`・`Badge`・`MiniButton`・`IndexRow`。
  - どれも theme に無い。案は 4 章に書く。

## 4. 寸法と色

| 要素 | 大きさ | 余白 | 地 | 枠 | 文字の Style | 文字の色 | 角丸 | アイコン |
|---|---|---|---|---|---|---|---|---|
| 見出し（header-section） | 幅 Fill・高さ 73（説明が 1 行のとき） | Padding 16 | Bg.Surface（#FFFFFF） | 下 1 Border.Soft（#E0E2E5） | | | | |
| 題 | 中身 | 題とアイコンの間 6 | | | PageTitle（16 Bold） | Ink.Strong（#202124） | | help-icon 14（線 1.17、色は未定） |
| 説明 | 幅 Fill、折り返す | 題との間 4 | | | Cell（12 Regular） | Ink.Body（#5F6368） | | |
| ボタンの並び | 中身 | ボタンの間 8（「決めること」1） | | | | | | |
| `ImportIndexButton` | 高さ 32・幅 中身（Figma 96） | 左右 14 | Bg.Surface（#FFFFFF） | 1 Border.Soft（#E0E2E5） | Label（12 SemiBold） | Ink.Strong（#202124） | 6 | なし |
| `NewIndexButton` | 中身（Figma 140×31） | 上下 7・左右 16 | Accent（#0078D4） | なし | Body（13）＋SemiBold（Style は未定。「決めること」2） | Ink.OnAccent（#FFFFFF） | 6 | なし（「＋」は文字） |
| 列の見出し（grid-headers） | 高さ 28・最小幅 991 | 左右 12・列の間 12 | Bg.Subtle（#F9FAFA） | 下 1 Border.Soft（#E0E2E5） | `IndexColumnHeader`（11 Bold。「決めること」3） | Ink.Body（#5F6368） | | 名前の列 chevron-down 8（線 0.67）・ステータスの列 help-icon 12（線 1.0） |
| 行 | 高さ 36・最小幅 991 | 左右 12・列の間 12 | 偶数行 Bg.Surface（#FFFFFF）・奇数行 Bg.Stripe（#FAFBFC）・選んだ行 Select.Soft（#E1F2FF） | 下 1 Border.Soft（#E0E2E5） | | | | |
| チェック | 13×13 | | 付いている: Accent（#0078D4）・外れている: Bg.Surface | 外れている: 1 Ink.Body（#5F6368） | | 印 Ink.OnAccent | 2 | |
| インデックス名 | 120 | アイコンとの間 4 | | | `IndexRowName`（12 SemiBold。「決めること」4） | Ink.Strong（#202124） | | folder 14（File.Folder #E8A020 で塗り。線の色は未定） |
| パス | Fill（最小 160） | | | | Meta（11 Regular） | Ink.Body（#5F6368） | | |
| ファイル数 | 60・右寄せ | | | | Meta（11） | Ink.Strong（#202124） | | |
| 最終更新 | 120 | | | | Meta（11） | Ink.Body（#5F6368） | | |
| ステータスのバッジ | 中身・高さ 18 | 上下 2・左右 8 | 区分ごと（下の表） | なし | Chip（11 Medium） | 区分ごと | 3 | |
| 進み具合のバー | 120×4 | バッジとの間 4 | 下地 Border.Soft（#E0E2E5）・中 File.Folder（#E8A020） | | | | 2 | |
| 「残り N 件」 | 中身 | バッジとの間 4 | | | Meta（11） | Warn（#BA7D00） | | |
| 高速検索のバッジ | 中身・高さ 18 | 上下 2・左右 8 | 区分ごと | なし | Chip（11 Medium） | 区分ごと | 3 | |
| 操作の枠 | 150・右寄せ | ボタンの間 6 | | | | | | |
| `MiniButton`（更新・中止・再設定・元のフォルダを設定） | 高さ 24・幅 中身 | 左右 10・アイコンと文字の間 4 | Bg.Surface（#FFFFFF） | 1 Border.Soft（#E0E2E5） | Chip（11 Medium） | Ink.Strong（#202124） | 4 | refresh-cw 10（線 0.83）／stop-circle 10（線 0.83） |
| ［⋯］（mini-button-more） | 26×24 | 左右 8 | Bg.Surface（#FFFFFF） | 1 Border.Soft（#E0E2E5） | | | 4 | ellipsis 10×14（線 0.83） |
| 行のメニュー（row-menu） | 幅 200（理由の行が付くとき 300）・高さ 中身 | 上下 4 | Bg.Surface（#FFFFFF） | 1 Border.Input（#D1D1D1） | | | 6 | 影 1 つ（値は未定） |
| メニューの項目 | 198×31 | 上下 7・左右 12 | ホバー #F0F4F9（キー無し） | | Body（13 Regular） | 「エクスポート…」Ink.Strong（#202124）・「削除」#B3261E（キー無し）・無効 #A0A5AB（キー無し） | | |
| メニューの区切り | 198×9（線 1） | 上下 4 | 線 Border.Row（#EDF0F2） | | | | | |
| 無効の理由（メニューの中の 2 行目） | 幅 274 で折り返す | 項目名との間 2 | | | Meta（11） | Ink.Body（#5F6368） | | |
| 無効の理由のツールチップ | 中身 | 上下 7・左右 10 | #333840（キー無し） | | Meta（11） | Ink.OnAccent（#FFFFFF） | 5 | 下向きの矢印 10×6 |
| 空のとき（Empty State） | 主な領域いっぱい（Figma 1060×663） | | Bg.Surface | | 未定 | 未定 | | parts/illust_first_run |

バッジの区分（`Badge` の `Tag`）。色は Figma の値をそのまま書く。

| Level | 地 | 文字 | 使う文言 |
|---|---|---|---|
| `Ok` | Ok.Soft（#E0F7E0） | Ok（#218A21） | 最新・検索できます・可 |
| `Warn` | Warn.Soft（#FFF5E0） | Warn（#BA7D00） | 要更新・元のフォルダ未設定・反映中 N%・反映待ち |
| `Busy` | Select.Soft（#E1F2FF） | Accent（#0078D4） | 更新中 N%・エクスポート中 N% |
| `Ng` | Danger.Soft（#FFE6E6） | Danger.Text（#D13438） | エラー・不可 |
| `None` | Bg.Tag（#F1F3F4） | Ink.Body（#5F6368） | 未作成・－ |

- キーの無い色（`../figma_wpf_map.md` の「食い違い」に足す案）:
  - #F0F4F9: メニューのホバー。案 `Bg.MenuHover`。
  - #B3261E: 危険な項目の文字。案 `Danger.Strong`。
  - #A0A5AB: 無効の項目の文字。案 `Ink.Disabled`。
  - #333840: ツールチップの地。案 `Bg.Tooltip`。
  - 無効の理由付きのメニューでは、項目名の無効の色が #A0A4A8 だった。#A0A5AB にまとめる案。
- 新しい Style の案:
  - `IndexColumnHeader`: 11 Bold・Ink.Body。
  - `IndexRowName`: 12 SemiBold。Cell.Key は 12 Medium なので別のキーにする。
  - `Badge`: Border。上下 2・左右 8、角丸 3。地と文字の色を `Tag` の値に DataTrigger で結ぶ。
  - `MiniButton`、`IndexRow`（ListBoxItem）: 地を AlternationIndex と IsSelected に結ぶ。

## 5. 状態ごとの見え方

一覧全体の状態。

| 状態（記号） | 入る条件（判断層の入力） | 見出し | 列の見出し・行 | 空のとき | 選んだ行 |
|---|---|---|---|---|---|
| X0・P-X0 | 行が 0 件（`getIndexListView` の `Empty`＝`$true`） | 出す | 隠す | 出す | なし |
| X・既定・P-X-U | 行が 1 件以上 | 出す | 出す | 隠す | なし（P-X-U は 106 の X と同じ中身。画面の違いはこの部品の外） |
| P-X | 行を選んだ | 出す | 出す | 隠す | 営業部を Select.Soft |
| P-X-顧客 | 2 行目を選んだ | 同 | 同 | 隠す | 顧客 |
| P-X-del | 削除したあと | 同 | 削除した行が消える | 隠す | 次の行（顧客）を選ぶ |
| X-⋯ | 行の［⋯］を開いた | 同 | 同 | 隠す | 未定（Figma は X と同じ） |

行の状態（ステータスの列）。上から順に判定し、最初に当てはまったものにする（9 章の `getIndexRowView`）。

| 状態の文言 | Level | 入る条件 | Figma の例 | バー | 「残り N 件」 | 操作のボタン |
|---|---|---|---|---|---|---|
| エクスポート中 {N}% | Busy | `exportPercent` が $null でない | X-エクスポート中・P-X-exporting の営業部 | 未定（Figma はバーなし） | なし | 更新（無効か未定） |
| 更新中 {N}% | Busy | その行を更新中（`indexing`） | X の営業部2025 | 出す（幅 120 × N/100。「決めること」5） | なし | 中止 |
| 元のフォルダ未設定 | Warn | インポートしたインデックスで、元のフォルダが無い（`folderState` が `NotSet`） | X-取り込み後2 | なし | なし | 元のフォルダを設定（「決めること」6） |
| エラー | Ng | 元のフォルダが見つからない（`folderState` が `Missing`。ヘルプの文「エラー: フォルダが見つかりません」） | X のアーカイブ（0・－・チェックなし） | なし | なし | 再設定 |
| 未作成 | None | まだ取り込んでいない（`stat` が $null か `Total` が 0） | X-N | なし | なし | 更新 |
| 検索できます | Ok | インポートしたあと、まだ更新していない（`imported`） | X-取り込み後・P-X-imported | なし | なし | 更新 |
| 要更新 | Warn | 取り込みの残りがある（`Pending` ≥ 1）か、元のフォルダに変更がある（`changed`） | X の顧客・X-P の営業部2025・P-X-refolder のアーカイブ | なし | `Pending` ≥ 1 のとき「残り {N} 件」（X-P） | 更新 |
| 最新 | Ok | 上のどれでもない | X の営業部 | なし | なし | 更新 |

- 取り込みに失敗したファイルがある（今の「一部失敗」）ときの見え方は、Figma に無い（「決めること」8）。
- チェックが外れた行（X のアーカイブ）も、見た目はチェック以外同じ。

行の高速検索の列。今の `getFastSearchRowView` の判定のまま、色の区分だけ写す。

| 文言 | Level | 入る条件 | Figma の例 |
|---|---|---|---|
| 可 | Ok | 反映待ちのフォルダが 0 | X の営業部 |
| 反映中 {N}% | Warn | 反映待ちが 1 以上で全部ではない。N は反映済みの割合の切り捨て、0 なら 1 | X の顧客（62%） |
| 反映待ち | Warn | 反映済みが 0（Waiting＝Folders）か、Windows Search がまだ索引していない（NotYet） | X の営業部2025 |
| 不可 | Ng | Windows Search に接続できない・対象外・インデックスに高速検索用のデータが無い | 高速検索が使えない・P-X-S・X-取り込み後 |
| － | None | 確かめられない・まだ作っていない・更新中 | X のアーカイブ・X-N |
| 確認中… | None | まだ一度も確かめていない（`reason` が $null） | Figma に無い（「決めること」9） |

variant ごとの中身（行は 営業部／顧客／営業部2025／アーカイブ の順）。

| 状態 | 営業部 | 顧客 | 営業部2025 | アーカイブ |
|---|---|---|---|---|
| 既定・X・X-⋯・X-高速検索-反映中・P-X・P-X-U・thumb-3・thumb-7・thumb-10 | 245・2024/10/14 15:30・最新・可・更新 | 1,830・2024/10/14 15:28・要更新・反映中 62%・更新 | 455・2024/10/13 09:00・更新中 45%（バー）・反映待ち・中止 | 0・-・エラー・－・再設定（チェックなし） |
| X-P・thumb-4・P-X-P | X と同じ | X と同じ | 要更新＋「残り 155 件」・反映待ち・更新 | X と同じ |
| X-C・thumb-6・P-X-C | 最新・2024/10/14 15:42 | 最新・15:42・反映中 62% | 最新・15:42・反映待ち・更新 | X と同じ |
| X-N・thumb-2・P-X-N | -・-・未作成・－・更新（1 行だけ） | なし | なし | なし |
| 高速検索が使えない・P-X-S | 不可 | 不可 | 最新・15:42・不可 | － |
| X-エクスポート中・thumb-8・P-X-exporting | エクスポート中 42% | X と同じ | X と同じ | X と同じ |
| X-取り込み後・thumb-9・P-X-imported | パス C:\共有\営業部・2,530・2026/09/27 10:00・検索できます・不可・更新 | X と同じ | X と同じ | X と同じ |
| X-取り込み後2 | 元のフォルダ未設定・不可・元のフォルダを設定 | X と同じ | X と同じ | X と同じ |
| P-X-refolder | 選んでいる | X と同じ | X と同じ | パス C:\Share\アーカイブ2024・要更新・－・更新 |
| P-X-del | 行が無い | 選んでいる | X と同じ | X と同じ |

ボタンの可否。

| 状態 | `NewIndexButton` | `ImportIndexButton` | 行の操作 | ［⋯］の「エクスポート…」 | ［⋯］の「削除」 |
|---|---|---|---|---|---|
| 何も動いていない | 有効 | 有効 | 有効 | 有効（作成を最後まで終えていないものは無効＋理由） | 有効 |
| その行を更新中 | 有効か未定（「決めること」10） | 同 | 中止だけ有効 | 未定 | 無効＋ツールチップ「更新中は削除できません」 |
| エクスポート・インポート中 | 無効＋ツールチップ | 無効＋ツールチップ | 無効＋ツールチップ | 無効 | 無効 |
| 削除中 | 無効 | 無効 | 無効 | 無効 | 無効 |

## 6. 文言

| 文言の ID | 全文 | 使う状態 |
|---|---|---|
| index.list.title | インデックス管理 | すべて |
| index.list.help | 未定（help-icon のツールチップ） | すべて |
| index.list.description | 検索するフォルダとインデックスを管理します。ファイルを変更したら［更新］でインデックスを最新にします。 | すべて |
| index.list.import | インポート… | すべて |
| index.list.new | ＋ フォルダを追加 | すべて・空のとき |
| index.list.col.name | インデックス名 | 行があるとき |
| index.list.col.path | パス | 同 |
| index.list.col.files | ファイル数 | 同 |
| index.list.col.updated | 最終更新 | 同 |
| index.list.col.status | ステータス | 同 |
| index.list.col.status.help | 各フォルダのインデックスの状態です。<br>・最新: インデックスは最新です<br>・要更新: 前回の更新のあとにファイルが変更されています<br>・更新中: インデックスを更新しています<br>・エラー: フォルダが見つかりません（`<br>` は改行。106 の tooltip 108:881） | 「ステータス」の help-icon のツールチップ |
| index.list.col.fast | 高速検索 | 同 |
| index.list.col.actions | 操作 | 同 |
| index.list.row.files | {件数}（3 桁ごとに「,」。例 1,830） | 行 |
| index.list.row.updated | {yyyy/MM/dd HH:mm}。無ければ - （半角のハイフン） | 行 |
| index.list.status.ok | 最新 | 行 |
| index.list.status.stale | 要更新 | 行 |
| index.list.status.rest | 残り {件数} 件 | X-P |
| index.list.status.running | 更新中 {N}% | X |
| index.list.status.error | エラー | X |
| index.list.status.none | 未作成 | X-N |
| index.list.status.exporting | エクスポート中 {N}% | X-エクスポート中 |
| index.list.status.imported | 検索できます | X-取り込み後 |
| index.list.status.nofolder | 元のフォルダ未設定 | X-取り込み後2 |
| index.list.fast.ok | 可 | 行 |
| index.list.fast.partial | 反映中 {N}% | 行 |
| index.list.fast.waiting | 反映待ち | 行 |
| index.list.fast.ng | 不可 | 行 |
| index.list.fast.unknown | －（全角のマイナス） | 行 |
| index.list.action.update | 更新 | 行 |
| index.list.action.stop | 中止 | 行（更新中） |
| index.list.action.refolder | 再設定 | 行（エラー） |
| index.list.action.setfolder | 元のフォルダを設定 | 行（元のフォルダ未設定） |
| index.list.menu.export | エクスポート… | 行のメニュー |
| index.list.menu.export.disabled | インデックスの更新が完了するとエクスポートできます | 行のメニュー（作成を最後まで終えていない） |
| index.list.menu.delete | 削除 | 行のメニュー |
| index.list.menu.delete.disabled | 更新中は削除できません | 「削除」のツールチップ（更新中） |
| index.list.busy | エクスポート・インポートが終わるまでお待ちください。 | ボタンのツールチップ（エクスポート・インポート中） |
| index.list.empty.title | フォルダを追加してください | X0 |
| index.list.empty.body | 検索したいフォルダを追加すると、ここに表示されます。 | X0 |

- 行のツールチップ（ステータス・高速検索）を付ける（決定）。文は今の `getIndexRowView`・`getFastSearchRowView` の文を使う。高速検索の詳しい説明（反映済みのフォルダ数・使えないわけ・最終確認の時刻）は、このツールチップに置く。見本は Figma の「06 ツールチップ」の TT-X の 8・9。

## 7. 操作と振る舞い

| きっかけ | 何が起きるか | 呼ぶ関数 | 次の状態 |
|---|---|---|---|
| `NewIndexButton`（空のときは `EmptyNewIndexButton`）をクリック | フォルダを選ぶ。選んだら行を足して選び、詳細欄に出す | `newIndex`（今の `addIndexForFolder` を流用） | X（足した行を選ぶ） |
| `ImportIndexButton` をクリック | インポートのダイアログ（`dialog/` の側） | `newImportIndex` | X-取り込み後・X-取り込み後2 |
| 行をクリック | 行を選び、詳細欄をその行にする | `selectIndexRow` | P-X・P-X-顧客 |
| チェックをクリック | 「すべて更新」のときに取り込む対象を切り替えて保存する | `saveTargets` | 同じ |
| ［更新］ | その行のインデックスを更新する | `startIndexUpdate`（新規。今の `startIndexing` を 1 行に絞る） | 更新中 N% |
| ［中止］ | 確認のダイアログ「更新の中止」（「「{名前}」の更新を中止しますか？」［中止する］。`dialog/` の側）を出し、［中止する］で止める。止まるまでダイアログは「中止しています…」 | `stopIndexing` | X-P（要更新＋残り N 件） |
| ［再設定］・［元のフォルダを設定］ | 元のフォルダを選び直す | `editIndexFolder`（新規。詳細欄の［...］と同じ処理） | P-X-refolder（要更新） |
| ［⋯］をクリック | 行のメニューを、ボタンの下に右端をそろえて開く（位置は未定） | `showIndexRowMenu` | X-⋯ |
| メニューの「エクスポート…」 | エクスポートの確認（DX）を出す | `newExportIndex` | エクスポート中 N% |
| メニューの「削除」 | 削除の確認（DC-del、赤）を出し、［削除する］で消す。ステータスバーに「削除しました」 | `deleteIndex` | P-X-del |
| 無効の「削除」にホバー | ツールチップ「更新中は削除できません」 | | |
| 無効のボタンにホバー（エクスポート・インポート中） | ツールチップ「エクスポート・インポートが終わるまでお待ちください。」 | | |
| パス・名前にホバー | 全文のツールチップ | | |
| 列の見出し「インデックス名」をクリック | 並べ替え（chevron-down がある）。するかどうかは未定（「決めること」11） | | |
| 行を外からドロップ | 今の `AllowDrop` の動き（フォルダを足す）を残すか未定（「決めること」12） | | |

- ホバー・押した・フォーカスの見た目（ボタン・行）は Figma に無い。今の theme の `Button.Base`・`Primary` の Trigger を使い、行は `Bg.Hover`（#F3F3F4）の案（「決めること」13）。
- 無効の見た目: ボタンは今の theme のまま。メニューの項目は文字を #A0A5AB。
- Tab の順: `ImportIndexButton` → `NewIndexButton` → `IndexGrid`（行の中は チェック → 操作 → ［⋯］）。
- 高速検索の確かめの時機は今のまま: 起動・画面を前に出したとき（60 秒以上たっていれば）・更新のあと・追加／編集／削除／インポートのあと、反映待ち・反映中の行があれば 5 分ごと。

## 8. リサイズ

一覧が多いとき（スクロールする所・固定する所・仮想化・未定のこと）は `../scroll.md`（見本 X-多・X-多-詳細・X-多-1024）。

- 一覧の行は `*`（MinHeight 225）。225 は「見出し 89＋表の見出し 28＋3 行 36×3」（Figma の注記）。
- 窓を高くすると伸びるのは一覧。詳細欄は 397 のまま（`index_detail.md`）。
- 窓を低くすると、先に詳細欄が 397 から 240 まで縮み、一覧は 225 で止まる。
- 縮む順番はコードで決める。インデックス管理の画面の Grid の `SizeChanged` で、次の式で詳細欄の高さを決める。
  - `$avail = $grid.ActualHeight - 5`
  - `$detail = [Math]::Max(240, [Math]::Min(397, $avail - 225))`
  - 詳細欄の `Height` に `$detail` を入れる。
  - 境目をドラッグしたあとは、その高さ（240〜397 に丸めた値）を上限にする。
- 見出し: 題と説明の列は `*`、説明は `TextWrapping="Wrap"`。ボタンは中身の幅で、右に寄せる。
- 表: 最小幅 991。伸びるのは「パス」の列だけ（`*`、MinWidth 160）。ほかは固定の幅。
  - 幅が足りないときは、外の `ScrollViewer` で横にスクロールする（`HorizontalScrollBarVisibility="Auto"`）。
  - 行が多いときは、`ListBox` で縦にスクロールする（`Auto`）。
- パス・名前は 1 行で `TextTrimming="CharacterEllipsis"`。
- 1024×640（408:33785）:
  - 主な領域は 804×608。一覧 225・境目 5・詳細欄 354・ステータスバー 24。
  - 表は 991 > 主な領域の 804 なので、横にスクロールが出る。
  - 説明は折り返して 2 行になる（見出しは 89）。
- 1280×820:
  - 主な領域は 1060×788。詳細欄 397、一覧はその残り（788 − 24 − 5 − 397 ＝ 362）。
- 1600×1000（408:33812）:
  - 主な領域は 1380×968。一覧 542・詳細欄 397。
  - パスの列が伸びる。横のスクロールは出ない。

## 9. 判断層（`ui/index/index_list_view.ps1`）

今の `index_view.ps1` の関数の名前の付け方（`get<何>View`・`test<何>`）に合わせる。今の関数は移すか、そのまま呼ぶ。

### getIndexRowView（作り直す）

- 入力:
  - `stat`（Total; Done; Pending; Failed）
  - `indexing`（その行を更新中か。bool）
  - `percent`（更新の進み具合。int か $null）
  - `exportPercent`（int か $null）
  - `folderState`（`Ok`・`Missing`・`NotSet`）
  - `imported`（インポートのあと更新していないか。bool）
  - `changed`（元のフォルダに変更があるか。bool）
- 出力: `@{ Text = [string]; Level = [string]; ToolTip = [string]; SubText = [string]; BarPercent = [int] or $null }`。

| stat | indexing | percent | exportPercent | folderState | imported | changed | → Text | Level | SubText | BarPercent |
|---|---|---|---|---|---|---|---|---|---|---|
| @{Total=245;Pending=0;Failed=0} | $false | $null | 42 | Ok | $false | $false | エクスポート中 42% | Busy | "" | $null |
| @{Total=455;Pending=155;Failed=0} | $true | 45 | $null | Ok | $false | $false | 更新中 45% | Busy | "" | 45 |
| @{Total=2530;Pending=0;Failed=0} | $false | $null | $null | NotSet | $true | $false | 元のフォルダ未設定 | Warn | "" | $null |
| $null | $false | $null | $null | Missing | $false | $false | エラー | Ng | "" | $null |
| $null | $false | $null | $null | Ok | $false | $false | 未作成 | None | "" | $null |
| @{Total=0;Pending=0;Failed=0} | $false | $null | $null | Ok | $false | $false | 未作成 | None | "" | $null |
| @{Total=2530;Pending=0;Failed=0} | $false | $null | $null | Ok | $true | $false | 検索できます | Ok | "" | $null |
| @{Total=455;Pending=155;Failed=0} | $false | $null | $null | Ok | $false | $false | 要更新 | Warn | 残り 155 件 | $null |
| @{Total=1830;Pending=0;Failed=0} | $false | $null | $null | Ok | $false | $true | 要更新 | Warn | "" | $null |
| @{Total=245;Pending=0;Failed=0} | $false | $null | $null | Ok | $false | $false | 最新 | Ok | "" | $null |

- `indexing` が真で `percent` が $null のときは、「更新中 0%」にするか未定（「決めること」5）。

### getProgressBarWidth

- 入力: `percent`（int）、`width`（double）。出力: double（0〜width）。

| percent | width | → |
|---|---|---|
| 45 | 120 | 54 |
| 0 | 120 | 0 |
| 100 | 120 | 120 |
| 150 | 120 | 120 |
| -5 | 120 | 0 |

### getFastSearchRowView（今のまま。Level の区分だけ変える）

- 出力の `Level` を、今の `Ok`・`Wait`・`Ng`・`None` から `Badge` の区分 `Ok`・`Warn`・`Ng`・`None` に替える（`Wait` → `Warn`）。
- 今のテストに次を足す。

| reason | progress.ByIndex[name] | hasContent | indexing | → Text | Level |
|---|---|---|---|---|---|
| Ok | @{Folders=150;Waiting=0} | $true | $false | 可 | Ok |
| Ok | @{Folders=150;Waiting=57} | $true | $false | 反映中 62% | Warn |
| Ok | @{Folders=150;Waiting=150} | $true | $false | 反映待ち | Warn |
| Ok | @{Folders=200;Waiting=199} | $true | $false | 反映中 1% | Warn |
| NotYet | @{Folders=150;Waiting=10} | $true | $false | 反映待ち | Warn |
| NoConnection | （無し） | $true | $false | 不可 | Ng |
| NotInScope | （無し） | $true | $false | 不可 | Ng |
| Ok | （無し） | $true | $false | 不可 | Ng |
| Ok | （無し） | $true | $true | － | None |
| NoFolder | （無し） | $false | $false | － | None |
| Ok | progress が $null | $true | $false | － | None |
| $null | | | | 確認中… | None |

### getIndexRowAction

- 入力: `rowView`（`getIndexRowView` の出力の Text の種類。`indexing`・`folderState` から決まる）、`blocker`（`getIndexJobBlocker` の値）。
- 出力: `@{ Kind = [string]; Text = [string]; Enabled = [bool]; ToolTip = [string] }`。Kind は `Update`・`Stop`・`Refolder`・`SetFolder`。

| indexing | folderState | blocker | → Kind | Text | Enabled | ToolTip |
|---|---|---|---|---|---|---|
| $false | Ok | "" | Update | 更新 | $true | "" |
| $true | Ok | インデックス作成中 | Stop | 中止 | $true | "" |
| $false | Missing | "" | Refolder | 再設定 | $true | "" |
| $false | NotSet | "" | SetFolder | 元のフォルダを設定 | $true | "" |
| $false | Ok | エクスポート・インポート中 | Update | 更新 | $false | エクスポート・インポートが終わるまでお待ちください。 |
| $false | Ok | インデックス作成中 | Update | 更新 | 未定（「決めること」10） | 未定 |

### getIndexRowMenu

- 入力: `stat`、`indexing`、`blocker`。
- 出力: `@{ ExportEnabled = [bool]; ExportReason = [string]; DeleteEnabled = [bool]; DeleteReason = [string] }`。

| stat | indexing | blocker | → ExportEnabled | ExportReason | DeleteEnabled | DeleteReason |
|---|---|---|---|---|---|---|
| @{Total=245;Pending=0} | $false | "" | $true | "" | $true | "" |
| @{Total=455;Pending=155} | $false | "" | $false | インデックスの更新が完了するとエクスポートできます | $true | "" |
| @{Total=455;Pending=155} | $true | インデックス作成中 | $false | インデックスの更新が完了するとエクスポートできます | $false | 更新中は削除できません |
| @{Total=245;Pending=0} | $false | エクスポート・インポート中 | $false | エクスポート・インポートが終わるまでお待ちください。 | $false | エクスポート・インポートが終わるまでお待ちください。 |

### getIndexListButtonsEnabled（今の `getIndexTabButtonsEnabled` を置き換える）

- 入力: `blocker`。
- 出力: `@{ New = [bool]; Import = [bool]; ToolTip = [string] }`。編集・削除・エクスポートは行に移ったので持たない。

| blocker | → New | Import | ToolTip |
|---|---|---|---|
| "" | $true | $true | "" |
| エクスポート・インポート中 | $false | $false | エクスポート・インポートが終わるまでお待ちください。 |
| 削除中 | $false | $false | 未定 |
| インデックス作成中 | 未定（「決めること」10） | 未定 | 未定 |

### getIndexListView

- 入力: `count`（行の数）。出力: `@{ Empty = [bool] }`。

| count | → Empty |
|---|---|
| 0 | $true |
| 1 | $false |

### formatIndexRowUpdated・formatFileCount

| 関数 | 入力 | → |
|---|---|---|
| formatIndexRowUpdated | [datetime]'2024-10-14 15:30' | 2024/10/14 15:30 |
| formatIndexRowUpdated | $null | - |
| formatFileCount | 1830 | 1,830 |
| formatFileCount | 0 | 0 |

## 10. 画面層（`ui/index/index_list.ps1`）

- 読み込みのとき: `FindName` で `ImportIndexButton`・`NewIndexButton`・`IndexGrid`・`IndexGridScroll`・`IndexEmptyState`・`EmptyNewIndexButton` を引く。
- 行のデータ（PSCustomObject）の項目:
  - Enabled・Name・Path・FilesText・UpdatedText
  - StatusText・StatusLevel・StatusToolTip・StatusSubText・BarWidth・BarVisibility・SubTextVisibility
  - FastText・FastLevel・FastToolTip
  - ActionKind・ActionText・ActionEnabled・ActionToolTip
  - `INotifyPropertyChanged` が要るので、今の `newFolderItem` の作りを使う。
- 判断層の出力を写す手順（`updateIndexListView`）:
  1. `getIndexListView` で `IndexEmptyState` と表の `Visibility` を切り替える。
  2. 行ごとに `getIndexRowView`・`getFastSearchRowView`・`getIndexRowAction` を呼び、項目に入れる。BarWidth は `getProgressBarWidth`。
  3. `getIndexListButtonsEnabled` で見出しのボタンの `IsEnabled`・`ToolTip` を決める。
  4. 選んでいる行が消えたら、次の行（無ければ前の行）を選ぶ（P-X-del）。
- 行の中のボタン:
  - `$IndexGrid.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, ...)` で受ける。
  - `OriginalSource` から上に辿って `Button` を見つけ、`Tag`（ActionKind）と `DataContext`（行）で処理を分ける。
  - ［⋯］は `ContextMenu` を `PlacementTarget` にボタンを入れて開く。
- 重い処理（更新・エクスポート・削除・高速検索の確かめ）は今のまま別のスレッドのジョブで行い、進み具合は `DispatcherTimer` で拾って行の項目を書き換える（今の `updateIndexingProgress`・`updateFastSearchRows` と同じ）。UI の要素には Dispatcher のスレッドからだけ触る。
- 今の `IndexingProgressPanel`・`FailedPanel`・`IndexSummaryText`・`IndexingStateText` に当たる表示は、この領域に無い。進み具合は行のバッジとバーに、合計はステータスバー（`shell/status_bar.md`）に移す。失敗したファイルの一覧の置き場は「決めること」14。

## 11. 受け入れ

- 見比べる画像:
  - 1280×820 で、Figma の get_screenshot の 352:24953（X）と比べる。主な領域は 1060 幅。
  - 1024×640 で 408:33785 と、1600×1000 で 408:33812 と比べる。
- 確かめる点:
  - 見出しの高さが 73 である（1060 幅で説明が 1 行のとき）。下に 1px の線がある。
  - ボタンが［インポート…］［＋ フォルダを追加］の順に右に並ぶ。主ボタンは Accent（#0078D4）で角丸 6 である。
  - 列の見出しが高さ 28 で、地が #F9FAFA である。列の幅は 13／120／Fill／60／120／164／96／150、間は 12 である。
  - 行の高さが 36 で、1 行おきに #FAFBFC である。選んだ行は #E1F2FF である。
  - バッジの地と文字の色が 4 章の表のとおりである。高さは 18、角丸は 3 である。
  - 更新中の行のバッジの下に、120×4 のバーがある（下地 #E0E2E5・中 #E8A020）。
  - 操作の列のボタンの高さが 24、［⋯］が 26×24 で、右端にそろっている。
  - 1024×640 で、一覧が 225・詳細欄が 354 である。表が横にスクロールする。
  - 1600×1000 で、詳細欄が 397 のままで一覧が 542 に伸びる。
  - 長いパスが 1 行で「…」になる。
  - 行が 0 件のとき（X0）、表の見出しが消え、空のときの絵と 2 行と［＋ フォルダを追加］が中央に出る。
  - 行のメニューは幅 200、項目の高さは 31 である。「削除」は #B3261E である。

## 決めること

1. **見出しのボタンの数と間。**
   - Figma の部品には［すべて更新］（`IndexingButton`、枠だけ、refresh-cw 12）がある。メンテナの決定は「上は［＋ フォルダを追加］と［インポート…］の 2 つ」。
   - この設計書は 2 つにした。Figma から［すべて更新］を消すか、残して決定を直すかを決める。消すなら、チェックの列（取り込む対象）の要否も決める。
   - ボタンの間も食い違う。Figma の注記は 16 px、描かれた値は 8 px。
2. **［＋ フォルダを追加］の文字の Style。** 13 SemiBold にあたるキーが theme に無い（Heading は 13 SemiBold だが用途が違う）。`Primary` の Style に持たせるか決める。
3. **列の見出しの文字。** Figma は 11 Bold、theme の `ColumnHeader` は 10 Bold。新しいキー（`IndexColumnHeader`）にするか、どちらかにそろえるかを決める。
4. **インデックス名の文字。** Figma は 12 SemiBold、`Cell.Key` は 12 Medium。
5. **進み具合のバー。**
   - Figma では 45% のときに 78 px（65%）が塗られていて、割合と合わない。この設計書は「幅 × 割合」にした。
   - 更新の始めで割合が分からないときの表示（「更新中 0%」か「更新中」か）も決める。
   - エクスポート中にバーを出すかも決める。
6. **［元のフォルダを設定］。** メンテナの決定の行のボタンは 更新・中止・再設定 だけ。X-取り込み後2 には［元のフォルダを設定］がある。［再設定］にそろえるか、4 つ目として認めるか。
7. **「エラー」に入る条件。** ヘルプの文は「フォルダが見つかりません」。ネットワークのフォルダに一時的に届かないときも「エラー」にするか（今の「フォルダの状態」の列は、届かないときと無いときを分けている）。
8. **取り込みに失敗したファイルがある行の見え方。** 今の「一部失敗」に当たる状態が Figma に無い。
9. **「確認中…」。** 高速検索をまだ一度も確かめていないときの表示が Figma に無い。今は「確認中…」を出す。「－」にそろえるか。
10. **更新中にほかの操作をどこまで許すか。**
    - 更新中の［＋ フォルダを追加］［インポート…］とほかの行の［更新］が、できるかどうか。
    - 今は作成中は追加・編集・削除を止める（E32「作成中に追加・名前の変更・削除」）。
    - 行ごとの更新で、ほかの行を並べて（待ち行列で）更新できるようにするかも決める。
11. **列の並べ替え。** 「インデックス名」に chevron-down がある。並べ替えるか、ただの印か。
12. **フォルダのドロップ。** 今の一覧はフォルダをドロップして足せる。新しい一覧でも残すか。
13. **ホバー・押した・フォーカスの見た目。** 行・小さなボタン・メニューの項目のホバー（メニューは #F0F4F9 が 1 か所だけ描かれている）・押した・フォーカスの見た目が Figma に無い。
14. **取り込み中のパネルと失敗の一覧の置き場。** 今の `IndexingProgressPanel`（ログを開く などを含む）と `FailedPanel`（失敗したファイルの一覧・フォルダを開く）が、部品に無い。消すのか、詳細欄か別の画面に移すのか。
15. **題の横の help-icon の中身。** 題の横（14）のツールチップの文言が読み取れなかった。アイコンの色は詳細欄の help-icon と同じ Ink.Faint（#99A1AB）にする案。
16. **注記の「場所」と見出しの「パス」。** 注記は「場所」の列、見出しの文言は「パス」。この設計書は見出しの「パス」にした。
17. **行のメニューの位置と影。** 開く位置（ボタンの下で右そろえか）と影の値が読み取れなかった。
18. **境目の線。** `IndexSplitter` の Border.Soft の線が上下どちらに付くかが読み取れなかった。
19. **空のとき（Empty State）の寸法と文字の Style。** 部品のインスタンスの中は読んでいない。`parts/illust_first_run` の寸法と、2 行の文字の Style を決める。
20. **キーの無い色 4 つ**（#F0F4F9・#B3261E・#A0A5AB・#333840）を theme のキーにするか。案は 4 章。
