# インデックスの詳細欄（index_detail）

## 1. 役割

一覧（`index_list.md`）で選んだインデックスの詳細を、一覧の下に出す。中身は 3 つ。

- 左の列「基本設定」: 名前とフォルダパスの設定はここだけで行う。下にインデックスの状態の枠を置く。
- 右の列「インデックス情報」: 件数・サイズ・最終更新。
- 右の列「高速検索」: 状態・反映の進み具合・不可の理由と直し方・最終確認。

- Figma の部品: `xaml/index/index_detail` 352:39102（既定 340:1203）。
- variant（106）:
  - 既定・X 352:37690・X-C 352:37796・X-N 352:37892・X0 352:37974。
  - X-取り込み後 352:38785・X-取り込み後2 352:38889・X-高速検索-反映中 352:38995・インデックス管理（高速検索が使えない）352:38081。
  - thumb 352:38163・thumb-2〜thumb-6。
- variant（148）: P-X-edit 352:39101。
- 使う画面: インデックス管理の画面すべて（Figma のページ「01 画面」「03 正常系」、「07 プロトタイプ」の P-X 系）。
- リサイズの見本: X 1024×640（408:33785）・X 1600×1000（408:33812）。
- `../png/` にこの領域の画像は無い。見比べは Figma の get_screenshot で行う（11 章）。

## 2. 置き場所

- XAML: `scripts/tebunko/xaml/index/index_detail.xaml`。
- 画面層 `ui/index/index_detail.ps1`、判断層 `ui/index/index_detail_view.ps1`。
- 親: インデックス管理の画面の Grid の 2 行目（`IndexDetailHost`。境目 `IndexSplitter` の下）。組み立ては `index_list.md` の 2 章。
- 読み込み: 起動のときに 1 回だけ `XamlReader::Load` で読んで `IndexDetailHost.Content` に入れ、`FindName` で x:Name を引く。選ぶ行が変わっても作り直さず、値だけ入れ替える。

## 3. 部品の木

```xml
<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
      x:Name="IndexDetailRoot" Background="{DynamicResource Bg.Surface}" MinHeight="240" MaxHeight="397">
  <Grid.RowDefinitions>
    <RowDefinition Height="40" />  <!-- preview-toolbar -->
    <RowDefinition Height="*" />   <!-- stats-columns（縦にスクロール） -->
  </Grid.RowDefinitions>

  <!-- 上の帯 -->
  <Border Grid.Row="0" Background="{DynamicResource Bg.Subtle}" BorderBrush="{DynamicResource Border.Soft}"
          BorderThickness="0,0,0,1" Padding="16,8">
    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
      <ContentControl Width="14" Height="14" /> <!-- info -->
      <TextBlock x:Name="IndexDetailTitle" Margin="8,0,0,0" Style="{StaticResource Label.Strong}"
                 Foreground="{DynamicResource Ink.Strong}" TextTrimming="CharacterEllipsis" Text="{index.detail.title}" />
    </StackPanel>
  </Border>

  <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
    <Grid Margin="16,12">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*" />   <!-- basic-settings -->
        <ColumnDefinition Width="Auto" /> <!-- divider（左右 32） -->
        <ColumnDefinition Width="*" />   <!-- index-info -->
      </Grid.ColumnDefinitions>

      <!-- 基本設定 -->
      <StackPanel Grid.Column="0" Orientation="Vertical">
        <StackPanel Orientation="Horizontal">
          <TextBlock Style="{StaticResource Label}" Foreground="{DynamicResource Ink.Value}" Text="{index.detail.basic}" />
          <ContentControl Width="12" Height="12" Margin="5,0,0,0" ToolTip="{index.detail.basic.help}" /> <!-- help-icon -->
        </StackPanel>
        <Grid Margin="0,10,0,0">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="90" />
            <ColumnDefinition Width="8" />
            <ColumnDefinition Width="*" />
          </Grid.ColumnDefinitions>
          <TextBlock Grid.Column="0" Style="{StaticResource Cell}" Foreground="{DynamicResource Ink.Muted}" VerticalAlignment="Center"
                     Text="{index.detail.name}" />
          <TextBox x:Name="IndexNameBox" Grid.Column="2" Style="{StaticResource DetailInput}" Height="28" Padding="8,5" />
        </Grid>
        <Grid Margin="0,10,0,0">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="90" />
            <ColumnDefinition Width="8" />
            <ColumnDefinition Width="*" />
            <ColumnDefinition Width="34" />
          </Grid.ColumnDefinitions>
          <TextBlock Grid.Column="0" Style="{StaticResource Cell}" Foreground="{DynamicResource Ink.Muted}" VerticalAlignment="Center"
                     Text="{index.detail.folder}" />
          <TextBox x:Name="IndexFolderBox" Grid.Column="2" Style="{StaticResource DetailInput}" Height="28" Padding="8,5"
                   IsReadOnly="True" ToolTip="{Binding Text, RelativeSource={RelativeSource Self}}" />
          <Button x:Name="BrowseIndexFolderButton" Grid.Column="3" Style="{StaticResource DetailBrowseButton}" Height="28"
                  Padding="12,5" Content="{index.detail.browse}" />
        </Grid>
        <!-- 状態の枠（status-box-body） -->
        <Border x:Name="IndexStatusBox" Margin="0,10,0,0" Padding="10,8" Background="{DynamicResource Bg.Stripe}"
                BorderBrush="{DynamicResource Border.Soft}" BorderThickness="1" CornerRadius="4">
          <StackPanel>
            <StackPanel Orientation="Horizontal">
              <TextBlock Style="{StaticResource Label}" Foreground="{DynamicResource Ink.Strong}" VerticalAlignment="Center"
                         Text="{index.detail.status.head}" />
              <Border x:Name="IndexStatusBadge" Margin="8,0,0,0" Style="{StaticResource Badge}">
                <TextBlock x:Name="IndexStatusBadgeText" Style="{StaticResource Chip}" />
              </Border>
            </StackPanel>
            <TextBlock x:Name="IndexStatusLine1" Margin="0,4,0,0" Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Body}"
                       TextTrimming="CharacterEllipsis" />
            <TextBlock x:Name="IndexStatusLine2" Margin="0,4,0,0" Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Body}"
                       TextTrimming="CharacterEllipsis" />
          </StackPanel>
        </Border>
      </StackPanel>

      <Border Grid.Column="1" Width="1" Margin="32,0" Background="{DynamicResource Border.Divider}" />

      <!-- インデックス情報と高速検索 -->
      <StackPanel Grid.Column="2" Orientation="Vertical">
        <StackPanel Orientation="Horizontal">
          <TextBlock Style="{StaticResource Label}" Foreground="{DynamicResource Ink.Value}" Text="{index.detail.info}" />
          <ContentControl Width="12" Height="12" Margin="5,0,0,0" ToolTip="{index.detail.info.help}" /> <!-- help-icon -->
        </StackPanel>
        <Border Margin="0,8,0,0" Background="{DynamicResource Bg.Surface}" BorderBrush="{DynamicResource Border.Normal}"
                BorderThickness="1" CornerRadius="4">
          <!-- 見出し 1 行＋値 6 行。行の間は 1 px の線（Border.Row） -->
          <ItemsControl x:Name="IndexInfoRows">
            <ItemsControl.ItemTemplate>
              <DataTemplate>
                <Border BorderBrush="{DynamicResource Border.Row}" BorderThickness="0,1,0,0">
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*" />
                      <ColumnDefinition Width="130" />
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Margin="10,5" Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Muted}"
                               TextTrimming="CharacterEllipsis" Text="{Binding Label}" />
                    <TextBlock Grid.Column="1" Margin="10,5" Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Value}"
                               HorizontalAlignment="Right" TextTrimming="CharacterEllipsis" Text="{Binding Value}" />
                  </Grid>
                </Border>
              </DataTemplate>
            </ItemsControl.ItemTemplate>
            <ItemsControl.Template>
              <ControlTemplate TargetType="ItemsControl">
                <StackPanel>
                  <Grid Background="{DynamicResource Bg.Window}">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*" />
                      <ColumnDefinition Width="130" />
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Margin="10,6" Style="{StaticResource Chip}" Foreground="{DynamicResource Ink.Value}" Text="{index.detail.info.col.item}" />
                    <TextBlock Grid.Column="1" Margin="10,6" Style="{StaticResource Chip}" Foreground="{DynamicResource Ink.Value}"
                               HorizontalAlignment="Right" Text="{index.detail.info.col.value}" />
                  </Grid>
                  <ItemsPresenter />
                </StackPanel>
              </ControlTemplate>
            </ItemsControl.Template>
          </ItemsControl>
        </Border>

        <!-- 高速検索 -->
        <StackPanel x:Name="FastSearchSection" Margin="0,8,0,0">
          <TextBlock Style="{StaticResource Label}" Foreground="{DynamicResource Ink.Value}" Text="{index.detail.fast}" />
          <Border Margin="0,8,0,0" Padding="12,10" Background="{DynamicResource Bg.Surface}" BorderBrush="{DynamicResource Border.Normal}"
                  BorderThickness="1" CornerRadius="4">
            <StackPanel>
              <!-- 各行: 項目 90 ＋間 8 ＋値。行の間 6 -->
              <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="90" /><ColumnDefinition Width="8" /><ColumnDefinition Width="*" /></Grid.ColumnDefinitions>
                <TextBlock Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Muted}" VerticalAlignment="Center" Text="{index.detail.fast.state}" />
                <Border x:Name="FastSearchBadge" Grid.Column="2" Style="{StaticResource Badge}" HorizontalAlignment="Left">
                  <TextBlock x:Name="FastSearchBadgeText" Style="{StaticResource Chip}" />
                </Border>
              </Grid>
              <Grid x:Name="FastSearchProgressLine" Margin="0,6,0,0">
                <Grid.ColumnDefinitions><ColumnDefinition Width="98" /><ColumnDefinition Width="*" /></Grid.ColumnDefinitions>
                <!-- 線と文。幅の決め方は 8 章（SizeChanged で 1 行か 2 行かを決める） -->
                <Grid x:Name="FastSearchProgressBox" Grid.Column="1">
                  <Grid.ColumnDefinitions><ColumnDefinition Width="Auto" /><ColumnDefinition Width="*" /></Grid.ColumnDefinitions>
                  <Grid.RowDefinitions><RowDefinition Height="Auto" /><RowDefinition Height="Auto" /></Grid.RowDefinitions>
                  <Grid x:Name="FastSearchProgressBar" Height="4" Width="100" MinWidth="60" MaxWidth="100" VerticalAlignment="Center">
                    <Border Background="{DynamicResource Border.Soft}" CornerRadius="2" />
                    <Border x:Name="FastSearchProgressFill" HorizontalAlignment="Left" CornerRadius="2" />
                  </Grid>
                  <TextBlock x:Name="FastSearchProgressText" Grid.Column="1" Margin="8,0,0,0" Style="{StaticResource Meta}"
                             Foreground="{DynamicResource Ink.Value}" TextWrapping="NoWrap" />
                </Grid>
              </Grid>
              <Grid x:Name="FastSearchReasonLine" Margin="0,6,0,0">
                <Grid.ColumnDefinitions><ColumnDefinition Width="90" /><ColumnDefinition Width="8" /><ColumnDefinition Width="*" /></Grid.ColumnDefinitions>
                <TextBlock Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Muted}" Text="{index.detail.fast.reason}" />
                <TextBlock x:Name="FastSearchReasonText" Grid.Column="2" Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Value}" TextWrapping="Wrap" />
              </Grid>
              <Grid x:Name="FastSearchFixLine" Margin="0,6,0,0">
                <Grid.ColumnDefinitions><ColumnDefinition Width="90" /><ColumnDefinition Width="8" /><ColumnDefinition Width="*" /></Grid.ColumnDefinitions>
                <TextBlock Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Muted}" Text="{index.detail.fast.fix}" />
                <TextBlock x:Name="FastSearchFixText" Grid.Column="2" Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Value}" TextWrapping="Wrap" />
              </Grid>
              <Grid x:Name="FastSearchCheckedLine" Margin="0,6,0,0">
                <Grid.ColumnDefinitions><ColumnDefinition Width="90" /><ColumnDefinition Width="8" /><ColumnDefinition Width="*" /></Grid.ColumnDefinitions>
                <TextBlock Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Muted}" VerticalAlignment="Center" Text="{index.detail.fast.checked}" />
                <TextBlock x:Name="FastSearchCheckedText" Grid.Column="2" Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Value}"
                           TextTrimming="CharacterEllipsis" />
              </Grid>
              <TextBlock x:Name="FastSearchNote" Margin="0,6,0,0" Style="{StaticResource Micro}" Foreground="{DynamicResource Ink.Muted}"
                         TextWrapping="Wrap" Text="{index.detail.fast.note}" />
            </StackPanel>
          </Border>
        </StackPanel>
      </StackPanel>
    </Grid>
  </ScrollViewer>
</Grid>
```

- Figma の x:Name は `IndexDetailTitle` だけ。次の名前は、この設計書で新しく付けた。
  - `IndexNameBox`・`IndexFolderBox`・`BrowseIndexFolderButton`
  - `IndexStatusBox`・`IndexStatusBadge(Text)`・`IndexStatusLine1/2`
  - `IndexInfoRows`
  - `FastSearch*`
- Figma には hidden の層 `field-ステータス`・`field-最終更新` がある。X0 以外では出さないので、XAML に作らない（「決めること」1）。
- 新しく足す Style のキー:
  - `DetailInput`（TextBox）・`DetailBrowseButton`。
  - `Badge` は `index_list.md` と同じものを使う。

## 4. 寸法と色

| 要素 | 大きさ | 余白 | 地 | 枠 | 文字の Style | 文字の色 | 角丸 | アイコン |
|---|---|---|---|---|---|---|---|---|
| 詳細欄の全体 | 横 Fill・高さ 既定 397（最小 240・最大 397） | | Bg.Surface（#FFFFFF） | | | | | |
| 上の帯（preview-toolbar） | 高さ 40 | 上下 8・左右 16 | Bg.Subtle（#F9FAFA） | 下 1 Border.Soft（#E0E2E5） | | | | info 14（線 1.17、Ink.Strong #202124） |
| `IndexDetailTitle` | 中身 | アイコンとの間 8 | | | Label.Strong（12 Bold） | Ink.Strong（#202124） | | |
| 中（stats-columns） | 横 Fill | 上下 12・左右 16・列の間 32 | | | | | | |
| 左右の列 | 等分（1060 幅で各 482） | 中の行の間: 基本設定 10・インデックス情報 8 | | | | | | |
| 区切り（divider） | 幅 1・高さ Fill | | Border.Divider（#E5E8ED） | | | | | |
| 見出し（基本設定・インデックス情報・高速検索） | 中身 | help-icon との間 5 | | | Label（12 SemiBold） | Ink.Value（#212126） | | help-icon 12（線 1.0、Ink.Faint #99A1AB）。「高速検索」には無い |
| 項目の名前（インデックス名・フォルダパス） | 幅 90 | 入力欄との間 8 | | | Cell（12 Regular） | Ink.Muted（#6B737D） | | |
| `IndexNameBox` | 高さ 28・幅 Fill | 上下 5・左右 8 | Bg.Surface（#FFFFFF） | 1 Border.Normal（#D9DEE3）。フォーカスのとき 2 Accent（#0078D4）で高さ 30（P-X-edit） | Cell（12） | Ink.Value（#212126） | 4 | |
| `IndexFolderBox` | 高さ 28・幅 Fill | 上下 5・左右 8 | Bg.Surface（#FFFFFF） | 1 Border.Normal（#D9DEE3） | Cell（12） | Ink.Value（#212126） | 左だけ 4 | |
| `BrowseIndexFolderButton` | 34×28 | 上下 5・左右 12 | Bg.Button（#F2F2F5） | 1 Border.Normal（#D9DEE3） | 12 Medium（キー無し。「決めること」7） | Button.Text（#4D4D54 をまとめた #4D4D4D） | 右だけ 4 | なし（「...」は文字） |
| 状態の枠（`IndexStatusBox`） | 幅 Fill・高さ 中身（72） | 上下 8・左右 10・行の間 4 | Bg.Stripe（#FAFBFC） | 1 Border.Soft（#E0E2E5） | | | 4 | |
| 枠の見出し「インデックス」 | 中身 | バッジとの間 8 | | | Label（12 SemiBold） | Ink.Strong（#202124） | | |
| 状態のバッジ | 中身・高さ 18 | 上下 2・左右 8 | `index_list.md` の Level の表 | | Chip（11 Medium） | 同 | 3 | |
| 状態の 2 行 | 幅 Fill | | | | Meta（11） | Ink.Body（#5F6368） | | |
| 情報の表（info-table） | 幅 Fill | | Bg.Surface（#FFFFFF） | 1 Border.Normal（#D9DEE3） | | | 4 | |
| 表の見出しの行 | 高さ 26 | 上下 6・左右 10。値の列 130 | Bg.Window（#F5F7FA） | | Chip（11 Medium） | Ink.Value（#212126） | | |
| 表の行 | 高さ 24 | 上下 5・左右 10。値の列 130・右寄せ | | 上 1 Border.Row（#EDF0F2） | Meta（11） | 項目 Ink.Muted（#6B737D）・値 Ink.Value（#212126） | | |
| 高速検索の枠（fast-search-box） | 幅 Fill・高さ 中身（99。状態だけのとき 40） | 上下 10・左右 12・行の間 6 | Bg.Surface（#FFFFFF） | 1 Border.Normal（#D9DEE3） | | | 4 | |
| 高速検索の項目の名前 | 幅 90 | 値との間 8 | | | Meta（11） | Ink.Muted（#6B737D） | | |
| 高速検索の値（理由・直し方・最終確認・進み具合の文） | Fill | | | | Meta（11） | Ink.Value（#212126） | | |
| 進み具合の線 | 幅 60〜100・高さ 4 | 文との間 8（折り返したときは上下 4） | 下地 Border.Soft（#E0E2E5）。中: 可のとき Ok（#218A21）・反映中のとき File.Folder（#E8A121 をまとめた #E8A020） | | | | 2（「決めること」8） | |
| 注記（FastSearchNote） | 幅 Fill、折り返す | | | | Micro（10 Regular） | Ink.Muted（#6B737D） | | |

## 5. 状態ごとの見え方

状態に入る条件は、9 章の判断層の入力。行の状態（「最新」など）は `index_list.md` の `getIndexRowView` と同じ判定。

| 状態（記号） | 入る条件 | タイトル | 名前・パス | 状態の枠 | 情報の表（対象・xlsx・docx・pptx・サイズ・最終更新） | 高速検索 | 高さ |
|---|---|---|---|---|---|---|---|
| 既定・X・thumb-3 | 営業部（最新・可） | 営業部 - 詳細 | 営業部・C:\Share\営業部\2024\見積もり | 最新／最終更新 2024/10/14 15:30／245 ファイル | 245・147・61・37・9.1 MB・2024/10/14 15:30 | 状態 可・線 100%・「反映済み 150 / 150 フォルダ」・最終確認 2026/09/27 10:05・注記 | 397 |
| X-C・thumb-4 | 更新を終えた | 同 | 同 | 最新／最終更新 2024/10/14 15:42／245 ファイル | 同（最終更新 15:42） | X と同じ | 397 |
| P-X-edit（148） | インデックス名の欄にフォーカス | 同 | 名前の欄の枠が 2 px の Accent | X と同じ | X と同じ | X と同じ | 397 |
| X-N・thumb-2 | まだ取り込んでいない（未作成） | 営業部 - 詳細 | 営業部・パス | 未作成／「まだ取り込んでいません」（2 行目なし） | 0・0・0・0・-・- | 状態 －（ほかの行と注記を隠す） | 338 |
| 高速検索が使えない | Windows Search に接続できない（`reason`＝`NoConnection`） | 営業部 - 詳細 | X と同じ | X と同じ | X と同じ | 状態 不可・理由「Windows Search に接続できません。」・直し方「Windows Search のサービスが実行中かを確認します。」・最終確認・注記（線と進み具合は隠す） | 417（最大 397 を超えるので中が縦にスクロール） |
| X-取り込み後・thumb-5 | インポートした（`imported`） | 営業部 - 詳細 | 営業部・C:\共有\営業部 | 検索できます／最終更新 2026/09/27 10:00／2,530 ファイル | 2,530・1,518・633・379・94.0 MB・2026/09/27 10:00 | 状態 不可・理由「このインデックスには高速検索用のデータがありません。」・直し方の行は出さない・最終確認・注記 | 397 |
| X-取り込み後2 | インポートして、元のフォルダが無い（`folderState`＝`NotSet`） | 同 | 同 | 元のフォルダ未設定／同 | 同 | X-取り込み後と同じ | 397 |
| X-高速検索-反映中・thumb-6 | 顧客を選んだ（要更新・反映中） | 顧客 - 詳細 | 顧客・C:\Share\顧客 | 要更新／最終更新 2024/10/14 15:28／1,830 ファイル | 1,830・1,098・458・274・68.0 MB・2024/10/14 15:28 | 状態 反映中 62%・線 62%・「反映済み 93 / 150 フォルダ（反映待ち 57）」・最終確認・注記 | 397 |
| X0・thumb | 一覧に行が無い | 営業部 - 詳細（Figma のまま） | 前の作り（ステータス 最新・最終更新 2024/10/14 15:30 の欄） | なし | なし（情報の表だけ） | なし | 266 |

- X0 の中身は前の作りのまま残っていて、一覧が空のときの見え方として読めない（「決めること」1）。
- 高速検索の各行の出し方（9 章の `getFastSearchDetailView` の出力）:

| 状態 | 状態のバッジ | 進み具合の行 | 理由 | 直し方 | 最終確認 | 注記 |
|---|---|---|---|---|---|---|
| 可 | 可（Ok） | 出す（線 100%） | 隠す | 隠す | 出す | 出す |
| 反映中 N% | 反映中 N%（Warn） | 出す（線 N%） | 隠す | 隠す | 出す | 出す |
| 反映待ち | 反映待ち（Warn） | 出す（線 0%）。Figma に無い（「決めること」3） | 隠す | 隠す | 出す | 出す |
| 不可（接続できない・対象外） | 不可（Ng） | 隠す | 出す | 出す | 出す | 出す |
| 不可（インポートしたもの） | 不可（Ng） | 隠す | 出す | 隠す | 出す | 出す |
| － | －（None） | 隠す | 隠す | 隠す | 隠す | 隠す |

## 6. 文言

| 文言の ID | 全文 | 使う状態 |
|---|---|---|
| index.detail.title | {インデックス名} - 詳細（名前の後ろに半角の空白・半角のハイフン・半角の空白） | すべて |
| index.detail.basic | 基本設定 | すべて |
| index.detail.basic.help | 未定 | すべて |
| index.detail.name | インデックス名 | すべて |
| index.detail.folder | フォルダパス | すべて |
| index.detail.browse | ...（半角のピリオド 3 つ） | すべて |
| index.detail.status.head | インデックス | X0 以外 |
| index.detail.status.updated | 最終更新 {yyyy/MM/dd HH:mm} | 取り込んだもの |
| index.detail.status.files | {件数} ファイル（件数は 3 桁ごとに「,」） | 取り込んだもの |
| index.detail.status.none | まだ取り込んでいません | X-N |
| index.detail.info | インデックス情報 | すべて |
| index.detail.info.help | 未定 | すべて |
| index.detail.info.col.item | 項目 | すべて |
| index.detail.info.col.value | 値 | すべて |
| index.detail.info.files | 対象ファイル数 | すべて |
| index.detail.info.xlsx | Excel (.xlsx) | すべて |
| index.detail.info.docx | Word (.docx) | すべて |
| index.detail.info.pptx | PowerPoint (.pptx) | すべて |
| index.detail.info.size | インデックスサイズ | すべて |
| index.detail.info.updated | 最終更新 | すべて |
| index.detail.info.size.value | {数} MB（小数 1 桁。例 9.1 MB・94.0 MB）。取り込んでいなければ - | すべて |
| index.detail.info.empty | -（半角のハイフン） | X-N のサイズ・最終更新 |
| index.detail.fast | 高速検索 | X0 以外 |
| index.detail.fast.state | 状態 | 同 |
| index.detail.fast.progress | 反映済み {F−W} / {F} フォルダ（{W} が 0 のとき） | 可 |
| index.detail.fast.progress.wait | 反映済み {F−W} / {F} フォルダ（反映待ち {W}）（括弧は全角） | 反映中 |
| index.detail.fast.reason | 理由 | 不可 |
| index.detail.fast.fix | 直し方 | 不可（インポートしたものを除く） |
| index.detail.fast.checked | 最終確認 | －以外 |
| index.detail.fast.checked.value | {yyyy/MM/dd HH:mm} | 同 |
| index.detail.fast.note | 反映待ちのフォルダは通常の検索で調べます。検索結果は変わりませんが、時間がかかります。 | －以外 |
| index.detail.fast.reason.noconn | Windows Search に接続できません。 | 不可（NoConnection） |
| index.detail.fast.fix.noconn | Windows Search のサービスが実行中かを確認します。 | 同 |
| index.detail.fast.reason.nodata | このインデックスには高速検索用のデータがありません。 | 不可（インポートしたもの） |
| index.detail.fast.reason.scope | 未定（Figma に無い。「決めること」4） | 不可（NotInScope） |
| index.detail.fast.fix.scope | 未定（同） | 同 |

## 7. 操作と振る舞い

| きっかけ | 何が起きるか | 呼ぶ関数 | 次の状態 |
|---|---|---|---|
| 一覧で行を選ぶ | タイトル・値をその行のものに入れ替える。スクロールは先頭に戻す | `updateIndexDetail` | X・X-高速検索-反映中 など |
| `IndexNameBox` にフォーカス | 枠を 2 px の Accent（#0078D4）にする | | P-X-edit |
| `IndexNameBox` の名前を確定する（確定の仕方は「決めること」5） | 入力を調べ、よければ名前を変えて保存する。一覧の名前・タイトルも変える。だめなら理由を出す（出し方は未定） | `testIndexEditInput`・`saveIndexName`（新規。今の `editIndex` の名前の部分） | X |
| `BrowseIndexFolderButton` をクリック | フォルダを選ぶダイアログ。選んだら元のフォルダを変えて保存し、状態を「要更新」にする | `editIndexFolder`（一覧の［再設定］［元のフォルダを設定］と同じ） | P-X-refolder |
| `IndexFolderBox` にホバー | 全文のツールチップ | | |
| 境目（`IndexSplitter`）をドラッグ | 詳細欄の高さを 240〜397 で変える | 8 章 | |
| 更新中・高速検索の確かめが終わった | 選んでいる行なら値を入れ替える | `updateIndexDetail` | |

- 更新中・削除中・エクスポート／インポート中に名前とフォルダを変えられるかは未定（「決めること」6）。今は作成中に名前の変更を止める（E32）。
- Tab の順: `IndexNameBox` → `BrowseIndexFolderButton`（`IndexFolderBox` は読むだけなので Tab で止めない案）。
- ホバー・押したときの見た目（`BrowseIndexFolderButton`）は Figma に無い（「決めること」7）。

## 8. リサイズ

- 高さ:
  - 既定 397、最小 240、最大 397。窓を高くしても 397 のまま。
  - 窓を低くすると、一覧より先に 240 まで縮む。
  - 中は `ScrollViewer` で縦にスクロールする（`Auto`）。上の帯（40）はスクロールしない。
- 縮む順番を決めるコード（`index_list.md` の 8 章と同じ）。インデックス管理の画面の Grid の `SizeChanged` で次を行う。
  - `$detail = [Math]::Max(240, [Math]::Min($userMax, $grid.ActualHeight - 5 - 225))`
    - `$userMax` は境目のドラッグで決めた高さ。初めは 397。
  - `IndexDetailRoot.Height` に入れる。
- 境目のドラッグ: `GridSplitter` の `DragDelta` で `$userMax` を 240〜397 に丸めて決め直し、上の式を呼ぶ。`GridSplitter` だけに任せると、行の高さが固定の値になり、窓を低くしたときに先に縮まなくなる。
- 横:
  - 主な領域いっぱい。左右の列は等分（`*`）で、区切りの左右に 32。
  - 値の行（入力欄・状態の 2 行・情報の表・最終確認）は 1 行で `TextTrimming="CharacterEllipsis"`。
  - 理由・直し方・注記は折り返す（`TextWrapping="Wrap"`）。
- 進み具合の線と文（`FastSearchProgressBox` の `SizeChanged`）:
  1. `$avail` は箱の幅、`$textW` は文を `Measure` した幅。
  2. `$avail - 8 - $textW -ge 60` なら 1 行にする。
     - 線の幅は `[Math]::Min(100, $avail - 8 - $textW)`。
     - 文は 0 行目の 1 列目に置き、左の間を 8 にする。
  3. そうでなければ 2 行にする。
     - 線の幅は `[Math]::Min(100, [Math]::Max(60, $avail))`。
     - 文は 1 行目の 0 列目に置き、上の間を 4・左の間を 0 にする。
  - 文は切らない（`NoWrap`・`TextTrimming="None"`）。
- 1024×640（408:33785）: 詳細欄は 804×354。左右の列は各 354 前後で、中が縦にスクロールする（397 − 354 ＝ 43 ぶん）。
- 1280×820: 1060×397（Figma の X と同じ）。
- 1600×1000（408:33812）: 1380×397。左右の列が広がるだけ。

## 9. 判断層（`ui/index/index_detail_view.ps1`）

今の `index_view.ps1` の名前の付け方（`get<何>View`・`format<何>`）に合わせる。

### getIndexDetailTitle

| name | → |
|---|---|
| 営業部 | 営業部 - 詳細 |
| 顧客 | 顧客 - 詳細 |
| "" | "" |

### getIndexDetailStatusView

- 入力: `rowView`（`getIndexRowView` の出力）、`stat`（Total; Done）、`updatedAt`（DateTime か $null）。
- 出力: `@{ Text; Level; Line1; Line2 }`。

| rowView.Text | stat.Total | updatedAt | → Line1 | Line2 |
|---|---|---|---|---|
| 最新 | 245 | 2024/10/14 15:30 | 最終更新 2024/10/14 15:30 | 245 ファイル |
| 要更新 | 1830 | 2024/10/14 15:28 | 最終更新 2024/10/14 15:28 | 1,830 ファイル |
| 検索できます | 2530 | 2026/09/27 10:00 | 最終更新 2026/09/27 10:00 | 2,530 ファイル |
| 未作成 | 0 | $null | まだ取り込んでいません | "" |

- Text・Level は rowView をそのまま写す。

### getIndexInfoRows

- 入力: `stat`（Total; Xlsx; Docx; Pptx）、`sizeBytes`（long か $null）、`updatedAt`。
- 出力: `@{ Label; Value }` の配列（6 行）。

| stat | sizeBytes | updatedAt | → Value の並び |
|---|---|---|---|
| @{Total=245;Xlsx=147;Docx=61;Pptx=37} | 9542041 | 2024/10/14 15:30 | 245・147・61・37・9.1 MB・2024/10/14 15:30 |
| @{Total=2530;Xlsx=1518;Docx=633;Pptx=379} | 98566144 | 2026/09/27 10:00 | 2,530・1,518・633・379・94.0 MB・2026/09/27 10:00 |
| $null | $null | $null | 0・0・0・0・-・- |

- テキストファイル（.txt など）は件数の行が無い（「決めること」2）。

### formatIndexSize

- MB は 1024×1024 で割り、小数 1 桁に四捨五入する。

| bytes | → |
|---|---|
| 9542041 | 9.1 MB |
| 98566144 | 94.0 MB |
| 71303168 | 68.0 MB |
| $null | - |
| 0 | 未定（「決めること」9） |

### getFastSearchDetailView

- 入力: `getFastSearchRowView` と同じ（`reason`・`progress`・`name`・`hasContent`・`checkedAt`・`indexing`）。
- 出力: `@{ Text; Level; ShowProgress = [bool]; Percent = [int]; ProgressText; Reason; Fix; CheckedText; ShowNote = [bool] }`。
  - Text・Level は `getFastSearchRowView` を呼んで写す。判定を二重に書かない。

| reason | ByIndex[name] | hasContent | checkedAt | → Text | ShowProgress | Percent | ProgressText | Reason | Fix | CheckedText | ShowNote |
|---|---|---|---|---|---|---|---|---|---|---|---|
| Ok | @{Folders=150;Waiting=0} | $true | 2026/09/27 10:05 | 可 | $true | 100 | 反映済み 150 / 150 フォルダ | "" | "" | 2026/09/27 10:05 | $true |
| Ok | @{Folders=150;Waiting=57} | $true | 2026/09/27 10:05 | 反映中 62% | $true | 62 | 反映済み 93 / 150 フォルダ（反映待ち 57） | "" | "" | 2026/09/27 10:05 | $true |
| Ok | @{Folders=150;Waiting=150} | $true | 2026/09/27 10:05 | 反映待ち | $true | 0 | 反映済み 0 / 150 フォルダ（反映待ち 150） | "" | "" | 2026/09/27 10:05 | $true |
| NoConnection | （無し） | $true | 2026/09/27 10:05 | 不可 | $false | 0 | "" | Windows Search に接続できません。 | Windows Search のサービスが実行中かを確認します。 | 2026/09/27 10:05 | $true |
| Ok | （無し） | $true | 2026/09/27 10:05 | 不可 | $false | 0 | "" | このインデックスには高速検索用のデータがありません。 | "" | 2026/09/27 10:05 | $true |
| NoFolder | （無し） | $false | $null | － | $false | 0 | "" | "" | "" | "" | $false |
| Ok | progress が $null | $true | 2026/09/27 10:05 | － | $false | 0 | "" | "" | "" | "" | $false |

- 反映待ちの ProgressText は Figma に無い。反映中の形にそろえた案（「決めること」3）。
- NotInScope の Reason・Fix は Figma に無い（「決めること」4）。
- 「－」のとき最終確認と注記を隠すのは、X-N の見え方による。確かめられなかった「－」でも隠すかは「決めること」10。

### getFastSearchBarFill

| Text | → 中の色のキー |
|---|---|
| 可 | Ok |
| 反映中 62% | File.Folder |
| 反映待ち | File.Folder |

## 10. 画面層（`ui/index/index_detail.ps1`）

- 読み込みのとき、`FindName` で 3 章の x:Name を引く。
- 配線:
  - `IndexNameBox`: `GotKeyboardFocus`・`LostKeyboardFocus` は見た目だけ（Style の Trigger で足りるなら不要）。確定のイベントは「決めること」5。
  - `BrowseIndexFolderButton.Click` → `editIndexFolder`。
  - `FastSearchProgressBox.SizeChanged` → 8 章の 1 行／2 行の切り替え。
  - インデックス管理の画面の Grid の `SizeChanged`、`IndexSplitter.DragDelta` → 詳細欄の高さ（8 章）。
- `updateIndexDetail`（一覧の選択が変わったとき・行の値が変わったときに一覧の画面層から呼ぶ）:
  1. 選んでいる行が無ければ、詳細欄の中身をどうするか（「決めること」1）。
  2. `getIndexDetailTitle` → `IndexDetailTitle.Text`。
  3. 名前・パスを入れる。入力中（`IndexNameBox` にフォーカスがある）なら、名前は上書きしない。
  4. `getIndexDetailStatusView` → バッジの `Tag`・文字、`IndexStatusLine1/2`（Line2 が空なら Collapsed）。
  5. `getIndexInfoRows` → `IndexInfoRows.ItemsSource`。
  6. `getFastSearchDetailView` を呼ぶ。
     - 結果をバッジ・各行の `Visibility`・`FastSearchProgressFill.Width`（線の幅 × Percent / 100）と色（`getFastSearchBarFill`）に写す。
     - Reason・Fix・CheckedText が空の行は Collapsed にする。
- インデックスのサイズ・種類ごとの件数は、ファイルを数えるので別のスレッドのジョブで求め、`Dispatcher.BeginInvoke` で戻して書く。求める間の表示は未定（「決めること」9）。高速検索の状態は、一覧が持っている最後の確かめの結果を使い、ここでは問い合わせない。

## 11. 受け入れ

- 見比べる画像:
  - 1280×820 で、Figma の get_screenshot と比べる。主な領域は 1060 幅、詳細欄は 397。
    - 352:37690（X）
    - 352:38995（X-高速検索-反映中）
    - 352:38081（高速検索が使えない）
    - 352:37892（X-N）
    - 352:38785（X-取り込み後）
  - 1024×640 で 408:33785 と比べる（詳細欄 354）。
- 確かめる点:
  - 上の帯の高さが 40 で、地が #F9FAFA、下に 1px の #E0E2E5 がある。
  - タイトルが「営業部 - 詳細」（12 Bold）で、info のアイコンとの間が 8 である。
  - 左右の列の間に 1px の #E5E8ED の線があり、その左右が 32 である。
  - 入力欄の高さが 28、枠が #D9DEE3、角丸が 4 である。［...］が 34×28 で地が #F2F2F5 である。
  - 状態の枠の地が #FAFBFC、高さが 72（X）である。
  - 情報の表の見出しの行の地が #F5F7FA で、値の列が 130 幅で右寄せである。行の間に #EDF0F2 の線がある。
  - 高速検索の枠の高さが 99（X）・40（X-N）である。
  - 進み具合の線が 100×4 で、可のとき #218A21・反映中のとき #E8A020 で割合どおりに塗られている。
  - 1024×640 で、進み具合の文が切れていない。入らないときは線の下に 4 空けて折り返している。
  - 1024×640 で、詳細欄の中が縦にスクロールする。一覧は 225 のままである。
  - 長いパスが 1 行で「…」になる。

## 決めること

1. **一覧に行が無いとき・行を選んでいないときの詳細欄。**
   - X0 の variant は前の作り（ステータス・最終更新の欄）のままで、空の一覧と合わない。
   - X では一覧の行を選んでいないのに営業部が出ている。決めることは 3 つ。
     - 選んでいないときに先頭の行を出すか。
     - 空のときに詳細欄を隠すか、空の案内を出すか。
     - hidden の `field-ステータス`・`field-最終更新` を Figma から消すか。
2. **テキストファイルの件数の行。** 情報の表は Excel・Word・PowerPoint の 3 行だけで、.txt などの行が無い（テキストファイルも検索の対象になっている）。足すか、「その他」にまとめるか。
3. **反映待ちのときの進み具合の行。** Figma に反映待ちの詳細が無い。「反映済み 0 / 150 フォルダ（反映待ち 150）」と線 0% を出す案にした。
4. **対象外（NotInScope）の理由と直し方。** Figma にあるのは「接続できない」と「インポートしたもの」だけ。対象外のときの全文を決める。今のツールチップは「ワークスペースが Windows Search の索引の対象外。［インデックスのオプション］でワークスペースの system_index を対象に加える（管理者の権限が要る PC では、PC の管理者に頼む）」。
5. **名前の変更の確定の仕方。**
   - P-X-edit はフォーカスの見た目だけ。何で確定するか（フォーカスが外れたとき・Enter・保存のボタン）が Figma に無い。
   - 入力の誤り（同じ名前・使えない文字）の出し方も無い。
6. **更新中などに名前・フォルダを変えられるか。** 更新中・削除中・エクスポート／インポート中に、入力欄と［...］を無効にするか。
7. **［...］の文字の Style と見た目。**
   - 12 Medium に当たるキーが theme に無い。
   - ホバー・押したときの見た目が Figma に無い。
8. **進み具合の線の角丸。** 2 は一覧のバーから当てた値で、詳細欄の線では読み取っていない。
9. **インデックスのサイズ。** 1 MB 未満・0 のときの表し方（KB を使うか「0.0 MB」か）と、求めている間の表示。
10. **「－」のときの最終確認と注記。** X-N（まだ作っていない）は状態だけを出す。確かめられなかった「－」と更新中の「－」でも同じにするか。
11. **help-icon のツールチップ。** 「基本設定」「インデックス情報」の help-icon の文言が読み取れなかった。
12. **理由・直し方の折り返し。** 注記は「値の行は 1 行で…」だが、Figma の理由・直し方の文は折り返す作り。この設計書は折り返すことにした。
