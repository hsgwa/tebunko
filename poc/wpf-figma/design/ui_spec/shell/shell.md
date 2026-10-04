# 窓（shell）

## 1. 役割

アプリの窓そのもの。上のタイトルバー（名前と窓のボタン）と、その下の本体を受け持つ。本体は左の欄（ナビ）と主な領域に分け、主な領域の下にステータスバーを置く。各画面の中身は差し替えの口（`NavHost`・`ContentHost`・`StatusBarHost`）に差すだけで、この領域は画面の中身を知らない。

- Figma の部品: `xaml/shell/shell` 340:1650（variant なし）。中のタイトルバー 340:1599、本体 `app-body-split`、ステータスバーのインスタンス 354:213。
- 使う画面: 106・148 のすべての画面。見本は 148 の H0（148:67・`../png/148_67.png`）、106 の H（108:230・`../png/108_230.png`）。
- Figma の注記（全文）: 「窓: 最小 1024×640、既定 1280×820。上のタイトルバーは高さ 32 固定・横いっぱい（左に名前、右に窓のボタン）。その下を左の欄と主な領域に分ける。左の欄は既定 220、最小 180、最大 360（右端をドラッグで変える）。主な領域は残りの幅と高さを全部使う。」

## 2. 置き場所

- XAML: `scripts/tebunko/xaml/tebunko.xaml`（Window。今のファイルを置き換える）。
- 画面層: `scripts/tebunko/ui/shell/shell.ps1`。判断層: `scripts/tebunko/ui/shell/shell_view.ps1`。
- 読み込み: `gui.ps1` が `tebunko.xaml` を読み、`theme.xaml`（`../../shared/xaml/theme.xaml`）を MergedDictionaries に入れる（今と同じ）。続けて次を読み込み、口に差す。
  - `xaml/shell/nav.xaml` → `NavHost.Content`
  - `xaml/shell/status_bar.xaml` → `StatusBarHost.Content`
  - 選んでいる画面の XAML（`xaml/search/search.xaml` など）→ `ContentHost.Content`
- 読み込みの仕方は `figma_wpf_structure.md` の移行の案のとおり: `loadXaml` で読み、`ContentControl.Content` に差し、読んだファイルの `x:Name` を `$ui` に足す。

## 3. 部品の木

```xml
<Window x:Name="MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="{Window.Title}"
        Width="1280" Height="820" MinWidth="1024" MinHeight="640"
        WindowStartupLocation="CenterScreen"
        WindowStyle="None" ResizeMode="CanResize"
        Background="{StaticResource Bg.Window}"
        BorderBrush="{StaticResource Border.Window}" BorderThickness="1"
        FontFamily="Segoe UI, Yu Gothic UI">
  <!-- 自前のタイトルバーにする（WindowChrome を使うかは「決めること」） -->
  <WindowChrome.WindowChrome>
    <WindowChrome CaptionHeight="32" ResizeBorderThickness="4" GlassFrameThickness="0" CornerRadius="0" UseAeroCaptionButtons="False"/>
  </WindowChrome.WindowChrome>
  <Window.Resources>
    <ResourceDictionary>
      <ResourceDictionary.MergedDictionaries>
        <ResourceDictionary Source="../../shared/xaml/theme.xaml"/>
      </ResourceDictionary.MergedDictionaries>
    </ResourceDictionary>
  </Window.Resources>

  <Grid x:Name="WindowRoot">
    <Grid.RowDefinitions>
      <RowDefinition Height="32"/>   <!-- タイトルバー -->
      <RowDefinition Height="*"/>    <!-- 本体 -->
    </Grid.RowDefinitions>

    <!-- タイトルバー（340:1599） -->
    <Border x:Name="TitleBar" Grid.Row="0"
            Background="{StaticResource Bg.TitleBar}"
            BorderBrush="{StaticResource Border.Window}" BorderThickness="0,0,0,1"
            Padding="8,0,8,0">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>  <!-- title-left -->
          <ColumnDefinition Width="*"/>     <!-- 間（ドラッグで窓を動かす所） -->
          <ColumnDefinition Width="Auto"/>  <!-- window-controls -->
        </Grid.ColumnDefinitions>

        <StackPanel x:Name="TitleLeft" Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
          <Border x:Name="AppIcon" Width="32" Height="32" CornerRadius="4"
                  Background="{StaticResource Accent}" VerticalAlignment="Center">
            <!-- 新しいロゴ（illust/app_logo.svg）22×22 を真ん中に。地の色・角丸は無し。下の Path は前の形なので、ロゴの絵に置き換える -->
            <Viewbox Width="10" Height="10" HorizontalAlignment="Center" VerticalAlignment="Center">
              <Path x:Name="AppIconGlyph" Stroke="{StaticResource Ink.OnAccent}" StrokeThickness="2"
                    StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round"
                    Data="{lucide:file-search}"/>
            </Viewbox>
          </Border>
          <TextBlock x:Name="AppTitleText" Margin="8,0,0,0" VerticalAlignment="Center"
                     Style="{StaticResource Brand}" Foreground="{StaticResource Ink.Strong}"
                     Text="{Title.AppName}"/>
        </StackPanel>

        <StackPanel x:Name="WindowControls" Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
          <Button x:Name="MinimizeButton" Width="46" Height="32" Style="{StaticResource TitleBarButton}"
                  WindowChrome.IsHitTestVisibleInChrome="True">
            <Rectangle Width="14" Height="1" RadiusX="1" RadiusY="1" Fill="{StaticResource Ink.Strong}"/>
          </Button>
          <Button x:Name="MaximizeButton" Width="46" Height="32" Style="{StaticResource TitleBarButton}"
                  WindowChrome.IsHitTestVisibleInChrome="True">
            <Border Width="12" Height="12" CornerRadius="1" BorderThickness="1" BorderBrush="{StaticResource Ink.Strong}"/>
          </Button>
          <Button x:Name="CloseButton" Width="46" Height="32" Style="{StaticResource TitleBarButton}"
                  WindowChrome.IsHitTestVisibleInChrome="True">
            <!-- Lucide x、線 0.9333、Danger.Text -->
            <Path x:Name="CloseGlyph" Stroke="{StaticResource Danger.Text}" StrokeThickness="0.9333"
                  StrokeStartLineCap="Round" StrokeEndLineCap="Round" Data="{lucide:x}"/>
          </Button>
        </StackPanel>
      </Grid>
    </Border>

    <!-- 本体（app-body-split） -->
    <Grid x:Name="AppBody" Grid.Row="1">
      <Grid.ColumnDefinitions>
        <ColumnDefinition x:Name="NavColumn" Width="220" MinWidth="180" MaxWidth="360"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>

      <!-- 左の欄。線とつまみの絵は nav.xaml が描く -->
      <ContentControl x:Name="NavHost" Grid.Column="0" Focusable="False"
                      HorizontalContentAlignment="Stretch" VerticalContentAlignment="Stretch"/>

      <!-- 幅を変える口。見た目は透明で、nav.xaml の右端の線（1px）とつまみに重ねる -->
      <GridSplitter x:Name="NavSplitter" Grid.Column="0" Width="5"
                    HorizontalAlignment="Right" VerticalAlignment="Stretch"
                    ResizeDirection="Columns" ResizeBehavior="CurrentAndNext"
                    Background="Transparent" Focusable="False" Cursor="SizeWE"/>

      <!-- 主な領域（right）。ステータスバーは主な領域の下だけに付く（左の欄の下には付かない） -->
      <Grid x:Name="MainArea" Grid.Column="1">
        <Grid.RowDefinitions>
          <RowDefinition Height="*"/>    <!-- ContentHost -->
          <RowDefinition Height="24"/>   <!-- StatusBarHost -->
        </Grid.RowDefinitions>
        <Border Grid.Row="0" Background="{StaticResource Bg.Stripe}">
          <ContentControl x:Name="ContentHost" Focusable="False"
                          HorizontalContentAlignment="Stretch" VerticalContentAlignment="Stretch"/>
        </Border>
        <ContentControl x:Name="StatusBarHost" Grid.Row="1" Focusable="False"
                        HorizontalContentAlignment="Stretch" VerticalContentAlignment="Stretch"/>
      </Grid>
    </Grid>
  </Grid>
</Window>
```

- `{lucide:<名前>}` は lucide-static 1.50.0 の SVG の path を Geometry に写したもの（`../lucide/` にある）。1 つのアイコンが複数の path のときは `GeometryGroup` にする。
- `TitleBarButton` は新しく足す Style（Background=`Bg.TitleBar`、BorderThickness=0、中身を中央）。ホバー・押したときの見た目は未定。

## 4. 寸法と色

| 要素 | 大きさ | 余白 | 地の色 | 枠 | 文字の Style | 文字の色 | 角丸 | アイコン |
|---|---|---|---|---|---|---|---|---|
| MainWindow | 既定 1280×820、最小 1024×640 | – | Bg.Window（#F5F7FA） | 外周の線は未定（148_67 では灰色の 1px） | – | – | 0 | – |
| TitleBar | 横 Fill × 32 | 左右 8 | Bg.TitleBar（#F0F0F0） | 下 1px Border.Window（#999999） | – | – | 0 | – |
| TitleLeft | Hug | 子の間 8 | – | – | – | – | – | – |
| AppIcon | 32×32 | – | 無し | – | – | – | 0 | アプリのロゴ（`illust/app_logo.svg`）22×22 を真ん中に（README の 10） |
| AppTitleText | Hug | – | – | – | Brand（12 SemiBold・字間 0.2） | Ink.Strong（#202124） | – | – |
| WindowControls | 138×32（46×3）。右端は窓の右から 8 | – | – | – | – | – | – | – |
| MinimizeButton | 46×32 | – | Bg.TitleBar（#F0F0F0） | – | – | – | – | 14×1 の棒、Ink.Strong、角丸 1 |
| MaximizeButton | 46×32 | – | Bg.TitleBar | – | – | – | – | 12×12 の四角、1px Ink.Strong、角丸 1 |
| CloseButton | 46×32 | – | Bg.TitleBar | – | – | – | – | Lucide `x`、線 0.9333（11.2 相当）、Danger.Text（#D13438） |
| AppBody | 1280×788 | 0 | – | – | – | – | – | – |
| NavColumn / NavHost | 既定 220、最小 180、最大 360、縦 Fill | – | nav.xaml が持つ（Bg.Subtle） | – | – | – | – | – |
| NavSplitter | 5×縦 Fill（透明） | – | Transparent | – | – | – | – | – |
| MainArea | Fill（1280 のとき 1060×788） | 0 | – | – | – | – | – | – |
| ContentHost | Fill（1060×764） | 0 | Bg.Stripe（#FAFBFC） | – | – | – | – | – |
| StatusBarHost | 横 Fill × 24 | – | status_bar.xaml が持つ | – | – | – | – | – |

- AppIcon は Figma のメタデータで 32×32・角丸 4。`diff_round2.md` には「40×32 の青い札」とあり、食い違う（決めること）。
- 閉じるの印の大きさは Figma のレイヤーの枠が 9×18 で、線の太さからは 11.2 相当。正確な描き方は未定（決めること）。
- フォントは `Segoe UI, Yu Gothic UI`（英数字は Segoe UI、日本語は Yu Gothic UI）。決まりは `figma_theme_tokens.md` の「文字」（2026-10-04 から試し中）。

## 5. 状態ごとの見え方

窓そのものに Figma の variant は無い。選んでいる画面で、口に差すものだけが変わる。

| 状態（選んでいる画面） | 入る条件（shell_view の入力 `screen`） | NavHost | ContentHost | StatusBarHost |
|---|---|---|---|---|
| 検索（H・H0・H-R・E13・E14 など） | `search` | nav（検索の variant） | `search/search.xaml` | status_bar |
| インデックス管理（X・X0・X-U・X-P など） | `index` | nav（インデックス管理の variant） | `index/index_list.xaml` と詳細欄 | status_bar |
| 設定（C） | `settings` | nav（設定の variant） | `settings/settings.xaml` | status_bar |
| Office の終了（P） | `office` | nav（Office の終了の variant） | `office/office.xaml` | status_bar |

| 窓の状態 | 見え方 |
|---|---|
| 通常 | 上の部品の木のとおり |
| 最大化 | 未定（最大化の印を「元に戻す」の印に替えるか、WindowChrome の最大化のときのはみ出し分の余白を足すか） |
| 窓が前面でない | 未定 |

## 6. 文言

| 文言の ID | 全文 | 使う状態 |
|---|---|---|
| Window.Title | `tebunko` | すべて（タスクバーの名前） |
| Title.AppName | `tebunko` | すべて（タイトルバーの左） |

- 窓のボタンのツールチップ（「最小化」など）は Figma に無い。出すかどうかは未定。

## 7. 操作と振る舞い

| きっかけ | 何が起きるか | 呼ぶ関数 | 次の状態 |
|---|---|---|---|
| タイトルバーの空き（TitleLeft と WindowControls の間・TitleLeft の上）をドラッグ | 窓を動かす（WindowChrome の CaptionHeight=32 による） | – | – |
| タイトルバーの空きをダブルクリック | 最大化／元に戻す（WindowChrome の既定） | – | – |
| MinimizeButton をクリック | 最小化 | `[System.Windows.SystemCommands]::MinimizeWindow($window)` | – |
| MaximizeButton をクリック | 最大化と元に戻すを切り替える | `MaximizeWindow` / `RestoreWindow` | 最大化 |
| CloseButton をクリック | 窓を閉じる（今の Closing の処理を通す） | `[System.Windows.SystemCommands]::CloseWindow($window)` | – |
| NavSplitter をドラッグ | 左の欄の幅を 180〜360 で変える。主な領域が残りを使う | – | – |
| ナビの項目を選ぶ（nav.md） | 選んだ画面を ContentHost に差す | `showScreen`（画面層） | その画面 |
| 起動 | 検索の画面を選ぶ（今の `$ui.Tabs.SelectedItem = $ui.SearchTab` の置き換え） | `showScreen 'search'` | 検索 |

- 窓のボタンのホバー・押したときの見た目は Figma に無い（未定）。
- Tab の順: TitleBar の窓のボタンは Tab で止めない（今の OS の窓と同じ）。止めるかは未定。本体は NavHost → ContentHost → StatusBarHost の順。
- 窓の大きさ・左の欄の幅を覚えて次の起動で戻すかは Figma に無い（未定）。

## 8. リサイズ

| 部分 | 動き |
|---|---|
| 窓 | 最小 1024×640。既定 1280×820。本体（Row 1）が縦に伸びる |
| タイトルバー | 高さ 32 固定・横 Fill。左（名前）と右（窓のボタン）の間が伸びる |
| 本体 | 左の欄は固定の幅（既定 220、180〜360 を GridSplitter で）。主な領域が Fill |
| 主な領域 | ContentHost が Fill、StatusBarHost は高さ 24 固定 |

- 1024×640 のとき: 主な領域は 804×608（左の欄 220 のとき）。ContentHost は 804×584。
- 1600×1000 のとき: 主な領域は 1380×968、ContentHost は 1380×944。
- 左の欄を 360 にしても、主な領域は最小の窓で 664 残る。主な領域の最小幅は Figma に無い（未定）。
- 文字の切り方: AppTitleText は切らない（固定の文言）。

## 9. 判断層

`scripts/tebunko/ui/shell/shell_view.ps1`

### getScreenParts

画面の名前から、口に差す XAML を返す。今の `Tabs.SelectedItem = $ui.<名前>Tab` の置き換え。

- 入力: `[string]$screen`（`search` / `index` / `settings` / `office`）
- 出力: ハッシュテーブル

| キー | 型 | 中身 |
|---|---|---|
| Content | string | ContentHost に差す XAML（`xaml/` からの相対パス） |
| NavPane | string | NavPaneHost に差す XAML。無ければ `$null` |
| NavItem | string | 選んだ見た目にするナビの項目の x:Name |

```powershell
It "画面 <screen> の部品を返す" -TestCases @(
    @{ screen = 'search';   content = 'search/search.xaml';      navPane = 'search/target_tree.xaml'; navItem = 'SearchTab' }
    @{ screen = 'index';    content = 'index/index_list.xaml';   navPane = $null;                     navItem = 'IndexTab' }
    @{ screen = 'settings'; content = 'settings/settings.xaml';   navPane = $null;                     navItem = 'SettingsTab' }
    @{ screen = 'office';   content = 'office/office.xaml';      navPane = $null;                     navItem = 'KillTab' }
) {
    param ($screen, $content, $navPane, $navItem)
    $parts = getScreenParts $screen
    $parts.Content | Should -Be $content
    $parts.NavPane | Should -Be $navPane
    $parts.NavItem | Should -Be $navItem
}

It "知らない画面の名前は検索の画面にする" {
    (getScreenParts 'unknown').NavItem | Should -Be 'SearchTab'
}
```

- インデックス管理の ContentHost に差すファイル（一覧と詳細欄を 1 つの XAML にするか）は `index/index_list.md` の決めに従う。表の `index/index_list.xaml` はその仮の名前。

## 10. 画面層

`scripts/tebunko/ui/shell/shell.ps1`

| x:Name | 配線 |
|---|---|
| MainWindow | `Loaded` で `showScreen 'search'`。`Closing` は今の処理のまま |
| MinimizeButton | `Add_Click` → `[System.Windows.SystemCommands]::MinimizeWindow($window)` |
| MaximizeButton | `Add_Click` → `$window.WindowState` が `Maximized` なら `RestoreWindow`、でなければ `MaximizeWindow` |
| CloseButton | `Add_Click` → `[System.Windows.SystemCommands]::CloseWindow($window)` |
| NavHost・StatusBarHost | 起動のときに 1 回だけ `loadXaml` して差す |
| ContentHost | `showScreen` が差し替える |

`showScreen($screen)` の手順:

1. `$parts = getScreenParts $screen`
2. 画面ごとの XAML は初めて選んだときに 1 回だけ読み、`$script:screenViews[$screen]` に取っておく（2 回目からは同じものを差す。選び直すたびに読み直さない）。
3. `$ui.ContentHost.Content = $script:screenViews[$screen]`
4. `$ui.NavPaneHost.Content` に `$parts.NavPane` の XAML（無ければ `$null`）を差す。
5. ナビの選んだ見た目を `$parts.NavItem` にする（nav.md の `applyNavView`）。
6. 今の `Tabs.Add_SelectionChanged` でしていたことを移す: `office` なら `refreshProcesses` と `processTimer` の開始、`index` なら `refreshIndexingState`。

- 今 `$ui.Tabs.SelectedItem = $ui.SearchTab` と書いている所（Ctrl+F・起動・`GoIndexTabButton`・`indexing_tab`）は、すべて `showScreen` を呼ぶように直す。
- 重い処理は無い。すべて UI スレッドで行う。

## 11. 受け入れ

見比べる画像: `../png/148_67.png`（H0）と `../png/108_230.png`（H）。窓 1280×820 で撮る。

- タイトルバーの高さが 32、地が #F0F0F0、下の線が 1px の #999999。
- 左端から 8 の所に 32×32・角丸 4 の青（#0078D4）の札があり、中に白いアイコン。札の右 8 に「tebunko」（12 SemiBold）。
- 窓のボタンが 46×32 で 3 つ並び、右端が窓の右から 8。閉じるの印が赤（#D13438）。
- 左の欄の幅が 220、右端に 1px の線と、縦の中央に 3×28 のつまみ（nav.xaml）。ドラッグで 180〜360 に変わり、それより外には動かない。
- 主な領域の地が #FAFBFC。ステータスバー（高さ 24）が主な領域の下だけにあり、左の欄の下には無い。
- 窓を 1024×640 より小さくできない。

## 決めること

- 窓の枠: `WindowStyle="None"`＋`WindowChrome` で自前のタイトルバーにするか、OS の枠のままにするか。自前にするなら、最大化のときのはみ出し分の余白・外周の線（148_67 では灰色の 1px。色は未定）・窓のスナップの扱い。
- AppIcon の大きさ: Figma のメタデータは 32×32・角丸 4、`diff_round2.md` は 40×32。どちらに合わせるか。
- 窓のボタン（最小化・最大化・閉じる）のホバー・押したときの見た目、最大化のときの「元に戻す」の印、窓が前面でないときの見た目、ツールチップの有無。
- 閉じるの印の大きさと描き方（レイヤーの枠 9×18・線 0.9333）。
- 主な領域の最小幅（左の欄を 360 にしたときの下限）。
- 窓の大きさ・左の欄の幅を覚えて次の起動で戻すか。
- 新しい Style `TitleBarButton` を theme に足すこと（`figma_wpf_map.md` の「食い違い」に足す案）。
- インデックス管理の画面の ContentHost に差す XAML の名前（`index/index_list.md` で決める）。
