# 左の欄のナビ（nav）

## 1. 役割

窓の左の欄。上に 4 つの画面（検索・インデックス管理・設定・Office の終了）の切り替えを並べ、下に「バージョン情報」を置く。検索の画面のときは、間に検索対象のツリー（`search/target_tree.md`）を差す。インデックスを作っている間・中断している間は、「インデックス管理」の右に小さな札を出す。今の TabControl（`Tabs`）と `MoreButton`（⋯ のメニュー）の置き換え。

- Figma の部品: `xaml/shell/nav` 353:30193（既定 340:553）。
- variant（16）: 既定 340:553・検索 366:41383・検索-空 366:45903・インデックス管理 366:46031・設定 366:47153・Office の終了 366:47393（106）、P-H0 366:31854・P-H 366:31985・P-H-B 366:35949・P-H-P 366:36476・P-E13 366:37000・P-X0 353:29951・P-X-U 353:29980・P-X-P 353:30008・P-C 353:30037・P-P 353:30066（148）。
- 使う画面: 106・148 のすべての画面。見本は `../png/148_67.png`（H0）・`../png/108_230.png`（H）・`../png/108_1116.png`（H-B）・`../png/108_1419.png`（H-P）・`../png/175_4085.png`（E13）。

## 2. 置き場所

- XAML: `scripts/tebunko/xaml/shell/nav.xaml`（ルートは `Grid`。UserControl 相当）。
- 画面層: `scripts/tebunko/ui/shell/nav.ps1`。判断層: `scripts/tebunko/ui/shell/nav_view.ps1`。
- 親: `tebunko.xaml` の `NavHost`（`shell.md`）。起動のときに 1 回だけ `loadXaml` して `NavHost.Content` に差し、x:Name を `$ui` に足す。
- 子: `NavPaneHost` に、検索の画面のときだけ `xaml/search/target_tree.xaml` を差す（`shell.md` の `showScreen`）。

## 3. 部品の木

```xml
<Grid x:Name="NavRoot"
      xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
      Background="{StaticResource Bg.Subtle}" ClipToBounds="True">
  <Grid.RowDefinitions>
    <RowDefinition Height="Auto"/>  <!-- 0: nav-list（上 8・下 8） -->
    <RowDefinition Height="Auto"/>  <!-- 1: separator（検索の画面だけ） -->
    <RowDefinition Height="*"/>     <!-- 2: NavPaneHost（検索の画面）／空き（spacer。ほかの画面） -->
    <RowDefinition Height="1"/>     <!-- 3: footer-line -->
    <RowDefinition Height="36"/>    <!-- 4: help-about-section -->
  </Grid.RowDefinitions>

  <!-- nav-list: 項目 36、項目の間 1 -->
  <StackPanel x:Name="NavList" Grid.Row="0" Orientation="Vertical" Margin="0,8,0,8">
    <RadioButton x:Name="SearchTab"   GroupName="Nav" Style="{StaticResource NavItem}" Height="36" Margin="0,0,0,1" Content="{Nav.Search}"/>
    <Grid>
      <RadioButton x:Name="IndexTab"  GroupName="Nav" Style="{StaticResource NavItem}" Height="36" Margin="0,0,0,1" Content="{Nav.Index}"/>
      <!-- 札（nav-badge）。左の欄の左端から 163、行の上下の中央 -->
      <Border x:Name="IndexTabBadge" Visibility="Collapsed"
              HorizontalAlignment="Left" VerticalAlignment="Center" Margin="163,0,0,1"
              Height="15" CornerRadius="9" Padding="6,1,6,1"
              Background="{StaticResource Select.Soft}" IsHitTestVisible="False">
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <Ellipse x:Name="IndexTabBadgeDot" Width="6" Height="6" Margin="0,0,4,0"
                   Fill="{StaticResource Accent}" VerticalAlignment="Center"/>
          <TextBlock x:Name="IndexTabBadgeText" Style="{StaticResource Badge.Micro}"
                     Foreground="{StaticResource Accent}" VerticalAlignment="Center" Text="{Nav.Badge.Progress}"/>
        </StackPanel>
      </Border>
    </Grid>
    <RadioButton x:Name="SettingsTab" GroupName="Nav" Style="{StaticResource NavItem}" Height="36" Margin="0,0,0,1" Content="{Nav.Settings}"/>
    <RadioButton x:Name="KillTab"     GroupName="Nav" Style="{StaticResource NavItem}" Height="36" Content="{Nav.Office}"/>
  </StackPanel>

  <Rectangle x:Name="NavSeparator" Grid.Row="1" Height="1" Fill="{StaticResource Border.Separator}"/>

  <ContentControl x:Name="NavPaneHost" Grid.Row="2" Focusable="False"
                  HorizontalContentAlignment="Stretch" VerticalContentAlignment="Stretch"/>

  <Rectangle x:Name="NavFooterLine" Grid.Row="3" Height="1" Fill="{StaticResource Border.Soft}"/>

  <!-- help-about-section: 左 12、縦の中央 -->
  <Grid Grid.Row="4">
    <Button x:Name="AboutLink" Style="{StaticResource NavHelpLink}"
            HorizontalAlignment="Left" VerticalAlignment="Center" Margin="12,0,0,0" Padding="0,4,0,4">
      <StackPanel Orientation="Horizontal">
        <!-- info 14（丸は線 1.1667、中の「i」は README の 9 のとおり線 1.5）、Ink.Body -->
        <Path x:Name="AboutLinkIcon" Width="14" Height="14" Stretch="Uniform"
              Stroke="{StaticResource Ink.Body}" StrokeThickness="1.1667"
              StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round"
              Data="{lucide:info}" VerticalAlignment="Center"/>
        <TextBlock Margin="4,0,0,0" Style="{StaticResource Meta}" Foreground="{StaticResource Ink.Body}"
                   VerticalAlignment="Center" Text="{Nav.About}"/>
      </StackPanel>
    </Button>
  </Grid>

  <!-- 右端の線とつまみ（絵だけ。ドラッグは shell.xaml の NavSplitter が受ける） -->
  <Rectangle x:Name="NavResizeLine" Grid.RowSpan="5" Width="1" HorizontalAlignment="Right"
             VerticalAlignment="Stretch" Fill="{StaticResource Border.Splitter}" IsHitTestVisible="False"/>
  <Rectangle x:Name="NavResizeGrip" Grid.RowSpan="5" Width="3" Height="28" RadiusX="1.5" RadiusY="1.5"
             HorizontalAlignment="Right" VerticalAlignment="Center"
             Fill="{StaticResource Border.Grip}" IsHitTestVisible="False"/>
</Grid>
```

ナビの項目の Style（`NavItem`、theme に新しく足す）:

```xml
<Style x:Key="NavItem" TargetType="RadioButton">
  <Setter Property="Foreground" Value="{StaticResource Ink.Strong}"/>
  <Setter Property="FontSize" Value="13"/>
  <Setter Property="FontWeight" Value="Medium"/>          <!-- Nav（13 Medium） -->
  <Setter Property="Cursor" Value="Hand"/>
  <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
  <Setter Property="Template">
    <Setter.Value>
      <ControlTemplate TargetType="RadioButton">
        <Grid x:Name="Root" Background="Transparent" Height="36">
          <Rectangle x:Name="ActiveBar" Width="3" RadiusX="1" RadiusY="1" HorizontalAlignment="Left"
                     VerticalAlignment="Stretch" Fill="{StaticResource Accent}" Visibility="Collapsed"/>
          <StackPanel Orientation="Horizontal" Margin="12,0,12,0" VerticalAlignment="Center">
            <!-- 14×14 のアイコンの枠。Figma では空（何も描いていない） -->
            <Border x:Name="IconFrame" Width="14" Height="14" Margin="0,0,8,0"/>
            <ContentPresenter VerticalAlignment="Center" RecognizesAccessKey="False"/>
          </StackPanel>
        </Grid>
        <ControlTemplate.Triggers>
          <Trigger Property="IsChecked" Value="True">
            <Setter TargetName="Root" Property="Background" Value="{StaticResource Select.Soft}"/>
            <Setter TargetName="ActiveBar" Property="Visibility" Value="Visible"/>
            <Setter Property="Foreground" Value="{StaticResource Accent}"/>
            <Setter Property="FontWeight" Value="SemiBold"/>    <!-- Heading（13 SemiBold） -->
          </Trigger>
        </ControlTemplate.Triggers>
      </ControlTemplate>
    </Setter.Value>
  </Setter>
</Style>
```

- `Badge.Micro`（10 SemiBold）は theme に無い（`figma_wpf_map.md` の「食い違い」の「10px」）。足す案。
- `NavHelpLink` は新しく足す Style（Background=Transparent・BorderThickness=0・Cursor=Hand）。
- 札の中身は 2 種類（下の 5 章）。「中断」は点を出さず、地と文字の色を替える。

## 4. 寸法と色

窓 1280×820、左の欄 220×788 のときの値。y は左の欄の上端から。

| 要素 | 大きさ・位置 | 余白 | 地の色 | 枠 | 文字の Style | 文字の色 | 角丸 | アイコン |
|---|---|---|---|---|---|---|---|---|
| NavRoot | 220×788 | – | Bg.Subtle（#F9FAFA） | – | – | – | 0 | – |
| NavList | 横 Fill、y 8〜155 | 上 8・下 8、項目の間 1 | – | – | – | – | – | – |
| 項目（SearchTab・IndexTab・SettingsTab・KillTab） | 220×36、y 8・45・82・119 | 左右 12、アイコンの枠と文字の間 8（文字は x=34） | 選んでいない: なし。選んだ: Select.Soft（#E1F2FF） | – | 選んでいない: Nav（13 Medium）。選んだ: Heading（13 SemiBold） | 選んでいない: Ink.Strong（#202124）。選んだ: Accent（#0078D4） | 0 | 14×14 の枠（空。未定） |
| ActiveBar | 3×36、項目の左端 | – | Accent（#0078D4） | – | – | – | 1 | – |
| IndexTabBadge（進み具合） | 43×15（「58%」のとき）、x=163・y=56（IndexTab の行の上下の中央） | 左右 6・上下 1、点と文字の間 4 | Select.Soft（#E1F2FF） | – | 10 SemiBold（Style 無し） | Accent（#0078D4） | 9 | 点 6×6、Accent |
| IndexTabBadge（中断） | 32×15、x=163・y=56 | 左右 6・上下 1 | Warn.Soft（#FFF5E0） | – | 10 SemiBold（Style 無し） | Warn（#BA7D00） | 9 | 点なし |
| NavSeparator | 220×1、y=163 | – | Border.Separator（#D5D9DE） | – | – | – | – | – |
| NavPaneHost | 220×587、y 164〜751 | 0 | target_tree が持つ（Bg.Pane） | – | – | – | – | – |
| 空き（spacer。検索以外） | 220×596、y 155〜751 | – | – | – | – | – | – | – |
| NavFooterLine | 220×1、y=751 | – | Border.Soft（#E0E2E5） | – | – | – | – | – |
| help-about-section | 220×36、y=752 | 左 12 | – | – | – | – | – | – |
| AboutLink | 95×22（x=12・y=7） | 上下 4、アイコンと文字の間 4 | – | – | Meta（11 Regular） | Ink.Body（#5F6368） | – | `info` 14（丸は線 1.1667、中の「i」は README の 9 のとおり線 1.5）、Ink.Body |
| NavResizeLine | 1×788、右端 | – | Border.Splitter（#D0D4D9） | – | – | – | – | – |
| NavResizeGrip | 3×28、右端（x=217）、縦の中央（y=380） | – | Border.Grip（#A9AFB6） | – | – | – | 1.5 | – |

- 検索以外の画面では、NavList の下の余白 8 が無く（nav-list の高さ 147）、すぐ下から空き（spacer）が始まる。見た目は区切りの線が無いだけで、ほかは同じ。WPF では NavList の Margin を変えず、NavSeparator を `Collapsed` にする（空きの始まりが 8 下がるが、空きなので見た目は変わらない）。
- 札の左端は、Figma ではどの variant も左の欄の左端から 163 で固定。左の欄の幅を変えたときの位置は未定（決めること）。

## 5. 状態ごとの見え方

入力は判断層 `getNavView` の `screen`（選んでいる画面）・`indexing`（インデックス作成の様子）・`hasIndex`（検索できるフォルダがあるか）・`checkedCount`（検索対象のチェックの数）。

| 状態 | 入る条件 | 選んだ項目 | IndexTabBadge | NavSeparator | NavPaneHost の中身 |
|---|---|---|---|---|---|
| 既定（340:553） | 部品の既定 | SearchTab | 出さない | 出す | 空 |
| 検索（366:41383）・P-H（366:31985） | `screen=search`・`hasIndex`・作成していない | SearchTab | 出さない | 出す | target_tree（既定。「検索対象 3 / 4」） |
| 検索-空（366:45903）・P-H0（366:31854） | `screen=search`・`hasIndex=$false` | SearchTab | 出さない | 出す | target_tree（空） |
| P-H-B（366:35949） | `screen=search`・`indexing.State=running`（58%） | SearchTab | 進み具合「58%」 | 出す | target_tree（既定） |
| P-H-P（366:36476） | `screen=search`・`indexing.State=paused` | SearchTab | 中断 | 出す | target_tree（既定） |
| P-E13（366:37000） | `screen=search`・`checkedCount=0` | SearchTab | 出さない | 出す | target_tree（P-E13） |
| インデックス管理（366:46031）・P-X0（353:29951） | `screen=index`・作成していない | IndexTab | 出さない | 出さない | 空き（spacer） |
| P-X-U（353:29980） | `screen=index`・`indexing.State=running`（66%） | IndexTab | 進み具合「66%」 | 出さない | 空き |
| P-X-P（353:30008） | `screen=index`・`indexing.State=paused` | IndexTab | 中断 | 出さない | 空き |
| 設定（366:47153）・P-C（353:30037） | `screen=settings` | SettingsTab | 出さない | 出さない | 空き |
| Office の終了（366:47393）・P-P（353:30066） | `screen=office` | KillTab | 出さない | 出さない | 空き |

- 札は画面に関係なく、インデックス作成の様子で出す（P-H-B・P-X-U、P-H-P・P-X-P）。
- 札は選んだ項目の上（Select.Soft の地）でも、地の色を替えない（P-X-U・P-X-P）。
- 今の「⚠ 1 インデックス管理」（失敗があるとき）と「⚠ 9 プロセス停止」（バックグラウンドの Office があるとき）に当たる印は Figma に無い（未定）。
- すべての項目が押せる（無効になる状態は Figma に無い）。

## 6. 文言

| 文言の ID | 全文 | 使う状態 |
|---|---|---|
| Nav.Search | `検索` | すべて |
| Nav.Index | `インデックス管理` | すべて |
| Nav.Settings | `設定` | すべて |
| Nav.Office | `Office の終了`（「Office」と「の」の間に半角の空白） | すべて |
| Nav.Badge.Progress | `{進み具合}%`（数と % の間に空白なし。例 `58%`・`66%`） | P-H-B・P-X-U |
| Nav.Badge.Paused | `中断` | P-H-P・P-X-P |
| Nav.About | `バージョン情報` | すべて |

- 項目に番号（今の「1 インデックス管理」「9 プロセス停止」）を付けない。キー操作のヒントは出さない。

## 7. 操作と振る舞い

| きっかけ | 何が起きるか | 呼ぶ関数 | 次の状態 |
|---|---|---|---|
| SearchTab をクリック | 検索の画面にする。NavPaneHost に target_tree を差す | `showScreen 'search'`（shell.ps1） | 検索・P-H 系 |
| IndexTab をクリック | インデックス管理の画面にする。`refreshIndexingState` | `showScreen 'index'` | P-X 系 |
| SettingsTab をクリック | 設定の画面にする | `showScreen 'settings'` | P-C |
| KillTab をクリック | Office の終了の画面にする。`refreshProcesses` と `processTimer` の開始 | `showScreen 'office'` | P-P |
| 選んでいる項目をもう一度クリック | 何もしない | – | そのまま |
| AboutLink をクリック | バージョン情報のダイアログを開く（今の `MoreButton` → `AboutMenuItem` の置き換え） | `showAboutDialog` | – |
| インデックス作成の進み具合が変わる | 札の数を書き換える | `applyNavView` | P-H-B・P-X-U |
| インデックス作成を中断した・再開した・終わった | 札を「中断」にする・進み具合に戻す・消す | `applyNavView` | P-H-P・P-X-P／札なし |

- ホバー・押したときの見た目は Figma に無い（未定。`figma_wpf_map.md` は使うキーに Bg.Hover を挙げている）。
- フォーカスの見た目は Figma に無い（未定）。
- Tab の順: SearchTab → IndexTab → SettingsTab → KillTab → NavPaneHost の中 → AboutLink。
- ツールチップ: 項目・札・AboutLink とも Figma に無い（出さない）。

## 8. リサイズ

| 部分 | 動き |
|---|---|
| 左の欄の幅 | 既定 220、最小 180、最大 360（shell.xaml の NavColumn と NavSplitter） |
| NavList・項目 | 横 Fill。高さ Hug（36×4＋間 1×3＋上下 8） |
| NavPaneHost／空き | 縦 Fill（高さは窓の高さ − 32 − 155 − 8 − 1 − 1 − 36） |
| footer-line・help | 下に固定（高さ 1・36） |
| 右端の線とつまみ | 線は縦いっぱい、つまみは縦の中央 |

- 項目の文字は切らない（固定の文言で、最小幅 180 に入る）。
- 窓の高さ 640 のとき、NavPaneHost の高さは 431（検索の画面）。1000 のときは 791。
- 札は左端 163 の固定で置くと、左の欄 180 のとき右端（163＋43＝206）が欄からはみ出す（決めること）。

## 9. 判断層

`scripts/tebunko/ui/shell/nav_view.ps1`。今の `applyIndexingState`（`IndexTabHeader` の文字）と `updateKillBadge`（`KillTabHeader` の文字）の、ナビに当たる部分をここに移す。

### getNavView

- 入力:
  - `[string]$screen`（`search` / `index` / `settings` / `office`）
  - `[hashtable]$indexing`（`State` = `idle` / `running` / `paused`、`Percent` = 0〜100 の整数。`running` のときだけ使う）
- 出力: ハッシュテーブル

| キー | 型 | 中身 |
|---|---|---|
| Selected | string | 選んだ見た目にする項目の x:Name（`SearchTab` / `IndexTab` / `SettingsTab` / `KillTab`） |
| ShowPane | bool | NavSeparator と NavPaneHost を出すか（検索の画面だけ `$true`） |
| Badge | hashtable | `Visible`（bool）・`Kind`（`progress` / `paused` / `$null`）・`Text`（string） |

```powershell
It "画面 <screen> では <selected> を選び、ツリーの欄は <showPane>" -TestCases @(
    @{ screen = 'search';   selected = 'SearchTab';   showPane = $true }
    @{ screen = 'index';    selected = 'IndexTab';    showPane = $false }
    @{ screen = 'settings'; selected = 'SettingsTab'; showPane = $false }
    @{ screen = 'office';   selected = 'KillTab';     showPane = $false }
) {
    param ($screen, $selected, $showPane)
    $view = getNavView $screen @{ State = 'idle' }
    $view.Selected | Should -Be $selected
    $view.ShowPane | Should -Be $showPane
}

It "作成の様子 <state>（<percent>）では札が <kind>・「<text>」" -TestCases @(
    @{ state = 'idle';    percent = $null; visible = $false; kind = $null;      text = '' }
    @{ state = 'running'; percent = 58;    visible = $true;  kind = 'progress'; text = '58%' }
    @{ state = 'running'; percent = 66;    visible = $true;  kind = 'progress'; text = '66%' }
    @{ state = 'running'; percent = 0;     visible = $true;  kind = 'progress'; text = '0%' }
    @{ state = 'running'; percent = 100;   visible = $true;  kind = 'progress'; text = '100%' }
    @{ state = 'paused';  percent = 40;    visible = $true;  kind = 'paused';   text = '中断' }
) {
    param ($state, $percent, $visible, $kind, $text)
    $badge = (getNavView 'search' @{ State = $state; Percent = $percent }).Badge
    $badge.Visible | Should -Be $visible
    $badge.Kind | Should -Be $kind
    $badge.Text | Should -Be $text
}

It "札は画面に関係なく出す" -TestCases @(
    @{ screen = 'search' }, @{ screen = 'index' }, @{ screen = 'settings' }, @{ screen = 'office' }
) {
    param ($screen)
    (getNavView $screen @{ State = 'paused' }).Badge.Text | Should -Be '中断'
}
```

- 進み具合の数（`Percent`）をどこから取るか（今の `IndexingProgressText` の元の値）は、`index/index_detail.md` の進み具合と同じ値を使う。小数の丸め方は未定（表は整数で渡す前提）。

## 10. 画面層

`scripts/tebunko/ui/shell/nav.ps1`

| x:Name | 配線 |
|---|---|
| SearchTab・IndexTab・SettingsTab・KillTab | `Add_Checked` → `showScreen <画面>`（shell.ps1）。`showScreen` の中から IsChecked を書くときに Checked が再び来ても、同じ画面なら何もしない |
| AboutLink | `Add_Click` → `showAboutDialog` |
| IndexTabBadge・IndexTabBadgeDot・IndexTabBadgeText | `applyNavView` が書く |
| NavSeparator・NavPaneHost | `applyNavView` が `ShowPane` で Visibility を書く（中身を差すのは `showScreen`） |

`applyNavView` の手順:

1. `$view = getNavView $script:currentScreen (getIndexingSnapshot)`
2. `$ui.<$view.Selected>.IsChecked = $true`
3. `$ui.NavSeparator.Visibility` を `ShowPane` で `Visible` / `Collapsed`
4. 札: `Visible` が `$false` なら `Collapsed`。`progress` なら地 `Select.Soft`・点を出す・文字 `Accent`。`paused` なら地 `Warn.Soft`・点を `Collapsed`・文字 `Warn`。文字は `Text`。色は `$window.FindResource('<キー>')` で取る。

- `applyNavView` は UI スレッドで呼ぶ。インデックス作成の進み具合は別のスレッドで動くので、今の `applyIndexingState` と同じく、タイマー（UI スレッド）で読んだ値から呼ぶ。別のスレッドから呼ぶときは `$window.Dispatcher.Invoke`。
- 消すもの: `Tabs`・`IndexTabHeader`・`KillTabHeader`・`MoreButton`・`AboutMenuItem` と、それを書いている所（`gui.ps1` の名前の一覧、`index_tab.ps1` の 908 行目、`process_tab.ps1` の 89 行目、`about_dialog.ps1` の MoreButton の配線）。

## 11. 受け入れ

見比べる画像: `../png/148_67.png`（P-H0）・`../png/108_230.png`（H）・`../png/108_1116.png`（H-B）・`../png/108_1419.png`（H-P）・`../png/175_4085.png`（E13）。窓 1280×820。インデックス管理・設定・Office の終了は Figma の nav の variant（P-X0・P-X-U・P-X-P・P-C・P-P）のスクリーンショットと見比べる。

- 左の欄の地が #F9FAFA、幅 220。
- 項目が上から 8 の所から 36 ずつ、間 1 で 4 つ並び、文字の左端が x=34。
- 選んだ項目の地が #E1F2FF、左端に 3px の青（#0078D4）の帯、文字が青の SemiBold。ほかは #202124 の Medium。
- 検索の画面だけ、y=163 に 1px の線（#D5D9DE）があり、その下にツリーの欄。ほかの画面では線が無い。
- 下から 37 の所に 1px の線（#E0E2E5）、その下の 36 の帯に、ⓘ のアイコンと「バージョン情報」（11、#5F6368）が左 12 から。
- 右端に 1px の線（#D0D4D9）と、縦の中央に 3×28 のつまみ（#A9AFB6）。
- インデックスを作っている間、「インデックス管理」の行の x=163 に、淡い青の札（角丸 9・高さ 15）で青い点と「58%」。中断している間は淡い橙の札（#FFF5E0）に「中断」（#BA7D00）で、点が無い。
- 画面を切り替えても、左の欄の幅とツリーの開閉・チェックが変わらない。

## 決めること

- 項目のアイコン: Figma は 14×14 の枠だけで中身が空。アイコンを入れるか（入れるなら Lucide の名前）、枠ごと消して文字を x=34 のまま保つか。
- 項目のホバー・押したとき・フォーカスの見た目（候補は Bg.Hover #F3F3F4 だが Figma に variant が無い）。
- 札の位置: Figma は左端から 163 の固定。左の欄の幅を変えたとき（180 だとはみ出す）に、右端から 14 に寄せるか、文字の直後に置くか。
- 失敗があるとき（今の「⚠ 1 インデックス管理」）と、バックグラウンドの Office があるとき（今の「⚠ 9 プロセス停止」）の印を出すか、出すならどんな札か。
- 進み具合の数の丸め方（切り捨て・四捨五入）と、100% に達したあと札を消すまでの扱い。
- 新しい Style `NavItem`・`NavHelpLink`・`Badge.Micro`（10 SemiBold）を theme に足すこと（`figma_wpf_map.md` の「食い違い」に足す案）。
