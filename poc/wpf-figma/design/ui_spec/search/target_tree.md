# 検索対象のツリー（target_tree）

## 1. 役割

検索の画面のとき、左の欄（nav）の中に出す、どのフォルダを検索するかを選ぶツリー。一番上の項目がインデックス 1 件で、開くとその下のフォルダが並ぶ。チェックの付いたフォルダだけを検索する。見出しに選んだ数を出し、「すべて」「解除」でまとめて付け外しし、「フォルダを探す」で名前から絞り込む。インデックスの状態（要更新・更新中・エラー）は名前の直後の点で示す。

- Figma の部品: `xaml/search/target_tree` 366:313（既定 340:564）。variant: 既定 340:564・空 366:215・P-E13 366:312。
- 使う画面: 検索の画面のすべて（nav の variant 検索・検索-空・P-H 系・P-E13 の NavPaneHost の中）。見本は `../png/108_230.png`（H）・`../png/148_67.png`（H0。空）・`../png/175_4085.png`（E13）・`../png/176_2571.png`（フルパスのツールチップ）。
- 今の `tab_search.xaml` の左のカード（見出し「検索対象」・「すべて選択」「すべて解除」・`IndexTree`・`IndexTreePlaceholder`）と `ui/index_tree.ps1` の置き換え。

## 2. 置き場所

- XAML: `scripts/tebunko/xaml/search/target_tree.xaml`（ルートは `Grid`。UserControl 相当）。
- 画面層: `scripts/tebunko/ui/search/target_tree.ps1`（今の `ui/index_tree.ps1` を移す）。判断層: `scripts/tebunko/ui/search/target_tree_view.ps1`。
- 親: `nav.xaml` の `NavPaneHost`（`shell.md` の `showScreen 'search'` が差す。検索以外の画面では外す）。
- 読み込み: 起動のときに 1 回だけ `loadXaml` して取っておき、検索の画面を選ぶたびに同じものを差す（チェック・開閉・絞り込みの文字・スクロールの位置を保つため）。x:Name は `$ui` に足す。

## 3. 部品の木

```xml
<Grid x:Name="TargetTreeRoot"
      xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
      Background="{StaticResource Bg.Pane}" Margin="0" ClipToBounds="True">
  <Grid Margin="10,12,10,12">
    <Grid.RowDefinitions>
      <RowDefinition Height="18"/>    <!-- 0: pane-header -->
      <RowDefinition Height="Auto"/>  <!-- 1: IndexTreeFilterBox（上 8） -->
      <RowDefinition Height="*"/>     <!-- 2: IndexTree（上 8）／IndexTreeEmpty（上 8。空のとき） -->
      <RowDefinition Height="Auto"/>  <!-- 3: IndexTreePlaceholder（上 8。何も選んでいないとき） -->
    </Grid.RowDefinitions>

    <!-- pane-header: 両端寄せ -->
    <Grid Grid.Row="0" VerticalAlignment="Center">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <TextBlock x:Name="SearchTargetText" Grid.Column="0" Style="{StaticResource Label}"
                 Foreground="{StaticResource Ink.Strong}" VerticalAlignment="Center"
                 TextTrimming="CharacterEllipsis" Text="{Tree.Header}"/>
      <StackPanel x:Name="TreeLinks" Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
        <Button x:Name="CheckAllIndexButton" Style="{StaticResource TextLink}" Content="{Tree.CheckAll}"/>
        <Button x:Name="UncheckAllIndexButton" Style="{StaticResource TextLink}" Margin="10,0,0,0" Content="{Tree.UncheckAll}"/>
      </StackPanel>
    </Grid>

    <!-- IndexTreeFilterBox: 高さ 28、左 8、アイコンと文字の間 6 -->
    <Border x:Name="IndexTreeFilterFrame" Grid.Row="1" Margin="0,8,0,0" Height="28"
            Background="{StaticResource Bg.Surface}" BorderBrush="{StaticResource Border.Strong}"
            BorderThickness="1" CornerRadius="4" Padding="8,0,8,0">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="14"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <!-- Lucide search 14、線 1.1667、Ink.Subtle -->
        <Path Grid.Column="0" Width="14" Height="14" Stretch="Uniform" VerticalAlignment="Center"
              Stroke="{StaticResource Ink.Subtle}" StrokeThickness="1.1667"
              StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round"
              Data="{lucide:search}"/>
        <TextBox x:Name="IndexTreeFilterBox" Grid.Column="1" Margin="6,0,0,0"
                 Style="{StaticResource TextBox.Bare}" FontSize="12"
                 Foreground="{StaticResource Ink.Value}" VerticalAlignment="Center"/>
        <TextBlock x:Name="IndexTreeFilterPlaceholder" Grid.Column="1" Margin="6,0,0,0"
                   Style="{StaticResource Cell}" Foreground="{StaticResource Ink.Placeholder}"
                   VerticalAlignment="Center" IsHitTestVisible="False" Text="{Tree.FilterPlaceholder}"/>
      </Grid>
    </Border>

    <TreeView x:Name="IndexTree" Grid.Row="2" Margin="0,8,0,0"
              Background="Transparent" BorderThickness="0" Padding="0"
              ScrollViewer.HorizontalScrollBarVisibility="Disabled"
              ScrollViewer.VerticalScrollBarVisibility="Auto"
              VirtualizingStackPanel.IsVirtualizing="True">
      <TreeView.ItemContainerStyle>
        <Style TargetType="TreeViewItem">
          <Setter Property="IsExpanded" Value="{Binding IsExpanded, Mode=TwoWay}"/>
          <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
          <Setter Property="Template">
            <Setter.Value>
              <ControlTemplate TargetType="TreeViewItem">
                <StackPanel>
                  <!-- 行: 高さ 24、左 2（根）、右 4。子の段は ItemsPresenter の左 16 で下げる（子の行は左 18） -->
                  <Border x:Name="Row" Height="24" Padding="2,0,4,0" Background="Transparent">
                    <ContentPresenter ContentSource="Header" VerticalAlignment="Center"/>
                  </Border>
                  <ItemsPresenter x:Name="ItemsHost" Margin="16,0,0,0"/>
                </StackPanel>
                <ControlTemplate.Triggers>
                  <Trigger Property="IsExpanded" Value="False">
                    <Setter TargetName="ItemsHost" Property="Visibility" Value="Collapsed"/>
                  </Trigger>
                  <!-- ホバーの地の色は未定（176_2571 では灰色・枠なし） -->
                </ControlTemplate.Triggers>
              </ControlTemplate>
            </Setter.Value>
          </Setter>
        </Style>
      </TreeView.ItemContainerStyle>
      <TreeView.ItemTemplate>
        <HierarchicalDataTemplate ItemsSource="{Binding Children}">
          <!-- 開閉の印 12・チェック 14・フォルダ 14・名前・状態の点 7。間はすべて 4 -->
          <Grid ToolTip="{Binding ToolTip}">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="12"/>   <!-- 開閉の印 -->
              <ColumnDefinition Width="Auto"/> <!-- チェック -->
              <ColumnDefinition Width="Auto"/> <!-- フォルダ -->
              <ColumnDefinition Width="*"/>    <!-- 名前と点 -->
            </Grid.ColumnDefinitions>
            <ToggleButton x:Name="Chevron" Grid.Column="0" Width="12" Height="12"
                          Style="{StaticResource TreeChevron}" Focusable="False"
                          IsChecked="{Binding IsExpanded, RelativeSource={RelativeSource AncestorType=TreeViewItem}}"
                          Visibility="{Binding HasItems, RelativeSource={RelativeSource AncestorType=TreeViewItem}, Converter={StaticResource BoolToVisibility}}"/>
            <CheckBox x:Name="Check" Grid.Column="1" Margin="4,0,0,0" Width="14" Height="14"
                      Style="{StaticResource TreeCheck}" IsThreeState="True" Focusable="False"
                      IsChecked="{Binding IsChecked, Mode=OneWay}" VerticalAlignment="Center"/>
            <!-- Lucide folder 14、線 1.1667、File.Folder -->
            <Path x:Name="FolderIcon" Grid.Column="2" Margin="4,0,0,0" Width="14" Height="14" Stretch="Uniform"
                  Stroke="{StaticResource File.Folder}" StrokeThickness="1.1667"
                  StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round"
                  Data="{lucide:folder}" VerticalAlignment="Center"/>
            <!-- 名前は Hug、点は名前の直後。入りきらないときは名前だけを「…」で縮め、点は残す -->
            <Grid Grid.Column="3" Margin="4,0,0,0" HorizontalAlignment="Left">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBlock x:Name="Label" Grid.Column="0" Text="{Binding Name}"
                         Style="{StaticResource Cell}" Foreground="{StaticResource Ink.Strong}"
                         TextTrimming="CharacterEllipsis" TextWrapping="NoWrap" VerticalAlignment="Center"/>
              <Ellipse x:Name="StatusDot" Grid.Column="1" Width="7" Height="7" Margin="4,0,0,0"
                       VerticalAlignment="Center" Visibility="Collapsed"/>
            </Grid>
          </Grid>
          <HierarchicalDataTemplate.Triggers>
            <DataTrigger Binding="{Binding IsRoot}" Value="True">
              <Setter TargetName="Label" Property="Style" Value="{StaticResource Cell.Key}"/>
            </DataTrigger>
            <DataTrigger Binding="{Binding Status}" Value="stale">
              <Setter TargetName="StatusDot" Property="Fill" Value="{StaticResource Warn.Dot}"/>
              <Setter TargetName="StatusDot" Property="Visibility" Value="Visible"/>
            </DataTrigger>
            <DataTrigger Binding="{Binding Status}" Value="updating">
              <Setter TargetName="StatusDot" Property="Fill" Value="{StaticResource Accent}"/>
              <Setter TargetName="StatusDot" Property="Visibility" Value="Visible"/>
            </DataTrigger>
            <DataTrigger Binding="{Binding Status}" Value="error">
              <Setter TargetName="StatusDot" Property="Fill" Value="{StaticResource Danger.Dot}"/>
              <Setter TargetName="StatusDot" Property="Visibility" Value="Visible"/>
            </DataTrigger>
            <!-- IsFiles（直下のファイル）・IsPlaceholder（読み込み前）・Exists=False の見た目は未定 -->
          </HierarchicalDataTemplate.Triggers>
        </HierarchicalDataTemplate>
      </TreeView.ItemTemplate>
    </TreeView>

    <!-- pane-empty（空）: 高さ 120、中身を上下左右の中央、間 8 -->
    <StackPanel x:Name="IndexTreeEmpty" Grid.Row="2" Margin="0,8,0,0" Height="120"
                VerticalAlignment="Top" HorizontalAlignment="Stretch" Visibility="Collapsed">
      <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center" Margin="0,40,0,0">
        <!-- Lucide folder-closed 22.4、線 1.8667、File.Folder -->
        <Path Width="22.4" Height="22.4" Stretch="Uniform" HorizontalAlignment="Center"
              Stroke="{StaticResource File.Folder}" StrokeThickness="1.8667"
              StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round"
              Data="{lucide:folder-closed}"/>
        <TextBlock Margin="0,8,0,0" HorizontalAlignment="Center" Style="{StaticResource Cell}"
                   Foreground="{StaticResource Ink.Body}" Text="{Tree.Empty}"/>
      </StackPanel>
    </StackPanel>

    <!-- 何も選んでいないときの案内（P-E13）: 欄のいちばん下 -->
    <StackPanel x:Name="IndexTreePlaceholder" Grid.Row="3" Margin="0,8,0,0" Orientation="Horizontal"
                HorizontalAlignment="Left" Visibility="Collapsed">
      <!-- info 14（丸は線 1.1667、中の「i」は README の 9 のとおり線 1.5）、Ink.Body -->
      <Path Width="14" Height="14" Stretch="Uniform" VerticalAlignment="Center"
            Stroke="{StaticResource Ink.Body}" StrokeThickness="1.1667"
            StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round"
            Data="{lucide:info}"/>
      <TextBlock Margin="6,0,0,0" Style="{StaticResource Meta}" Foreground="{StaticResource Ink.Body}"
                 VerticalAlignment="Center" Text="{Tree.NoneChecked}"/>
    </StackPanel>
  </Grid>
</Grid>
```

- `pane-empty` の中身の縦の位置: Figma は高さ 120 の枠の中で上下の中央（アイコン 22.4＋間 8＋文字 15 ≒ 45 なので、上に約 37）。上の `Margin="0,40,0,0"` は仮。`StackPanel` の `Height="120"` の中で `VerticalAlignment="Center"` が効くよう、外側を `Grid`（Height 120）にしてもよい。
- `IndexTreeEmpty` は `figma_wpf_map.md` の x:Name の一覧に無い。Figma の `pane-empty` に当たる名前として足す案（決めること）。Figma の `IndexTreePlaceholder` は P-E13 の下の案内の名前で、今の `IndexTreePlaceholder`（インデックスが無いときの文）とは役目が違う。
- `TreeChevron`（Lucide `chevron-right` を閉、`chevron-down` を開で出す 12×12 の ToggleButton。中の印 10×10・線 0.8333・Ink.Body）、`TreeCheck`（チェック 14×14）、`TextLink`（文字だけのボタン）、`TextBox.Bare`（枠なしの TextBox）は Style。`TreeChevron` は今の theme にある。`TreeCheck`・`TextLink`・`TextBox.Bare` は足す案。`BoolToVisibility` は `BooleanToVisibilityConverter`。
- 名前の行の `Grid HorizontalAlignment="Left"`（列 `*`＋`Auto`）は、空きがあるときは名前の幅だけ取り（点が名前の直後に付く）、足りないときは `*` の列が縮んで名前が「…」になり、点は残る。

`TreeCheck` の見た目（Figma の値）:

| チェック | 地 | 枠 | 角丸 | 印 |
|---|---|---|---|---|
| 付いている（true） | Accent（#0078D4） | Accent（#0078D4） | 2.5 | 白（Ink.OnAccent）のチェック。Lucide `check` 相当、線 1.8 |
| 付いていない（false） | Bg.Surface（#FFFFFF） | 1.2px Ink.Subtle（#80868B） | 3 | なし |
| 一部（null） | 未定 | 未定 | 未定 | 未定 |

## 4. 寸法と色

左の欄 220 のとき（欄の中身は 219 幅に描かれている。右の 1px は nav の線）。

| 要素 | 大きさ | 余白 | 地の色 | 枠 | 文字の Style | 文字の色 | 角丸 | アイコン |
|---|---|---|---|---|---|---|---|---|
| TargetTreeRoot | 横 Fill × 縦 Fill（220×587） | 左右 10・上下 12、子の間 8 | Bg.Pane（#F0F2F4） | – | – | – | – | – |
| pane-header | 横 Fill × 18 | – | – | – | – | – | – | – |
| SearchTargetText | Hug | – | – | – | Label（12 SemiBold） | Ink.Strong（#202124） | – | – |
| TreeLinks | Hug | 間 10 | – | – | Meta（11 Regular） | Accent（#0078D4） | – | – |
| IndexTreeFilterFrame | 横 Fill × 28 | 左 8、アイコンと文字の間 6 | Bg.Surface（#FFFFFF） | 1px Border.Strong（#C9CED4） | – | – | 4 | Lucide `search` 14、線 1.1667、Ink.Subtle（#80868B） |
| IndexTreeFilterPlaceholder | Hug | – | – | – | Cell（12 Regular） | Ink.Placeholder（#9AA0A6） | – | – |
| IndexTreeFilterBox（打った文字） | Fill | – | – | – | 12 Regular | Ink.Value（#212126）（Figma に打った状態が無いので未定） | – | – |
| IndexTree | 横 Fill × 縦 Fill | 0 | 透明（Bg.Pane が見える） | – | – | – | – | – |
| 行（Row） | 横 Fill × 24 | 根: 左 2・右 4。子: 左 18・右 4。部品の間 4 | 透明 | – | – | – | – | – |
| Chevron | 12×12（中の印 10×10） | – | – | – | – | – | – | Lucide `chevron-down`（開）・`chevron-right`（閉）、線 0.8333、Ink.Body（#5F6368）。子の無い行は枠だけ（空） |
| Check | 14×14 | – | 上の表 | 上の表 | – | – | 上の表 | – |
| FolderIcon | 14×14 | – | – | – | – | – | – | Lucide `folder`、線 1.1667、File.Folder（#E8A020） |
| Label（根） | Hug（入りきらないと縮む） | – | – | – | Cell.Key（12 Medium） | Ink.Strong（#202124） | – | – |
| Label（子） | Hug（入りきらないと縮む） | – | – | – | Cell（12 Regular） | Ink.Strong（#202124） | – | – |
| StatusDot | 7×7、名前の直後（間 4） | – | 要更新: Warn.Dot（#E8A400）。更新中: Accent（#0078D4）。エラー: Danger.Dot（#D93025） | – | – | – | 円 | – |
| IndexTreeEmpty（pane-empty） | 横 Fill × 120 | 中身の間 8、上下左右の中央 | – | – | Cell（12 Regular） | Ink.Body（#5F6368） | – | Lucide `folder-closed` 22.4、線 1.8667、File.Folder（#E8A020） |
| IndexTreePlaceholder | Hug | アイコンと文字の間 6 | – | – | Meta（11 Regular） | Ink.Body（#5F6368） | – | `info` 14（丸は線 1.1667、中の「i」は README の 9 のとおり線 1.5）、Ink.Body |

- 点は根（インデックス）の行だけに出る（Figma の見本では子の行に点が無い）。
- 子の行の名前の段下げは、Figma では 2 段（根と子）だけ。3 段目より下の段下げ（仮に 16 ずつ）は未定。
- `figma_wpf_map.md` は使うキーに `Link（Accent）`・`Ok`・`Warn`・`Danger.Text` を挙げているが、Figma の部品の値は「すべて／解除」が 11 Regular（Link の Style は 11 SemiBold）、点が Warn.Dot・Accent・Danger.Dot。ここでは Figma の部品の値に合わせた。
- `diff_round3.md`・`diff_round4.md` に「『検索対象 3 / 4』の数字は太字」とあるが、Figma の部品は見出し全体が 1 つの 12 SemiBold の文字。部品の値に合わせた。
- 空の状態のアイコンは、Figma のレイヤーの名前が `folder-closed`（任された文書では「folder-open」）。Figma のとおり `folder-closed` にした。

## 5. 状態ごとの見え方

入力は判断層 `getTargetTreeView` の `total`（インデックスの数＝根の数）と `checked`（チェックの付いた根の数）。

| 状態 | 入る条件 | SearchTargetText | TreeLinks | IndexTreeFilterFrame | IndexTree | IndexTreeEmpty | IndexTreePlaceholder |
|---|---|---|---|---|---|---|---|
| 既定（340:564）・H・P-H・P-H-B・P-H-P | `total ≥ 1`・`checked ≥ 1` | `検索対象 3 / 4` | 出す | 出す | 出す | 出さない | 出さない |
| 空（366:215）・H0・P-H0 | `total = 0` | `検索対象` | 出さない | 出さない | 出さない | 出す | 出さない |
| P-E13（366:312）・E13 | `total ≥ 1`・`checked = 0` | `検索対象 0 / 4` | 出す | 出す | 出す（すべてチェックなし） | 出さない | 出す |

行ごとの見え方（IndexNode の値）:

| 行の状態 | 条件 | 見え方 |
|---|---|---|
| 根 | `IsRoot = $true` | 名前が Cell.Key（12 Medium）、左 2 |
| 子 | `IsRoot = $false` | 名前が Cell（12 Regular）、左 18 |
| 開いている | `IsExpanded = $true` | 印が `chevron-down`、子の行を出す |
| 閉じている | `IsExpanded = $false`・子がある | 印が `chevron-right`（例 アーカイブ） |
| 子が無い | 子が無い | 印の枠（12）だけで、何も描かない |
| チェックあり／なし／一部 | `IsChecked = $true / $false / $null` | 3 章の `TreeCheck` の表（一部は未定） |
| 要更新 | `Status = 'stale'` | 名前の直後に橙の点（例 顧客） |
| 更新中 | `Status = 'updating'` | 名前の直後に青の点（例 営業部2025） |
| エラー | `Status = 'error'` | 名前の直後に赤の点（例 アーカイブ） |
| 状態なし | `Status = ''` | 点を出さない（例 営業部） |
| ホバー | カーソルが行の上 | 行の地が灰色・枠なし（`../png/176_2571.png`）。色の値は未定 |
| 名前が長い | 名前が欄に入りきらない | 名前の末尾を「…」で切る（例 `提案書_2024年度下…`）。点は切らずに残す |

- 点はチェックが外れていても出す（P-E13）。
- 「（このフォルダ直下のファイル）」の行（`IsFiles`）・開く前の仮の行（`IsPlaceholder`）・元のフォルダが見つからない根（`Exists = $false`）の見え方は Figma に無い（未定）。
- 選んだ行（TreeView の選択）の見た目は Figma に無い（未定）。

## 6. 文言

| 文言の ID | 全文 | 使う状態 |
|---|---|---|
| Tree.Header | `検索対象 {チェックの数} / {インデックスの数}`（「検索対象」の後・「/」の前後は半角の空白。例 `検索対象 3 / 4`・`検索対象 0 / 4`） | 既定・P-E13 |
| Tree.HeaderEmpty | `検索対象` | 空 |
| Tree.CheckAll | `すべて` | 既定・P-E13 |
| Tree.UncheckAll | `解除` | 既定・P-E13 |
| Tree.FilterPlaceholder | `フォルダを探す` | 既定・P-E13（打った文字が空のとき） |
| Tree.Empty | `検索できるフォルダがありません` | 空 |
| Tree.NoneChecked | `検索するフォルダを選んでください` | P-E13 |
| Tree.ToolTip | `{元のフォルダのフルパス}`（例 `C:\共有\営業部\2024\見積もり\提案書_2024年度下期_大口顧客向け`。拡張子などを足さない） | 行にカーソルを当てたとき |

- ツールチップは、フルパスだけを出す（今の「元のフォルダ：…」「インデックス：…」の 2 行はやめる）。見本 `../png/176_2571.png` の文は `C:\Share\営業部\2024\見積もり\提案書_2024年度下期_大口顧客向け`。ツールチップの見た目（地・文字・余白）は `search/overlays.md` に従う。
- 元のフォルダが見つからないときに足す文（今の「（フォルダが見つかりません。検索時はスキップします）」）は Figma に無い（未定）。
- キー操作のヒント（今のコードの Space で付け外し・F5 で読み直し）は出さない。

## 7. 操作と振る舞い

| きっかけ | 何が起きるか | 呼ぶ関数 | 次の状態 |
|---|---|---|---|
| チェックをクリック | その行のチェックを切り替え、子と親に伝える（一部の印は親に）。外したフォルダを設定に保存し、見出しの数・検索ボタンを直す | `$node.Toggle()` → `onIndexTreeChecked` | 既定／P-E13 |
| 開閉の印をクリック | 開く・閉じる。初めて開くときは子のフォルダを読む | `TreeViewItem.Expanded` → `$node.LoadChildren()` | – |
| 「すべて」をクリック | すべての根にチェックを付け、保存する | `setAllIndexChecked $true` | 既定 |
| 「解除」をクリック | すべての根のチェックを外し、保存する | `setAllIndexChecked $false` | P-E13 |
| 「フォルダを探す」に文字を打つ | 名前に文字を含む行と、その上の段だけを出す（文字を消すと元に戻す） | `applyTreeFilter`（`testTreeFilterMatch`） | – |
| 行にカーソルを当てる | 行の地を灰色にし、元のフォルダのフルパスをツールチップで出す | – | – |
| 行の名前をクリック | 未定（今は TreeView の選択だけで、何も起きない） | – | – |
| インデックスが増えた・減った・状態が変わった | ツリーを読み直す（チェックと開閉は保つ） | `loadIndexTree` | 既定／空 |

- 「すべて」「解除」のホバー（下線を引くか）・押したとき・どちらも効かないとき（すでにすべて付いている）の見た目は Figma に無い（未定）。
- 絞り込みの欄のフォーカスの見た目（枠の色）は Figma に無い（未定）。
- Tab の順: CheckAllIndexButton → UncheckAllIndexButton → IndexTreeFilterBox → IndexTree。
- 絞り込みで、まだ開いていない（子を読んでいない）段のフォルダを探すかは未定（読んでいない段まで読むと遅くなる）。
- 絞り込みのあいだに「すべて」「解除」を押したとき、隠れている行にも効かせるかは未定。

## 8. リサイズ

一覧が多いとき（スクロールする所・固定する所・仮想化・未定のこと）は `../scroll.md`（見本 T-多・T-多-1024）。

| 部分 | 動き |
|---|---|
| 欄 | 左の欄の幅（180〜360）と NavPaneHost の高さに合わせて、横・縦 Fill |
| pane-header | 横 Fill。見出しは左、「すべて」「解除」は右。入りきらないときは見出しを「…」で切る |
| 絞り込みの欄 | 横 Fill、高さ 28 固定 |
| ツリー | 横・縦 Fill。行が多いと縦にスクロール（`Auto`）。横にはスクロールしない（`Disabled`） |
| 名前 | 1 行。入りきらないときは末尾を「…」（`TextTrimming="CharacterEllipsis"`）。点は名前の直後に付けたまま残す |
| 開閉の印 | 12 固定 |
| 空の状態 | 高さ 120 の枠の中で中央 |
| 下の案内（P-E13） | 欄のいちばん下に Hug。1 行（左の欄が 180 のときに入りきるかは未定） |

- 左の欄 180 のとき: 名前の幅は、根で 180 − 1（右の線）− 20（左右 10）− 2 − 4 − 12 − 4 − 14 − 4 − 14 − 4 ≒ 101。
- 1024×640 のとき: 欄の高さ 431（`nav.md`）。ツリーの高さは 431 − 24（上下 12）− 18 − 8 − 28 − 8 ＝ 345（行 14 本分）。
- 1600×1000 のとき: 欄の高さ 791、ツリーの高さ 705。

## 9. 判断層

`scripts/tebunko/ui/search/target_tree_view.ps1`。今の `index_tree.ps1` のうち、`$ui`・`$script:indexRoots` に触らない部分（`describeSearchTargets`）もここに移す。`getSearchTargets`・`isAllIndexChecked`・`saveSearchExcludes`・`onIndexTreeChecked`・`setAllIndexChecked`・`loadIndexTree` は画面層に残す（名前は今のまま）。

### getTargetTreeView

- 入力: `[int]$total`（根の数）・`[int]$checked`（チェックの付いた根の数）
- 出力: ハッシュテーブル `Header`（string）・`ShowTools`（bool。リンク・絞り込み・ツリー）・`ShowEmpty`（bool）・`ShowHint`（bool。下の案内）

```powershell
It "根 <total> 件・チェック <checked> 件 → 「<header>」" -TestCases @(
    @{ total = 4; checked = 3; header = '検索対象 3 / 4'; tools = $true;  empty = $false; hint = $false }
    @{ total = 4; checked = 4; header = '検索対象 4 / 4'; tools = $true;  empty = $false; hint = $false }
    @{ total = 4; checked = 0; header = '検索対象 0 / 4'; tools = $true;  empty = $false; hint = $true }
    @{ total = 1; checked = 1; header = '検索対象 1 / 1'; tools = $true;  empty = $false; hint = $false }
    @{ total = 0; checked = 0; header = '検索対象';       tools = $false; empty = $true;  hint = $false }
) {
    param ($total, $checked, $header, $tools, $empty, $hint)
    $view = getTargetTreeView $total $checked
    $view.Header | Should -Be $header
    $view.ShowTools | Should -Be $tools
    $view.ShowEmpty | Should -Be $empty
    $view.ShowHint | Should -Be $hint
}
```

- 一部だけチェックの付いた根（`IsChecked = $null`）を `checked` に数えるかは未定。画面層で数えるときの決まりが決まったら、表に行を足す。

### getTreeStatus

インデックスの状態から、点の種類を返す。IndexNode の `Status` に入れる値。

- 入力: `[string]$state`（インデックスの状態。今の状態の名前との対応は下の注）
- 出力: string（`stale` / `updating` / `error` / `''`）

```powershell
It "インデックスの状態 <state> → 点 <status>" -TestCases @(
    @{ state = '要更新'; status = 'stale' }
    @{ state = '更新中'; status = 'updating' }
    @{ state = 'エラー'; status = 'error' }
    @{ state = '最新';   status = '' }
) {
    param ($state, $status)
    getTreeStatus $state | Should -Be $status
}
```

- 入力の値は Figma のレイヤーの名前（`status-要更新`・`status-更新中`・`status-エラー`）。今のコードの状態（`docs/design/gui/state-flow.md` の未作成・中断中・失敗あり・作成中・通常、インデックス管理の「最新」「要更新」「反映中」のバッジ）との対応は、`index/index_list.md` の状態の名前と合わせて決める（決めること）。

### getTreeItemToolTip

- 入力: `[string]$sourcePath`（根の元のフォルダ）・`[string]$relPath`（根からの相対パス。根は空）
- 出力: string（フルパス）

```powershell
It "<sourcePath> と <relPath> → <expected>" -TestCases @(
    @{ sourcePath = 'C:\共有\営業部';  relPath = '';                expected = 'C:\共有\営業部' }
    @{ sourcePath = 'C:\共有\営業部';  relPath = 'A社';             expected = 'C:\共有\営業部\A社' }
    @{ sourcePath = 'C:\共有\営業部\'; relPath = '2024\見積もり';   expected = 'C:\共有\営業部\2024\見積もり' }
    @{ sourcePath = '\\server\共有';   relPath = '顧客\取引先台帳'; expected = '\\server\共有\顧客\取引先台帳' }
) {
    param ($sourcePath, $relPath, $expected)
    getTreeItemToolTip $sourcePath $relPath | Should -Be $expected
}
```

- 元のフォルダが分からない根（今の `SourcePath` が空）のツールチップは未定。

### testTreeFilterMatch

- 入力: `[string]$name`（行の名前）・`[string]$filter`（打った文字）
- 出力: bool（その行を出すか。上の段を出すかは画面層が子の結果から決める）

```powershell
It "名前 <name> と絞り込み <filter> → <expected>" -TestCases @(
    @{ name = '営業部';     filter = '';     expected = $true }
    @{ name = '営業部';     filter = '営業'; expected = $true }
    @{ name = '営業部2025'; filter = '2025'; expected = $true }
    @{ name = '顧客';       filter = '営業'; expected = $false }
    @{ name = '提案書_2024年度下期_大口顧客向け'; filter = '大口'; expected = $true }
) {
    param ($name, $filter, $expected)
    testTreeFilterMatch $name $filter | Should -Be $expected
}
```

- 英字の大文字・小文字、全角・半角を区別するか、前後の空白を除くかは未定（決まったら表に行を足す）。

### describeSearchTargets（今のまま移す）

今の `index_tree.ps1` の `describeSearchTargets`（先頭の 3 件と「ほか N か所」、直下だけのときの「（直下のファイル）」）。Figma の見出しは数だけなので、この文をどこに出すか（検索の画面の別の所か、出さないか）は未定。テストは今の `tests/` のものを移す。

## 10. 画面層

`scripts/tebunko/ui/search/target_tree.ps1`（今の `ui/index_tree.ps1` を移し、`tebunko/lib.ps1` ではなく `gui.ps1` から読む）。

| x:Name | 配線 |
|---|---|
| IndexTree | `ItemsSource = $script:indexRoots`。`AddHandler([ButtonBase]::ClickEvent)` で、元が `CheckBox` のときに `DataContext.Toggle()` → `onIndexTreeChecked`（今のまま）。`AddHandler([TreeViewItem]::ExpandedEvent)` で `LoadChildren()`（今のまま） |
| CheckAllIndexButton | `Add_Click` → `setAllIndexChecked $true` |
| UncheckAllIndexButton | `Add_Click` → `setAllIndexChecked $false` |
| IndexTreeFilterBox | `Add_TextChanged` → 短い間（未定）をおいて `applyTreeFilter`。`IndexTreeFilterPlaceholder.Visibility` を、文字が空なら `Visible` |
| SearchTargetText・TreeLinks・IndexTreeFilterFrame・IndexTree・IndexTreeEmpty・IndexTreePlaceholder | `applyTargetTreeView` が書く |

`applyTargetTreeView` の手順（`loadIndexTree` の終わりと `onIndexTreeChecked` から呼ぶ。今の `IndexTreePlaceholder` の切り替えを置き換える）:

1. `$total = $script:indexRoots.Count`、`$checked` = チェックの付いた根の数（一部の扱いは未定）
2. `$view = getTargetTreeView $total $checked`
3. `$ui.SearchTargetText.Text = $view.Header`
4. `ShowTools` で `TreeLinks`・`IndexTreeFilterFrame`・`IndexTree` の Visibility、`ShowEmpty` で `IndexTreeEmpty`、`ShowHint` で `IndexTreePlaceholder`
5. 検索ボタンの可否は今の `updateSearchTarget` → `updateSearchButton` のまま（`search/search_bar.md`）

IndexNode（`types.ps1`）に足すもの:

| プロパティ | 型 | 中身 |
|---|---|---|
| IsRoot | bool | 根なら `$true`（`Parent` が無い） |
| Status | string | `getTreeStatus` の値（根だけ。子は `''`） |

- `ToolTip` は `getTreeItemToolTip $SourcePath $RelPath` の値にする（今の 2 行の文はやめる）。
- インデックスの状態（`Status`）は、`loadIndexTree` のときと、インデックスの作成の様子が変わったとき（今の `applyIndexingState` のタイマー）に書き換える。別のスレッドで調べた状態は、UI スレッド（タイマーか `$window.Dispatcher.Invoke`）で書く。
- 子のフォルダを読む（`LoadChildren`）のは今と同じく UI スレッド。フォルダが多くて遅いときの扱いは今のまま。

## 11. 受け入れ

見比べる画像: `../png/108_230.png`（H。既定）・`../png/148_67.png`（H0。空）・`../png/175_4085.png`（E13）・`../png/176_2571.png`（ツールチップ）。窓 1280×820、左の欄 220。

- 欄の地が #F0F2F4、中身が左右 10・上 12 から始まる。
- 見出し「検索対象 3 / 4」（12 SemiBold、#202124）と、同じ行の右に「すべて」「解除」（11、#0078D4、間 10）。
- 見出しの 8 下に、高さ 28・角丸 4・枠 #C9CED4・地 #FFFFFF の「フォルダを探す」（虫眼鏡 14、#9AA0A6 の文字）。
- 行の高さが 24。根の行の開閉の印が x=12（欄の左端から。10＋2）、チェック、フォルダ（#E8A020）、名前（12 Medium）の順に間 4。子の行は 16 右。
- 「顧客」の直後に橙（#E8A400）、「営業部2025」の直後に青（#0078D4）、「アーカイブ」の直後に赤（#D93025）の 7px の点。点は右端にそろえず、名前のすぐ後ろ。
- 長い名前が「提案書_2024年度下…」のように 1 行で切れ、カーソルを当てると行が灰色になり、フルパスがツールチップで出る（拡張子を足さない）。
- チェックの付いた箱は青の地に白の印、付いていない箱は白の地に 1.2px の灰色（#80868B）の枠。
- インデックスが無いとき: 見出しが「検索対象」だけで、リンク・絞り込みの欄が無く、橙のフォルダのアイコンと「検索できるフォルダがありません」（12、#5F6368）が中央に出る。
- 「解除」を押すと見出しが「検索対象 0 / 4」になり、点は残り、欄のいちばん下に ⓘ と「検索するフォルダを選んでください」（11、#5F6368）が出る。
- 左の欄を 180 にしても、点が名前の直後に残り、名前だけが「…」で縮む。

## 決めること

- 一部だけチェックの付いた箱（3 つ目の状態）の見た目。Figma に見本が無い。
- 見出しの数に、一部だけチェックの付いた根を数えるか。
- 点の状態（要更新・更新中・エラー）と、今のインデックスの状態（未作成・中断中・失敗あり・作成中・通常、インデックス管理のバッジ）の対応。中断中・未作成・元のフォルダが見つからない根に点を出すか。
- 点は色だけで状態を示している（`docs/design/gui/common.md` の「色だけで示さない」と食い違う）。点にツールチップで状態の名前を出すか、別の形にするか。
- 行のホバーの地の色（176_2571 では灰色・枠なし。値は未定）と、選んだ行の見た目、行の名前をクリックしたときの動き。
- 「すべて」「解除」のホバー・押したとき・効かないときの見た目。絞り込みの欄のフォーカスの見た目と、打った文字の色。
- 絞り込みの決まり: 大文字・小文字と全角・半角を区別するか、まだ開いていない段まで探すか、打ってから絞り込むまでの間、絞り込みのあいだの「すべて」「解除」の効く範囲。
- 「（このフォルダ直下のファイル）」の行・開く前の仮の行・元のフォルダが見つからない根の見た目と、見つからないときのツールチップの文。
- 3 段目より下の段下げの幅（Figma は 2 段だけ。仮に 16 ずつ）。
- 空の状態の x:Name `IndexTreeEmpty` を足すこと（`figma_wpf_map.md` に無い）。今の `IndexTreePlaceholder`（インデックスが無いときの文）を、Figma のとおり P-E13 の下の案内に役目替えすること。
- `describeSearchTargets` の文（今の「検索対象：すべて（集約ファイル N 件 ・ 最終取り込み …）」など）をどこに出すか、やめるか。Figma の部品の説明の「データ: getUsedIndexNames・describeSearchOption」との関係。
- 下の案内（P-E13）が左の欄 180 のときに入りきらない場合の扱い（切る・折り返す）。
- 新しい Style `TreeCheck`・`TextLink`・`TextBox.Bare` を theme に足すこと、`figma_wpf_map.md` の使うキー（Link・Ok・Warn・Danger.Text）を Figma の部品の値（Meta＋Accent・Warn.Dot・Accent・Danger.Dot）に直すこと。
