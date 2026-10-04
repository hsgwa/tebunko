# 検索結果の一覧

## 1. 役割

検索の画面の真ん中に置く欄。上から次のものが並ぶ。

- 件数の行
  - 件数
  - すべて開く・すべて折りたたむ
  - 絞り込み
  - ［保存］
- 列の見出し
- ファイルごとにまとめた該当行

ファイルの見出しで開く・折りたたみ、行を選ぶとプレビュー（preview.md）にその場所を出す。

**Figma**

- 部品は `xaml/search/result_list`（352:18716。既定は 340:723）。
- variant の node ID:

  | variant | node ID |
  |---|---|
  | H | 352:16918 |
  | H-B | 352:17031 |
  | H-W | 352:17144 |
  | H-R | 352:17288 |
  | H-F | 352:17353 |
  | P-H | 352:17811 |
  | P-H-E2 | 352:17924 |
  | P-H-S | 352:18037 |
  | P-E16 | 352:18150 |
  | P-E14 | 352:18263 |
  | P-E12 | 352:18376 |
  | P-H-saved | 352:18489 |
  | P-H-row3 | 352:18602 |
  | P-H-row4 | 352:18715 |
  | H-PP | 426:50 |
  | H-TX | 426:183 |
  | P-H-folded | 501:10889 |
  | P-H-open | 501:11032 |

**使う画面と `../png/` の画像**

| 画面 | 画像 |
|---|---|
| H | `108_230.png` |
| H-S | `157_944.png` |
| H-R | `135_1498.png` |
| H-W | `125_1839.png` |
| H-B | `108_1116.png` |
| H-P | `108_1419.png` |
| E12 | `157_2424.png` |
| E16 | `157_1238.png` |
| H-saved | `158_527.png` |
| H-row3 | `158_823.png` |
| H-row4 | `158_1118.png` |
| H-PP | `427_36175.png` |
| H-TX | `427_36508.png` |

- H-F（絞り込み中）は画像が無い。Figma の 352:17353 を `get_screenshot` で見る。

## 2. 置き場所

- XAML は `scripts/tebunko/xaml/search/result_list.xaml`。根は `Grid`（x:Name `ResultList`）。
- `search.xaml` の **ResultListHost**（`SearchRoot` の MainRow。`*`・MinHeight 120）に差す。
- 空の状態（H0）・該当なし（E14）は search.md の `EmptyState`・`NoResultState` が受け持つ。そのとき、この部品は Collapsed にする。
- 画面層は `ui/search/result_list.ps1`。
  - 今の `ui/result_list.ps1` の `clearResults`・`newFileGroup`・`toggleFileGroup`・`setAllFileGroupsExpanded`・`applyResultFilter`・`sortResults`・`finishResults`・`getShownHitCount` を移す。
- 判断層は `ui/search/result_list_view.ps1`。
  - 今の `search_view.ps1` から次を移して直す。
    - `getSearchSummaryText`・`getSearchProgressText`
    - `describeFileLocations`・`getAppKind`
    - `selectShownRows`・`getResultItems`・`getShownHitRows`・`prepareHitRow`
    - `sortFileGroups`

## 3. 部品の木

```xml
<Grid x:Name="ResultList" Background="{StaticResource Bg.Surface}">
  <Grid.RowDefinitions>
    <RowDefinition Height="Auto"/>  <!-- 件数の行 -->
    <RowDefinition Height="Auto"/>  <!-- 列の見出し -->
    <RowDefinition Height="*"/>     <!-- 一覧 -->
  </Grid.RowDefinitions>

  <!-- 件数の行 -->
  <Border x:Name="SummaryRow" Grid.Row="0" Background="{StaticResource Bg.Subtle}"
          BorderBrush="{StaticResource Border.Soft}" BorderThickness="0,0,0,1" Padding="16,4">
    <WrapPanel x:Name="SummaryWrap" Orientation="Horizontal">
      <StackPanel x:Name="SummaryLeft" Orientation="Horizontal" MinHeight="24">
        <TextBlock x:Name="SummaryText" VerticalAlignment="Center"
                   Style="{StaticResource Meta.Strong}" Foreground="{StaticResource Ink.Strong}"
                   Text="{result.summary}"/>
        <StackPanel x:Name="ToggleLinks" Orientation="Horizontal" Margin="12,0,0,0">
          <Button x:Name="ExpandAllButton" Style="{StaticResource Link.Small}" Content="{result.expandall}"/>
          <Rectangle Width="1" Height="10" Margin="12,0" Fill="{StaticResource Border.Normal}"
                     VerticalAlignment="Center"/>
          <Button x:Name="CollapseAllButton" Style="{StaticResource Link.Small}" Content="{result.collapseall}"/>
        </StackPanel>
      </StackPanel>
      <StackPanel x:Name="SummaryRight" Orientation="Horizontal" MinHeight="24">
        <TextBlock Text="{result.filter.label}" VerticalAlignment="Center"
                   Style="{StaticResource Meta}" Foreground="{StaticResource Ink.Body}"/>
        <Grid Margin="6,0,0,0" Width="220" Height="24">
          <TextBox x:Name="FilterBox" Padding="8,0" VerticalContentAlignment="Center"
                   Style="{StaticResource TextBox.Base}" FontSize="11"
                   Background="{StaticResource Bg.Surface}" BorderBrush="{StaticResource Border.Input}"
                   Foreground="{StaticResource Ink.Strong}"/>                <!-- 角丸 4 -->
          <TextBlock x:Name="FilterPlaceholder" Margin="9,0,0,0" VerticalAlignment="Center"
                     IsHitTestVisible="False" Style="{StaticResource Meta}"
                     Foreground="{StaticResource Ink.Subtle}" Text="{result.filter.placeholder}"/>
        </Grid>
        <Button x:Name="ExportButton" Margin="12,0,0,0" Padding="8,4,6,4" Style="{StaticResource ToolButton}">
          <StackPanel Orientation="Horizontal">
            <Viewbox Width="12" Height="12"><Path Data="{StaticResource Icon.Download}"
                     Stroke="{StaticResource Button.Icon}" StrokeThickness="2"/></Viewbox>
            <TextBlock Margin="4,0,0,0" Style="{StaticResource Meta}"
                       Foreground="{StaticResource Button.Text}" Text="{result.export}"/>
          </StackPanel>
        </Button>
      </StackPanel>
    </WrapPanel>
  </Border>

  <!-- 列の見出し（1 本だけ。全幅） -->
  <Border x:Name="ColumnHeader" Grid.Row="1" Height="22" Padding="16,0"
          Background="{StaticResource Bg.Subtle}" BorderBrush="{StaticResource Border.Soft}"
          BorderThickness="0,0,0,1">
    <Grid>
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="180"/><ColumnDefinition Width="12"/>
        <ColumnDefinition Width="48"/><ColumnDefinition Width="12"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <TextBlock Grid.Column="0" Style="{StaticResource ColumnHeader}" Foreground="{StaticResource Ink.Body}"
                 VerticalAlignment="Center" Text="{result.col.place}"/>
      <TextBlock Grid.Column="2" Style="{StaticResource ColumnHeader}" Foreground="{StaticResource Ink.Body}"
                 VerticalAlignment="Center" HorizontalAlignment="Right" Text="{result.col.kind}"/>
      <TextBlock Grid.Column="4" Style="{StaticResource ColumnHeader}" Foreground="{StaticResource Ink.Body}"
                 VerticalAlignment="Center" Text="{result.col.text}"/>
    </Grid>
  </Border>

  <!-- 一覧。仮想化のため ListBox＋GroupStyle にする -->
  <ListBox x:Name="ResultGrid" Grid.Row="2" BorderThickness="0"
           ScrollViewer.VerticalScrollBarVisibility="Auto"
           ScrollViewer.HorizontalScrollBarVisibility="Disabled"
           VirtualizingPanel.IsVirtualizing="True" VirtualizingPanel.IsVirtualizingWhenGrouping="True"
           VirtualizingPanel.VirtualizationMode="Recycling" AlternationCount="2"
           ItemContainerStyle="{StaticResource HitRow}">
    <ListBox.GroupStyle>
      <GroupStyle ContainerStyle="{StaticResource FileGroup}"/>   <!-- 下の「ファイルの見出し」 -->
    </ListBox.GroupStyle>
  </ListBox>
</Grid>
```

**ファイルの見出し（`FileGroup` の Expander の Header。開いているとき）**

```xml
<Border Height="36" Padding="12,0,16,0" Background="{StaticResource Bg.Window}"
        BorderBrush="{StaticResource Border.Soft}" BorderThickness="0,0,0,1">
  <DockPanel>
    <Border DockPanel.Dock="Left" Width="3" Height="16" CornerRadius="1.5" Background="{StaticResource Accent}"/>
    <Path DockPanel.Dock="Left" Margin="8,0,0,0" Width="8" Height="5" Data="{StaticResource Icon.ChevronDown}"
          Stretch="Uniform" Stroke="{StaticResource Ink.Body}" StrokeThickness="1.2"/>
    <Viewbox DockPanel.Dock="Left" Margin="8,0,0,0" Width="14" Height="14">
      <Path x:Name="FileIcon" Data="{StaticResource Icon.FileSpreadsheet}" Stroke="{StaticResource File.Excel}"
            StrokeThickness="2"/></Viewbox>                                 <!-- 線 14 × 2/24 = 1.17 -->
    <TextBlock DockPanel.Dock="Left" Margin="8,0,0,0" x:Name="FileName" VerticalAlignment="Center"
               Style="{StaticResource Body.Strong}" Foreground="{StaticResource Ink.Strong}"
               TextTrimming="CharacterEllipsis"/>
    <Border DockPanel.Dock="Left" Margin="8,0,0,0" x:Name="PathBadge" Padding="6,1" CornerRadius="4"
            Background="{StaticResource Bg.Hover}" VerticalAlignment="Center">
      <TextBlock Style="{StaticResource Micro}" Foreground="{StaticResource Ink.Body}"
                 TextTrimming="CharacterEllipsis"/>                         <!-- 例「営業部/2024/見積もり」 -->
    </Border>
    <TextBlock DockPanel.Dock="Right" x:Name="FileSummary" VerticalAlignment="Center"
               Style="{StaticResource Meta.Strong}" Foreground="{StaticResource Accent}"/>
    <Border/>                                                               <!-- 空き Fill -->
  </DockPanel>
</Border>
```

折りたたんだときは、次のように変える。

| 部品 | 開いているとき | 折りたたんだとき |
|---|---|---|
| 地 | Bg.Window | Bg.Subtle |
| 青い棒 | 出す | 出さない（Collapsed） |
| chevron | chevron-down（8×5） | chevron-right（5×8） |
| FileName | Body.Strong | Heading（13 SemiBold） |
| PathBadge | 出す | 出さない |
| FileSummary | Meta.Strong・Accent | Meta（11 Regular）・Ink.Body |

**該当行（`HitRow` の ListBoxItem の Template）**

```xml
<Border x:Name="RowBg" Height="30" Padding="16,0" Background="{StaticResource Bg.Surface}">
  <Grid>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="180"/><ColumnDefinition Width="12"/>
      <ColumnDefinition Width="48"/><ColumnDefinition Width="12"/>
      <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>
    <TextBlock Grid.Column="0" Text="{Binding Place}" VerticalAlignment="Center"
               Style="{StaticResource Meta}" Foreground="{StaticResource Ink.Body}"
               TextTrimming="CharacterEllipsis"/>
    <TextBlock Grid.Column="2" Text="{Binding Kind}" VerticalAlignment="Center" HorizontalAlignment="Right"
               Style="{StaticResource Meta}" Foreground="{StaticResource Ink.Body}"/>
    <TextBlock Grid.Column="4" x:Name="HitText" VerticalAlignment="Center" TextWrapping="NoWrap"
               Style="{StaticResource Body}" Foreground="{StaticResource Ink.Strong}"
               ClipToBounds="True"/>                                         <!-- Inlines を画面層で組む -->
  </Grid>
</Border>
<!-- Trigger: ItemsControl.AlternationIndex=1 → RowBg.Background = Bg.Stripe
              IsSelected=True              → RowBg.Background = Select.Soft -->
```

- 当たった文字は、`HitText` の Inlines に組んで目立たせる。
  - 地は Hit（#FFF176）、文字は Bold。
  - Figma の余白（左右 2・上下 1）と角丸 2 は、Run では出せない。そこで `InlineUIContainer` に Border＋TextBlock を入れる。
  - 行の高さがずれないよう、Border の Margin を 0,-1 にする（受け入れで見比べる）。

## 4. 寸法と色

| 要素 | 大きさ | 余白 | 地の色 | 枠 | 文字の Style | 文字の色 | 角丸 | アイコン |
|---|---|---|---|---|---|---|---|---|
| SummaryRow | 横 Fill × Hug（32。折り返すと 24＋4＋24＋8＝60） | 上下 4・左右 16。左の塊と右の塊は両端に寄せ、入らなければ右の塊を次の行へ（行の間 4） | Bg.Subtle（#F9FAFA） | 下 1 Border.Soft（#E0E2E5） | – | – | 0 | – |
| SummaryText | Hug | – | – | – | Meta.Strong（11 Bold） | Ink.Strong（#202124） | – | – |
| ExpandAllButton・CollapseAllButton | Hug | 件数との間 12・2 つの間 12（間に 1×10 の縦線 Border.Normal #D9DEE3） | 無し | 無し | Meta（11 Regular）・下線 | Accent（#0078D4） | – | – |
| 「絞り込み」 | Hug | 欄との間 6 | – | – | Meta | Ink.Body（#5F6368） | – | – |
| FilterBox | 220 × 24 | 左右 8 | Bg.Surface（#FFFFFF） | 1 Border.Input（#D1D1D1） | 11 Regular | Ink.Strong。薄い文字は Ink.Subtle（#80868B） | 4 | – |
| ExportButton | Hug × 24 | 左 8・右 6・上下 4、中の間 4、欄との間 12 | Bg.Surface | 1 Border.Input | Meta | Button.Text（#4D4D4D） | 4 | download 12・線 1.0（Button.Icon #666666） |
| ColumnHeader | 横 Fill × 22 | 左右 16、列の間 12 | Bg.Subtle | 下 1 Border.Soft | ColumnHeader（10 Bold） | Ink.Body | 0 | – |
| 列「場所」 | 180 | – | – | – | – | – | – | – |
| 列「種別」 | 48・右寄せ | – | – | – | – | – | – | – |
| 列「該当行」 | Fill | – | – | – | – | – | – | – |
| ファイルの見出し（開いている） | 横 Fill × 36 | 左 12・右 16、間 8 | Bg.Window（#F5F7FA） | 下 1 Border.Soft | 名前 Body.Strong（13 Bold）、パス Micro（10）、まとめ Meta.Strong（11 Bold） | 名前 Ink.Strong、パス Ink.Body、まとめ Accent | 棒 1.5、パス 4 | 棒 3×16 Accent、chevron-down 8×5、種類 14・線 1.17 |
| ファイルの見出し（折りたたみ） | 横 Fill × 36 | 左 12・右 16、間 8 | Bg.Subtle（#F9FAFA） | 下 1 Border.Soft | 名前 Heading（13 SemiBold）、まとめ Meta（11 Regular） | 名前 Ink.Strong、まとめ Ink.Body | – | chevron-right 5×8、種類 14 |
| パスのバッジ | Hug | 左右 6・上下 1 | Bg.Hover（#F3F3F4） | 無し | Micro（10 Regular） | Ink.Body | 4 | – |
| 該当行 | 横 Fill × 30 | 左右 16、列の間 12 | 偶数行 Bg.Surface・奇数行 Bg.Stripe（#FAFBFC）・選んだ行 Select.Soft（#E1F2FF） | **行の間に線を引かない** | – | – | 0 | – |
| 行の「場所」 | 180 | – | – | – | Meta（11 Regular） | Ink.Body（#5F6368） | – | – |
| 行の「種別」 | 48・右寄せ | – | – | – | Meta | Ink.Body | – | – |
| 行の「該当行」 | Fill | – | – | – | Body（13 Regular） | Ink.Strong | – | – |
| 当たった文字 | Hug | 左右 2・上下 1 | Hit（#FFF176） | 無し | Body.Strong（13 Bold） | Ink.Strong | 2 | – |

**ファイルの種類のアイコン**（`figma_wpf_map.md`。14px・線 1.17）

| 種類 | アイコン | 色 |
|---|---|---|
| Excel | file-spreadsheet（開いているとき）。**折りたたんだ見出しでは chart-column**（`figma_wpf_map.md`。diff_round2 で「揺れではない」と決めた） | File.Excel（#107C41） |
| Word | file-text | File.Word（#185ABD） |
| PowerPoint | presentation | File.PowerPoint（#C43E1C） |
| テキスト（.txt・.md） | file | Ink.Body（#5F6368） |

## 5. 状態ごとの見え方

| 状態 | SummaryRow | SummaryText | すべて開く／折りたたむ | 絞り込み | ［保存］ | ColumnHeader | 一覧 |
|---|---|---|---|---|---|---|---|
| H0・E13・E14 | – | – | – | – | – | – | **この部品を出さない**（search.md） |
| H-S（検索中） | 出す | `{result.progress}`「検索中… 該当 5 件」 | **出さない** | 出す | **出さない** | 出す | 見つかった順に足していく |
| H | 出す | `{result.summary}`＋`{result.mode.fast}` | 出す | 出す | 出す | 出す | 先頭のファイルだけ開き、1 行目を選ぶ |
| H-R・E12 | 出す | `{result.summary}`＋`{result.mode.normal}` | 出す | 出す | 出す | 出す | H と同じ形 |
| H-B・H-P | 出す | `{result.summary}`（方式を付けない） | 出す | 出す | 出す | 出す | H と同じ形 |
| H-F（絞り込み中） | 出す | `{result.filter.summary}` | 出す | 出す（打った文字） | 出す | 出す | 合うファイル・合う行だけ。見出しのまとめも、見えている数で数え直す |
| H-W | 出す | H と同じ | 出す | 出す | 出す | 出す | Word のファイルが開いている |
| H-PP・H-TX | 出す | H と同じ | 出す | 出す | 出す | 出す | PowerPoint・テキストのファイルが開いている |
| E16（中止） | **出さない** | – | – | – | – | 出す（一覧の先頭が列の見出し） | 止めるまでに見つかった分 |
| H-saved | 出す | H と同じ | 出す | 出す | 出す | 出す | H と同じ（保存の帯は search.md） |
| row3・row4 | 出す | H と同じ | 出す | 出す | 出す | 出す | 3 行目・4 行目を選んだ |
| P-H-folded（すべて折りたたんだ） | 出す | H と同じ | 出す | 出す | 出す | 出す | 6 つのファイルの見出しだけ。行は出さない。プレビューは折りたたむ前に選んでいた行のまま |
| P-H-open（すべて開いた） | 出す | H と同じ | 出す | 出す | 出す | 出す | 6 つのファイルをすべて開く。入りきらない分は縦にスクロール |

**H の一覧（Figma の例。並びは下の「決めること」）**

| 順 | ファイル | 開閉 | まとめ |
|---|---|---|---|
| 1 | A社_見積書.xlsx | 開く | [シート]見積書 ほか 1 か所 ・ 5 件 |
| 2 | A社_見積書_改訂.xlsx | 折りたたみ | [シート]見積書 ・ 3 件 |
| 3 | 2025年4月.pptx | 折りたたみ | スライド 2 ・ 1 件 |
| 4 | 基本契約書.docx | 折りたたみ | 1 ページ（目安） ・ 3 件 |
| 5 | 顧客一覧.xlsx | 折りたたみ | [シート]顧客 ・ 2 件 |

H の 1 番目（A社_見積書.xlsx）の該当行:

| 場所 | 種別 | 該当行 |
|---|---|---|
| [シート]見積書!A3 | セル | 見積先：**(株)山田商事**(御中) │ 見積番号：12-2324-0143 |
| [シート]見積書!B18 ほか 1 | セル | 4月4日 │ 値引き：**(株)山田商事** 特別割引あり |
| [シート]見積書!A41 | セル | 納品場所：**(株)山田商事** │ 本社ビル |
| [シート]見積書!D5 | 図形 | 納品場所：**(株)山田商事** 本社 4F |

ほかの variant の例:

| variant | ファイル | まとめ | 該当行 |
|---|---|---|---|
| H-W | 基本契約書.docx | 1 ページ（目安） ほか 2 か所 ・ 3 件 | 「1 ページ（目安）」本文「甲：**(株)山田商事**（以下「甲」という）」・「第3条**(株)山田商事**は毎月末日までに支払う」・「署名欄：**(株)山田商事** 代表取締役 山田 太郎」 |
| H-PP | 2025年4月.pptx | スライド 2・1 件 | 「スライド 2」本文「導入効果：**(株)山田商事**様の事例」 |
| H-TX | .md のファイル | 行 12 ほか 1 か所 ・ 2 件 | 「行 12」本文「定例会の出席：**(株)山田商事** 佐藤様・山田」・「次回までに**(株)山田商事** 向けの見積を送る」 |
| H-TX | .txt のファイル | 29 行目 ・ 1 件 | 「29 行目」本文 |

**PoC の誤り（diff_round2〜4 で指摘したもの）**

| 誤り | 正しい形 | 出どころ |
|---|---|---|
| 列の見出しをファイルごとに出した | 一覧の上に 1 本だけ、全幅で出す | diff_round2 の 1 |
| 行が x 237〜1264 に収まっていた。列の見出しも字下げしてあった | 行も見出しも全幅。見出しの地は Bg.Subtle（#F9FAFA） | diff_round3 の 1 |
| ファイルの行の間に線が無かった | ファイルの見出しの下に 1 本の線（Border.Soft）。該当行の間には引かない | diff_round2 の 1・diff_round3 の 1 |
| 選んだ行が光っていなかった | Select.Soft（#E1F2FF）。H・H-row3・H-row4 | diff_round2・3 |
| 「｜」を列と列の間に置いた | 区切りは該当行の文の中だけ。列の間には置かない | diff_round3 の 1 |
| 結果の無い状態（H0・H-E・H-E2・H-1・E14）に列の見出しを出した | この部品ごと出さない | diff_round3 の 4（★） |
| E14 で件数の行に「0件」を出した | 件数の行を出さず、ステータスバーを「0 件」にする（search.md） | diff_round3 の 4 |
| ［すべて開く／すべて折りたたむ］の置き場所・形が違った | 件数の右の中ほど。下線を付け、間に区切りの縦線 | diff_round2（H）・diff_round3 |
| H-B で件数に「・高速検索」が出ていた | 更新中（H-B・H-P）は方式を付けない | diff_round2（H-B） |
| 絞り込み中も一覧が全部出ていた。欄の文字の前に余白があった | 「15 件中 3 件を表示」なら一覧も 3 件。欄の左の余白は 8 だけ | diff_round2（絞り込み中） |
| 1024 で縦のスクロールバーが常に出ていた | `VerticalScrollBarVisibility="Auto"` | diff_round3 の 6 |
| 該当行の文が Figma と違った | Figma の文をそのまま出す（上の表） | diff_round4 の 1 |
| 場所・種別が本文と同じ大きさ・色で、場所の列が狭かった（種別 x 393・該当行 x 453） | 11px の灰色。種別の右端 x 476・該当行 x 488（窓 1280） | diff_round4 の 2 |
| 開いたファイルの青い帯が見えなかった。閉じた「顧客一覧.xlsx」が file-spreadsheet だった | 帯 3×16 Accent。閉じた Excel は chart-column | diff_round4 の 3 |
| 件数・まとめ・当たった文字が太字でなかった | 件数・まとめは Meta.Strong（11 Bold）、当たった文字は Bold | diff_round2 の 5 |

## 6. 文言

| 文言の ID | 全文 | 使う状態 |
|---|---|---|
| `result.summary` | {該当数} 件（{ファイル数} ファイル） | すべて。例「15 件（6 ファイル）」。かかった秒数と検索の方式は出さない（2026-10-04 に決定。利用者に要る情報だけを出す。方式は検索バーの高速検索のバッジで分かる） |
| `result.filter.summary` | {該当数} 件中 {見えている数} 件を表示 | H-F。例「15 件中 3 件を表示」 |
| `result.progress` | 検索中… 該当 {該当数} 件 | H-S。例「検索中… 該当 5 件」。「…」の後ろは**半角**の空白 |
| `result.expandall` | すべて開く | 件数の行 |
| `result.collapseall` | すべて折りたたむ | 件数の行 |
| `result.filter.label` | 絞り込み | 件数の行 |
| `result.filter.placeholder` | ファイル名・場所・中身で絞り込む | 絞り込みの欄が空のとき |
| `result.export` | 保存 | 件数の行 |
| `result.col.place` | 場所 | 列の見出し |
| `result.col.kind` | 種別 | 列の見出し |
| `result.col.text` | 該当行 | 列の見出し |
| `result.group.summary` | {場所} ・ {件数} 件 | ファイルの見出し。例「[シート]見積書 ・ 3 件」。「・」の前後は半角の空白 |
| `result.group.summary.more` | {最初の場所} ほか {残りの数} か所 ・ {件数} 件 | 例「[シート]見積書 ほか 1 か所 ・ 5 件」 |
| `result.kind.cell` | セル | 行の種別（Excel のセル） |
| `result.kind.shape` | 図形 | 行の種別 |
| `result.kind.body` | 本文 | 行の種別（Word・PowerPoint・テキスト） |
| `result.place.sheet` | [シート]{シート名}!{セル} | 例「[シート]見積書!A3」 |
| `result.place.sheet.more` | [シート]{シート名}!{セル} ほか {残りの数} | 同じ行に当たったセルが複数。例「[シート]見積書!B18 ほか 1」 |
| `result.place.page` | {ページ} ページ（目安） | Word。例「1 ページ（目安）」 |
| `result.place.slide` | スライド {番号} | PowerPoint。例「スライド 2」 |
| `result.place.line.md` | 行 {番号} | .md。例「行 12」 |
| `result.place.line.txt` | {番号} 行目 | .txt。例「29 行目」（「行 12」との食い違い → 決めること） |
| `result.sep` | │ | 行の中の区切り（U+2502）。前後は半角の空白 |

- 数は、今のコードと同じく桁区切りを付けない。付けるか（N0）は search.md の決めることと同じ。
- 「・」は全角の中黒。かっこは全角。

## 7. 操作と振る舞い

| きっかけ | 何が起きるか | 呼ぶ関数 | 次の状態 |
|---|---|---|---|
| 検索が始まる | 一覧を空にし、件数の行を `{result.progress}` にする | `clearResults`・`getSearchProgressText` | H-S |
| 検索の途中で見つかる | ファイルの見出しと行を足す（まとめて `Dispatcher` で）。件数を更新する | `newFileGroup`・`prepareHitRow`・`getSearchProgressText` | H-S |
| 検索が終わる | 並べ直し、先頭のファイルだけを開き、1 行目を選ぶ。件数を `{result.summary}` にする | `finishResults`・`sortFileGroups`・`getSearchSummaryText` | H・H-R |
| 中止で止まる | 件数の行を消す。見つかった分は残す | `getResultListView` | E16 |
| ファイルの見出しをクリック | そのファイルを開く・折りたたむ | `toggleFileGroup` | 変わらない |
| 「すべて開く」 | すべてのファイルを開く | `setAllFileGroupsExpanded $true` | 変わらない |
| 「すべて折りたたむ」 | すべてのファイルを折りたたむ | `setAllFileGroupsExpanded $false` | 変わらない |
| 絞り込みの欄に打つ | 少し待ってから（今のコードは 300ms）、ファイル名・場所・中身のどれかに合う行だけを見せる。件数の行を `{result.filter.summary}` にする | `applyResultFilter`・`getFilterSummaryText` | H-F（空に戻すと H） |
| 行をクリック・矢印キーで選ぶ | 選んだ行を Select.Soft にし、プレビューにその場所を出す | `onResultSelected`（preview.md） | row3・row4 |
| 行をダブルクリック | そのファイルを開く（Office ならその場所へ） | `openHitLocation` | 変わらない |
| ［保存］ | 結果を保存し、保存の帯を出す | `exportResults` | H-saved |
| 行を右クリック | 行のメニューを、押した位置に出す。項目は上から「コピー」「パスをコピー」（区切り線）「開く」「フォルダを開く」。メニューの外を押すと閉じる。キー操作のヒント（InputGestureText）は出さない | 今の右クリックメニューと同じ処理 | 変わらない |

- **Figma の例**: H の一覧の 6 つ目は、折りたたんだテキストのファイル（定例会議事録.txt。まとめは「29 行目 ・ 1 件」）。押すと H-TX。件数は「15 件（6 ファイル）」。
- **プロトタイプでの見せ方**: 「07 プロトタイプ」に H-fold（501:53806）・H-open（501:53846）・H-ctx（行のメニュー。501:53886）がある。Figma には右クリックのきっかけが無いので、H-ctx は H の行の長押しで出す（実物は右クリック）。
- **ホバー**: 行・ファイルの見出しのホバーの見た目は Figma に無い。決まるまでは、地を変えない。
- **キー操作**: ダブルクリックと同じことを Enter でもできる。キー操作のヒントは画面に出さない。
- **Tab の順**: ExpandAllButton → CollapseAllButton → FilterBox → ExportButton → ResultGrid。

## 8. リサイズ

一覧が多いとき（スクロールする所・固定する所・仮想化・未定のこと）は `../scroll.md`（見本 H-多・H-多-上限・H-多-1024）。

- 横の幅:
  - 「場所」180・「種別」48 は固定。「該当行」だけが伸びる（`*`）。
  - 場所は「…」で切る。該当行は折り返さず、右端で切る（当たった文字が見えるよう、`prepareHitRow` が当たった場所の前後を切り出す）。
- 件数の行:
  - 左の塊（件数・開閉）と右の塊（絞り込み・［保存］）を両端に寄せる。
  - 入らないときは、右の塊を次の行の左へ落とす（行の間 4）。
  - 両端に寄せる計算は、search_bar.md と同じく `SizeChanged` で右の塊の左の Margin を決める。
- 縦の高さ:
  - 一覧は MainRow の `*` を使い、最小 120。
  - 入りきらない分は、縦のスクロールバーで見る（Auto）。横のスクロールは出さない。
- 窓ごとの見え方:
  - 1280: 件数の行は 1 行。
  - 1024（主な領域 804）: 件数の行が「15 件（6 ファイル）」＋開閉（約 360）と、絞り込み＋［保存］（約 340）。
    - Padding を足しても 804 に入るので、1 行になる見込み（幅は目安。測っていない）。
    - H-F の長い文言でも入るかは、受け入れで確かめる。
  - 1600: 該当行の列が広がるだけ。

## 9. 判断層

`ui/search/result_list_view.ps1`。WPF の型に触らない。

### getSearchSummaryText

今の関数の形を変える。

入力: `hits`（int）・`files`（int）。

出力: string。

```powershell
It "<Name>" -TestCases @(
  @{ Name="H";   Hits=15; Files=6; Expected="15 件（6 ファイル）" }
  @{ Name="H-R"; Hits=16; Files=6; Expected="16 件（6 ファイル）" }
  @{ Name="0 件"; Hits=0; Files=0; Expected="0 件（0 ファイル）" }
) { param($Hits,$Files,$Expected) getSearchSummaryText $Hits $Files | Should -Be $Expected }
```

- 今のコードの「該当 N 件（F ファイル） ・ S 秒」から、次の 2 点を直す。
  - 「該当」を取る。
  - 秒数を取る（方式も付けない。2026-10-04 に決定）。

### getFilterSummaryText

入力: `shown`（int）・`hits`（int）。

出力: string。

```powershell
It "<Name>" -TestCases @(
  @{ Name="H-F";  Shown=3;  Hits=15; Expected="15 件中 3 件を表示" }
  @{ Name="0 件"; Shown=0;  Hits=16; Expected="16 件中 0 件を表示" }
  @{ Name="全部"; Shown=15; Hits=15; Expected="15 件中 15 件を表示" }
) { param($Shown,$Hits,$Expected) getFilterSummaryText $Shown $Hits | Should -Be $Expected }
```

- 絞り込んで 0 件になったときの見え方（一覧が空のまま・別の文言を出す）は、Figma に無い → 決めること。

### getSearchProgressText

```powershell
It "<Name>" -TestCases @(
  @{ Name="0 件"; Hits=0; Expected="検索中… 該当 0 件" }
  @{ Name="5 件"; Hits=5; Expected="検索中… 該当 5 件" }
) { param($Hits,$Expected) getSearchProgressText $Hits | Should -Be $Expected }
```

- 今のコードは「検索中…」の後ろが全角の空白。半角に直す（Figma）。

### getFileGroupSummary

今の `describeFileLocations` を使って、件数まで付ける。

入力: `labels`（string[]。そのファイルで当たった場所の名前。重なりは除いて順を保つ）・`count`（int）。

出力: string。

```powershell
It "<Name>" -TestCases @(
  @{ Name="1 か所";   Labels=@("[シート]見積書");                          Count=3; Expected="[シート]見積書 ・ 3 件" }
  @{ Name="2 か所";   Labels=@("[シート]見積書","[シート]単価表");          Count=5; Expected="[シート]見積書 ほか 1 か所 ・ 5 件" }
  @{ Name="Word";     Labels=@("1 ページ（目安）","2 ページ（目安）","3 ページ（目安）"); Count=3; Expected="1 ページ（目安） ほか 2 か所 ・ 3 件" }
  @{ Name="スライド"; Labels=@("スライド 2");                              Count=1; Expected="スライド 2 ・ 1 件" }
  @{ Name="md";       Labels=@("行 12","行 30");                           Count=2; Expected="行 12 ほか 1 か所 ・ 2 件" }
) { param($Labels,$Count,$Expected) getFileGroupSummary $Labels $Count | Should -Be $Expected }
```

- Excel は、場所をシートの単位でまとめる。Word・PowerPoint・テキストは、行の「場所」をそのまま使う。
- H-PP の「スライド 2・1 件」（空白が無い）は、ほかと形をそろえて「スライド 2 ・ 1 件」にする案。決めることに挙げる。

### getHitPlaceText

今の `describeHitPlace` を直す。

入力: `appKind`（`'Excel'`・`'Word'`・`'PowerPoint'`・`'Markdown'`・`'テキスト'`）・`location`（hashtable）。

出力: `@{ Place = string; Kind = string }`。

```powershell
It "<Name>" -TestCases @(
  @{ Name="セル";     App='Excel';      Loc=@{Sheet="見積書"; Cell="A3"; Part='cell'};             Place="[シート]見積書!A3";      Kind="セル" }
  @{ Name="セル複数"; App='Excel';      Loc=@{Sheet="見積書"; Cell="B18"; More=1; Part='cell'};    Place="[シート]見積書!B18 ほか 1"; Kind="セル" }
  @{ Name="図形";     App='Excel';      Loc=@{Sheet="見積書"; Cell="D5"; Part='shape'};            Place="[シート]見積書!D5";      Kind="図形" }
  @{ Name="Word";     App='Word';       Loc=@{Page=1; Part='body'};                                Place="1 ページ（目安）";        Kind="本文" }
  @{ Name="PPT";      App='PowerPoint'; Loc=@{Slide=2; Part='body'};                               Place="スライド 2";              Kind="本文" }
  @{ Name="md";       App='Markdown';   Loc=@{Line=12; Part='body'};                               Place="行 12";                   Kind="本文" }
  @{ Name="txt";      App='テキスト';   Loc=@{Line=29; Part='body'};                               Place="29 行目";                 Kind="本文" }
) { param($App,$Loc,$Place,$Kind) $r = getHitPlaceText $App $Loc; $r.Place | Should -Be $Place; $r.Kind | Should -Be $Kind }
```

- コメント・ノートの種別の文言（「コメント」「ノート」か）と、その場所の書き方は、Figma に無い → 決めること。
- 決まるまでは表に入れない。

### getFileIconKind

入力: `appKind`・`isExpanded`（bool）。

出力: `@{ Icon = string; ColorKey = string }`。

```powershell
It "<Name>" -TestCases @(
  @{ Name="Excel 開";   App='Excel';      Open=$true;  Icon='file-spreadsheet'; Color='File.Excel' }
  @{ Name="Excel 閉";   App='Excel';      Open=$false; Icon='chart-column';     Color='File.Excel' }
  @{ Name="Word";       App='Word';       Open=$false; Icon='file-text';        Color='File.Word' }
  @{ Name="PowerPoint"; App='PowerPoint'; Open=$true;  Icon='presentation';     Color='File.PowerPoint' }
  @{ Name="md";         App='Markdown';   Open=$true;  Icon='file';             Color='Ink.Body' }
  @{ Name="txt";        App='テキスト';   Open=$false; Icon='file';             Color='Ink.Body' }
) { param($App,$Open,$Icon,$Color) $r = getFileIconKind $App $Open; $r.Icon | Should -Be $Icon; $r.ColorKey | Should -Be $Color }
```

- Excel の折りたたみを chart-column にするのは、`figma_wpf_map.md` と diff_round2 で決めたとおり。
- Figma の H-W では、折りたたんだ A社_見積書.xlsx だけが file-spreadsheet になっている。これは Figma の描き違いとして扱い、合わせない。

### getResultListView

入力: `state`（search.md の `getSearchScreenState` の戻り値。`'H'`・`'H-S'`・`'H-F'`・`'E16'` など）。

出力: `@{ Visible; ShowSummary; ShowToggles; ShowFilter; ShowExport; ShowHeader }`。

```powershell
It "<State>" -TestCases @(
  @{ State='H0';  Visible=$false; Summary=$false; Toggles=$false; Filter=$false; Export=$false; Header=$false }
  @{ State='E13'; Visible=$false; Summary=$false; Toggles=$false; Filter=$false; Export=$false; Header=$false }
  @{ State='E14'; Visible=$false; Summary=$false; Toggles=$false; Filter=$false; Export=$false; Header=$false }
  @{ State='H-S'; Visible=$true;  Summary=$true;  Toggles=$false; Filter=$true;  Export=$false; Header=$true }
  @{ State='H';   Visible=$true;  Summary=$true;  Toggles=$true;  Filter=$true;  Export=$true;  Header=$true }
  @{ State='H-F'; Visible=$true;  Summary=$true;  Toggles=$true;  Filter=$true;  Export=$true;  Header=$true }
  @{ State='H-B'; Visible=$true;  Summary=$true;  Toggles=$true;  Filter=$true;  Export=$true;  Header=$true }
  @{ State='E12'; Visible=$true;  Summary=$true;  Toggles=$true;  Filter=$true;  Export=$true;  Header=$true }
  @{ State='E16'; Visible=$true;  Summary=$false; Toggles=$false; Filter=$false; Export=$false; Header=$true }
) {
  param($State,$Visible,$Summary,$Toggles,$Filter,$Export,$Header)
  $v = getResultListView $State
  $v.Visible | Should -Be $Visible; $v.ShowSummary | Should -Be $Summary; $v.ShowToggles | Should -Be $Toggles
  $v.ShowFilter | Should -Be $Filter; $v.ShowExport | Should -Be $Export; $v.ShowHeader | Should -Be $Header
}
```

### そのまま移すもの

- `getAppKind`・`selectShownRows`・`getResultItems`・`getShownHitRows`・`prepareHitRow`・`sortFileGroups` は、今の関数とテストを移す。
- `prepareHitRow` は、区切りを「│」（U+2502）で返すことをテストに足す。

## 10. 画面層

`ui/search/result_list.ps1`。

- `initResultList` で次の配線をする。
  - `ExpandAllButton.Add_Click` → `setAllFileGroupsExpanded $true`
  - `CollapseAllButton.Add_Click` → `setAllFileGroupsExpanded $false`
  - `FilterBox.Add_TextChanged` → DispatcherTimer（300ms）を張り直す → `applyResultFilter`
  - `FilterBox` の文字の有無で `FilterPlaceholder.Visibility` を決める。
  - `ExportButton.Add_Click` → `exportResults`
  - `ResultGrid.Add_SelectionChanged` → `onResultSelected`
  - `ResultGrid.Add_MouseDoubleClick`・`Add_KeyDown`（Enter）→ `openHitLocation`
  - `SummaryWrap.Add_SizeChanged` → 右の塊を右に寄せる。
- `updateResultList $state` は `getResultListView $state` を見て、次の Visibility を決める。
  - `ResultList`・`SummaryRow`・`ToggleLinks`・`SummaryRight` の絞り込みの部分・`ExportButton`・`ColumnHeader`
- `ResultGrid` の項目:
  - 項目は、`prepareHitRow` で作った PSCustomObject（Place・Kind・Segments・File・Location）。
  - ファイルの単位で並べる。CollectionViewSource の `GroupDescriptions` に `File` を入れる。
  - グループの開閉は、ファイルごとの `IsExpanded` を辞書に持つ。Expander の `IsExpanded` に結ぶ。
- `HitText` の Inlines は、`Segments` から組む（`Loaded` と、再利用されたときの `DataContextChanged`）。
  - 当たった部分 → `InlineUIContainer`（Hit の地・Bold）
  - ほか → `Run`
- 見出しのアイコンは `getFileIconKind` の Icon を、`Icon.<名前>` の StaticResource に引き当てる。
- 見出しのまとめは `getFileGroupSummary`。
- 絞り込みの中は `getShownHitRows` で合う行を決め、CollectionView の `Filter` に入れる。
  - 見出しのまとめは、見えている行だけで数え直す（H-F）。
- 検索の途中は、見つかった分を 200ms ごとにまとめて足す（今のコードと同じ）。
- 上限（今のコードは 10,000 件）を超えたら止める。超えたことを示す文言は、Figma に無い → 決めること。

## 11. 受け入れ

窓 1280×820 で `../png/` と見比べる。

- H（`108_230.png`）:
  - 件数の行:
    - 高さ 32。地が #F9FAFA。
    - 左に「15 件（6 ファイル）」（11 Bold）。
    - 続けて、青い下線の「すべて開く」｜「すべて折りたたむ」。
    - 右に「絞り込み」と幅 220 の欄（薄い文字「ファイル名・場所・中身で絞り込む」）、［保存］（download のアイコン）。
  - 列の見出し:
    - 高さ 22。「場所」「種別」「該当行」が 10 Bold。
    - 「種別」の右端が x 476、「該当行」の左が x 488（diff_round4 の 2 の x≈455・488 と合う）。行の「セル」も同じ位置。
  - 1 番目のファイルの見出し:
    - 高さ 36。地が #F5F7FA。
    - 左から、青い棒・chevron-down・緑の file-spreadsheet・「A社_見積書.xlsx」（Bold）。
    - 続けて、灰色の地のパス。右端に青の「[シート]見積書 ほか 1 か所 ・ 5 件」。
  - 該当行:
    - 高さ 30。1 行目が淡い青（#E1F2FF）、ほかは白と #FAFBFC が交互。
    - 「(株)山田商事」が黄色の地で Bold。
    - 行の間に線が無い。
  - 2〜5 番目のファイル: 地が #F9FAFA。chevron-right で、名前が SemiBold、まとめは灰色の Regular。パスと青い棒は無い。
- H-S（`157_944.png`）:
  - 件数の行が「検索中… 該当 5 件」。
  - 「すべて開く」と［保存］が無く、絞り込みはある。
- H-R（`135_1498.png`）: 件数の行が「16 件（6 ファイル）」。
- H-F（Figma の 352:17353）:
  - 件数の行が「15 件中 3 件を表示」で、欄に打った文字がある。
  - 合う行だけが見える。
- E16（`157_1236.png`）: 件数の行が無く、一覧が列の見出しから始まる。
- H-W・H-PP・H-TX:
  - それぞれ、Word（青の file-text）・PowerPoint（赤橙の presentation）・テキスト（灰色の file）が開いている。
  - 場所が「1 ページ（目安）」「スライド 2」「行 12」「29 行目」で、種別が「本文」。
- 1024×640:
  - 件数の行が重ならない（入らなければ 2 行に折り返す）。
  - 「場所」が「…」で切れ、「該当行」が右端で切れる。
  - 一覧が 120 より低くならない。
- 当たった文字の黄色の地で、行の高さ（30）が変わらない。

## 決めること

1. （決めた、10-04）該当行の「場所」の文字色は Figma のとおり Ink.Body（#5F6368）。diff_round4 を訂正した。
2. （決めた、10-04）行の中の区切りは Figma のとおり「│」（U+2502）。diff_round4 を訂正した。
3. テキストの場所の書き方。.md は「行 12」、.txt は「29 行目」で、そろっていない。
4. ファイルの見出しのまとめの空白。H-PP は「スライド 2・1 件」、ほかは「スライド 2 ・ 1 件」。
5. ファイルの並び順。H と H-W で順が違う。列の見出しを押して並べ替えるかも決まっていない。
6. 種別「コメント」「ノート」の文言と、その場所の書き方。
7. 絞り込みで 0 件になったときの見え方。
8. 件数の上限（今のコードは 10,000 件）を超えたときの文言と、出す場所。
9. 行・ファイルの見出しにマウスを乗せたときの見た目（Figma に無い）。
10. 絞り込みを待つ時間（今のコードは 300ms）をそのままにするか。
11. 縦のスクロールバーの見た目（theme の既定でよいか）。
12. ダブルクリック（と Enter）で開く先。Office のファイルでその場所まで動かすか、ファイルを開くだけか。
