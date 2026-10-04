# 検索バー

## 1. 役割

検索の画面の上に置く欄。次のものを並べる。

- 検索ワードを打つ欄と［検索］（検索中は［中止］）
- 調べるファイルの種類（チップ）
- 見出し「種類」とファイルの種類のチップ、［ファイル内の対象］（2026-10-04 に「探す範囲」から改名。種類のチップ＝どのファイルを探すか、ファイル内の対象＝ファイルの中のどこを探すか、と役目を分けた）
- 大文字・小文字の区別と正規表現の切り替え
- 高速検索を使えるかの表示（バッジ）

**Figma**

- 部品は `xaml/search/search_bar`（348:9775。既定は 340:610）。
- variant:
  - ページ 106: 既定・H0・H・H-B・H-R・H-範囲2・E13・thumb-H-R
  - ページ 148: P-H・P-H-R・P-H-B・P-H-E・P-H-1・P-E12・P-E13・P-E14
- 中身を細かく読んだのは H0（348:9139）・P-E12（348:9774）・H-B（348:9223）。
- 高速検索の表示の一覧は 136:89（`../png/136_89.png`）にある。
- 正規表現の吹き出しは 135:1801（`135_1801.png`）にある。
- ファイル内の対象のメニューは 201:5488（`201_5488.png`）・213:6592・213:6958 にある。

**使う画面と `../png/` の画像**

| 画面 | 画像 |
|---|---|
| H | `108_230.png` |
| H0 | `148_67.png` |
| H-E | `157_654.png` |
| H-E2 | `157_359.png` |
| H-1 | `157_1835.png` |
| H-S | `157_944.png` |
| H-R | `135_1498.png` |
| H-B | `108_1116.png` |
| H-P | `108_1419.png` |
| E12 | `157_2424.png` |
| E13 | `175_4085.png` |
| E14 | `157_1538.png` |
| H-範囲 | `213_6592.png` |
| H-範囲2 | `213_6958.png` |

## 2. 置き場所

- XAML は `scripts/tebunko/xaml/search/search_bar.xaml`。根は `Border`（x:Name `SearchBar`）。
- `search.xaml` の **SearchBarHost**（`SearchRoot` の 1 行目）に差す。
- `XamlReader.Load` で読み、`SearchBarHost.Content` に入れる。
- 画面層は `ui/search/search_bar.ps1`、判断層は `ui/search/search_bar_view.ps1`。
- 判断層は、今の `ui/search_view.ps1` の `newSearchButtonState`・`getWordNotice`・`getFastSearchView`・`describeSearchOption` を移して直す。

## 3. 部品の木

```xml
<Border x:Name="SearchBar" Background="{StaticResource Bg.Surface}"
        BorderBrush="{StaticResource Border.Soft}" BorderThickness="0,0,0,1"
        Padding="16,12,16,10">
  <StackPanel>
    <!-- 1 行目: ラベル・ワード欄・［検索］ -->
    <Grid x:Name="WordRow">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="80"/>     <!-- ラベル -->
        <ColumnDefinition Width="*"/>      <!-- ワード欄 -->
        <ColumnDefinition Width="12"/>     <!-- 間 -->
        <ColumnDefinition Width="100"/>    <!-- ［検索］ -->
      </Grid.ColumnDefinitions>
      <TextBlock x:Name="WordLabel" Grid.Column="0" VerticalAlignment="Center"
                 Style="{StaticResource Body}" Foreground="{StaticResource Ink.Strong}"
                 Text="{search.word.label}"/>
      <Grid Grid.Column="1">
        <TextBox x:Name="WordBox" Height="32" Padding="10,0,4,0" VerticalContentAlignment="Center"
                 Style="{StaticResource TextBox.Base}" FontSize="13"
                 Background="{StaticResource Bg.Surface}" BorderBrush="{StaticResource Border.Input}"
                 BorderThickness="1"/>                                   <!-- 角丸 6 は TextBox.Base の Template で -->
        <TextBlock x:Name="WordPlaceholder" Margin="11,0,0,0" VerticalAlignment="Center"
                   IsHitTestVisible="False" Visibility="Collapsed"
                   Style="{StaticResource Body}" Foreground="#B2BAC4"
                   Text="{search.word.placeholder}"/>                   <!-- #B2BAC4 はキーが無い（決めること） -->
      </Grid>
      <Button x:Name="SearchButton" Grid.Column="3" Height="32"
              Style="{StaticResource Primary}" Padding="16,7"
              Content="{search.button.search}"/>
    </Grid>

    <!-- 正規表現の誤り（E12 だけ） -->
    <StackPanel x:Name="WordErrorRow" Orientation="Horizontal" Margin="92,2,0,0"
                Visibility="Collapsed">
      <Border Width="12" Height="12" CornerRadius="6" Background="{StaticResource Danger.Text}"
              VerticalAlignment="Center">
        <TextBlock Text="!" FontSize="9" FontWeight="Bold" Foreground="{StaticResource Ink.OnAccent}"
                   HorizontalAlignment="Center" VerticalAlignment="Center"/>
      </Border>
      <TextBlock x:Name="WordErrorText" Margin="4,0,0,0" VerticalAlignment="Center"
                 Style="{StaticResource Chip}" Foreground="{StaticResource Danger.Text}"
                 Text="{search.word.error.unclosed}"/>
    </StackPanel>

    <!-- 2 行目: 折り返す行。左の塊と右の塊を DockPanel で両端に寄せ、入らなければ右の塊を次の行へ -->
    <WrapPanel x:Name="OptionRow" Margin="0,8,0,0" Orientation="Horizontal">
      <StackPanel x:Name="LeftOptions" Orientation="Horizontal" Margin="0,0,8,0">
        <TextBlock x:Name="KindLabel" Style="{StaticResource Meta}" Foreground="{StaticResource Ink.Body}"
                   VerticalAlignment="Center" Margin="0,0,8,0" Text="{search.kind.label}"/>
        <ItemsControl x:Name="FileKindChips">
          <ItemsControl.ItemsPanel>
            <ItemsPanelTemplate><StackPanel Orientation="Horizontal"/></ItemsPanelTemplate>
          </ItemsControl.ItemsPanel>
          <!-- 子: ToggleButton Style="FileKindChip"（下の表）。Excel・Word・PowerPoint・テキストの 4 つ。間 6 -->
        </ItemsControl>
        <Button x:Name="ScopeButton" Margin="8,0,0,0" Height="24" MinWidth="80"
                Padding="10,0,8,0" Style="{StaticResource ToolButton}">
          <StackPanel Orientation="Horizontal">
            <TextBlock x:Name="ScopeButtonText" Style="{StaticResource Meta}"
                       Foreground="{StaticResource Ink.Strong}" VerticalAlignment="Center"
                       Text="{search.scope.button}"/>
            <Viewbox Width="10" Height="10" Margin="4,0,0,0"><Path Data="{StaticResource Icon.ChevronDown}"
                     Stroke="{StaticResource Ink.Strong}" StrokeThickness="2.4"/></Viewbox>
          </StackPanel>
        </Button>
      </StackPanel>
      <StackPanel x:Name="RightOptions" Orientation="Horizontal">
        <CheckBox x:Name="CaseCheck" Height="24" Padding="4,0,6,0" Style="{StaticResource Choice.Small}"
                  Content="{search.case}"/>
        <CheckBox x:Name="RegexCheck" Height="24" Margin="8,0,0,0" Padding="4,0,6,0"
                  Style="{StaticResource Choice.Small}" Content="{search.regex}"
                  ToolTip="{search.regex.tooltip}"/>
        <Border x:Name="FastSearchText" Margin="8,0,0,0" Padding="8,3" CornerRadius="11"
                VerticalAlignment="Center" Background="{StaticResource Ok.Soft}">
          <StackPanel Orientation="Horizontal">
            <Viewbox Width="12" Height="12"><Path x:Name="FastSearchIcon" Data="{StaticResource Icon.Zap}"
                     Stroke="{StaticResource Ok}" StrokeThickness="2"/></Viewbox>
            <TextBlock x:Name="FastSearchLabel" Margin="4,0,0,0" Style="{StaticResource Chip}"
                       Foreground="{StaticResource Ok}" VerticalAlignment="Center"
                       Text="{search.fast.ok}"/>
            <Viewbox x:Name="FastSearchInfo" Width="13" Height="13" Margin="4,0,0,0"
                     Visibility="Collapsed"><Path Data="{StaticResource Icon.Info}"
                     Stroke="{StaticResource Ok}" StrokeThickness="2"/></Viewbox>
          </StackPanel>
        </Border>
      </StackPanel>
    </WrapPanel>
  </StackPanel>
</Border>
```

**2 行目の並べ方**

- Figma の 2 行目は、次の順に並べた折り返しの行（gap 8）になっている。
  1. チップ
  2. ［ファイル内の対象］
  3. 空き（Fill）
  4. 切り替えの 2 つとバッジ
- WPF の WrapPanel は「空き Fill」を持てない。そこで画面層の `SizeChanged` で、`RightOptions` の左の Margin を計算して右端に寄せる。
  - 計算: 行の幅 − 左の塊の幅 − 右の塊の幅。
  - 右の塊が入らないときは Margin を 0 にし、次の行の左から並べる。
- 1024 の窓では、バッジだけが 3 行目に落ちる（`figma_wpf_map.md` の「検索バー」: 幅 804 で高さ 86 → 114）。
  - これを出すには、`RightOptions` を 1 つの塊にせず、`CaseCheck`・`RegexCheck`・`FastSearchText` を WrapPanel の子として並べる。
  - そのうえで、先頭の `CaseCheck` にだけ、右寄せの Margin を付ける（下の「8. リサイズ」）。

**種別のチップ（Style `FileKindChip`、ToggleButton）**

```xml
<Border Height="24" Padding="8,0,10,0" CornerRadius="12" BorderThickness="1"
        Background="{StaticResource Accent.Soft}" BorderBrush="{StaticResource Accent}">
  <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
    <Viewbox Width="10" Height="10"><Path Data="{StaticResource Icon.Check}"
             Stroke="{StaticResource Accent.Hover}" StrokeThickness="2"/></Viewbox>  <!-- 線 = 10 × 2/24 を 24 の座標で 2 -->
    <TextBlock Margin="4,0,0,0" Style="{StaticResource Chip}" Foreground="{StaticResource Accent.Hover}"
               Text="{search.kind.excel}"/>
  </StackPanel>
</Border>
```

## 4. 寸法と色

| 要素 | 大きさ | 余白 | 地の色 | 枠 | 文字の Style | 文字の色 | 角丸 | アイコン |
|---|---|---|---|---|---|---|---|---|
| SearchBar | 横 Fill × Hug（86。E12 は 112。幅 804 では 114） | 上 12・左右 16・下 10、行の間 8 | Bg.Surface（#FFFFFF） | 下 1 Border.Soft（#E0E2E5） | – | – | 0 | – |
| WordLabel | 80 × Hug | – | – | – | Body（13 Regular） | Ink.Strong（#202124）。H0 は #B2BAC4（キーが無い） | – | – |
| WordBox | Fill × 32 | 左 10・右 4 | Bg.Surface | 1 Border.Input（#D1D1D1）。E12 は 1 Danger.Dot（#D93025） | 13 Regular（Body） | Ink.Strong | 6 | – |
| WordPlaceholder | – | 左 11 | – | – | Body | #B2BAC4（キーが無い） | – | – |
| SearchButton（使える） | 100 × 32 | 16/7 | Accent（#0078D4） | – | 13 SemiBold（Heading） | Ink.OnAccent（#FFFFFF） | 6 | – |
| SearchButton（中止） | 100 × 32 | 16/7 | Accent（#0078D4） | – | 13 SemiBold | Ink.OnAccent | 6 | – |
| SearchButton（使えない・H0） | 100 × 32 | 16/7 | Border.Normal（#D9DEE3） | – | 13 **Bold** | Ink.Faint（#99A1AB） | 6 | – |
| SearchButton（使えない・E12・H-E） | 100 × 32 | 16/7 | `#E8EAED`（キーが無い） | – | 13 SemiBold | Ink.Placeholder（#9AA0A6） | 6 | – |
| WordErrorRow | Hug | 左 92・上下 2、間 4 | – | – | – | – | – | 12 の丸（Danger.Text #D13438）に白の「!」9 Bold |
| WordErrorText | Hug | – | – | – | Chip（11 Medium） | Danger.Text（#D13438） | – | – |
| OptionRow | 横 Fill × Hug（24。折り返すと 24＋8＋24） | 上 8、子の間 8 | – | – | – | – | – | – |
| FileKindChips の各チップ | Hug × 24 | 左 8・右 10、中の間 4、チップの間 6 | Accent.Soft（#E5F1FB） | 1 Accent（#0078D4） | Chip（11 Medium） | Accent.Hover（#0B5CAD） | 12 | check 10・線 0.83 |
| ScopeButton | 最小 80 × 24（文字に合わせて伸びる） | 左 10・右 8、中の間 4 | Bg.Surface | 1 Border.Input（#D1D1D1） | Meta（11 Regular） | Ink.Strong（#202124） | 4 | chevron-down 10・線 0.83 |
| CaseCheck／RegexCheck | Hug × 24 | 左 4・右 6、箱と文字の間 5、2 つの間 8 | – | 箱 12×12・1 Border.Check（#9EA3AB） | Meta（11 Regular） | Ink.Body（#5F6368。Figma の #5E6369 をまとめた） | 箱 2 | – |
| RegexCheck（オン） | 同上 | 同上 | 箱 Accent（#0078D4） | – | **11 Medium** | Accent.Hover（#0B5CAD。Figma は #0A5CAD をまとめた） | 箱 2 | check（白） |
| FastSearchText（使用可） | Hug × Hug（約 22） | 8/3、中の間 4 | Ok.Soft（#E0F7E0） | 無し | Chip（11 Medium） | Ok（#218A21） | 11 | zap 12・線 1.0、info 13・丸の線 1.08、中の「i」は線 1.5（README の 9）（Ok） |
| FastSearchText（一部で使用可） | 同上 | 同上 | Warn.Soft（#FFF5E0） | 無し | Chip | Warn（#BA7D00） | 11 | zap 12・info 13（Warn） |
| FastSearchText（使用不可） | 同上 | 同上 | Bg.Tag（#F1F3F4） | 無し | Chip | Ink.Body（#5F6368） | 11 | zap 12・info 13（Ink.Body）。info はどの理由でも出す |

- 高速検索のバッジとチップのアイコンは Lucide（lucide-static 1.50.0）を使う。描き方は Viewbox＋Path で、線の太さは 24 の座標で 2（表示の大きさ × 2/24 になる）。
- **バッジに枠を付けない**（diff_round2 の 7・diff_round3 の 7）。
- **チップにファイルの種類のアイコンを付けない。** チェックと文言だけにする（`figma_wpf_map.md`）。

## 5. 状態ごとの見え方

| 状態 | WordLabel | WordBox | 薄い文字 | SearchButton | 誤りの行 | チップ・ファイル内の対象 | 大文字・小文字／正規表現 | 高速検索のバッジ |
|---|---|---|---|---|---|---|---|---|
| H0 | #B2BAC4 | 空 | 出す `{search.word.placeholder}` | 使えない（H0 の色）`{search.button.search}` | 出さない | 出す（使える） | 出す（オフ） | **出さない** |
| H-E | Ink.Strong | 空 | **出さない** | 使えない（E12 の色） | 出さない | 出す | 出す | 出す `{search.fast.ok}` |
| H-E2 | Ink.Strong | 「(株)山田商事」 | 出さない | 使える | 出さない | 出す | 出す | `{search.fast.ok}` |
| H-1 | Ink.Strong | 「山」 | 出さない | 使える | 出さない | 出す | 出す | `{search.fast.short}` |
| H-S | Ink.Strong | 「(株)山田商事」 | 出さない | **`{search.button.stop}`（Accent の地）** | 出さない | 出す | 出す | `{search.fast.ok}` |
| H | Ink.Strong | 「(株)山田商事」 | 出さない | 使える | 出さない | 出す | 出す | `{search.fast.ok}` |
| H-R | Ink.Strong | 「(株)山田商事」 | 出さない | 使える | 出さない | 出す | 正規表現オン | `{search.fast.regex}` |
| H-B・H-P | Ink.Strong | 「(株)山田商事」 | 出さない | 使える | 出さない | 出す | 出す | **出さない** |
| E12 | Ink.Strong | 「(株)山田(商事」・**枠 Danger.Dot** | 出さない | **使えない**（E12 の色） | **出す** | 出す | 正規表現オン | `{search.fast.regex}` |
| E13 | Ink.Strong | 「(株)山田商事」 | 出さない | **使えない** | 出さない | 出す（**ファイル内の対象は使える**） | 出す | `{search.fast.ok}` |
| E14 | Ink.Strong | 「(株)山田商店」 | 出さない | 使える | 出さない | 出す | 正規表現オン | `{search.fast.regex}` |
| H-範囲2 | Ink.Strong | 「(株)山田商事」 | 出さない | 使える | 出さない | ［ファイル内の対象］の文言が `{search.scope.button.partial}` になり幅が伸びる | 出す | `{search.fast.ok}` |

- **H0 でワード欄に打てるか**（IsEnabled）は Figma で決まっていない → 決めること。
  - Figma の H0 は、ラベルと薄い文字が灰色で、チップ・切り替えは普通の色になっている。
- **E13 で［ファイル内の対象］を使えなくしない**（diff_round2 の E13。PoC が無効にした）。使えなくするのは［検索］だけ。
- **E12 は前の版の「文字どおり検索します」をやめ、［検索］を使えなくして誤りを出す。**
  - 今のコードの `getWordNotice` は「正規表現として不正なため、文字どおり検索します。」を返す。これは Figma と違う。
- 検索バーの高さ:
  - E12 だけ誤りの行の分だけ高く、112 になる（Figma の値）。
  - ほかの状態は 86（窓 1280 のとき）。

### 高速検索のバッジ（上から順に、最初に当てはまるもの）

| 順 | 条件（`getFastSearchView` の入力） | 文言の ID | 地・文字の色 | info | ツールチップ |
|---|---|---|---|---|---|
| 0 | `IndexState` が `'none'`・`'updating'`・`'interrupted'`（H0・H-B・H-P） | – | **バッジを出さない** | – | – |
| 1 | `Available = $false`（Windows Search に接続できない） | `search.fast.nows` | Bg.Tag／Ink.Body | 出す | `{search.fast.nows.tip}` |
| 2 | `UseRegex = $true` | `search.fast.regex` | Bg.Tag／Ink.Body | 出す | `{search.fast.regex.tip}` |
| 3 | ワードが 1 文字（空は当てはめない） | `search.fast.short` | Bg.Tag／Ink.Body | **出さない** | 無し |
| 4 | `Pending = $true`（反映を待っているフォルダがある） | `search.fast.pending` | Bg.Tag／Ink.Body | 出す | `{search.fast.pending.tip}` |
| 5 | `LongPathFolders -gt 0`（パスが長く、通常の検索で調べるフォルダがある） | `search.fast.partial` | Warn.Soft／Warn | 出す | `{search.fast.partial.tip}` |
| 6 | それ以外（ワードが空のときも） | `search.fast.ok` | Ok.Soft／Ok | **出さない** | 無し |

- ワードが空のときも「使用可」を出す（Figma の H-E）。今の `testFastSearchUsable` は空を「使用不可」にするので、判断層で空を先に外す。
- ツールチップの見た目は overlays.md に書く。
  - Figma の値: 地 `#333840`（キーが無い）、10/7、角丸 5、影 0 2 8 rgba(0,0,0,.15)、文字 11 白。
- ツールチップの中の［インデックス管理で確認］はリンクで、押すとインデックス管理を開く。
  - ツールチップの中のリンクを押せるようにするには、ToolTip ではなく Popup にする必要がある（WPF の ToolTip は押せない）。どちらにするかは overlays.md と合わせて決めること。

## 6. 文言

| 文言の ID | 全文 | 使う状態 |
|---|---|---|
| `search.word.label` | 検索ワード | すべて |
| `search.word.placeholder` | 検索ワードを入力 | H0 |
| `search.button.search` | 検索 | 検索していないとき |
| `search.button.stop` | 中止 | H-S |
| `search.word.error.unclosed` | 正規表現が正しくありません（閉じていない括弧があります） | E12（閉じていない括弧のとき） |
| `search.kind.excel` | Excel | すべて |
| `search.kind.word` | Word | すべて |
| `search.kind.powerpoint` | PowerPoint | すべて |
| `search.kind.text` | テキスト | すべて |
| `search.kind.label` | 種類 | すべて。チップの左の見出し（Meta 11・Ink.Body #5F6368。チップとの間 8） |
| `search.scope.button` | ファイル内の対象 | 既定（すべてオン）から変えていないとき |
| `search.scope.button.partial` | ファイル内の対象・{既定から変えた項目の数} 件変更 | 例「ファイル内の対象・2 件変更」（H-範囲2。コメント・ノートをオフにした）。文字は Medium・#0B5CAD、枠は Accent #0078D4 |
| `search.scope.body` | 本文　（常に対象） | ファイル内の対象のメニュー（使えない・常にオン）。「本文」と「（」の間は全角の空白 |
| `search.scope.shape` | 図形 | ファイル内の対象のメニュー |
| `search.scope.comment` | コメント | ファイル内の対象のメニュー |
| `search.scope.note` | ノート | メニューの「PowerPoint」の組 |
| `search.scope.group.common` | 共通 | メニューの組の見出し（本文・図形・コメント） |
| `search.scope.group.powerpoint` | PowerPoint | メニューの組の見出し（ノート） |
| `search.scope.reset` | 既定に戻す | メニューの下（区切りの線の下）。既定のままなら灰色（#9AA0A6）で押せない。変えていたら Accent #0078D4・下線で押せる |
| `search.case` | 大文字・小文字を区別 | すべて |
| `search.regex` | 正規表現 | すべて |
| `search.regex.tooltip` | 正規表現を使う<br>オンにすると高速検索は使えません | 正規表現にマウスを乗せたとき（2 行） |
| `search.fast.ok` | 高速検索：使用可 | 上の表の 6 |
| `search.fast.partial` | 高速検索：一部で使用可 | 5 |
| `search.fast.partial.tip` | {フォルダ数} フォルダはパスが長いため、通常の検索で調べます（検索結果は変わりません） | 5。例「2 フォルダは…」 |
| `search.fast.nows` | 高速検索：使用不可 | 1 |
| `search.fast.nows.tip` | Windows Search に接続できません。検索はできますが時間がかかります　［インデックス管理で確認］ | 1。［ ］の前は全角の空白。［インデックス管理で確認］はリンク |
| `search.fast.regex` | 高速検索：使用不可 | 2 |
| `search.fast.regex.tip` | 正規表現では使えません。正規表現をオフにすると速く検索できます | 2 |
| `search.fast.short` | 高速検索：使用不可 | 3 |
| `search.fast.short.tip` | 2 文字以上で使えます | 3 |
| `search.fast.pending` | 高速検索：使用不可 | 4 |
| `search.fast.pending.tip` | 反映待ちです。インデックスの更新が終わると使えます　［インデックス管理で確認］ | 4 |

- **使用不可のバッジに理由のカッコ書きは付けない（2026-10-04 に決定）。** どの理由でも文言は「高速検索：使用不可」だけにし、理由は info の吹き出し（ツールチップ）の先頭に書く。2 文字未満のときも info を出す。
- 「：」は全角。かっこは全角の「（）」。数と単位の間（「2 文字」「2 フォルダ」）は半角の空白。
- キー操作のヒント（「Enter で検索」など）は書かない・出さない。
- ファイル内の対象のメニューの見た目は overlays.md に書く。ここでは文言とボタンの文言だけを決める。

## 7. 操作と振る舞い

| きっかけ | 何が起きるか | 呼ぶ関数 | 次の状態 |
|---|---|---|---|
| WordBox に文字を打つ | ［検索］の可否・誤りの行・バッジを決め直す。H0 の薄い文字は空のときだけ出す | `newSearchButtonState`・`getWordError`・`getFastSearchView` → `updateSearchScreen` | H-E・H-1・H-E2・E12 |
| WordBox で Enter | ［検索］と同じ（使えるときだけ） | `startSearch` | H-S |
| ［検索］をクリック | 検索を始める。ボタンは［中止］になる | `startSearch` | H-S |
| ［中止］をクリック | 中止を頼む。止まるまでボタンを使えなくする（`Stopping`） | `stopSearch` | E16 |
| チップをクリック | その種類のオン・オフを切り替える。チップの見た目は Accent.Soft の地・check（オン）と、オフの見た目（Figma に無い → 決めること） | `setFileKind` | 変わらない |
| ［ファイル内の対象］をクリック | メニューをボタンの左下に開く（overlays.md）。図形・コメント・ノートを切り替える。本文は切り替えられない | `getScopeButtonText` | 変わらない（ボタンの文言と幅だけ変わる） |
| CaseCheck を切り替える | 大文字・小文字の区別を切り替える | – | 変わらない |
| RegexCheck を切り替える | 正規表現を切り替える。バッジ・誤りの行を決め直す | `getWordError`・`getFastSearchView` | E12 か元の状態 |
| RegexCheck にマウスを乗せる | ツールチップ `{search.regex.tooltip}` を出す | – | – |
| バッジにマウスを乗せる | info があるときだけ、ツールチップを出す | `getFastSearchView` の `Tooltip` | – |
| ツールチップの［インデックス管理で確認］ | インデックス管理を開く | `selectNav 'Index'` | インデックス管理の画面 |

- **ホバー・押した・フォーカス**
  - ［検索］: theme の Primary を使う（ホバーは Accent.Hover #0B5CAD）。Figma に別の定めは無い。
  - ［ファイル内の対象］: ToolButton を使う（ホバーは Bg.Hover）。Figma に別の定めは無い。
  - フォーカスの枠: theme の既定を使う。
- **Tab の順**: WordBox → SearchButton → チップ 4 つ → ScopeButton → CaseCheck → RegexCheck。バッジにはフォーカスを当てない。
- **ツールチップ**: 正規表現とバッジ（info があるとき）だけに付ける。チップ・［ファイル内の対象］には付けない（Figma に無い）。

## 8. リサイズ

- 横は 1 行目のワード欄だけが伸びる（`*`）。ラベル 80・間 12・［検索］100 は固定。
- 2 行目は WrapPanel で折り返す。行の間は 8。
  - 左寄せ: チップ・［ファイル内の対象］。
  - 右寄せ: 大文字・小文字・正規表現・バッジ。
  - 入らないときは、後ろから順に次の行の左へ落とす。
- 画面層の `OptionRow.SizeChanged` での動き:
  1. 左の塊（チップと［ファイル内の対象］）・`CaseCheck`・`RegexCheck`・`FastSearchText` の DesiredSize を測る。
  2. 1 行に入るなら、`CaseCheck.Margin.Left` を「行の幅 − ほかの幅の合計」にして右へ寄せる。
  3. 入らないなら、入るところまでを 1 行目の右に寄せ、残りを 2 行目の左から並べる。
- 窓ごとの見え方:
  - 1280×820（主な領域 1060）: 2 行目は 1 行で、高さ 86。
  - 1024×640（主な領域 804）:
    - Figma では、バッジ「高速検索：使用可」が 3 行目に落ちて、高さ 114。
    - PoC では、部品が少し細くて 2 行目に収まった。WrapPanel で折り返す作りになっていれば、これを許す（diff_round4 の「許」）。
  - 1600×1000（主な領域 1380）: 1 行に収まり、間が広がる。
- 文字は切らない（ワード欄は横にスクロールする TextBox）。
- 誤りの行は 1 行。1024 で入りきらないときの扱いは Figma に無い → 決めること。

## 9. 判断層

`ui/search/search_bar_view.ps1`。WPF の型に触らない。

### newSearchButtonState

今の関数に `wordValid` を足す。

入力:

| 引数 | 型 |
|---|---|
| searching | bool |
| stopping | bool |
| word | string |
| hasIndex | bool |
| targetCount | int |
| wordValid | bool |

`wordValid` は、正規表現がオフなら `$true`、オンなら `isValidRegex`。

出力: `@{ Content = string; Enabled = bool; Look = 'Primary' | 'DisabledNoIndex' | 'Disabled' }`。

- `DisabledNoIndex` は H0 の色（Border.Normal の地・Ink.Faint・Bold）。
- `Disabled` は E12・H-E・E13 の色（`#E8EAED`・Ink.Placeholder）。
- 色を 2 つに分けるかは決めること。Figma が 2 通りあるので、そのとおりに分けておく。

```powershell
It "<Name>" -TestCases @(
  @{ Name="H0";          Searching=$false; Stopping=$false; Word="";              HasIndex=$false; Target=3; Valid=$true;  Content="検索"; Enabled=$false; Look="DisabledNoIndex" }
  @{ Name="H-E";         Searching=$false; Stopping=$false; Word="";              HasIndex=$true;  Target=3; Valid=$true;  Content="検索"; Enabled=$false; Look="Disabled" }
  @{ Name="H-E2";        Searching=$false; Stopping=$false; Word="(株)山田商事";  HasIndex=$true;  Target=3; Valid=$true;  Content="検索"; Enabled=$true;  Look="Primary" }
  @{ Name="H-1";         Searching=$false; Stopping=$false; Word="山";            HasIndex=$true;  Target=3; Valid=$true;  Content="検索"; Enabled=$true;  Look="Primary" }
  @{ Name="E12";         Searching=$false; Stopping=$false; Word="(株)山田(商事"; HasIndex=$true;  Target=3; Valid=$false; Content="検索"; Enabled=$false; Look="Disabled" }
  @{ Name="E13";         Searching=$false; Stopping=$false; Word="(株)山田商事";  HasIndex=$true;  Target=0; Valid=$true;  Content="検索"; Enabled=$false; Look="Disabled" }
  @{ Name="H-S";         Searching=$true;  Stopping=$false; Word="(株)山田商事";  HasIndex=$true;  Target=3; Valid=$true;  Content="中止"; Enabled=$true;  Look="Primary" }
  @{ Name="中止を頼んだ"; Searching=$true;  Stopping=$true;  Word="(株)山田商事";  HasIndex=$true;  Target=3; Valid=$true;  Content="中止"; Enabled=$false; Look="Disabled" }
) {
  param($Searching,$Stopping,$Word,$HasIndex,$Target,$Valid,$Content,$Enabled,$Look)
  $s = newSearchButtonState $Searching $Stopping $Word $HasIndex $Target $Valid
  $s.Content | Should -Be $Content; $s.Enabled | Should -Be $Enabled; $s.Look | Should -Be $Look
}
```

### getWordError

今の `getWordNotice` を置き換える。

入力: `word`（string）・`useRegex`（bool）。

出力: `@{ Visible = bool; Text = string }`。

```powershell
It "<Name>" -TestCases @(
  @{ Name="正規表現オフ";       Word="(株)山田(商事"; Regex=$false; Visible=$false; Text="" }
  @{ Name="空";                 Word="";              Regex=$true;  Visible=$false; Text="" }
  @{ Name="正しい";             Word="山田.*商事";    Regex=$true;  Visible=$false; Text="" }
  @{ Name="閉じていない括弧";   Word="(株)山田(商事"; Regex=$true;  Visible=$true;  Text="正規表現が正しくありません（閉じていない括弧があります）" }
) { param($Word,$Regex,$Visible,$Text) $e = getWordError $Word $Regex; $e.Visible | Should -Be $Visible; $e.Text | Should -Be $Text }
```

- 閉じていない括弧の外の誤り（`[` の閉じ忘れ・`*` で始まるなど）の文言は、Figma に無い → 決めること。
- 決まるまでは、どの誤りでも「正規表現が正しくありません」だけを出す案にする。

### getFastSearchView

今の関数を広げる。

入力:

| 引数 | 型・値 |
|---|---|
| available | `$null`・`$true`・`$false` |
| useRegex | bool |
| word | string |
| indexState | `'none'`・`'ready'`・`'updating'`・`'interrupted'` |
| pending | bool |
| longPathFolders | int |

出力:

| キー | 型・値 |
|---|---|
| Visible | bool |
| Usable | bool |
| Tone | `'Ok'`・`'Warn'`・`'Off'` |
| Text | string |
| ShowInfo | bool |
| Tooltip | string |
| TooltipLink | bool。［インデックス管理で確認］を付けるか |

- `Usable` は、件数の文言には使わない（2026-10-04。件数の行に方式を出さない）。
- 「一部で使用可」は `Usable = $true` とする。

```powershell
It "<Name>" -TestCases @(
  @{ Name="H0";        A=$true;  R=$false; W="";             S='none';        P=$false; L=0; Visible=$false; Tone='';     Text="";                                                   Info=$false; Tip="" }
  @{ Name="H-B";       A=$true;  R=$false; W="(株)山田商事"; S='updating';    P=$false; L=0; Visible=$false; Tone='';     Text="";                                                   Info=$false; Tip="" }
  @{ Name="H-P";       A=$true;  R=$false; W="(株)山田商事"; S='interrupted'; P=$false; L=0; Visible=$false; Tone='';     Text="";                                                   Info=$false; Tip="" }
  @{ Name="接続不可";  A=$false; R=$true;  W="山";           S='ready';       P=$true;  L=2; Visible=$true;  Tone='Off';  Text="高速検索：使用不可"; Info=$true;  Tip="Windows Search に接続できません。検索はできますが時間がかかります" }
  @{ Name="正規表現";  A=$true;  R=$true;  W="山";           S='ready';       P=$true;  L=2; Visible=$true;  Tone='Off';  Text="高速検索：使用不可"; Info=$true;  Tip="正規表現では使えません。正規表現をオフにすると速く検索できます" }
  @{ Name="1 文字";    A=$true;  R=$false; W="山";           S='ready';       P=$true;  L=2; Visible=$true;  Tone='Off';  Text="高速検索：使用不可"; Info=$true;  Tip="2 文字以上で使えます" }
  @{ Name="反映待ち";  A=$true;  R=$false; W="(株)山田商事"; S='ready';       P=$true;  L=2; Visible=$true;  Tone='Off';  Text="高速検索：使用不可"; Info=$true;  Tip="反映待ちです。インデックスの更新が終わると使えます" }
  @{ Name="一部";      A=$true;  R=$false; W="(株)山田商事"; S='ready';       P=$false; L=2; Visible=$true;  Tone='Warn'; Text="高速検索：一部で使用可";                              Info=$true;  Tip="2 フォルダはパスが長いため、通常の検索で調べます（検索結果は変わりません）" }
  @{ Name="使用可";    A=$true;  R=$false; W="(株)山田商事"; S='ready';       P=$false; L=0; Visible=$true;  Tone='Ok';   Text="高速検索：使用可";                                    Info=$false; Tip="" }
  @{ Name="空・H-E";   A=$true;  R=$false; W="";             S='ready';       P=$false; L=0; Visible=$true;  Tone='Ok';   Text="高速検索：使用可";                                    Info=$false; Tip="" }
  @{ Name="未確認";    A=$null;  R=$false; W="(株)山田商事"; S='ready';       P=$false; L=0; Visible=$true;  Tone='Ok';   Text="高速検索：使用可";                                    Info=$false; Tip="" }
) {
  param($A,$R,$W,$S,$P,$L,$Visible,$Tone,$Text,$Info,$Tip)
  $v = getFastSearchView $A $R $W $S $P $L
  $v.Visible | Should -Be $Visible
  if ($Visible) { $v.Tone | Should -Be $Tone; $v.Text | Should -Be $Text; $v.ShowInfo | Should -Be $Info; $v.Tooltip | Should -Be $Tip }
}
```

- `Tooltip` には［インデックス管理で確認］を含めない。リンクの有無は `TooltipLink`（接続不可・反映待ちで `$true`）で返し、画面層が付ける。

### getScopeButtonText

入力: `includeShapes`・`includeComments`・`includeNotes`（bool）。

出力: string。

```powershell
It "<Name>" -TestCases @(
  @{ Name="すべてオン";   S=$true;  C=$true;  N=$true;  Expected="ファイル内の対象" }
  @{ Name="図形だけ";     S=$true;  C=$false; N=$false; Expected="ファイル内の対象・2 件変更" }
) { param($S,$C,$N,$Expected) getScopeButtonText $S $C $N | Should -Be $Expected }
```

- Figma にあるのは、上の 2 通り（既定と H-範囲2）だけ。次は決めること。
  - 数えるのは、既定（オン）から変えた項目の数。図形・コメント・ノートをすべてオフにしたら「・3 件変更」。

### describeSearchOption

ステータスバー・保存のための今の関数。検索バーの見た目には使わない。今のまま残す。

## 10. 画面層

`ui/search/search_bar.ps1`。

- `initSearchBar` で次の配線をする。
  - `WordBox.Add_TextChanged` → `onSearchConditionChanged`
  - `WordBox.Add_KeyDown`（Enter）→ `SearchButton` を押したのと同じ
  - `SearchButton.Add_Click`
    - `newSearchButtonState` の Content が「中止」なら `stopSearch`
    - そうでなければ `startSearch`
  - チップ（`FileKindChips` の各 ToggleButton）の `Checked`・`Unchecked` → `setFileKind`
  - `ScopeButton.Add_Click` → ファイル内の対象のメニューを開く（overlays.md の `openScopeMenu`。PlacementTarget は ScopeButton、左下に付ける）
    - メニューが閉じたら `ScopeButtonText.Text = getScopeButtonText ...`
  - `CaseCheck`・`RegexCheck` の `Checked`・`Unchecked` → `onSearchConditionChanged`
  - `OptionRow.Add_SizeChanged` → `arrangeOptionRow`（上の「8. リサイズ」の右寄せ）
- `onSearchConditionChanged` は次を写してから、`updateSearchScreen`（search.md）を呼ぶ。

  | 判断層の出力 | 写す先 |
  |---|---|
  | `newSearchButtonState` | `SearchButton.Content`・`IsEnabled` と、Look から Style か地・文字の色 |
  | `getWordError` | `WordErrorRow.Visibility`・`WordErrorText.Text`・`WordBox.BorderBrush`（誤りのとき Danger.Dot） |
  | `getFastSearchView` | `setFastSearchView`（下） |

  - H0 の `WordLabel.Foreground`・`WordPlaceholder.Visibility` は、`HasIndex` から決める。
- `setFastSearchView $view` で次を写す。
  - `FastSearchText.Visibility` は Visible から。
  - Tone から、地（Ok.Soft・Warn.Soft・Bg.Tag）と文字・アイコンの色（Ok・Warn・Ink.Body）を決める。
  - `FastSearchLabel.Text` に Text を入れる。
  - `FastSearchInfo.Visibility` は ShowInfo から。
  - `FastSearchText.ToolTip` には Tooltip を入れる（空なら `$null`）。
- Windows Search に接続できるかの確認（`testWindowsSearch`）は重いので、別のスレッドで行う。
  - 結果は `Dispatcher.BeginInvoke` で戻し、`available` を入れて `onSearchConditionChanged` を呼ぶ。
  - 確認が終わるまでは `$null`（使えるものとして扱う）にする。
- 前の版の x:Name のうち、Figma に無いものは置かない。

  | 前の x:Name | 移り先 |
  |---|---|
  | `ShapeCheck`・`CommentCheck` | ファイル内の対象のメニュー |
  | `FileFilterBox` | チップ |
  | `WordNotice` | WordErrorRow |
  | `GoIndexTabButton` | 空の状態のボタン（search.md） |
  | `SearchProgress` | 件数の行の文言（result_list.md） |

## 11. 受け入れ

窓 1280×820 で `../png/` と見比べる（主な領域の左上が x 220・y 32）。

- H（`108_230.png`）:
  - ラベル「検索ワード」の左が x 236。
  - ワード欄が y 44〜76（高さ 32）で、右端が x 1152。
  - ［検索］が x 1164〜1264（幅 100）・高さ 32。
  - 2 行目のチップの上端が y 84、高さ 24。
  - 検索バーの下の線が y 118。
- H（同じ画像）の 2 行目:
  - チップ 4 つの地が淡い青で、枠が青、文字の前に check がある。
  - ［ファイル内の対象］は白地に灰色の枠で、chevron-down がある。
  - 右端のバッジは緑の地で「高速検索：使用可」、右に info が無い。
- H0（`148_67.png`）:
  - ラベルと「検索ワードを入力」が淡い灰色。
  - ［検索］の地が灰色（#D9DEE3）で、文字が太い灰色。
  - バッジが無い。
- H-E（`157_654.png`）:
  - ワード欄が空で、薄い文字が無い。
  - ［検索］が淡い灰色（#E8EAED）。
  - バッジ「使用可」がある。
- H-1（`157_1835.png`）: バッジが灰色の地で「高速検索：使用不可」。info があり、乗せると「2 文字以上で使えます」。
- H-S（`157_944.png`）: ボタンが青の地に「中止」。
- E12（`157_2424.png`）:
  - ワード欄の枠が赤。
  - 欄の下（x 328 から）に赤丸の「!」と「正規表現が正しくありません（閉じていない括弧があります）」がある。
  - ［検索］が灰色。
  - 正規表現のチェックが青で、文字が青の Medium。
  - バッジが灰色の地で、info がある。
  - 検索バーの下の線が y 144（高さ 112）。
- H-B・H-P（`108_1116.png`・`108_1419.png`）: バッジが無く、正規表現のチェックが右端に寄っている。
- H-範囲2（`213_6958.png`）: ボタンが「ファイル内の対象・2 件変更」（青い字・青い枠）で、幅が伸びている。
- 1024×640: 2 行目が折り返し、検索バーの高さが 114 になる（バッジが 3 行目に落ちる。落ちなくても、重ならなければ許す）。

## 決めること

1. 使えない［検索］の色が 2 通りある。
   - H0: 地 Border.Normal（#D9DEE3）・文字 Ink.Faint（#99A1AB）・Bold
   - E12・H-E: 地 `#E8EAED`・文字 Ink.Placeholder（#9AA0A6）・SemiBold
   - diff_round3 は `#E8EAED`・`#9FA4AA` と書いている。
   - 1 つにまとめるか、H0 だけ別にするか。
2. キーの無い色を足す案（`../figma_wpf_map.md` の「食い違い」へ）。
   - `#B2BAC4`（H0 のラベルと薄い文字）
   - `#E8EAED`（使えない［検索］の地）
   - `#333840`（ツールチップの地。overlays.md と共通）
3. H0 でワード欄・チップ・切り替えを使えるようにしておくか（IsEnabled）。
4. 「検索ワードを入力」の薄い文字を、H0 のほかの空のとき（H-E）にも出すか。Figma の H-E は出していない。
5. 正規表現の誤りのうち、閉じていない括弧の外の誤りの文言。
6. 誤りの行が 1024 で入りきらないときの扱い（折り返すか・「…」で切るか）。
7. ［ファイル内の対象］の文言で、コメント・ノートを選んだときの並べ方とノートの短い名前、すべてオフのときの文言。
8. ファイル内の対象のメニューの寸法（幅・行の高さ・余白）。Figma の値をまだ読んでいない（overlays.md の担当と合わせる）。
9. 種別のチップをオフにしたときの見た目（Figma はすべてオンの例だけ）。すべてオフにできるか。
10. バッジのツールチップの［インデックス管理で確認］を押せるようにするか（押せるなら ToolTip ではなく Popup にする）。
11. H-P（前回の更新が途中）でバッジを出さない理由。Figma のとおりにするが、「反映待ち」の表示と使い分けを確かめる。
