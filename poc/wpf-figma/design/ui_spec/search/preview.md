# プレビュー（Excel・Word・PowerPoint・テキストと開くメニュー）

## 1. 役割

検索の画面の下の欄に、結果の一覧で選んだ行の **前後の中身** を出す。上のツールバーから元のファイルを開いたり、フォルダを開いたりできる。

ファイルの種類ごとに、中身の出し方を変える。

| 種類 | 出し方 | Figma |
|---|---|---|
| Excel | 罫線の格子。行番号・列の見出し・一致したセルの地 | 340:954（既定）・353:33663（row3）・353:33741（row4） |
| Word | 段落（折り返す）。一致した語は太字 | 353:33429（H-W） |
| PowerPoint | 段落（折り返す）。一致した語に印 | 427:228（H-PP） |
| テキスト | 行番号と 1 行ずつの文。選んだ行の地を変える。一致した語に印 | 427:284（H-TX） |

ほかに、次の 3 つの状態を持つ。

- 未選択（324:8558）
- 読めない（324:8672）
- 列を省いた（324:8866）

開くメニュー（148:5884）もこの文書で決める。ツールバーの［開く ▾］の ▾ を押すと出る。

境目（PreviewSplitter）・プレビューを出すかどうか・欄の高さは `search.md` が持つ。ここでは欄の中だけを決める。

## 2. 置き場所

| もの | パス |
|---|---|
| XAML | `scripts/tebunko/xaml/search/preview.xaml`（根は `Grid`、x:Name `PreviewRoot`） |
| 差す口 | `search.md` の `PreviewHost`（ContentControl。Grid.Row 5、既定の高さ 180） |
| 画面層 | `scripts/tebunko/ui/search/preview.ps1`（今の `ui/preview.ps1` と `ui/open_source.ps1` の開くところを移す） |
| 判断層 | `scripts/tebunko/ui/search/preview_view.ps1`（今の `ui/preview_view.ps1` を移して足す） |
| テスト | `tests/tebunko/ui/search/preview_view.Tests.ps1` |
| 読み込み口 | `gui.ps1`（画面層）・`tebunko/lib.ps1`（判断層） |

- Figma のコンポーネント: `xaml/search/preview`（353:33742）。variant は 既定・H-W・thumb・P-H・P-H-row3・P-H-row4・H-PP・H-TX。
- 見本の集まり: 「02 画面のバリエーション」の Preview（324:8867）。画面で使うのは `xaml/search/preview` だけ（決定）。
- 画面の H（108:230）と H-P のプレビューは、コンポーネントではなく「preview-panel」という名前の部品（328:218・328:675）。中身は 340:954 と同じ。

## 3. 部品の木

文言は `{文言ID}` で書く。値は窓 1280×820 のときの DIP。

```xml
<Grid x:Name="PreviewRoot" Background="{StaticResource Bg.Surface}">
  <Grid.RowDefinitions>
    <RowDefinition x:Name="PreviewToolbarRow" Height="44"/>   <!-- 0: ツールバー -->
    <RowDefinition x:Name="PreviewNoteRow"    Height="Auto"/> <!-- 1: 注記（読めない・列を省いた のときだけ） -->
    <RowDefinition x:Name="PreviewBodyRow"    Height="*"/>    <!-- 2: 中身 -->
  </Grid.RowDefinitions>

  <!-- 0: ツールバー -->
  <Border Grid.Row="0" Background="{StaticResource Bg.Window}"
          BorderBrush="{StaticResource Border.Soft}" BorderThickness="0,0,0,1"
          Padding="16,8,16,8">
    <Grid>
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>      <!-- ファイル名 -->
        <ColumnDefinition Width="16"/>     <!-- 間 16 -->
        <ColumnDefinition Width="Auto"/>   <!-- ボタン -->
      </Grid.ColumnDefinitions>
      <TextBlock x:Name="DetailTitle" Grid.Column="0" VerticalAlignment="Center"
                 Style="{StaticResource Meta}" Foreground="{StaticResource Ink.Strong}"
                 TextTrimming="CharacterEllipsis" TextWrapping="NoWrap"
                 Text="{preview.title}"/>
      <StackPanel x:Name="PreviewActions" Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
        <!-- 分かれたボタン（Excel）。Word・PowerPoint・テキストでは OpenMenuPart を隠し、40×24 になる -->
        <Border x:Name="OpenButtonFrame" Height="24" CornerRadius="4"
                BorderBrush="{StaticResource Accent}" BorderThickness="1" Background="{StaticResource Bg.Surface}">
          <StackPanel Orientation="Horizontal">
            <Button x:Name="OpenButton" Style="{StaticResource PreviewOpenPart}" Padding="10,0,10,0">
              <TextBlock Style="{StaticResource ColumnHeader}" Foreground="{StaticResource Accent}" Text="{preview.open}"/>
            </Button>
            <Grid x:Name="OpenMenuPart">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="1"/>     <!-- 区切りの線 -->
                <ColumnDefinition Width="Auto"/>  <!-- ▾ -->
              </Grid.ColumnDefinitions>
              <Rectangle x:Name="OpenDivider" Grid.Column="0" Fill="{StaticResource Accent}"/>
              <Button x:Name="OpenMenuButton" Grid.Column="1" Style="{StaticResource PreviewOpenPart}" Padding="6,0,10,0">
                <Viewbox Width="10" Height="10">
                  <Path Data="{StaticResource Icon.ChevronDown}" Stroke="{StaticResource Accent}" StrokeThickness="2"/>
                </Viewbox>
              </Button>
            </Grid>
          </StackPanel>
        </Border>
        <Button x:Name="OpenFolderButton" Margin="12,0,0,0" Style="{StaticResource LinkButton}" VerticalAlignment="Center">
          <TextBlock Style="{StaticResource Micro}" Foreground="{StaticResource Accent}"
                     TextDecorations="Underline" Text="{preview.openFolder}"/>
        </Button>
      </StackPanel>
    </Grid>
  </Border>

  <!-- 1: 注記 -->
  <Border x:Name="PreviewNoteBar" Grid.Row="1" Padding="16,6,16,6" Visibility="Collapsed"
          Background="{StaticResource Bg.Surface}">
    <TextBlock x:Name="PreviewNote" Style="{StaticResource Micro}" Foreground="{StaticResource Ink.Body}"
               TextTrimming="CharacterEllipsis" Text="{preview.note.columns}"/>
  </Border>

  <!-- 2: 中身。種類ごとに 1 つだけ Visible にする -->
  <Grid Grid.Row="2">
    <!-- 2a: 未選択 -->
    <TextBlock x:Name="PreviewPlaceholder" HorizontalAlignment="Center" VerticalAlignment="Center"
               Style="{StaticResource Meta}" Foreground="{StaticResource Ink.Body}"
               Text="{preview.placeholder}"/>

    <!-- 2b: Excel の格子 -->
    <ScrollViewer x:Name="PreviewGridScroll" HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto"
                  Visibility="Collapsed">
      <StackPanel HorizontalAlignment="Left">
        <!-- 列の見出しの行（高さ 22）。左端は行番号の列（幅 30）の空きのセル -->
        <ItemsControl x:Name="PreviewHeader" ItemTemplate="{StaticResource PreviewHeaderCell}">
          <ItemsControl.ItemsPanel><ItemsPanelTemplate><StackPanel Orientation="Horizontal"/></ItemsPanelTemplate></ItemsControl.ItemsPanel>
        </ItemsControl>
        <!-- 行（高さ 22）。各行は 行番号のセル（30）＋ 列のセル -->
        <ItemsControl x:Name="PreviewRows" ItemTemplate="{StaticResource PreviewGridRow}"/>
      </StackPanel>
    </ScrollViewer>

    <!-- 2c: Word・PowerPoint の段落 -->
    <ScrollViewer x:Name="PreviewDocScroll" HorizontalScrollBarVisibility="Disabled" VerticalScrollBarVisibility="Auto"
                  Visibility="Collapsed">
      <ItemsControl x:Name="PreviewDocument" Margin="24,14,24,14" ItemTemplate="{StaticResource PreviewParagraph}">
        <!-- 段落の間 6（PreviewParagraph の Margin 0,0,0,6。最後の段落は 0） -->
      </ItemsControl>
    </ScrollViewer>

    <!-- 2d: テキストの行 -->
    <ScrollViewer x:Name="PreviewLinesScroll" HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto"
                  Visibility="Collapsed">
      <ItemsControl x:Name="PreviewLines" Margin="0,8,0,8" ItemTemplate="{StaticResource PreviewTextLine}"/>
    </ScrollViewer>
  </Grid>

  <!-- 開くメニュー（重ねて出す） -->
  <Popup x:Name="OpenMenuPopup" PlacementTarget="{Binding ElementName=OpenButtonFrame}"
         Placement="Bottom" StaysOpen="False" AllowsTransparency="True">
    <Border Width="180" Background="{StaticResource Bg.Surface}" BorderBrush="{StaticResource Border.Soft}"
            BorderThickness="1" CornerRadius="6">
      <StackPanel>
        <Button x:Name="MenuOpenNormal"   Style="{StaticResource PreviewMenuItem}" Content="{preview.menu.open}"/>
        <Rectangle Height="1" Fill="{StaticResource Border.Soft}"/>
        <Button x:Name="MenuOpenNew"      Style="{StaticResource PreviewMenuItem}" Content="{preview.menu.new}"/>
        <Rectangle Height="1" Fill="{StaticResource Border.Soft}"/>
        <Button x:Name="MenuOpenReadOnly" Style="{StaticResource PreviewMenuItem}" Content="{preview.menu.readOnly}"/>
      </StackPanel>
    </Border>
  </Popup>
</Grid>
```

### 行のテンプレート（DataTemplate）

```xml
<!-- 列の見出しのセル。行番号の列の上は空き（幅 30） -->
<DataTemplate x:Key="PreviewHeaderCell">
  <Border Width="{Binding Width}" Height="22" Padding="8,0,8,0" Background="{StaticResource Bg.Hover}"
          BorderBrush="{StaticResource Border.Soft}" BorderThickness="0,0,1,1">
    <TextBlock Text="{Binding Text}" Style="{StaticResource Micro}" Foreground="{StaticResource Ink.Strong}"
               VerticalAlignment="Center" HorizontalAlignment="Center" TextTrimming="CharacterEllipsis"/>
  </Border>
</DataTemplate>

<!-- Excel の 1 行。Cells の [0] は行番号のセル -->
<DataTemplate x:Key="PreviewGridRow">
  <ItemsControl ItemsSource="{Binding Cells}">
    <ItemsControl.ItemsPanel><ItemsPanelTemplate><StackPanel Orientation="Horizontal"/></ItemsPanelTemplate></ItemsControl.ItemsPanel>
    <ItemsControl.ItemTemplate>
      <DataTemplate>
        <Border Width="{Binding Width}" Height="22" Padding="8,0,8,0" Background="{Binding Background}"
                BorderBrush="{StaticResource Border.Soft}" BorderThickness="0,0,1,1">
          <TextBlock Text="{Binding Text}" FontWeight="{Binding Weight}" Style="{StaticResource Micro}"
                     Foreground="{StaticResource Ink.Strong}" VerticalAlignment="Center"
                     TextTrimming="CharacterEllipsis" TextWrapping="NoWrap"/>
        </Border>
      </DataTemplate>
    </ItemsControl.ItemTemplate>
  </ItemsControl>
</DataTemplate>

<!-- Word・PowerPoint の段落。Runs は 文・一致した語 の並び -->
<DataTemplate x:Key="PreviewParagraph">
  <TextBlock Margin="0,0,0,6" TextWrapping="Wrap" Style="{StaticResource Body}" Foreground="{StaticResource Ink.Strong}"
             local:PreviewRuns.Source="{Binding Runs}"/>
</DataTemplate>

<!-- テキストの 1 行（高さ 20） -->
<DataTemplate x:Key="PreviewTextLine">
  <Border Height="20" Padding="16,0,16,0" Background="{Binding Background}">
    <StackPanel Orientation="Horizontal">
      <TextBlock Width="24" TextAlignment="Right" Text="{Binding Number}" Style="{StaticResource Body}"
                 Foreground="{StaticResource Ink.Body}" VerticalAlignment="Center"/>
      <TextBlock Margin="12,0,0,0" TextWrapping="NoWrap" Style="{StaticResource Body}" Foreground="{StaticResource Ink.Strong}"
                 VerticalAlignment="Center" local:PreviewRuns.Source="{Binding Runs}"/>
    </StackPanel>
  </Border>
</DataTemplate>
```

- `local:PreviewRuns.Source` は書き方の例。PowerShell からは、画面層で `TextBlock.Inlines` に `Run`（と一致した語の印の `InlineUIContainer`）を足して作る（10 章）。
- Style `PreviewOpenPart`・`PreviewMenuItem`・`LinkButton`・`Micro.Strong` は theme に無い。足す案を 4 章の下と「決めること」に書く。

## 4. 寸法と色

色は theme のキー（値）で書く。

| 要素 | 大きさ | 余白 | 地の色 | 枠 | 文字の Style | 文字の色 | 角丸 | アイコン |
|---|---|---|---|---|---|---|---|---|
| PreviewRoot | 横 Fill × 180（search.md の PreviewRow） | 0 | Bg.Surface（#FFFFFF） | – | – | – | 0 | – |
| ツールバー | 横 Fill × 44 | 上下 8・左右 16。間 16 | Bg.Window（#F5F7FA） | 下 1 Border.Soft（#E0E2E5） | – | – | 0 | – |
| DetailTitle | 横 Fill × Hug | – | – | – | Meta（11 Regular） | Ink.Strong（#202124） | – | – |
| OpenButtonFrame（Excel） | 63 × 24 | – | Bg.Surface | 1 Accent（#0078D4） | – | – | 4 | – |
| OpenButton | Hug × 24 | 左右 10 | 透明 | – | ColumnHeader（10 Bold） | Accent | – | – |
| OpenDivider | 1 × 24（上下いっぱい） | – | Accent | – | – | – | – | – |
| OpenMenuButton | Hug × 24 | 左 6・右 10 | 透明 | – | – | – | – | chevron-down 10、線 0.83、Accent |
| OpenButtonFrame（Word・PowerPoint・テキスト） | 40 × 24 | OpenButton の左右 10 | Bg.Surface | 1 Accent | ColumnHeader | Accent | 4 | 無し（OpenMenuPart を Collapsed） |
| OpenFolderButton | Hug（約 70）× Hug | 左に 12（ボタンの間） | 透明 | – | Micro（10 Regular）・下線 | Accent | – | – |
| PreviewNoteBar | 横 Fill × Hug | 上下 6・左右 16 | Bg.Surface | – | Micro | Ink.Body（#5F6368） | – | – |
| PreviewPlaceholder | Hug、上下左右の中央 | – | – | – | Meta | Ink.Body | – | – |
| 列の見出しのセル | 幅は列ごと（下の表）× 22 | 左右 8 | Bg.Hover（#F3F3F4） | 右・下 1 Border.Soft | Micro（太さは決めること） | Ink.Strong | – | – |
| 行番号のセル | 30 × 22 | 左右 8 | Bg.Hover | 右・下 1 Border.Soft | Micro | Ink.Strong | – | – |
| 行番号のセル（選んだ行） | 30 × 22 | 同上 | Select.Soft（#E1F2FF） | 同上 | Micro・Bold | Ink.Strong | – | – |
| ふつうのセル | 列の幅 × 22 | 左右 8 | Bg.Surface | 右・下 1 Border.Soft | Micro | Ink.Strong | – | – |
| 選んだ行のセル | 同上 | 同上 | Select.Soft | 同上 | Micro | Ink.Strong | – | – |
| 一致したセル | 同上 | 同上 | Hit.Cell（#FFF3CD） | 同上 | Micro・Bold | Ink.Strong | – | – |
| PreviewDocument（Word・PowerPoint） | 横 Fill | 上下 14・左右 24。段落の間 6 | Bg.Surface | – | Body（13） | Ink.Strong | – | – |
| 見出しの段落（Word「第3条（支払条件）」・PowerPoint「導入事例のご紹介」） | – | – | – | – | Heading（13 SemiBold） | Ink.Strong | – | – |
| 一致した語（Word） | – | – | 無し | – | Heading（13 SemiBold） | Ink.Strong | – | – |
| 一致した語（PowerPoint・テキスト） | Hug | 左右 2・上下 1 | Hit（#FFF176） | – | Body.Strong（13 Bold） | Ink.Strong | 2 | – |
| PreviewLines（テキスト） | 横 Fill | 上下 8 | Bg.Surface | – | – | – | – | – |
| テキストの行 | 横 Fill × 20 | 左右 16。番号と文の間 12 | Bg.Surface | – | Body | Ink.Strong | – | – |
| テキストの行（選んだ行） | 同上 | 同上 | Select.Soft（Figma は変数でなく #E1F2FF） | – | Body | Ink.Strong | – | – |
| テキストの行番号 | 24 × 20、右寄せ | – | – | – | Body | Ink.Body（Figma は変数でなく #5F6368） | – | – |
| OpenMenuPopup | 180 × 72 | 0 | Bg.Surface | 1 Border.Soft | – | – | 6 | – |
| メニューの項目 | 横 Fill × 24 | 上下 4・左右 8 | 透明 | 項目の間に 1 Border.Soft の線 | Micro | Ink.Strong | – | – |

### Excel の列の幅

| 列 | 幅 | Figma |
|---|---|---|
| 行番号 | 30 | 340:954・324:8866 |
| 一致したセルの列 | 220 | 340:954（A3 の A）・324:8866（BXV） |
| ほかの列 | 200 | 同上 |

- 格子の幅は 30 ＋ 列の幅の和。340:954 は 850（A〜D）、324:8866 は 1050（5 列）。
- 列の幅は固定。欄の幅に合わせて引き伸ばさない。入らなければ横にスクロールする。
- 今のコードの行番号の列は 44（`types.ps1` の NumberWidth）。Figma の 30 にする。

### theme に無いもの（`../figma_wpf_map.md` の「食い違い」に足す案）

- **Style `Micro.Strong`（10 Bold）**: 一致したセル・選んだ行の行番号に使う。今は `Micro` に FontWeight Bold を重ねて書く。`ColumnHeader` も 10 Bold だが、名前が合わないので借りるのは［開く］の文字だけにする。
- **Style `PreviewOpenPart`**: 地が透明・枠が無い・高さ 24 のボタン。ホバーなどの見た目は「決めること」。
- **Style `PreviewMenuItem`**: 高さ 24・余白 8/4・Micro・Ink.Strong・左寄せ。ホバーの見た目は「決めること」。
- **Style `LinkButton`**: 地も枠も無く、文字だけのボタン。
- **使えないときの文字の色 #9AA0A6**: 値は Ink.Placeholder と同じ。Ink.Placeholder を使う（意味が違うので、別のキー `Ink.Disabled` を足すかは決めること）。
- **figma_wpf_map は「プレビューの文字は Cell（12）」としている。** Figma のセルは 10 なので、Figma に合わせて Micro（10）にする。figma_wpf_map を直す案。

## 5. 状態ごとの見え方

上から順に、最初に当てはまるものにする（判断層 `getPreviewView` の `Mode`）。

| 状態（Mode） | 入る条件 | ツールバー | 注記 | 中身 | Figma |
|---|---|---|---|---|---|
| `none` 未選択 | 結果の一覧で行を選んでいない | DetailTitle は空。［開く］は使えない見た目（枠 Border.Soft、文字 #9AA0A6）。OpenMenuPart も使えない。［フォルダを開く］は #9AA0A6 | 出さない | PreviewPlaceholder（高さ 136 の真ん中） | 324:8558 |
| `loading` 読んでいる | 行を選び、前後の中身を読み終えていない | 選んだ行の題を出す。ボタンは使える | 出さない | **未定**（Figma に無い） | – |
| `unreadable` 読めない | 前後の行を読めなかった（元のファイルが無い・開けない・時間切れ） | ふつう | `{preview.note.unreadable}` | 選んだ行だけ（Excel は見出し＋1 行の格子、それ以外は 1 段落・1 行） | 324:8672（高さ 113） |
| `excel` | 種類が Excel（セル・図形・コメント） | Excel の題。分かれたボタン［開く ▾］ | 列を省いたときだけ `{preview.note.columns}` | 格子。選んだ行の上下に行を足す。一致したセルは Hit.Cell・Bold | 340:954・353:33663・353:33741・324:7302 |
| `word` | 種類が Word | Word の題。［開く］だけ（▾ なし） | 出さない | 段落 | 353:33429 |
| `powerpoint` | 種類が PowerPoint（本文・ノート） | PowerPoint の題。［開く］だけ | 出さない | 段落 | 427:228 |
| `text` | 種類がテキスト（txt・md・csv・tsv・log など） | テキストの題。［開く］だけ | 出さない | 行。選んだ行は Select.Soft | 427:284 |

- 結果の一覧が無い状態（H0・H-E・H-E2・H-1・H-S・E14・E16）では、プレビューそのものを出さない（search.md の `ShowPreview`）。
- `excel` で列を省くのは、出す列の数が全体の列の数より少ないとき（例: 2,000 列のうち BQG〜BXX の 200 列）。

### 例（Figma の中身。名前は架空のものに置き換えてある）

**Excel・既定（H、340:954）**

- 題: `営業部\A社_見積書.xlsx ・ [シート]見積書!A3 ・ セル`
- 行 1〜5。
  - 1 行目: A「取引規模」、B「Aクラス」
  - 3 行目（選んだ行）: A3「見積先：(株)山田商事(御中)」（一致）、B「見積番号：12-2324-0143」
  - 4 行目: A「件名：基幹システム導入一式」、B「発行日：2024/10/15」

**Excel・row3（H-row3、353:33663）**

- 題の終わりは `…!A41 ・ セル`。
- 行 39〜43。
- A41「納品場所：(株)山田商事」（一致）、B41「本社ビル」。

**Excel・row4（H-row4、353:33741）**

- 題の終わりは `…!D5 ・ 図形`。
- 行 3〜7。
- D5「納品場所：(株)山田商事 本社 4F」（一致）。

**Excel・コメント（324:7302）**

- 行 10〜14。
  - 10 行目: 小計 1,200,000
  - 11 行目: 消費税 120,000
  - 12 行目: 値引き「(株)山田商事 佐藤様と合意済みの値引き」
  - 13 行目: 合計 1,300,000

**Excel・列を省いた（324:8866）**

- 題: `経理部\集計_2024.xlsx ・ [シート]明細!BXV3 ・ セル`
- 注記: `表示は BQG〜BXX の 200 列（全 2,000 列）`

**Word（H-W、353:33429）**

- 題: `顧客\契約書\基本契約書.docx ・ 1 ページ（目安） ・ 本文`
- 段落:
  - 「第3条（支払条件）」（見出し）
  - 「(株)山田商事（以下「甲」という）は、本契約に基づく代金を、毎月末日までに乙の指定する口座に振り込む。」（(株)山田商事 が一致。太字だけで、印は無い）
  - 「2 振込手数料は甲の負担とする。」

**PowerPoint（H-PP、427:228）**

- 題: `営業部\2025年4月.pptx ・ スライド 2 ・ 本文`
- 段落:
  - 「導入事例のご紹介」（見出し）
  - 「導入効果：(株)山田商事様の事例」（(株)山田商事 に印）
  - 「（以下、効果の数値を記載）」

**テキスト（H-TX、427:284）**

- 題: `営業部\議事録\定例会議事録.txt ・ 29 行目 ・ 本文`
- 行:
  - 28「■ 決定事項」
  - 29「次回までに(株)山田商事向けの見積を送る」（選んだ行。(株)山田商事 に印）
  - 30「・担当：佐藤（見積）、山田（日程の調整）。見積には保守の費用と、導入後 1 年間の問い合わせ窓口の費用を含める（詳しくは別紙の見積の条件を参照）。」（折り返さず、横にスクロール）
  - 31「■ 次回」
  - 32「2025/5/12（月）10:00〜 第2会議室」

### 開くメニュー

| 状態 | 見え方 |
|---|---|
| 閉じている | 何も出さない（既定） |
| 開いている | OpenButtonFrame の下に、右端をそろえて出す（148:5884。メニューの右端 1182 ＝ ボタンの右端）。上の 3 項目 |

- 開いたときの［開く ▾］の見た目は Figma に無い（変えない）。

## 6. 文言

| ID | 全文 | 使う状態 |
|---|---|---|
| `preview.title` | {パス} ・ {場所} ・ {種別} | excel・word・powerpoint・text・unreadable・loading。「・」の前後は半角の空白 |
| `preview.open` | 開く | すべて |
| `preview.openFolder` | フォルダを開く | すべて |
| `preview.placeholder` | 結果の表で行を選ぶと、ここにその前後の行を表示します。 | none |
| `preview.note.unreadable` | 前後の行を読めなかったため、選んだ行だけを表示しています。 | unreadable |
| `preview.note.columns` | 表示は {先頭の列}〜{最後の列} の {出した列の数} 列（全 {全体の列の数} 列） | excel（列を省いたとき）。数は 3 桁ごとに「,」。例「表示は BQG〜BXX の 200 列（全 2,000 列）」 |
| `preview.menu.open` | 開く | 開くメニュー |
| `preview.menu.new` | 新規で開く | 開くメニュー |
| `preview.menu.readOnly` | 読み取り専用で開く | 開くメニュー |

- {パス}・{場所}・{種別} は結果の行の値をそのまま使う（result_list.md）。
  - {パス} の例: `営業部\A社_見積書.xlsx`
  - {場所} の例: `[シート]見積書!A3`・`1 ページ（目安）`・`スライド 2`・`29 行目`
  - {種別} の例: `セル`・`図形`・`コメント`・`本文`・`ノート`
- 今のコードの注記「表示は {範囲} の {n} 列（全 {m} 列）」は、Figma の形（「〜」でつなぐ）にそろえる。
- キー操作のヒントは書かない。今の右クリックメニューの InputGestureText（「Enter」）もこの欄には出さない。

## 7. 操作と振る舞い

| きっかけ | 何が起きるか | 呼ぶ関数 | 次の状態 |
|---|---|---|---|
| 結果の一覧で行を選ぶ | 題をすぐ変え、120ms 待ってから前後の中身を裏で読む。読めたら中身を出す | `requestPreview` → `getPreviewView` → `setPreviewView` | loading → excel 等・unreadable |
| 選びを外す | 未選択に戻す | `getPreviewView $null` | none |
| ［開く］（OpenButton）をクリック | 選んだ行の元のファイルを、通常の開き方で開く。Excel は該当のセルを選ぶ | `openSource 'normal'` | 変わらない |
| ▾（OpenMenuButton）をクリック | 開くメニューを出す。出ていれば閉じる | `toggleOpenMenu` | メニューが開く |
| メニューの［開く］ | 通常で開き、メニューを閉じる | `openSource 'normal'` | 閉じる |
| メニューの［新規で開く］ | 新規（元のファイルを雛形に）で開き、メニューを閉じる | `openSource 'new'` | 閉じる |
| メニューの［読み取り専用で開く］ | 読み取り専用で開き、メニューを閉じる | `openSource 'readOnly'` | 閉じる |
| メニューの外をクリック | メニューを閉じる（148:5885「外側（押すと閉じる）」。Popup の StaysOpen=False） | – | 閉じる |
| ［フォルダを開く］ | 元のファイルのフォルダをエクスプローラーで開き、そのファイルを選ぶ | `openSourceFolder` | 変わらない |
| 欄の中をホイール | 中身を縦にスクロール。Shift なしの横は無い | – | – |
| 境目をドラッグ | 欄の高さを変える（search.md）。高さが変われば出す行の数を決め直す | `getPreviewRowCounts` | 変わらない |

- **使えないとき**（none）: OpenButton・OpenMenuButton・OpenFolderButton の IsEnabled を False にする。見た目は 5 章の none のとおり。
- **ホバー・押した**: Figma に無い。**未定**（決めること）。
- **フォーカス**: theme の既定のフォーカスの枠を使う。
- **Tab の順**: OpenButton → OpenMenuButton（Excel のときだけ）→ OpenFolderButton → 中身の ScrollViewer。DetailTitle にはフォーカスを当てない。
  - メニューが開いているときは MenuOpenNormal → MenuOpenNew → MenuOpenReadOnly。メニューを閉じたら OpenMenuButton にフォーカスを戻す。
- **ツールチップ**: 付けない（Figma に無い）。DetailTitle が「…」で切れたときに全文を出すかは決めること。
- 開いたあとの報告（「開きました：…」など）はステータスバーに出す（今の `openFoundSource` のまま。shell/status_bar.md）。

## 8. リサイズ

一覧が多いとき（スクロールする所・固定する所・仮想化・未定のこと）は `../scroll.md`（見本 PV-多・PV-多-1024）。

- 欄の高さは search.md の PreviewRow（既定 180・最小 120）。窓の高さを変えても 180 のまま。一覧（MainRow）が伸び縮みする。
- ツールバーは高さ 44 のまま。DetailTitle が縮み、入らないところは「…」で切る。ボタンは縮めない。
- 中身の高さ ＝ 欄の高さ − 44 −（注記の高さ）。180 のとき 136。
- **Excel の出す行の数**: 今の `getPreviewRowCounts`（行の高さ 22・最大 101 行）で、中身の高さから決める。136 なら 見出し 22 ＋ 5 行（110）で、選んだ行の上 2 行・下 2 行。
- **Excel の列**: 幅は固定（30 ＋ 220 ＋ 200 × n）。欄より広ければ横のスクロールバーを出す。欄より狭くても引き伸ばさず、右は空ける。
- **Word・PowerPoint**: 欄の幅で折り返す。入らなければ縦にスクロール。
- **テキスト**: 折り返さない。入らなければ横・縦にスクロール。

| 窓 | 主な領域の幅 | 見え方 |
|---|---|---|
| 1024×640 | 804（ナビ 220） | 欄の高さ 180。Excel の格子（850）は欄（804）より広く、横のスクロールバーが出る（D 列の右端が隠れる）。題は「…」で切れることがある。ボタンは右端に 16 空けて並ぶ |
| 1280×820 | 1060 | Figma のとおり。格子 850 の右に 210 の空き |
| 1600×1000 | 1380 | 欄の高さ 180 のまま（一覧が伸びる）。格子の右が広く空く。広いときに列を増やして読むかは **未定** |

- 開くメニューはボタンに付いて動く（PlacementTarget）。窓の幅を変えても、ボタンの右下から外れない。

## 9. 判断層

`scripts/tebunko/ui/search/preview_view.ps1`。`$ui`・`$window`・WPF の型に触らない。戻り値は hashtable。

### getPreviewView

選んだ結果の行と、裏で読んだ前後の中身から、プレビューの見え方を決める。

- 入力
  - `$row`: 結果の行（`$null` なら未選択）。使う値は `Kind`（`excel`・`word`・`powerpoint`・`text`）・`PathDisplay`・`PlaceDisplay`・`KindLabel`（セル・本文など）
  - `$context`: 読んだ中身（`$null` なら読んでいる途中）。`Ok`（読めたか）・`TotalColumns`・`FirstColumn`・`LastColumn`・`ShownColumns`
- 出力

| キー | 型 | 中身 |
|---|---|---|
| `Mode` | string | `none`・`loading`・`unreadable`・`excel`・`word`・`powerpoint`・`text` |
| `Title` | string | `getPreviewTitle` の値。none は空 |
| `OpenEnabled` | bool | none だけ False |
| `ShowOpenMenu` | bool | Excel のときだけ True（Figma のとおり。決めること） |
| `FolderEnabled` | bool | none だけ False |
| `Note` | string | 注記の文。出さないときは空 |
| `Body` | string | 中身のどれを出すか: `placeholder`・`grid`・`document`・`lines`・`''`（loading） |

```powershell
It '<Name>' -TestCases @(
  @{ Name='未選択';       Row=$null;                    Ctx=$null;                                 Mode='none';       Open=$false; Menu=$false; Body='placeholder'; Note='' }
  @{ Name='読んでいる';   Row=@{Kind='excel'};          Ctx=$null;                                 Mode='loading';    Open=$true;  Menu=$true;  Body='';            Note='' }
  @{ Name='読めない';     Row=@{Kind='word'};           Ctx=@{Ok=$false};                          Mode='unreadable'; Open=$true;  Menu=$false; Body='document';    Note='前後の行を読めなかったため、選んだ行だけを表示しています。' }
  @{ Name='Excel';        Row=@{Kind='excel'};          Ctx=@{Ok=$true;TotalColumns=4;ShownColumns=4}; Mode='excel'; Open=$true; Menu=$true; Body='grid'; Note='' }
  @{ Name='Excel 列を省く'; Row=@{Kind='excel'};        Ctx=@{Ok=$true;TotalColumns=2000;ShownColumns=200;FirstColumn='BQG';LastColumn='BXX'}; Mode='excel'; Open=$true; Menu=$true; Body='grid'; Note='表示は BQG〜BXX の 200 列（全 2,000 列）' }
  @{ Name='Word';         Row=@{Kind='word'};           Ctx=@{Ok=$true};                           Mode='word';       Open=$true;  Menu=$false; Body='document';    Note='' }
  @{ Name='PowerPoint';   Row=@{Kind='powerpoint'};     Ctx=@{Ok=$true};                           Mode='powerpoint'; Open=$true;  Menu=$false; Body='document';    Note='' }
  @{ Name='テキスト';     Row=@{Kind='text'};           Ctx=@{Ok=$true};                           Mode='text';       Open=$true;  Menu=$false; Body='lines';       Note='' }
) {
  param($Row, $Ctx, $Mode, $Open, $Menu, $Body, $Note)
  $v = getPreviewView $Row $Ctx
  $v.Mode | Should -Be $Mode; $v.OpenEnabled | Should -Be $Open; $v.ShowOpenMenu | Should -Be $Menu
  $v.Body | Should -Be $Body; $v.Note | Should -Be $Note
}
```

- テストの `Row` には、ほかの値（`PathDisplay` など）も入れておく（ここでは省いた）。
- 読めないときの Excel は `grid`（選んだ行だけ）、テキストは `lines`。

### getPreviewTitle

- 入力: `$path`・`$place`・`$kindLabel`（文字列）
- 出力: 文字列 `"{path} ・ {place} ・ {kindLabel}"`。どれかが空なら、その部分と前の「 ・ 」を省く。

```powershell
It '<Expected>' -TestCases @(
  @{ Path='営業部\A社_見積書.xlsx';      Place='[シート]見積書!A3'; Kind='セル'; Expected='営業部\A社_見積書.xlsx ・ [シート]見積書!A3 ・ セル' }
  @{ Path='顧客\契約書\基本契約書.docx'; Place='1 ページ（目安）';   Kind='本文'; Expected='顧客\契約書\基本契約書.docx ・ 1 ページ（目安） ・ 本文' }
  @{ Path='営業部\2025年4月.pptx';       Place='スライド 2';         Kind='本文'; Expected='営業部\2025年4月.pptx ・ スライド 2 ・ 本文' }
  @{ Path='営業部\議事録\定例会議事録.txt'; Place='29 行目';         Kind='本文'; Expected='営業部\議事録\定例会議事録.txt ・ 29 行目 ・ 本文' }
  @{ Path='営業部\メモ.txt';             Place='';                   Kind='本文'; Expected='営業部\メモ.txt ・ 本文' }
) { param($Path,$Place,$Kind,$Expected); getPreviewTitle $Path $Place $Kind | Should -Be $Expected }
```

### getPreviewColumnNote

- 入力: `$first`・`$last`（列の名前）・`$shown`・`$total`（数）
- 出力: 文字列。`$shown -ge $total` なら空。

```powershell
It '<Expected>' -TestCases @(
  @{ First='BQG'; Last='BXX'; Shown=200; Total=2000; Expected='表示は BQG〜BXX の 200 列（全 2,000 列）' }
  @{ First='A';   Last='D';   Shown=4;   Total=4;    Expected='' }
  @{ First='A';   Last='CV';  Shown=100; Total=1500; Expected='表示は A〜CV の 100 列（全 1,500 列）' }
) { param($First,$Last,$Shown,$Total,$Expected); getPreviewColumnNote $First $Last $Shown $Total | Should -Be $Expected }
```

### getPreviewColumnWidths

- 入力: `$columns`（列の名前の配列）・`$hitColumn`（一致したセルの列の名前）
- 出力: hashtable `@{ Number = 30; Widths = @(…); Total = <和> }`。一致した列は 220、ほかは 200。

```powershell
It '<Name>' -TestCases @(
  @{ Name='A が一致';   Cols=@('A','B','C','D');           Hit='A';   Widths=@(220,200,200,200);     Total=850 }
  @{ Name='D が一致';   Cols=@('A','B','C','D');           Hit='D';   Widths=@(200,200,200,220);     Total=850 }
  @{ Name='BXV が一致'; Cols=@('BXT','BXU','BXV','BXW','BXX'); Hit='BXV'; Widths=@(200,200,220,200,200); Total=1050 }
) { param($Cols,$Hit,$Widths,$Total); $v = getPreviewColumnWidths $Cols $Hit; $v.Widths | Should -Be $Widths; $v.Total | Should -Be $Total }
```

- 「一致した列を 220 にする」は 340:954 と 324:8866 から読んだ決まり。H-row4（D5 が一致）の D 列の幅が 220 かは確かめていない（決めること）。

### getPreviewCellStyle

- 入力: `$isNumber`（行番号のセルか）・`$isSelectedRow`・`$isHit`
- 出力: `@{ Background = '<theme のキー>'; Bold = <bool> }`

```powershell
It '<Name>' -TestCases @(
  @{ Name='ふつう';           Num=$false; Sel=$false; Hit=$false; Bg='Bg.Surface';  Bold=$false }
  @{ Name='選んだ行';         Num=$false; Sel=$true;  Hit=$false; Bg='Select.Soft'; Bold=$false }
  @{ Name='一致したセル';     Num=$false; Sel=$true;  Hit=$true;  Bg='Hit.Cell';    Bold=$true }
  @{ Name='行番号';           Num=$true;  Sel=$false; Hit=$false; Bg='Bg.Hover';    Bold=$false }
  @{ Name='選んだ行の行番号'; Num=$true;  Sel=$true;  Hit=$false; Bg='Select.Soft'; Bold=$true }
) { param($Num,$Sel,$Hit,$Bg,$Bold); $v = getPreviewCellStyle $Num $Sel $Hit; $v.Background | Should -Be $Bg; $v.Bold | Should -Be $Bold }
```

### getPreviewRuns

- 入力: `$text`（1 段落・1 行の文）・`$matches`（一致した位置 `@{Start;Length}` の配列）・`$kind`（`word`・`powerpoint`・`text`）
- 出力: `@{ Text; Hit }` の配列。`Hit` は `''`（ふつう）・`'bold'`（Word）・`'mark'`（PowerPoint・テキストの印）。

```powershell
It '<Name>' -TestCases @(
  @{ Name='Word';       Text='(株)山田商事（以下「甲」という）は'; M=@(@{Start=0;Length=7}); Kind='word';       Out=@(@{Text='(株)山田商事';Hit='bold'}, @{Text='（以下「甲」という）は';Hit=''}) }
  @{ Name='PowerPoint'; Text='導入効果：(株)山田商事様の事例';     M=@(@{Start=5;Length=7}); Kind='powerpoint'; Out=@(@{Text='導入効果：';Hit=''}, @{Text='(株)山田商事';Hit='mark'}, @{Text='様の事例';Hit=''}) }
  @{ Name='テキスト';   Text='次回までに(株)山田商事向けの見積を送る'; M=@(@{Start=5;Length=7}); Kind='text';    Out=@(@{Text='次回までに';Hit=''}, @{Text='(株)山田商事';Hit='mark'}, @{Text='向けの見積を送る';Hit=''}) }
  @{ Name='一致なし';   Text='■ 決定事項';                         M=@();                    Kind='text';       Out=@(@{Text='■ 決定事項';Hit=''}) }
) { param($Text,$M,$Kind,$Out); $r = getPreviewRuns $Text $M $Kind; $r.Count | Should -Be $Out.Count
    for ($i = 0; $i -lt $Out.Count; $i++) { $r[$i].Text | Should -Be $Out[$i].Text; $r[$i].Hit | Should -Be $Out[$i].Hit } }
```

### getOpenMenuItems

- 入力: `$kind`
- 出力: `@{ Mode; TextId; Enabled }` の配列。並びは 開く・新規で開く・読み取り専用で開く。

```powershell
It '<Kind>' -TestCases @(
  @{ Kind='excel'; Modes=@('normal','new','readOnly'); Ids=@('preview.menu.open','preview.menu.new','preview.menu.readOnly'); Enabled=@($true,$true,$true) }
) { param($Kind,$Modes,$Ids,$Enabled); $v = getOpenMenuItems $Kind
    ($v | ForEach-Object Mode) | Should -Be $Modes; ($v | ForEach-Object TextId) | Should -Be $Ids; ($v | ForEach-Object Enabled) | Should -Be $Enabled }
```

- Word・PowerPoint・テキストではメニューを出さない（`ShowOpenMenu = $false`）。出すことになったら、この表に行を足す（決めること）。

### 今の関数

- `getPreviewRowCounts($height, $rowHeight, $maxRows)` は今のまま使う（出す行の数を前と後ろに分ける）。
- `toStatusText` は今のまま（ステータスバーに出すときに 40 文字で切る）。

## 10. 画面層

`scripts/tebunko/ui/search/preview.ps1`。

1. **初期化 `initPreview`**: `preview.xaml` を XamlReader で読み、`PreviewHost.Content` に差す。x:Name の要素を `$ui` に入れる。次のイベントをつなぐ。
   - `OpenButton.Click` → `openSource 'normal'`
   - `OpenMenuButton.Click` → `toggleOpenMenu`（`OpenMenuPopup.IsOpen` を反転）
   - `MenuOpenNormal`・`MenuOpenNew`・`MenuOpenReadOnly` の `Click` → `OpenMenuPopup.IsOpen = $false` のあと `openSource '<mode>'`
   - `OpenMenuPopup.Closed` → `OpenMenuButton.Focus()`
   - `OpenFolderButton.Click` → `openSourceFolder`
   - `PreviewRoot.SizeChanged` → 高さが変わったら `requestPreview`（行の数を決め直す）
   - 最初は `setPreviewView (getPreviewView $null $null)`。
2. **行を選んだとき `requestPreview $row`**（result_list.md の選びの変化から呼ぶ）
   - すぐ `setPreviewView (getPreviewView $row $null)`（題を出し、中身は loading）。
   - 120ms の `DispatcherTimer`（今の detailTimer）を止めてかけ直す。続けて選びを動かしたときに、読むのを 1 回にするため。
   - 時間が来たら、`startJob` で裏のスレッドに `readPackContext` を投げる（今のまま）。行の数は `getPreviewRowCounts` で決めて渡す。
   - 結果は `Dispatcher.BeginInvoke` で画面のスレッドに戻す。そのときの選んだ行が、頼んだ行と違えば捨てる。
   - 戻ったら `setPreviewView (getPreviewView $row $ctx)` と、中身を作る `fillPreviewBody`。
3. **`setPreviewView $v`**
   - `DetailTitle.Text = $v.Title`
   - `OpenButton.IsEnabled`・`OpenMenuButton.IsEnabled = $v.OpenEnabled`、`OpenFolderButton.IsEnabled = $v.FolderEnabled`
   - `OpenMenuPart.Visibility` を `$v.ShowOpenMenu` で切り替える。
   - 使えないときは `OpenButtonFrame.BorderBrush = Border.Soft`、文字を Ink.Placeholder にする（Style の Trigger で書いてもよい）。
   - `PreviewNoteBar` は `$v.Note` が空なら Collapsed。
   - `$v.Body` で `PreviewPlaceholder`・`PreviewGridScroll`・`PreviewDocScroll`・`PreviewLinesScroll` のどれか 1 つを Visible にする。
4. **`fillPreviewBody $ctx`**
   - grid: `getPreviewColumnWidths` と `getPreviewCellStyle` で、`PreviewHeader.ItemsSource`・`PreviewRows.ItemsSource` を作る（今の `types.ps1` の PreviewTable・PreviewRow・PreviewColumn に Width・Background・Weight を足す）。そのあと、選んだ行と一致した列が見えるところまでスクロールする（`BringIntoView`）。
   - document・lines: `getPreviewRuns` の結果から、`TextBlock.Inlines` に `Run` を足す。
     - `bold` は `FontWeight = SemiBold`。
     - `mark` は `InlineUIContainer` に Border（地 Hit・角丸 2・余白 2/1）と Bold の TextBlock を入れる。
   - テキストの選んだ行は、その行の Border の地を Select.Soft にし、見えるところまでスクロールする。
5. **重い処理を画面のスレッドでしない。** ファイルを読むのは裏のスレッドだけ。元のファイルを探すのも今の `findSourceFile`（見つかるまで画面のスレッドで待たない）を使う。
6. 今の `OpenModeCombo`（ComboBox。選んだ開き方を設定に残す）は、このメニューに置き換える。ダブルクリック・Enter の開き方をどうするかは決めること。

## 11. 受け入れ

| 見比べる png（`../png/`） | 状態 | 見るところ |
|---|---|---|
| `108_230.png`（H） | Excel 既定 | ツールバー 44・地 #F5F7FA・下の線。［開く ▾］63×24 の青い枠・区切りの線・▾。［フォルダを開く］の下線。格子の行番号 30・A 220・B〜D 200・行 22。A3 の地 #FFF3CD・太字、3 行目の地 #E1F2FF、行番号 3 が太字 |
| `158_823.png`（H-row3） | Excel row3 | 行 39〜43 が出る。A41 と B41 に分かれている |
| `158_1118.png`（H-row4） | Excel row4（図形） | 行 3〜7。D5 が一致。題の終わりが「図形」 |
| `125_1839.png`（H-W） | Word | ［開く］40×24（▾ なし）。段落の左右 24・上 14、13 の文字、見出しが SemiBold、一致した語が太字で印が無い |
| `427_36175.png`（H-PP） | PowerPoint | 一致した語の地 #FFF176・角丸 2・太字。2 行目が折り返す |
| `427_36508.png`（H-TX） | テキスト | 行の高さ 20、番号 24 の右寄せ、選んだ 29 行目の地 #E1F2FF、30 行目が折り返さず右へはみ出す |
| `148_5884.png`（開くメニュー） | メニュー | 180×72、角丸 6、項目 24 が 3 つ、間の線。ボタンの右下に右端をそろえて付く |
| Figma 324:8558・324:8672・324:8866（png なし） | 未選択・読めない・列を省いた | 使えないボタンの色 #9AA0A6。注記の余白 16/6・10 の文字 |

- 窓を 1024×640 にして、格子に横のスクロールバーが出ること・題が「…」で切れること・メニューがボタンに付いて出ることを見る。
- 1600×1000 にして、欄の高さが 180 のままであること・格子が引き伸ばされないことを見る。
- 判断層のテスト（9 章）が通ること。

## 決めること

1. **Word・PowerPoint・テキストに ▾（開くメニュー）を付けるか。** Figma は Excel だけ分かれたボタン。今のコードは Word・PowerPoint でも読み取り専用・新規で開ける（テキストは開き方を使わない）。
2. **メニューで選んだ開き方を「既定」として残すか。** 今は［開き方］の ComboBox で選んだものを設定に残し、ダブルクリック・Enter もそれで開く。Figma のメニューはその場で開くだけに見える。
3. **figma_wpf_map の x:Name `OpenModeCombo` を `OpenMenuButton`・`OpenMenuPopup` に改めるか。**
4. **Excel の列の見出しの太さ。** 340:954 は Bold、324:8867 の見本は Regular。
5. **一致した列を 220 にする決まりでよいか。** 340:954 は A（A3 が一致）、324:8866 は BXV が 220。diff_round2 は「等幅」と書いている。H-row4（D5 が一致）の D 列は確かめていない。
6. **PowerPoint のノートとテキストの log の出し方。** 324:8867 の見本は格子（30 ＋ 640 の列、行 22、log の長い行は 50 に折り返す）。353:33742 の H-PP・H-TX は段落・行（H-TX は折り返さない、行 20）。この文書は H-PP・H-TX に合わせた。
7. **Word の一致した語の見せ方。** Word は太字だけ（印なし）、PowerPoint・テキストは印（Hit #FFF176）と太字。そろえるか。
8. **読んでいる途中（loading）の見え方。** Figma に無い。
9. **ボタン・メニューの項目のホバー・押した見た目。** Figma に無い（Style `PreviewOpenPart`・`PreviewMenuItem`・`LinkButton` を足す案）。
10. **Style `Micro.Strong`（10 Bold）を theme に足すか。** figma_wpf_map の「プレビューは Cell（12）」を Figma の 10 に直すか。
11. **使えないときの色 #9AA0A6 を Ink.Placeholder で書くか、`Ink.Disabled` を足すか。**
12. **テキストの選んだ行（#E1F2FF）と行番号（#5F6368）が Figma で変数でなく直に書かれている。** Select.Soft・Ink.Body につなぐよう Figma を直すか。
13. **境目（search.md の PreviewSplitter）の見た目。** Figma 108:449 は地 Bg.Hover（#F3F3F4）・上下 1 Border.Soft・つまみ 32×2 Ink.Body（#5F6368）・角丸 1。search.md は地 Bg.Surface・つまみ Border.Grip（#A9AFB6）。今のコードは高さ 11・つまみ 48×5。search.md と合わせて決める。
14. **欄の高さ。** variant の高さは 178・135・160・176 と種類で違う。画面では 180 で固定と読んだ。種類で高さを変えないでよいか。
15. **窓が広いとき（1600）に、列や行を増やして読むか。**
16. **DetailTitle が「…」で切れたとき、全文のツールチップを出すか。**
17. **開くメニューを開いている間の［開く ▾］の見た目。** Figma に無い。
