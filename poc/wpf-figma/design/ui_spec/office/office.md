# Office の終了（office）

## 1. 役割

動いている Excel・Word・PowerPoint を一覧にし、バックグラウンドに残ったもの（画面に出ていないもの）をまとめて終了する画面。選んだものだけを終了することもできる。画面に表示中のものを終了するときは、保存していない変更が失われることを確かめる（`../dialog/confirm.md` の DC-kill・DC-kill-vis）。

- Figma の部品: `xaml/office/office` 374:50。variant は次の 5 つで、どれも 1060×788。

  | variant | ノード | 使う画面 |
  |---|---|---|
  | 既定 | 340:1438 | 106 の P |
  | P | 374:283 | 106 の P（109:1638） |
  | P-P | 374:908 | 148 の P-P |
  | P-P-killed | 374:1497 | 148 の P-P-killed（バックグラウンドを終了した後） |
  | E60 | 381:76 | 106 の E60（117:1303。動いている Office が無い） |

- 既定・P・P-P は同じ見た目（3 行）。
- 絵: `parts/illust_no_office` 379:35945（64×64、Lucide circle-check、線 2px、塗りなし、色 Illust.Line）。
- `../png/` にこの画面の画像は無い。見比べは Figma の get_screenshot で行う（11 章）。

## 2. 置き場所

- XAML: `scripts/tebunko/xaml/office/office.xaml`（今の `tab_kill.xaml` を置き換える）。
- 画面層 `ui/office/office.ps1`（今の `ui/process_tab.ps1` から移す）、判断層 `ui/office/office_view.ps1`（新しく分ける）。
- 親: shell の `ContentHost`（`../shell/shell.md`）。ナビで「Office の終了」（`KillTab`）を選んだときに差す。
- 読み込み: 起動のときに 1 回だけ `XamlReader::Load` で読み、`FindName` で x:Name を引く。
- `ProcessSummaryText` に当たる文は、Figma ではステータスバーにある（`../shell/status_bar.md`）。この部品には置かない。

## 3. 部品の木

```xml
<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
      x:Name="OfficeRoot" Background="{DynamicResource Bg.Surface}">
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
        <TextBlock Style="{StaticResource PageTitle}" Foreground="{DynamicResource Ink.Strong}" Text="{office.title}" />
        <TextBlock Style="{StaticResource Cell}" Foreground="{DynamicResource Ink.Body}" Margin="0,4,0,0"
                   TextWrapping="Wrap" Text="{office.description}" />
      </StackPanel>
      <!-- actions-col（E60 では縦に 2 段。右寄せ・間 4） -->
      <StackPanel Grid.Column="1" Orientation="Vertical" HorizontalAlignment="Right" VerticalAlignment="Center">
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
          <Button x:Name="KillSelectedButton" Style="{StaticResource Button.Base}" Height="32" Padding="14,0"
                  Content="{office.killSelected}" />
          <Button x:Name="KillAllButton" Style="{StaticResource Primary}" Margin="8,0,0,0" Padding="16,7"
                  Content="{office.killAll}" />
        </StackPanel>
        <TextBlock x:Name="KillReasonText" Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Body}"
                   HorizontalAlignment="Right" Margin="0,4,0,0" Visibility="Collapsed" Text="{office.killReason}" />
      </StackPanel>
    </Grid>
  </Border>

  <!-- 表 -->
  <ScrollViewer Grid.Row="1" x:Name="ProcessScroll" HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto">
    <DataGrid x:Name="ProcessGrid" MinWidth="1047" AutoGenerateColumns="False" HeadersVisibility="Column"
              GridLinesVisibility="None" CanUserAddRows="False" CanUserResizeRows="False" IsReadOnly="True"
              SelectionMode="Extended" RowHeight="36" ColumnHeaderHeight="28"
              Background="{DynamicResource Bg.Surface}" RowBackground="{DynamicResource Bg.Surface}"
              AlternatingRowBackground="{DynamicResource Bg.Stripe}" AlternationCount="2" BorderThickness="0">
      <DataGrid.Columns>
        <DataGridTemplateColumn Width="37">                     <!-- 選ぶ（チェック 13 + 左右 12） -->
          <DataGridTemplateColumn.CellTemplate>
            <DataTemplate><CheckBox Style="{StaticResource Choice}" IsChecked="{Binding Selected}" VerticalAlignment="Center" /></DataTemplate>
          </DataGridTemplateColumn.CellTemplate>
        </DataGridTemplateColumn>
        <DataGridTextColumn Header="{office.col.app}" Width="120" Binding="{Binding AppName}"
                            ElementStyle="{StaticResource Label}" />                       <!-- Ink.Strong -->
        <DataGridTextColumn Header="{office.col.file}" Width="*" MinWidth="160" Binding="{Binding FileText}"
                            ElementStyle="{StaticResource Meta.Trim}" />                   <!-- Ink.Body・「…」で切る -->
        <DataGridTemplateColumn Header="{office.col.view}" Width="180">
          <DataGridTemplateColumn.CellTemplate>
            <DataTemplate>
              <Border Padding="8,2" CornerRadius="3" HorizontalAlignment="Left" VerticalAlignment="Center"
                      Background="{Binding TagBackground}">
                <TextBlock Style="{StaticResource Chip}" Foreground="{Binding TagForeground}" Text="{Binding TagText}" />
              </Border>
            </DataTemplate>
          </DataGridTemplateColumn.CellTemplate>
        </DataGridTemplateColumn>
      </DataGrid.Columns>
    </DataGrid>
  </ScrollViewer>

  <!-- 空のとき（E60） -->
  <Border Grid.Row="1" x:Name="OfficeEmptyPanel" Visibility="Collapsed" BorderBrush="{DynamicResource Border.Soft}">
    <StackPanel HorizontalAlignment="Center" VerticalAlignment="Center">
      <Image Width="64" Height="64" Source="{StaticResource Illust.NoOffice}" HorizontalAlignment="Center" />
      <TextBlock Margin="0,16,0,0" FontSize="16" FontWeight="SemiBold" Foreground="{DynamicResource Ink.Value}"
                 HorizontalAlignment="Center" Text="{office.empty.title}" />
      <TextBlock Margin="0,16,0,0" Style="{StaticResource Body}" LineHeight="23.4" Foreground="{DynamicResource Ink.Muted}"
                 HorizontalAlignment="Center" TextAlignment="Center" Text="{office.empty.body}" />
    </StackPanel>
  </Border>
</Grid>
```

- Figma の列は「チェック 13／アプリ 120／開いているファイル Fill（最小 160）／空 60／空 120／表示 180／空 110」。空の 3 列（60・120・110）と、最後の列の透明（opacity 0）の「更新」「中止」と「…」は、インデックス管理の一覧の写しの残り。上の木では置いていない（「決めること」）。置くなら表の最小幅 1047 はそのままで、空の列を足す。
- `Meta.Trim`（Meta に `TextTrimming="CharacterEllipsis"` を足した Style）は theme に無いキー。`Illust.NoOffice` も theme に足す（DrawingImage）。
- 今の XAML の PID・開始時刻・メモリの列、［バックグラウンドのみ終了］（`KillBackgroundButton`）、［すべて終了］、［一覧を更新］（`RefreshProcessButton`）は置かない。Figma に無い（「決めること」）。

## 4. 寸法と色

| 要素 | 大きさ | 余白 | 地 | 枠 | 文字の Style | 文字の色 | 角丸 |
|---|---|---|---|---|---|---|---|
| 見出しの区画 | 幅いっぱい | 16 | Bg.Surface（#FFFFFF） | 下 1px Border.Soft（#E0E2E5） | — | — | — |
| 題 | Hug | — | — | — | PageTitle（16 Bold） | Ink.Strong（#202124） | — |
| 説明 | 幅いっぱい | 上 4 | — | — | Cell（12） | Ink.Body（#5F6368） | — |
| KillSelectedButton（有効） | 高さ 32 | 左右 14 | Bg.Surface（#FFFFFF） | 1px Border.Soft（#E0E2E5） | 12 SemiBold | Ink.Strong（#202124） | 6 |
| KillSelectedButton（無効） | 同 | 同 | Bg.Tag（#F1F3F4） | 1px Border.Soft | 12 SemiBold | #A0A5AB（キー無し） | 6 |
| KillAllButton（有効） | Hug | 16/7 | Accent（#0078D4） | なし | 13 SemiBold | Ink.OnAccent（#FFFFFF） | 6 |
| KillAllButton（無効、E60） | Hug | 16/7 | #E8EAED（キー無し） | なし | 13 SemiBold | Ink.Placeholder（#9AA0A6） | 6 |
| ボタンの間 | — | 8 | — | — | — | — | — |
| KillReasonText | Hug | 上 4 | — | — | Meta（11） | Ink.Body（#5F6368） | — |
| 表の見出しの行 | 高さ 28・最小幅 859 | 左右 12・列の間 12 | Bg.Subtle（#F9FAFA） | 下 1px Border.Soft | 11 Bold（Meta.Strong） | Ink.Body（#5F6368） | — |
| 行 | 高さ 36 | 左右 12・列の間 12 | 交互に Bg.Surface（#FFFFFF）・Bg.Stripe（#FAFBFC） | — | — | — | — |
| アプリの名前 | 120 | — | — | — | Label（12 SemiBold） | Ink.Strong（#202124） | — |
| 開いているファイル | Fill（最小 160） | — | — | — | Meta（11） | Ink.Body（#5F6368） | — |
| タグ「バックグラウンド」 | Hug | 8/2 | Bg.Tag（#F1F3F4） | — | Chip（11 Medium） | Ink.Body（#5F6368） | 3 |
| タグ「画面に表示中」 | Hug | 8/2 | Select.Soft（#E1F2FF） | — | Chip（11 Medium） | Accent（#0078D4） | 3 |
| OfficeEmptyPanel | 表の領域いっぱい | 中身の間 16 | — | Border.Soft（#E0E2E5。辺は未定） | — | — | — |
| 絵 | 64×64 | — | — | — | — | 線 Illust.Line（キー無し）・2px | — |
| 空の見出し | Hug | — | — | — | 16 SemiBold | Ink.Value（#212126） | — |
| 空の説明 | Hug | — | — | — | Body（13）・行の高さ 1.8 | Ink.Muted（#6B737D） | — |

- 選ぶチェックは 13×13。`Choice` の Style に従う。

## 5. 状態ごとの見え方

| 状態 | 入る条件 | 表の行 | KillSelectedButton | KillAllButton | KillReasonText | OfficeEmptyPanel |
|---|---|---|---|---|---|---|
| 既定・P・P-P | 動いている Office が 1 つ以上、うちバックグラウンドが 1 つ以上 | 出す（例 3 行） | 有効 | 有効 `office.killAll`（例「（2 件）」） | 出さない | 出さない |
| P-P-killed | 動いている Office が 1 つ以上、バックグラウンドが 0 | 出す（例 1 行） | 有効 | 無効 `office.killAll`（「（0 件）」） | 出さない | 出さない |
| E60 | 動いている Office が 0 | 出さない | 無効 | 無効 `office.killAll.none`（件数なし） | 出す | 出す |

- 例の行（既定・P・P-P）:

  | アプリ | 開いているファイル | 表示 |
  |---|---|---|
  | Excel | （開いているファイルなし） | バックグラウンド |
  | Word | （開いているファイルなし） | バックグラウンド |
  | Excel | 見積書.xlsx | 画面に表示中 |

- P-P-killed の行は「Excel／見積書.xlsx／画面に表示中」の 1 行だけ。
- P-P-killed で KillSelectedButton が有効なのは、何も選んでいなくても有効に見える。選んでいないときに無効にするかは未定（「決めること」）。
- 行の並びは、バックグラウンドが先、その中は起動の古い順（今のコードの並び）。Figma の例もバックグラウンドが先。

## 6. 文言

| ID | 全文 | 使う状態 |
|---|---|---|
| office.title | Office の終了 | すべて |
| office.description | バックグラウンドに残った Excel・Word・PowerPoint を終了します。 | すべて |
| office.killSelected | 選んだものを終了 | すべて |
| office.killAll | バックグラウンドを終了（{件数} 件） | 既定・P・P-P・P-P-killed |
| office.killAll.none | バックグラウンドを終了 | E60 |
| office.killReason | 終了するものがありません | E60 |
| office.col.app | アプリ | 既定・P・P-P・P-P-killed |
| office.col.file | 開いているファイル | 同上 |
| office.col.view | 表示 | 同上 |
| office.file.none | （開いているファイルなし） | 同上 |
| office.tag.background | バックグラウンド | 同上 |
| office.tag.visible | 画面に表示中 | 同上 |
| office.empty.title | 動いている Office はありません | E60 |
| office.empty.body | Excel・Word・PowerPoint が起動していると、ここに表示されます。 | E60 |

- 「（2 件）」は全角の括弧、数字と「件」の間は半角の空白。
- 確認の文言は `../dialog/confirm.md` の DC-kill・DC-kill-vis。ステータスバーの文は `../shell/status_bar.md`。

## 7. 操作と振る舞い

| きっかけ | 何が起きるか | 呼ぶ関数 | 次の状態 |
|---|---|---|---|
| 画面を開く | 一覧を取り直す。開いている間は 5 秒ごとに別スレッドで取り直す（今の `refreshProcessesInBackground`） | `getOfficeProcesses`（状態層）→ `getOfficeView` | 既定・P-P-killed・E60 のどれか |
| 行のチェック・行をクリック | 選ぶ。取り直しても選んだ行（PID）は残す | — | そのまま |
| KillAllButton をクリック | DC-kill を出す。［終了する］でバックグラウンドのものだけを終了する | `newKillConfirm`（`../dialog/confirm.md`） | P-P-killed（表示中が残る）か E60 |
| KillSelectedButton をクリック | 選んだものに画面に表示中のものがあれば DC-kill-vis、無ければ DC-kill を出す。［終了する］で選んだものを終了する | `newKillConfirm` | 残りに合わせて決まる |

- ボタンのホバー・押した・無効の見た目は `Button.Base`・`Primary` に従う。無効の色は 4 章。
- Tab の順: KillSelectedButton → KillAllButton → ProcessGrid。
- ツールチップ: 無し（Figma に無い）。開いているファイルが「…」で切れたときに全文を出すかは未定。

## 8. リサイズ

一覧が多いとき（スクロールする所・固定する所・仮想化・未定のこと）は `../scroll.md`（見本 P-多・P-多-1024）。

| 窓 | 主な領域 | 見え方 |
|---|---|---|
| 1024×640（最小） | 804 幅 | 表の最小幅 1047 より狭いので、表が横にスクロールする。見出しは折り返さずに収まる（ボタンは右端） |
| 1280×820（基本） | 1060 幅 | 表が 1047 以上で横のスクロールは出ない |
| 1600×1000 | 1380 幅 | 「開いているファイル」の列だけが伸びる |

- 伸びるのは「開いているファイル」の列だけ（Fill、最小 160）。ほかの列は固定。
- 文字は 1 行で、はみ出したら「…」（`TextTrimming="CharacterEllipsis"`）。
- 縦に収まらないときは縦にスクロールする。見出しはスクロールしない。
- 空のとき（E60）の中身は、表の領域の中央。

## 9. 判断層（`ui/office/office_view.ps1`）

### getOfficeView

- 入力: `[object[]]$processes`（`getOfficeProcesses` の結果。各要素は `Id`・`AppName`・`Background [bool]`・`Title`・`StartTime`）
- 出力: `@{ Rows; BackgroundCount [int]; VisibleCount [int]; KillAllText [string]; KillAllEnabled [bool]; KillReason [string]; IsEmpty [bool] }`
  - Rows の各要素: `@{ Id; AppName; FileText; Tag = "background" | "visible"; TagText }`
  - FileText は `Title` が空なら「（開いているファイルなし）」。

| processes | BackgroundCount | VisibleCount | KillAllText | KillAllEnabled | KillReason | IsEmpty |
|---|---|---|---|---|---|---|
| 0 件 | 0 | 0 | バックグラウンドを終了 | `$false` | 終了するものがありません | `$true` |
| Excel（背景）・Word（背景）・Excel 見積書.xlsx（表示） | 2 | 1 | バックグラウンドを終了（2 件） | `$true` | `""` | `$false` |
| Excel 見積書.xlsx（表示）だけ | 0 | 1 | バックグラウンドを終了（0 件） | `$false` | `""` | `$false` |
| PowerPoint（背景）だけ | 1 | 0 | バックグラウンドを終了（1 件） | `$true` | `""` | `$false` |

### getOfficeRowView

- 入力: 1 つのプロセス
- 出力: `@{ FileText; Tag; TagText }`

| Background | Title | FileText | Tag | TagText |
|---|---|---|---|---|
| `$true` | `""` | （開いているファイルなし） | background | バックグラウンド |
| `$true` | `$null` | （開いているファイルなし） | background | バックグラウンド |
| `$false` | 見積書.xlsx | 見積書.xlsx | visible | 画面に表示中 |

- Title（窓の題）からファイルの名前を取り出す決まり（「見積書.xlsx - Excel」→「見積書.xlsx」など）は、Figma から読めない（「決めること」）。

### 行の並び（sortOfficeRows）

- バックグラウンドが先、その中は `StartTime` の古い順。今の `applyProcesses` の並びを判断層に移す。

| 入力（Background, StartTime） | 出力の順 |
|---|---|
| (F, 9:00)・(T, 9:10)・(T, 9:05) | (T, 9:05)・(T, 9:10)・(F, 9:00) |

## 10. 画面層（`ui/office/office.ps1`）

- 画面を開いている間、5 秒ごとに `refreshProcessesInBackground`（`startJob` で別スレッド）。結果は Dispatcher で `applyProcesses` に戻す。
- `applyProcesses`: 選んでいた行の Id を覚える → `getOfficeView` → `ProcessGrid.ItemsSource` を差し替え → 覚えた Id の行を選び直す → ボタンの文言・有効を写す → `IsEmpty` なら `ProcessScroll` を隠して `OfficeEmptyPanel` と `KillReasonText` を出す。ステータスバーの文を更新し、ナビのバッジ（`updateKillBadge`）に BackgroundCount を渡す。
- KillAllButton・KillSelectedButton: 確認（`showConfirm`）→ 終了（`Stop-Process` 相当の状態層の関数）→ すぐに `refreshProcesses`。
- タグの色は `Tag` から画面層の対応表で引く（background → Bg.Tag／Ink.Body、visible → Select.Soft／Accent）。

## 11. 受け入れ

見比べる Figma: 既定 340:1438、P-P 374:908、P-P-killed 374:1497、E60 381:76（`../png/` に画像は無い。get_screenshot で撮って並べる）。

- 見出しの余白 16、下の線が Border.Soft。
- ボタンが右端に「選んだものを終了」「バックグラウンドを終了（2 件）」の順で、間 8。
- 表の見出しの行が高さ 28、行が高さ 36 で交互の地。
- タグ「バックグラウンド」が灰の地、「画面に表示中」が淡い青の地に青の文字。
- P-P-killed で KillAllButton が無効になり「（0 件）」になる。
- E60 で表が消え、中央に 64 の絵と 2 行の文、ボタン 2 つが無効、右下に「終了するものがありません」。
- 1024×640 で表が横にスクロールし、見出しは崩れない。

## 決めること

- 空の 3 列（60・120・110）と透明の「更新」「中止」「…」を置くか（インデックス管理の一覧の写しの残りに見える）。
- 今の PID・開始時刻・メモリの列、［すべて終了］、［一覧を更新］を消してよいか。
- ProcessSummaryText をステータスバーに移すときの文言（`../shell/status_bar.md` と合わせる）。
- KillAllButton の文言: E60 は件数なし「バックグラウンドを終了」、P-P-killed は「（0 件）」。0 件のときにどちらにそろえるか。
- 無効のボタンの色 #A0A5AB（文字）・#E8EAED（主のボタンの地）に theme のキーを足すか。
- `Illust.Line` の色の値（theme の表に無い）。
- 既定・P・P-P に違いがあるか（どれも同じ見た目だった）。
- 選ぶチェックの振る舞い（行のクリックとチェックのどちらで選ぶか、全部を選ぶチェックを見出しに置くか）。何も選んでいないときに KillSelectedButton を無効にするか。
- 窓の題からファイルの名前を取り出す決まり。
- OfficeEmptyPanel の線がどの辺に付くか。
- 開いているファイルが切れたときのツールチップ。
