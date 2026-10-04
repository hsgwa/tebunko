# インデックス更新の確認（indexing_confirm）

## 1. 役割

インデックスの更新を始める前に、更新するフォルダとファイルの数を示し、始めてよいかを確かめるダイアログ（DC-start）。最新のフォルダは更新しないことと、Office のファイルはコピーしてから読むことを伝える。

- Figma の部品: `xaml/dialog/indexing_confirm` 372:669。

  | variant | ノード | 大きさ | 使う画面 |
  |---|---|---|---|
  | 既定 | 340:1544 | 520×401 | 106 の DC-start |
  | P-DC-start | 372:788 | 520×401 | 148 の P-DC-start（既定と同じ見た目。スクリーンショットだけで確かめた） |
  | P-DC-start-1 | 372:948 | 520×335 | 148 の P-DC-start-1（更新するフォルダが 1 つ） |

- 開く所: インデックス管理の一覧の更新のボタン（`IndexingButton`、`../index/index_list.md`）。
- メンテナの決定: Office のファイルはコピーしてから読むことを 1 行で示す（office-copy-note）。
- `../png/` にこのダイアログの画像は無い。見比べは Figma の get_screenshot で行う（11 章）。

## 2. 置き場所

- XAML: `scripts/tebunko/xaml/dialog_indexing_confirm.xaml`（今のファイルを書き直す。今は 880×500）。
- 画面層 `ui/indexing_tab.ps1` の `showIndexingConfirmDialog`、判断層 `ui/indexing_view.ps1`（`newPlanViewRows`・`getIndexingConfirmText`）。
- 型: Window。`ShowDialog()` で出す。Owner は主の窓。
- 置き方: 主の窓の中央（`WindowStartupLocation="CenterOwner"`）。
- 暗幕: 主の窓全体を rgba(0,0,0,0.3) で覆う（`confirm.md` の 2 章と同じ）。
- 窓の枠: `WindowStyle="None"`・`AllowsTransparency="True"`・`ResizeMode="NoResize"`。

## 3. 部品の木

```xml
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="520" SizeToContent="Height" WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        ResizeMode="NoResize" WindowStartupLocation="CenterOwner" ShowInTaskbar="False">
  <Border Margin="20" Background="{DynamicResource Bg.Surface}" BorderBrush="{DynamicResource Border.Normal}"
          BorderThickness="1" CornerRadius="8">
    <Border.Effect>
      <DropShadowEffect BlurRadius="20" ShadowDepth="4" Direction="270" Opacity="0.18" Color="#000000" />
    </Border.Effect>
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto" />  <!-- title-bar -->
        <RowDefinition Height="Auto" />  <!-- body -->
      </Grid.RowDefinitions>

      <!-- title-bar -->
      <Border Grid.Row="0" x:Name="TitleBar" Background="{DynamicResource Bg.Window}" CornerRadius="8,8,0,0"
              BorderBrush="{DynamicResource Border.Normal}" BorderThickness="0,0,0,1" Padding="16,10,12,10">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*" />
            <ColumnDefinition Width="Auto" />
          </Grid.ColumnDefinitions>
          <TextBlock Grid.Column="0" FontSize="13" FontWeight="Medium" Foreground="{DynamicResource Ink.Value}"
                     VerticalAlignment="Center" Text="{dcstart.title}" />
          <Button x:Name="TitleCloseButton" Grid.Column="1" Width="20" Height="20" Margin="8,0,0,0"
                  Style="{StaticResource TitleBarButton}" IsCancel="False">
            <Image Width="20" Height="20" Source="{StaticResource Icon.X}" />
          </Button>
        </Grid>
      </Border>

      <!-- body（上下 16・左右 20、間 14） -->
      <StackPanel Grid.Row="1" Margin="20,16,20,16" Orientation="Vertical">
        <TextBlock Style="{StaticResource Body}" Foreground="{DynamicResource Ink.Value}" TextWrapping="Wrap"
                   Text="{dcstart.lead}" />

        <!-- folder-table（PlanGrid） -->
        <Border Margin="0,14,0,0" BorderBrush="{DynamicResource Border.Normal}" BorderThickness="1" CornerRadius="5">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto" />  <!-- 見出し -->
              <RowDefinition Height="Auto" />  <!-- 行 -->
              <RowDefinition Height="Auto" />  <!-- 合計 -->
            </Grid.RowDefinitions>
            <Border Grid.Row="0" Background="{DynamicResource Bg.Window}" CornerRadius="5,5,0,0"
                    BorderBrush="{DynamicResource Border.Row}" BorderThickness="0,0,0,1">
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*" />
                  <ColumnDefinition Width="90" />
                  <ColumnDefinition Width="100" />
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Margin="10,7" FontSize="11" FontWeight="Medium" Foreground="{DynamicResource Ink.Value}" Text="{dcstart.col.folder}" />
                <TextBlock Grid.Column="1" Margin="10,7" FontSize="11" FontWeight="Medium" Foreground="{DynamicResource Ink.Value}" TextAlignment="Right" Text="{dcstart.col.count}" />
                <TextBlock Grid.Column="2" Margin="10,7" FontSize="11" FontWeight="Medium" Foreground="{DynamicResource Ink.Value}" TextAlignment="Right" Text="{dcstart.col.status}" />
              </Grid>
            </Border>
            <ItemsControl Grid.Row="1" x:Name="PlanGrid">
              <ItemsControl.ItemTemplate>
                <DataTemplate>
                  <Border BorderBrush="{DynamicResource Border.Row}" BorderThickness="0,0,0,1">
                    <Grid>
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*" />
                        <ColumnDefinition Width="90" />
                        <ColumnDefinition Width="100" />
                      </Grid.ColumnDefinitions>
                      <TextBlock Grid.Column="0" Margin="10,7" Style="{StaticResource Cell}" Foreground="{DynamicResource Ink.Value}"
                                 TextTrimming="CharacterEllipsis" Text="{Binding Name}" />
                      <TextBlock Grid.Column="1" Margin="10,7" Style="{StaticResource Cell}" Foreground="{DynamicResource Ink.Muted}"
                                 TextAlignment="Right" Text="{Binding CountText}" />
                      <Border Grid.Column="2" Margin="10,5" Padding="8,2" CornerRadius="3" HorizontalAlignment="Right"
                              VerticalAlignment="Center" Background="{Binding BadgeBackground}">
                        <TextBlock Style="{StaticResource Chip}" Foreground="{Binding BadgeForeground}" Text="{Binding StatusText}" />
                      </Border>
                    </Grid>
                  </Border>
                </DataTemplate>
              </ItemsControl.ItemTemplate>
            </ItemsControl>
            <Border Grid.Row="2" Background="{DynamicResource Bg.Summary}" CornerRadius="0,0,5,5" Padding="10,8">
              <TextBlock TextWrapping="Wrap">
                <Run x:Name="PlanSummaryRun" FontSize="11" FontWeight="Medium" Foreground="{DynamicResource Ink.Value}" Text="{dcstart.summary}" />
                <Run x:Name="PlanSummaryNoteRun" FontSize="11" Foreground="{DynamicResource Ink.Muted}" Text="{dcstart.summary.note}" />
              </TextBlock>
            </Border>
          </Grid>
        </Border>

        <TextBlock Margin="0,14,0,0" Style="{StaticResource Note}" Foreground="{DynamicResource Ink.Note}" TextWrapping="Wrap"
                   Text="{dcstart.note}" />

        <!-- office-copy-note -->
        <Grid Margin="0,14,0,0">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="*" />
          </Grid.ColumnDefinitions>
          <Image Grid.Column="0" Width="12" Height="12" Margin="0,2,4,0" VerticalAlignment="Top" Source="{StaticResource Icon.Info.Note}" />
          <TextBlock Grid.Column="1" Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Note}" TextWrapping="Wrap"
                     Text="{dcstart.officeCopy}" />
        </Grid>

        <!-- buttons（右寄せ、間 8） -->
        <StackPanel Margin="0,14,0,0" Orientation="Horizontal" HorizontalAlignment="Right">
          <Button x:Name="CancelButton" Style="{StaticResource DialogButton}" Padding="16,7" IsCancel="True" Content="{dcstart.cancel}" />
          <Button x:Name="StartButton" Style="{StaticResource Primary}" Margin="8,0,0,0" Padding="16,7" IsDefault="True" Content="{dcstart.start}" />
        </StackPanel>
      </StackPanel>
    </Grid>
  </Border>
</Window>
```

- 外側の `Margin="20"` は影を描くための透明の余白。見た目の幅は 520。
- 今の XAML の `RetryCheck`（失敗分も再取り込みする）と、広い表の列（取り込み対象・内訳・合計など）は置かない。Figma に無い（「決めること」）。
- 今の XAML の名前 `PlanGrid` は、Figma の `folder-table` の中の行の入れ物に付けた。
- 中の線の太さ・行の高さは Figma の値（セルの padding 10/7、線 1px）。行の高さは中身で決まる。

## 4. 寸法と色

| 要素 | 大きさ | 余白 | 地 | 枠 | 文字の Style | 文字の色 | 角丸 | アイコン |
|---|---|---|---|---|---|---|---|---|
| 外枠 | 520×401（既定）・520×335（1 行） | — | Bg.Surface（#FFFFFF） | 1px Border.Normal（#D9DEE3） | — | — | 8 | — |
| 影 | 0 4 20 rgba(0,0,0,.18) | — | — | — | — | — | — | — |
| TitleBar | Hug | 上下 10・右 12・左 16 | Bg.Window（#F5F7FA） | 下 1px Border.Normal | 13 Medium | Ink.Value（#212126） | 上 8 | 閉じる x 20、線 1.67、色は未定 |
| body | — | 上下 16・左右 20、間 14 | — | — | — | — | — | — |
| 前置き | 幅いっぱい | — | — | — | Body（13） | Ink.Value（#212126） | — | — |
| folder-table | 幅いっぱい | — | — | 1px Border.Normal | — | — | 5 | — |
| 表の見出し | — | セル 10/7 | Bg.Window（#F5F7FA） | 下 1px Border.Row（#EDF0F2） | 11 Medium | Ink.Value | — | — |
| 列 | フォルダ名 Fill・対象ファイル数 90（右揃え）・ステータス 100（右揃え） | — | — | — | — | — | — | — |
| 行 | — | セル 10/7 | Bg.Surface | 下 1px Border.Row | Cell（12） | 名前 Ink.Value・数 Ink.Muted（#6B737D） | — | — |
| バッジ「要更新」 | Hug | 8/2 | Warn.Soft（#FFF5E0） | — | Chip（11 Medium） | Warn（#BA7D00） | 3 | — |
| バッジ「最新」 | Hug | 8/2 | Ok.Soft（#E0F7E0） | — | Chip | Ok（#218A21） | 3 | — |
| 合計の行 | — | 10/8 | Bg.Summary（#F7FAFC） | — | 11 Medium／補足 11 | Ink.Value／補足 Ink.Muted | 下 5 | — |
| 注意 | 幅いっぱい | — | — | — | Note（11/160%） | Ink.Note（#8C949E） | — | — |
| office-copy-note | 幅いっぱい | アイコンと文の間 4 | — | — | Meta（11） | Ink.Note | — | info 12、丸の線 1、中の「i」は線 1.5（README の 9）、Ink.Note |
| CancelButton | Hug | 16/7 | Bg.Surface | 1px Border.Dialog（#D1D6E0） | 13 Medium | Ink.Strong（#202124） | 6 | — |
| StartButton | Hug | 16/7 | Accent（#0078D4） | なし | 13 SemiBold | Ink.OnAccent（#FFFFFF） | 6 | — |

- バッジの文字の大きさは Chip（11 Medium）と読んだ。数値は Figma で確かめる（11 章）。

## 5. 状態ごとの見え方

| 状態 | 入る条件 | 行 | PlanSummaryRun | PlanSummaryNoteRun | StartButton |
|---|---|---|---|---|---|
| 既定・P-DC-start | 更新するフォルダが 1 つ以上、最新のフォルダが 1 つ以上 | 全部（例 3 行） | `dcstart.summary` | 出す `dcstart.summary.note` | 有効 |
| P-DC-start-1 | 更新するフォルダが 1 つ、最新のフォルダが 0 | 1 行 | `dcstart.summary` | 出さない | 有効 |
| 更新が要らない | 更新するフォルダが 0 | 未定 | 未定 | 未定 | 未定 |

- 例の行（既定）:

  | フォルダ名 | 対象ファイル数 | ステータス |
  |---|---|---|
  | 営業部 | 245 | 要更新 |
  | 顧客 | 1,830 | 要更新 |
  | 営業部2025 | 455 | 最新 |

- 合計は「更新対象: 2 フォルダ / 2,075 ファイル」（要更新の行の数とファイル数の合計）。
- 更新するものが 0 のとき（今のコードは「更新が必要なファイルはありません」と［閉じる］）は、Figma に variant が無い（「決めること」）。
- 元のフォルダが見つからない行・チェックを外した行（今の「取り込めません」「取り込みません」）の見え方も Figma に無い（「決めること」）。

## 6. 文言

| ID | 全文 | 使う状態 |
|---|---|---|
| dcstart.title | インデックス更新の確認 | すべて |
| dcstart.lead | 次のフォルダのインデックスを更新します。 | すべて |
| dcstart.col.folder | フォルダ名 | すべて |
| dcstart.col.count | 対象ファイル数 | すべて |
| dcstart.col.status | ステータス | すべて |
| dcstart.row.count | {件数}（3 桁ごとに「,」。単位は付けない。例「1,830」） | すべて |
| dcstart.status.update | 要更新 | すべて |
| dcstart.status.latest | 最新 | 既定・P-DC-start |
| dcstart.summary | 更新対象: {フォルダ数} フォルダ / {ファイル数} ファイル | すべて |
| dcstart.summary.note | （最新のフォルダは更新しません） | 最新の行があるとき |
| dcstart.note | ※ 更新には時間がかかることがあります。途中で中止しても、次回は続きから再開できます。 | すべて |
| dcstart.officeCopy | 更新のとき、Office のファイルはいったんこの PC にコピーしてから読み込み、読み終えたらコピーを削除します。元のファイルは開いたままにしないため、更新中も編集・保存できます。 | すべて |
| dcstart.cancel | キャンセル | すべて |
| dcstart.start | 更新を開始 | すべて |

- 「更新対象:」のコロンは半角で、後ろに半角の空白。「/」は半角で、前後に半角の空白。
- 合計の補足の括弧は全角。合計と補足の間に空白は無い（Figma の 2 つの文字の並び。間の数値は 11 章で確かめる）。

## 7. 操作と振る舞い

| きっかけ | 何が起きるか | 呼ぶ関数 | 次の状態 |
|---|---|---|---|
| 一覧の更新のボタンをクリック | 取り込み予定を作り、ダイアログを出す | `readIngestPlan` → `newPlanViewRows` → `getIndexingConfirmView` | 既定・P-DC-start-1 |
| StartButton・Enter | 閉じて、更新を始める | 今の `showIndexingConfirmDialog` の戻り値 `$true` | インデックス管理の一覧の更新中 |
| CancelButton・Esc・TitleCloseButton | 閉じて、何もしない | 戻り値 `$false` | そのまま |

- 既定のボタン: StartButton（`IsDefault`）。Esc で閉じる: する（CancelButton が `IsCancel`）。今の XAML と同じ。
- ボタンの並び: 左から［キャンセル］［更新を開始］。
- Tab の順: CancelButton → StartButton → TitleCloseButton。表の行はフォーカスを受けない。
- ツールチップ: フォルダ名が「…」で切れたときに全文と元のフォルダのパスを出すかは未定。

## 8. リサイズ

- 幅 520 で固定。高さは中身に合わせる（`SizeToContent="Height"`）。最小の高さは今の中身の高さ。
- 主の窓の大きさにかかわらず、同じ大きさで主の窓の中央に出す。暗幕は主の窓全体を覆う。
- 題は 1 行。前置き・注意・office-copy-note は折り返す。
- フォルダ名は 1 行で「…」で切る。
- 行が多くて主の窓に収まらないときの扱い（表の最大の高さとスクロール）は Figma に無い（「決めること」）。

## 9. 判断層（`ui/indexing_view.ps1`）

### newPlanViewRows（形を変える）

- 入力: `$plan`（`readIngestPlan` の結果。各要素は `インデックス名`・`元のフォルダ`・`区分`・`ファイル数`）
- 出力: 行の配列。各要素 `@{ Name [string]; Path [string]; CountText [string]; Status = "update" | "latest"; StatusText [string] }`
  - 今の `TargetText`・`DetailText`・`TotalText`・`Tone` は使わない。
  - 未チェック・見つからない区分の扱いは「決めること」。決まるまで表から外す。

| 入力（名前, 区分, ファイル数） | Name | CountText | Status | StatusText |
|---|---|---|---|---|
| 営業部, 更新あり, 245 | 営業部 | 245 | update | 要更新 |
| 顧客, 新規, 1830 | 顧客 | 1,830 | update | 要更新 |
| 営業部2025, 更新なし, 455 | 営業部2025 | 455 | latest | 最新 |
| （空の配列） | 行なし | — | — | — |

- 区分の名前（「更新あり」「新規」「更新なし」）は今の `${planKind*}` の値に合わせる。どの区分を「要更新」とするかは、今の区分の定義に従う（`latest` は取り込むファイルが 0 のもの）。

### getIndexingConfirmView（`getIndexingConfirmText` を置き換える）

- 入力: `$rows`（`newPlanViewRows` の結果）
- 出力: `@{ SummaryText [string]; SummaryNote [string]; FolderCount [int]; FileCount [int] }`

| rows | SummaryText | SummaryNote | FolderCount | FileCount |
|---|---|---|---|---|
| 営業部 245 update・顧客 1830 update・営業部2025 455 latest | 更新対象: 2 フォルダ / 2,075 ファイル | （最新のフォルダは更新しません） | 2 | 2075 |
| 営業部 245 update | 更新対象: 1 フォルダ / 245 ファイル | `""` | 1 | 245 |
| 営業部2025 455 latest | 未定 | 未定 | 0 | 0 |

- 今の `getIndexingConfirmText` の `retryFailed`（失敗分も再取り込みする）は、RetryCheck を置かないなら要らない（「決めること」）。

## 10. 画面層（`ui/indexing_tab.ps1` の `showIndexingConfirmDialog`）

- 取り込み予定の作成（`readIngestPlan`）は今のとおり。重いときは別スレッドで作り、Dispatcher でダイアログを出す。
- `XamlReader::Load` → Owner を主の窓 → `PlanGrid.ItemsSource` に `newPlanViewRows` の行を入れる。バッジの色は `Status` から画面層の対応表で引く（update → Warn.Soft／Warn、latest → Ok.Soft／Ok）。
- `getIndexingConfirmView` の出力を `PlanSummaryRun`・`PlanSummaryNoteRun` に写す。`SummaryNote` が空なら `PlanSummaryNoteRun` を消す。
- 暗幕を出す（`confirm.md` の `showScrim`）→ `ShowDialog()` → 閉じたら暗幕を消す。
- `StartButton.Add_Click` で `DialogResult = $true`。`TitleCloseButton.Add_Click` で `DialogResult = $false`。

## 11. 受け入れ

見比べる Figma: 既定 340:1544、P-DC-start 372:788、P-DC-start-1 372:948（`../png/` に画像は無い。get_screenshot で撮って並べる）。

- 幅 520、既定で高さ 401、1 行で 335。
- 題のバーが淡い灰（Bg.Window）で、題が 13 Medium。
- 表の見出しが淡い灰、列が「フォルダ名」（伸びる）・「対象ファイル数」90・「ステータス」100 で、数とステータスが右揃え。
- 「要更新」が淡い黄の地に黄の文字、「最新」が淡い緑の地に緑の文字。
- 合計の行が「更新対象: 2 フォルダ / 2,075 ファイル（最新のフォルダは更新しません）」で、補足だけ薄い色。
- 注意の文と、info アイコン付きの Office のコピーの文がある。
- 右下に［キャンセル］［更新を開始］。

## 決めること

- 今の 880 幅の表（取り込み対象・内訳・合計の列）を、Figma の 3 列に減らしてよいか。
- 今の RetryCheck（失敗分も再取り込みする）を消してよいか。消さないなら置き場所。
- 題が「インデックス更新の確認」・ボタンが「更新を開始」。今のコードの「インデックス作成」の言い方を、ほかの画面も含めて「更新」にそろえるか。
- 更新するものが 0 のときの見え方（今は「更新が必要なファイルはありません」と［閉じる］）。
- 元のフォルダが見つからない行・チェックを外した行の見え方。
- 行が多いときの表の最大の高さとスクロール。
- 閉じる x の色、フォルダ名のツールチップ。
- P-DC-start は、スクリーンショットだけで既定と同じと確かめた。数値は読んでいない。
