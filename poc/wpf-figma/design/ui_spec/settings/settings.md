# 設定（settings）

## 1. 役割

インデックスとログの保存先、インデックス作成のときの並列処理数、検索結果の 1 ページの件数を決める画面。保存先を変えると、今のインデックスとログを新しい場所へ移す（確認と進み具合は `../dialog/confirm.md` の DC-ws・DC-ws-idx・DC-reset・DC-move・E54）。

- Figma の部品: `xaml/settings/settings` 352:44948。variant は次の 4 つで、どれも 1060×767。

  | variant | ノード | 使う画面 |
  |---|---|---|
  | 既定 | 340:1241 | 106 の C |
  | C | 352:44869 | 106 の C |
  | P-C-moved | 352:44908 | 148 の P-C-moved（保存先を移した直後） |
  | P-C-reset | 352:44947 | 148 の P-C-reset（既定の場所に戻した直後） |

- リサイズの見本: C 1024×640（408:33839）・C 1600×1000（408:33865）。
- 次の項目はメンテナの決定で、この画面には置かない。
  - 検索の条件の 4 項目
  - 保存の文字コード
  - 自動更新
  - 検索結果のプレビュー
- `../png/` にこの画面の画像は無い。見比べは Figma の get_screenshot で行う（11 章）。

## 2. 置き場所

- XAML: `scripts/tebunko/xaml/settings/settings.xaml`（今の `tab_settings.xaml` を置き換える）。
- 画面層 `ui/settings/settings.ps1`、判断層 `ui/settings/settings_view.ps1`。今の `ui/settings_tab.ps1`・`ui/settings_view.ps1` から移す。
- 親: shell の `ContentHost`（`../shell/shell.md`）。ナビで「設定」（`SettingsTab`）を選んだときに差す。
- 読み込み: 起動のときに 1 回だけ `XamlReader::Load` で読み、`FindName` で x:Name を引く。画面を切り替えても作り直さない。

## 3. 部品の木

```xml
<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
      x:Name="SettingsRoot" Background="{DynamicResource Bg.Surface}">
  <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
    <StackPanel Orientation="Vertical">

      <!-- 見出し（header） -->
      <StackPanel Margin="24,16,24,0" Orientation="Vertical">
        <TextBlock Style="{StaticResource Title}" Foreground="{DynamicResource Ink.Value}" Text="{settings.title}" />
        <TextBlock Style="{StaticResource Body}" Foreground="{DynamicResource Ink.Muted}" Margin="0,4,0,0"
                   TextWrapping="Wrap" Text="{settings.description}" />
      </StackPanel>

      <!-- 区画（sections）。区画の間は 24、左右と下の余白は 24 -->
      <StackPanel Margin="24,24,24,24" Orientation="Vertical">

        <!-- 区画: インデックス設定 -->
        <StackPanel Orientation="Vertical">
          <TextBlock Style="{StaticResource Focal}" Foreground="{DynamicResource Ink.Value}" Text="{settings.section.index}" />
          <Border Height="1" Margin="0,8,0,0" Background="{DynamicResource Bg.Section}" /> <!-- 見出しと線の間は未定（決めること） -->

          <!-- 行: 保存先 -->
          <Grid Margin="0,12,0,0" MinHeight="33">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*" />     <!-- 左の文（伸びる） -->
              <ColumnDefinition Width="Auto" />  <!-- 右の操作 -->
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0" Orientation="Vertical" VerticalAlignment="Center" Margin="0,0,16,0">
              <TextBlock Style="{StaticResource Nav}" Foreground="{DynamicResource Ink.Value}" Text="{settings.workspace.label}" />
              <TextBlock Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Muted}" Margin="0,2,0,0"
                         TextWrapping="Wrap" Text="{settings.workspace.sub}" />
            </StackPanel>
            <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
              <TextBlock x:Name="WorkspaceDoneText" Style="{StaticResource Cell}" FontWeight="Medium"
                         Foreground="{DynamicResource Ok}" VerticalAlignment="Center" Margin="0,0,8,0"
                         Visibility="Collapsed" Text="{settings.workspace.done}" />
              <Border Background="{DynamicResource Bg.Subtle}" BorderBrush="{DynamicResource Border.Soft}"
                      BorderThickness="1" CornerRadius="4" Padding="10,6" VerticalAlignment="Center">
                <TextBlock x:Name="WorkspaceText" Style="{StaticResource Cell}" Foreground="{DynamicResource Ink.Strong}"
                           TextTrimming="CharacterEllipsis" Text="{settings.workspace.path}" />
              </Border>
              <Button x:Name="ChangeWorkspaceButton" Style="{StaticResource Button.Base}" Margin="8,0,0,0"
                      Height="30" Padding="12,6" Content="{settings.workspace.change}" />
              <Button x:Name="ResetWorkspaceButton" Style="{StaticResource Button.Base}" Margin="8,0,0,0"
                      Height="30" Padding="12,6" Content="{settings.workspace.reset}" />
            </StackPanel>
          </Grid>

          <!-- 行: 最大並列処理数 -->
          <Grid Margin="0,12,0,0" MinHeight="33">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*" />
              <ColumnDefinition Width="Auto" />
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0" Orientation="Vertical" VerticalAlignment="Center" Margin="0,0,16,0">
              <TextBlock Style="{StaticResource Nav}" Foreground="{DynamicResource Ink.Value}" Text="{settings.parallel.label}" />
              <TextBlock Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Muted}" Margin="0,2,0,0"
                         TextWrapping="Wrap" Text="{settings.parallel.sub}" />
            </StackPanel>
            <ComboBox x:Name="MaxParallelCombo" Grid.Column="1" Style="{StaticResource ComboBox.Base}"
                      Width="160" Height="28" VerticalAlignment="Center" />
          </Grid>
        </StackPanel>

        <!-- 区画: 表示設定 -->
        <StackPanel Orientation="Vertical" Margin="0,24,0,0">
          <TextBlock Style="{StaticResource Focal}" Foreground="{DynamicResource Ink.Value}" Text="{settings.section.view}" />
          <Border Height="1" Margin="0,8,0,0" Background="{DynamicResource Bg.Section}" />

          <!-- 行: 1ページあたりの表示件数 -->
          <Grid Margin="0,12,0,0" MinHeight="33">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*" />
              <ColumnDefinition Width="Auto" />
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0" Orientation="Vertical" VerticalAlignment="Center" Margin="0,0,16,0">
              <TextBlock Style="{StaticResource Nav}" Foreground="{DynamicResource Ink.Value}" Text="{settings.pagesize.label}" />
              <TextBlock Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Muted}" Margin="0,2,0,0"
                         TextWrapping="Wrap" Text="{settings.pagesize.sub}" />
            </StackPanel>
            <ComboBox x:Name="PageSizeCombo" Grid.Column="1" Style="{StaticResource ComboBox.Base}"
                      Width="160" Height="28" VerticalAlignment="Center" />
          </Grid>
        </StackPanel>

      </StackPanel>
    </StackPanel>
  </ScrollViewer>
</Grid>
```

- `Nav`（13 Medium）を行の見出しに使う。Figma の行の見出しは 13 Medium で、Text style の名前は付いていない。
- `ComboBox.Base` は theme にまだ無いキー（「決めること」）。
- 見出しと区切り線の間の余白は、Figma の数値を読み取れていない（未定）。上の 8 は仮の値ではなく、置き場所を示すだけ。
- 今の XAML の `WorkspaceNote`・`SettingsFileText`・`SettingsFileNote` は置かない。Figma に無い（「決めること」）。

## 4. 寸法と色

| 要素 | 大きさ | 余白 | 地 | 枠 | 文字の Style | 文字の色 | 角丸 |
|---|---|---|---|---|---|---|---|
| SettingsRoot | 主な領域いっぱい（1280 幅で 1060×767） | — | Bg.Surface（#FFFFFF） | — | — | — | — |
| 見出しの区画 | 幅いっぱい | 左右 24・上下 16 | — | — | — | — | — |
| 題「設定」 | Hug | — | — | — | Title（18 SemiBold） | Ink.Value（#212126） | — |
| 説明 | 幅いっぱい | 上 4 | — | — | Body（13） | Ink.Muted（#6B737D） | — |
| 区画の見出し | Hug | — | — | — | Focal（14 SemiBold） | Ink.Value（#212126） | — |
| 区画の線 | 高さ 1・幅いっぱい | — | Bg.Section（#E8EBF0） | — | — | — | — |
| 行 | 高さ 33（中身に合わせて伸びる） | 行の間 12 | — | — | — | — | — |
| 行の見出し | Hug | — | — | — | 13 Medium（Nav） | Ink.Value（#212126） | — |
| 行の補足 | 幅いっぱい | 上 2 | — | — | Meta（11） | Ink.Muted（#6B737D） | — |
| パスの箱（WorkspaceText の Border） | Hug（最大幅は未定） | 10/6 | Bg.Subtle（#F9FAFA） | 1px Border.Soft（#E0E2E5） | Cell（12） | Ink.Strong（#202124） | 4 |
| ChangeWorkspaceButton | 58×30 | 12/6 | Bg.Surface（#FFFFFF） | 1px Border.Soft（#E0E2E5） | 12 SemiBold | Ink.Strong（#202124） | 6 |
| ResetWorkspaceButton | 86×30 | 12/6 | Bg.Surface | 1px Border.Soft | 12 SemiBold | Ink.Strong | 6 |
| ボタンの間 | — | 8 | — | — | — | — | — |
| WorkspaceDoneText | Hug | — | — | — | 12 Medium | Ok（#218A21） | — |
| MaxParallelCombo・PageSizeCombo | 160×28 | 左 10・右 8・上下 5 | Bg.Surface（#FFFFFF） | 1px Border.Normal（#D9DEE3） | Cell（12） | Ink.Strong（#202124） | 4 |
| コンボの ▼ | 6×3 | — | — | — | — | 未定 | — |

- 区画の上端は y0 と y145（既定の 1060×767 のとき）。
- アイコンは使わない（コンボの ▼ は Lucide ではなく三角形の Path。色は未定）。

## 5. 状態ごとの見え方

| 状態 | 入る条件 | WorkspaceText | WorkspaceDoneText | ResetWorkspaceButton | MaxParallelCombo | PageSizeCombo |
|---|---|---|---|---|---|---|
| 既定・C | 画面を開いた | 今の保存先（例 `C:\Tools\tebunko\work`） | 出さない | 出す・有効 | 今の値（例「4」） | 今の値（例「50 件」） |
| P-C-moved | 保存先の移動（DC-ws → DC-move）が終わった | 新しい保存先（例 `D:\tebunko_index`） | 出す `settings.workspace.done` | 出す・有効 | 同上 | 同上 |
| P-C-reset | 既定に戻す移動（DC-reset → DC-move）が終わった | 既定の保存先 | 出す `settings.workspace.done` | 出す・有効 | 同上 | 同上 |
| 移動中 | DC-move を出している | そのまま | 出さない | 暗幕の下（操作できない） | 同左 | 同左 |

- Figma では、既定の場所にいるとき（既定・C）も［既定に戻す］が出ている。今のコードは `CanReset` が偽のとき隠す。どちらにするかは未定（「決めること」）。
- WorkspaceDoneText をいつ消すか（画面を切り替えたら・時間で）は未定。

## 6. 文言

| ID | 全文 | 使う状態 |
|---|---|---|
| settings.title | 設定 | すべて |
| settings.description | インデックスの保存先と表示を設定します。 | すべて |
| settings.section.index | インデックス設定 | すべて |
| settings.workspace.label | 保存先 | すべて |
| settings.workspace.sub | インデックスとログを置くフォルダ | すべて |
| settings.workspace.path | `{保存先のパス}`（例 `C:\Tools\tebunko\work`・`D:\tebunko_index`） | すべて |
| settings.workspace.change | 変更… | すべて |
| settings.workspace.reset | 既定に戻す | すべて |
| settings.workspace.done | 移動しました | P-C-moved・P-C-reset |
| settings.parallel.label | 最大並列処理数 | すべて |
| settings.parallel.sub | インデックス作成時の並列処理数 | すべて |
| settings.parallel.value | `{数}`（例「4」） | すべて |
| settings.section.view | 表示設定 | すべて |
| settings.pagesize.label | 1ページあたりの表示件数 | すべて |
| settings.pagesize.sub | 検索結果の1ページに表示する件数 | すべて |
| settings.pagesize.value | `{件数} 件`（例「50 件」） | すべて |

- 「1ページあたり」「1ページに」は、数字と「ページ」の間に空白を入れない（Figma のとおり）。「50 件」は数字と「件」の間に半角の空白を入れる。
- 確認・進み具合・エラーの文言は `../dialog/confirm.md`（DC-ws・DC-ws-idx・DC-reset・DC-move・E54）。

## 7. 操作と振る舞い

| きっかけ | 何が起きるか | 呼ぶ関数 | 次の状態 |
|---|---|---|---|
| ChangeWorkspaceButton をクリック | Windows の標準のフォルダの選択（`selectFolder`、IFileOpenDialog）を開く | `selectFolder` | 選んだら下へ。取りやめたら何もしない |
| フォルダを選んだ | 可否を調べる | `testWorkspaceChoice` | `same` → 何もしない。`error` → E54。`ok` → 中にインデックスがあれば DC-ws-idx、無ければ DC-ws |
| DC-ws で［移動する］ | 移動を始める（別スレッド） | `moveWorkspace`（状態層） | DC-move → 終わったら P-C-moved |
| DC-ws-idx で［そのフォルダのインデックスを使う］ | 選んだフォルダを保存先にする（今の中身は移さない） | `newWorkspaceConfirm` の Value `use` | 終わったら P-C-moved |
| DC-ws-idx で［今のインデックスを移動する］ | 今のインデックスを移す | Value は未定（`reset` に当たるかは「決めること」） | DC-move → P-C-moved |
| E54 で［別のフォルダを選ぶ］ | フォルダの選択をもう一度開く | `selectFolder` | 上と同じ |
| ResetWorkspaceButton をクリック | DC-reset を出す | `getWorkspaceView`（既定の場所を得る） | ［戻す］→ DC-move → P-C-reset |
| MaxParallelCombo で選ぶ | 設定ファイルに書く | 書き込みの関数は未定 | そのまま |
| PageSizeCombo で選ぶ | 設定ファイルに書き、検索結果の 1 ページの件数を変える | 書き込みの関数は未定 | そのまま |

- ボタンのホバー・押したときの見た目は `Button.Base` に従う（Figma に variant が無い）。
- Tab の順: ChangeWorkspaceButton → ResetWorkspaceButton → MaxParallelCombo → PageSizeCombo。
- ツールチップ: 無し（Figma に無い）。WorkspaceText が「…」で切れたときに全文をツールチップで出すかは未定。

## 8. リサイズ

| 窓 | 主な領域 | 見え方 |
|---|---|---|
| 1024×640（最小） | 804×608 | 行の左の文が縮む。右のパスの箱・ボタン・コンボは右端に付いたまま。補足は折り返す |
| 1280×820（基本） | 1060×767 | Figma の既定のとおり |
| 1600×1000 | 1380 幅 | 左の文が伸びる。右の操作は右端に付いたまま |

- 縦に収まらないときは `ScrollViewer` が縦にスクロールする（横は出さない）。
- 見出しの左右 24・上下 16、区画の左右と下 24、区画の間 24、行の間 12 は窓の大きさで変えない。
- パスの箱は中身の幅（Hug）。長いパスで右の操作が左の文を押しつぶさないよう最大幅を付けるかは未定（「決めること」）。付けるなら `TextTrimming="CharacterEllipsis"` で切る。

## 9. 判断層（`ui/settings/settings_view.ps1`）

### getWorkspaceView

- 入力: `[string]$workDir`（今の保存先）、`[string]$defaultDir`（既定の保存先）
- 出力: `@{ Path [string]; IsDefault [bool]; CanReset [bool] }`
  - 今の `Note` は画面に出さない（Figma に無い）。
  - `CanReset` を Figma のとおり常に `$true` にするかは「決めること」。下の表は今の決まり（既定なら `$false`）で書く。

| workDir | defaultDir | Path | IsDefault | CanReset |
|---|---|---|---|---|
| `C:\Tools\tebunko\work` | `C:\Tools\tebunko\work` | `C:\Tools\tebunko\work` | `$true` | `$false` |
| `C:\Tools\tebunko\work\` | `C:\Tools\tebunko\work` | `C:\Tools\tebunko\work\` | `$true` | `$false` |
| `c:\tools\TEBUNKO\work` | `C:\Tools\tebunko\work` | `c:\tools\TEBUNKO\work` | `$true` | `$false` |
| `D:\tebunko_index` | `C:\Tools\tebunko\work` | `D:\tebunko_index` | `$false` | `$true` |

### testWorkspaceChoice（今のまま）

- 入力: `$folder`・`$current`・`[bool]$writable`・`$drives`
- 出力: `@{ Kind = "same" | "error" | "ok"; Message [string] }`。`error` のとき、画面層は E54 の detail-box に `Message` を出す。

| folder | current | writable | Kind | Message |
|---|---|---|---|---|
| `""` | `C:\Tools\tebunko\work` | `$true` | error | フォルダを選んでください。 |
| `C:\Tools\tebunko\work` | `C:\Tools\tebunko\work` | `$true` | same | `""` |
| `C:\Tools\tebunko\work\index\営業部` | `C:\Tools\tebunko\work` | `$true` | error | 「C:\Tools\tebunko\work\index\営業部」は今のインデックスのフォルダの中です。インデックスの外のフォルダを選んでください。 |
| `D:\tebunko_index` | `C:\Tools\tebunko\work` | `$false` | error | 「D:\tebunko_index」にはファイルを作れません。書き込めるフォルダを選んでください。 |
| `D:\tebunko_index` | `C:\Tools\tebunko\work` | `$true` | ok | `""` |

- Figma の E54 の detail-box は「書き込みが拒否されました」。今の Message と文言が違う。どちらを出すかは「決めること」。

### newWorkspaceConfirm（形を変える）

- 入力: 今のまま（`$folder`・`$current`・`$entryCount`・`$sampleNames`・`$countCapped`・`$workspaceNames`・`$canMakeSub`）。
- 出力: 確認の中身。形は `../dialog/confirm.md` の 9 章の `newConfirmView` の出力と同じ `@{ State; Type; Title; Message; Detail; DetailBox; Buttons }`。今の `Heading`・`Facts`・`Hint`・`Choices` は使わない。

| 入力の要点 | State | Type | 主のボタン（Value） |
|---|---|---|---|
| workspaceNames が 1 つ以上 | DC-ws-idx | choice | そのフォルダのインデックスを使う（use） |
| entryCount = 0 | DC-ws | normal | 移動する（change） |
| entryCount ≥ 1・workspaceNames なし | 未定 | 未定 | 未定 |

- 空でないフォルダを選んだとき（今の `sub`・`asis` の選択）は、Figma に当たる確認が無い（「決めること」）。

### getParallelOptions・getPageSizeOptions（新しい）

- 出力: コンボに並べる値の配列と、今の値。並べる値と設定ファイルのキーは Figma から読めない（未定。「決めること」）。テストの表は値が決まってから書く。

## 10. 画面層（`ui/settings/settings.ps1`）

- 画面を開いたとき: `getWorkspaceView $workDir $defaultDir` → `WorkspaceText.Text = Path`、`ResetWorkspaceButton.Visibility` を `CanReset` から決める。コンボに今の値を選ぶ。
- `ChangeWorkspaceButton.Add_Click`: `selectFolder` → `testWorkspaceWritable`（今のコードの書き込みの確かめ）→ `testWorkspaceChoice` → `newWorkspaceConfirm` → `showConfirm`（`../dialog/confirm.md`）。
- 移動（`moveWorkspace`）は別スレッドで行う（`startJob`）。進み具合は DC-move に Dispatcher で書く。終わったら Dispatcher で `WorkspaceText` を書き直し、`WorkspaceDoneText` を出す。
- `ResetWorkspaceButton.Add_Click`: DC-reset を出し、［戻す］なら上と同じ移動を既定の場所へ行う。
- `MaxParallelCombo.Add_SelectionChanged`・`PageSizeCombo.Add_SelectionChanged`: 設定ファイルに書く。画面のスレッドで書いてよい（小さな書き込み）。

## 11. 受け入れ

見比べる Figma: 既定 340:1241、P-C-moved 352:44908、P-C-reset 352:44947、C 1024×640 408:33839、C 1600×1000 408:33865（`../png/` に画像は無い。get_screenshot で撮って並べる）。

- 題「設定」が左 24・上 16 にあり、説明との間が 4。
- 「インデックス設定」と「表示設定」の区画の上端の差が 145（1280×820 のとき）。
- 行の高さが 33、行の間が 12。
- 保存先の行の右に、パスの箱・［変更…］（58×30）・［既定に戻す］（86×30）が 8 の間で並ぶ。
- コンボが 160×28 で右端に付く。
- P-C-moved・P-C-reset で「移動しました」が緑（Ok #218A21）で出る。
- 1024×640 で横のスクロールバーが出ない。右の操作が右端に付いたまま。
- 設定の画面に、検索の条件・保存の文字コード・自動更新・プレビューの項目が無い。

## 決めること

- ［既定に戻す］を、既定の場所にいるときも出すか（Figma は出している。今のコードは隠す）。
- P-C-reset の完了の文言が P-C-moved と同じ「移動しました」でよいか。
- WorkspaceDoneText をいつ消すか。
- MaxParallelCombo・PageSizeCombo に並べる値と、設定ファイル（`setting.config`）のキー。書き込みの関数。
- `ComboBox.Base` の Style と、▼ の色（theme にキーが無い）。
- 今の `WorkspaceNote`（既定かどうかの補足）・`SettingsFileText`・`SettingsFileNote`（設定ファイルの場所）を消してよいか。
- パスの箱の最大幅と、切れたときのツールチップ。
- 区画の見出しと区切り線の間の余白（Figma から読み取れていない）。
- 空でないフォルダを選んだとき（今の「中に workspace を作る」「そのまま使う」）の確認。Figma に当たるダイアログが無い。
- DC-ws-idx の［今のインデックスを移動する］が、今の Value `reset`（選んだフォルダのインデックスを消して移す）に当たるか。
- E54 の detail-box に、Figma の「書き込みが拒否されました」と、今の `testWorkspaceChoice` の Message のどちらを出すか。
