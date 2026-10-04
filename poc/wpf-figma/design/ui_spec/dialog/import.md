# インデックスのインポート（import）

## 1. 役割

エクスポートした zip を読み込んで、インデックスに加えるダイアログ（DI）。zip の中身（名前・日時・数・大きさ・版）を示し、名前と元のフォルダを決めさせる。同じ名前があるとき（DI-2）、元のフォルダが見つからないとき（DI-3）、元のフォルダがほかと重なるとき（DI-4）を、その場で示す。

- Figma の部品: `xaml/dialog/import` 372:1112。

  | variant | ノード | 大きさ | 使う画面 |
  |---|---|---|---|
  | 既定 | 340:1597 | 520×346 | 106・148 の DI（元のフォルダが見つかった） |
  | DI-2 | 372:1216 | 520×422 | 106・148（同じ名前がある） |
  | DI-3 | 372:1338 | 520×378 | 106（元のフォルダが見つからない） |
  | DI-4 | 372:1407 | 520×346 | 106（元のフォルダが重なる） |
  | P-DI-1 | 372:1475 | 520×346 | 148（数値は読んでいない。既定と同じ大きさ） |

- 開く所: インデックス管理の一覧の見出しの［インポート］（`ImportIndexButton`、`../index/index_list.md`）。先に Windows の標準のファイルの選択で zip を選ぶ。
- 読めない zip・新しい版の zip・失敗は、`confirm.md` の DI-E1〜DI-E3 で示す。
- `../png/` にこのダイアログの画像は無い。見比べは Figma の get_screenshot で行う（11 章）。

## 2. 置き場所

- XAML: `scripts/tebunko/xaml/dialog_import.xaml`（新しいファイル）。
- 画面層 `ui/index/import_dialog.ps1`、判断層 `ui/index/import_dialog_view.ps1`（新しい）。今の `ui/index_view.ps1` の `testIndexImportInput`・`getImportSuggestedName`・`testImportNameCollision`・`getIndexImportNotice` を移す。
- 読み込みは状態層の `importIndex`（`index/index_archive.ps1`）。
- 型: Window。`ShowDialog()` で出す。Owner は主の窓。
- 置き方: 主の窓の中央。暗幕: 主の窓全体を rgba(0,0,0,0.3) で覆う（`confirm.md` の 2 章と同じ）。
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
        <RowDefinition Height="Auto" />
        <RowDefinition Height="Auto" />
      </Grid.RowDefinitions>

      <!-- title-bar -->
      <Border Grid.Row="0" Background="{DynamicResource Bg.Window}" CornerRadius="8,8,0,0"
              BorderBrush="{DynamicResource Border.Normal}" BorderThickness="0,0,0,1" Padding="16,10,16,10">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*" />
            <ColumnDefinition Width="Auto" />
          </Grid.ColumnDefinitions>
          <TextBlock FontSize="13" FontWeight="Medium" Foreground="{DynamicResource Ink.Value}" VerticalAlignment="Center" Text="{di.title}" />
          <Button x:Name="TitleCloseButton" Grid.Column="1" Width="20" Height="20" Style="{StaticResource TitleBarButton}">
            <Image Width="20" Height="20" Source="{StaticResource Icon.X}" />
          </Button>
        </Grid>
      </Border>

      <!-- body -->
      <StackPanel Grid.Row="1" Margin="20,16,20,16" Orientation="Vertical">
        <!-- zip-info -->
        <Border Background="{DynamicResource Bg.Subtle}" BorderBrush="{DynamicResource Border.Soft}" BorderThickness="1"
                CornerRadius="6" Padding="12,10">
          <StackPanel Orientation="Vertical">
            <TextBlock x:Name="ZipNameText" Style="{StaticResource Heading}" Foreground="{DynamicResource Ink.Strong}"
                       TextTrimming="CharacterEllipsis" Text="{di.zip.name}" />
            <TextBlock x:Name="ZipDateText" Margin="0,4,0,0" Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Body}" Text="{di.zip.date}" />
            <TextBlock x:Name="ZipSizeText" Margin="0,4,0,0" Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Body}" Text="{di.zip.size}" />
          </StackPanel>
        </Border>

        <!-- 名前 -->
        <StackPanel Margin="0,14,0,0" Orientation="Vertical">
          <TextBlock FontSize="11" FontWeight="Medium" Foreground="{DynamicResource Ink.Body}" Text="{di.name.label}" />
          <TextBox x:Name="NameBox" Margin="0,6,0,0" Width="300" Height="28" HorizontalAlignment="Left"
                   Style="{StaticResource TextBox.Base}" Padding="8,0" VerticalContentAlignment="Center"
                   BorderBrush="{DynamicResource Border.Input}" Text="{di.name.value}" />
          <!-- DI-2 だけ出す -->
          <StackPanel x:Name="NameConflictPanel" Margin="0,6,0,0" Orientation="Vertical" Visibility="Collapsed">
            <TextBlock FontSize="11" FontWeight="Medium" Foreground="{DynamicResource Error.Text}" Text="{di.name.conflict}" />
            <StackPanel Margin="0,8,0,0" Orientation="Horizontal">
              <RadioButton x:Name="RenameRadio" GroupName="ImportCollision" Style="{StaticResource Radio}" IsChecked="True"
                           VerticalAlignment="Center" Content="{di.rename}" />
              <TextBox x:Name="RenameBox" Margin="8,0,0,0" Width="140" Height="28" Style="{StaticResource TextBox.Base}"
                       Padding="8,0" VerticalContentAlignment="Center" BorderBrush="{DynamicResource Border.Input}" Text="{di.rename.value}" />
            </StackPanel>
            <StackPanel Margin="0,8,0,0" Orientation="Horizontal">
              <RadioButton x:Name="OverwriteRadio" GroupName="ImportCollision" Style="{StaticResource Radio}"
                           VerticalAlignment="Center" Content="{di.overwrite}" />
              <TextBlock Margin="8,0,0,0" FontSize="12" Foreground="{DynamicResource Error.Text}" VerticalAlignment="Center"
                         Text="{di.overwrite.note}" />
            </StackPanel>
          </StackPanel>
        </StackPanel>

        <!-- 元のフォルダ -->
        <StackPanel Margin="0,14,0,0" Orientation="Vertical">
          <TextBlock FontSize="11" FontWeight="Medium" Foreground="{DynamicResource Ink.Body}" Text="{di.folder.label}" />
          <StackPanel Margin="0,6,0,0" Orientation="Horizontal">
            <TextBox x:Name="FolderBox" Width="380" Height="28" Style="{StaticResource TextBox.Base}" Padding="8,0"
                     VerticalContentAlignment="Center" BorderBrush="{DynamicResource Border.Input}" Text="{di.folder.value}" />
            <Button x:Name="FolderBrowseButton" Margin="8,0,0,0" Style="{StaticResource DialogButton}" Padding="16,7" Content="{di.browse}" />
          </StackPanel>
          <!-- status-line（既定・DI-4） -->
          <StackPanel x:Name="FolderStatusPanel" Margin="0,6,0,0" Orientation="Horizontal">
            <Image x:Name="FolderStatusIcon" Width="14" Height="14" VerticalAlignment="Center" Source="{StaticResource Icon.Check.Ok}" />
            <TextBlock x:Name="FolderStatusText" Margin="5,0,0,0" FontSize="11" FontWeight="Medium"
                       Foreground="{DynamicResource Ok.Strong}" VerticalAlignment="Center" Text="{di.folder.found}" />
          </StackPanel>
          <!-- warn box（DI-3） -->
          <Border x:Name="FolderMissingPanel" Margin="0,6,0,0" Visibility="Collapsed" Background="{DynamicResource Warn.Note}"
                  BorderBrush="{DynamicResource Warn.Line}" BorderThickness="1" CornerRadius="6" Padding="10,8">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto" />
                <ColumnDefinition Width="*" />
              </Grid.ColumnDefinitions>
              <Image Width="14" Height="14" Margin="0,1,6,0" VerticalAlignment="Top" Source="{StaticResource Icon.TriangleAlert.Warn}" />
              <TextBlock Grid.Column="1" Style="{StaticResource Meta}" Foreground="{DynamicResource Warn.Strong}" TextWrapping="Wrap"
                         Text="{di.folder.missing}" />
            </Grid>
          </Border>
        </StackPanel>

        <!-- buttons -->
        <StackPanel Margin="0,14,0,0" Orientation="Horizontal" HorizontalAlignment="Right">
          <Button x:Name="CancelButton" Style="{StaticResource DialogButton}" Padding="16,7" IsCancel="True" Content="{di.cancel}" />
          <Button x:Name="ImportButton" Margin="8,0,0,0" Style="{StaticResource Primary}" Padding="16,7" IsDefault="True" Content="{di.import}" />
        </StackPanel>
      </StackPanel>
    </Grid>
  </Border>
</Window>
```

- `Error.Text`（#B3261E）と `Radio` の Style は theme に無いキー（「決めること」）。
- body の余白（上下 16・左右 20）と区画の間（14）は indexing_confirm と同じ値を入れた。import の Figma では読み取れていない（「決めること」）。
- 「参照…」は Figma で x:Name が付いていない。`FolderBrowseButton` は新しく付けた名前。DI-2 の部品（`NameConflictPanel`・`RenameRadio`・`RenameBox`・`OverwriteRadio`）、`FolderStatusPanel`・`FolderMissingPanel` も同じ。

## 4. 寸法と色

| 要素 | 大きさ | 余白 | 地 | 枠 | 文字の Style | 文字の色 | 角丸 | アイコン |
|---|---|---|---|---|---|---|---|---|
| 外枠 | 幅 520・高さ Hug | — | Bg.Surface（#FFFFFF） | 1px Border.Normal（#D9DEE3） | — | — | 8 | — |
| title-bar | Hug | 左右 16・上下 10 | Bg.Window（#F5F7FA） | 下 1px Border.Normal | 13 Medium | Ink.Value（#212126） | 上 8 | x 20 |
| zip-info | 幅いっぱい | 12/10、行の間 4 | Bg.Subtle（#F9FAFA） | 1px Border.Soft（#E0E2E5） | 名前 Heading・ほか Meta | 名前 Ink.Strong・ほか Ink.Body | 6 | — |
| 項目名（名前・元のフォルダ） | Hug | 下 6 | — | — | 11 Medium | Ink.Body（#5F6368） | — | — |
| NameBox | 300×28 | 左右 8 | Bg.Surface | 1px Border.Input（#D1D1D1）。DI-2 は Danger.Dot（#D93025） | Cell（12） | Ink.Strong | 4 | — |
| 同じ名前の文 | Hug | — | — | — | 11 Medium | #B3261E（キー無し） | — | — |
| ラジオ | 14×14 | 文との間 8 | — | — | Cell（12） | Ink.Strong（未定） | — | — |
| RenameBox | 140×28 | 左右 8 | Bg.Surface | 1px Border.Input | Cell | Ink.Strong | 4 | — |
| 置き換えの補足 | Hug | — | — | — | 12 | #B3261E | — | — |
| FolderBox | 380×28 | 左右 8 | Bg.Surface | 1px Border.Input。DI-4 は Danger.Dot | Cell・「…」で切る | Ink.Strong | 4 | — |
| FolderBrowseButton | Hug | 16/7、左 8 | Bg.Surface | 1px Border.Dialog（#D1D6E0） | 13 Medium | Ink.Strong | 6 | — |
| status-line（見つかった） | Hug | アイコンと文の間 5 | — | — | 11 Medium | Ok.Strong（#1E8E3E） | — | check 14、線 1.17、Ok.Strong |
| status-line（重なる、DI-4） | Hug | 同 | — | — | 11 Medium | #B3261E | — | triangle-alert 14、線 1.17、色は未定 |
| FolderMissingPanel（DI-3） | 幅いっぱい | 10/8、アイコンと文の間 6 | Warn.Note（#FFF7E0） | 1px Warn.Line（#F0C36D） | Meta（11） | Warn.Strong（#6B4E00） | 6 | triangle-alert 14、線 1.17、色は未定 |
| CancelButton | Hug | 16/7 | Bg.Surface | 1px Border.Dialog | 13 Medium | Ink.Strong | 6 | — |
| ImportButton | Hug | 16/7 | Accent（#0078D4） | なし | 13 SemiBold | Ink.OnAccent（#FFFFFF） | 6 | — |
| ImportButton（無効、DI-4） | 同 | 同 | Accent・不透明度 0.4 | なし | 同 | 同 | 6 | — |

## 5. 状態ごとの見え方

| 状態 | 入る条件（`getImportDialogView` の出力） | NameBox の枠 | NameConflictPanel | FolderBox の枠 | FolderStatusPanel | FolderMissingPanel | ImportButton |
|---|---|---|---|---|---|---|---|
| 既定・P-DI-1 | NameCollision = 偽、FolderState = found | Border.Input | 出さない | Border.Input | 出す（check・`di.folder.found`・Ok.Strong） | 出さない | 有効 |
| DI-2 | NameCollision = 真 | Danger.Dot | 出す | Border.Input | 既定と同じ | 出さない | 有効 |
| DI-3 | FolderState = missing | Border.Input | 出さない | Border.Input | 出さない | 出す | 有効 |
| DI-4 | FolderState = overlap | Border.Input | 出さない | Danger.Dot | 出す（triangle-alert・`di.folder.overlap`・#B3261E） | 出さない | 無効（不透明度 0.4） |

- DI-3 でも ImportButton は有効（検索はできる。元のフォルダを設定するまで更新できない）。
- DI-2 と DI-3・DI-4 が同時に起きたときの見え方は Figma に無い。2 つの区画は別の行なので、両方を出す（それぞれの条件で決まる）。ImportButton は DI-4 の条件だけで無効になる。
- 例の値: 既定は名前「営業部」、元のフォルダ `C:\共有\営業部`。DI-4 は元のフォルダ `C:\共有\顧客\営業部`。

## 6. 文言

| ID | 全文 | 使う状態 |
|---|---|---|
| di.title | インデックスのインポート | すべて |
| di.zip.name | {zip のインデックス名}（例「営業部」） | すべて |
| di.zip.date | エクスポートした日時 {YYYY/MM/DD HH:MM}（例「エクスポートした日時 2026/09/27 10:00」） | すべて |
| di.zip.size | {ファイル数} ファイル・{大きさ} MB・{版}（例「2,530 ファイル・56.6 MB・v0.3.0」） | すべて |
| di.name.label | 名前 | すべて |
| di.name.value | {名前}（例「営業部」） | すべて |
| di.name.conflict | 同じ名前のインデックスがあります | DI-2 |
| di.rename | 別の名前にする | DI-2 |
| di.rename.value | {別の名前}（例「営業部（2）」） | DI-2 |
| di.overwrite | 置き換える | DI-2 |
| di.overwrite.note | （今の「{名前}」は削除されます） | DI-2 |
| di.folder.label | 元のフォルダ | すべて |
| di.folder.value | {元のフォルダ}（例 `C:\共有\営業部`） | すべて |
| di.browse | 参照… | すべて |
| di.folder.found | 見つかりました | 既定・DI-2 |
| di.folder.missing | 元のフォルダがこの PC に見つかりません。検索はできますが、元のフォルダを設定するまで更新できません。 | DI-3 |
| di.folder.overlap | 「{重なるインデックス名}」の元のフォルダと重なっています | DI-4 |
| di.cancel | キャンセル | すべて |
| di.import | インポート | すべて |

- 「エクスポートした日時」と日時の間は半角の空白。「・」は全角。
- 「営業部（2）」の括弧は全角。今の `newIndexName` は半角の「営業部(2)」を作る（「決めること」）。
- 「同じ名前のインデックスがあります」「「顧客」の元のフォルダと重なっています」は句点なし。DI-3 の文は句点あり（Figma のとおり）。

## 7. 操作と振る舞い

| きっかけ | 何が起きるか | 呼ぶ関数 | 次の状態 |
|---|---|---|---|
| ［インポート］（一覧） | 標準のファイルの選択で zip を選ぶ → 目録を読む → ダイアログを出す | `readIndexArchiveInfo` → `getImportSuggestedName` → `getImportDialogView` | 既定・DI-2・DI-3・DI-4 |
| zip が読めない | ダイアログを出さずにエラー | — | `confirm.md` の DI-E1 |
| zip が新しい版 | ダイアログを出さずにエラー | — | `confirm.md` の DI-E2 |
| NameBox に打つ | 同じ名前かを調べ直す | `getImportDialogView` | 既定 ⇄ DI-2 |
| FolderBox に打つ・FolderBrowseButton で選ぶ | 元のフォルダを調べ直す（選ぶのは `selectFolder`） | `getImportDialogView` | 既定・DI-3・DI-4 |
| RenameRadio・OverwriteRadio を選ぶ | 使う名前を変える | — | DI-2 のまま |
| ImportButton・Enter | 閉じて、別スレッドで読み込む | `importIndex` | 一覧の X-取り込み後 |
| 読み込みに失敗 | エラー | — | `confirm.md` の DI-E3 |
| CancelButton・Esc・TitleCloseButton | 閉じて、何もしない | — | そのまま |

- 既定のボタン: ImportButton（無効のときは Enter で何もしない）。Esc で閉じる: する（CancelButton が `IsCancel`）。Figma に定めは無い（「決めること」）。
- ボタンの並び: 左から［キャンセル］［インポート］。
- Tab の順: NameBox → RenameRadio → RenameBox → OverwriteRadio → FolderBox → FolderBrowseButton → CancelButton → ImportButton。
- 打つたびに調べ直すか、フォーカスが離れたときに調べるかは未定（「決めること」）。

## 8. リサイズ

- 幅 520 で固定。高さは中身に合わせる（346・378・422）。最小の高さは今の中身の高さ。
- 主の窓の大きさにかかわらず、同じ大きさで主の窓の中央に出す。暗幕は主の窓全体を覆う。
- 題は 1 行。DI-3 の文は折り返す。パス（FolderBox）と名前は 1 行で「…」で切る。

## 9. 判断層（`ui/index/import_dialog_view.ps1`）

### getImportDialogView（新しい）

- 入力: `[string]$name`、`[string]$folder`、`[bool]$folderExists`、`$items`（今の一覧。Name・Path を持つ行）、`$usedNames`
- 出力: `@{ NameCollision [bool]; FolderState = "found" | "missing" | "overlap" | "empty"; OverlapName [string]; CanImport [bool]; RenameSuggestion [string] }`

| name | folder | folderExists | items | NameCollision | FolderState | OverlapName | CanImport |
|---|---|---|---|---|---|---|---|
| 営業部 | `C:\共有\営業部` | `$true` | なし | `$false` | found | `""` | `$true` |
| 営業部 | `C:\共有\営業部` | `$true` | 営業部（`C:\共有\旧営業部`） | `$true` | found | `""` | `$true` |
| 営業部 | `C:\共有\営業部` | `$false` | なし | `$false` | missing | `""` | `$true` |
| 営業部 | `C:\共有\顧客\営業部` | `$true` | 顧客（`C:\共有\顧客`） | `$false` | overlap | 顧客 | `$false` |
| 営業部 | `""` | `$false` | なし | `$false` | empty | `""` | `$false` |

- `empty`（元のフォルダが空）の見え方は Figma に無い（「決めること」）。
- 重なりの判定は今の `getIndexFolderConflict` を使う。同じ名前の行は比べる相手から外す（今の `testIndexImportInput` と同じ）。ただし DI-2 で［別の名前にする］を選んだときは外さない。

### getImportSuggestedName（今のまま）

| indexName | usedNames | 出力 |
|---|---|---|
| 営業部 | なし | 営業部 |
| 営業部 | 営業部 | 営業部(2)（今の決まり。Figma は「営業部（2）」） |
| 営業部 | 営業部・営業部(2) | 営業部(3) |
| 営業部 | 営業部（大文字・小文字違いの英字名でも重なる） | 名前(2) の形 |

### getImportZipView（`getIndexImportNotice` を置き換える）

- 入力: `$info`（`readIndexArchiveInfo` の結果。Name・ExportedAt・Files・Bytes・Version）
- 出力: `@{ NameText; DateText; SizeText }`

| Files | Bytes | ExportedAt | Version | DateText | SizeText |
|---|---|---|---|---|---|
| 2530 | 59349606 | 2026-09-27 10:00 | v0.3.0 | エクスポートした日時 2026/09/27 10:00 | 2,530 ファイル・56.6 MB・v0.3.0 |
| 1 | 1024 | 2026-01-05 09:05 | v0.3.0 | エクスポートした日時 2026/01/05 09:05 | 1 ファイル・0.1 MB・v0.3.0 |

- 目録に `ExportedAt`・`Version` があるかは、今の `readIndexArchiveInfo` で確かめていない（「決めること」）。
- 0.1 MB より小さいときは 0.1 に切り上げる（今の `getIndexImportNotice` の決まり）。

## 10. 画面層（`ui/index/import_dialog.ps1`）

- `XamlReader::Load` → Owner を主の窓 → `getImportZipView` の出力を zip-info に写す → NameBox・FolderBox に初めの値を入れる。
- `NameBox.Add_TextChanged`・`FolderBox.Add_TextChanged`・`FolderBrowseButton.Add_Click`: `Test-Path` で元のフォルダがあるかを調べ（ネットワークのフォルダで遅いときは別スレッド。戻りは Dispatcher）、`getImportDialogView` → 5 章の表のとおりに見え方を写す。
- `ImportButton.Add_Click`: `CanImport` が偽なら何もしない。真なら `DialogResult = $true`。閉じた後、名前（DI-2 なら RenameBox か上書き）と元のフォルダを `importIndex` に渡し、`startJob` で別スレッドで読み込む。
- 暗幕を出す（`confirm.md` の `showScrim`）→ `ShowDialog()` → 閉じたら暗幕を消す。

## 11. 受け入れ

見比べる Figma: 既定 340:1597、DI-2 372:1216、DI-3 372:1338、DI-4 372:1407、P-DI-1 372:1475（`../png/` に画像は無い。get_screenshot で撮って並べる）。

- 幅 520、高さ 既定 346・DI-2 422・DI-3 378・DI-4 346。
- 淡い灰の箱に「営業部」と 2 行の情報。
- 名前の入力が 300×28、元のフォルダの入力が 380×28 と［参照…］。
- 既定で緑の check と「見つかりました」。
- DI-2 で名前の枠が赤、赤の文、ラジオ 2 つ（［別の名前にする］が選ばれ、140 の入力）。
- DI-3 で淡い黄の箱と注意の文。［インポート］は有効。
- DI-4 で元のフォルダの枠が赤、赤の文、［インポート］が薄い（0.4）。

## 決めること

- 赤の文字 #B3261E に theme のキー（上では `Error.Text`）を足すか。ラジオの Style（`Radio`）。
- DI-3 でもインポートできるままでよいか（Figma は有効）。
- 別の名前の括弧: Figma の全角「営業部（2）」と、今の `newIndexName` の半角「営業部(2)」。
- 元のフォルダが空のときの見え方。DI-2 と DI-4 が同時に起きたときの見え方。
- 打つたびに調べるか、フォーカスが離れたときに調べるか。
- 既定のボタンと Esc（Figma に定めが無い）。
- body の余白と区画の間（import の Figma では読み取れていない）。DI-2 のラジオの文字色。triangle-alert の色。
- zip の目録に、エクスポートした日時と版があるか。
- P-DI-1 の中身（数値を読んでいない）。
