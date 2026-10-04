# インデックスのエクスポート（export）

## 1. 役割

1 つのインデックスを zip に書き出すダイアログ（DX）。保存先のフォルダを選び、作る zip の名前を示す。zip には元のファイルの本文と元のフォルダの場所が入ることを注意する。

- Figma の部品: `xaml/dialog/export` 372:1008。variant は既定 340:1567 だけ（106・148 の DX で使う）。
- 開く所: インデックス管理の一覧の行の［⋯］の「エクスポート」（`../index/index_list.md`）。
- `../png/` にこのダイアログの画像は無い。見比べは Figma の get_screenshot で行う（11 章）。

## 2. 置き場所

- XAML: `scripts/tebunko/xaml/dialog_export.xaml`（新しいファイル）。
- 画面層 `ui/index/export_dialog.ps1`、判断層 `ui/index/export_dialog_view.ps1`（どちらも新しい。今はダイアログを出さず、保存の標準のダイアログだけを使っている）。
- 書き出しは状態層の `exportIndex`（`index/index_archive.ps1`）。
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

      <!-- title-bar（左右 16） -->
      <Border Grid.Row="0" Background="{DynamicResource Bg.Window}" CornerRadius="8,8,0,0"
              BorderBrush="{DynamicResource Border.Normal}" BorderThickness="0,0,0,1" Padding="16,10,16,10">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*" />
            <ColumnDefinition Width="Auto" />
          </Grid.ColumnDefinitions>
          <TextBlock FontSize="13" FontWeight="Medium" Foreground="{DynamicResource Ink.Value}" VerticalAlignment="Center"
                     Text="{dx.title}" />
          <Button x:Name="TitleCloseButton" Grid.Column="1" Width="20" Height="20" Style="{StaticResource TitleBarButton}">
            <Image Width="20" Height="20" Source="{StaticResource Icon.X}" />
          </Button>
        </Grid>
      </Border>

      <!-- body -->
      <StackPanel Grid.Row="1" Margin="20,16,20,16" Orientation="Vertical">
        <!-- info box -->
        <Border Background="{DynamicResource Bg.Subtle}" BorderBrush="{DynamicResource Border.Soft}" BorderThickness="1"
                CornerRadius="6" Padding="12,10">
          <StackPanel Orientation="Vertical">
            <TextBlock x:Name="ExportNameText" Style="{StaticResource Heading}" Foreground="{DynamicResource Ink.Strong}"
                       TextTrimming="CharacterEllipsis" Text="{dx.name}" />
            <TextBlock x:Name="ExportSizeText" Margin="0,4,0,0" Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Body}"
                       Text="{dx.size}" />
          </StackPanel>
        </Border>

        <!-- 保存先 -->
        <StackPanel Margin="0,14,0,0" Orientation="Vertical">
          <TextBlock FontSize="11" FontWeight="Medium" Foreground="{DynamicResource Ink.Body}" Text="{dx.dest.label}" />
          <StackPanel Margin="0,6,0,0" Orientation="Horizontal">
            <TextBox x:Name="ExportPathBox" Width="380" Height="28" Style="{StaticResource TextBox.Base}" Padding="8,0"
                     VerticalContentAlignment="Center" BorderBrush="{DynamicResource Border.Input}" Text="{dx.dest.path}" />
            <Button x:Name="ExportBrowseButton" Margin="8,0,0,0" Style="{StaticResource DialogButton}" Padding="16,7"
                    Content="{dx.browse}" />
          </StackPanel>
          <TextBlock x:Name="ExportFileNameText" Margin="0,6,0,0" Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Body}"
                     TextTrimming="CharacterEllipsis" Text="{dx.dest.file}" />
        </StackPanel>

        <!-- note-warn -->
        <Border Margin="0,14,0,0" Background="{DynamicResource Warn.Note}" BorderBrush="{DynamicResource Warn.Line}"
                BorderThickness="1" CornerRadius="6" Padding="10,8">
          <TextBlock Style="{StaticResource Meta}" Foreground="{DynamicResource Warn.Strong}" TextWrapping="Wrap" Text="{dx.warn}" />
        </Border>

        <!-- buttons -->
        <StackPanel Margin="0,14,0,0" Orientation="Horizontal" HorizontalAlignment="Right">
          <Button x:Name="CancelButton" Style="{StaticResource DialogButton}" Padding="16,7" IsCancel="True" Content="{dx.cancel}" />
          <Button x:Name="ExportButton" Margin="8,0,0,0" Style="{StaticResource Primary}" Padding="16,7" IsDefault="True" Content="{dx.export}" />
        </StackPanel>
      </StackPanel>
    </Grid>
  </Border>
</Window>
```

- body の余白（上下 16・左右 20）と区画の間（14）は、ほかのダイアログ（indexing_confirm・import）と同じ値を入れた。export の Figma で数値を読んだのは、題のバーの左右 16 と、中の部品の値（下の 4 章）だけ（「決めること」）。
- 「参照…」は Figma で x:Name が付いていない。`ExportBrowseButton` は新しく付けた名前。

## 4. 寸法と色

| 要素 | 大きさ | 余白 | 地 | 枠 | 文字の Style | 文字の色 | 角丸 |
|---|---|---|---|---|---|---|---|
| 外枠 | 幅 520・高さ Hug | — | Bg.Surface（#FFFFFF） | 1px Border.Normal（#D9DEE3） | — | — | 8 |
| title-bar | Hug | 左右 16・上下 10 | Bg.Window（#F5F7FA） | 下 1px Border.Normal | 13 Medium | Ink.Value（#212126） | 上 8 |
| info box | 幅いっぱい | 12/10、行の間 4 | Bg.Subtle（#F9FAFA） | 1px Border.Soft（#E0E2E5） | — | — | 6 |
| ExportNameText | Hug | — | — | — | Heading（13 SemiBold） | Ink.Strong（#202124） | — |
| ExportSizeText | Hug | — | — | — | Meta（11） | Ink.Body（#5F6368） | — |
| 「保存先」 | Hug | 下 6 | — | — | 11 Medium | Ink.Body | — |
| ExportPathBox | 380×28 | 左右 8 | Bg.Surface | 1px Border.Input（#D1D1D1） | Cell（12）・「…」で切る | Ink.Strong（未定） | 4 |
| ExportBrowseButton | Hug | 16/7、左 8 | Bg.Surface | 1px Border.Dialog（#D1D6E0） | 13 Medium | Ink.Strong | 6 |
| ExportFileNameText | 幅いっぱい | 上 6 | — | — | Meta | Ink.Body | — |
| note-warn | 幅いっぱい | 10/8 | Warn.Note（#FFF7E0） | 1px Warn.Line（#F0C36D） | Meta | Warn.Strong（#6B4E00） | 6 |
| CancelButton | Hug | 16/7 | Bg.Surface | 1px Border.Dialog | 13 Medium | Ink.Strong | 6 |
| ExportButton | Hug | 16/7 | Accent（#0078D4） | なし | 13 SemiBold | Ink.OnAccent（#FFFFFF） | 6 |

## 5. 状態ごとの見え方

| 状態 | 入る条件 | ExportPathBox | ExportFileNameText | ExportButton |
|---|---|---|---|---|
| 既定（DX） | ダイアログを開いた | 初めの保存先（例 `C:\Users\test\Documents`） | `dx.dest.file` | 有効 |
| 保存先が空・使えない | 未定 | 未定 | 未定 | 未定 |

- Figma の variant は既定だけ。保存先が空のとき・書き込めないとき・同じ名前の zip があるときの見え方は無い（「決めること」）。
- 書き出し中・書き出した後は、インデックス管理の一覧の X-エクスポート中・X-取り込み後（`../index/index_list.md`）で示す。ダイアログは［エクスポート］で閉じる。

## 6. 文言

| ID | 全文 | 使う状態 |
|---|---|---|
| dx.title | インデックスのエクスポート | 既定 |
| dx.name | {インデックス名}（例「営業部」） | 既定 |
| dx.size | {ファイル数} ファイル・{大きさ} MB（例「2,530 ファイル・94.0 MB」） | 既定 |
| dx.dest.label | 保存先 | 既定 |
| dx.dest.path | {保存先のフォルダ}（例 `C:\Users\test\Documents`） | 既定 |
| dx.browse | 参照… | 既定 |
| dx.dest.file | {zip の名前} として保存します（例「営業部_インデックス_20260930.zip として保存します」） | 既定 |
| dx.warn | このファイルには、元のファイルの本文と元のフォルダの場所が含まれます。元のファイルと同じように取り扱ってください。 | 既定 |
| dx.cancel | キャンセル | 既定 |
| dx.export | エクスポート | 既定 |

- 「2,530 ファイル・94.0 MB」: 数と単位の間は半角の空白、区切りは全角の「・」、大きさは小数 1 桁。
- zip の名前の決まりは Figma の例「{インデックス名}_インデックス_{YYYYMMDD}.zip」から読める形。日付が書き出した日か、ほかの日かは未定（「決めること」）。

## 7. 操作と振る舞い

| きっかけ | 何が起きるか | 呼ぶ関数 | 次の状態 |
|---|---|---|---|
| 行の［⋯］の「エクスポート」 | ダイアログを出す | `getExportDialogView` | 既定 |
| ExportBrowseButton をクリック | Windows の標準のフォルダの選択（`selectFolder`）を開き、選んだら ExportPathBox に入れる | `selectFolder` | 既定 |
| ExportButton・Enter | 閉じて、別スレッドで書き出す | `exportIndex` | 一覧の X-エクスポート中 → X-取り込み後 |
| CancelButton・Esc・TitleCloseButton | 閉じて、何もしない | — | そのまま |
| 書き出しに失敗 | エラーを出す | — | `confirm.md` の DC-err |

- 既定のボタン: ExportButton。Esc で閉じる: する（CancelButton が `IsCancel`）。Figma に定めは無い（「決めること」）。
- ボタンの並び: 左から［キャンセル］［エクスポート］。
- Tab の順: ExportPathBox → ExportBrowseButton → CancelButton → ExportButton。
- ExportPathBox に直接打てるか（読み取り専用にするか）は未定（「決めること」）。

## 8. リサイズ

- 幅 520 で固定。高さは中身に合わせる。最小の高さは今の中身の高さ。
- 主の窓の大きさにかかわらず、同じ大きさで主の窓の中央に出す。暗幕は主の窓全体を覆う。
- 題は 1 行。注意（note-warn）は折り返す。パス（ExportPathBox）と zip の名前は 1 行で「…」で切る。

## 9. 判断層（`ui/index/export_dialog_view.ps1`）

### getExportDialogView

- 入力: `[string]$name`（インデックス名）、`[int]$files`（ファイル数）、`[long]$bytes`（インデックスの大きさ）、`[datetime]$now`
- 出力: `@{ NameText; SizeText; FileName; FileNameText }`

| name | files | bytes | now | SizeText | FileName | FileNameText |
|---|---|---|---|---|---|---|
| 営業部 | 2530 | 98566144 | 2026-09-30 | 2,530 ファイル・94.0 MB | 営業部_インデックス_20260930.zip | 営業部_インデックス_20260930.zip として保存します |
| 顧客 | 1 | 1024 | 2026-01-05 | 1 ファイル・未定 MB | 顧客_インデックス_20260105.zip | 顧客_インデックス_20260105.zip として保存します |

- 大きさが 0.1 MB より小さいときの書き方は未定（import の `getIndexImportNotice` は 0.1 に切り上げている）。
- インデックス名にファイル名に使えない文字は来ない（`testIndexName` が断る）。

### getExportDefaultFolder

- 初めの保存先。Figma の例は利用者のドキュメント（`C:\Users\test\Documents`）。前回の保存先を覚えるかは未定（「決めること」）。テストの表は決まってから書く。

## 10. 画面層（`ui/index/export_dialog.ps1`）

- `XamlReader::Load` → Owner を主の窓 → `getExportDialogView` の出力を写す。
- `ExportBrowseButton.Add_Click`: `selectFolder` の結果を `ExportPathBox.Text` に入れる。
- `ExportButton.Add_Click`: `DialogResult = $true`。閉じた後、`Join-Path ExportPathBox.Text FileName` を `exportIndex` に渡し、`startJob` で別スレッドで書き出す。終わりは Dispatcher で一覧に戻す。
- 暗幕を出す（`confirm.md` の `showScrim`）→ `ShowDialog()` → 閉じたら暗幕を消す。

## 11. 受け入れ

見比べる Figma: 既定 340:1567（`../png/` に画像は無い。get_screenshot で撮って並べる）。

- 幅 520、題「インデックスのエクスポート」。
- 淡い灰の箱に「営業部」（13 SemiBold）と「2,530 ファイル・94.0 MB」。
- 「保存先」の下に 380×28 の入力と［参照…］、その下に「… として保存します」。
- 淡い黄の箱に注意の文（濃い黄の文字）。
- 右下に［キャンセル］［エクスポート］。

## 決めること

- 判断層・画面層を新しく分ける（今はダイアログが無い）。ファイルの名前。
- zip の名前の決まり（日付は書き出した日か）。同じ名前の zip があるときの扱い。
- ExportPathBox に直接打てるか。初めの保存先（ドキュメントか、前回の場所か）。
- 保存先が空・書き込めないときの見え方。
- body の余白と区画の間（export の Figma では読み取れていない）。
- 既定のボタンと Esc（Figma に定めが無い）。
- 0.1 MB より小さいときの大きさの書き方。
- ExportPathBox の文字色。
