# tebunko について（about）

## 1. 役割

tebunko の版・ビルド・動作環境と、ライセンスを示すダイアログ。リリースノートとライセンスへのリンクを置く。

- Figma の部品: `xaml/dialog/about` 372:54。variant は既定 340:1489（106）と P-DV 372:538（148）。2 つは同じ見た目。
- 開く所: 左の欄のナビの「バージョン情報」（`AboutLink`、`../shell/nav.md`）。
- 次の項目はメンテナの決定で、置かない。
  - 著作権の行
  - 対応形式の一覧
- `../png/` にこのダイアログの画像は無い。見比べは Figma の get_screenshot で行う（11 章）。

## 2. 置き場所

- XAML: `scripts/tebunko/xaml/dialog_about.xaml`（今のファイルを書き直す）。
- 画面層 `ui/about_dialog.ps1`、判断層 `ui/about_view.ps1`（今のファイル）。
- 型: Window。`ShowDialog()` で出す。Owner は主の窓。
- 置き方: 主の窓の中央（`WindowStartupLocation="CenterOwner"`）。
- 暗幕: 出している間、主の窓全体を黒の 30%（rgba(0,0,0,0.3)）で覆う。覆い方は `confirm.md` の 2 章と同じ。
- 窓の枠: `WindowStyle="None"`・`AllowsTransparency="True"`・`ResizeMode="NoResize"`。題のバーは自分で描く（下の木）。

## 3. 部品の木

```xml
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="460" SizeToContent="Height" WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        ResizeMode="NoResize" WindowStartupLocation="CenterOwner" ShowInTaskbar="False">
  <Border Background="{DynamicResource Bg.Surface}" BorderBrush="{DynamicResource Border.Dialog}" BorderThickness="1"
          CornerRadius="8" Margin="24">
    <Border.Effect>
      <DropShadowEffect BlurRadius="24" ShadowDepth="8" Direction="270" Opacity="0.15" Color="#000000" />
    </Border.Effect>
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="32" />    <!-- title-bar -->
        <RowDefinition Height="Auto" />  <!-- 中身 -->
      </Grid.RowDefinitions>

      <!-- title-bar -->
      <Border x:Name="TitleBar" Grid.Row="0" Background="{DynamicResource Bg.Hover}" CornerRadius="8,8,0,0"
              BorderBrush="{DynamicResource Border.Dialog}" BorderThickness="0,0,0,1" Padding="16,0,0,0">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="*" />
            <ColumnDefinition Width="46" />
          </Grid.ColumnDefinitions>
          <Border Grid.Column="0" Width="16" Height="16" CornerRadius="3" Background="{DynamicResource Accent}" VerticalAlignment="Center">
            <Image Width="10" Height="10" Source="{StaticResource Icon.FileSearch.OnAccent}" />
          </Border>
          <TextBlock Grid.Column="1" Margin="8,0,0,0" Style="{StaticResource Label}" Foreground="{DynamicResource Ink.Strong}"
                     VerticalAlignment="Center" Text="{about.windowTitle}" />
          <Button x:Name="TitleCloseButton" Grid.Column="2" Width="46" Height="32" Style="{StaticResource TitleBarButton}">
            <Image Width="10" Height="10" Source="{StaticResource Icon.CircleX}" />
          </Button>
        </Grid>
      </Border>

      <!-- 中身（padding 24/20、間 16、中央揃え） -->
      <StackPanel Grid.Row="1" Margin="20,24,20,24" Orientation="Vertical">
        <!-- app-info-block -->
        <StackPanel HorizontalAlignment="Center" Orientation="Vertical">
          <Border x:Name="AppIcon" Width="48" Height="48" CornerRadius="12" Background="{DynamicResource Select.Soft}" HorizontalAlignment="Center">
            <Image Width="28" Height="28" Source="{StaticResource Icon.FileSearch.Accent}" />
          </Border>
          <TextBlock Margin="0,16,0,0" Style="{StaticResource AppTitle}" Foreground="{DynamicResource Ink.Strong}"
                     HorizontalAlignment="Center" Text="{about.appName}" />
          <TextBlock Margin="0,6,0,0" Style="{StaticResource Nav.Tall}" Foreground="{DynamicResource Ink.Body}"
                     HorizontalAlignment="Center" Text="{about.tagline}" />
          <StackPanel Margin="0,6,0,0" Orientation="Horizontal" HorizontalAlignment="Center">
            <TextBlock x:Name="VersionText" Style="{StaticResource Meta.Tall}" Foreground="{DynamicResource Ink.Body}" Text="{about.version}" />
            <TextBlock x:Name="VersionSeparator" Margin="8,0,0,0" Style="{StaticResource Meta.Tall}" Foreground="{DynamicResource Ink.Placeholder}" Text="{about.sep}" />
            <TextBlock x:Name="CommitText" Margin="8,0,0,0" Style="{StaticResource Meta.Tall}" Foreground="{DynamicResource Ink.Placeholder}" Text="{about.build}" />
          </StackPanel>
        </StackPanel>

        <Border Margin="0,16,0,0" Height="1" Background="{DynamicResource Border.Dialog}" /> <!-- 線の色は未定 -->

        <!-- system-info-panel -->
        <Border Margin="0,16,0,0" Background="{DynamicResource Bg.Subtle}" BorderBrush="{DynamicResource Border.Dialog}"
                BorderThickness="1" CornerRadius="6" Padding="16,8">
          <StackPanel Orientation="Vertical">
            <StackPanel Orientation="Horizontal">
              <TextBlock Width="66" Style="{StaticResource Meta.Key}" Foreground="{DynamicResource Ink.Body}" Text="{about.env.label}" />
              <TextBlock Style="{StaticResource Meta.Tall}" Foreground="{DynamicResource Ink.Strong}" Text="{about.env.value}" />
            </StackPanel>
            <StackPanel Margin="0,6,0,0" Orientation="Horizontal">
              <TextBlock Width="66" Style="{StaticResource Meta.Key}" Foreground="{DynamicResource Ink.Body}" Text="{about.runtime.label}" />
              <TextBlock Style="{StaticResource Meta.Tall}" Foreground="{DynamicResource Ink.Strong}" Text="{about.runtime.value}" />
            </StackPanel>
          </StackPanel>
        </Border>

        <!-- footer（両端に寄せる） -->
        <Grid Margin="0,16,0,0">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*" />
            <ColumnDefinition Width="Auto" />
          </Grid.ColumnDefinitions>
          <StackPanel Grid.Column="0" Orientation="Vertical" VerticalAlignment="Center">
            <TextBlock Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Placeholder}" Text="{about.license}" />
            <StackPanel Margin="0,4,0,0" Orientation="Horizontal">
              <TextBlock Style="{StaticResource Link}">
                <Hyperlink x:Name="ReleaseNotesLink" Foreground="{DynamicResource Accent}" TextDecorations="{x:Null}">
                  <Run Text="{about.link.release}" />
                </Hyperlink>
              </TextBlock>
              <Image Margin="4,0,0,0" Width="7" Height="7" Source="{StaticResource Icon.ExternalLink.Accent}" VerticalAlignment="Center" />
              <TextBlock Margin="12,0,0,0" Style="{StaticResource Meta}" Foreground="{DynamicResource Ink.Faint}" Text="{about.sep}" />
              <TextBlock Margin="12,0,0,0" Style="{StaticResource Link}">
                <Hyperlink x:Name="LicenseLink" Foreground="{DynamicResource Accent}" TextDecorations="{x:Null}">
                  <Run Text="{about.link.license}" />
                </Hyperlink>
              </TextBlock>
              <Image Margin="4,0,0,0" Width="7" Height="7" Source="{StaticResource Icon.ExternalLink.Accent}" VerticalAlignment="Center" />
            </StackPanel>
          </StackPanel>
          <Button x:Name="CloseButton" Grid.Column="1" Style="{StaticResource Primary}" Padding="16,7"
                  VerticalAlignment="Bottom" IsDefault="True" IsCancel="True" Content="{about.ok}" />
        </Grid>
      </StackPanel>
    </Grid>
  </Border>
</Window>
```

- 外側の `Margin="24"` は影を描くための透明の余白。見た目の幅は 460。
- タイトルバーは上に置く（決定。Figma の部品も上に直した）。
- footer の中の行の組み方（「MIT ライセンスで公開しています」とリンクの行が縦か横か、その間）は Figma の数値を読み取れていない。上の縦並び・間 4 は置き場所を示すだけ（「決めること」）。
- アイコン（`Icon.*`）は theme.xaml の DrawingImage として足す。

## 4. 寸法と色

| 要素 | 大きさ | 余白 | 地 | 枠 | 文字の Style | 文字の色 | 角丸 | アイコン |
|---|---|---|---|---|---|---|---|---|
| 外枠 | 幅 460・高さ Hug | — | Bg.Surface（#FFFFFF） | 1px Border.Dialog（#D1D6E0） | — | — | 8 | — |
| 影 | 0 8 24 rgba(0,0,0,.15) | — | — | — | — | — | — | — |
| TitleBar | 高さ 32 | 左 16 | Bg.Hover（#F3F3F4） | 下 1px Border.Dialog | — | — | 上 8 | — |
| app-logo | 16×16 | — | 無し | — | — | — | 0 | アプリのロゴ（`illust/app_logo.svg`）16×16（README の 10） |
| 窓の題 | Hug | 左 8（未定） | — | — | Label（12 SemiBold） | Ink.Strong（#202124） | — | — |
| TitleCloseButton | 46×32 | — | 透明 | — | — | — | — | circle-x 10、線 0.83、色は未定 |
| 中身 | — | 上下 24・左右 20、間 16 | — | — | — | — | — | — |
| AppIcon | 48×48 | — | 無し | — | — | — | 0 | アプリのロゴ（`illust/app_logo.svg`）48×48（README の 10） |
| アプリ名 | Hug | 上 16 | — | — | AppTitle（20 SemiBold/28） | Ink.Strong（#202124） | — | — |
| 説明 | Hug | 上 6 | — | — | Nav.Tall（13 Medium/18） | Ink.Body（#5F6368） | — | — |
| VersionText | Hug | — | — | — | Meta.Tall（11/16） | Ink.Body（#5F6368） | — | — |
| 区切り・CommitText | Hug | 間 8 | — | — | Meta.Tall | Ink.Placeholder（#9AA0A6） | — | — |
| 線 | 高さ 1 | — | 未定 | — | — | — | — | — |
| system-info-panel | 幅いっぱい | 16/8、行の間 6 | Bg.Subtle（#F9FAFA） | 1px Border.Dialog | — | — | 6 | — |
| 項目名 | 幅 66 | — | — | — | Meta.Key（11 SemiBold/16） | Ink.Body | — | — |
| 項目の値 | Hug | — | — | — | Meta.Tall | Ink.Strong | — | — |
| ライセンスの文 | Hug | — | — | — | Meta（11） | Ink.Placeholder（#9AA0A6） | — | — |
| リンク | Hug | 間 12 | — | — | Link（11 SemiBold） | Accent（#0078D4） | — | external-link 7、線 0.58、Accent |
| リンクの区切り「|」 | Hug | — | — | — | Meta | Ink.Faint（#99A1AB） | — | — |
| CloseButton | Hug | 16/7 | Accent | なし | 13 SemiBold | Ink.OnAccent（#FFFFFF） | 6 | — |

## 5. 状態ごとの見え方

| 状態 | 入る条件 | VersionText | CommitText と区切り |
|---|---|---|---|
| 既定・P-DV（配布版） | `readVersionFile` が値を返した | `about.version`（例「バージョン 1.2.0」） | 出す `about.build`（例「ビルド 2024.10.15」） |
| 開発版 | `readVersionFile` が `$null`（git から直接起動） | `about.version.dev` | 出さない（今のコードの決まり。Figma に variant は無い） |

- Figma の「ビルド 2024.10.15」は日付。今の `getAboutView` はコミットの先頭 7 文字を返す。どちらを出すかは未定（「決めること」）。

## 6. 文言

| ID | 全文 | 使う状態 |
|---|---|---|
| about.windowTitle | tebunko について | すべて |
| about.appName | tebunko | すべて |
| about.tagline | Office ファイル内テキスト検索ツール | すべて |
| about.version | バージョン {版} | 配布版 |
| about.version.dev | 開発版 | 開発版（今のコードの文言。Figma に無い） |
| about.sep | \| | すべて |
| about.build | ビルド {ビルド} | 配布版 |
| about.env.label | 動作環境： | すべて |
| about.env.value | Windows 10 / 11 | すべて |
| about.runtime.label | ランタイム： | すべて |
| about.runtime.value | PowerShell 5.1 + WPF | すべて |
| about.license | MIT ライセンスで公開しています | すべて |
| about.link.release | リリースノート | すべて |
| about.link.license | ライセンス（MIT） | すべて |
| about.ok | OK | すべて |

- 「動作環境：」「ランタイム：」のコロンは全角。「ライセンス（MIT）」の括弧は全角。
- 例の値「1.2.0」「2024.10.15」は Figma の見本で、実際は `getAboutView` の値を入れる。

## 7. 操作と振る舞い

| きっかけ | 何が起きるか | 呼ぶ関数 | 次の状態 |
|---|---|---|---|
| ナビの「バージョン情報」をクリック | ダイアログを出す | `readVersionFile` → `getAboutView` | 既定 |
| ReleaseNotesLink をクリック | 既定のブラウザでリリースノートを開く | URL は未定 | そのまま |
| LicenseLink をクリック | 既定のブラウザでライセンスを開く | URL は未定 | そのまま |
| CloseButton・TitleCloseButton をクリック | 閉じる | — | 閉じる |
| Esc | 閉じる（CloseButton が `IsCancel`） | — | 閉じる |
| Enter | 閉じる（CloseButton が `IsDefault`） | — | 閉じる |

- 既定のボタン: CloseButton。Esc で閉じる: する。Figma に定めは無く、ボタンが 1 つだけのため（「決めること」で確かめる）。
- Tab の順: ReleaseNotesLink → LicenseLink → CloseButton。
- リンクのホバーの見た目は Figma に無い（未定）。
- ツールチップ: 無し。

## 8. リサイズ

- 幅 460 で固定。高さは中身に合わせる（`SizeToContent="Height"`）。最小の高さは今の中身の高さ。
- 主の窓の大きさ（1024×640・1280×820・1600×1000）にかかわらず、同じ大きさで主の窓の中央に出す。
- 暗幕は主の窓全体を覆う。
- 文は折り返さない（どれも 1 行に収まる）。版が長くて収まらないときの扱いは未定。

## 9. 判断層（`ui/about_view.ps1`）

### getAboutView

- 入力: `$versionInfo`（`readVersionFile` の結果。`Tag`・`Sha`、または `$null`）
- 出力: `@{ VersionText [string]; BuildText [string]; ShowBuild [bool] }`
  - 今の出力（`Version`・`Commit`）から、画面に出す全文を返す形に変える。
  - BuildText にコミットを出すか日付を出すかは「決めること」。下の表は今の決まり（コミットの先頭 7 文字）で書く。

| versionInfo | VersionText | BuildText | ShowBuild |
|---|---|---|---|
| `$null` | 開発版 | `""` | `$false` |
| `@{ Tag = "v0.3.0"; Sha = "0123456789abcdef" }` | バージョン v0.3.0 | ビルド 0123456 | `$true` |
| `@{ Tag = "1.2.0"; Sha = "abcdef0123456789" }` | バージョン 1.2.0 | ビルド abcdef0 | `$true` |

- Figma は「バージョン 1.2.0」で、先頭の `v` が無い。タグの `v` を外すかは「決めること」。

## 10. 画面層（`ui/about_dialog.ps1`）

- `AboutLink` のクリックで、`XamlReader::Load` でダイアログを読み、Owner を主の窓にする。
- `getAboutView (readVersionFile)` の出力を `VersionText`・`CommitText` に写し、`ShowBuild` が偽なら `CommitText` と `VersionSeparator` を隠す。
- 暗幕を出す（`confirm.md` の 10 章の `showScrim`）→ `ShowDialog()` → 閉じたら暗幕を消す。
- ハイパーリンクは `Add_RequestNavigate`（または `Add_Click`）で `Start-Process <URL>`。
- 重い処理は無い（すべて画面のスレッドでよい）。

## 11. 受け入れ

見比べる Figma: 既定 340:1489、P-DV 372:538（`../png/` に画像は無い。get_screenshot で撮って並べる）。

- 幅 460、角丸 8、枠 Border.Dialog。
- 48 のアプリのロゴ（青・緑・橙の 3 段）。地のタイルは無い。
- 「tebunko」が 20 SemiBold で中央、その下に説明、その下に版の行。
- 動作環境の箱が淡い灰の地で、項目名が幅 66 にそろう。
- 左下に「MIT ライセンスで公開しています」とリンク 2 つ、右下に青の［OK］。
- 著作権の行と対応形式の一覧が無い。

## 決めること

- 版の表示: Figma の「バージョン 1.2.0」「ビルド 2024.10.15」（日付）と、今の `getAboutView`（タグとコミットの先頭 7 文字）のどちらに合わせるか。タグの `v` を外すか。
- リリースノートとライセンスのリンクの URL。
- 影が他のダイアログ（0 4 20 .18）と違う（0 8 24 .15）。そろえるか。
- 既定のボタンと Esc（Figma に定めが無い。上は CloseButton を既定・Esc で閉じるとした）。
- 「動作環境」「ランタイム」の行は新しい。値を固定の文にするか、実行中の環境から取るか。
- footer の中の組み方と間（Figma の数値を読み取れていない）。中身の上の線の色。
- 窓の題とロゴの間、閉じるアイコンの色、リンクのホバーの見た目。
