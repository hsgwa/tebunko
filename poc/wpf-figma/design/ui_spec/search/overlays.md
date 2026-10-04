# 重ねて出すもの（バナー・吹き出し・ツールチップ・ファイル内の対象のメニュー）

## 1. 役割

検索の画面で、ほかの部品の上や間に出す小さなものの見た目と出し方を決める。

| もの | 何を知らせるか | Figma | 中身（文言・条件）を決める文書 |
|---|---|---|---|
| バナー | インデックスの更新中・前回の更新が途中・保存した・検索を中止した | 106:54（info）・106:60（warn）・106:72（success）・106:66（danger） | `search.md`（どれを出すか）・この文書（見た目と `setBanner`） |
| 高速検索の吹き出し | 高速検索を使えない・一部だけ使えるわけ | 135:1801（正規表現の吹き出し）・136:89（高速検索の表示） | `search_bar.md` の `getFastSearchView` |
| ツールチップ（暗い地） | 正規表現のチェックの説明・検索対象のツリーのフルパス | 201:5488・176:2571 | `search_bar.md`・`target_tree.md` |
| ファイル内の対象のメニュー | 本文のほかに、図形・コメント・ノートも探すか | 201:5488（見本）・213:6592（H-範囲）・213:6958（H-範囲2） | 文言とボタンの文言は `search_bar.md` |

開くメニュー（プレビューの［開く ▾］）は `preview.md` で決める。

## 2. 置き場所

| もの | XAML | 差す・付ける先 |
|---|---|---|
| バナー | `scripts/tebunko/xaml/search/banner.xaml`（根は `Border`、x:Name `BannerRoot`） | `search.md` の `TopBannerHost`・`MidBannerHost`（ContentControl、高さ 40）。2 つの口に別々に読んで差す |
| 高速検索の吹き出し | `search_bar.xaml` の中に Popup で置く（x:Name `FastSearchTipPopup`） | PlacementTarget は `FastSearchText`（search_bar.md のバッジ） |
| ツールチップ（暗い地） | theme に Style `Tooltip.Dark`（ToolTip 用）を置く | 正規表現: `RegexCheck`。フルパス: ツリーの行の Grid（target_tree.md） |
| ファイル内の対象のメニュー | `search_bar.xaml` の中に Popup で置く（x:Name `ScopeMenuPopup`） | PlacementTarget は `ScopeButton` |

- 画面層: `scripts/tebunko/ui/search/overlays.ps1`（`setBanner`・`openScopeMenu`・`showFastSearchTip`）。
- 判断層: `scripts/tebunko/ui/search/overlays_view.ps1`（`getBannerView`・`getScopeMenuItems`・`splitTipLink`）。
- テスト: `tests/tebunko/ui/search/overlays_view.Tests.ps1`。
- 読み込み口: `gui.ps1`（画面層）・`tebunko/lib.ps1`（判断層）。

## 3. 部品の木

文言は `{文言ID}` で書く。

### バナー（banner.xaml）

```xml
<Border x:Name="BannerRoot" Height="40" Background="{StaticResource Banner.Info.Bg}"
        BorderBrush="{StaticResource Accent}" BorderThickness="3,0,0,0" CornerRadius="0"
        Padding="16,8,12,8">
  <Grid>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="Auto"/>  <!-- アイコン 16 -->
      <ColumnDefinition Width="10"/>    <!-- 間 10 -->
      <ColumnDefinition Width="*"/>     <!-- 文 -->
      <ColumnDefinition Width="Auto"/>  <!-- ボタン（無いことがある） -->
    </Grid.ColumnDefinitions>
    <Viewbox Grid.Column="0" Width="16" Height="16" VerticalAlignment="Center">
      <Path x:Name="BannerIcon" Data="{StaticResource Icon.Info}" Stroke="{StaticResource Accent}"
            StrokeThickness="2" StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
    </Viewbox>
    <TextBlock x:Name="BannerText" Grid.Column="2" VerticalAlignment="Center"
               Style="{StaticResource Label.Medium}" Foreground="{StaticResource Banner.Info.Ink}"
               TextTrimming="CharacterEllipsis" TextWrapping="NoWrap" Text="{banner.updating}"/>
    <Button x:Name="BannerAction" Grid.Column="3" Margin="10,0,0,0" VerticalAlignment="Center"
            Style="{StaticResource BannerButton}" Background="{StaticResource Bg.Surface}"
            BorderBrush="{StaticResource Accent}" BorderThickness="1" Padding="10,4,10,4">
      <TextBlock x:Name="BannerActionText" Style="{StaticResource Meta.SemiBold}"
                 Foreground="{StaticResource Accent}" Text="{banner.updating.action}"/>
    </Button>
  </Grid>
</Border>
```

- 色（地・帯・アイコン・文・ボタンの枠と文字）は種類ごとに `setBanner` が入れ替える（4 章の表）。
- `Banner.*`・`Label.Medium`・`Meta.SemiBold`・`BannerButton` は theme に無い。足す案を 4 章と「決めること」に書く。

### 高速検索の吹き出し（search_bar.xaml の中）

```xml
<Popup x:Name="FastSearchTipPopup" PlacementTarget="{Binding ElementName=FastSearchText}"
       Placement="Bottom" StaysOpen="True" AllowsTransparency="True">
  <Grid Margin="8,0,8,8">  <!-- 影のための余白 -->
    <Grid.RowDefinitions>
      <RowDefinition Height="6"/>     <!-- 上向きの矢印 -->
      <RowDefinition Height="Auto"/>  <!-- 本体 -->
    </Grid.RowDefinitions>
    <Path x:Name="FastSearchTipArrow" Grid.Row="0" Width="10" Height="6" HorizontalAlignment="Center"
          Data="M0,6 L5,0 L10,6 Z" Fill="{StaticResource Tip.Bg}"/>
    <Border Grid.Row="1" Background="{StaticResource Tip.Bg}" CornerRadius="5" Padding="10,7,10,7">
      <Border.Effect><DropShadowEffect BlurRadius="8" ShadowDepth="2" Direction="270" Opacity="0.15" Color="#000000"/></Border.Effect>
      <TextBlock x:Name="FastSearchTipText" MaxWidth="280" TextWrapping="Wrap"
                 Style="{StaticResource Meta}" Foreground="{StaticResource Tip.Ink}">
        <Run x:Name="FastSearchTipBody" Text="{search.fast.nows.tip の［ ］より前}"/>
        <Hyperlink x:Name="FastSearchTipLink" Foreground="{StaticResource Tip.Ink}">
          <Run Text="［インデックス管理で確認］"/>
        </Hyperlink>
      </TextBlock>
    </Border>
  </Grid>
</Popup>
```

- 文と［インデックス管理で確認］の分け方は `splitTipLink`（9 章）で決める。リンクが無い文言のときは `FastSearchTipLink` を消す（`Inlines` から外す）。
- 押せるリンクを入れるため ToolTip ではなく Popup にする（WPF の ToolTip の中は押せない）。ToolTip にするかは「決めること」。

### ツールチップ（暗い地）の Style

```xml
<Style x:Key="Tooltip.Dark" TargetType="ToolTip">
  <Setter Property="Background" Value="{StaticResource Tip.Bg}"/>
  <Setter Property="Foreground" Value="{StaticResource Tip.Ink}"/>
  <Setter Property="BorderThickness" Value="0"/>
  <Setter Property="Padding" Value="10,6,10,6"/>
  <Setter Property="FontSize" Value="12"/>   <!-- Cell -->
  <Setter Property="HasDropShadow" Value="False"/>
  <Setter Property="Template">
    <Setter.Value>
      <ControlTemplate TargetType="ToolTip">
        <Border Background="{TemplateBinding Background}" CornerRadius="4" Padding="{TemplateBinding Padding}" Margin="0,0,6,6">
          <Border.Effect><DropShadowEffect BlurRadius="6" ShadowDepth="2" Direction="270" Opacity="0.2" Color="#000000"/></Border.Effect>
          <ContentPresenter TextElement.Foreground="{TemplateBinding Foreground}"/>
        </Border>
      </ControlTemplate>
    </Setter.Value>
  </Setter>
</Style>
```

使う所:

```xml
<!-- 正規表現（search_bar.xaml） -->
<CheckBox x:Name="RegexCheck" ...>
  <CheckBox.ToolTip>
    <ToolTip Style="{StaticResource Tooltip.Dark}">
      <StackPanel>
        <TextBlock Text="{search.regex.tooltip の 1 行目}"/>
        <TextBlock Text="{search.regex.tooltip の 2 行目}"/>
      </StackPanel>
    </ToolTip>
  </CheckBox.ToolTip>
</CheckBox>

<!-- 検索対象のツリーの行（target_tree.xaml） -->
<Grid ToolTipService.Placement="Mouse" ToolTipService.HorizontalOffset="10" ToolTipService.VerticalOffset="2">
  <Grid.ToolTip>
    <ToolTip Style="{StaticResource Tooltip.Dark}">
      <TextBlock Text="{Binding ToolTip}" TextWrapping="NoWrap"/>
    </ToolTip>
  </Grid.ToolTip>
</Grid>
```

- `Placement="Mouse"` は、カーソルの絵の下に出す。Figma（176:2571）はカーソルの先から右 10・下 22。`VerticalOffset` はカーソルの大きさで変わるので、2 は案。画面で確かめて合わせる（受け入れ）。

### ファイル内の対象のメニュー（search_bar.xaml の中）

```xml
<Popup x:Name="ScopeMenuPopup" PlacementTarget="{Binding ElementName=ScopeButton}"
       Placement="Bottom" VerticalOffset="8" StaysOpen="False" AllowsTransparency="True">
  <Border Margin="0,0,12,12" Width="249" Background="{StaticResource Bg.Surface}"
          BorderBrush="{StaticResource Border.Soft}" BorderThickness="1" CornerRadius="6" Padding="0,6,0,8">
    <Border.Effect><DropShadowEffect BlurRadius="12" ShadowDepth="4" Direction="270" Opacity="0.14" Color="#000000"/></Border.Effect>
    <StackPanel>
      <!-- 1 行目: 本文（常にオン・使えない） -->
      <CheckBox x:Name="ScopeBody" Style="{StaticResource ScopeMenuItem}" IsChecked="True" IsEnabled="False">
        <TextBlock>
          <Run Text="本文" Foreground="{StaticResource Ink.Placeholder}"/>
          <Run Text="　（常に対象）" FontSize="11" Foreground="{StaticResource Ink.Placeholder}"/>
        </TextBlock>
      </CheckBox>
      <CheckBox x:Name="ScopeShape"   Style="{StaticResource ScopeMenuItem}" Content="{search.scope.shape}"/>
      <CheckBox x:Name="ScopeComment" Style="{StaticResource ScopeMenuItem}" Content="{search.scope.comment}"/>
      <!-- 組の見出し「PowerPoint」（search.scope.group.powerpoint）。「共通」の見出しは本文の行の上に置く -->
      <CheckBox x:Name="ScopeNote"    Style="{StaticResource ScopeMenuItem}" Content="{search.scope.note}"/>
      <Rectangle Height="1" Margin="0,4,0,0" Fill="{StaticResource Border.Soft}"/>
      <TextBlock x:Name="ScopeFooter" Margin="12,6,12,0" Style="{StaticResource Micro}"
                 Foreground="{StaticResource Ink.Body}" TextWrapping="NoWrap" Text="{search.scope.reset}"/> <!-- 「既定に戻す」。押せる行（Button にする）。既定のままなら IsEnabled=False -->
    </StackPanel>
  </Border>
</Popup>
```

- 「本文」の文字は `search.scope.body`（「本文　（常に対象）」）を、`splitTipLink` と同じく 2 つに分けて入れる（「（常に対象）」だけ 11）。
- 区切りの線の上の間（4）は Figma に値が無い。**未定**。
- `ScopeMenuItem`（CheckBox の Style）は theme に無い。4 章の値で足す。

## 4. 寸法と色

色は theme のキー（値）で書く。キーが無い色は、足すキーの名前の案を書く（`../figma_wpf_map.md` の「食い違い」に足す案）。

### バナー

| 要素 | 大きさ | 余白 | 地の色 | 枠 | 文字の Style | 文字の色 | 角丸 | アイコン |
|---|---|---|---|---|---|---|---|---|
| BannerRoot | 横 Fill × 40 | 左 16・右 12・上下 8。中の間 10 | 種類ごと（下の表） | 左 3（種類の色） | – | – | 0 | – |
| BannerIcon | 16 × 16 | – | – | – | – | – | – | 種類ごと（下の表）。線 1.33、種類の色 |
| BannerText | 横 Fill × Hug | – | – | – | 12 Medium（theme に無い。`Label.Medium` を足す案） | 種類ごと | – | – |
| BannerAction | Hug × Hug（約 24） | 左右 10・上下 4 | Bg.Surface（#FFFFFF） | 1（種類の色） | 11 SemiBold（theme に無い。`Meta.SemiBold` を足す案） | 種類の色 | 4 | – |

| 種類 | Figma | 地 | 帯・アイコン・ボタンの色 | 文の色 | アイコン（Lucide） |
|---|---|---|---|---|---|
| info | 106:54 | Select.Soft（#E1F2FF） | Accent（#0078D4） | `Banner.Info.Ink`（#0B3D66） | `info` |
| warn | 106:60 | Warn.Soft（#FFF5E0） | Warn（#BA7D00） | `Banner.Warn.Ink`（#5C3D00） | `circle-alert` |
| success | 106:72 | Ok.Soft（#E0F7E0） | Ok（#218A21） | `Banner.Ok.Ink`（#124D12） | `circle-check` |
| danger | 106:66 | Danger.Soft（#FFE6E6） | Danger.Text（#D13438） | `Banner.Danger.Ink`（#7A1C1F） | `circle-x` |

- 文字の色は変数にした（決定。Figma の変数 `Banner/Info/Ink`・`Banner/Warn/Ink`・`Banner/Ok/Ink`・`Banner/Danger/Ink`）。地は今ある変数（Select.Soft・Warn.Soft・Ok.Soft・Danger.Soft）につないである。theme に `Banner.*.Bg` を足すときは、この 4 つの別名にする。
- danger は検索の画面では使わない（Figma の画面に無い）。部品としては持つ。

### 高速検索の吹き出し・ツールチップ

| 要素 | 大きさ | 余白 | 地の色 | 枠 | 文字の Style | 文字の色 | 角丸 | アイコン |
|---|---|---|---|---|---|---|---|---|
| FastSearchTipPopup の本体 | 幅 Hug（文の幅は最大 280）× Hug | 左右 10・上下 7 | #333840（キー無し。`Tip.Bg` を足す案） | 無し。影 0 2 8 rgba(0,0,0,.15) | Meta（11 Regular） | #FFFFFF（キー無し。`Tip.Ink` を足す案。Bg.Surface と同じ値） | 5 | – |
| FastSearchTipArrow | 10 × 6（上向き） | 本体の上に重ねる（本体の上端から −5） | Tip.Bg | – | – | – | – | – |
| FastSearchTipLink | Hug | – | – | – | Meta・下線 | Tip.Ink | – | – |
| Tooltip.Dark（正規表現・フルパス） | Hug × Hug | 左右 10・上下 6 | Tip.Bg | 無し。影 0 2 6 rgba(0,0,0,.2) | 12 Regular（Cell） | Tip.Ink | 4 | – |

- 135:1801 の矢印は、本体（幅 300）の左から 145 にある。真ん中と読んで `HorizontalAlignment="Center"` にした。
- ツールチップは 2 つの見た目がある（吹き出し: 角丸 5・余白 7・11、ツールチップ: 角丸 4・余白 6・12）。1 つにそろえるかは決めること。

### ファイル内の対象のメニュー

- **メニューの並び（2026-10-04 に決定）**: 上から、組の見出し「共通」（10 Medium・Ink.Body #5F6368、余白 左右 12・上 6・下 2）→ 本文（常に対象。使えない）・図形・コメント → 組の見出し「PowerPoint」→ ノート → 区切りの線 → 「既定に戻す」。
  - アプリごとの組にしたのは、あとで項目（例: Excel の組）を足せるようにするため。今回足すのはここに書いた項目だけ。
  - 種類のチップで外したアプリの組は、見出しと項目を薄くする（押せない）。例: PowerPoint のチップを外したら「PowerPoint」の組が薄くなる。
  - 以前の下の注意「Excel・Word・PowerPoint のファイルが対象です」は消した。

| 要素 | 大きさ | 余白 | 地の色 | 枠 | 文字の Style | 文字の色 | 角丸 | アイコン |
|---|---|---|---|---|---|---|---|---|
| ScopeMenuPopup の本体 | 249 × Hug | 上 6・下 8 | Bg.Surface | 1 Border.Soft（#E0E2E5）。影 0 4 12 rgba(0,0,0,.14) | – | – | 6 | – |
| 項目（ScopeMenuItem） | 横 Fill × 28 | 左右 12。箱と文字の間 8 | 透明 | – | Cell（12 Regular） | Ink.Strong（#202124） | – | – |
| チェックの箱（オフ） | 14 × 14 | – | Bg.Surface | 1.2 Ink.Subtle（#80868B） | – | – | 3 | – |
| チェックの箱（オン） | 14 × 14 | – | Accent | 無し | – | – | 3 | `check` 白、線は 14 × 2/24 ≒ 1.17 |
| 本文の項目（使えない・オン） | 横 Fill × 28 | 同上 | 透明 | – | Cell ＋「（常に対象）」は Meta（11） | Ink.Placeholder（#9AA0A6） | – | 箱は薄い青（値は Figma で決めること） |
| 区切りの線 | 横 Fill × 1 | – | Border.Soft | – | – | – | – | – |
| ScopeFooter | Hug × Hug、1 行 | 上 6・左右 12 | – | – | Micro（10 Regular） | Ink.Body（#5F6368） | – | – |
| ScopeButton（開いているとき） | search_bar.md のとおり | 同 | Bg.Surface | 1 Accent（201:5488 の見本だけ。決めること） | – | – | 4 | chevron-down 10 |

- ホバー（項目の地）は Figma に無い。**未定**。

## 5. 状態ごとの見え方

### バナー

`search.md` の `getSearchScreenView` が `TopBanner`・`MidBanner` を決め、`getBannerView`（9 章）が種類と文言を決める。

| 種類の名前（`getSearchScreenView` の値） | 入る条件 | 出す口 | 種類（色） | 文 | ボタン | 画面 |
|---|---|---|---|---|---|---|
| `updating` | インデックスを更新している | TopBannerHost（主な領域の一番上） | info | `banner.updating` | `banner.updating.action` | H-B（108:1116） |
| `interrupted` | 前回の更新が途中で止まっている | TopBannerHost | warn | `banner.interrupted` | `banner.interrupted.action` | H-P（108:1419） |
| `saved` | 検索結果を保存した | TopBannerHost | success | `banner.saved` | `banner.saved.action` | H-saved（158:527） |
| `stopped` | 検索を中止した | MidBannerHost（検索バーの下） | info | `banner.stopped` | 無し | E16（157:1238） |
| `''` | 上のどれでもない | – | 口を Collapsed | – | – | – |

- バナーは主な領域の幅（ナビの右だけ）。端まで伸ばし、角は四角。
- `saved` をいつ消すか（次の検索で消す・閉じるボタンを付けるなど）は Figma に無い。**未定**。閉じるボタンもどのバナーにも無い。

### 高速検索の吹き出し

| 状態 | 入る条件 | 見え方 |
|---|---|---|
| 閉じている | 既定。バッジに info が無いとき（`getFastSearchView` の `Tooltip` が空）はいつもこれ | 出さない |
| 開いている | info のあるバッジにマウスを乗せた | バッジの下、真ん中に矢印を合わせて出す。文言は `getFastSearchView` の `Tooltip` |

- 文言は search_bar.md のバッジの表の 1・2・4・5（`search.fast.nows.tip`・`search.fast.regex.tip`・`search.fast.pending.tip`・`search.fast.partial.tip`）。
- 135:1801（正規表現の吹き出し）は 2 の吹き出しと読んだ。H-R（135:1498）の枠の外に置いてあり、画面のどこに出すかは Figma に無い。バッジの下に出す案にした（決めること）。

### ツールチップ（暗い地）

| どこ | 入る条件 | 文言 | 見え方 |
|---|---|---|---|
| RegexCheck | マウスを乗せた | `search.regex.tooltip`（2 行） | Tooltip.Dark。WPF の既定の位置（カーソルの下） |
| ツリーの行 | マウスを乗せた | `Tree.ToolTip`（target_tree.md。元のフォルダのフルパス、1 行・折り返さない） | Tooltip.Dark。カーソルの右 10・下 22（176:2571） |

- ツリーの行のホバーの地（#E4E7EA、キー無し）は target_tree.md で決める。

### ファイル内の対象のメニュー

| 状態 | 入る条件 | 見え方 |
|---|---|---|
| 閉じている | 既定 | 出さない。ScopeButton は search_bar.md の見た目 |
| 開いている | ScopeButton を押した | ボタンの下 8 に、左端をそろえて出す（201:5488）。項目のオン・オフは今のファイル内の対象のとおり |

- 画面 H-範囲・H-範囲2 では、メニューが x780・y112 にあり、ボタン（x552〜632、H-範囲2 は 696 まで）の下に付いていない。見本 201:5488 に合わせた（決めること）。

## 6. 文言

バナーの文言は `search.md`、ファイル内の対象・高速検索・正規表現の文言は `search_bar.md` で決めたものを使う。同じ ID を使い、ここで足すものは無い。

| ID | 全文 | 使う状態 |
|---|---|---|
| `banner.updating` | インデックスを更新しています（{インデックス名} {済み件数} / {全件数} 件） | updating |
| `banner.updating.action` | 進み具合を見る | updating |
| `banner.interrupted` | 前回の更新が途中です（残り {件数} 件） | interrupted |
| `banner.interrupted.action` | 続きから再開 | interrupted |
| `banner.saved` | 検索結果を保存しました | saved |
| `banner.saved.action` | フォルダを開く | saved |
| `banner.stopped` | 検索を中止しました（見つかった {件数} 件を表示しています） | stopped（ボタン無し） |
| `search.regex.tooltip` | 正規表現を使う<br>オンにすると高速検索は使えません | RegexCheck のツールチップ |
| `search.fast.*.tip` | search_bar.md の 6 章 | 高速検索の吹き出し |
| `search.scope.*` | search_bar.md の 6 章 | ファイル内の対象のメニュー |

- 例: 「インデックスを更新しています（営業部 1,200 / 3,400 件）」「前回の更新が途中です（残り 820 件）」「検索を中止しました（見つかった 14 件を表示しています）」。数は 3 桁ごとに「,」。
- キー操作のヒントは書かない。

## 7. 操作と振る舞い

| きっかけ | 何が起きるか | 呼ぶ関数 | 次の状態 |
|---|---|---|---|
| バナーの［進み具合を見る］ | インデックス管理を開き、更新中のインデックスを選ぶ | `selectNav 'Index'`（search.md） | インデックス管理の画面 |
| バナーの［続きから再開］ | 前回の続きから更新を始める | 今の再開の処理（search.md） | H-B |
| バナーの［フォルダを開く］ | 保存したファイルのフォルダをエクスプローラーで開く | `Start-Process explorer.exe "/select,…"` | H-saved のまま |
| info のあるバッジにマウスを乗せる | 400ms 待って吹き出しを開く | `showFastSearchTip` | 開いている |
| バッジからも吹き出しからもマウスが離れる | 200ms 待って閉じる（吹き出しへ移る間に閉じないため） | `hideFastSearchTip` | 閉じている |
| 吹き出しの［インデックス管理で確認］ | 吹き出しを閉じ、インデックス管理を開く | `selectNav 'Index'` | インデックス管理の画面 |
| ScopeButton を押す | メニューを開く。開いていれば閉じる | `openScopeMenu` | 開いている |
| 図形・コメント・ノートを押す | その範囲のオン・オフを切り替える。メニューは開いたまま | `setSearchScope` | 開いている |
| 本文を押す | 何も起きない（使えない） | – | – |
| メニューの外を押す | 閉じる（StaysOpen=False）。ボタンの文言を決め直す | `getScopeButtonText`（search_bar.md） | 閉じている |
| RegexCheck・ツリーの行にマウスを乗せる | WPF の ToolTip の既定の間で出し、離れたら消す | – | – |

- 待つ時間（400ms・200ms）は Figma に無い案。決めること。
- **ホバー・押した**
  - バナーのボタン: Figma に無い。**未定**。
  - メニューの項目: Figma に無い。**未定**。
  - 使えない本文の項目: 5 章のとおり（文字 #9AA0A6、箱は薄い青）。
- **フォーカス**: theme の既定のフォーカスの枠。
- **Tab の順**
  - バナー: BannerAction だけがフォーカスを受ける。検索バーの前（TopBannerHost）・後（MidBannerHost）の順になる。
  - メニューが開いているとき: ScopeShape → ScopeComment → ScopeNote。本文はフォーカスを受けない。閉じたら ScopeButton に戻す。
  - 吹き出し: マウスでだけ出すので、Tab で入らない。リンクにキーボードで届く道は無い（決めること）。
- **ツールチップの文言**: 6 章のとおり。ほかには付けない。

## 8. リサイズ

- **バナー**: 横は主な領域いっぱいに伸びる。高さ 40 のまま。文は 1 行で、入らないところは「…」で切る。ボタンは縮めない。
- **吹き出し・メニュー**: Popup なので、付けた先（バッジ・ScopeButton）に付いて動く。窓の端からはみ出すときは、WPF が内側に寄せる。
- **ツールチップ**: カーソルに付く。

| 窓 | 主な領域の幅 | 見え方 |
|---|---|---|
| 1024×640 | 804 | バナーの文が長いとき（`banner.updating` に長いインデックス名）は「…」で切れる。ボタンは右に残る。検索バーが 2 行に折り返すと、ScopeButton とバッジの位置が変わるが、メニュー・吹き出しは付いて動く |
| 1280×820 | 1060 | Figma のとおり |
| 1600×1000 | 1380 | バナーは右へ伸び、文の右が空く。メニュー・吹き出しの位置は付けた先のまま |

## 9. 判断層

`scripts/tebunko/ui/search/overlays_view.ps1`。`$ui`・`$window`・WPF の型に触らない。戻り値は hashtable。

### getBannerView

- 入力: `$name`（`getSearchScreenView` の `TopBanner`・`MidBanner` の値）・`$values`（文言の {…} に入れる値の hashtable）
- 出力

| キー | 中身 |
|---|---|
| `Show` | bool。`$name` が空なら False |
| `Tone` | `info`・`warn`・`success`・`danger` |
| `Icon` | Lucide の名前（`info`・`circle-alert`・`circle-check`・`circle-x`） |
| `Text` | 文（{…} を入れたもの） |
| `ActionText` | ボタンの文。ボタンが無ければ空 |
| `Action` | ボタンで呼ぶものの名前: `openIndexProgress`・`resumeIndex`・`openSavedFolder`・`''` |

```powershell
It '<Name>' -TestCases @(
  @{ Name='更新中'; N='updating'; V=@{ IndexName='営業部'; Done=1200; Total=3400 }; Show=$true; Tone='info'; Icon='info'
     Text='インデックスを更新しています（営業部 1,200 / 3,400 件）'; Act='進み具合を見る'; Action='openIndexProgress' }
  @{ Name='途中';   N='interrupted'; V=@{ Rest=820 }; Show=$true; Tone='warn'; Icon='circle-alert'
     Text='前回の更新が途中です（残り 820 件）'; Act='続きから再開'; Action='resumeIndex' }
  @{ Name='保存';   N='saved'; V=@{}; Show=$true; Tone='success'; Icon='circle-check'
     Text='検索結果を保存しました'; Act='フォルダを開く'; Action='openSavedFolder' }
  @{ Name='中止';   N='stopped'; V=@{ Count=14 }; Show=$true; Tone='info'; Icon='info'
     Text='検索を中止しました（見つかった 14 件を表示しています）'; Act=''; Action='' }
  @{ Name='無し';   N=''; V=@{}; Show=$false; Tone=''; Icon=''; Text=''; Act=''; Action='' }
) {
  param($N, $V, $Show, $Tone, $Icon, $Text, $Act, $Action)
  $v = getBannerView $N $V
  $v.Show | Should -Be $Show; $v.Tone | Should -Be $Tone; $v.Icon | Should -Be $Icon
  $v.Text | Should -Be $Text; $v.ActionText | Should -Be $Act; $v.Action | Should -Be $Action
}
```

### getBannerColors

- 入力: `$tone`
- 出力: `@{ Background; Accent; Ink }`（theme のキーの名前）

```powershell
It '<Tone>' -TestCases @(
  @{ Tone='info';    Bg='Banner.Info.Bg';    Accent='Accent';      Ink='Banner.Info.Ink' }
  @{ Tone='warn';    Bg='Banner.Warn.Bg';    Accent='Warn';        Ink='Banner.Warn.Ink' }
  @{ Tone='success'; Bg='Banner.Ok.Bg';      Accent='Ok';          Ink='Banner.Ok.Ink' }
  @{ Tone='danger';  Bg='Banner.Danger.Bg';  Accent='Danger.Text'; Ink='Banner.Danger.Ink' }
) { param($Tone,$Bg,$Accent,$Ink); $v = getBannerColors $Tone; $v.Background | Should -Be $Bg; $v.Accent | Should -Be $Accent; $v.Ink | Should -Be $Ink }
```

- キーの名前は「決めること」が決まるまでの案。

### getScopeMenuItems

- 入力: `$includeShapes`・`$includeComments`・`$includeNotes`（bool）
- 出力: `@{ Key; Checked; Enabled }` の配列。並びは 本文・図形・コメント・ノート。

```powershell
It '<Name>' -TestCases @(
  @{ Name='すべてオン'; S=$true;  C=$true;  N=$true;  Checked=@($true,$true,$true,$true);   Enabled=@($false,$true,$true,$true) }
  @{ Name='図形だけ';   S=$true;  C=$false; N=$false; Checked=@($true,$true,$false,$false); Enabled=@($false,$true,$true,$true) }
  @{ Name='すべてオフ'; S=$false; C=$false; N=$false; Checked=@($true,$false,$false,$false); Enabled=@($false,$true,$true,$true) }
) { param($S,$C,$N,$Checked,$Enabled); $v = getScopeMenuItems $S $C $N
    ($v | ForEach-Object Key) | Should -Be @('body','shape','comment','note')
    ($v | ForEach-Object Checked) | Should -Be $Checked; ($v | ForEach-Object Enabled) | Should -Be $Enabled }
```

### splitTipLink

文言を、ふつうの文と、全角の［ ］で囲んだリンクに分ける。

- 入力: `$text`
- 出力: `@{ Body; Link }`。［ ］が無ければ `Link` は空。`Body` の末尾の全角の空白は取る。

```powershell
It '<Text>' -TestCases @(
  @{ Text='検索はできますが時間がかかります　［インデックス管理で確認］'; Body='検索はできますが時間がかかります'; Link='インデックス管理で確認' }
  @{ Text='インデックスの更新が終わると使えます　［インデックス管理で確認］'; Body='インデックスの更新が終わると使えます'; Link='インデックス管理で確認' }
  @{ Text='正規表現をオフにすると速く検索できます'; Body='正規表現をオフにすると速く検索できます'; Link='' }
  @{ Text='2 フォルダはパスが長いため、通常の検索で調べます（検索結果は変わりません）'; Body='2 フォルダはパスが長いため、通常の検索で調べます（検索結果は変わりません）'; Link='' }
) { param($Text,$Body,$Link); $v = splitTipLink $Text; $v.Body | Should -Be $Body; $v.Link | Should -Be $Link }
```

- 画面ではリンクを「［インデックス管理で確認］」と［ ］付きで出す（Figma のとおり）。

## 10. 画面層

`scripts/tebunko/ui/search/overlays.ps1`。

1. **`setBanner $host $name $values`**（search.md の `updateSearchScreen` から呼ぶ）
   - `$v = getBannerView $name $values`。`$v.Show` が False なら `$host.Visibility = Collapsed` で終わる。
   - `$host.Content` が空なら、`banner.xaml` を XamlReader で読んで差す（口ごとに 1 回）。
   - `$c = getBannerColors $v.Tone` の キーで、`BannerRoot.Background`・`BorderBrush`、`BannerIcon.Stroke`・`Data`、`BannerText.Foreground`、`BannerAction.BorderBrush`、`BannerActionText.Foreground` を入れ替える（`$window.FindResource`）。
   - `BannerText.Text = $v.Text`。`$v.ActionText` が空なら `BannerAction` を Collapsed。
   - `BannerAction.Click` は、読んだときに 1 回だけつなぎ、押したら `$host.Tag`（そのときの `$v.Action`）の名前の関数を呼ぶ。
   - 進み具合（{済み件数}）は、インデックスの更新の通知のたびに呼び直す。通知は裏のスレッドから来るので、`Dispatcher.BeginInvoke` で画面のスレッドに戻してから呼ぶ。
2. **`showFastSearchTip` / `hideFastSearchTip`**（search_bar.md の `initSearchBar` から、バッジの `MouseEnter`・`MouseLeave` につなぐ）
   - `DispatcherTimer` で 400ms 待ってから開く。開く前に `getFastSearchView` の `Tooltip` を `splitTipLink` で分け、`FastSearchTipBody.Text`・リンクの有無を入れる。
   - `FastSearchTipPopup` の中の `MouseEnter` で閉じるタイマーを止め、`MouseLeave` で 200ms のタイマーをかける。
   - `FastSearchTipLink.Click` → `FastSearchTipPopup.IsOpen = $false` → `selectNav 'Index'`。
   - バッジの中身が変わったら（`setFastSearchView`）、開いていれば閉じる。
3. **`openScopeMenu`**（search_bar.md の `ScopeButton.Click`）
   - `getScopeMenuItems` の値を `ScopeShape`・`ScopeComment`・`ScopeNote` の `IsChecked` に入れてから `ScopeMenuPopup.IsOpen = $true`。
   - 各 CheckBox の `Checked`・`Unchecked` → `setSearchScope`（今の検索範囲の設定を変える）。
   - `ScopeMenuPopup.Closed` → `ScopeButtonText.Text = getScopeButtonText …`（search_bar.md）、`ScopeButton.Focus()`。
4. **ツールチップ**: XAML の `ToolTip` だけで出す。画面層のコードは要らない。ツリーの文言は target_tree.md の `getTreeItemToolTip`。
5. 重い処理は無い。バナーの件数は、更新の処理が裏で数えたものを受け取るだけにする。

## 11. 受け入れ

| 見比べる png（`../png/`） | もの | 見るところ |
|---|---|---|
| `108_1116.png`（H-B） | info のバナー | 主な領域の一番上、高さ 40、左の帯 3 の青、地 #E1F2FF、info のアイコン 16、文 12 の濃い青、白地に青い枠のボタン |
| `108_1419.png`（H-P） | warn のバナー | 地 #FFF5E0、帯・アイコン・ボタンが #BA7D00、文 #5C3D00。［続きから再開］が 40 の中に収まり、下で切れない |
| `158_527.png`（H-saved） | success のバナー | 地 #E0F7E0、帯 #218A21、circle-check |
| `157_1238.png`（E16） | 中止のバナー | 検索バーの下、info、ボタン無し |
| `135_1801.png`・`136_89.png` | 高速検索の吹き出し | 地 #333840、角丸 5、11 の白い文、上向きの矢印が真ん中、最大幅 280 |
| `201_5488.png`（ファイル内の対象） | メニュー・正規表現のツールチップ | メニュー 249 幅・角丸 6・影、項目 28、本文が薄い字で使えない、組の見出し「共通」「PowerPoint」、区切りの線の下に「既定に戻す」。ボタンの下 8 に左端をそろえて付く。ツールチップは角丸 4・12 の白い文・2 行 |
| `213_6592.png`・`213_6958.png`（H-範囲・H-範囲2） | メニュー | 項目と文言。位置は見本 201:5488 に合わせる（画面の位置とは違ってよい） |
| `176_2571.png`（フルパスのツールチップ） | ツリーのツールチップ | カーソルの右下に、1 行・折り返さないフルパス。角丸 4 |

- 窓を 1024×640 と 1600×1000 にして、バナーが主な領域いっぱいに伸びること・メニューと吹き出しが付けた先から外れないことを見る。
- 判断層のテスト（9 章）が通ること。

## 決めること

1. **高速検索の説明を Popup（押せるリンク付き）にするか、ToolTip（押せない）にするか。** Figma の文に［インデックス管理で確認］がある。この文書は Popup にした。キーボードでリンクに届く道も無い。
2. **吹き出し（135:1801）を出す位置。** H-R の枠の外にあり、画面のどこに出すかが無い。バッジの下、真ん中に矢印を合わせる案にした。
3. **吹き出しを開く・閉じるまでの待ち時間**（案 400ms・200ms）。
4. **暗い地の見た目を 1 つにそろえるか。** 吹き出し（角丸 5・余白 10/7・11・影 0 2 8 .15）とツールチップ（角丸 4・余白 10/6・12・影 0 2 6 .2）の 2 通りがある。
5. **theme に足すキー**
   - `Tip.Bg`（#333840）・`Tip.Ink`（#FFFFFF）
   - `Banner.Info.Ink`（#0B3D66）・`Banner.Warn.Ink`（#5C3D00）・`Banner.Ok.Ink`（#124D12）・`Banner.Danger.Ink`（#7A1C1F）
   - バナーの地: info は Select.Soft と同じ値、warn・success は Warn.Soft・Ok.Soft と同じ値だが、Figma は変数につないでいない。danger の #FDE7E9 は Danger.Soft（#FFE6E6）と違う。どちらに合わせるか。
   - Style `Label.Medium`（12 Medium）・`Meta.SemiBold`（11 SemiBold）・`BannerButton`・`ScopeMenuItem`・`Tooltip.Dark`
6. **バナーの帯の太さ。** Figma の variant は 3。diff_round3 は「4px の色の帯」と書いている。この文書は Figma の 3 にした。
7. **ファイル内の対象のメニューの位置。** 見本 201:5488 はボタンの下 8 に付く。画面 H-範囲・H-範囲2 も 2026-10-04 にボタンの左端・下 8 に付け直した（決定）。
8. **メニューを開いている間の ScopeButton の枠。** 見本は Accent、画面は Border.Input のまま。
9. **本文の項目の箱の色（使えない・オン）。** diff_round4 は「薄い青」とだけあり、値が無い。区切りの線の上の間（4）も値が無い。
10. **バナーのボタン・メニューの項目のホバー・押した見た目。** Figma に無い。
11. **保存のバナー（saved）をいつ消すか。** 閉じるボタンはどのバナーにも無い。
12. **danger のバナーを使う場面。** 検索の画面には無い。
13. **フルパスのツールチップの位置の合わせ方。** Figma はカーソルの先から右 10・下 22。WPF の `Placement="Mouse"` はカーソルの絵の下が基準なので、`VerticalOffset` を画面で合わせる。

## ツールチップをすべて出した見本（2026-10-04）

Figma のページ「06 ツールチップ」に、画面ごとのツールチップ・吹き出しを 1 枚にまとめた見本がある（実際に出るのは、カーソルを当てた 1 つだけ）。

| 画面 | 付ける先 | 文言 | 決まり具合 |
|---|---|---|---|
| 検索 | RegexCheck | `search.regex.tooltip`（2 行） | 決定 |
| 検索 | 検索対象のツリーの行 | 元のフォルダのフルパス | 決定 |
| 検索 | 高速検索のバッジ（info があるとき） | `search.fast.*.tip` | 決定 |
| 検索・インデックス管理 | ステータスバーの文 | 同じ文の全文 | 決定 |
| インデックス管理 | 題の横の help-icon | 検索するフォルダを登録する画面です。<br>登録したフォルダは、インデックスを更新すると検索できます。 | 案（文言は未定） |
| インデックス管理 | 「ステータス」の列の help-icon | `index.list.col.status.help` | 決定 |
| インデックス管理 | 「基本設定」の help-icon | インデックスの名前と、検索するフォルダです。 | 案（文言は未定） |
| インデックス管理 | 「インデックス情報」の help-icon | 最後に更新した日時と、読み込んだファイルの数です。 | 案（文言は未定） |
| インデックス管理 | 一覧の名前・パス、詳細の「フォルダ」の箱 | 全文 | 決定 |
| インデックス管理 | 無効の「削除」（更新中） | `index.list.menu.delete.disabled` | 決定 |
| インデックス管理 | 無効のボタン（エクスポート・インポート中） | `index.list.busy` | 決定 |

- 付けないもの: 種類のチップ・［ファイル内の対象］・検索ボタン・結果の行・プレビューの題・ナビの項目・窓のボタン。
- 決定: 一覧の行のステータス・高速検索の札にツールチップを付ける（文は `../index/index_list.md`。見本は TT-X の 8・9）。高速検索の詳しい説明はツールチップに置く。

## 異常系のバナーの画面（「07 プロトタイプ」。2026-10-04 に足した）

「04 異常系」のカードのうち、既存の画面にバナーを載せるだけのものを、画面として足した。文言の「案」は、カードに「推測」とあり、まだ決まっていないもの。

| 記号 | node ID | 元の画面 | バナー | 置き場所 | 文言 | ボタン |
|---|---|---|---|---|---|---|
| E2 | 505:18949 | H0 | warn | 主な領域の一番上 | 既定の設定で起動しました | ［設定を開く］→ 設定 |
| E3 | 505:19073 | X | danger | 主な領域の一番上 | 保存先が使えません（アクセスが拒否されました） | ［設定を開く］→ 設定 |
| E11 | 505:19410 | H | info | 検索バーの下（結果の上） | まだ取り込んでいないファイルがあります | なし |
| E15 | 505:19853 | H | warn | 検索バーの下（結果の上） | 10,000 件で打ち切りました。条件を絞り込むと、すべての結果を表示できます。 | なし |
| E23 | 505:20289 | H | danger | 主な領域の一番上 | 保存できませんでした（保存先に書き込めません）（案） | なし |
| E42 | 505:20725 | X | danger | 主な領域の一番上 | 更新に失敗しました（営業部） | ［ログを開く］［もう一度］→ 更新中 |
| E45 | 505:21059 | X | danger | 主な領域の一番上 | 削除できませんでした（営業部）（案） | なし |
| E64 | 505:21392 | P | danger | 主な領域の一番上 | 終了できませんでした（1 件）（案） | なし |

- E42 の一覧の行の札は、元の画面（X）のままにしてある（失敗した行を「エラー」にする絵はまだ無い）。
- バナーのほかの見せ方（欄の赤字・無効＋理由・空の表示・ダイアログの中）の異常系は、まだ画面にしていない。
