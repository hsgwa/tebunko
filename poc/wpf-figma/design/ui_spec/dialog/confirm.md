# 確認・進み具合・エラー（confirm）

## 1. 役割

操作の前に確かめる・待たせる・失敗を伝える、共通のダイアログ。題・主文・補足（または詳細の箱）・ボタンだけの小さな形で、5 つの型（Type）がある。どの画面からも同じ 1 つの XAML を使い、中身は判断層の `newConfirmView` が状態（DC-stop など）から決める。

- Figma の部品: `xaml/dialog/confirm` 373:50。`Type` と `状態` の 2 つの項目で表す。部品の項目は title・message・detail・showDetail・primary・secondary・choice。
- 型ごとの既定: normal 373:51・danger 373:66・choice 373:83・progress 373:100・error 373:115。
- 状態の一覧（106 の DC-*・E54・DI-E*、148 の P-DC-*）:

  | Type | 状態 | ノード | 大きさ | 使う所 |
  |---|---|---|---|---|
  | normal | DC-stop | 373:338 | 520×200 | インデックス管理で更新を中止する |
  | normal | DC-kill | 373:879 | 520×200 | Office の終了で［バックグラウンドを終了］ |
  | normal | DC-ws | 373:924 | 520×200 | 設定で保存先を変える |
  | normal | DC-reset | 373:969 | 520×200 | 設定で既定の場所に戻す |
  | normal | DC-close | 373:1014 | 520×200 | 更新中に窓を閉じる |
  | normal | DC-src | 373:1059 | 520×200 | 検索結果から開こうとしたファイルが無い |
  | normal | P-DC-stop | 373:1148 | 520×200 | DC-stop と同じ文言 |
  | normal | P-DC-ws | 373:1217 | 520×200 | DC-ws と同じ文言 |
  | normal | P-DC-reset | 373:1251 | 520×200 | DC-reset と同じ文言 |
  | normal | P-DC-kill | 373:1285 | 520×200 | DC-kill と同じ文言 |
  | normal | P-DC-close | 373:1319 | 520×200 | DC-close と同じ文言 |
  | normal | P-DC-refolder | 373:1353 | 520×200 | インデックス管理で、元のフォルダを選び直す |
  | danger | DC-del | 373:1489 | 520×200 | インデックスを削除する |
  | danger | DC-kill-vis | 373:1540 | 520×200 | 画面に表示中の Office も終了する |
  | danger | P-DC-del | 373:1607 | 520×200 | DC-del と同じ文言 |
  | choice | DC-ws-idx | 373:1647 | 620×200 | 保存先に選んだフォルダにインデックスがある |
  | progress | DC-stopping | 373:1715 | 520×200 | 更新を中止している |
  | progress | DC-close-wait | 373:1760 | 520×200 | 閉じる前に中止を待っている |
  | progress | DC-move | 373:1825 | 520×200 | 保存先を移している |
  | error | DC-err | 373:1959 | 520×200 | 予期しないエラー |
  | error | E54 | 373:1999 | 520×200 | 保存先に使えないフォルダを選んだ |
  | error | DI-E1 | 377:11760 | 440×139 | インポートの zip が読めない |
  | error | DI-E2 | 377:11962 | 440×156 | インポートの zip が新しい版 |
  | error | DI-E3 | 377:11990 | 440×165 | インポートに失敗した |

- P-DC-* は、主のボタンの作りだけが違う（文字色が直書きの白・中央揃え・はみ出しを切らない）。見た目は DC-* と同じなので、実装では分けない。
- 画面の上に重ねた見本: DC-ws 1024×640（408:33891）・DC-ws 1600×1000（408:33936）。
- `../png/` にこのダイアログの画像は無い。見比べは Figma の get_screenshot で行う（11 章）。

## 2. 置き場所

- XAML: `scripts/shared/xaml/dialog_confirm.xaml`（今のファイルを書き直す。今は幅 560・事実の行・選択肢の縦並び）。
- 画面層: `scripts/shared/ui/shell.ps1` の `showConfirm`（書き直す）と、新しい `showScrim`・`showProgress`。判断層: `scripts/tebunko/ui/confirm_view.ps1`（新しい。文言はツール固有なので `tebunko/` に置く。`shared/` はツールを知らない決まり）。
- 型: Window。`ShowDialog()` で出す。Owner は主の窓。progress だけは `Show()` で出し、処理の終わりに画面層から閉じる（10 章）。
- 置き方: 主の窓の中央（`WindowStartupLocation="CenterOwner"`）。見本では 1024×640 で x=252・y=220、1600×1000 で x=540・y=400（どちらも上下左右の真ん中）。
- 暗幕（scrim）: 出している間、主の窓全体（窓のタイトルバーを含む）を黒の 30%（rgba(0,0,0,0.3)）で覆う。about・indexing_confirm・export・import も同じ暗幕を使う。
  - 作り方: 主の窓と同じ位置・大きさの、枠なし・透明の Window（`AllowsTransparency="True"`・`Background="#4D000000"`・`ShowInTaskbar="False"`・Owner は主の窓）を先に `Show()` し、ダイアログの Owner をこの暗幕の窓にする。閉じたら暗幕も閉じる。
- 窓の枠: `WindowStyle="None"`・`AllowsTransparency="True"`・`ResizeMode="NoResize"`。題のバーは自分で描く。

## 3. 部品の木

```xml
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="ConfirmWindow" Width="560" SizeToContent="Height" WindowStyle="None" AllowsTransparency="True"
        Background="Transparent" ResizeMode="NoResize" WindowStartupLocation="CenterOwner" ShowInTaskbar="False">
  <!-- Width は 外枠の幅（520・620・440）+ 影の余白 40。画面層が型から決める -->
  <Border x:Name="DialogFrame" Margin="20" MinHeight="200" Background="{DynamicResource Bg.Surface}"
          BorderBrush="{DynamicResource Border.Normal}" BorderThickness="1" CornerRadius="8">
    <Border.Effect>
      <DropShadowEffect BlurRadius="20" ShadowDepth="4" Direction="270" Opacity="0.18" Color="#000000" />
    </Border.Effect>
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto" />  <!-- title-bar -->
        <RowDefinition Height="*" />     <!-- body -->
      </Grid.RowDefinitions>

      <!-- title-bar -->
      <Border x:Name="TitleBar" Grid.Row="0" Background="{DynamicResource Bg.Window}" CornerRadius="8,8,0,0"
              BorderBrush="{DynamicResource Border.Normal}" BorderThickness="0,0,0,1" Padding="16,10,12,10">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*" />
            <ColumnDefinition Width="Auto" />
          </Grid.ColumnDefinitions>
          <TextBlock x:Name="TitleText" Grid.Column="0" FontSize="13" FontWeight="Medium" Foreground="{DynamicResource Ink.Value}"
                     VerticalAlignment="Center" TextTrimming="CharacterEllipsis" Text="{confirm.title}" />
          <Button x:Name="TitleCloseButton" Grid.Column="1" Margin="8,0,0,0" Style="{StaticResource TitleBarButton}"
                  Width="10" Height="21">
            <Image x:Name="TitleCloseIcon" Width="10" Height="10" Source="{StaticResource Icon.X}" />
          </Button>
        </Grid>
      </Border>

      <!-- body（上 16・左右 20・下 20、間 12） -->
      <StackPanel x:Name="BodyPanel" Grid.Row="1" Margin="20,16,20,20" Orientation="Vertical">
        <!-- message-row -->
        <Grid x:Name="MessageRow">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="*" />
          </Grid.ColumnDefinitions>
          <!-- danger・error の印（DI-E* では出さない） -->
          <Border x:Name="AlertMark" Grid.Column="0" Width="20" Height="20" CornerRadius="10" Margin="0,0,10,0"
                  VerticalAlignment="Top" Background="{DynamicResource Danger.Text}" Visibility="Collapsed">
            <TextBlock FontSize="12" FontWeight="Bold" Foreground="{DynamicResource Ink.OnAccent}"
                       HorizontalAlignment="Center" VerticalAlignment="Center" Text="!" />
          </Border>
          <TextBlock x:Name="MessageText" Grid.Column="1" Style="{StaticResource Body}" Foreground="{DynamicResource Ink.Value}"
                     TextWrapping="Wrap" VerticalAlignment="Center" Text="{confirm.message}" />
        </Grid>

        <!-- progress の型だけ -->
        <Grid x:Name="ProgressTrack" Margin="0,12,0,0" Height="6" Visibility="Collapsed">
          <Border Background="{DynamicResource Border.Soft}" CornerRadius="3" />
          <ProgressBar x:Name="ProgressFill" Height="6" Minimum="0" Maximum="100" Style="{StaticResource Progress.Thin}" />
        </Grid>

        <!-- 補足（normal・danger・choice・progress） -->
        <TextBlock x:Name="DetailText" Margin="0,12,0,0" FontSize="11" Foreground="{DynamicResource Ink.Note}"
                   TextWrapping="Wrap" Text="{confirm.detail}" />

        <!-- 詳細の箱（error の DC-err・E54） -->
        <Border x:Name="DetailBox" Margin="0,12,0,0" Visibility="Collapsed" Background="{DynamicResource Bg.Window}"
                BorderBrush="{DynamicResource Border.Normal}" BorderThickness="1" CornerRadius="5" Padding="10,8">
          <TextBlock x:Name="DetailBoxText" FontSize="11" Foreground="{DynamicResource Ink.Body}" TextWrapping="Wrap" />
        </Border>

        <!-- buttons（右寄せ、間 8、上 4）。中身は画面層が newConfirmView の Buttons から作る -->
        <StackPanel x:Name="ButtonPanel" Margin="0,16,0,0" Orientation="Horizontal" HorizontalAlignment="Right" />
      </StackPanel>
    </Grid>
  </Border>
</Window>
```

- ボタンの Style は 4 つ。
  - `DialogButton`: secondary と choice。choice は見た目が secondary と同じで、別の Style にしない。
  - `Primary`
  - `Danger`
  - `DialogButton.Disabled`: progress の無効。新しいキー。
- ボタンの上の余白は、本文の間 12 と上の padding 4 を足した 16。
- 今の XAML の `HeadingText`・`FactsPanel`・`FactsList`・`FactDetail`・`ChoicePanel`・`HintText` は置かない（「決めること」）。Figma の確認は、題・主文・補足・ボタンだけ。
- DI-E1〜DI-E3 は寸法が少し違う（4 章の表）。同じ XAML で、画面層が `newConfirmView` の `Compact` を見て値を差し替える。
- `Progress.Thin`（地を透明にし、中身を Accent・角丸 3 にした ProgressBar）は theme に無いキー。

## 4. 寸法と色

### 共通（373 番台）

| 要素 | 大きさ | 余白 | 地 | 枠 | 文字 | 文字の色 | 角丸 | アイコン |
|---|---|---|---|---|---|---|---|---|
| DialogFrame | 幅 520（DC-ws-idx は 620）・最小の高さ 200・高さ Hug | — | Bg.Surface（#FFFFFF） | 1px Border.Normal（#D9DEE3） | — | — | 8 | — |
| 影 | 0 4 20 rgba(0,0,0,.18) | — | — | — | — | — | — | — |
| TitleBar | Hug | 左 16・右 12・上下 10、間 8 | Bg.Window（#F5F7FA） | 下 1px Border.Normal | 13 Medium | Ink.Value（#212126） | 上 8 | x 10×21（Lucide x）。線の太さと色は未定 |
| BodyPanel | — | 上 16・左右 20・下 20、間 12 | — | — | — | — | — | — |
| AlertMark | 20×20 | 文との間 10 | Danger.Text（#D13438） | — | 「!」12 Bold | Ink.OnAccent（#FFFFFF） | 10 | — |
| MessageText | 幅いっぱい | — | — | — | Body（13） | Ink.Value（#212126） | — | — |
| ProgressTrack | 高さ 6・幅いっぱい | — | Border.Soft（#E0E2E5） | — | — | — | 3 | — |
| ProgressFill | 高さ 6 | — | Accent（#0078D4） | — | — | — | 3 | — |
| DetailText | 幅いっぱい | — | — | — | 11 | Ink.Note（#8C949E） | — | — |
| DetailBox | 幅いっぱい | 10/8 | Bg.Window（#F5F7FA） | 1px Border.Normal | 11 | Ink.Body（#5F6368） | 5 | — |
| ButtonPanel | Hug・右寄せ | 上 4、間 8 | — | — | — | — | — | — |
| secondary・choice | 高さ 33 | 16/7 | Bg.Surface | 1px Border.Dialog（#D1D6E0） | 13 Medium | Ink.Strong（#202124） | 6 | — |
| primary | 高さ 31 | 16/7 | Accent（#0078D4） | なし | 13 SemiBold | Ink.OnAccent（#FFFFFF） | 6 | — |
| danger の primary | 高さ 31 | 16/7 | Danger.Text（#D13438） | なし | 13 SemiBold | Ink.OnAccent | 6 | — |
| 無効（progress） | 高さ 33 | 16/7 | Bg.Tag（#F1F3F4） | 1px Border.Soft（#E0E2E5） | 13 Medium | #A0A5AB（キー無し） | 6 | — |

- secondary（33）と primary（31）の高さの差は線の 1px ずつ。WPF では primary にも透明の 1px の線を付けて 33 にそろえるか、Figma のとおり 31 にするかは「決めること」。

### DI-E1〜DI-E3（377 番台）で違う値

| 要素 | 373 番台 | 377 番台 |
|---|---|---|
| DialogFrame | 幅 520・最小の高さ 200 | 幅 440・最小の高さなし（139・156・165） |
| TitleBar の余白 | 左 16・右 12・上下 10 | 左右 16・上下 10 |
| 閉じるアイコン | 10×21 | 20×20 |
| BodyPanel | 上 16・左右 20・下 20、間 12 | 上下 16・左右 20、間 14 |
| MessageText の色 | Ink.Value（#212126） | Ink.Strong（#202124） |
| ButtonPanel の上の余白 | 4 | なし |
| AlertMark | error は出す | 出さない |
| 理由の行（DI-E3） | — | 11、Ink.Body（#5F6368）、枠なし |
| DI-E2 の主のボタン | — | 文字の右に external-link 9.8×9.8（間 6）、線 0.82、Ink.OnAccent |

## 5. 状態ごとの見え方

### 型ごと

| Type | AlertMark | ProgressTrack | DetailText | DetailBox | ボタン（左から） | 既定のボタン | Esc で閉じる |
|---|---|---|---|---|---|---|---|
| normal | 出さない | 出さない | 出す | 出さない | secondary・primary | primary | する（secondary と同じ） |
| danger | 出す | 出さない | 出す | 出さない | secondary・danger | secondary | する |
| choice | 出さない | 出さない | 出す | 出さない | secondary・choice・primary | primary | する |
| progress | 出さない | 出す | 出す | 出さない | 無効のボタン 1 つ | なし | しない |
| error（373） | 出す | 出さない | 出さない | 出す | choice・secondary・primary（状態でボタンの数が変わる） | primary | する |
| error（DI-E*） | 出さない | 出さない | 出さない（DI-E3 は理由の行） | 出さない | secondary・primary、または primary だけ | primary | する |

- 既定のボタン・Esc は Figma に定めが無い。上は今の `showConfirm` の決まり（危ないものはキャンセルを既定、ほかは主のボタンを既定）を引き継いだもの（「決めること」）。
- 題のバーの x は、Esc と同じ扱い（progress では押せない）。

### 状態ごと

| 状態 | Type | 題 | 主文 | 補足／詳細 | ボタン（左から。太字は主） |
|---|---|---|---|---|---|
| DC-stop | normal | `dc.stop.title` | `dc.stop.message` | `dc.stop.detail` | キャンセル・**中止する** |
| DC-kill | normal | `dc.kill.title` | `dc.kill.message` | `dc.kill.detail` | キャンセル・**終了する** |
| DC-ws | normal | `dc.ws.title` | `dc.ws.message` | `dc.ws.detail` | キャンセル・**移動する** |
| DC-reset | normal | `dc.reset.title` | `dc.reset.message` | `dc.ws.detail` | キャンセル・**戻す** |
| DC-close | normal | `dc.close.title` | `dc.close.message` | `dc.close.detail` | 閉じない・**中止して閉じる** |
| DC-src | normal | `dc.src.title` | `dc.src.message` | `dc.src.detail` | キャンセル・**フォルダを選ぶ** |
| P-DC-refolder | normal | `dc.refolder.title` | `dc.refolder.message` | `dc.refolder.detail` | キャンセル・**フォルダを選ぶ** |
| DC-del | danger | `dc.del.title` | `dc.del.message` | `dc.del.detail` | キャンセル・**削除する**（赤） |
| DC-kill-vis | danger | `dc.kill.title` | `dc.killvis.message` | `dc.killvis.detail` | キャンセル・**終了する**（赤） |
| DC-ws-idx | choice | `dc.ws.title` | `dc.wsidx.message` | `dc.wsidx.detail` | キャンセル・今のインデックスを移動する・**そのフォルダのインデックスを使う** |
| DC-stopping | progress | `dc.stop.title` | `dc.stopping.message` | `dc.stopping.detail` | キャンセル（無効） |
| DC-close-wait | progress | `dc.close.title` | `dc.stopping.message` | `dc.closewait.detail` | 閉じない（無効） |
| DC-move | progress | `dc.ws.title` | `dc.move.message` | `dc.move.detail` | キャンセル（無効） |
| DC-err | error | `dc.err.title` | `dc.err.message` | 詳細の箱 `dc.err.box` | 内容をコピー・ログを開く・**閉じる** |
| E54 | error | `e54.title` | `e54.message` | 詳細の箱 `e54.box` | キャンセル・**別のフォルダを選ぶ** |
| DI-E1 | error | `di.e.title` | `di.e1.message` | なし | 閉じる・**別のファイルを選ぶ** |
| DI-E2 | error | `di.e.title` | `di.e2.message` | なし | 閉じる・**リリースノート**（external-link） |
| DI-E3 | error | `di.e.title` | `di.e3.message` | 理由の行 `di.e3.reason` | **閉じる** |

- DC-reset の補足は DC-ws と同じ文。DC-kill-vis の題は DC-kill と同じ。
- progress の 3 つは、ボタンが無効のまま、処理が終わったら画面層が閉じる。
- ProgressFill の幅: DC-move は件数の割合（`{済み} / {全部}`）。DC-stopping・DC-close-wait は割合が無い。見本はどれも中身の幅 300 で止まっている。不定（動き続ける）にするかは未定（「決めること」）。

## 6. 文言

| ID | 全文 | 使う状態 |
|---|---|---|
| confirm.title | 確認 | 既定（部品の初期値） |
| confirm.message | 操作を続けますか？ | 既定 |
| confirm.detail | 補足の説明 | 既定 |
| confirm.secondary | キャンセル | 既定 |
| confirm.primary | OK | 既定 |
| confirm.choice | 選択 | 既定（choice・error） |
| dc.stop.title | 更新の中止 | DC-stop・DC-stopping |
| dc.stop.message | 「{インデックス名}」の更新を中止しますか？ | DC-stop |
| dc.stop.detail | 取り込んだところまでは残ります。あとで続きから再開できます。 | DC-stop |
| dc.stop.ok | 中止する | DC-stop |
| dc.kill.title | Office の終了 | DC-kill・DC-kill-vis |
| dc.kill.message | バックグラウンドの {アプリの並び}（{件数} 件）を終了しますか？ | DC-kill |
| dc.kill.detail | 画面に表示されていないものだけを終了します。 | DC-kill |
| dc.kill.ok | 終了する | DC-kill・DC-kill-vis |
| dc.killvis.message | 画面に表示されている {アプリ}（{ファイル名}）も終了します。保存していない変更は失われます。 | DC-kill-vis |
| dc.killvis.detail | 終了する前に、{アプリ} で保存したかを確認してください。 | DC-kill-vis |
| dc.ws.title | 保存先の変更 | DC-ws・DC-ws-idx・DC-move |
| dc.ws.message | インデックスとログを「{フォルダ}」へ移動します。 | DC-ws |
| dc.ws.detail | 移動中は、検索とインデックスの更新はできません。 | DC-ws・DC-reset |
| dc.ws.ok | 移動する | DC-ws |
| dc.reset.title | 既定の場所に戻す | DC-reset |
| dc.reset.message | インデックスとログを既定の場所「{既定のフォルダ}」へ移動します。 | DC-reset |
| dc.reset.ok | 戻す | DC-reset |
| dc.close.title | 更新中です | DC-close・DC-close-wait |
| dc.close.message | インデックスを更新中です。中止して閉じますか？ | DC-close |
| dc.close.detail | 取り込んだところまでは残ります。次に起動したときに続きから再開できます。 | DC-close |
| dc.close.cancel | 閉じない | DC-close・DC-close-wait |
| dc.close.ok | 中止して閉じる | DC-close |
| dc.src.title | 元のファイルが見つかりません | DC-src |
| dc.src.message | 元のファイルが見つかりません。フォルダを移動した場合は、移動先のフォルダを選んでください。 | DC-src |
| dc.src.detail | 元の場所: {ファイルのパス} | DC-src |
| dc.src.ok | フォルダを選ぶ | DC-src・P-DC-refolder |
| dc.refolder.title | フォルダの再設定 | P-DC-refolder |
| dc.refolder.message | 「{インデックス名}」のフォルダが見つかりません。フォルダを選び直しますか？ | P-DC-refolder |
| dc.refolder.detail | {元のフォルダ} が移動されたか、名前が変更された可能性があります。選び直したフォルダは、次の更新で取り込みます。 | P-DC-refolder |
| dc.del.title | インデックスの削除 | DC-del |
| dc.del.message | 「{インデックス名}」のインデックスを削除しますか？ | DC-del |
| dc.del.detail | 元のファイルは削除されません。 | DC-del |
| dc.del.ok | 削除する | DC-del |
| dc.wsidx.message | 「{フォルダ}」には、すでにインデックスがあります。 | DC-ws-idx |
| dc.wsidx.detail | そのフォルダのインデックスを使うか、今のインデックスを移動するかを選んでください。 | DC-ws-idx |
| dc.wsidx.move | 今のインデックスを移動する | DC-ws-idx |
| dc.wsidx.use | そのフォルダのインデックスを使う | DC-ws-idx |
| dc.stopping.message | 中止しています… | DC-stopping・DC-close-wait |
| dc.stopping.detail | 取り込み中のファイルが終わるまでお待ちください。 | DC-stopping |
| dc.closewait.detail | あと {秒} 秒で強制終了します | DC-close-wait |
| dc.move.message | 移動しています…（{済み} / {全部} 件） | DC-move |
| dc.move.detail | 終わるまでお待ちください。 | DC-move |
| dc.cancel | キャンセル | 多くの状態 |
| dc.err.title | エラー | DC-err |
| dc.err.message | 予期しないエラーが起きました。 | DC-err |
| dc.err.box | {例外の型}: {例外の文} | DC-err |
| dc.err.copy | 内容をコピー | DC-err |
| dc.err.log | ログを開く | DC-err |
| dc.err.close | 閉じる | DC-err |
| e54.title | 保存先の確認 | E54 |
| e54.message | このフォルダは保存先に使えません | E54 |
| e54.box | {理由}（Figma の例「書き込みが拒否されました」） | E54 |
| e54.ok | 別のフォルダを選ぶ | E54 |
| di.e.title | インデックスのインポート | DI-E1〜DI-E3 |
| di.e.close | 閉じる | DI-E1〜DI-E3 |
| di.e1.message | このファイルは tebunko のインデックスではないか、壊れています | DI-E1 |
| di.e1.ok | 別のファイルを選ぶ | DI-E1 |
| di.e2.message | 新しい版の tebunko で作ったインデックスです。tebunko を更新してからインポートしてください | DI-E2 |
| di.e2.ok | リリースノート | DI-E2 |
| di.e3.message | インポートできませんでした。インデックスと設定は元のままです | DI-E3 |
| di.e3.reason | 理由: {理由}（例「理由: 保存先の空きが足りません」） | DI-E3 |

- Figma の例の値:

  | 置き換える所 | 例 |
  |---|---|
  | {インデックス名} | 営業部2025・顧客・アーカイブ・営業部 |
  | {アプリの並び}（{件数} 件） | Excel・Word（2 件） |
  | {アプリ}（{ファイル名}） | Excel（見積書.xlsx） |
  | {フォルダ} | D:\tebunko_index |
  | {既定のフォルダ} | C:\Tools\tebunko\work |
  | {ファイルのパス} | C:\Share\営業部\2024\見積もり\A社_見積書.xlsx |
  | {元のフォルダ} | C:\Share\アーカイブ |
  | {秒} | 15 |
  | {済み} / {全部} | 1,200 / 3,400 |
  | DC-err の詳細の箱 | System.IO.IOException: ファイル「C:\Share\顧客\一覧.xlsx」にアクセスできません。 |

- 句点の無い文（Figma のとおり）:
  - dc.closewait.detail
  - e54.message
  - di.e1.message・di.e2.message・di.e3.message
- 「元の場所:」「理由:」のコロンは半角で、後ろに半角の空白。「（1,200 / 3,400 件）」の括弧は全角、「/」の前後は半角の空白。
- 「…」は 1 字の三点リーダー。

## 7. 操作と振る舞い

| きっかけ | 何が起きるか | 呼ぶ関数 | 次の状態 |
|---|---|---|---|
| 呼び出し側が確認を求める | 暗幕を出し、`newConfirmView` の出力でダイアログを組み、`ShowDialog()` | `newConfirmView` → `showConfirm` | 選んだボタンの Value を返す（閉じた・Esc は `$null`） |
| ボタンをクリック | そのボタンの Value を返して閉じる | — | 呼び出し側へ |
| x・Esc | `$null` を返して閉じる（progress では押せない） | — | 呼び出し側へ |
| DC-err の［内容をコピー］ | 詳細の箱の文をクリップボードに入れる。閉じない | — | DC-err のまま |
| DC-err の［ログを開く］ | ログのフォルダ（またはファイル）を開く。閉じるかは未定 | — | 未定 |
| DI-E2 の［リリースノート］ | 既定のブラウザでリリースノートを開く（URL は未定） | — | 閉じる |
| progress を出す | `showProgress` が `Show()` で出し、更新用の入れ物を返す | `newConfirmView` | 処理の終わりに閉じる |
| DC-close-wait の秒が進む | 1 秒ごとに補足の秒を書き直す。0 で強制終了 | `getCloseWaitDetail` | 窓を閉じる |
| DC-move の件数が進む | 主文とバーを書き直す | `getMoveProgressView` | 終わったら閉じる |

- ボタンのホバー・押した・フォーカスの見た目は、各 Style に従う（Figma に variant が無い）。
- Tab の順: ボタンを左から、最後に x。
- ツールチップ: 無し。

## 8. リサイズ

- 幅は型で固定（520・620・440）。高さは中身に合わせる（`SizeToContent="Height"`）。最小の高さは 200（DI-E* は無し）。
- 主の窓の大きさ（1024×640・1280×820・1600×1000）にかかわらず、同じ大きさで主の窓の中央に出す。見本 408:33891・408:33936 のとおり。
- 暗幕は主の窓全体を覆う。主の窓を動かした・大きさを変えたときは、暗幕も合わせる（`LocationChanged`・`SizeChanged`）。
- 題は 1 行（はみ出したら「…」）。主文・補足・詳細の箱は折り返す。パスは折り返しの中に入る（Figma の DC-src・P-DC-refolder は 2 行に折り返している）。
- ボタンは右寄せのまま。ボタンの文が長くて 1 行に収まらないときの扱いは未定。

## 9. 判断層（`scripts/tebunko/ui/confirm_view.ps1`）

### newConfirmView

- 入力: `[string]$state`（DC-stop など）、`[hashtable]$values`（文言に入れる値。Name・Folder・Count・Apps・App・File・Path・Seconds・Done・Total・Error・Reason）
- 出力: `@{ State; Type; Width [int]; Compact [bool]; Title; Message; Detail; DetailBox; Reason; ShowProgress [bool]; ProgressValue [double or $null]; Buttons }`
  - Buttons の各要素: `@{ Text; Value; Kind = "secondary" | "choice" | "primary" | "danger" | "disabled"; IsDefault [bool]; IsCancel [bool]; Icon }`
  - Width は外枠の幅（520・620・440）。Compact は DI-E* のとき真（4 章の 377 番台の値を使う）。

| state | values | Type | Width | Title | Message | Detail | Buttons（Text:Kind:Value） |
|---|---|---|---|---|---|---|---|
| DC-stop | Name=営業部2025 | normal | 520 | 更新の中止 | 「営業部2025」の更新を中止しますか？ | 取り込んだところまでは残ります。あとで続きから再開できます。 | キャンセル:secondary:$null／中止する:primary:stop |
| DC-kill | Apps=Excel・Word, Count=2 | normal | 520 | Office の終了 | バックグラウンドの Excel・Word（2 件）を終了しますか？ | 画面に表示されていないものだけを終了します。 | キャンセル:secondary:$null／終了する:primary:kill |
| DC-ws | Folder=D:\tebunko_index | normal | 520 | 保存先の変更 | インデックスとログを「D:\tebunko_index」へ移動します。 | 移動中は、検索とインデックスの更新はできません。 | キャンセル:secondary:$null／移動する:primary:change |
| DC-reset | Folder=C:\Tools\tebunko\work | normal | 520 | 既定の場所に戻す | インデックスとログを既定の場所「C:\Tools\tebunko\work」へ移動します。 | 移動中は、検索とインデックスの更新はできません。 | キャンセル:secondary:$null／戻す:primary:reset |
| DC-close | なし | normal | 520 | 更新中です | インデックスを更新中です。中止して閉じますか？ | 取り込んだところまでは残ります。次に起動したときに続きから再開できます。 | 閉じない:secondary:$null／中止して閉じる:primary:close |
| DC-src | Path=C:\Share\営業部\2024\見積もり\A社_見積書.xlsx | normal | 520 | 元のファイルが見つかりません | 元のファイルが見つかりません。フォルダを移動した場合は、移動先のフォルダを選んでください。 | 元の場所: C:\Share\営業部\2024\見積もり\A社_見積書.xlsx | キャンセル:secondary:$null／フォルダを選ぶ:primary:refolder |
| P-DC-refolder | Name=アーカイブ, Folder=C:\Share\アーカイブ | normal | 520 | フォルダの再設定 | 「アーカイブ」のフォルダが見つかりません。フォルダを選び直しますか？ | C:\Share\アーカイブ が移動されたか、名前が変更された可能性があります。選び直したフォルダは、次の更新で取り込みます。 | キャンセル:secondary:$null／フォルダを選ぶ:primary:refolder |
| DC-del | Name=営業部 | danger | 520 | インデックスの削除 | 「営業部」のインデックスを削除しますか？ | 元のファイルは削除されません。 | キャンセル:secondary:$null／削除する:danger:delete |
| DC-kill-vis | App=Excel, File=見積書.xlsx | danger | 520 | Office の終了 | 画面に表示されている Excel（見積書.xlsx）も終了します。保存していない変更は失われます。 | 終了する前に、Excel で保存したかを確認してください。 | キャンセル:secondary:$null／終了する:danger:kill |
| DC-ws-idx | Folder=D:\tebunko_index | choice | 620 | 保存先の変更 | 「D:\tebunko_index」には、すでにインデックスがあります。 | そのフォルダのインデックスを使うか、今のインデックスを移動するかを選んでください。 | キャンセル:secondary:$null／今のインデックスを移動する:choice:move／そのフォルダのインデックスを使う:primary:use |
| DC-stopping | なし | progress | 520 | 更新の中止 | 中止しています… | 取り込み中のファイルが終わるまでお待ちください。 | キャンセル:disabled:$null |
| DC-close-wait | Seconds=15 | progress | 520 | 更新中です | 中止しています… | あと 15 秒で強制終了します | 閉じない:disabled:$null |
| DC-move | Done=1200, Total=3400 | progress | 520 | 保存先の変更 | 移動しています…（1,200 / 3,400 件） | 終わるまでお待ちください。 | キャンセル:disabled:$null |
| DC-err | Error=System.IO.IOException: ファイル「C:\Share\顧客\一覧.xlsx」にアクセスできません。 | error | 520 | エラー | 予期しないエラーが起きました。 | （DetailBox に Error） | 内容をコピー:choice:copy／ログを開く:secondary:log／閉じる:primary:$null |
| E54 | Reason=書き込みが拒否されました | error | 520 | 保存先の確認 | このフォルダは保存先に使えません | （DetailBox に Reason） | キャンセル:secondary:$null／別のフォルダを選ぶ:primary:choose |
| DI-E1 | なし | error | 440 | インデックスのインポート | このファイルは tebunko のインデックスではないか、壊れています | `""` | 閉じる:secondary:$null／別のファイルを選ぶ:primary:choose |
| DI-E2 | なし | error | 440 | インデックスのインポート | 新しい版の tebunko で作ったインデックスです。tebunko を更新してからインポートしてください | `""` | 閉じる:secondary:$null／リリースノート:primary:release |
| DI-E3 | Reason=保存先の空きが足りません | error | 440 | インデックスのインポート | インポートできませんでした。インデックスと設定は元のままです | （Reason に「理由: 保存先の空きが足りません」） | 閉じる:primary:$null |

- Value は呼び出し側が見る値。名前（stop・kill など）は今のコード（`process_tab.ps1` の stop、`newWorkspaceConfirm` の change・use）に合わせ、無いものは新しく付けた。DC-ws-idx の move を今の `reset` に当てるかは「決めること」。
- IsDefault・IsCancel は 5 章の「型ごと」の表から決める。

| state | IsDefault のボタン | IsCancel のボタン |
|---|---|---|
| DC-stop | 中止する | キャンセル |
| DC-del | キャンセル | キャンセル |
| DC-kill-vis | キャンセル | キャンセル |
| DC-ws-idx | そのフォルダのインデックスを使う | キャンセル |
| DC-move | なし | なし |
| DC-err | 閉じる | 閉じる |
| DI-E3 | 閉じる | 閉じる |

- 知らない state を渡したら例外にする。

| state | 期待 |
|---|---|
| DC-unknown | 例外 |

### formatAppList

- 入力: アプリ名の配列。出力: 「・」でつないだ文（同じアプリは 1 回）。

| 入力 | 出力 |
|---|---|
| Excel, Word | Excel・Word |
| Excel, Excel, Word | Excel・Word |
| PowerPoint | PowerPoint |

- 並びの順（Excel・Word・PowerPoint の順か、出てきた順か）は未定。上の表は出てきた順。

### getCloseWaitDetail

| seconds | 出力 |
|---|---|
| 15 | あと 15 秒で強制終了します |
| 1 | あと 1 秒で強制終了します |
| 0 | 未定（0 になったら閉じるので出さない） |

### getMoveProgressView

| done | total | Message | ProgressValue |
|---|---|---|---|
| 1200 | 3400 | 移動しています…（1,200 / 3,400 件） | 35.29…（done / total × 100） |
| 0 | 3400 | 移動しています…（0 / 3,400 件） | 0 |
| 0 | 0 | 未定 | 未定 |

## 10. 画面層（`scripts/shared/ui/shell.ps1`）

- `showScrim($owner)`: 主の窓と同じ位置・大きさの暗幕の窓を `Show()` して返す。主の窓の `LocationChanged`・`SizeChanged` に合わせて動かす。
- `showConfirm($view, $owner)`:
  1. `showScrim` → `loadWindow dialog_confirm.xaml` → Owner を暗幕の窓にする。
  2. `ConfirmWindow.Width = $view.Width + 40`。Compact なら 4 章の 377 番台の値を差し替える。
  3. 題・主文・補足・詳細の箱・理由を写し、空の要素は `Collapsed` にする。
  4. Type が danger・error（Compact でない）なら AlertMark を出す。
  5. Buttons から Button を作って ButtonPanel に足す（Style は Kind から。Value は `Tag`。IsDefault・IsCancel を写す。Icon があれば文字の右に Image を足す）。
  6. `ShowDialog()` → 選んだ Value を返す → 暗幕を閉じる。
- `showProgress($view, $owner)`: 上と同じに組み、`Show()` で出し、`@{ Update = { param($view) ... }; Close = { ... } }` を返す。x と Esc は効かなくする（`Closing` で `Cancel`、処理の側から閉じるときだけ通す）。
- 進み具合の書き直しは、処理のスレッドから `Dispatcher.Invoke` で行う。
- DC-close-wait の秒は `DispatcherTimer`（1 秒）で書き直す。

## 11. 受け入れ

見比べる Figma: 1 章の表のノード全部と、DC-ws 1024×640（408:33891）・DC-ws 1600×1000（408:33936）（`../png/` に画像は無い。get_screenshot で撮って並べる）。

- normal: 幅 520・高さ 200、題のバーが淡い灰、右下に［キャンセル］（白）と主のボタン（青）。
- danger: 主文の左に赤の丸の「!」、主のボタンが赤。
- choice（DC-ws-idx）: 幅 620、ボタンが 3 つ（白・白・青）。
- progress: 主文と補足の間に高さ 6 のバー、ボタンは灰の無効 1 つ。
- error（DC-err・E54）: 赤の「!」、補足の代わりに淡い灰の箱。
- DI-E1〜E3: 幅 440、印なし、高さ 139・156・165。DI-E2 の主のボタンに external-link。DI-E3 は青の［閉じる］1 つ。
- 1024×640 と 1600×1000 で、ダイアログが同じ大きさで主の窓の真ん中に出て、窓全体（タイトルバーも）が黒の 30% で覆われる。
- 文言が 6 章のとおり一字一句同じ（句点の有り無しを含む）。

## 決めること

- 今の確認（見出し 16 SemiBold の `DialogHeading`・事実の行 `FactsPanel`・選択肢の縦並び `ChoicePanel`・`HintText`）を、Figma の題・主文 13 Regular・補足・ボタンだけの形にしてよいか。今これを使っている確認は、すべて 6 章の状態のどれかに置き換える必要がある。当たる状態が無いもの（ワークスペースの「中に workspace を作る」「そのまま使う」など）の扱い。
- 既定のボタンと Esc（Figma に定めが無い。上は今の `showConfirm` の決まりを引き継いだ）。
- 題のバーの x の大きさ（373 番台 10×21、377 番台 20×20）・線の太さ・色。そろえるか。
- DI-E* の主文の色（Ink.Strong）・余白が 373 番台と違う。そろえるか。
- 無効のボタンの文字 #A0A5AB に theme のキーを足すか。`DialogButton.Disabled`・`Progress.Thin` の Style。
- secondary（33）と primary（31）の高さの差をそろえるか。
- DC-stopping・DC-close-wait のバーを不定（動き続ける）にするか。DC-move の総数が 0 のとき。
- DC-close-wait の秒数（例 15）と、0 になったときの強制終了の中身。
- DC-ws-idx の［今のインデックスを移動する］が、今の Value `reset`（選んだフォルダのインデックスを消して移す）に当たるか。
- DC-err の［ログを開く］で、ダイアログを閉じるか。開くのはログのフォルダかファイルか。
- DI-E2 の［リリースノート］の URL。
- E54 の詳細の箱に出す理由（Figma の「書き込みが拒否されました」と、今の `testWorkspaceChoice` の Message）。
- DC-kill の {アプリの並び} の順。
- DC-src は 106 だけで 148 に P 版が無い。P-DC-refolder は 148 だけで 106 に DC 版が無い。両方を実装するか。
- 判断層を `tebunko/` に置き、`shared/` の `showConfirm` は Spec を受け取って描くだけにする分け方でよいか。
