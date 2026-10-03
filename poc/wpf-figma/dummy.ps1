# poc/wpf-figma/dummy.ps1
#
# 26 枚のフレームの一覧と、状態だけが違うフレーム（主に xaml/search.xaml を使う 21 枚）の
# ダミーデータを定義する。show.ps1 / compare.ps1 はここをドットソースして
# Get-FigmaFrames と Set-FigmaFrameState を呼ぶ。
#
# 考え方（計画どおり）: 骨組みは 1 つの xaml/search.xaml を使い回し、
# フレームごとに名前を付けた要素の Visibility・Text・Style などだけを書き換える。
# root は毎回 XamlReader で新しく読み込んだものを渡す（使い回して前の状態が残らないように）。

$script:FigmaFrames = @(
    [pscustomobject]@{ Name = "SP";               Xaml = "xaml\splash.xaml";            Width = 480;  Height = 300 }
    [pscustomobject]@{ Name = "H0";                Xaml = "xaml\search.xaml";            Width = 1280; Height = 820 }
    [pscustomobject]@{ Name = "H-E2";              Xaml = "xaml\search.xaml";            Width = 1280; Height = 820 }
    [pscustomobject]@{ Name = "H-E";               Xaml = "xaml\search.xaml";            Width = 1280; Height = 820 }
    [pscustomobject]@{ Name = "H-S";               Xaml = "xaml\search.xaml";            Width = 1280; Height = 820 }
    [pscustomobject]@{ Name = "E16";                Xaml = "xaml\search.xaml";           Width = 1280; Height = 820 }
    [pscustomobject]@{ Name = "E14";                Xaml = "xaml\search.xaml";           Width = 1280; Height = 820 }
    [pscustomobject]@{ Name = "H-1";               Xaml = "xaml\search.xaml";            Width = 1280; Height = 820 }
    [pscustomobject]@{ Name = "E12";                Xaml = "xaml\search.xaml";           Width = 1280; Height = 820 }
    [pscustomobject]@{ Name = "H-saved";           Xaml = "xaml\search.xaml";            Width = 1280; Height = 820 }
    [pscustomobject]@{ Name = "H-row3";            Xaml = "xaml\search.xaml";            Width = 1280; Height = 820 }
    [pscustomobject]@{ Name = "H-row4";            Xaml = "xaml\search.xaml";            Width = 1280; Height = 820 }
    [pscustomobject]@{ Name = "open-menu";         Xaml = "xaml\search.xaml";            Width = 1280; Height = 820 }
    [pscustomobject]@{ Name = "H";                 Xaml = "xaml\search.xaml";            Width = 1280; Height = 820 }
    [pscustomobject]@{ Name = "H-R";               Xaml = "xaml\search.xaml";            Width = 1280; Height = 820 }
    [pscustomobject]@{ Name = "H-W";               Xaml = "xaml\search.xaml";            Width = 1280; Height = 820 }
    [pscustomobject]@{ Name = "H-B";               Xaml = "xaml\search.xaml";            Width = 1280; Height = 820 }
    [pscustomobject]@{ Name = "H-P";               Xaml = "xaml\search.xaml";            Width = 1280; Height = 820 }
    [pscustomobject]@{ Name = "E13";                Xaml = "xaml\search.xaml";           Width = 1280; Height = 820 }
    [pscustomobject]@{ Name = "H-T1";              Xaml = "xaml\search.xaml";            Width = 1280; Height = 820 }
    [pscustomobject]@{ Name = "H-範囲";             Xaml = "xaml\search.xaml";            Width = 1280; Height = 820 }
    [pscustomobject]@{ Name = "H-範囲2";            Xaml = "xaml\search.xaml";            Width = 1280; Height = 820 }
    [pscustomobject]@{ Name = "H-絞り込み中";        Xaml = "xaml\search.xaml";            Width = 1280; Height = 820 }
    [pscustomobject]@{ Name = "探す範囲";            Xaml = "xaml\探す範囲.xaml";            Width = 817;  Height = 202 }
    [pscustomobject]@{ Name = "フルパスのツールチップ"; Xaml = "xaml\フルパスのツールチップ.xaml"; Width = 571;  Height = 100 }
    [pscustomobject]@{ Name = "正規表現の吹き出し";   Xaml = "xaml\正規表現の吹き出し.xaml";    Width = 300;  Height = 28 }
    [pscustomobject]@{ Name = "高速検索の表示";      Xaml = "xaml\高速検索の表示.xaml";        Width = 758;  Height = 296 }
    # 下の「プレビューの種類ごと見本」は Figma のフレームではなく、メンテナ指摘
    # （2026-10-03・追加の 3）向けに新規に作った参考シート。26 枚には含まれず、参照 PNG も無い。
    [pscustomobject]@{ Name = "プレビューの種類ごと見本"; Xaml = "xaml\プレビューの種類ごと見本.xaml"; Width = 820; Height = 1450 }
)

function Get-FigmaFrames {
    return $script:FigmaFrames
}

function Get-FigmaFrame([string]$Name) {
    return $script:FigmaFrames | Where-Object { $_.Name -eq $Name } | Select-Object -First 1
}

# ---- 欧文フォントをインストールせずに実物へ差し替える ----
#
# theme.xaml の Font.UI（DynamicResource）をあとから書き換えようとすると、
# 画面にすでに結び付いた要素への通知が WPF の検査にひっかかり
# 「file:///...が FontFamily として正しくない」という例外になる（原因不明の挙動。
# 直接プロパティへ代入する分には問題が起きない）。そのため、ここでは
# ツリーの各要素に直接 FontFamily を上書きする（ローカル値はスタイルの設定より優先される）。
function Set-RealFont($Root, [System.Windows.Media.FontFamily]$FontFamily) {
    $stack = New-Object System.Collections.Generic.Stack[object]
    $stack.Push($Root)
    while ($stack.Count -gt 0) {
        $node = $stack.Pop()
        if ($node -isnot [System.Windows.DependencyObject]) { continue }
        try {
            $node.SetValue([System.Windows.Documents.TextElement]::FontFamilyProperty, $FontFamily)
        } catch {
            # FontFamily を持たない要素型は無視する。
        }
        $children = [System.Windows.LogicalTreeHelper]::GetChildren($node)
        foreach ($child in $children) {
            if ($child -is [System.Windows.DependencyObject]) { $stack.Push($child) }
        }
    }
}

# ---- 要素を操作する小さな道具（見つからない名前は黙って無視する。見本 4 種には無い名前もあるため） ----

function Find-Named($Root, [string]$Name) {
    return $Root.FindName($Name)
}

function Set-ElVisible($Root, [string]$Name, [bool]$Visible) {
    $el = Find-Named $Root $Name
    if ($null -ne $el) {
        $el.Visibility = if ($Visible) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    }
}

function Set-ElOpen($Root, [string]$Name, [bool]$Open) {
    # Popup は Visibility ではなく IsOpen で出し入れする（diff_round3.md 3 ★）。
    # IsOpen は PresentationSource（実際の窓）が無いと WPF 側で false に戻される
    # （ヘッドレスの show/compare/screenshots では実物の窓を開かないため）。
    # show.ps1（実物の窓）では IsOpen がそのまま効くので両方とも設定し、
    # ヘッドレスの合成（Merge-PopupOverlay）は Tag を目印に見る。
    $el = Find-Named $Root $Name
    if ($null -ne $el) {
        $el.IsOpen = $Open
        $el.Tag = if ($Open) { "open" } else { $null }
    }
}

function Set-ElMinHeight($Root, [string]$Name, $MinHeight) {
    $el = Find-Named $Root $Name
    if ($null -ne $el) { $el.MinHeight = $MinHeight }
}

# 要素の Grid.Row を書き換える（E16 の帯を、検索バーの上の行から件数の行の位置へ動かす。
# diff_round3.md 4 ★）。
function Set-ElGridRow($Root, [string]$Name, [int]$Row) {
    $el = Find-Named $Root $Name
    if ($null -ne $el) { [System.Windows.Controls.Grid]::SetRow($el, $Row) }
}

# 結果の無い状態（H0・H-E・H-E2・H-S・E14・E16）向けに、プレビューの欄・境目・
# 列の見出しの行をまとめて隠す（diff_round3.md 4）。
function Hide-ResultsColumnHeader($Root) {
    Set-ElVisible $Root "ResultsColumnHeader" $false
    Set-ElHeight $Root "ResultsColumnHeaderRow" 0
}

function Hide-PreviewPane($Root) {
    Set-ElHeight $Root "PreviewDividerRow" 0
    Set-ElVisible $Root "PreviewDivider" $false
    Set-ElVisible $Root "PreviewToolbar" $false
    Set-ElVisible $Root "PreviewContentHost" $false
    Set-ElMinHeight $Root "PreviewContentRow" 0
    Set-ElHeight $Root "PreviewContentRow" 0
}

function Set-ElText($Root, [string]$Name, [string]$Text) {
    $el = Find-Named $Root $Name
    if ($null -ne $el) { $el.Text = $Text }
}

function Set-ElChecked($Root, [string]$Name, $Checked) {
    # $Checked は $true / $false / $null（中間状態）
    $el = Find-Named $Root $Name
    if ($null -ne $el) { $el.IsChecked = $Checked }
}

function Set-ElEnabled($Root, [string]$Name, [bool]$Enabled) {
    $el = Find-Named $Root $Name
    if ($null -ne $el) { $el.IsEnabled = $Enabled }
}

function Set-ElHeight($Root, [string]$Name, $Height) {
    $el = Find-Named $Root $Name
    if ($null -ne $el) { $el.Height = $Height }
}

function Set-ElStyleKey($Root, [string]$Name, [string]$StyleKey) {
    $el = Find-Named $Root $Name
    if ($null -ne $el) { $el.Style = $Root.FindResource($StyleKey) }
}

function Set-ElBrushKey($Root, [string]$Name, [string]$Property, [string]$BrushKey) {
    $el = Find-Named $Root $Name
    if ($null -ne $el) { $el.$Property = $Root.FindResource($BrushKey) }
}

# Banner・ContentBanner のバッジの中の線（Path の Data）を、お知らせの種類に合わせて差し替える
# （既定は Icon.BadgeGlyphInfo の「i」。メンテナ指摘 2026-10-03・動作確認で の i・!・×・✓ に対応）。
function Set-ElGeometryKey($Root, [string]$Name, [string]$GeometryKey) {
    $el = Find-Named $Root $Name
    if ($null -ne $el) { $el.Data = $Root.FindResource($GeometryKey) }
}

# 検索結果の行の文字列（ヒット語だけ黄色の背景）を差し替える。
# $Parts は @(@{ Text = "見積先："}, @{ Text = "(株)山田商事"; Hit = $true }, @{ Text = "（御中）" }) の形。
function Set-ElRuns($Root, [string]$Name, [array]$Parts) {
    $el = Find-Named $Root $Name
    if ($null -eq $el) { return }
    $el.Inlines.Clear()
    foreach ($part in $Parts) {
        $run = New-Object System.Windows.Documents.Run($part.Text)
        if ($part.Hit) { $run.Background = $Root.FindResource("Hit.FFF176") }
        if ($part.Bold) { $run.FontWeight = "Bold" }
        $el.Inlines.Add($run)
    }
}

# ---- 高速検索の表示（6 状態） ----
# $State: "unavailable-connect" / "unavailable-regex" / "unavailable-short" / "unavailable-pending" / "partial" / "ok" / "hidden"
function Set-FastSearchState($Root, [string]$State) {
    $badge = Find-Named $Root "FastSearchBadge"
    $icon = Find-Named $Root "FastSearchIcon"
    $text = Find-Named $Root "FastSearchText"
    $info = Find-Named $Root "FastSearchInfo"
    if ($null -eq $badge) { return }

    if ($State -eq "hidden") {
        $badge.Visibility = [System.Windows.Visibility]::Collapsed
        return
    }
    $badge.Visibility = [System.Windows.Visibility]::Visible

    # unavailable-* はボーダーレス（枠なし）・くすんだ薄い色ではなく濃い灰色の文字にする（diff_round2.md 2 回目指摘）
    switch ($State) {
        "unavailable-connect" {
            $badge.Background = $Root.FindResource("Bg.F3F3F4"); $badge.BorderBrush = [System.Windows.Media.Brushes]::Transparent
            $icon.Stroke = $Root.FindResource("Ink.5F6368"); $text.Foreground = $Root.FindResource("Ink.5F6368")
            $text.Text = "高速検索：使用不可（Windows Search に接続できません）"
            $info.Visibility = [System.Windows.Visibility]::Visible; $info.Stroke = $Root.FindResource("Ink.5F6368")
        }
        "unavailable-regex" {
            $badge.Background = $Root.FindResource("Bg.F3F3F4"); $badge.BorderBrush = [System.Windows.Media.Brushes]::Transparent
            $icon.Stroke = $Root.FindResource("Ink.5F6368"); $text.Foreground = $Root.FindResource("Ink.5F6368")
            $text.Text = "高速検索：使用不可（正規表現では使えません）"
            $info.Visibility = [System.Windows.Visibility]::Visible; $info.Stroke = $Root.FindResource("Ink.5F6368")
        }
        "unavailable-short" {
            $badge.Background = $Root.FindResource("Bg.F3F3F4"); $badge.BorderBrush = [System.Windows.Media.Brushes]::Transparent
            $icon.Stroke = $Root.FindResource("Ink.5F6368"); $text.Foreground = $Root.FindResource("Ink.5F6368")
            $text.Text = "高速検索：使用不可（2文字以上で使えます）"
            $info.Visibility = [System.Windows.Visibility]::Visible; $info.Stroke = $Root.FindResource("Ink.5F6368")
        }
        "unavailable-pending" {
            $badge.Background = $Root.FindResource("Bg.F3F3F4"); $badge.BorderBrush = [System.Windows.Media.Brushes]::Transparent
            $icon.Stroke = $Root.FindResource("Ink.5F6368"); $text.Foreground = $Root.FindResource("Ink.5F6368")
            $text.Text = "高速検索：使用不可（反映待ちです）"
            $info.Visibility = [System.Windows.Visibility]::Visible; $info.Stroke = $Root.FindResource("Ink.5F6368")
        }
        "partial" {
            $badge.Background = $Root.FindResource("Warn.FFF5E0"); $badge.BorderBrush = $Root.FindResource("Warn.BA7D00")
            $icon.Stroke = $Root.FindResource("Warn.BA7D00"); $text.Foreground = $Root.FindResource("Warn.BA7D00")
            $text.Text = "高速検索：一部で使用可"
            $info.Visibility = [System.Windows.Visibility]::Visible; $info.Stroke = $Root.FindResource("Warn.BA7D00")
        }
        "ok" {
            $badge.Background = $Root.FindResource("Ok.E0F7E0"); $badge.BorderBrush = $Root.FindResource("Ok.218A21")
            $icon.Stroke = $Root.FindResource("Ok.218A21"); $text.Foreground = $Root.FindResource("Ok.218A21")
            $text.Text = "高速検索：使用可"
            $info.Visibility = [System.Windows.Visibility]::Visible; $info.Stroke = $Root.FindResource("Ok.218A21")
        }
    }
}

# ---- フレームごとの違い ----
# $Root は xaml/search.xaml を XamlReader.Load した直後（既定値＝H のベースライン）の状態を渡す。
function Set-FigmaFrameState($Root, [string]$FrameName) {
    switch ($FrameName) {

        "H" {
            # ベースライン。xaml の既定値がそのまま H（既定の 14 件・1 行目選択・使用可）になっている。
            # 参照画像は「開く ▾」のドロップダウンが開いた状態のため、合わせる。
            Set-ElOpen $Root "ContextMenuPopup" $true
        }

        "H0" {
            # インデックス未作成（検索対象なし）。
            Set-ElText $Root "SearchTargetCountText" ""
            Set-ElVisible $Root "SearchTargetLinksRow" $false
            Set-ElVisible $Root "SearchTargetSearchBoxBorder" $false
            Set-ElVisible $Root "TreePanel" $false
            Set-ElVisible $Root "EmptyTreeState" $true

            Set-ElText $Root "SearchWordBox" ""
            # 検索欄が空のとき薄い文字を出し、ラベルは灰色にする（diff_round3.md 8）
            Set-ElVisible $Root "SearchWordPlaceholder" $true
            Set-ElBrushKey $Root "SearchWordLabel" "Foreground" "Ink.5F6368"
            Set-ElEnabled $Root "SearchButton" $false
            Set-FastSearchState $Root "hidden"

            Set-ElVisible $Root "ResultsPanel" $false
            Set-ElVisible $Root "ResultsEmptyState" $true
            $title = Find-Named $Root "ResultsEmptyTitle"
            if ($title) { $title.Text = "インデックスが作成されていません"; $title.FontSize = 18; $title.FontWeight = "Bold" }
            # 2 行で、Figma と同じ位置で折る（diff_round3.md 8）
            Set-ElText $Root "ResultsEmptyBody" "検索を行うには、まずインデックス管理からフォルダを登録し、`r`nインデックスを作成してください。"
            $cta = Find-Named $Root "ResultsEmptyCta"
            if ($cta) { $cta.Content = "インデックス管理を開く"; $cta.Width = 176 }
            Set-ElVisible $Root "ExpandCollapseLinksRow" $false
            Set-ElVisible $Root "SaveButton" $false
            Set-ElText $Root "ResultsSummaryText" ""

            Set-ElVisible $Root "PreviewOpenBody" $false
            Set-ElVisible $Root "PreviewOpenArrow" $false
            Set-ElVisible $Root "PreviewOpenFolderLink" $false
            Set-ElText $Root "PreviewBreadcrumbText" ""
            Set-ElVisible $Root "PreviewContentHost" $false
            Set-ElText $Root "StatusBarText" "インデックスが未作成です"

            # 結果の無い状態。プレビューの欄・境目・絞り込みの欄・列の見出しの行を出さない（diff_round3.md 4 ★）
            Hide-PreviewPane $Root
            Hide-ResultsColumnHeader $Root
            Set-ElVisible $Root "ResultsFilterRow" $false
        }

        "H-E" {
            # 検索前（検索ワードが空、検索ボタンは不可）。高速検索バッジは「使用可」を出す（diff_round4.md 6）。
            Set-ElText $Root "SearchWordBox" ""
            Set-ElEnabled $Root "SearchButton" $false
            Set-FastSearchState $Root "ok"
            Set-ElVisible $Root "ResultsPanel" $false
            # 検索ワードが無いときは、主要エリアに何も出さない（diff_round3.md 4）
            Set-ElVisible $Root "ResultsEmptyState" $false
            Set-ElVisible $Root "ExpandCollapseLinksRow" $false
            Set-ElVisible $Root "SaveButton" $false
            Set-ElText $Root "ResultsSummaryText" ""
            Set-ElVisible $Root "PreviewOpenBody" $false
            Set-ElVisible $Root "PreviewOpenArrow" $false
            Set-ElVisible $Root "PreviewOpenFolderLink" $false
            Set-ElText $Root "PreviewBreadcrumbText" ""
            Set-ElVisible $Root "PreviewContentHost" $false
            Set-ElText $Root "StatusBarText" ""

            Hide-PreviewPane $Root
            Hide-ResultsColumnHeader $Root
            Set-ElVisible $Root "ResultsFilterRow" $false
        }

        "H-E2" {
            # 検索ワードは入っているが、まだ検索していない（検索ボタンは押せる）。
            Set-ElText $Root "SearchWordBox" "(株)山田商事"
            Set-ElEnabled $Root "SearchButton" $true
            Set-ElVisible $Root "ResultsPanel" $false
            Set-ElVisible $Root "ResultsEmptyState" $true
            Set-ElVisible $Root "ResultsEmptyIcon" $false
            Set-ElText $Root "ResultsEmptyTitle" "検索ワードを入力してください"
            Set-ElText $Root "ResultsEmptyBody" "［検索］を押すと、検索を始めます。"
            Set-ElVisible $Root "ResultsEmptyCta" $false
            Set-ElVisible $Root "ExpandCollapseLinksRow" $false
            Set-ElVisible $Root "SaveButton" $false
            Set-ElText $Root "ResultsSummaryText" ""
            Set-ElVisible $Root "PreviewOpenBody" $false
            Set-ElVisible $Root "PreviewOpenArrow" $false
            Set-ElVisible $Root "PreviewOpenFolderLink" $false
            Set-ElText $Root "PreviewBreadcrumbText" ""
            Set-ElVisible $Root "PreviewContentHost" $false
            Set-ElText $Root "StatusBarText" ""

            Hide-PreviewPane $Root
            Hide-ResultsColumnHeader $Root
            Set-ElVisible $Root "ResultsFilterRow" $false
        }

        "H-S" {
            # 検索中（中止できる）。まだ 5 件しか見つかっていないため、
            # 先頭の A社_見積書.xlsx のグループだけを出し、ほかのファイルの行は出さない。
            Set-ElText $Root "ResultsSummaryText" "検索中… 該当 5 件"
            # 中止ボタンは赤ではなく青（Btn.Primary のまま。diff_round2.md 2 回目指摘）
            $btn = Find-Named $Root "SearchButton"
            if ($btn) { $btn.Content = "中止" }
            Set-ElVisible $Root "ExpandCollapseLinksRow" $false
            Set-ElVisible $Root "SaveButton" $false
            Set-ElVisible $Root "FileListRow1" $false
            Set-ElVisible $Root "FileListRow2" $false
            Set-ElVisible $Root "FileListRow3" $false
            Set-ElVisible $Root "FileListRow4" $false
            Set-ElText $Root "StatusBarText" ""

            Hide-PreviewPane $Root
        }

        "E16" {
            # 中止した（見つかった分だけ表示。帯は検索バーの上ではなく、件数行の位置に出す
            # ★ diff_round3.md 4。ほかの帯フレームは検索バーの上の行のまま）。
            Set-ElGridRow $Root "Banner" 4
            Set-ElText $Root "BannerText" "検索を中止しました（見つかった 5 件を表示しています）"
            Set-ElVisible $Root "BannerButton" $false
            Set-ElVisible $Root "ResultsSummaryText" $false
            Set-ElVisible $Root "ExpandCollapseLinksRow" $false
            Set-ElVisible $Root "SaveButton" $false
            Set-ElVisible $Root "FileListRow1" $false
            Set-ElVisible $Root "FileListRow2" $false
            Set-ElVisible $Root "FileListRow3" $false
            Set-ElVisible $Root "FileListRow4" $false
            # ステータスバーは空にする（diff_round4.md 6・E16）
            Set-ElText $Root "StatusBarText" ""

            Hide-PreviewPane $Root
            Set-ElVisible $Root "ResultsFilterRow" $false
        }

        "E14" {
            # 見つからなかった。
            Set-ElText $Root "SearchWordBox" "(株)山田商店"
            Set-ElVisible $Root "ResultsPanel" $false
            Set-ElVisible $Root "ResultsEmptyState" $true
            Set-ElVisible $Root "ResultsEmptyIcon" $false
            $title = Find-Named $Root "ResultsEmptyTitle"
            if ($title) { $title.Text = "見つかりませんでした"; $title.FontSize = 16; $title.FontWeight = "Bold" }
            Set-ElText $Root "ResultsEmptyBody" "検索対象のフォルダ・［探す範囲］・［正規表現］を見直してください。"
            Set-ElVisible $Root "ResultsEmptyNote" $true
            Set-ElVisible $Root "ResultsEmptyCta" $false
            Set-ElChecked $Root "RegexCheck" $true
            Set-FastSearchState $Root "unavailable-regex"
            Set-ElVisible $Root "ExpandCollapseLinksRow" $false
            Set-ElVisible $Root "SaveButton" $false
            # 件数の行には「0件」を出さない。状態バーには出す（diff_round3.md 4）
            Set-ElVisible $Root "ResultsSummaryText" $false
            Set-ElVisible $Root "PreviewOpenBody" $false
            Set-ElVisible $Root "PreviewOpenArrow" $false
            Set-ElVisible $Root "PreviewOpenFolderLink" $false
            Set-ElText $Root "PreviewBreadcrumbText" ""
            Set-ElVisible $Root "PreviewContentHost" $false
            Set-ElText $Root "StatusBarText" "検索しました：(株)山田商店 0 件"

            Hide-PreviewPane $Root
            Hide-ResultsColumnHeader $Root
            Set-ElVisible $Root "ResultsFilterRow" $false
        }

        "H-1" {
            # 検索ワードが 1 文字（まだ足りない）。
            Set-ElText $Root "SearchWordBox" "山"
            Set-FastSearchState $Root "unavailable-short"
            Set-ElVisible $Root "ResultsPanel" $false
            Set-ElVisible $Root "ResultsEmptyState" $false
            Set-ElVisible $Root "ExpandCollapseLinksRow" $false
            Set-ElVisible $Root "SaveButton" $false
            Set-ElText $Root "ResultsSummaryText" ""
            Set-ElVisible $Root "PreviewOpenBody" $false
            Set-ElVisible $Root "PreviewOpenArrow" $false
            Set-ElVisible $Root "PreviewOpenFolderLink" $false
            Set-ElText $Root "PreviewBreadcrumbText" ""
            Set-ElVisible $Root "PreviewContentHost" $false
            Set-ElText $Root "StatusBarText" ""

            Hide-ResultsColumnHeader $Root
        }

        "E12" {
            # 正規表現の構文エラー（直前の結果は残したまま、赤枠＋赤文字で知らせる）。
            Set-ElText $Root "SearchWordBox" "(株)山田(商事"
            Set-ElChecked $Root "RegexCheck" $true
            Set-ElStyleKey $Root "SearchWordBox" "TextBox.Error"
            Set-ElVisible $Root "SearchWordError" $true
            Set-FastSearchState $Root "unavailable-regex"
            Set-ElEnabled $Root "SearchButton" $false
            Set-ElText $Root "ResultsSummaryText" "14件（5ファイル）・0.8秒・通常の検索"
        }

        "H-saved" {
            # 保存に成功した（緑の帯）。
            Set-ElHeight $Root "TopBannerRow" "Auto"
            Set-ElBrushKey $Root "Banner" "Background" "Ok.E0F7E0"
            Set-ElBrushKey $Root "Banner" "BorderBrush" "Ok.218A21"
            Set-ElBrushKey $Root "BannerAccent" "Fill" "Ok.218A21"
            Set-ElBrushKey $Root "BannerIcon" "Stroke" "Ok.218A21"
            Set-ElBrushKey $Root "BannerGlyph" "Stroke" "Ok.218A21"
            Set-ElGeometryKey $Root "BannerGlyph" "Icon.BadgeGlyphOk"
            Set-ElText $Root "BannerText" "検索結果を保存しました"
            $bb = Find-Named $Root "BannerButton"
            if ($bb) { $bb.Content = "フォルダを開く"; $bb.BorderBrush = $Root.FindResource("Ok.218A21"); $bb.Foreground = $Root.FindResource("Ok.218A21") }
        }

        "H-row3" {
            # ヒットした行が検索結果の 3 件目（セル A41）。選んでいる行は TopRow3（diff_round3.md 1）。
            # プレビューは前後の行（40・42行目）を含めて表示し、ヒット行（41行目）だけ列 A の内容を出す
            # （列 B・C は空）。
            Set-ElBrushKey $Root "TopRow1" "Background" "Bg.FAFBFC"
            Set-ElBrushKey $Root "TopRow3" "Background" "Select.E1F2FF"
            Set-ElText $Root "PreviewBreadcrumbText" "営業部\A社_見積書.xlsx ・ [シート]見積書!A41 ・ セル"
            Set-ElText $Root "PreviewTopRowNum" "39"
            Set-ElText $Root "PreviewTopText1" ""
            Set-ElText $Root "PreviewTopText2" ""
            Set-ElText $Root "PreviewMidRowNum" "40"
            Set-ElText $Root "PreviewMidText1" ""
            Set-ElText $Root "PreviewMidText2" ""
            Set-ElText $Root "PreviewHighlightRowNum" "41"
            Set-ElText $Root "PreviewHighlightText1" "納品場所：(株)山田商事"
            Set-ElText $Root "PreviewHighlightText2" "本社ビル"
            Set-ElText $Root "PreviewBottomRowNum" "42"
            Set-ElText $Root "PreviewBottomText1" ""
            Set-ElText $Root "PreviewBottomText2" ""
            Set-ElText $Root "PreviewExtraRowNum" "43"
        }

        "H-row4" {
            # ヒットした行が検索結果の 4 件目（セル D5・図形）。選んでいる行は TopRow4（diff_round3.md 1）。
            # プレビューは前後の行（4・6行目）を含めて表示し、ヒット行（5行目）は D 列を光らせる
            # （列 A・B・C は空。diff_round3.md 10）。
            Set-ElBrushKey $Root "TopRow1" "Background" "Bg.FAFBFC"
            Set-ElBrushKey $Root "TopRow4" "Background" "Select.E1F2FF"
            Set-ElText $Root "PreviewBreadcrumbText" "営業部\A社_見積書.xlsx ・ [シート]見積書!D5 ・ 図形"
            # 同じシートの 3・4 行目は、既定（H）の 3・4 行目と同じ中身（diff_round5.md 1）
            Set-ElText $Root "PreviewTopRowNum" "3"
            Set-ElText $Root "PreviewTopText1" "見積先：(株)山田商事(御中)"
            Set-ElText $Root "PreviewTopText2" "見積番号：12-2324-0143"
            Set-ElText $Root "PreviewMidRowNum" "4"
            Set-ElText $Root "PreviewMidText1" "件名：基幹システム導入一式"
            Set-ElText $Root "PreviewMidText2" "発行日：2024/10/15"
            Set-ElText $Root "PreviewHighlightRowNum" "5"
            Set-ElText $Root "PreviewHighlightText1" ""
            Set-ElText $Root "PreviewHighlightText2" ""
            Set-ElVisible $Root "PreviewHighlightCellA" $false
            Set-ElVisible $Root "PreviewHighlightCellD" $true
            Set-ElText $Root "PreviewHighlightTextD" "納品場所：(株)山田商事 本社 4F"
            Set-ElText $Root "PreviewBottomRowNum" "6"
            Set-ElText $Root "PreviewBottomText1" ""
            Set-ElText $Root "PreviewBottomText2" ""
            Set-ElText $Root "PreviewExtraRowNum" "7"
        }

        "open-menu" {
            # Figma の書き出しが、右クリックメニューだけを残した白紙のキャンバスのため、
            # 画面の飾り（タイトルバー・本体）をすべて隠し、メニューだけを出す。
            Set-ElVisible $Root "TitleBarGrid" $false
            Set-ElVisible $Root "BodyGrid" $false
            Set-ElOpen $Root "ContextMenuPopup" $true
        }

        "H-R" {
            # 正規表現で検索した状態。ヒット件数が増え（議事録メモ.md が加わる）、
            # 高速検索は正規表現では使えないため「使用不可」になる。
            Set-ElChecked $Root "RegexCheck" $true
            Set-FastSearchState $Root "unavailable-regex"
            Set-ElText $Root "ResultsSummaryText" "16 件（6 ファイル）・0.8 秒・通常の検索"
            # H-R のステータスバー（diff_round3.md 4 は Figma の読み違いだった。diff_round6.md 5 で訂正）
            Set-ElText $Root "StatusBarText" "検索しました：(株)山田商事 16 件"

            Set-ElVisible $Root "MinutesGroupHeader" $true
            Set-ElVisible $Root "MinutesSubHeader" $true
            Set-ElVisible $Root "MinutesRow1" $true
            Set-ElVisible $Root "MinutesRow2" $true
        }

        "H-W" {
            # Word ファイル（基本契約書.docx）が一覧の先頭に展開された状態。
            # 先頭グループが Excel から Word に替わるため、見出し・列名・3 行の中身・
            # ファイル一覧の並び（基本契約書.docx が一覧から抜け、代わりに A社_見積書.xlsx が入る）
            # をすべて書き換える。
            $icon = Find-Named $Root "TopGroupIcon"
            if ($icon) { $icon.Data = $Root.FindResource("Icon.FileText"); $icon.Stroke = $Root.FindResource("Word.185ABD") }
            Set-ElText $Root "TopGroupFileName" "基本契約書.docx"
            Set-ElText $Root "TopGroupLocation" "　顧客/契約書"
            Set-ElText $Root "TopGroupCount" "1 ページ（目安） ほか 2 か所 ・ 3 件"

            Set-ElText $Root "TopRow1Col1" "1 ページ（目安）"
            Set-ElText $Root "TopRow1Col2" "本文"
            Set-ElRuns $Root "TopRow1Col3" @(
                @{ Text = "甲：" }, @{ Text = "(株)山田商事"; Hit = $true }, @{ Text = "（以下「甲」という）" }
            )
            Set-ElText $Root "TopRow2Col1" "2 ページ（目安）"
            Set-ElText $Root "TopRow2Col2" "本文"
            Set-ElRuns $Root "TopRow2Col3" @(
                @{ Text = "第3条" }, @{ Text = "(株)山田商事"; Hit = $true }, @{ Text = "は毎月末日までに支払う" }
            )
            Set-ElText $Root "TopRow3Col1" "3 ページ（目安）"
            Set-ElText $Root "TopRow3Col2" "本文"
            Set-ElRuns $Root "TopRow3Col3" @(
                @{ Text = "署名欄：" }, @{ Text = "(株)山田商事"; Hit = $true }, @{ Text = " 代表取締役 山田 太郎" }
            )
            Set-ElVisible $Root "TopRow4" $false

            # 閉じた Excel は chart-column（diff_round6.md の訂正 a059c9d で file-spreadsheet から戻した）
            $fIcon = Find-Named $Root "FileListItem3Icon"
            if ($fIcon) { $fIcon.Data = $Root.FindResource("Icon.ChartColumn"); $fIcon.Stroke = $Root.FindResource("Excel.107C41") }
            Set-ElText $Root "FileListItem3Name" "A社_見積書.xlsx"
            Set-ElText $Root "FileListItem3Count" "[シート]見積書 ほか 1 か所 ・ 5 件"

            # 1 件しか選べない Word の段落プレビュー。「開く」は分割せず単独ボタン。
            Set-ElVisible $Root "PreviewOpenArrow" $false
            Set-ElStyleKey $Root "PreviewOpenBody" "Btn.OpenPlain"
            Set-ElVisible $Root "PreviewSheetGrid" $false
            Set-ElVisible $Root "PreviewParagraphView" $true
            Set-ElText $Root "PreviewParagraphHeading" "第3条（支払条件）"
            # Figma の 2 行（diff_round5.md 4）。間にあった「1 甲は乙に対し…」の行は無く、
            # 「2 振込手数料は甲の負担とする。」の行がある
            Set-ElRuns $Root "PreviewParagraphText" @(
                @{ Text = "(株)山田商事"; Bold = $true }, @{ Text = "（以下「甲」という）は、本契約に基づく代金を、毎月末日までに乙の指定する口座に振り込む。" }
            )
            Set-ElText $Root "PreviewParagraphText2" "2 振込手数料は甲の負担とする。"
            Set-ElText $Root "PreviewBreadcrumbText" "顧客\契約書\基本契約書.docx ・ 1 ページ（目安） ・ 本文"
        }

        "H-B" {
            # インデックス更新中（進み具合）。更新中は高速検索が使えないため、バッジを隠し、
            # 件数の末尾の「・高速検索」も外す（diff_round2.md 2 回目指摘）。
            Set-ElHeight $Root "TopBannerRow" "Auto"
            Set-ElText $Root "BannerText" "インデックスを更新しています（営業部 1,200 / 2,075 件）"
            Set-ElVisible $Root "NavBadge_Index" $true
            Set-ElVisible $Root "NavBadge_IndexDot" $true
            Set-ElText $Root "NavBadge_IndexText" "58%"
            # 「58%」は pill にする（diff_round3.md 12・小）
            Set-ElBrushKey $Root "NavBadge_IndexPill" "Background" "Select.E5F1FB"
            Set-ElBrushKey $Root "NavBadge_IndexPill" "BorderBrush" "Accent.0078D4"
            Set-FastSearchState $Root "hidden"
            Set-ElText $Root "ResultsSummaryText" "14 件（5 ファイル）・0.8 秒"
        }

        "H-P" {
            # インデックス更新が中断している。
            Set-ElHeight $Root "TopBannerRow" "Auto"
            $icon = Find-Named $Root "BannerIcon"
            if ($icon) { $icon.Stroke = $Root.FindResource("Warn.BA7D00") }
            $glyph = Find-Named $Root "BannerGlyph"
            if ($glyph) { $glyph.Stroke = $Root.FindResource("Warn.BA7D00") }
            $banner = Find-Named $Root "Banner"
            if ($banner) { $banner.Background = $Root.FindResource("Warn.FFF5E0"); $banner.BorderBrush = $Root.FindResource("Warn.BA7D00") }
            Set-ElGeometryKey $Root "BannerGlyph" "Icon.BadgeGlyphWarn"
            # バナーの文言とボタンは「前回の更新が途中です」「続きから再開」にする（diff_round2.md 2 回目指摘）
            Set-ElText $Root "BannerText" "前回の更新が途中です（残り 875 件）"
            # H-P のボタンは橙にする（diff_round3.md 7）
            $bb = Find-Named $Root "BannerButton"
            if ($bb) { $bb.Content = "続きから再開"; $bb.BorderBrush = $Root.FindResource("Warn.BA7D00"); $bb.Foreground = $Root.FindResource("Warn.BA7D00"); $bb.Height = 24 }
            Set-ElVisible $Root "NavBadge_Index" $true
            Set-ElVisible $Root "NavBadge_IndexDot" $false
            Set-ElText $Root "NavBadge_IndexText" "中断"
            # H-P は高速検索のバッジを出さない。ナビの「中断」は枠の無い橙の pill にする
            # （diff_round3.md 4、diff_round6.md 6・小で枠を外した）
            Set-FastSearchState $Root "hidden"
            Set-ElBrushKey $Root "NavBadge_IndexPill" "Background" "Warn.FFF5E0"
            $pill = Find-Named $Root "NavBadge_IndexPill"
            if ($pill) { $pill.BorderBrush = [System.Windows.Media.Brushes]::Transparent }
            Set-ElBrushKey $Root "NavBadge_IndexText" "Foreground" "Warn.BA7D00"
            # 件数行は「・高速検索」を付けない（diff_round4.md 6）
            Set-ElText $Root "ResultsSummaryText" "14 件（5 ファイル）・0.8 秒"
        }

        "E13" {
            # 検索対象 0/4（ツリーはすべて外す。案内の一行を出す）。
            # フォルダを 1 つも選んでいないため、検索の実行に関わる操作は不可にする。
            Set-ElText $Root "SearchTargetCountText" "0 / 4"
            foreach ($n in @("TreeCheck_営業部","TreeCheck_A社","TreeCheck_B社","TreeCheck_提案書",
                             "TreeCheck_顧客","TreeCheck_取引先台帳","TreeCheck_契約書",
                             "TreeCheck_営業部2025","TreeCheck_見積もり","TreeCheck_請求書",
                             "TreeCheck_アーカイブ")) {
                Set-ElChecked $Root $n $false
            }
            # 3 つの点（顧客＝橙・営業部2025＝青・アーカイブ＝赤）はすべて出す（diff_round3.md 9）
            Set-ElHeight $Root "NavNoticeRow" "Auto"
            Set-ElVisible $Root "NavNotice" $true

            # 「探す範囲」は検索対象 0 件でも消さない・無効にしない（diff_round2.md 2 回目指摘）
            Set-ElEnabled $Root "SearchButton" $false
            Set-ElEnabled $Root "ResultsFilterBox" $false
            Set-ElEnabled $Root "CaseCheck" $false
            Set-ElEnabled $Root "RegexCheck" $false
        }

        "H-T1" {
            # 顧客フォルダが中間状態（契約書だけ外した）。
            Set-ElChecked $Root "TreeCheck_契約書" $false
            Set-ElChecked $Root "TreeCheck_顧客" $null
            Set-ElText $Root "SearchTargetCountText" "2 / 4"
        }

        "H-範囲" {
            Set-ElVisible $Root "RangePopup" $true
        }

        "H-範囲2" {
            Set-ElVisible $Root "RangePopup" $true
            Set-ElChecked $Root "RangeCheck_Comment" $false
            Set-ElChecked $Root "RangeCheck_Note" $false
            Set-ElText $Root "RangeButtonText" "探す範囲：本文・図形"
        }

        "H-絞り込み中" {
            # メンテナ指摘（2026-10-03・追加の 1）: 絞り込み中は件数を「N件中K件を表示」にする見本。
            $filterBox = Find-Named $Root "ResultsFilterBox"
            if ($filterBox) { $filterBox.Text = "山田"; $filterBox.Foreground = $Root.FindResource("Ink.202124") }
            Set-ElText $Root "ResultsSummaryText" "14 件中 3 件を表示"
            # 絞り込み中は、合うものだけを並べる（diff_round3.md 10）。
            # 「山田」に合う A社_見積書.xlsx の 3 行だけ残し、ほかのファイルの行は隠す。
            Set-ElVisible $Root "FileListRow1" $false
            Set-ElVisible $Root "FileListRow2" $false
            Set-ElVisible $Root "FileListRow3" $false
            Set-ElVisible $Root "FileListRow4" $false
            Set-ElVisible $Root "TopRow4" $false
            Set-ElText $Root "TopGroupCount" "[シート]見積書 ・ 3 件"
        }

        default {
            # 見本 4 種（splash・探す範囲・フルパスのツールチップ・正規表現の吹き出し・高速検索の表示）は
            # それぞれの xaml に固定の見本を書いてあるため、差し替えは不要。
        }
    }
}
