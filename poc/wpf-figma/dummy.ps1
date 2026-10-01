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
    [pscustomobject]@{ Name = "探す範囲";            Xaml = "xaml\探す範囲.xaml";            Width = 817;  Height = 202 }
    [pscustomobject]@{ Name = "フルパスのツールチップ"; Xaml = "xaml\フルパスのツールチップ.xaml"; Width = 571;  Height = 100 }
    [pscustomobject]@{ Name = "正規表現の吹き出し";   Xaml = "xaml\正規表現の吹き出し.xaml";    Width = 300;  Height = 28 }
    [pscustomobject]@{ Name = "高速検索の表示";      Xaml = "xaml\高速検索の表示.xaml";        Width = 758;  Height = 296 }
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

    switch ($State) {
        "unavailable-connect" {
            $badge.Background = $Root.FindResource("Bg.F3F3F4"); $badge.BorderBrush = $Root.FindResource("Disabled.9AA0A6")
            $icon.Fill = $Root.FindResource("Disabled.9AA0A6"); $text.Foreground = $Root.FindResource("Disabled.8C949E")
            $text.Text = "高速検索：使用不可（Windows Search に接続できません）"
            $info.Visibility = [System.Windows.Visibility]::Visible; $info.Stroke = $Root.FindResource("Disabled.9AA0A6")
        }
        "unavailable-regex" {
            $badge.Background = $Root.FindResource("Bg.F3F3F4"); $badge.BorderBrush = $Root.FindResource("Disabled.9AA0A6")
            $icon.Fill = $Root.FindResource("Disabled.9AA0A6"); $text.Foreground = $Root.FindResource("Disabled.8C949E")
            $text.Text = "高速検索：使用不可（正規表現では使えません）"
            $info.Visibility = [System.Windows.Visibility]::Visible; $info.Stroke = $Root.FindResource("Disabled.9AA0A6")
        }
        "unavailable-short" {
            $badge.Background = $Root.FindResource("Bg.F3F3F4"); $badge.BorderBrush = $Root.FindResource("Disabled.9AA0A6")
            $icon.Fill = $Root.FindResource("Disabled.9AA0A6"); $text.Foreground = $Root.FindResource("Disabled.8C949E")
            $text.Text = "高速検索：使用不可（2文字以上で使えます）"
            $info.Visibility = [System.Windows.Visibility]::Collapsed
        }
        "unavailable-pending" {
            $badge.Background = $Root.FindResource("Bg.F3F3F4"); $badge.BorderBrush = $Root.FindResource("Disabled.9AA0A6")
            $icon.Fill = $Root.FindResource("Disabled.9AA0A6"); $text.Foreground = $Root.FindResource("Disabled.8C949E")
            $text.Text = "高速検索：使用不可（反映待ちです）"
            $info.Visibility = [System.Windows.Visibility]::Visible; $info.Stroke = $Root.FindResource("Disabled.9AA0A6")
        }
        "partial" {
            $badge.Background = $Root.FindResource("Warn.FFF5E0"); $badge.BorderBrush = $Root.FindResource("Warn.BA7D00")
            $icon.Fill = $Root.FindResource("Warn.BA7D00"); $text.Foreground = $Root.FindResource("Warn.BA7D00")
            $text.Text = "高速検索：一部で使用可"
            $info.Visibility = [System.Windows.Visibility]::Visible; $info.Stroke = $Root.FindResource("Warn.BA7D00")
        }
        "ok" {
            $badge.Background = $Root.FindResource("Ok.E0F7E0"); $badge.BorderBrush = $Root.FindResource("Ok.218A21")
            $icon.Fill = $Root.FindResource("Ok.218A21"); $text.Foreground = $Root.FindResource("Ok.218A21")
            $text.Text = "高速検索：使用可"
            $info.Visibility = [System.Windows.Visibility]::Collapsed
        }
    }
}

# ---- フレームごとの違い ----
# $Root は xaml/search.xaml を XamlReader.Load した直後（既定値＝H のベースライン）の状態を渡す。
function Set-FigmaFrameState($Root, [string]$FrameName) {
    switch ($FrameName) {

        "H" {
            # ベースライン。xaml の既定値がそのまま H（既定の 14 件・1 行目選択・使用可）になっている。
        }

        "H0" {
            # インデックス未作成（検索対象なし）。
            Set-ElText $Root "SearchTargetCountText" ""
            Set-ElVisible $Root "SearchTargetLinksRow" $false
            Set-ElVisible $Root "SearchTargetSearchBoxBorder" $false
            Set-ElVisible $Root "TreePanel" $false
            Set-ElVisible $Root "EmptyTreeState" $true

            Set-ElText $Root "SearchWordBox" ""
            $swb = Find-Named $Root "SearchWordBox"
            if ($swb) { $swb.Tag = "検索ワードを入力" }
            Set-ElEnabled $Root "SearchButton" $false
            Set-FastSearchState $Root "hidden"

            Set-ElVisible $Root "ResultsPanel" $false
            Set-ElVisible $Root "ResultsEmptyState" $true
            $title = Find-Named $Root "ResultsEmptyTitle"
            if ($title) { $title.Text = "インデックスが作成されていません"; $title.FontSize = 18; $title.FontWeight = "Bold" }
            Set-ElText $Root "ResultsEmptyBody" "検索するフォルダを選び、インデックスを作成してください。"
            $cta = Find-Named $Root "ResultsEmptyCta"
            if ($cta) { $cta.Content = "インデックス管理を開く" }
            Set-ElVisible $Root "ExpandAllLink" $false
            Set-ElVisible $Root "CollapseAllLink" $false
            Set-ElVisible $Root "SaveButton" $false
            Set-ElText $Root "ResultsSummaryText" ""

            Set-ElVisible $Root "PreviewOpenBody" $false
            Set-ElVisible $Root "PreviewOpenArrow" $false
            Set-ElVisible $Root "PreviewOpenFolderLink" $false
            Set-ElText $Root "PreviewBreadcrumbText" ""
            Set-ElVisible $Root "PreviewContentHost" $false
            Set-ElText $Root "StatusBarText" "インデックスが未作成です"
        }

        "H-E" {
            # 検索前（検索ワードが空、検索ボタンは不可）。
            Set-ElText $Root "SearchWordBox" ""
            Set-ElEnabled $Root "SearchButton" $false
            Set-FastSearchState $Root "hidden"
            Set-ElVisible $Root "ResultsPanel" $false
            Set-ElVisible $Root "ResultsEmptyState" $true
            Set-ElText $Root "ResultsEmptyTitle" "検索ワードを入力してください"
            Set-ElText $Root "ResultsEmptyBody" "検索ワードを入力し、［検索］を押してください。"
            Set-ElVisible $Root "ResultsEmptyCta" $false
            Set-ElVisible $Root "ExpandAllLink" $false
            Set-ElVisible $Root "CollapseAllLink" $false
            Set-ElVisible $Root "SaveButton" $false
            Set-ElText $Root "ResultsSummaryText" ""
            Set-ElVisible $Root "PreviewOpenBody" $false
            Set-ElVisible $Root "PreviewOpenArrow" $false
            Set-ElVisible $Root "PreviewOpenFolderLink" $false
            Set-ElText $Root "PreviewBreadcrumbText" ""
            Set-ElVisible $Root "PreviewContentHost" $false
            Set-ElText $Root "StatusBarText" ""
        }

        "H-E2" {
            # 検索ワードは入っているが、まだ検索していない（検索ボタンは押せる）。
            Set-ElText $Root "SearchWordBox" "（株）山田商事"
            Set-ElEnabled $Root "SearchButton" $true
            Set-ElVisible $Root "ResultsPanel" $false
            Set-ElVisible $Root "ResultsEmptyState" $true
            Set-ElText $Root "ResultsEmptyTitle" "検索ワードを入力してください"
            Set-ElText $Root "ResultsEmptyBody" "［検索］を押すと、検索を始めます。"
            Set-ElVisible $Root "ResultsEmptyCta" $false
            Set-ElVisible $Root "ExpandAllLink" $false
            Set-ElVisible $Root "CollapseAllLink" $false
            Set-ElVisible $Root "SaveButton" $false
            Set-ElText $Root "ResultsSummaryText" ""
            Set-ElVisible $Root "PreviewOpenBody" $false
            Set-ElVisible $Root "PreviewOpenArrow" $false
            Set-ElVisible $Root "PreviewOpenFolderLink" $false
            Set-ElText $Root "PreviewBreadcrumbText" ""
            Set-ElVisible $Root "PreviewContentHost" $false
            Set-ElText $Root "StatusBarText" ""
        }

        "H-S" {
            # 検索中（中止できる）。
            Set-ElText $Root "ResultsSummaryText" "検索中… 該当 5 件"
            $btn = Find-Named $Root "SearchButton"
            if ($btn) { $btn.Content = "中止"; $btn.Style = $Root.FindResource("Btn.Danger") }
            Set-ElVisible $Root "ExpandAllLink" $false
            Set-ElVisible $Root "CollapseAllLink" $false
            Set-ElVisible $Root "SaveButton" $false
        }

        "E16" {
            # 中止した（見つかった分だけ表示。青の帯で知らせる）。
            Set-ElHeight $Root "ContentBannerRow" "Auto"
            Set-ElVisible $Root "ContentBanner" $true
            Set-ElText $Root "ContentBannerText" "検索を中止しました（見つかった 5 件を表示しています）"
            Set-ElVisible $Root "ResultsSummaryText" $false
            Set-ElVisible $Root "ExpandAllLink" $false
            Set-ElVisible $Root "CollapseAllLink" $false
            Set-ElVisible $Root "SaveButton" $false
        }

        "E14" {
            # 見つからなかった。
            Set-ElVisible $Root "ResultsPanel" $false
            Set-ElVisible $Root "ResultsEmptyState" $true
            $title = Find-Named $Root "ResultsEmptyTitle"
            if ($title) { $title.Text = "見つかりませんでした"; $title.FontSize = 16; $title.FontWeight = "Bold" }
            Set-ElText $Root "ResultsEmptyBody" "検索対象のフォルダ・［探す範囲］・［正規表現］を見直してください。"
            Set-ElVisible $Root "ResultsEmptyNote" $true
            Set-ElVisible $Root "ResultsEmptyCta" $false
            Set-ElChecked $Root "RegexCheck" $true
            Set-ElVisible $Root "ExpandAllLink" $false
            Set-ElVisible $Root "CollapseAllLink" $false
            Set-ElVisible $Root "SaveButton" $false
            Set-ElText $Root "ResultsSummaryText" ""
            Set-ElVisible $Root "PreviewOpenBody" $false
            Set-ElVisible $Root "PreviewOpenArrow" $false
            Set-ElVisible $Root "PreviewOpenFolderLink" $false
            Set-ElText $Root "PreviewBreadcrumbText" ""
            Set-ElVisible $Root "PreviewContentHost" $false
        }

        "H-1" {
            # 検索ワードが 1 文字（まだ足りない）。
            Set-ElText $Root "SearchWordBox" "山"
            Set-FastSearchState $Root "unavailable-short"
            Set-ElVisible $Root "ResultsPanel" $false
            Set-ElVisible $Root "ResultsEmptyState" $false
            Set-ElVisible $Root "ExpandAllLink" $false
            Set-ElVisible $Root "CollapseAllLink" $false
            Set-ElVisible $Root "SaveButton" $false
            Set-ElText $Root "ResultsSummaryText" ""
            Set-ElVisible $Root "PreviewOpenBody" $false
            Set-ElVisible $Root "PreviewOpenArrow" $false
            Set-ElVisible $Root "PreviewOpenFolderLink" $false
            Set-ElText $Root "PreviewBreadcrumbText" ""
            Set-ElVisible $Root "PreviewContentHost" $false
            Set-ElText $Root "StatusBarText" ""
        }

        "E12" {
            # 正規表現の構文エラー（直前の結果は残したまま、赤枠＋赤文字で知らせる）。
            Set-ElChecked $Root "RegexCheck" $true
            Set-ElStyleKey $Root "SearchWordBox" "TextBox.Error"
            Set-ElVisible $Root "SearchWordError" $true
            Set-FastSearchState $Root "unavailable-regex"
        }

        "H-saved" {
            # 保存に成功した（緑の帯）。
            Set-ElHeight $Root "ContentBannerRow" "Auto"
            Set-ElVisible $Root "ContentBanner" $true
            Set-ElBrushKey $Root "ContentBanner" "Background" "Ok.E0F7E0"
            Set-ElBrushKey $Root "ContentBanner" "BorderBrush" "Ok.218A21"
            Set-ElBrushKey $Root "ContentBannerIcon" "Stroke" "Ok.218A21"
            Set-ElText $Root "ContentBannerText" "✓ 検索結果を保存しました"
            Set-ElBrushKey $Root "ContentBannerText" "Foreground" "Ok.218A21"
            Set-ElVisible $Root "ContentBannerButton" $true
        }

        "H-row3" {
            Set-ElChecked $Root "ResultRow1Check" $false
            Set-ElChecked $Root "ResultRow41Check" $true
            Set-ElText $Root "PreviewBreadcrumbText" "A社_見積書.xlsx > [シート] 見積書 > セル A41"
            Set-ElText $Root "PreviewHighlightText1" "納品場所：（株）山田商事 本社ビル"
            Set-ElText $Root "PreviewHighlightText2" ""
        }

        "H-row4" {
            Set-ElChecked $Root "ResultRow1Check" $false
            Set-ElChecked $Root "ResultRow5Check" $true
            Set-ElText $Root "PreviewBreadcrumbText" "A社_見積書.xlsx > [シート] 見積書 > セル A5"
            Set-ElText $Root "PreviewHighlightText1" "83 納品場所：（株）山田商事 本社4F"
            Set-ElText $Root "PreviewHighlightText2" ""
        }

        "open-menu" {
            # Figma の書き出しが、右クリックメニューだけを残した白紙のキャンバスのため、
            # 画面の飾り（タイトルバー・本体）をすべて隠し、メニューだけを出す。
            Set-ElVisible $Root "TitleBarGrid" $false
            Set-ElVisible $Root "BodyGrid" $false
            Set-ElVisible $Root "ContextMenuPopup" $true
        }

        "H-R" {
            # 検索中に別の画面（インデックス管理）へ移っても続く……という場面向けの予備。
            # 参照画像は H とほぼ同じ構図のため、既定のまま。
        }

        "H-W" {
            # 1 件しか選べない Word の段落プレビュー。「開く」は分割せず単独ボタン。
            Set-ElVisible $Root "PreviewOpenArrow" $false
            Set-ElStyleKey $Root "PreviewOpenBody" "Btn.OpenPlain"
            Set-ElVisible $Root "PreviewSheetGrid" $false
            Set-ElVisible $Root "PreviewParagraphView" $true
            Set-ElText $Root "PreviewParagraphText" "第3条（支払条件）`r`n`r`n甲は乙に対し、本契約に基づく対価を、検収完了日の属する月の翌月末日までに、乙が指定する銀行口座へ振り込む方法により支払う。"
            Set-ElText $Root "PreviewBreadcrumbText" "基本契約書.docx > p.1 段落4"
        }

        "H-B" {
            # インデックス更新中（進み具合）。
            Set-ElHeight $Root "BannerRow" "Auto"
            Set-ElText $Root "BannerText" "インデックスを更新しています（営業部 1,200 / 2,075 件）"
            Set-ElVisible $Root "NavBadge_Index" $true
            Set-ElVisible $Root "NavBadge_IndexDot" $true
            Set-ElText $Root "NavBadge_IndexText" "58%"
        }

        "H-P" {
            # インデックス更新が中断している。
            Set-ElHeight $Root "BannerRow" "Auto"
            $icon = Find-Named $Root "BannerIcon"
            if ($icon) { $icon.Stroke = $Root.FindResource("Warn.BA7D00") }
            $banner = Find-Named $Root "Banner"
            if ($banner) { $banner.Background = $Root.FindResource("Warn.FFF5E0"); $banner.BorderBrush = $Root.FindResource("Warn.BA7D00") }
            Set-ElText $Root "BannerText" "インデックスの更新が中断しました"
            $bb = Find-Named $Root "BannerButton"
            if ($bb) { $bb.Content = "続ける" }
            Set-ElVisible $Root "NavBadge_Index" $true
            Set-ElVisible $Root "NavBadge_IndexDot" $false
            Set-ElText $Root "NavBadge_IndexText" "中断"
        }

        "E13" {
            # 検索対象 0/4（ツリーはすべて外す。案内の一行を出す）。
            Set-ElText $Root "SearchTargetCountText" "0 / 4"
            foreach ($n in @("TreeCheck_営業部","TreeCheck_A社","TreeCheck_B社","TreeCheck_提案書",
                             "TreeCheck_顧客","TreeCheck_取引先台帳","TreeCheck_契約書",
                             "TreeCheck_営業部2025","TreeCheck_見積もり","TreeCheck_請求書",
                             "TreeCheck_アーカイブ")) {
                Set-ElChecked $Root $n $false
            }
            Set-ElVisible $Root "Dot_顧客" $false
            Set-ElVisible $Root "Dot_営業部2025" $false
            Set-ElVisible $Root "Dot_アーカイブ" $false
            Set-ElHeight $Root "NavNoticeRow" "Auto"
            Set-ElVisible $Root "NavNotice" $true
        }

        "H-T1" {
            # 顧客フォルダが中間状態（契約書だけ外した）。
            Set-ElChecked $Root "TreeCheck_契約書" $false
            Set-ElChecked $Root "TreeCheck_顧客" $null
        }

        "H-範囲" {
            Set-ElVisible $Root "RangePopup" $true
        }

        "H-範囲2" {
            Set-ElVisible $Root "RangePopup" $true
            Set-ElChecked $Root "RangeCheck_Comment" $false
            Set-ElChecked $Root "RangeCheck_Note" $false
            Set-ElText $Root "RangeButtonText" "探す範囲：本文・図形 ▾"
        }

        default {
            # 見本 4 種（splash・探す範囲・フルパスのツールチップ・正規表現の吹き出し・高速検索の表示）は
            # それぞれの xaml に固定の見本を書いてあるため、差し替えは不要。
        }
    }
}
