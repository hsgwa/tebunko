# スクリプトの構成（パス定義・構文・XAML・型の読み込み）のテスト
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
}

Describe "パス定義" -Tag Meta {
    It "リポジトリ直下を基準にする（書き込めるため、設定ファイルもリポジトリ直下に置く）" {
        $rootDir | Should -Be (Resolve-Path "$here\..").Path
        $settingsFile | Should -Be "$rootDir\setting.config"
    }

    It "work の中身は work の置き場所（既定はリポジトリ直下の work。setting.config の workspaceFolder で変わる）を基準にする" {
        $workspace.Dir | Should -Be (getWorkDir $settingsFile)
        $workspace.IndexDir | Should -Be "$($workspace.Dir)\store\content"
        $workspace.LegacyIndexDir | Should -Be "$($workspace.Dir)\index"
        $workspace.TmpRoot | Should -Be "$($workspace.Dir)\tmp"
        $workspace.PublishDir | Should -Be "$($workspace.Dir)\publish\$PID"
        $workspace.ResultFile | Should -Be "$($workspace.Dir)\search_results.txt"
    }
}

Describe "画面の部品でのパスの組み立て" -Tag Meta {
    # ui\ 配下のファイルは gui.ps1 から dot-source する部品。中で $PSScriptRoot を使うと ui\ を指すため、
    # "${PSScriptRoot}\tebunko\indexer.ps1" のように起動口からの相対パスを書くと存在しないパスになる
    # （［すべて更新］でインデクサが起動しなかった不具合）。パスは起動口（gui.ps1）で決めて変数で渡す
    It "ui 配下のスクリプトで `$PSScriptRoot を使っていない" {
        # コードで使っているかだけを見る（コメントで説明に触れているだけの行は対象外）
        $found = @(Get-ChildItem "$here\..\scripts" -Recurse -Filter "*.ps1" |
            Where-Object { $_.DirectoryName -match '\\ui$' } |
            Select-String -Pattern '\$\{?PSScriptRoot\}?' |
            Where-Object { $_.Line.TrimStart() -notmatch '^#' } |
            ForEach-Object { "$($_.Filename):$($_.LineNumber)" })
        ($found -join ", ") | Should -Be ""
    }
}

Describe "カバレッジの計測の対象（画面層だけを分母から外す）" -Tag Meta {
    BeforeAll {
        . "$PSScriptRoot\..\helpers\coverage_targets.ps1"
        $script:targets = @(getCoverageTargets "$here\..\scripts" | ForEach-Object { $_.Substring(([System.IO.Path]::GetFullPath("$here\..\scripts")).Length + 1) })
    }

    It "テストのある ui\ の下のファイル（<file>）は計測の対象に入っている" -ForEach @(
        @{ file = "tebunko\ui\result_list.ps1" }
        @{ file = "tebunko\ui\preview.ps1" }
        @{ file = "tebunko\ui\index_tree.ps1" }
        @{ file = "tebunko\ui\open_source.ps1" }
        @{ file = "tebunko\ui\search\result_types.ps1" }
        @{ file = "tebunko\ui\search\target_types.ps1" }
        @{ file = "tebunko\ui\index\index_types.ps1" }
        @{ file = "tebunko\ui\startup_error.ps1" }
        @{ file = "shared\ui\types.ps1" }
        @{ file = "tebunko\ui\index\index_list_view.ps1" }
        @{ file = "tebunko\ui\index\index_edit_view.ps1" }
        @{ file = "tebunko\ui\index\index_detail_view.ps1" }
        @{ file = "tebunko\ui\index\index_archive_view.ps1" }
        @{ file = "tebunko\ui\wait_view.ps1" }
        @{ file = "tebunko\ui\shell\nav_view.ps1" }
        @{ file = "tebunko\ui\shell\nav.ps1" }
        @{ file = "tebunko\ui\shell\status_bar.ps1" }
        @{ file = "tebunko\ui\shell\status_bar_view.ps1" }
        @{ file = "tebunko\ui\shell\notice_banner_view.ps1" }
        @{ file = "tebunko\ui\search\search_bar_view.ps1" }
        @{ file = "tebunko\ui\search\result_list_view.ps1" }
        @{ file = "tebunko\ui\search\open_source_view.ps1" }
    ) {
        $script:targets | Should -Contain $file
    }

    It "ui\ の下の判断層（*_view.ps1）は、すべて計測の対象に入っている" {
        $root = [System.IO.Path]::GetFullPath("$here\..\scripts")
        $views = @(Get-ChildItem -LiteralPath $root -Recurse -Filter "*_view.ps1" |
            Where-Object { $_.FullName -like "*\ui\*" } |
            ForEach-Object { $_.FullName.Substring($root.Length + 1) })
        $views.Count | Should -BeGreaterThan 0
        foreach ($view in $views) {
            $script:targets | Should -Contain $view
        }
    }

    It "画面層（<file>）は計測の対象に入っていない" -ForEach @(
        @{ file = "tebunko\ui\gui_main.ps1" }
        @{ file = "tebunko\ui\index\index_list.ps1" }
        @{ file = "tebunko\ui\index\index_store.ps1" }
        @{ file = "tebunko\ui\index\index_archive.ps1" }
        @{ file = "tebunko\ui\index\index_detail.ps1" }
        @{ file = "tebunko\ui\index\index_events.ps1" }
        @{ file = "tebunko\ui\settings\settings.ps1" }
        @{ file = "tebunko\ui\about_dialog.ps1" }
        @{ file = "tebunko\ui\leftover_dialog.ps1" }
        @{ file = "shared\ui\shell.ps1" }
        @{ file = "shared\ui\app_host.ps1" }
        @{ file = "shared\ui\trimmed_tooltip.ps1" }
        @{ file = "tebunko\ui\splash.ps1" }
        @{ file = "tebunko\ui\shell\notice_banner.ps1" }
        @{ file = "tebunko\ui\shell\regions.ps1" }
        @{ file = "tebunko\ui\shell\screens.ps1" }
        @{ file = "tebunko\ui\shell\services.ps1" }
        @{ file = "tebunko\ui\shell\settings_recovery.ps1" }
        @{ file = "tebunko\ui\shell\window_events.ps1" }
        @{ file = "tebunko\ui\shell\closing.ps1" }
        @{ file = "tebunko\ui\shell\startup.ps1" }
        @{ file = "tebunko\ui\shell\activate_watch.ps1" }
        @{ file = "shared\ui\folder_dialog.ps1" }
        @{ file = "tebunko\ui\search\search_bar.ps1" }
        @{ file = "tebunko\ui\search\search_session.ps1" }
        @{ file = "tebunko\ui\search\result_filter.ps1" }
    ) {
        $script:targets | Should -Not -Contain $file
    }

    It "起動口 gui.ps1 は入っていない。状態層（core）は入っている" {
        @($script:targets | Where-Object { (Split-Path $_ -Leaf) -eq "gui.ps1" }).Count | Should -Be 0
        @($script:targets | Where-Object { $_.Contains("\core\") }).Count | Should -BeGreaterThan 0
    }
}

Describe "画面の既定のフォント（起動した窓の Font.Body）" -Tag Meta {
    # フォントは Windows 標準のものだけを使う（Segoe UI・Yu Gothic UI・Meiryo UI）。同梱しないので、読み込みの呼び出しはパスだけを渡す
    It "画面を読み込む呼び出しは、パスだけを渡す（フォントの置き場所を渡さない）" {
        $problems = New-Object System.Collections.Generic.List[string]
        $calls = 0
        foreach ($file in (Get-ChildItem "$here\..\scripts" -Recurse -Filter "*.ps1")) {
            if ($file.Name -eq "app_host.ps1") { continue }
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$null)
            $found = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -in @("loadWindow", "loadXaml") }, $true)
            foreach ($call in $found) {
                $calls++
                if ($call.CommandElements.Count -gt 2) { $problems.Add("$($file.Name):$($call.Extent.StartLineNumber)") }
            }
        }
        $calls | Should -BeGreaterThan 5
        $problems -join ", " | Should -Be ""
    }
}

Describe "同梱する第三者の部品" -Tag Meta {
    It "同梱するのはアイコンの形のライセンスの文面だけで、フォントのファイルは無い" {
        Test-Path -LiteralPath "$here\..\scripts\shared\licenses\LICENSE-Lucide.txt" | Should -Be $true
        @(Get-ChildItem "$here\..\scripts" -Recurse -File | Where-Object { $_.Extension -in ".ttf", ".otf", ".ttc" -or $_.Name -eq "OFL.txt" }).Count | Should -Be 0
    }

    It "ライセンス文の置き場（scripts/shared/licenses/）は、new_sbom.ps1 の第三者の部品（`$thirdParty の name）と 1 対 1 になっている" {
        # new_sbom.ps1 はドットソースすると本体が動くので、構文木で読む。
        # 左辺が $thirdParty の代入の右辺の配列の、直下の要素（[ordered] を外したハッシュテーブル）の "name" だけを拾う
        $parsed = [System.Management.Automation.Language.Parser]::ParseFile("$here\..\tools\new_sbom.ps1", [ref]$null, [ref]$null)
        $assign = @($parsed.FindAll({ param($n)
                    $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
                    $n.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
                    $n.Left.VariablePath.UserPath -ceq "thirdParty" }, $true))[0]
        $assign | Should -Not -BeNullOrEmpty
        $array = $assign.Right.Find({ param($n) $n -is [System.Management.Automation.Language.ArrayExpressionAst] }, $false)
        $names = New-Object System.Collections.Generic.List[string]
        $body = $array.SubExpression.Statements[0].PipelineElements[0].Expression
        $items = if ($body -is [System.Management.Automation.Language.ArrayLiteralAst]) { $body.Elements } else { @($body) }
        foreach ($item in $items) {
            $table = if ($item -is [System.Management.Automation.Language.ConvertExpressionAst]) { $item.Child } else { $item }
            $table | Should -BeOfType [System.Management.Automation.Language.HashtableAst]
            foreach ($pair in $table.KeyValuePairs) {
                if ($pair.Item1 -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $pair.Item1.Value -ceq "name" -and
                    $pair.Item2.PipelineElements[0].Expression -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                    $names.Add($pair.Item2.PipelineElements[0].Expression.Value)
                }
            }
        }
        # 拾えた数が 0 なら、読み方が壊れている（見逃さない）
        $names.Count | Should -BeGreaterThan 0
        # name が文字列の定数でない要素があると、黙って飛ばさずに落とす（要素の数と同じだけ拾えている）
        $names.Count | Should -Be @($items).Count

        $files = @(Get-ChildItem -LiteralPath "$here\..\scripts\shared\licenses" -File)
        @($files | Where-Object { $_.Name -cnotmatch '^LICENSE-.+\.txt$' }).Count | Should -Be 0
        $fileNames = @($files | ForEach-Object { $_.Name -creplace '^LICENSE-(.+)\.txt$', '$1' } | Sort-Object)
        $fileNames -join "," | Should -BeExactly (@($names | Sort-Object) -join ",")
    }
}

Describe "画面の既定のフォント（起動画面）" -Tag Meta {
    It "起動画面（splash.xaml）が直に書くフォントの並びは、newAppFontFamily（Font.Body）と同じ" {
        $splash = [System.IO.File]::ReadAllText("$here\..\scripts\tebunko\xaml\splash.xaml")
        $host_ = [System.IO.File]::ReadAllText("$here\..\scripts\shared\ui\app_host.ps1")
        $splash -match '<Window[^>]*?\sFontFamily="([^"]+)"' | Should -BeTrue
        $splashFont = $Matches[1]
        $host_ -match '(?s)function newAppFontFamily \{.*?-ArgumentList "([^"]+)"' | Should -BeTrue
        $splashFont | Should -Be $Matches[1]
    }
}

Describe "スクリプトの構文" -Tag Meta {
    BeforeDiscovery {
        $scriptFiles = @(Get-ChildItem "$PSScriptRoot\..\..\scripts" -Recurse -Filter "*.ps1" |
            ForEach-Object { @{ Name = $_.Name; FullName = $_.FullName } })
    }

    It "<name> に構文エラーが無い" -ForEach $scriptFiles {
        $errors = $null
        [System.Management.Automation.Language.Parser]::ParseFile($FullName, [ref]$null, [ref]$errors) | Out-Null
        # 継承元の型が別ファイルにある場合、1 ファイルだけを読むと型が見つからない（TypeNotFound）。
        # 読み込む順で解決できることは、下の「型の読み込み」で実際に読み込んで確かめる
        @($errors | Where-Object { $_.ErrorId -ne "TypeNotFound" }).Count | Should -Be 0
    }

    It "ファイル名・関数・変数・ハッシュテーブルのキー・メンバーの名前・呼ぶ名前・並列の照合に渡す名前に pack を使わない（本文インデックスのコードの名前は contentIndex）" {
        # コメントや package の URL を拾わないよう、正規表現で本文を探さずに AST から名前を拾う
        $found = New-Object System.Collections.Generic.List[string]
        foreach ($file in Get-ChildItem "$PSScriptRoot\..\..\scripts" -Recurse -Filter "*.ps1") {
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$null)
            $names = New-Object System.Collections.Generic.List[object]
            foreach ($node in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
                $names.Add(@($node.Name, $node.Extent.StartLineNumber))
            }
            $names.Add(@($file.Name, 1))
            foreach ($node in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.MemberExpressionAst] -and $n.Member -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true)) {
                $names.Add(@($node.Member.Value, $node.Extent.StartLineNumber))
            }
            foreach ($node in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.PropertyMemberAst] -or $n -is [System.Management.Automation.Language.FunctionMemberAst] }, $true)) {
                $names.Add(@($node.Name, $node.Extent.StartLineNumber))
            }
            foreach ($node in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)) {
                $names.Add(@($node.VariablePath.UserPath, $node.Extent.StartLineNumber))
            }
            foreach ($node in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.HashtableAst] }, $true)) {
                foreach ($pair in $node.KeyValuePairs) {
                    if ($pair.Item1 -is [System.Management.Automation.Language.StringConstantExpressionAst]) { $names.Add(@($pair.Item1.Value, $pair.Item1.Extent.StartLineNumber)) }
                }
            }
            foreach ($node in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)) {
                $commandName = $node.GetCommandName()
                if ($commandName) { $names.Add(@($commandName, $node.Extent.StartLineNumber)) }
                if ($commandName -eq "newWorkerState") {
                    foreach ($text in $node.FindAll({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true)) {
                        $names.Add(@($text.Value, $text.Extent.StartLineNumber))
                    }
                }
            }
            foreach ($entry in $names) {
                if ($entry[0] -match "(?i)pack(?!age)") { $found.Add("$($file.Name):$($entry[1]) $($entry[0])") }
            }
        }
        if ($found.Count) {
            throw "名前に pack を使っています。本文インデックスのコードの名前は contentIndex です（docs/design/index.md の用語の表）: $(($found | Select-Object -Unique) -join ' / ')"
        }
    }
}

Describe "画面定義（XAML）" -Tag Meta {
    BeforeDiscovery {
        $xamlFiles = @(Get-ChildItem "$PSScriptRoot\..\..\scripts" -Recurse -Filter "*.xaml" |
            ForEach-Object { @{ Name = $_.Name; FullName = $_.FullName } })
    }

    It "<name> が XML として読める" -ForEach $xamlFiles {
        { [xml](Get-Content $FullName -Raw -Encoding UTF8) } | Should -Not -Throw
    }

    It "行を分ける仕切りは、共通の Style（Splitter.Rows）を使い、自分のテンプレートを持たない" {
        $scripts = (Resolve-Path "$PSScriptRoot\..\..\scripts").Path
        $theme = [xml](Get-Content "$scripts\shared\xaml\theme.xaml" -Raw -Encoding UTF8)
        $x = "http://schemas.microsoft.com/winfx/2006/xaml"
        $keys = @($theme.DocumentElement.ChildNodes | Where-Object { $_ -is [System.Xml.XmlElement] } | ForEach-Object { $_.GetAttribute("Key", $x) })
        $keys | Should -Contain "Splitter.Rows"

        $splitters = @()
        foreach ($f in Get-ChildItem "$scripts\tebunko\xaml" -Recurse -Filter "*.xaml") {
            $xml = [xml](Get-Content $f.FullName -Raw -Encoding UTF8)
            $splitters += @($xml.SelectNodes("//*[local-name()='GridSplitter'][@ResizeDirection='Rows']"))
        }
        $splitters.Count | Should -BeGreaterThan 0
        foreach ($s in $splitters) {
            $s.GetAttribute("Style") | Should -Be "{StaticResource Splitter.Rows}"
            $s.SelectNodes("*[local-name()='GridSplitter.Template']").Count | Should -Be 0
        }
    }

    It "設定の節は、カードの Style（Section.Card・Section.Header・Section.Row）で作る" {
        $scripts = (Resolve-Path "$PSScriptRoot\..\..\scripts").Path
        $theme = [xml](Get-Content "$scripts\shared\xaml\theme.xaml" -Raw -Encoding UTF8)
        $x = "http://schemas.microsoft.com/winfx/2006/xaml"
        $keys = @($theme.DocumentElement.ChildNodes | Where-Object { $_ -is [System.Xml.XmlElement] } | ForEach-Object { $_.GetAttribute("Key", $x) })
        foreach ($k in "Section.Card", "Section.Header", "Section.Row", "Field.ReadOnly") { $keys | Should -Contain $k }
        $keys | Should -Not -Contain "Bg.Section"

        $settingsText = Get-Content "$scripts\tebunko\xaml\settings\settings.xaml" -Raw -Encoding UTF8
        $left = @(Get-ChildItem $scripts -Recurse -Include *.xaml, *.ps1 | Where-Object { (Get-Content $_.FullName -Raw -Encoding UTF8) -match "Bg\.Section" } | ForEach-Object Name)
        $left | Should -BeNullOrEmpty -Because "使われなくなった Bg.Section を scripts の下に残さない"
        $settingsText | Should -Match 'Style="\{StaticResource Section\.Card\}"'
        $settingsText | Should -Match 'Style="\{StaticResource Section\.Header\}"'
        $settingsText | Should -Match 'Style="\{StaticResource Section\.Row\}"'
    }
}

Describe "theme のキー（色の値は Figma の設計どおり）" -Tag Meta {
    BeforeDiscovery {
        # Figma の Variables の値。`/` を `.` に替えたキーで theme.xaml に置く。値を変えるときはこの表も同じ PR で直す
        $themeColors = @(
        @{ Key = "Bg.Window"; Color = "#F5F7FA"; Opacity = 1 }
        @{ Key = "Bg.Surface"; Color = "#FFFFFF"; Opacity = 1 }
        @{ Key = "Bg.Subtle"; Color = "#F9FAFA"; Opacity = 1 }
        @{ Key = "Bg.Stripe"; Color = "#FAFBFC"; Opacity = 1 }
        @{ Key = "Bg.Hover"; Color = "#F3F3F4"; Opacity = 1 }
        @{ Key = "Bg.Pane"; Color = "#F0F2F4"; Opacity = 1 }
        @{ Key = "Bg.Tag"; Color = "#F1F3F4"; Opacity = 1 }
        @{ Key = "Bg.Button"; Color = "#F2F2F5"; Opacity = 1 }
        @{ Key = "Bg.Summary"; Color = "#F7FAFC"; Opacity = 1 }
        @{ Key = "Bg.TitleBar"; Color = "#F0F0F0"; Opacity = 1 }
        @{ Key = "Border.Soft"; Color = "#E0E2E5"; Opacity = 1 }
        @{ Key = "Border.Normal"; Color = "#D9DEE3"; Opacity = 1 }
        @{ Key = "Border.Input"; Color = "#D1D1D1"; Opacity = 1 }
        @{ Key = "Border.Strong"; Color = "#C9CED4"; Opacity = 1 }
        @{ Key = "Border.Separator"; Color = "#D5D9DE"; Opacity = 1 }
        @{ Key = "Border.Splitter"; Color = "#D0D4D9"; Opacity = 1 }
        @{ Key = "Border.Grip"; Color = "#A9AFB6"; Opacity = 1 }
        @{ Key = "Border.Divider"; Color = "#E5E8ED"; Opacity = 1 }
        @{ Key = "Border.Row"; Color = "#EDF0F2"; Opacity = 1 }
        @{ Key = "Border.Dialog"; Color = "#D1D6E0"; Opacity = 1 }
        @{ Key = "Border.Check"; Color = "#9EA3AB"; Opacity = 1 }
        @{ Key = "Border.Window"; Color = "#999999"; Opacity = 1 }
        @{ Key = "Overlay.Scrim"; Color = "#000000"; Opacity = 0.35 }
        @{ Key = "Ink.Strong"; Color = "#202124"; Opacity = 1 }
        @{ Key = "Ink.Value"; Color = "#212126"; Opacity = 1 }
        @{ Key = "Ink.Body"; Color = "#5F6368"; Opacity = 1 }
        @{ Key = "Ink.Muted"; Color = "#6B737D"; Opacity = 1 }
        @{ Key = "Ink.Subtle"; Color = "#80868B"; Opacity = 1 }
        @{ Key = "Ink.Placeholder"; Color = "#9AA0A6"; Opacity = 1 }
        @{ Key = "Ink.Faint"; Color = "#99A1AB"; Opacity = 1 }
        @{ Key = "Ink.Note"; Color = "#8C949E"; Opacity = 1 }
        @{ Key = "Ink.OnAccent"; Color = "#FFFFFF"; Opacity = 1 }
        @{ Key = "Button.Text"; Color = "#4D4D4D"; Opacity = 1 }
        @{ Key = "Button.Icon"; Color = "#666666"; Opacity = 1 }
        @{ Key = "Accent"; Color = "#0078D4"; Opacity = 1 }
        @{ Key = "Accent.Hover"; Color = "#0B5CAD"; Opacity = 1 }
        @{ Key = "Accent.Soft"; Color = "#E5F1FB"; Opacity = 1 }
        @{ Key = "Select.Soft"; Color = "#E1F2FF"; Opacity = 1 }
        @{ Key = "Select.Text"; Color = "#0078D4"; Opacity = 1 }
        @{ Key = "Hit"; Color = "#FFF176"; Opacity = 1 }
        @{ Key = "Hit.Cell"; Color = "#FFF3CD"; Opacity = 1 }
        @{ Key = "Ok"; Color = "#218A21"; Opacity = 1 }
        @{ Key = "Ok.Strong"; Color = "#1E8E3E"; Opacity = 1 }
        @{ Key = "Ok.Soft"; Color = "#E0F7E0"; Opacity = 1 }
        @{ Key = "Warn"; Color = "#BA7D00"; Opacity = 1 }
        @{ Key = "Warn.Dot"; Color = "#E8A400"; Opacity = 1 }
        @{ Key = "Warn.Soft"; Color = "#FFF5E0"; Opacity = 1 }
        @{ Key = "Warn.Note"; Color = "#FFF7E0"; Opacity = 1 }
        @{ Key = "Warn.Line"; Color = "#F0C36D"; Opacity = 1 }
        @{ Key = "Warn.Strong"; Color = "#6B4E00"; Opacity = 1 }
        @{ Key = "Danger.Text"; Color = "#D13438"; Opacity = 1 }
        @{ Key = "Danger.Dot"; Color = "#D93025"; Opacity = 1 }
        @{ Key = "Danger.Soft"; Color = "#FFE6E6"; Opacity = 1 }
        @{ Key = "File.Excel"; Color = "#107C41"; Opacity = 1 }
        @{ Key = "File.Word"; Color = "#185ABD"; Opacity = 1 }
        @{ Key = "File.PowerPoint"; Color = "#C43E1C"; Opacity = 1 }
        @{ Key = "File.Folder"; Color = "#E8A020"; Opacity = 1 }
        @{ Key = "Illust.Line"; Color = "#D1D6DE"; Opacity = 1 }
        )
    }

    BeforeAll {
        $themeXml = New-Object System.Xml.XmlDocument
        $themeXml.Load("${scriptsDir}\shared\xaml\theme.xaml")
        $xns = "http://schemas.microsoft.com/winfx/2006/xaml"
        $brushes = @{}
        foreach ($node in $themeXml.DocumentElement.ChildNodes) {
            if ($node.LocalName -eq "SolidColorBrush") {
                $brushes[$node.GetAttribute("Key", $xns)] = $node
            }
        }
    }

    It "<key> が <color>（不透明度 <opacity>）" -ForEach $themeColors {
        $brushes.ContainsKey($key) | Should -Be $true
        $brushes[$key].GetAttribute("Color") | Should -Be $color
        $actualOpacity = if ($brushes[$key].HasAttribute("Opacity")) { [double]$brushes[$key].GetAttribute("Opacity") } else { 1 }
        $actualOpacity | Should -Be $opacity
    }

    It "入力欄の選択色は Select.Text を 40% で重ねる" {
        $setters = @{}
        foreach ($node in $themeXml.DocumentElement.ChildNodes) {
            if ($node.LocalName -eq "Style" -and $node.GetAttribute("Key", $xns) -eq "TextBox.Base") {
                foreach ($setter in $node.ChildNodes) {
                    if ($setter.LocalName -eq "Setter") { $setters[$setter.GetAttribute("Property")] = $setter.GetAttribute("Value") }
                }
            }
        }
        $setters["SelectionBrush"] | Should -Be "{StaticResource Select.Text}"
        $setters["SelectionOpacity"] | Should -Be "0.4"
    }
}

Describe "theme の文字の Style（大きさ・太さ・行の高さは Figma の設計どおり）" -Tag Meta {
    BeforeDiscovery {
        # Weight が空は Normal。LineHeight が空は指定しない
        $textStyles = @(
            @{ Key = "Micro";        Size = "10"; Weight = "";         LineHeight = "" }
            @{ Key = "ColumnHeader"; Size = "10"; Weight = "Bold";     LineHeight = "" }
            @{ Key = "Meta";         Size = "11"; Weight = "";         LineHeight = "" }
            @{ Key = "Meta.Tall";    Size = "";   Weight = "";         LineHeight = "16" }
            @{ Key = "Meta.Strong";  Size = "";   Weight = "Bold";     LineHeight = "" }
            @{ Key = "Meta.Key";     Size = "";   Weight = "SemiBold"; LineHeight = "" }   # 行の高さ 16 は Meta.Tall から継ぐ
            @{ Key = "Note";         Size = "11"; Weight = "";         LineHeight = "17.6" }
            @{ Key = "Chip";         Size = "11"; Weight = "Medium";   LineHeight = "" }
            @{ Key = "Link.Text";    Size = "11"; Weight = "SemiBold"; LineHeight = "" }
            @{ Key = "Cell";         Size = "12"; Weight = "";         LineHeight = "" }
            @{ Key = "Cell.Key";     Size = "";   Weight = "Medium";   LineHeight = "" }
            @{ Key = "Label";        Size = "12"; Weight = "SemiBold"; LineHeight = "" }
            @{ Key = "Label.Strong"; Size = "";   Weight = "Bold";     LineHeight = "" }
            @{ Key = "Brand";        Size = "12"; Weight = "SemiBold"; LineHeight = "" }
            @{ Key = "Body";         Size = "13"; Weight = "";         LineHeight = "" }
            @{ Key = "Body.Strong";  Size = "";   Weight = "Bold";     LineHeight = "" }
            @{ Key = "Nav";          Size = "13"; Weight = "Medium";   LineHeight = "" }
            @{ Key = "Nav.Tall";     Size = "";   Weight = "";         LineHeight = "18" }
            @{ Key = "Heading";      Size = "13"; Weight = "SemiBold"; LineHeight = "" }
            @{ Key = "Focal";        Size = "14"; Weight = "SemiBold"; LineHeight = "" }
            @{ Key = "PageTitle";    Size = "16"; Weight = "Bold";     LineHeight = "" }
            @{ Key = "Title";        Size = "18"; Weight = "SemiBold"; LineHeight = "" }
            @{ Key = "AppTitle";     Size = "20"; Weight = "SemiBold"; LineHeight = "28" }
        )
    }

    BeforeAll {
        $xns = "http://schemas.microsoft.com/winfx/2006/xaml"
        $themeXml = New-Object System.Xml.XmlDocument
        $themeXml.Load("${scriptsDir}\shared\xaml\theme.xaml")
        $styles = @{}
        foreach ($node in $themeXml.DocumentElement.ChildNodes) {
            if ($node.LocalName -eq "Style" -and $node.GetAttribute("TargetType") -eq "TextBlock") {
                $styles[$node.GetAttribute("Key", $xns)] = $node
            }
        }
        # Style の Setter の値。BasedOn を持つ Style は、継ぐ前の Style の値は見ずに、その Style が書いた分だけを見る
        function getSetter($style, [string]$property) {
            foreach ($s in $style.ChildNodes) {
                if ($s.LocalName -eq "Setter" -and $s.GetAttribute("Property") -eq $property) { return $s.GetAttribute("Value") }
            }
            return ""
        }
    }

    It "<key> の大きさ <size>・太さ <weight>・行の高さ <lineHeight>" -ForEach $textStyles {
        $styles.ContainsKey($key) | Should -Be $true
        getSetter $styles[$key] "FontSize" | Should -Be $size
        getSetter $styles[$key] "FontWeight" | Should -Be $weight
        getSetter $styles[$key] "LineHeight" | Should -Be $lineHeight
    }
}

Describe "theme のアイコン（Geometry）と図の Style" -Tag Meta {
    BeforeDiscovery {
        $iconKeys = @(
            "Folder", "FolderClosed", "FolderOpen", "ChevronDown", "ChevronRight", "RefreshCw", "TriangleAlert",
            "File", "FileText", "FileSpreadsheet", "FolderSearch", "Presentation", "ChartColumn", "Save",
            "StopCircle", "Close", "CircleX", "CircleCheck", "Search", "CircleQuestionMark", "Check", "Zap",
            "Info", "InfoCircle", "InfoGlyph.S12", "InfoGlyph.S13", "InfoGlyph.S14", "InfoGlyph.S16",
            "BadgeGlyphInfo", "BadgeGlyphWarn", "BadgeGlyphError", "BadgeGlyphOk", "Ellipsis"
        ) | ForEach-Object { @{ Key = "Icon.$_" } }
    }

    BeforeAll {
        $xns = "http://schemas.microsoft.com/winfx/2006/xaml"
        $themeXml = New-Object System.Xml.XmlDocument
        $themeXml.Load("${scriptsDir}\shared\xaml\theme.xaml")
        $geometries = @{}
        $keys = @()
        foreach ($node in $themeXml.DocumentElement.ChildNodes) {
            if ($node.NodeType -ne "Element") { continue }
            $k = $node.GetAttribute("Key", $xns)
            $keys += $k
            if ($node.LocalName -eq "Geometry") { $geometries[$k] = $node.InnerText }
        }
    }

    It "<key> が Geometry で、空でない" -ForEach $iconKeys {
        $geometries.ContainsKey($key) | Should -Be $true
        $geometries[$key].Trim() | Should -Not -BeNullOrEmpty
    }

    It "x:Key が重ならない（ResourceDictionary は重なると読み込みで失敗する）" {
        ($keys | Group-Object | Where-Object Count -gt 1 | ForEach-Object Name) | Should -BeNullOrEmpty
    }

    It "<_> がある" -ForEach @("Illust.Path", "Radius.Pill", "Chip.Box", "NavBadge.Count", "NavBadge.Dot", "StatusBadge", "Banner") {
        $keys | Should -Contain $_
    }
}

Describe "型の読み込み" -Tag Meta {
    # 画面で使う型は shared と tebunko に分かれている。gui.ps1 と同じ順で読み込めば、
    # 継承（NotifyBase を継承する型）が解決できることを確かめる
    It "shared と tebunko の型を順に読み込める" {
        $probe = Join-Path $TestDrive "probe.ps1"
        $scripts = (Resolve-Path "$here\..\scripts").Path
        Set-Content -LiteralPath $probe -Encoding UTF8 -Value @(
            'Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase'
            ". `"$scripts\shared\shared.ps1`""
            ". `"$scripts\tebunko\lib.ps1`""
            ". `"$scripts\shared\ui\types.ps1`""
            ". `"$scripts\tebunko\ui\search\result_types.ps1`""
            ". `"$scripts\tebunko\ui\search\target_types.ps1`""
            ". `"$scripts\tebunko\ui\index\index_types.ps1`""
            '([HitRow], [IndexNode], [ConfirmFact], [PreviewTable]).Count'
        )
        $output = & powershell -NoProfile -ExecutionPolicy Bypass -File $probe 2>&1
        ($output -join "") | Should -Be "4"
    }
}

Describe "単一 .ps1 化の決まり（AST。docs/design/structure/single-script.md）" -Tag Meta {
    BeforeAll {
        $singleScriptsDir = (Resolve-Path "$here\..\scripts").Path
        $m1ExemptFiles = @("parts.ps1", "paths.ps1", "gui.ps1", "indexer.ps1")

        function testPathPatternRisk {
            # パスをパターンとして読むコマンド（M3）に、変数や式のパスを -LiteralPath なしで渡している CommandAst なら $true。
            # M3 の本体と、M3 の例の表の両方がこの関数を呼ぶ
            param ($c)
            $pathCommands = @("Resolve-Path", "New-Item", "Get-ChildItem", "Get-Item", "Test-Path", "Get-Content", "Set-Content", "Add-Content",
                "Out-File", "Remove-Item", "Copy-Item", "Move-Item", "Rename-Item", "Convert-Path", "Import-Csv", "Export-Csv", "Select-String", "Get-FileHash")
            # パイプで渡された文字列を Path として受けるもの
            $pipeByValue = @("Resolve-Path", "Get-ChildItem", "Get-Item", "Test-Path", "Get-Content", "Remove-Item", "Convert-Path",
                "Copy-Item", "Move-Item", "Rename-Item", "Import-Csv", "Get-FileHash")
            # 値を取らない（スイッチの）名前。直後の要素は位置引数になる
            $valueFree = @("Recurse", "Force", "File", "Directory", "Hidden", "ReadOnly", "System", "Append", "NoClobber", "WhatIf", "Confirm",
                "NoNewline", "Raw", "SimpleMatch", "Quiet", "CaseSensitive", "NotMatch", "List", "AllMatches",
                "PassThru", "NoTypeInformation", "IsValid", "Container", "Wait")
            $paramType = [System.Management.Automation.Language.CommandParameterAst]
            $ignoreCase = [System.StringComparison]::OrdinalIgnoreCase
            $isRiskyValue = {
                param ($e)
                # @params（スプラット）は中身が見えないので対象にしない
                if ($e -is [System.Management.Automation.Language.VariableExpressionAst]) { return (-not $e.Splatted) }
                return ($e -is [System.Management.Automation.Language.ExpandableStringExpressionAst] -or
                    $e -is [System.Management.Automation.Language.ParenExpressionAst] -or
                    $e -is [System.Management.Automation.Language.SubExpressionAst] -or
                    $e -is [System.Management.Automation.Language.IndexExpressionAst] -or
                    $e -is [System.Management.Automation.Language.MemberExpressionAst])
            }

            $name = $c.GetCommandName()
            if (-not $name -or $pathCommands -notcontains $name) { return $false }
            $elements = @($c.CommandElements | Select-Object -Skip 1)
            if (@($elements | Where-Object { $_ -is $paramType -and $_.ParameterName -like "Lit*" }).Count -gt 0) { return $false }

            $skipNext = $false
            $positional = 0
            $namedPattern = ($name -eq "Select-String" -and @($elements | Where-Object { $_ -is $paramType -and $_.ParameterName.Length -ge 4 -and "Pattern".StartsWith($_.ParameterName, $ignoreCase) }).Count -gt 0)
            foreach ($e in $elements) {
                if ($e -is $paramType) {
                    $pn = $e.ParameterName
                    $isSwitchName = (@($valueFree | Where-Object { $_.Equals($pn, $ignoreCase) }).Count -gt 0)
                    # Select-String の -Pattern（-Patt 以上）は検索する文字。-Path と取り違えない
                    $isPatternName = ($name -eq "Select-String" -and $pn.Length -ge 4 -and "Pattern".StartsWith($pn, $ignoreCase))
                    # -File などのスイッチは、-FilePath の頭と一致しても対象にしない。-FilePath は 5 文字以上の一致だけ見る
                    if (-not $isSwitchName -and -not $isPatternName -and ("Path".StartsWith($pn, $ignoreCase) -or ($pn.Length -ge 5 -and "FilePath".StartsWith($pn, $ignoreCase)))) { return $true }
                    if ($null -ne $e.Argument) { continue }
                    $skipNext = (@($valueFree | Where-Object { $_.StartsWith($pn, $ignoreCase) }).Count -eq 0)
                    continue
                }
                if ($skipNext) { $skipNext = $false; continue }
                $index = $positional
                $positional++
                # Select-String の最初の位置引数は検索する文字（Pattern）で、パスではない
                if ($name -eq "Select-String" -and $index -eq 0 -and -not $namedPattern) { continue }
                if (& $isRiskyValue $e) { return $true }
            }

            # パスの指定が無く、パイプで変数や式の文字列が流れ込む形（$x | Get-Content）
            if ($positional -eq 0 -and $pipeByValue -contains $name -and $c.Parent -is [System.Management.Automation.Language.PipelineAst]) {
                $elementsOfPipeline = @($c.Parent.PipelineElements)
                $at = [array]::IndexOf($elementsOfPipeline, $c)
                if ($at -gt 0) {
                    $source = $elementsOfPipeline[$at - 1]
                    if ($source -is [System.Management.Automation.Language.CommandExpressionAst] -and (& $isRiskyValue $source.Expression)) { return $true }
                    if ($source -is [System.Management.Automation.Language.CommandAst] -and $source.GetCommandName() -eq "Join-Path") { return $true }
                }
            }
            return $false
        }

        function getRelPath {
            param ([string]$fullName)
            return $fullName.Substring($singleScriptsDir.Length + 1)
        }

        function isLoaderLine {
            # 読み込み口の行: . ([WildcardPattern]::Escape("$PSScriptRoot\...")) の形。
            # パスにワイルドカードの文字（[ ]）が入っていても読めるよう、相手のパスは必ず Escape を通す
            # （通さない . "$PSScriptRoot\..." は、読み込み口の行として認めない）
            param ($commandAst)
            if ($commandAst.InvocationOperator -ne "Dot") { return $false }
            if ($commandAst.CommandElements.Count -lt 1) { return $false }
            $first = $commandAst.CommandElements[0]
            if ($first -isnot [System.Management.Automation.Language.ParenExpressionAst]) { return $false }
            $inner = $first.Pipeline
            if ($inner -isnot [System.Management.Automation.Language.PipelineAst] -or $inner.PipelineElements.Count -ne 1) { return $false }
            $expr = $inner.PipelineElements[0].Expression
            if ($expr -isnot [System.Management.Automation.Language.InvokeMemberExpressionAst]) { return $false }
            if (-not $expr.Static -or $expr.Member.Value -ne "Escape") { return $false }
            if ($expr.Expression -isnot [System.Management.Automation.Language.TypeExpressionAst] -or $expr.Expression.TypeName.Name -ne "WildcardPattern") { return $false }
            if ($expr.Arguments.Count -ne 1) { return $false }
            # 引数は "<パス>".Replace('`', '``') の形（バッククォートは Escape がエスケープしないので、先に二重にする。対象は PowerShell 5.1 だけ）
            $arg = $expr.Arguments[0]
            if ($arg -isnot [System.Management.Automation.Language.InvokeMemberExpressionAst] -or $arg.Static -or $arg.Member.Value -ne "Replace") { return $false }
            if ($arg.Expression -isnot [System.Management.Automation.Language.ExpandableStringExpressionAst] -and
                $arg.Expression -isnot [System.Management.Automation.Language.StringConstantExpressionAst]) { return $false }
            if ($arg.Arguments.Count -ne 2) { return $false }
            return $arg.Arguments[0].Extent.Text -eq ("'" + [char]96 + "'") -and $arg.Arguments[1].Extent.Text -eq ("'" + [char]96 + [char]96 + "'")
        }

        function commandOf {
            # 文（PipelineAst）の中の、唯一の要素（CommandAst）を取り出す。それ以外はそのまま返す
            param ($stmt)
            if ($stmt -is [System.Management.Automation.Language.PipelineAst] -and $stmt.PipelineElements.Count -eq 1) {
                return $stmt.PipelineElements[0]
            }
            return $stmt
        }

        function isTopLevel {
            # 関数・スクリプトブロックの外（root との間に ScriptBlockAst が無い）かどうか
            param ($node, $root)
            $p = $node.Parent
            while ($null -ne $p -and $p -ne $root) {
                if ($p -is [System.Management.Automation.Language.ScriptBlockAst]) { return $false }
                $p = $p.Parent
            }
            return $true
        }

        $allScriptFiles = @(Get-ChildItem -LiteralPath $singleScriptsDir -Recurse -Filter "*.ps1")
        $parsedAsts = @{}
        foreach ($f in $allScriptFiles) {
            $parsedAsts[$f.FullName] = [System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$null, [ref]$null)
        }
    }

    It "M1: `$PSScriptRoot を書けるのは、読み込み口の行・parts.ps1・shared/core/paths.ps1・起動口（gui.ps1・indexer.ps1）だけ" {
        $bad = New-Object System.Collections.Generic.List[string]
        foreach ($f in $allScriptFiles) {
            if ($m1ExemptFiles -contains $f.Name) { continue }
            $ast = $parsedAsts[$f.FullName]
            $vars = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.VariableExpressionAst] -and $args[0].VariablePath.UserPath -eq "PSScriptRoot" }, $true)
            foreach ($v in $vars) {
                # 読み込み口の行（. "$PSScriptRoot\..."）の中で使うのは許す
                $cmd = $v
                while ($null -ne $cmd -and -not ($cmd -is [System.Management.Automation.Language.CommandAst])) { $cmd = $cmd.Parent }
                if ($null -ne $cmd -and (isLoaderLine $cmd)) { continue }
                $bad.Add("$(getRelPath $f.FullName):$($v.Extent.StartLineNumber)")
            }
        }
        ($bad -join ", ") | Should -Be ""
    }

    It "M2(i): パス区切りを含み .ps1 で終わる文字列は、読み込み口の行・parts.ps1 だけに書く" {
        # text_file.ps1 の ".ps1"（拡張子だけの一覧）のような、パスではない文字列は対象外にする
        $bad = New-Object System.Collections.Generic.List[string]
        foreach ($f in $allScriptFiles) {
            if ($f.Name -eq "parts.ps1") { continue }
            $ast = $parsedAsts[$f.FullName]
            $strs = $ast.FindAll({
                ($args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] -or $args[0] -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) -and
                $args[0].Value -match '\.ps1$' -and $args[0].Value -match '[\\/]'
            }, $true)
            foreach ($s in $strs) {
                $cmd = $s
                while ($null -ne $cmd -and -not ($cmd -is [System.Management.Automation.Language.CommandAst])) { $cmd = $cmd.Parent }
                if ($null -ne $cmd -and (isLoaderLine $cmd)) { continue }
                $bad.Add("$(getRelPath $f.FullName):$($s.Extent.StartLineNumber):$($s.Extent.Text)")
            }
        }
        ($bad -join ", ") | Should -Be ""
    }

    It "M2(ii): ドットソース（.）は、読み込み口の行だけ（変数・パスでの dot-source は認めない）" {
        $bad = New-Object System.Collections.Generic.List[string]
        foreach ($f in $allScriptFiles) {
            $ast = $parsedAsts[$f.FullName]
            $cmds = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] -and $args[0].InvocationOperator -eq "Dot" }, $true)
            foreach ($c in $cmds) {
                if (isLoaderLine $c) { continue }
                $bad.Add("$(getRelPath $f.FullName):$($c.Extent.StartLineNumber):$($c.Extent.Text)")
            }
        }
        ($bad -join ", ") | Should -Be ""
    }

    It "M2(ii) の例: 変数やプロパティを渡す形（. `$libPath・. `$x.Path）と、Escape を通さない形は読み込み口の行として検出しない" {
        $badSamples = @('. $libPath', '. $x.Path', '. "$PSScriptRoot\x.ps1"', '. (Escape "$PSScriptRoot\x.ps1")', '. ([WildcardPattern]::Escape("$PSScriptRoot\x.ps1"))')
        foreach ($sample in $badSamples) {
            $fixtureAst = [System.Management.Automation.Language.Parser]::ParseInput($sample, [ref]$null, [ref]$null)
            $cmd = @($fixtureAst.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true))[0]
            (isLoaderLine $cmd) | Should -Be $false -Because "形: $sample"
        }
    }

    It "M2(iv): & の相手に、パスの文字列は直接書かない（スクリプトのパスは Escape を通す。変数・式は、スクリプトブロックと見分けられないので検査しない）" {
        # 変数・式（& $x・& ($x)）は、スクリプトブロック（& { }）と構文だけでは見分けられないので検査しない（元の不具合の & $gui も、この検査では止まらない）。文字列で書いたパスだけを見る
        $bad = New-Object System.Collections.Generic.List[string]
        foreach ($f in $allScriptFiles) {
            $ast = $parsedAsts[$f.FullName]
            $cmds = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] -and $args[0].InvocationOperator -eq "Ampersand" }, $true)
            foreach ($c in $cmds) {
                $first = $c.CommandElements[0]
                if ($first -is [System.Management.Automation.Language.StringConstantExpressionAst] -or $first -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) {
                    $bad.Add("$(getRelPath $f.FullName):$($c.Extent.StartLineNumber):$($c.Extent.Text)")
                }
            }
        }
        ($bad -join ", ") | Should -Be ""
    }

    It "M3: パスをワイルドカードとして読むコマンド（-Path・-FilePath・位置引数・パイプ）に、変数や式のパスを渡さない（-LiteralPath を使う。-Filter・-like は別）" {
        $bad = New-Object System.Collections.Generic.List[string]
        foreach ($f in $allScriptFiles) {
            $ast = $parsedAsts[$f.FullName]
            $cmds = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)
            foreach ($c in $cmds) {
                if (testPathPatternRisk $c) { $bad.Add("$(getRelPath $f.FullName):$($c.Extent.StartLineNumber):$($c.GetCommandName())") }
            }
        }
        ($bad -join ", ") | Should -Be ""
    }

    It "M3 の例（本体と同じ判定）: <code> は検出<expected>" -ForEach @(
        @{ code = 'Resolve-Path $x'; expected = $true }
        @{ code = 'New-Item -Path $x -Force'; expected = $true }
        @{ code = 'Get-Item "$dir\a"'; expected = $true }
        @{ code = 'Get-ChildItem -Recurse $dir'; expected = $true }
        @{ code = 'Remove-Item -Force $x'; expected = $true }
        @{ code = 'New-Item -Force $x'; expected = $true }
        @{ code = 'Out-File -FilePath $x'; expected = $true }
        @{ code = '$x | Get-Content'; expected = $true }
        @{ code = 'Join-Path $a $b | Get-Item'; expected = $true }
        @{ code = 'Select-String $pattern $file'; expected = $true }
        @{ code = '$rows | Export-Csv -NoTypeInformation $path'; expected = $true }
        @{ code = 'Copy-Item -PassThru $x $y'; expected = $true }
        @{ code = 'Select-String -Pattern $p $file'; expected = $true }
        @{ code = 'Select-String $file -Pattern $p'; expected = $true }
        @{ code = 'Get-Item $a[0]'; expected = $true }
        @{ code = 'Get-Item $(Join-Path $a $b)'; expected = $true }
        @{ code = 'Get-ChildItem -File'; expected = $false }
        @{ code = 'Get-ChildItem -File -Filter *.x'; expected = $false }
        @{ code = 'Get-ChildItem @params'; expected = $false }
        @{ code = 'Select-String -Pattern $p -LiteralPath $file'; expected = $false }
        @{ code = 'Get-Item -LiteralPath $x'; expected = $false }
        @{ code = 'Get-ChildItem -Recurse -File -LiteralPath $dir'; expected = $false }
        @{ code = '$text | Select-String $pattern'; expected = $false }
        @{ code = '$x | Set-Content -LiteralPath $y'; expected = $false }
        @{ code = 'Get-ChildItem -LiteralPath $dir | Get-Content -Raw -LiteralPath $f'; expected = $false }
    ) {
        $fixtureAst = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$null, [ref]$null)
        $cmds = @($fixtureAst.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true))
        $hit = @($cmds | Where-Object { testPathPatternRisk $_ }).Count -gt 0
        $hit | Should -Be $expected
    }
    It "M2(iii): `&` での `.Path` 呼び出し・`-Part` を付けた呼び出しは書かない（自己起動を禁じる）" {
        $bad = New-Object System.Collections.Generic.List[string]
        foreach ($f in $allScriptFiles) {
            $ast = $parsedAsts[$f.FullName]
            $cmds = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] -and $args[0].InvocationOperator -eq "Ampersand" }, $true)
            foreach ($c in $cmds) {
                $first = $c.CommandElements[0]
                $isPathCall = $first -is [System.Management.Automation.Language.MemberExpressionAst] -and
                    $first.Member -is [System.Management.Automation.Language.StringConstantExpressionAst] -and
                    $first.Member.Value -eq "Path"
                $hasPartFlag = @($c.CommandElements | Where-Object {
                    $_ -is [System.Management.Automation.Language.CommandParameterAst] -and $_.ParameterName -eq "Part"
                }).Count -gt 0
                if ($isPathCall -or $hasPartFlag) {
                    $bad.Add("$(getRelPath $f.FullName):$($c.Extent.StartLineNumber):$($c.Extent.Text)")
                }
            }
        }
        ($bad -join ", ") | Should -Be ""
    }

    It "M2(iii) の例: `-Part` を含む文字列は自己起動（Start-Process の引数など）に書かない" {
        $bad = New-Object System.Collections.Generic.List[string]
        foreach ($f in $allScriptFiles) {
            $ast = $parsedAsts[$f.FullName]
            $strs = $ast.FindAll({
                ($args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] -or $args[0] -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) -and
                $args[0].Value -match '-Part\b'
            }, $true)
            foreach ($s in $strs) {
                $bad.Add("$(getRelPath $f.FullName):$($s.Extent.StartLineNumber):$($s.Extent.Text)")
            }
        }
        ($bad -join ", ") | Should -Be ""
    }

    It "M2(iv): `. ` で始まる文字列・importTebunkoPart の名前が書けるのは parts.ps1 だけ" {
        $bad = New-Object System.Collections.Generic.List[string]
        foreach ($f in $allScriptFiles) {
            if ($f.Name -eq "parts.ps1") { continue }
            $ast = $parsedAsts[$f.FullName]
            $strs = $ast.FindAll({
                ($args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] -or $args[0] -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) -and
                ($args[0].Value -like ". *" -or $args[0].Value -like "*importTebunkoPart*")
            }, $true)
            foreach ($s in $strs) {
                $bad.Add("$(getRelPath $f.FullName):$($s.Extent.StartLineNumber):$($s.Extent.Text)")
            }
        }
        ($bad -join ", ") | Should -Be ""
    }

    It "M2(v): `${bundledScriptPath}`（単一 .ps1 自身のパス）を dot-source・& での呼び出しに使わない" {
        $bad = New-Object System.Collections.Generic.List[string]
        foreach ($f in $allScriptFiles) {
            $ast = $parsedAsts[$f.FullName]
            $cmds = $ast.FindAll({
                $args[0] -is [System.Management.Automation.Language.CommandAst] -and
                ($args[0].InvocationOperator -eq "Dot" -or $args[0].InvocationOperator -eq "Ampersand")
            }, $true)
            foreach ($c in $cmds) {
                $hit = @($c.FindAll({ $args[0] -is [System.Management.Automation.Language.VariableExpressionAst] -and $args[0].VariablePath.UserPath -eq "bundledScriptPath" }, $true)).Count -gt 0
                if ($hit) {
                    $bad.Add("$(getRelPath $f.FullName):$($c.Extent.StartLineNumber):$($c.Extent.Text)")
                }
            }
        }
        ($bad -join ", ") | Should -Be ""
    }

    It "M3: XAML の ResourceDictionary Source= は theme.xaml を指すものだけ" {
        $bad = New-Object System.Collections.Generic.List[string]
        $xamlFiles = @(Get-ChildItem -LiteralPath $singleScriptsDir -Recurse -Filter "*.xaml")
        foreach ($f in $xamlFiles) {
            $text = [System.IO.File]::ReadAllText($f.FullName)
            foreach ($m in [regex]::Matches($text, 'ResourceDictionary\s+Source="([^"]+)"')) {
                if ($m.Groups[1].Value -notmatch 'theme\.xaml$') {
                    $bad.Add("$($f.Name): $($m.Groups[1].Value)")
                }
            }
        }
        ($bad -join ", ") | Should -Be ""
    }

    It "M4: トップレベルの trap・exit・`$ErrorActionPreference の代入は、起動口の決まった形だけ" {
        $bad = New-Object System.Collections.Generic.List[string]
        foreach ($f in $allScriptFiles) {
            $ast = $parsedAsts[$f.FullName]
            $nodes = $ast.FindAll({
                $args[0] -is [System.Management.Automation.Language.TrapStatementAst] -or
                $args[0] -is [System.Management.Automation.Language.ExitStatementAst] -or
                ($args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] -and
                 $args[0].Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
                 $args[0].Left.VariablePath.UserPath -eq "ErrorActionPreference")
            }, $true)
            foreach ($n in $nodes) {
                if (-not (isTopLevel $n $ast)) { continue }
                $sanctioned =
                    ($f.Name -eq "indexer.ps1" -and $n -is [System.Management.Automation.Language.ExitStatementAst] -and $n.Extent.Text -eq 'exit (invokeIndexerMain -RetryFailed:$RetryFailed -Channel $Channel)') -or
                    ($f.Name -eq "gui.ps1" -and $n -is [System.Management.Automation.Language.ExitStatementAst] -and $n.Extent.Text -eq 'exit (reportStartupFailure $_)')
                if (-not $sanctioned) {
                    $bad.Add("$(getRelPath $f.FullName):$($n.Extent.StartLineNumber):$($n.Extent.Text)")
                }
            }
        }
        ($bad -join ", ") | Should -Be ""
    }

    It "gui.ps1 の最上位は、起動の失敗の知らせ（paths.ps1・startup_error_view.ps1・startup_error.ps1）の読み込みと 1 つの try/catch だけ" {
        $ast = $parsedAsts["$singleScriptsDir\tebunko\gui.ps1"]
        $top = @($ast.EndBlock.Statements)
        $top.Count | Should -Be 4
        foreach ($loader in $top[0..2]) {
            (commandOf $loader) | Should -BeOfType [System.Management.Automation.Language.CommandAst]
            (isLoaderLine (commandOf $loader)) | Should -Be $true
        }
        $top[3] | Should -BeOfType [System.Management.Automation.Language.TryStatementAst]

        $try = $top[3]
        $try.CatchClauses.Count | Should -Be 1
        $catchBody = @($try.CatchClauses[0].Body.Statements)
        $catchBody.Count | Should -Be 1
        $catchBody[0].Extent.Text | Should -Be 'exit (reportStartupFailure $_)'
    }

    It "indexer.ps1 の最上位は、読み込み口の行と exit (invokeIndexerMain ...) だけ" {
        $ast = $parsedAsts["$singleScriptsDir\tebunko\indexer.ps1"]
        $top = @($ast.EndBlock.Statements)
        $last = $top[$top.Count - 1]
        $last.Extent.Text | Should -Be 'exit (invokeIndexerMain -RetryFailed:$RetryFailed -Channel $Channel)'
        foreach ($stmt in $top[0..($top.Count - 2)]) {
            (isLoaderLine (commandOf $stmt)) | Should -Be $true -Because "indexer.ps1 の最上位の文: $($stmt.Extent.Text)"
        }
    }

    It "M5: 結合の道具が出す単一 .ps1（`${bundledScriptPath} = `$PSCommandPath）がリポジトリに入っていない" {
        Push-Location $rootDir
        try {
            $files = git ls-files
        } finally {
            Pop-Location
        }
        $hits = New-Object System.Collections.Generic.List[string]
        foreach ($relFile in ($files | Where-Object { $_ -like "*.ps1" })) {
            # 結合の道具・そのテスト（new_single_script*・single_script_guard*）は、検査する文字列を、テストの見本として持つ
            if ($relFile -eq "tools/new_single_script.ps1" -or $relFile -like "tests/tools/new_single_script*" -or $relFile -like "tests/tools/single_script_guard*" -or $relFile -eq "tests/meta/structure.Tests.ps1") { continue }
            $full = Join-Path $rootDir $relFile
            try {
                $text = [System.IO.File]::ReadAllText($full)
            } catch { continue }
            if ($text.Contains('${bundledScriptPath} = $PSCommandPath')) {
                $hits.Add($relFile)
            }
        }
        ($hits -join ", ") | Should -Be ""
    }
}

Describe "画面の部品の名前" -Tag Meta {
    # startGui（gui_main.ps1。領域の表は shell\regions.ps1）が FindName で取る名前が、XAML に実在すること。
    # タブの中身を別ファイルに分けているため、名前を足したり動かしたりすると気づきにくい
    BeforeAll {
        $xamlNs = "http://schemas.microsoft.com/winfx/2006/xaml"
        $gui = [System.IO.File]::ReadAllText("$here\..\scripts\tebunko\ui\shell\regions.ps1")

        function getXamlNames {
            param ([string]$path)
            [xml]$xaml = Get-Content $path -Raw -Encoding UTF8
            return @($xaml.SelectNodes("//*") | ForEach-Object { $_.GetAttribute("Name", $xamlNs) } | Where-Object { $_ -ne "" })
        }
    }

    It "ウィンドウの枠の名前がある" {
        $names = getXamlNames "$here\..\scripts\tebunko\xaml\tebunko.xaml"
        foreach ($name in @("NavHost", "ContentHost", "StatusBarHost", "ScrimOverlay")) {
            $names -contains $name | Should -Be $true
        }
    }

    It "ナビ・ステータスバーに、画面の切り替えと他の画面が使う名前がある" {
        $nav = getXamlNames "$here\..\scripts\tebunko\xaml\shell\nav.xaml"
        foreach ($name in @("NavList", "SearchTab", "IndexTab", "SettingsTab", "IndexTabBadge", "AboutLink")) {
            $nav -contains $name | Should -Be $true
        }
        (getXamlNames "$here\..\scripts\tebunko\xaml\shell\status_bar.xaml") -contains "StatusText" | Should -Be $true
    }

    It "shell\regions.ps1 の領域の表の各ファイルに、使う名前がすべてある" {
        # $regions の `File = "..."; ... Names = @(...)` を全部取り出して、XAML と突き合わせる
        $entries = @([regex]::Matches($gui, 'File\s*=\s*"([^"]+)".*?Names\s*=\s*(?:@\(([^)]*)\)|\$(\w+))'))
        $entries.Count -gt 0 | Should -Be $true
        # 表の行の数と、読めた行の数が合う（形が崩れて読み落とした行が、確かめの外に出ない）
        $fileCount = @([regex]::Matches($gui, '@\{\s*File\s*=\s*"')).Count
        $entries.Count | Should -Be $fileCount
        $problems = New-Object System.Collections.Generic.List[string]
        foreach ($entry in $entries) {
            $file = $entry.Groups[1].Value
            $list = $entry.Groups[2].Value
            if ($entry.Groups[3].Success) {
                # Names = $変数名 のときは、変数の `@(...)` を読む
                $list = [regex]::Match($gui, '\$' + $entry.Groups[3].Value + '\s*=\s*@\(([^)]*)\)').Groups[1].Value
            }
            $wanted = @([regex]::Matches($list, '"([A-Za-z]+)"') | ForEach-Object { $_.Groups[1].Value })
            $path = "$here\..\scripts\tebunko\xaml\$file"
            if (-not (Test-Path -LiteralPath $path)) { $problems.Add("$file が無い"); continue }
            if ($wanted.Count -eq 0) { $problems.Add("$file の Names が読めない（空）") }
            $names = getXamlNames $path
            foreach ($w in $wanted) {
                if ($names -notcontains $w) { $problems.Add("$file に $w が無い") }
            }
        }
        ($problems -join ", ") | Should -Be ""
    }
}

Describe "ダイアログは暗幕付きで出す" -Tag Meta {
    It "scripts/tebunko の ShowDialog は、本体の窓を開く gui_main.ps1 だけ（ほかは showOwnedDialog で出す）" {
        $here = (Resolve-Path "$PSScriptRoot\..").Path
        $hits = @(Get-ChildItem "$here\..\scripts/tebunko" -Recurse -Filter *.ps1 | ForEach-Object {
            $file = $_
            @(Select-String -LiteralPath $file.FullName -Pattern '\.ShowDialog\(' | ForEach-Object { "$($file.Name):$($_.LineNumber)" })
        })
        $hits.Count | Should -Be 1
        $hits[0] | Should -BeLike "gui_main.ps1:*"
    }

    It "gui_main.ps1 は、ShowDialog() の finally のあとで画面のスレッドの Dispatcher を止める（止めないと、閉じたあとの PowerShell の終了で、まれに終了コード 5 になる。closing.md）" {
        $here = (Resolve-Path "$PSScriptRoot\..").Path
        $text = [System.IO.File]::ReadAllText("$here\..\scripts\tebunko\ui\gui_main.ps1")
        # 間には、コメントと、閉じる処理の記録（writeCloseTrace）の行だけを許す
        $text | Should -Match '\$window\.ShowDialog\(\)\r?\n    \} finally \{[\s\S]*?\r?\n    \}\r?\n(\s*(#|writeCloseTrace)[^\r\n]*\r?\n)*\s*\[System\.Windows\.Threading\.Dispatcher\]::CurrentDispatcher\.InvokeShutdown\(\)'
    }

    It "gui_main.ps1 は、ShowDialog() の finally の各段（Close・Abandon など）の前に、閉じる処理の記録（writeCloseTrace）を書く（止まった段を後から読めるようにする。closing.md）" {
        $here = (Resolve-Path "$PSScriptRoot\..").Path
        $lines = [System.IO.File]::ReadAllLines("$here\..\scripts\tebunko\ui\gui_main.ps1")
        $stages = @('\$script:indexingSession\.Close\(\)', '\$script:searchService\.Close\(\)', '\$script:backgroundQueue\.Abandon\(\)', '\$script:networkQueue\.Abandon\(\)', '\[System\.Windows\.Threading\.Dispatcher\]::CurrentDispatcher\.InvokeShutdown\(\)')
        foreach ($stage in $stages) {
            $index = @(0..($lines.Count - 1) | Where-Object { $lines[$_] -match "^\s*$stage" })
            $index.Count | Should -Be 1 -Because $stage
            $lines[$index[0] - 1] | Should -Match '^\s*writeCloseTrace ' -Because $stage
        }
    }

    It "startGui（gui_main.ps1）は 150 行以下。中身は shell\ の下のファイルに分けて読み込む" {
        $here = (Resolve-Path "$PSScriptRoot\..").Path
        $parseErrors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile("$here\..\scripts\tebunko\ui\gui_main.ps1", [ref]$null, [ref]$parseErrors)
        $startGuiAst = @($ast.FindAll({ param ($a) $a -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $a.Name -eq "startGui" }, $true))
        $startGuiAst.Count | Should -Be 1
        $lineCount = $startGuiAst[0].Extent.EndLineNumber - $startGuiAst[0].Extent.StartLineNumber + 1
        $lineCount | Should -BeLessOrEqual 150
    }
}

Describe "メッセージは自前の画面で出す" -Tag Meta {
    It "scripts の MessageBox の直接の呼び出しは、決めた所だけ（ほかは showMessage で出す）" {
        # shell.ps1 の showMessage … 自前の画面を出せないときの予備（別のスレッドから呼ばれた・画面の定義が読めない）
        # gui_main.ps1 … 二重起動の知らせ。画面の部品（loadWindow・テーマ）を読み込む前で、自前の画面を出せない
        # startup_error.ps1 … 起動の失敗の知らせ。制限言語モード・WPF が読めない場面でも知らせるため、標準のまま
        $allowed = @("shell.ps1", "gui_main.ps1", "startup_error.ps1")
        $here = (Resolve-Path "$PSScriptRoot\..").Path
        $hits = @(Get-ChildItem "$here\..\scripts" -Recurse -Filter *.ps1 | Where-Object { $allowed -notcontains $_.Name } | ForEach-Object {
            $file = $_
            @(Select-String -LiteralPath $file.FullName -Pattern 'MessageBox\]?::Show|Windows\.Forms\.MessageBox' | Where-Object { $_.Line.TrimStart() -notmatch '^#' } | ForEach-Object { "$($file.Name):$($_.LineNumber)" })
        })
        ($hits -join ", ") | Should -Be ""
    }

    It "shell.ps1 の MessageBox は、showMessage の予備の 2 か所だけ" {
        $here = (Resolve-Path "$PSScriptRoot\..").Path
        $hits = @(Select-String -LiteralPath "$here\..\scripts\shared\ui\shell.ps1" -Pattern 'MessageBox\]::Show' | Where-Object { $_.Line.TrimStart() -notmatch '^#' })
        $hits.Count | Should -Be 2
    }
}

Describe "フォルダ名に特別な文字が入っていても、本物の読み込み口を読める（scripts\ を写して読む）" -Tag Meta {
    BeforeDiscovery {
        . "$PSScriptRoot\..\helpers\path_chars.ps1"
        $pathCharCases = getPathCharCases
        # 名前の隣に、パターンとして読むと当たってしまう別のフォルダがある場面（実在する方を読むこと）
        $libCases = @($pathCharCases | Where-Object { @("一重引用符 '", "角括弧の中が空 []", "引用符・&・角括弧の組み合わせ", "バッククォートと開く角括弧", "日本語") -contains $_.label })
        $siblingCases = @(
            @{ label = "[1] と 1"; ch = "[1]"; sibling = "1" }
            @{ label = "バッククォートと、バッククォート抜き"; ch = ([string][char]96); sibling = "" }
        )
    }

    BeforeAll {
        $here = Split-Path -Parent $PSCommandPath
        $realScripts = (Resolve-Path -LiteralPath "$here\..\..\scripts").Path
        $powershell = Join-Path ([Environment]::GetFolderPath("System")) "WindowsPowerShell\v1.0\powershell.exe"
        $script:trialRoots = @()

        function copyTree {
            param ([string]$from, [string]$to)
            [void][System.IO.Directory]::CreateDirectory($to)
            foreach ($file in [System.IO.Directory]::GetFiles($from)) {
                [System.IO.File]::Copy($file, [System.IO.Path]::Combine($to, [System.IO.Path]::GetFileName($file)))
            }
            foreach ($dir in [System.IO.Directory]::GetDirectories($from)) {
                copyTree $dir ([System.IO.Path]::Combine($to, [System.IO.Path]::GetFileName($dir)))
            }
        }
    }

    AfterAll {
        # Pester の TestDrive は [ を含む名前を消せないので、TEMP に作って .NET で消す
        foreach ($root in $script:trialRoots) {
            if ([System.IO.Directory]::Exists($root)) { [System.IO.Directory]::Delete($root, $true) }
        }
    }

    It "フォルダ名に <label> を含む場所へ scripts\ を写しても、lib.ps1 を最後まで読める" -ForEach $libCases {
        $root = Join-Path ([System.IO.Path]::GetTempPath()) ("lib_" + [guid]::NewGuid().ToString("N"))
        $script:trialRoots += ,$root
        $copy = Join-Path $root "tool_a${ch}b\scripts"
        copyTree $realScripts $copy

        $psi = New-Object System.Diagnostics.ProcessStartInfo($powershell, '-NoProfile -ExecutionPolicy RemoteSigned -Command "& { $p = $env:TEBUNKO_TEST_LIB; . ([WildcardPattern]::Escape($p.Replace([string][char]96, [string][char]96 + [char]96))); if (Get-Command readSettings -ErrorAction SilentlyContinue) { Write-Output loaded } } 2>&1"')
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.RedirectStandardOutput = $true
        $psi.EnvironmentVariables["TEBUNKO_TEST_LIB"] = "$copy\tebunko\lib.ps1"
        $process = [System.Diagnostics.Process]::Start($psi)
        $outText = ""
        try {
            # エラーは -Command の中で標準出力にまとめてある（標準エラーを別に読むと、パイプがいっぱいになったときに固まるため）
            $outText = $process.StandardOutput.ReadToEnd()
            $process.WaitForExit(60000) | Out-Null
        } finally {
            stopStartedProcess $process
        }
        $outText.Trim() | Should -Be "loaded"
    }

    It "<label> の隣に別のフォルダがあっても、実在する方の部品を読む" -ForEach $siblingCases {
        $root = Join-Path ([System.IO.Path]::GetTempPath()) ("sib_" + [guid]::NewGuid().ToString("N"))
        $script:trialRoots += ,$root
        $target = Join-Path $root "a${ch}b"
        $other = Join-Path $root "a${sibling}b"
        [void][System.IO.Directory]::CreateDirectory($target)
        [void][System.IO.Directory]::CreateDirectory($other)
        [System.IO.File]::WriteAllText("$target\part.ps1", '$script:partOwner = "target"')
        [System.IO.File]::WriteAllText("$other\part.ps1", '$script:partOwner = "other"')

        $script:partOwner = $null
        $dir = $target
        . ([WildcardPattern]::Escape("$dir\part.ps1".Replace([string][char]96, [string][char]96 + [char]96)))
        $script:partOwner | Should -Be "target"
    }
}