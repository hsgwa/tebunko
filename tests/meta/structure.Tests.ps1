# スクリプトの構成（パス定義・構文・XAML・型の読み込み）のテスト
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
}

Describe "パス定義" -Tag Meta {
    It "リポジトリ直下を基準にする（書き込めるため、設定ファイルもリポジトリ直下に置く）" {
        $rootDir | Should -Be (Resolve-Path "$here\..").Path
        $dataDir | Should -Be $rootDir
        $settingsFile | Should -Be "$rootDir\setting.config"
    }

    It "work の中身は work の置き場所（既定はリポジトリ直下の work。setting.config の workspaceFolder で変わる）を基準にする" {
        $workspace.Dir | Should -Be (getWorkDir $settingsFile)
        $workspace.IndexDir | Should -Be "$($workspace.Dir)\content_index"
        $workspace.LegacyIndexDir | Should -Be "$($workspace.Dir)\index"
        $workspace.TmpRoot | Should -Be "$($workspace.Dir)\tmp"
        $workspace.PublishDir | Should -Be "$($workspace.Dir)\publish\$PID"
        $workspace.ResultFile | Should -Be "$($workspace.Dir)\search_results.txt"
    }
}

Describe "画面の部品でのパスの組み立て" -Tag Meta {
    # ui\ 配下のファイルは gui.ps1 から dot-source する部品。中で $PSScriptRoot を使うと ui\ を指すため、
    # "${PSScriptRoot}\tebunko\indexer.ps1" のように起動口からの相対パスを書くと存在しないパスになる
    # （［インデックス作成を開始］でインデクサが起動しなかった不具合）。パスは起動口（gui.ps1）で決めて変数で渡す
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
}

Describe "画面定義（XAML）" -Tag Meta {
    BeforeDiscovery {
        $xamlFiles = @(Get-ChildItem "$PSScriptRoot\..\..\scripts" -Recurse -Filter "*.xaml" |
            ForEach-Object { @{ Name = $_.Name; FullName = $_.FullName } })
    }

    It "<name> が XML として読める" -ForEach $xamlFiles {
        { [xml](Get-Content $FullName -Raw -Encoding UTF8) } | Should -Not -Throw
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
        @{ Key = "Bg.Section"; Color = "#E8EBF0"; Opacity = 1 }
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
}

Describe "theme の文字の Style（大きさ・太さ・行の高さは Figma の設計どおり）" -Tag Meta {
    BeforeDiscovery {
        # Weight が空は Normal。Figma の Medium は WPF の SemiBold に読み替えてある。LineHeight が空は指定しない
        $textStyles = @(
            @{ Key = "Micro";        Size = "10"; Weight = "";         LineHeight = "" }
            @{ Key = "ColumnHeader"; Size = "10"; Weight = "Bold";     LineHeight = "" }
            @{ Key = "Meta";         Size = "11"; Weight = "";         LineHeight = "" }
            @{ Key = "Meta.Tall";    Size = "";   Weight = "";         LineHeight = "16" }
            @{ Key = "Meta.Strong";  Size = "";   Weight = "Bold";     LineHeight = "" }
            @{ Key = "Meta.Key";     Size = "";   Weight = "SemiBold"; LineHeight = "" }   # 行の高さ 16 は Meta.Tall から継ぐ
            @{ Key = "Note";         Size = "11"; Weight = "";         LineHeight = "17.6" }
            @{ Key = "Chip";         Size = "11"; Weight = "SemiBold"; LineHeight = "" }
            @{ Key = "Link.Text";    Size = "11"; Weight = "SemiBold"; LineHeight = "" }
            @{ Key = "Cell";         Size = "12"; Weight = "";         LineHeight = "" }
            @{ Key = "Cell.Key";     Size = "";   Weight = "SemiBold"; LineHeight = "" }
            @{ Key = "Label";        Size = "12"; Weight = "SemiBold"; LineHeight = "" }
            @{ Key = "Label.Strong"; Size = "";   Weight = "Bold";     LineHeight = "" }
            @{ Key = "Brand";        Size = "12"; Weight = "SemiBold"; LineHeight = "" }
            @{ Key = "Body";         Size = "13"; Weight = "";         LineHeight = "" }
            @{ Key = "Body.Strong";  Size = "";   Weight = "Bold";     LineHeight = "" }
            @{ Key = "Nav";          Size = "13"; Weight = "SemiBold"; LineHeight = "" }
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
            "BadgeGlyphInfo", "BadgeGlyphWarn", "BadgeGlyphError", "BadgeGlyphOk", "Dots3"
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
            ". `"$scripts\shared\ui\types.ps1`""
            ". `"$scripts\tebunko\ui\types.ps1`""
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

        function getRelPath {
            param ([string]$fullName)
            return $fullName.Substring($singleScriptsDir.Length + 1)
        }

        function isLoaderLine {
            # 読み込み口の行: . "$PSScriptRoot\..." や . "$here\..." の形（$PSScriptRoot・固定文字列の組み立て）。
            # ここでは「読み込み口が使ってよい形」として、$PSScriptRoot を使った文字列展開のドットソースだけを対象にする
            param ($commandAst)
            if ($commandAst.InvocationOperator -ne "Dot") { return $false }
            if ($commandAst.CommandElements.Count -lt 1) { return $false }
            $first = $commandAst.CommandElements[0]
            return $first -is [System.Management.Automation.Language.ExpandableStringExpressionAst] -or
                $first -is [System.Management.Automation.Language.StringConstantExpressionAst]
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

    It "M2(ii) の例: 変数やプロパティを渡す形（. `$libPath・. `$x.Path）は読み込み口の行として検出しない" {
        $badSamples = @('. $libPath', '. $x.Path')
        foreach ($sample in $badSamples) {
            $fixtureAst = [System.Management.Automation.Language.Parser]::ParseInput($sample, [ref]$null, [ref]$null)
            $cmd = @($fixtureAst.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true))[0]
            (isLoaderLine $cmd) | Should -Be $false -Because "形: $sample"
        }
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

    It "gui.ps1 の最上位は、起動の失敗の知らせ（startup_error_view.ps1・startup_error.ps1）の読み込みと 1 つの try/catch だけ" {
        $ast = $parsedAsts["$singleScriptsDir\tebunko\gui.ps1"]
        $top = @($ast.EndBlock.Statements)
        $top.Count | Should -Be 3
        foreach ($loader in $top[0..1]) {
            (commandOf $loader) | Should -BeOfType [System.Management.Automation.Language.CommandAst]
            (isLoaderLine (commandOf $loader)) | Should -Be $true
        }
        $top[2] | Should -BeOfType [System.Management.Automation.Language.TryStatementAst]

        $try = $top[2]
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
            if ($relFile -eq "tools/new_single_script.ps1" -or $relFile -like "tests/tools/new_single_script*" -or $relFile -eq "tests/meta/structure.Tests.ps1") { continue }
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
    # gui_main.ps1（startGui）が FindName で取る名前が、XAML に実在すること。
    # タブの中身を別ファイルに分けているため、名前を足したり動かしたりすると気づきにくい
    BeforeAll {
        $xamlNs = "http://schemas.microsoft.com/winfx/2006/xaml"
        $gui = [System.IO.File]::ReadAllText("$here\..\scripts\tebunko\ui\gui_main.ps1")

        function getXamlNames {
            param ([string]$path)
            [xml]$xaml = Get-Content $path -Raw -Encoding UTF8
            return @($xaml.SelectNodes("//*") | ForEach-Object { $_.GetAttribute("Name", $xamlNs) } | Where-Object { $_ -ne "" })
        }
    }

    It "ウィンドウの枠の名前がある" {
        $names = getXamlNames "$here\..\scripts\tebunko\xaml\tebunko.xaml"
        foreach ($name in @("NavHost", "ContentHost", "StatusBarHost")) {
            $names -contains $name | Should -Be $true
        }
    }

    It "ナビ・ステータスバーに、画面の切り替えと他の画面が使う名前がある" {
        $nav = getXamlNames "$here\..\scripts\tebunko\xaml\shell\nav.xaml"
        foreach ($name in @("NavList", "SearchTab", "IndexTab", "SettingsTab", "KillTab", "IndexTabBadge", "KillTabBadge", "AboutLink")) {
            $nav -contains $name | Should -Be $true
        }
        (getXamlNames "$here\..\scripts\tebunko\xaml\shell\status_bar.xaml") -contains "StatusText" | Should -Be $true
    }

    It "gui_main.ps1 の領域の表の各ファイルに、使う名前がすべてある" {
        # $regions の `File = "..."; ... Names = @(...)` を全部取り出して、XAML と突き合わせる
        $entries = @([regex]::Matches($gui, 'File\s*=\s*"([^"]+)".*?Names\s*=\s*@\(([^)]*)\)'))
        $entries.Count -gt 0 | Should -Be $true
        $problems = New-Object System.Collections.Generic.List[string]
        foreach ($entry in $entries) {
            $file = $entry.Groups[1].Value
            $wanted = @([regex]::Matches($entry.Groups[2].Value, '"([A-Za-z]+)"') | ForEach-Object { $_.Groups[1].Value })
            $path = "$here\..\scripts\tebunko\xaml\$file"
            if (-not (Test-Path -LiteralPath $path)) { $problems.Add("$file が無い"); continue }
            $names = getXamlNames $path
            foreach ($w in $wanted) {
                if ($names -notcontains $w) { $problems.Add("$file に $w が無い") }
            }
        }
        ($problems -join ", ") | Should -Be ""
    }
}