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

    It "M2(ii) の例: 試作で漏れた形（. `$libPath・. `$x.Path）は読み込み口の行として検出しない" {
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

    It "gui.ps1 の最上位は、startup_error.ps1 の読み込みと 1 つの try/catch だけ" {
        $ast = $parsedAsts["$singleScriptsDir\tebunko\gui.ps1"]
        $top = @($ast.EndBlock.Statements)
        $top.Count | Should -Be 2
        (commandOf $top[0]) | Should -BeOfType [System.Management.Automation.Language.CommandAst]
        (isLoaderLine (commandOf $top[0])) | Should -Be $true
        $top[1] | Should -BeOfType [System.Management.Automation.Language.TryStatementAst]

        $try = $top[1]
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
        foreach ($name in @("Tabs", "IndexTab", "SearchTab", "SettingsTab", "KillTab", "IndexTabHeader", "KillTabHeader", "StatusText", "MoreButton", "AboutMenuItem")) {
            $names -contains $name | Should -Be $true
        }
    }

    It "<file> に、gui.ps1 が使う名前がすべてある" -ForEach @(
        @{ File = "tab_index.xaml"; Marker = 'Tab = "IndexTab"' }
        @{ File = "tab_search.xaml"; Marker = 'Tab = "SearchTab"' }
        @{ File = "tab_settings.xaml"; Marker = 'Tab = "SettingsTab"' }
        @{ File = "tab_kill.xaml"; Marker = 'Tab = "KillTab"' }
    ) {
        # gui.ps1 の $tabs から、そのタブの名前の一覧を取り出す
        $start = $gui.IndexOf($Marker)
        $start | Should -Not -Be -1
        $listStart = $gui.IndexOf("Names = @(", $start)
        $listEnd = $gui.IndexOf(") }", $listStart)
        $list = $gui.Substring($listStart, $listEnd - $listStart)
        $wanted = @([regex]::Matches($list, '"([A-Za-z]+)"') | ForEach-Object { $_.Groups[1].Value })
        $wanted.Count -gt 0 | Should -Be $true

        $names = getXamlNames "$here\..\scripts\tebunko\xaml\$File"
        $missing = @($wanted | Where-Object { $names -notcontains $_ })
        ($missing -join ", ") | Should -Be ""
    }
}