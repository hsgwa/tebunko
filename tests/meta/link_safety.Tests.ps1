# フォルダの再帰の列挙を、許した関数の一覧に限る。
# ジャンクション・シンボリックリンクを辿る列挙（Get-ChildItem -Recurse・SearchOption.AllDirectories）が、取り込み対象の外の
# ファイルを読む・書く・消す・zip に入れる原因になるため、リンクを辿らない自前の再帰（fs.ps1 の getFilesNoFollow）を使う。
# 列挙を足すときは、リンクを辿っても読むだけで安全な理由を確かめ、下の一覧に理由つきで足す。
#
# 見つけるもの（scripts の下の全ファイル。判定の単位は関数。関数の外は Function を空にする）:
#   1. Get-ChildItem（gci・ls・dir）の -Recurse・-Depth（パラメーター名の前方一致・@splat を含む）
#   2. SearchOption の型の使い方のうち [SearchOption]::TopDirectoryOnly の直接の参照以外（AllDirectories・変数・キャスト・型指定）
#   3. GetFiles・EnumerateFiles・GetDirectories などの呼び出しで、探す範囲の引数が [SearchOption]::TopDirectoryOnly でないもの
#   4. EnumerationOptions の RecurseSubdirectories
#   5. フォルダの下をまとめて扱う Copy-Item -Recurse（別名 copy・cp・cpi とスプラットも見る）・Compress-Archive・robocopy・xcopy・ZipFile.CreateFromDirectory（コマンド名・メソッド名で見る）
# 見つけないもの（人が見る）: 自前の再帰（関数が自分を呼ぶ・Stack で降りる）。これらは、フォルダの子を 1 段ずつ列挙するとき、
# リンクを辿るかを自分で決める。列挙には fs.ps1 の getFilesNoFollow を使う。
# 削除の検査は tests/meta/remove_entry.Tests.ps1 が受け持つため、ここでは調べない。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"

    # 許した関数: File（scripts からの相対パス。/ 区切り）・Function（関数・メソッドの名前。関数の外は空）・Reason（リンクを辿っても安全な理由）
    ${linkSafetyAllowed} = @(
        @{ File = "tebunko/gui.ps1"; Function = ""; Reason = "起動時に配布物のフォルダのファイルのブロックを外す（ファイルの属性の書き換え）。配布物のフォルダの中にリンクを置かない前提で、取り込み対象のファイルには届かない" }
        @{ File = "tebunko/core/workspace.ps1"; Function = "getLegacyIndexState"; Reason = "前の版のインデックスがあるかを調べる読み取りだけ。ファイルの有無しか見ない" }
        @{ File = "tebunko/index/content_index_store.ps1"; Function = "findContentIndexFiles"; Reason = "本文インデックスのファイルを探す読み取りだけ（インデックスの置き場の中を探し、取り込み対象のファイルは見ない）" }
        @{ File = "tebunko/index/content_index_store.ps1"; Function = "findIndexFoldersWithBooks"; Reason = "インデックスの置き場の中で、元のファイルごとの TSV のフォルダを探す。結果は本文インデックスのファイルへのまとめ直し（書き込み）に渡るが、書くのは置き場の中のフォルダで、置き場にリンクを置かない前提" }
        @{ File = "tebunko/index/index_store.ps1"; Function = "getIndexTsvCounts"; Reason = "TSV の数を数える読み取りだけ（インデックスの置き場の中）" }
        @{ File = "tebunko/search/search_run.ps1"; Function = "testIndexExists"; Reason = "インデックスの有無を調べる読み取りだけ" }
        @{ File = "tebunko/search/search_run.ps1"; Function = "getIndexSummary"; Reason = "インデックスの件数を数える読み取りだけ" }
    )

    ${enumerationMethods} = @(
        "GetFiles", "EnumerateFiles", "GetDirectories", "EnumerateDirectories",
        "GetFileSystemEntries", "EnumerateFileSystemEntries", "GetFileSystemInfos", "EnumerateFileSystemInfos"
    )

    function testParameterPrefix {
        # name が full の前方一致（大文字小文字を区別しない）で、minLength 文字以上のとき $true
        param ([string]$name, [string]$full, [int]$minLength)

        return ($name.Length -ge $minLength -and $name.Length -le $full.Length -and $full.StartsWith($name, [System.StringComparison]::OrdinalIgnoreCase))
    }

    function testTopDirectoryOnly {
        # [SearchOption]::TopDirectoryOnly（型名は SearchOption で終わるもの）の直接の参照なら $true
        param ($ast)

        if ($ast -isnot [System.Management.Automation.Language.MemberExpressionAst]) { return $false }
        if (!$ast.Static -or $ast.Member -isnot [System.Management.Automation.Language.StringConstantExpressionAst]) { return $false }
        if ($ast.Expression -isnot [System.Management.Automation.Language.TypeExpressionAst]) { return $false }
        return ($ast.Member.Value -eq "TopDirectoryOnly" -and $ast.Expression.TypeName.Name -match "(^|\.)SearchOption$")
    }

    function getEnclosingFunctionName {
        param ($ast)

        $current = $ast.Parent
        while ($null -ne $current) {
            if ($current -is [System.Management.Automation.Language.FunctionDefinitionAst] -or $current -is [System.Management.Automation.Language.FunctionMemberAst]) {
                return $current.Name
            }
            $current = $current.Parent
        }
        return ""
    }

    function findLinkFollowingAsts {
        # ast の下で、リンクを辿りうる列挙を @{ Line; Function; Kind } の配列で返す
        param ($ast)

        $found = New-Object System.Collections.Generic.List[object]
        $nodes = $ast.FindAll({ param ($n) $true }, $true)
        foreach ($node in $nodes) {
            $kind = $null
            if ($node -is [System.Management.Automation.Language.CommandAst]) {
                # モジュール付きの名前（Microsoft.PowerShell.Management\Get-ChildItem）は、最後の \ か / の後ろだけで比べる
                $commandName = $node.GetCommandName()
                if ($commandName) {
                    $commandName = $commandName.Substring($commandName.LastIndexOfAny([char[]]@("\", "/")) + 1)
                }
                if ($commandName -and $commandName -match "^(Get-ChildItem|gci|ls|dir)$") {
                    foreach ($element in $node.CommandElements) {
                        if ($element -is [System.Management.Automation.Language.CommandParameterAst]) {
                            # -s は Recurse の別名
                            if ($element.ParameterName -ieq "s" -or (testParameterPrefix $element.ParameterName "Recurse" 1) -or (testParameterPrefix $element.ParameterName "Depth" 2)) {
                                $kind = "Get-ChildItem -$($element.ParameterName)"
                            }
                        } elseif ($element -is [System.Management.Automation.Language.VariableExpressionAst] -and $element.Splatted) {
                            $kind = "Get-ChildItem @$($element.VariablePath.UserPath)（スプラット）"
                        }
                    }
                } elseif ($commandName -and $commandName -match "^(Compress-Archive|robocopy(\.exe)?|xcopy(\.exe)?)$") {
                    $kind = "$commandName（フォルダの下をまとめて扱う）"
                } elseif ($commandName -and $commandName -match "^(Copy-Item|copy|cp|cpi)$") {
                    foreach ($element in $node.CommandElements) {
                        if ($element -is [System.Management.Automation.Language.CommandParameterAst]) {
                            if (testParameterPrefix $element.ParameterName "Recurse" 1) {
                                $kind = "Copy-Item -$($element.ParameterName)"
                            }
                        } elseif ($element -is [System.Management.Automation.Language.VariableExpressionAst] -and $element.Splatted) {
                            $kind = "Copy-Item @$($element.VariablePath.UserPath)（スプラット）"
                        }
                    }
                }
            } elseif ($node -is [System.Management.Automation.Language.TypeExpressionAst]) {
                if ($node.TypeName.Name -match "(^|\.)SearchOption$" -and !(testTopDirectoryOnly $node.Parent)) {
                    $kind = "SearchOption（TopDirectoryOnly の直接の参照以外）"
                } elseif ($node.TypeName.Name -match "(^|\.)EnumerationOptions$") {
                    $kind = "EnumerationOptions"
                }
            } elseif ($node -is [System.Management.Automation.Language.ConvertExpressionAst] -or $node -is [System.Management.Automation.Language.TypeConstraintAst]) {
                if ($node.TypeName.Name -match "(^|\.)(SearchOption|EnumerationOptions)$") {
                    $kind = "$($node.TypeName.Name) のキャスト・型指定"
                }
            } elseif ($node -is [System.Management.Automation.Language.MemberExpressionAst] -and $node -isnot [System.Management.Automation.Language.InvokeMemberExpressionAst]) {
                if ($node.Member -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $node.Member.Value -eq "RecurseSubdirectories") {
                    $kind = "RecurseSubdirectories"
                }
            } elseif ($node -is [System.Management.Automation.Language.InvokeMemberExpressionAst]) {
                if ($node.Member -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $node.Member.Value -eq "CreateFromDirectory") {
                    $kind = "CreateFromDirectory（フォルダの下をまとめて zip に入れる）"
                } elseif ($node.Member -is [System.Management.Automation.Language.StringConstantExpressionAst] -and ${enumerationMethods} -contains $node.Member.Value) {
                    $arguments = @($node.Arguments)
                    # 静的メソッド（Directory.GetFiles(path, pattern, option)）は 3 番目、インスタンス（DirectoryInfo.GetFiles(pattern, option)）は 2 番目が探す範囲
                    $optionIndex = if ($node.Static) { 2 } else { 1 }
                    if ($arguments.Count -gt $optionIndex -and !(testTopDirectoryOnly $arguments[$optionIndex])) {
                        $kind = "$($node.Member.Value) の探す範囲が TopDirectoryOnly の直接の参照でない"
                    }
                }
            }
            if ($kind) {
                $found.Add(@{ Line = $node.Extent.StartLineNumber; Function = (getEnclosingFunctionName $node); Kind = $kind })
            }
        }
        return $found.ToArray()
    }

    function findLinkFollowingInScripts {
        # scripts の下の全ファイルで、リンクを辿りうる列挙を @{ File; Function; Line; Kind } の配列で返す
        $found = New-Object System.Collections.Generic.List[object]
        foreach ($file in @(Get-ChildItem -LiteralPath ${scriptsDir} -Filter "*.ps1" -Recurse)) {
            $errors = $null
            $tokens = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
            foreach ($item in @(findLinkFollowingAsts $ast)) {
                $relative = $file.FullName.Substring(([System.IO.Path]::GetFullPath(${scriptsDir}).TrimEnd("\").Length + 1)).Replace("\", "/")
                $found.Add(@{ File = $relative; Function = $item.Function; Line = $item.Line; Kind = $item.Kind })
            }
        }
        return $found.ToArray()
    }

    function findLinkFollowingInText {
        param ([string]$text)

        $errors = $null
        $tokens = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($text, [ref]$tokens, [ref]$errors)
        return @(findLinkFollowingAsts $ast)
    }
}

Describe "リンクを辿りうるフォルダの列挙" -Tag Meta {
    It "許した一覧に無い関数が、再帰の列挙をしていない" {
        $unknown = @(findLinkFollowingInScripts | Where-Object {
                $item = $_
                @(${linkSafetyAllowed} | Where-Object { $_.File -eq $item.File -and $_.Function -eq $item.Function }).Count -eq 0
            } | ForEach-Object { "$($_.File) › $($_.Function) 行 $($_.Line): $($_.Kind)" })
        $unknown | Should -BeNullOrEmpty -Because "再帰の列挙はリンクを辿る。fs.ps1 の getFilesNoFollow を使うか、読むだけで安全な理由を tests/meta/link_safety.Tests.ps1 の一覧に足す"
    }

    It "許した一覧の関数は、いまも存在し、再帰の列挙をしている（消した・書き換えた関数を一覧に残さない）" {
        $current = @(findLinkFollowingInScripts)
        $stale = @(${linkSafetyAllowed} | Where-Object {
                $item = $_
                @($current | Where-Object { $_.File -eq $item.File -and $_.Function -eq $item.Function }).Count -eq 0
            } | ForEach-Object { "$($_.File) › $($_.Function)" })
        $stale | Should -BeNullOrEmpty
    }

    It "許した一覧のどの項目にも理由がある" {
        @(${linkSafetyAllowed} | Where-Object { [string]::IsNullOrWhiteSpace($_.Reason) }).Count | Should -Be 0
    }
}

Describe "再帰の列挙の見つけ方（回避の見本）" -Tag Meta {
    It "見つける: <name>" -TestCases @(
        @{ name = "Get-ChildItem -Recurse"; text = 'Get-ChildItem -LiteralPath $d -Recurse' }
        @{ name = "別名 gci"; text = 'gci $d -Recurse' }
        @{ name = "別名 ls"; text = 'ls $d -Recurse' }
        @{ name = "別名 dir"; text = 'dir $d -Recurse' }
        @{ name = "省略形 -Rec"; text = 'Get-ChildItem $d -Rec' }
        @{ name = "省略形 -r"; text = 'Get-ChildItem $d -r' }
        @{ name = "大文字小文字違い"; text = 'GET-CHILDITEM $d -RECURSE' }
        @{ name = "Recurse の別名 -s"; text = 'Get-ChildItem $d -s' }
        @{ name = "モジュール付きの名前"; text = 'Microsoft.PowerShell.Management\Get-ChildItem $d -Recurse' }
        @{ name = "Copy-Item -Recurse"; text = 'Copy-Item -LiteralPath $a -Destination $b -Recurse' }
        @{ name = "Copy-Item の別名 cp"; text = 'cp $a $b -Recurse' }
        @{ name = "Copy-Item の別名 copy"; text = 'copy $a $b -r' }
        @{ name = "Copy-Item の別名 cpi"; text = 'cpi $a $b -Recurse' }
        @{ name = "Copy-Item のスプラット"; text = '$o = @{ Recurse = $true }; Copy-Item $a $b @o' }
        @{ name = "/ 区切りのモジュール付きの名前"; text = 'Microsoft.PowerShell.Management/Get-ChildItem $d -Recurse' }
        @{ name = "Compress-Archive"; text = 'Compress-Archive -Path $d -DestinationPath $z' }
        @{ name = "robocopy"; text = 'robocopy $a $b /E' }
        @{ name = "xcopy"; text = 'xcopy $a $b /s' }
        @{ name = "ZipFile.CreateFromDirectory"; text = '[System.IO.Compression.ZipFile]::CreateFromDirectory($d, $z)' }
        @{ name = "-Depth"; text = 'Get-ChildItem $d -Depth 3' }
        @{ name = "-Depth の省略形"; text = 'Get-ChildItem $d -Dep 3' }
        @{ name = "スプラット"; text = '$a = @{ Recurse = $true }; Get-ChildItem $d @a' }
        @{ name = "AllDirectories の直接の参照"; text = '[System.IO.Directory]::GetFiles($d, "*", [System.IO.SearchOption]::AllDirectories)' }
        @{ name = "短い型名の AllDirectories"; text = '$x = [SearchOption]::AllDirectories' }
        @{ name = "SearchOption を変数で渡す（静的メソッド）"; text = '[System.IO.Directory]::EnumerateFiles($d, "*", $option)' }
        @{ name = "SearchOption を変数で渡す（インスタンス）"; text = '(New-Object System.IO.DirectoryInfo $d).GetFiles("*", $option)' }
        @{ name = "文字列で渡す"; text = '[System.IO.Directory]::GetDirectories($d, "*", "AllDirectories")' }
        @{ name = "数で渡す"; text = '[System.IO.Directory]::GetFileSystemEntries($d, "*", 1)' }
        @{ name = "SearchOption へのキャスト"; text = '$o = [System.IO.SearchOption]"AllDirectories"' }
        @{ name = "SearchOption の型指定"; text = 'function f { param ([System.IO.SearchOption]$o) }' }
        @{ name = "EnumerationOptions"; text = '$o = New-Object System.IO.EnumerationOptions; $o.RecurseSubdirectories = $true' }
        @{ name = "式の中の AllDirectories"; text = '$x = if ($r) { [System.IO.SearchOption]::AllDirectories } else { [System.IO.SearchOption]::TopDirectoryOnly }' }
    ) {
        param ($name, $text)
        @(findLinkFollowingInText $text).Count | Should -BeGreaterThan 0
    }

    It "見つけない: <name>" -TestCases @(
        @{ name = "再帰でない Get-ChildItem"; text = 'Get-ChildItem -LiteralPath $d -File' }
        @{ name = "再帰でない Copy-Item"; text = 'Copy-Item -LiteralPath $a -Destination $b -Force' }
        @{ name = "再帰でない cp"; text = 'cp $a $b -Force' }
        @{ name = "ConvertTo-Json の -Depth"; text = 'ConvertTo-Json -InputObject $x -Depth 5' }
        @{ name = "TopDirectoryOnly（静的メソッド）"; text = '[System.IO.Directory]::GetFiles($d, "*.tsv", [System.IO.SearchOption]::TopDirectoryOnly)' }
        @{ name = "TopDirectoryOnly（インスタンス）"; text = '$info.EnumerateFiles("*", [System.IO.SearchOption]::TopDirectoryOnly)' }
        @{ name = "探す範囲を書かない列挙"; text = '[System.IO.Directory]::GetFiles($d, $pattern)' }
        @{ name = "パターンだけのインスタンスの列挙"; text = '$info.GetFiles("*.tsv")' }
        @{ name = "名前に SearchOption を含む関数"; text = 'function readSearchOption { }; readSearchOption' }
    ) {
        param ($name, $text)
        @(findLinkFollowingInText $text).Count | Should -Be 0
    }

    It "関数の中で見つけたときは、関数の名前を返す" {
        $found = @(findLinkFollowingInText 'function walk { Get-ChildItem $d -Recurse }')
        $found.Count | Should -Be 1
        $found[0].Function | Should -Be "walk"
    }

    It "関数の外で見つけたときは、関数の名前を空にする" {
        $found = @(findLinkFollowingInText 'Get-ChildItem $d -Recurse')
        $found[0].Function | Should -Be ""
    }
}
