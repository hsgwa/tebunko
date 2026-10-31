# 関数・クラスのメソッドの長さ（行数）と入れ子の深さ（段）を数え、上限を超えるものが許可の一覧に無ければ落とす。
# 長い関数・深い入れ子が新しく増えないようにする（今ある長いものは許可の一覧に載せ、長くしない）。
# 数え方・閾値・許可の一覧の決まりは docs/design/testing/unit-checks.md「設計の決まりの検査」。
#
# 本体は読み込まず、scripts/ の下のすべての .ps1 を AST で読むだけ（structure.Tests.ps1・ui_thread_io.Tests.ps1 と同じやり方）。
BeforeAll {
    ${scriptsRoot} = (Resolve-Path "$PSScriptRoot\..\..\scripts").Path

    # 上限。これを超えるものは許可の一覧に載せる
    ${maxLines} = 80
    ${maxDepth} = 4

    # 入れ子の段に数える文（elseif・else は if と同じ段、catch・finally は try と同じ段）
    ${nestingTypes} = @(
        [System.Management.Automation.Language.IfStatementAst],
        [System.Management.Automation.Language.SwitchStatementAst],
        [System.Management.Automation.Language.ForStatementAst],
        [System.Management.Automation.Language.ForEachStatementAst],
        [System.Management.Automation.Language.WhileStatementAst],
        [System.Management.Automation.Language.DoWhileStatementAst],
        [System.Management.Automation.Language.DoUntilStatementAst],
        [System.Management.Automation.Language.TryStatementAst]
    )

    # 許可の一覧: Name（<ファイル名>:<関数名>、メソッドは <ファイル名>:<クラス名>.<メソッド名>）・
    # Lines・Depth（今の長さと段。これを超えたら落ちる）・Reason。80 行以下かつ 4 段以下になったら、一覧から消す。
    # 分けるのは、それぞれの関数を触るときに行う（ここに載せるだけで、分けない）
    ${sizeAllowed} = @(
        @{ Name = "content_index_format.ps1:readContentIndexPlaces"; Lines = 65; Depth = 5; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "content_index_search.ps1:searchContentIndex"; Lines = 159; Depth = 7; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "content_index_search.ps1:searchContentIndexFiles"; Lines = 380; Depth = 11; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "content_index_store.ps1:readContentIndexLines"; Lines = 114; Depth = 6; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "content_index_store.ps1:readContentIndexRange"; Lines = 55; Depth = 5; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "extract_office.ps1:extractDocument"; Lines = 108; Depth = 4; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "extract_office.ps1:extractWorkbook"; Lines = 156; Depth = 5; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "gram_index_store.ps1:getFilesGramValues"; Lines = 91; Depth = 4; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "gram_index_store.ps1:updateGramIndexOf"; Lines = 112; Depth = 4; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "gram_search.ps1:getGramContentIndexFiles"; Lines = 99; Depth = 4; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "gui_main.ps1:startGui"; Lines = 128; Depth = 2; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "index_archive.ps1:exportIndexCore"; Lines = 113; Depth = 4; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "index_archive.ps1:importIndexCore"; Lines = 112; Depth = 5; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "index_archive.ps1:openIndexArchive"; Lines = 43; Depth = 5; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "index_archive_rules.ps1:testIndexArchiveManifest"; Lines = 99; Depth = 2; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "index_detail_view.ps1:getIndexDetailView"; Lines = 89; Depth = 4; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "index_format_store.ps1:removeOldIndexOf"; Lines = 48; Depth = 5; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "indexer_decide.ps1:convertFromIngestJournal"; Lines = 108; Depth = 4; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "indexer_plan.ps1:createTargetList"; Lines = 170; Depth = 3; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "indexer_run.ps1:completeIngestJob"; Lines = 90; Depth = 3; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "indexer_run.ps1:invokeIndexerBody"; Lines = 509; Depth = 5; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "indexer_run.ps1:invokeIngestTask"; Lines = 111; Depth = 3; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "indexing_tab.ps1:finishIndexing"; Lines = 83; Depth = 2; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "indexing_tab.ps1:updateIndexingProgress"; Lines = 90; Depth = 3; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "office_app.ps1:OfficeWatchdog.Start"; Lines = 96; Depth = 5; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "office_embedded.ps1:readEmbeddedPackageLines"; Lines = 43; Depth = 5; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "office_embedded.ps1:readXlsxSharedStrings"; Lines = 52; Depth = 5; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "office_embedded.ps1:readXlsxSheetCellLines"; Lines = 92; Depth = 6; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "office_reader.ps1:readDocxUnitsFromZip"; Lines = 94; Depth = 4; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "office_reader.ps1:readPptxUnitsFromZip"; Lines = 98; Depth = 4; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "office_reader.ps1:readXlsxHeaderFooterLines"; Lines = 64; Depth = 6; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "office_reader.ps1:readXlsxObjectUnits"; Lines = 113; Depth = 5; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "office_reader.ps1:readXlsxShapeRows"; Lines = 85; Depth = 6; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "office_reader.ps1:readXmlLines"; Lines = 259; Depth = 7; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "office_reader.ps1:readZipEntry"; Lines = 101; Depth = 4; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "open_source.ps1:continueFindSourceFile"; Lines = 86; Depth = 3; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "open_source.ps1:openFoundSource"; Lines = 101; Depth = 3; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "open_source.ps1:openInExcel"; Lines = 98; Depth = 4; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "pdf_crypt.ps1:getPdfStreamCryptMethod"; Lines = 29; Depth = 5; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "pdf_file.ps1:readPdfPageText"; Lines = 212; Depth = 7; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "pdf_file.ps1:readPdfToken"; Lines = 139; Depth = 6; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "pdf_file.ps1:readPdfToUnicode"; Lines = 74; Depth = 7; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "pdf_file.ps1:readPdfXrefSection"; Lines = 136; Depth = 5; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "shell.ps1:showConfirm"; Lines = 88; Depth = 3; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "status_file.ps1:readStatusFile"; Lines = 103; Depth = 4; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
        @{ Name = "version.ps1:readDevVersion"; Lines = 107; Depth = 5; Reason = "今ある長い・深い関数。分けるのは、この関数を触る refactor で行う" }
    )

    # 関数・メソッドの本体か（メソッドの本体の FunctionDefinitionAst は、メソッド側で数えるため飛ばす）
    function isOwnedByMember($ast) {
        return ($ast -is [System.Management.Automation.Language.FunctionDefinitionAst]) -and
            ($ast.Parent -is [System.Management.Automation.Language.FunctionMemberAst])
    }

    # 関数・メソッドの AST を求める（入れ子の関数も 1 つとして数える）
    function getFunctionAsts($rootAst) {
        $found = $rootAst.FindAll({
            param ($a)
            ($a -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                -not ($a.Parent -is [System.Management.Automation.Language.FunctionMemberAst])) -or
            ($a -is [System.Management.Automation.Language.FunctionMemberAst])
        }, $true)
        return @($found)
    }

    # 関数・メソッドの名前（<クラス名>.<メソッド名> / <関数名>）
    function getFunctionName($fn) {
        if ($fn -is [System.Management.Automation.Language.FunctionMemberAst]) {
            $cls = $fn.Parent
            return "$($cls.Name).$($fn.Name)"
        }
        return $fn.Name
    }

    # 入れ子の段の最大。関数の本体から数え、入れ子の関数・メソッドの中は別に数える（外の段に足さない）
    function getNestingDepth($fn) {
        $nestingTypes = ${nestingTypes}
        $statements = $fn.FindAll({
            param ($a)
            foreach ($t in $nestingTypes) { if ($a -is $t) { return $true } }
            return $false
        }, $true)
        $max = 0
        foreach ($st in $statements) {
            $depth = 0
            $p = $st
            $owned = $true
            while ($null -ne $p) {
                if (-not [object]::ReferenceEquals($p, $fn) -and (
                        ($p -is [System.Management.Automation.Language.FunctionMemberAst]) -or
                        (($p -is [System.Management.Automation.Language.FunctionDefinitionAst]) -and -not (isOwnedByMember $p)))) {
                    # 途中に別の関数・メソッドがある = この文は内側の関数のもの
                    $owned = $false
                    break
                }
                foreach ($t in $nestingTypes) { if ($p -is $t) { $depth++; break } }
                if ([object]::ReferenceEquals($p, $fn)) { break }
                $p = $p.Parent
            }
            if ($owned -and $depth -gt $max) { $max = $depth }
        }
        return $max
    }

    # 1 つのスクリプト（ファイルのパスか、ソースの文字列）にあるすべての関数・メソッドの { Name; Lines; Depth } を返す。
    # 名前が重なるもの（オーバーロード）は、長いほう・深いほうを 1 つにまとめる
    function measureFunctions([string]$FileName, $rootAst) {
        $byName = @{}
        foreach ($fn in (getFunctionAsts $rootAst)) {
            $name = "${FileName}:$(getFunctionName $fn)"
            $lines = $fn.Extent.EndLineNumber - $fn.Extent.StartLineNumber + 1
            $depth = getNestingDepth $fn
            if ($byName.ContainsKey($name)) {
                $old = $byName[$name]
                $byName[$name] = @{ Name = $name; Lines = [Math]::Max($old.Lines, $lines); Depth = [Math]::Max($old.Depth, $depth) }
            } else {
                $byName[$name] = @{ Name = $name; Lines = $lines; Depth = $depth }
            }
        }
        return @($byName.Values)
    }

    function measureSource([string]$FileName, [string]$Source) {
        $tokens = $null; $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($Source, [ref]$tokens, [ref]$errors)
        return @(measureFunctions $FileName $ast)
    }

    function measureAllScripts() {
        $all = New-Object System.Collections.Generic.List[object]
        foreach ($f in (Get-ChildItem ${scriptsRoot} -Filter "*.ps1" -Recurse)) {
            $tokens = $null; $errors = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$tokens, [ref]$errors)
            foreach ($m in (measureFunctions $f.Name $ast)) { $all.Add($m) }
        }
        # ファイル名だけで引くため、別のフォルダの同じ名前のファイルの同じ名前の関数は、長いほうにまとめる
        $merged = @{}
        foreach ($m in $all) {
            if ($merged.ContainsKey($m.Name)) {
                $old = $merged[$m.Name]
                $merged[$m.Name] = @{ Name = $m.Name; Lines = [Math]::Max($old.Lines, $m.Lines); Depth = [Math]::Max($old.Depth, $m.Depth) }
            } else {
                $merged[$m.Name] = $m
            }
        }
        return @($merged.Values)
    }

    # 上限を超えるものから、許可の一覧で許されていないものを選ぶ。許可の上限を超えたものも入る
    function findOverLimit($measured, $allowed) {
        $result = @()
        foreach ($m in $measured) {
            $over = ($m.Lines -gt ${maxLines}) -or ($m.Depth -gt ${maxDepth})
            if (-not $over) { continue }
            $allow = @($allowed | Where-Object { $_.Name -eq $m.Name })
            if ($allow.Count -eq 0) {
                $result += "$($m.Name)（$($m.Lines) 行・$($m.Depth) 段。許可の一覧に無い）"
            } elseif ($m.Lines -gt $allow[0].Lines -or $m.Depth -gt $allow[0].Depth) {
                $result += "$($m.Name)（$($m.Lines) 行・$($m.Depth) 段。許可した $($allow[0].Lines) 行・$($allow[0].Depth) 段を超えた）"
            }
        }
        return @($result)
    }

    # 許可の一覧に残った、要らない項目（上限以下になった・今のコードに無い）
    function findStaleAllowed($measured, $allowed) {
        $result = @()
        foreach ($a in $allowed) {
            $m = @($measured | Where-Object { $_.Name -eq $a.Name })
            if ($m.Count -eq 0) {
                $result += "$($a.Name)（今のコードに無い）"
            } elseif ($m[0].Lines -le ${maxLines} -and $m[0].Depth -le ${maxDepth}) {
                $result += "$($a.Name)（$($m[0].Lines) 行・$($m[0].Depth) 段に収まった）"
            }
        }
        return @($result)
    }

    ${measuredAll} = @(measureAllScripts)
}

Describe "関数の長さと入れ子の深さ" -Tag Meta {
    It "scripts/ の関数・メソッドが、80 行・4 段以内か、許可の一覧にある（長くしていない）" {
        $over = findOverLimit ${measuredAll} ${sizeAllowed}
        ($over -join ", ") | Should -Be ""
    }

    It "許可の一覧に、収まった・今のコードに無い項目が残っていない（一覧から消す）" {
        $stale = findStaleAllowed ${measuredAll} ${sizeAllowed}
        ($stale -join ", ") | Should -Be ""
    }

    It "許可の一覧の項目が重なっていない" {
        $dup = @(${sizeAllowed} | Group-Object { $_.Name } | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
        ($dup -join ", ") | Should -Be ""
    }

    It "許可の一覧の項目に、理由がある" {
        $noReason = @(${sizeAllowed} | Where-Object { [string]::IsNullOrWhiteSpace($_.Reason) } | ForEach-Object { $_.Name })
        ($noReason -join ", ") | Should -Be ""
    }
}

Describe "関数の長さと入れ子の深さの数え方（小さな例のスクリプト）" -Tag Meta {
    It "関数の行数は、最初の行から最後の行まで（コメント・空行・param を含む）" {
        $src = "function f {`n    param (`$a)`n`n    # c`n    `$a`n}`n"
        $m = @(measureSource "x.ps1" $src)
        $m.Count | Should -Be 1
        $m[0].Name | Should -Be "x.ps1:f"
        $m[0].Lines | Should -Be 6
    }

    It "入れ子の段: <Name>" -TestCases @(
        @{ Name = "文が無ければ 0 段"; Body = '$a = 1'; Expected = 0 }
        @{ Name = "if だけなら 1 段"; Body = 'if ($a) { $b = 1 }'; Expected = 1 }
        @{ Name = "elseif・else は同じ段"; Body = 'if ($a) { 1 } elseif ($b) { 2 } else { 3 }'; Expected = 1 }
        @{ Name = "try の catch・finally は同じ段"; Body = 'try { 1 } catch { 2 } finally { 3 }'; Expected = 1 }
        @{ Name = "if の中の foreach は 2 段"; Body = 'if ($a) { foreach ($x in $a) { 1 } }'; Expected = 2 }
        @{ Name = "switch・for・while・do も数える"; Body = 'switch ($a) { 1 { for ($i=0;$i -lt 1;$i++) { while ($a) { do { 1 } while ($b) } } } }'; Expected = 4 }
        @{ Name = "do-until も数える"; Body = 'do { 1 } until ($a)'; Expected = 1 }
        @{ Name = "スクリプトブロックの中の文も数える"; Body = '$x | ForEach-Object { if ($_) { try { 1 } catch { 2 } } }'; Expected = 2 }
        @{ Name = "5 段"; Body = 'if ($a) { if ($b) { if ($c) { if ($d) { if ($e) { 1 } } } } }'; Expected = 5 }
    ) {
        param ($Name, $Body, $Expected)
        $m = @(measureSource "x.ps1" "function f {`n$Body`n}`n")
        $m[0].Depth | Should -Be $Expected
    }

    It "入れ子の関数は別に数え、外の関数の段に足さない" {
        $src = "function outer {`n    if (`$a) {`n        function inner {`n            if (`$b) { if (`$c) { 1 } }`n        }`n    }`n}`n"
        $m = @(measureSource "x.ps1" $src)
        ($m | Where-Object { $_.Name -eq "x.ps1:outer" }).Depth | Should -Be 1
        ($m | Where-Object { $_.Name -eq "x.ps1:inner" }).Depth | Should -Be 2
    }

    It "クラスのメソッドは 1 つとして数え（本体の関数を二重に数えない）、名前は <クラス名>.<メソッド名>" {
        $src = "class Foo {`n    [void] Bar() {`n        if (`$true) {`n            foreach (`$x in 1) { 1 }`n        }`n    }`n}`n"
        $m = @(measureSource "x.ps1" $src)
        $m.Count | Should -Be 1
        $m[0].Name | Should -Be "x.ps1:Foo.Bar"
        $m[0].Lines | Should -Be 5
        $m[0].Depth | Should -Be 2
    }

    It "メソッドのオーバーロードは、長いほうを 1 つにまとめる" {
        $src = "class Foo {`n    [void] Bar() { 1 }`n    [void] Bar([int]`$a) {`n        1`n        2`n    }`n}`n"
        $m = @(measureSource "x.ps1" $src)
        $m.Count | Should -Be 1
        $m[0].Lines | Should -Be 4
    }

    It "上限の判定: <Name>" -TestCases @(
        @{ Name = "80 行・4 段は収まる"; Lines = 80; Depth = 4; Allowed = @(); Expected = 0 }
        @{ Name = "81 行は許可の一覧に無ければ落ちる"; Lines = 81; Depth = 1; Allowed = @(); Expected = 1 }
        @{ Name = "5 段は許可の一覧に無ければ落ちる"; Lines = 10; Depth = 5; Allowed = @(); Expected = 1 }
        @{ Name = "許可の一覧にあり、上限の中なら通る"; Lines = 100; Depth = 5; Allowed = @(@{ Name = "x.ps1:f"; Lines = 100; Depth = 5 }); Expected = 0 }
        @{ Name = "許可した行数を超えたら落ちる"; Lines = 101; Depth = 5; Allowed = @(@{ Name = "x.ps1:f"; Lines = 100; Depth = 5 }); Expected = 1 }
        @{ Name = "許可した段を超えたら落ちる"; Lines = 100; Depth = 6; Allowed = @(@{ Name = "x.ps1:f"; Lines = 100; Depth = 5 }); Expected = 1 }
    ) {
        param ($Name, $Lines, $Depth, $Allowed, $Expected)
        $measured = @(@{ Name = "x.ps1:f"; Lines = $Lines; Depth = $Depth })
        @(findOverLimit $measured $Allowed).Count | Should -Be $Expected
    }

    It "許可の一覧の残り: <Name>" -TestCases @(
        @{ Name = "収まったものは残りとして見つかる"; Lines = 80; Depth = 4; Expected = 1 }
        @{ Name = "まだ超えているものは残りではない"; Lines = 90; Depth = 1; Expected = 0 }
        @{ Name = "段だけ超えているものも残りではない"; Lines = 10; Depth = 5; Expected = 0 }
    ) {
        param ($Name, $Lines, $Depth, $Expected)
        $allowed = @(@{ Name = "x.ps1:f"; Lines = 100; Depth = 5 })
        @(findStaleAllowed @(@{ Name = "x.ps1:f"; Lines = $Lines; Depth = $Depth }) $allowed).Count | Should -Be $Expected
    }

    It "許可の一覧にあるのに、今のコードに無いものは残りとして見つかる" {
        @(findStaleAllowed @() @(@{ Name = "x.ps1:gone"; Lines = 100; Depth = 5 })).Count | Should -Be 1
    }
}
