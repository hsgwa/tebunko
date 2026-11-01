# 利用者に出すエラー・警告の文言が、設計書の一覧に載っているかを確かめる。
# scripts/ の .ps1 を AST で読み、次の 3 種類の日本語の文字列を取り出して、設計書の一覧と突き合わせる。
#   (A) throw の中の文字列
#   (B) writeIndexerLog の 1 つ目の引数のうち、2 つ目の引数が Yellow か Red のもの（警告・失敗。色なし・Green は進み具合の記録なので見ない）
#   (C) describeIngestError（indexer_state.ps1）の中の文字列（失敗の原因の言い換え）
# 変数・関数で渡す文言も追う（同じ関数の中 → 同じファイルの直下 → ほかのファイルの直下の代入。scripts/ の関数の呼び出しは、その関数の文字列）。
# 追えない呼び出しは「追えない呼び出しの一覧」に、載せない文言は「除外の一覧」に理由を書く。
# 突き合わせる一覧と決まりは docs/design/testing/unit-checks.md「設計の決まりの検査」。
#
# 本体は読み込まず、AST で読むだけ（safety.Tests.ps1・ui_thread_io.Tests.ps1 と同じやり方）。
BeforeAll {
    ${scriptsRoot} = (Resolve-Path "$PSScriptRoot\..\..\scripts").Path
    ${docsRoot}    = (Resolve-Path "$PSScriptRoot\..\..\docs\design").Path

    # 埋め込み（$x・${x}・$(…)・-f の {0}）の目印。突き合わせでは「1 文字以上の何か」になる
    ${mark} = [string][char]1

    # 突き合わせる一覧（どれか 1 行に一致すれば載っているとする）
    function getListDocPaths() {
        $paths = @(Get-ChildItem "$docsRoot\indexing" -Filter "log-warnings*.md" | ForEach-Object { $_.FullName })
        $paths += @(Get-ChildItem "$docsRoot\indexing" -Filter "errors*.md" | ForEach-Object { $_.FullName })
        $paths += @(Get-ChildItem "$docsRoot\gui" -Filter "messages*.md" | ForEach-Object { $_.FullName })
        return @($paths | Where-Object { Test-Path -LiteralPath $_ })
    }

    # 除外の一覧: Text（置き換えた形。埋め込みは {}）か Pattern（正規表現。置き換えた形に当てる）と Reason。
    # 利用者に届かない文言、または一覧に別の形でまとめて載せている文言だけを書く。
    # 場所（ファイル・関数・行）では引かない（refactor で動いても崩れないため）
    ${messageExcluded} = @(
        @{ Pattern = '^PDF.+読めません。$'; Reason = "PDF の読み取りの失敗の文言。errors.md の「PDF で始まる文言」の行にまとめて載せている" }
        @{ Text = "Windows のあるドライブを取れない"; Reason = "開発者向けの内部エラー（Office の入っているかを調べる処理が、環境の異常で止まる。利用者の画面・ログに出る前に別の文言になる）" }
        @{ Text = "newAppMutex: ツールの名前（app）が空です"; Reason = "開発者向けの内部エラー（呼び出し側の書き間違いを知らせる）" }
        @{ Text = "joinWorkerScript: つなぐ本体の構文エラーです: {}"; Reason = "開発者向けの内部エラー（単一 .ps1 版を作るときの確かめ）" }
        @{ Text = "WorkerPool は閉じています。"; Reason = "開発者向けの内部エラー（閉じたあとの呼び出しを知らせる）" }
        @{ Text = "(埋め込みから出す文字)"; Reason = "メッセージではない（サイズの上限のログに、埋め込む値として出る部品の名前 PartName。ログの行は log-warnings.md に載せている）" }
        @{ Text = "{} が起動しているため、更新に使用できません"; Reason = "内部の合図（利用者の Office アプリが起動しているため、そのファイルを次回に回す。取り込み一覧にもログにも出ない）" }
        @{ Pattern = '^インデックスのファイル「\{\}」を読み込めません（大きさが上限（\{\} MB）を超えています）。エクスポートした側で'; Reason = "変数をたどると、大きさの原因と読み取りの原因が組み合わさって出るが、実際には出ない組み合わせ（大きさで断る文言と読み取りで断る文言は、それぞれ gui/messages-index-ops.md に載せている）" }
        @{ Text = "読んだ量が上限を超えました"; Reason = "読み取りの途中で投げ、呼び出し側が「ファイルサイズが大きすぎるため更新できません。」などの文言に置き換える" }
        @{ Text = "不明な文字コードです（{}）。"; Reason = "開発者向けの内部エラー（文字コードの名前の書き間違いを知らせる）" }
        @{ Text = "getGuiErrorLogHeader: setGuiErrorLogSource で登録されていない"; Reason = "開発者向けの内部エラー（起動の順番の間違いを知らせる）" }
        @{ Text = "ensureNetworkQueue: setNetworkQueueFactory が呼ばれていない"; Reason = "開発者向けの内部エラー（起動の順番の間違いを知らせる）" }
        @{ Text = "単一 .ps1 版に {} の部品が入っていません。"; Reason = "開発者向けの内部エラー（単一 .ps1 版のビルドの確かめ）" }
        @{ Text = "このファイルの更新には Word・PowerPoint が要ります。"; Reason = "内部の合図（Office を使わない読み取りから、Office を使う読み取りへ切り替える）。取り込みの指揮が受けて続け、利用者には出ない" }
        @{ Text = "IngestPlanner: kinds に {} がありません"; Reason = "開発者向けの内部エラー（部品の渡し忘れを知らせる）" }
        @{ Text = "IngestPlanner: Crawl の前に AssignNames を呼んでください"; Reason = "開発者向けの内部エラー（呼ぶ順番の間違いを知らせる）" }
        @{ Text = "パスワード|password"; Reason = "describeIngestError が、例外のメッセージを分類するために使う語（利用者に出す文言ではない）" }
        @{ Text = "ファイル形式|破損|壊れ|corrupt|file format"; Reason = "describeIngestError が、例外のメッセージを分類するために使う語（利用者に出す文言ではない）" }
        @{ Text = "IngestDispatcher の options に {} がありません。"; Reason = "開発者向けの内部エラー（部品の渡し忘れを知らせる）" }
        @{ Text = "待つ間のダイアログの場面が不明です: {}"; Reason = "開発者向けの内部エラー（場面の名前の書き間違いを知らせる）" }
    )

    # 追えない呼び出しの一覧: Key（<ファイル名>:<関数名>:<引数の表記>）・Reason。
    # 変数・関数をたどっても日本語の文字列が取り出せない呼び出し（例外のメッセージをそのまま出すものなど）
    ${untraceableAllowed} = @(
        @{ Key = 'pdf_file.ps1:readPdfPages:$message'; Reason = "読み取りで投げた PDF の文言をそのまま投げ直す（元の文言は除外の一覧の Pattern で扱う）" }
        @{ Key = 'workspace_mover.ps1:WorkspaceMover.MoveEntry:$original'; Reason = "捕まえた例外を、後片付けのあとで投げ直す（元の文言は、それぞれの投げ元で一覧に載せている）" }
        @{ Key = 'index_archive.ps1:expandImportArchive:$original'; Reason = "捕まえた例外を、後片付けのあとで投げ直す（元の文言は、それぞれの投げ元で一覧に載せている）" }
        @{ Key = 'index_archive.ps1:importIndexCore:$original'; Reason = "捕まえた例外を、後片付けのあとで投げ直す（元の文言は、それぞれの投げ元で一覧に載せている）" }
        @{ Key = 'extract_office.ps1:throwProtectionFailure:$exception'; Reason = "保護の確認で捕まえた例外を投げ直す（元の文言は、それぞれの投げ元で一覧に載せている）" }
        @{ Key = 'indexer_run.ps1:invokeIndexer:$message'; Reason = "想定していない例外のメッセージをそのまま書く（errors.md の「例外のメッセージそのまま」の行）" }
        @{ Key = 'indexer_run.ps1:cancelBeforeIngest:$message'; Reason = "引数で受けた文言を書く。呼び出し側が渡す 2 つの文言は log-warnings.md に載せているが、引数をたどる検査はまだ無く、機械では確かめていない" }
        @{ Key = 'startup_error.ps1:reportStartupFailure:$err'; Reason = "起動の失敗の例外をそのまま画面に出す（元の文言は、それぞれの投げ元で一覧に載せている）" }
    )

    function normalizeSpace([string]$s) {
        return ([regex]::Replace($s, '\s+', ' ')).Trim()
    }

    function testHasJapanese([string]$s) {
        return [regex]::IsMatch($s, '[぀-ヿ一-鿿]')
    }

    # -f の左の文字列の {0} / {1:N0} を埋め込みにする
    function replaceFormatHolders([string]$s) {
        return [regex]::Replace($s, '\{\d+(,-?\d+)?(:[^}]*)?\}', ${mark})
    }

    # 展開される文字列（"…$x…"）を、埋め込みを目印にした文字列にする
    function getExpandableText($ast) {
        $text = $ast.Extent.Text
        $base = $ast.Extent.StartOffset
        $type = $ast.StringConstantType.ToString()
        $nested = @($ast.NestedExpressions | Sort-Object { $_.Extent.StartOffset } -Descending)
        foreach ($n in $nested) {
            $s = $n.Extent.StartOffset - $base
            $e = $n.Extent.EndOffset - $base
            $text = $text.Substring(0, $s) + ${mark} + $text.Substring($e)
        }
        if ($type -eq "DoubleQuoted") {
            $text = $text.Substring(1, $text.Length - 2)
            $text = $text.Replace('""', '"')
        } else {
            $text = [regex]::Replace($text, '^@"\r?\n', '')
            $text = [regex]::Replace($text, '\r?\n"@$', '')
        }
        # バッククォートのエスケープ（`n は空白、それ以外はその文字）
        $text = [regex]::Replace($text, '`(.)', {
            param ($m)
            if ($m.Groups[1].Value -eq "n" -or $m.Groups[1].Value -eq "r" -or $m.Groups[1].Value -eq "t") { " " } else { $m.Groups[1].Value }
        })
        return $text
    }

    function getVariableKey($ast) {
        $path = $ast.VariablePath.UserPath
        $i = $path.LastIndexOf(":")
        if ($i -ge 0) { $path = $path.Substring($i + 1) }
        return $path.ToLowerInvariant()
    }

    # 関数・メソッドを囲む AST（無ければ $null）。メソッドの本体の関数はメソッド側に寄せる
    function getEnclosingFunction($ast) {
        $p = $ast.Parent
        while ($null -ne $p) {
            if ($p -is [System.Management.Automation.Language.FunctionMemberAst]) { return $p }
            if ($p -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                -not ($p.Parent -is [System.Management.Automation.Language.FunctionMemberAst])) { return $p }
            $p = $p.Parent
        }
        return $null
    }

    function getFunctionLabel($fn) {
        if ($null -eq $fn) { return "" }
        if ($fn -is [System.Management.Automation.Language.FunctionMemberAst]) { return "$($fn.Parent.Name).$($fn.Name)" }
        return $fn.Name
    }

    # 関数の中の、日本語を含む文字列の候補
    function getFunctionStrings($fn) {
        if (${fnStringsCache}.ContainsKey($fn)) { return ${fnStringsCache}[$fn] }
        $result = @()
        $nodes = $fn.FindAll({
            param ($a)
            ($a -is [System.Management.Automation.Language.StringConstantExpressionAst]) -or
            ($a -is [System.Management.Automation.Language.ExpandableStringExpressionAst])
        }, $true)
        foreach ($n in $nodes) {
            if ($n -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) {
                $t = getExpandableText $n
            } else {
                $t = $n.Value
            }
            if (testHasJapanese $t) { $result += $t }
        }
        ${fnStringsCache}[$fn] = @($result)
        return @($result)
    }

    # 関数の中の、呼び出しと変数の参照（たどる候補）
    function getFunctionRefNodes($fn) {
        if (${fnRefsCache}.ContainsKey($fn)) { return ${fnRefsCache}[$fn] }
        $nodes = @($fn.FindAll({
            param ($a)
            ($a -is [System.Management.Automation.Language.CommandAst]) -or
            ($a -is [System.Management.Automation.Language.VariableExpressionAst])
        }, $true))
        ${fnRefsCache}[$fn] = $nodes
        return $nodes
    }

    # 変数の代入を、同じ関数の中 → 同じファイルの直下 → ほかのファイルの直下の順に探す
    function findAssigns($ctx, [string]$key) {
        if ($null -ne $ctx.Function) {
            $own = @(getAssignsOf $ctx.File $key | Where-Object { [object]::ReferenceEquals($_.Function, $ctx.Function) })
            if ($own.Count -gt 0) { return $own }
        }
        $top = @(getAssignsOf $ctx.File $key | Where-Object { $null -eq $_.Function })
        if ($top.Count -gt 0) { return $top }
        $others = @()
        foreach ($f in $ctx.Files) {
            if ([object]::ReferenceEquals($f, $ctx.File)) { continue }
            $others += @(getAssignsOf $f $key | Where-Object { $null -eq $_.Function })
        }
        return $others
    }

    function getAssignsOf($fileInfo, [string]$key) {
        if ($fileInfo.ByKey.ContainsKey($key)) { return $fileInfo.ByKey[$key] }
        return @()
    }

    # どこかのファイルの直下に、この名前の代入があるか
    function testGlobalAssign($ctx, [string]$key) {
        foreach ($f in $ctx.Files) {
            if (@(getAssignsOf $f $key | Where-Object { $null -eq $_.Function }).Count -gt 0) { return $true }
        }
        return $false
    }

    # 呼んだ関数が返す文字列。関数の中に日本語の文字列が無ければ、その中の変数・別の関数の呼び出しをたどる（3 段まで）
    function resolveFunctionStrings($fn, $ctx, $state) {
        $own = @(getFunctionStrings $fn)
        if ($own.Count -gt 0 -or $ctx.Depth -ge 3) { return $own }
        $inner = @{ File = $ctx.File; Files = $ctx.Files; Functions = $ctx.Functions; Function = $fn; Depth = ($ctx.Depth + 1) }
        $r = @()
        foreach ($n in (getFunctionRefNodes $fn)) {
            $name = $null
            if ($n -is [System.Management.Automation.Language.CommandAst]) { $name = $n.GetCommandName() }
            if ($null -ne $name -and -not $ctx.Functions.ContainsKey($name.ToLowerInvariant())) { continue }
            if ($n -is [System.Management.Automation.Language.VariableExpressionAst] -and
                -not (testGlobalAssign $inner (getVariableKey $n))) { continue }
            if ($null -ne $name -and [object]::ReferenceEquals($ctx.Functions[$name.ToLowerInvariant()], $fn)) { continue }
            $sub = @{ Opaque = $false }
            $r += @(resolveStrings $n $inner $sub | Where-Object { testHasJapanese $_ })
        }
        return $r
    }

    # { Files: ファイル名 → { Ast; Assigns } ; Functions: 小文字の関数名 → AST } を作る
    function loadScripts([string[]]$Paths) {
        $files = @()
        $functions = @{}
        foreach ($path in $Paths) {
            $tokens = $null; $errors = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
            $files += (newScriptInfo ([System.IO.Path]::GetFileName($path)) $ast $functions)
        }
        return @{ Files = $files; Functions = $functions }
    }

    function newScriptInfo([string]$Name, $ast, $functions) {
        $assigns = @()
        $byKey = @{}
        $throws = New-Object System.Collections.Generic.List[object]
        $logCalls = New-Object System.Collections.Generic.List[object]
        $functionAsts = New-Object System.Collections.Generic.List[object]
        # 1 回の走査で、代入・関数の定義・throw・writeIndexerLog の呼び出しを集める（AST の走査は遅いので回数を減らす）
        $wantLog = $ast.Extent.Text.Contains("writeIndexerLog")
        $nodes = $ast.FindAll({
            param ($x)
            ($x -is [System.Management.Automation.Language.AssignmentStatementAst]) -or
            ($x -is [System.Management.Automation.Language.FunctionDefinitionAst]) -or
            ($x -is [System.Management.Automation.Language.ThrowStatementAst]) -or
            ($wantLog -and ($x -is [System.Management.Automation.Language.CommandAst]))
        }, $true)
        foreach ($a in $nodes) {
            if ($a -is [System.Management.Automation.Language.ThrowStatementAst]) { $throws.Add($a); continue }
            if ($a -is [System.Management.Automation.Language.CommandAst]) {
                if ($a.GetCommandName() -ieq "writeIndexerLog") { $logCalls.Add($a) }
                continue
            }
            if ($a -is [System.Management.Automation.Language.FunctionDefinitionAst]) {
                if (-not ($a.Parent -is [System.Management.Automation.Language.FunctionMemberAst])) { $functionAsts.Add($a) }
                continue
            }
            $left = $a.Left
            if ($left -is [System.Management.Automation.Language.ConvertExpressionAst]) { $left = $left.Child }
            if ($left -is [System.Management.Automation.Language.VariableExpressionAst]) {
                $item = @{ Key = (getVariableKey $left); Right = $a.Right; Function = (getEnclosingFunction $a) }
                $assigns += $item
                if (-not $byKey.ContainsKey($item.Key)) { $byKey[$item.Key] = New-Object System.Collections.Generic.List[object] }
                $byKey[$item.Key].Add($item)
            }
        }
        foreach ($f in $functionAsts) {
            $k = $f.Name.ToLowerInvariant()
            if (-not $functions.ContainsKey($k)) { $functions[$k] = $f }
        }
        return @{ Name = $Name; Ast = $ast; Assigns = $assigns; ByKey = $byKey; Throws = $throws; LogCalls = $logCalls }
    }

    # 展開される文字列。埋め込みの変数（$x・${x}）をたどり、固定の部分と組み合わせる。
    # たどれない変数があり、日本語が 1 つも出なければ Opaque にする（$_ とメンバーの参照・式は、今までどおり埋め込みのまま）
    function resolveExpandable($ast, $ctx, $state) {
        $base = getExpandableText $ast
        $nested = @($ast.NestedExpressions | Sort-Object { $_.Extent.StartOffset })
        $pieces = $base.Split([char]1)
        $results = @($pieces[0])
        $unresolved = $false
        for ($i = 0; $i -lt $nested.Count; $i++) {
            $n = $nested[$i]
            $options = @(${mark})
            if ($n -is [System.Management.Automation.Language.VariableExpressionAst] -and (getVariableKey $n) -ne "_") {
                $sub = @{ Opaque = $false }
                $cands = @(resolveStrings $n $ctx $sub | Where-Object { testHasJapanese $_ } | Select-Object -First 8)
                if ($cands.Count -gt 0) { $options = $cands } else { $unresolved = $true }
            }
            $next = @()
            foreach ($r in $results) { foreach ($o in $options) { $next += ($r + $o + $pieces[$i + 1]) } }
            $results = @($next | Select-Object -First 64)
        }
        if ($unresolved -and @($results | Where-Object { testHasJapanese $_ }).Count -eq 0) { $state.Opaque = $true }
        return $results
    }

    # 式から文字列の候補を取り出す。$state.Opaque は「変数・関数をたどっても取り出せなかった」印
    function resolveStrings($ast, $ctx, $state) {
        if ($null -eq $ast) { return @() }
        $T = [System.Management.Automation.Language.TokenKind]
        if ($ast -is [System.Management.Automation.Language.StringConstantExpressionAst]) { return @($ast.Value) }
        if ($ast -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) { return @(resolveExpandable $ast $ctx $state) }
        if ($ast -is [System.Management.Automation.Language.PipelineAst]) {
            $r = @()
            foreach ($e in $ast.PipelineElements) { $r += @(resolveStrings $e $ctx $state) }
            return $r
        }
        if ($ast -is [System.Management.Automation.Language.CommandExpressionAst]) { return @(resolveStrings $ast.Expression $ctx $state) }
        if ($ast -is [System.Management.Automation.Language.ParenExpressionAst]) { return @(resolveStrings $ast.Pipeline $ctx $state) }
        if ($ast -is [System.Management.Automation.Language.ConvertExpressionAst]) { return @(resolveStrings $ast.Child $ctx $state) }
        if ($ast -is [System.Management.Automation.Language.ArrayLiteralAst]) {
            $r = @()
            foreach ($e in $ast.Elements) { $r += @(resolveStrings $e $ctx $state) }
            return $r
        }
        if ($ast -is [System.Management.Automation.Language.SubExpressionAst]) {
            $r = @()
            foreach ($s in $ast.SubExpression.Statements) { $r += @(resolveStrings $s $ctx $state) }
            return $r
        }
        if ($ast -is [System.Management.Automation.Language.BinaryExpressionAst]) {
            if ($ast.Operator -eq $T::Format) {
                return @(@(resolveStrings $ast.Left $ctx $state) | ForEach-Object { replaceFormatHolders $_ })
            }
            if ($ast.Operator -eq $T::Plus) {
                $left = @(resolveStrings $ast.Left $ctx $state)
                $right = @(resolveStrings $ast.Right $ctx $state)
                if ($left.Count -eq 0) { $left = @(${mark}) }
                if ($right.Count -eq 0) { $right = @(${mark}) }
                $r = @()
                foreach ($l in $left) { foreach ($x in $right) { $r += ($l + $x) } }
                return $r
            }
            $state.Opaque = $true
            return @()
        }
        if ($ast -is [System.Management.Automation.Language.InvokeMemberExpressionAst]) {
            $r = @()
            foreach ($a in $ast.Arguments) { $r += @(resolveStrings $a $ctx $state) }
            return $r
        }
        if ($ast -is [System.Management.Automation.Language.CommandAst]) {
            $name = $ast.GetCommandName()
            if ($null -eq $name) { return @() }
            if ($name -ieq "New-Object") {
                $r = @()
                for ($i = 1; $i -lt $ast.CommandElements.Count; $i++) {
                    $e = $ast.CommandElements[$i]
                    if ($e -is [System.Management.Automation.Language.CommandParameterAst]) { continue }
                    $r += @(resolveStrings $e $ctx $state)
                }
                return $r
            }
            $key = $name.ToLowerInvariant()
            if ($ctx.Functions.ContainsKey($key)) {
                $found = @(resolveFunctionStrings $ctx.Functions[$key] $ctx $state)
                if ($found.Count -eq 0) { $state.Opaque = $true }
                return $found
            }
            $state.Opaque = $true
            return @()
        }
        if ($ast -is [System.Management.Automation.Language.VariableExpressionAst]) {
            $key = getVariableKey $ast
            if ($key -eq "_" -or $ctx.Depth -ge 3) { return @() }
            $assigns = @(findAssigns $ctx $key)
            $r = @()
            $inner = @{ File = $ctx.File; Files = $ctx.Files; Functions = $ctx.Functions; Function = $ctx.Function; Depth = ($ctx.Depth + 1) }
            foreach ($a in $assigns) {
                $right = $a.Right
                if ($right -is [System.Management.Automation.Language.CommandExpressionAst] -or
                    $right -is [System.Management.Automation.Language.PipelineAst]) {
                    $r += @(resolveStrings $right $inner $state)
                } else {
                    # $x = if (…) { "…" } else { "…" } など
                    foreach ($s in @(getFunctionStrings $right)) { $r += $s }
                }
            }
            if (@($r | Where-Object { testHasJapanese $_ }).Count -eq 0) { $state.Opaque = $true }
            return $r
        }
        # 例外の投げ直し（$_・$x.Exception・$x.Message など）と数値などの定数は、文言を持たない
        if ($ast -is [System.Management.Automation.Language.MemberExpressionAst] -or
            $ast -is [System.Management.Automation.Language.ConstantExpressionAst]) { return @() }
        $state.Opaque = $true
        return @()
    }

    function addMessage($list, [string]$Kind, $fileInfo, $fn, $argAst, $allFiles, $functions) {
        $ctx = @{ File = $fileInfo; Files = $allFiles; Functions = $functions; Function = $fn; Depth = 0 }
        $state = @{ Opaque = $false }
        $strings = @(resolveStrings $argAst $ctx $state)
        $japanese = @($strings | Where-Object { testHasJapanese $_ })
        $label = getFunctionLabel $fn
        $argText = normalizeSpace $argAst.Extent.Text
        if ($japanese.Count -eq 0) {
            if ($state.Opaque) {
                $list.Add(@{ Kind = $Kind; File = $fileInfo.Name; Function = $label; Untraceable = $true; Key = "$($fileInfo.Name):${label}:$argText" })
            }
            return
        }
        foreach ($s in $japanese) {
            $list.Add(@{ Kind = $Kind; File = $fileInfo.Name; Function = $label; Untraceable = $false; Text = (normalizeSpace $s) })
        }
    }

    # スクリプトの集まりから、(A)(B)(C) の文言と、追えない呼び出しを取り出す
    function findMessages($scripts) {
        $list = New-Object System.Collections.Generic.List[object]
        foreach ($file in $scripts.Files) {
            foreach ($t in $file.Throws) {
                if ($null -eq $t.Pipeline) { continue }
                addMessage $list "throw" $file (getEnclosingFunction $t) $t.Pipeline $scripts.Files $scripts.Functions
            }
            foreach ($c in $file.LogCalls) {
                $cmdArgs = @($c.CommandElements | Select-Object -Skip 1 | Where-Object { -not ($_ -is [System.Management.Automation.Language.CommandParameterAst]) })
                if ($cmdArgs.Count -lt 2) { continue }
                $color = $cmdArgs[1]
                if (-not ($color -is [System.Management.Automation.Language.StringConstantExpressionAst])) { continue }
                if ($color.Value -ine "Yellow" -and $color.Value -ine "Red") { continue }
                addMessage $list "writeIndexerLog" $file (getEnclosingFunction $c) $cmdArgs[0] $scripts.Files $scripts.Functions
            }
            if ($file.Name -eq "indexer_state.ps1" -and $scripts.Functions.ContainsKey("describeingesterror")) {
                $fn = $scripts.Functions["describeingesterror"]
                foreach ($s in @(getFunctionStrings $fn)) {
                    $list.Add(@{ Kind = "describeIngestError"; File = $file.Name; Function = "describeIngestError"; Untraceable = $false; Text = (normalizeSpace $s) })
                }
            }
        }
        return $list.ToArray()
    }

    # 文言を、設計書の行に合わせた正規表現にする（埋め込みは 1 文字以上の何か）
    function getMessagePattern([string]$Text) {
        $parts = $Text.Split([char]1)
        $escaped = @($parts | ForEach-Object { [regex]::Escape($_) })
        return ($escaped -join ".+")
    }

    function getDocLines([string[]]$Paths) {
        $lines = @()
        foreach ($p in $Paths) {
            foreach ($l in [System.IO.File]::ReadAllLines($p, [System.Text.Encoding]::UTF8)) {
                $lines += (normalizeSpace $l.Replace('\|', '|'))
            }
        }
        return @($lines)
    }

    function testListed([string]$Text, $docLines) {
        $pattern = getMessagePattern (normalizeSpace $Text)
        foreach ($l in $docLines) {
            if ([regex]::IsMatch($l, $pattern)) { return $true }
        }
        return $false
    }

    function showText([string]$Text) {
        return $Text.Replace([string][char]1, "{}")
    }

    function toExcludedForm([string]$Text) {
        return (normalizeSpace (showText $Text))
    }

    # 一覧に載っていない文言（除外の一覧にあるものを除く）
    function findUnlisted($messages, $docLines, $excluded) {
        $result = @()
        $seen = @{}
        foreach ($m in $messages) {
            if ($m.Untraceable) { continue }
            if (testListed $m.Text $docLines) { continue }
            $form = toExcludedForm $m.Text
            if (@($excluded | Where-Object { testExcludedBy $_ $form }).Count -gt 0) { continue }
            if ($seen.ContainsKey($form)) { continue }
            $seen[$form] = $true
            $result += "$($m.File):$($m.Function) [$($m.Kind)] $form"
        }
        return @($result)
    }

    function findUnlistedUntraceable($messages, $allowed) {
        $keys = @($messages | Where-Object { $_.Untraceable } | ForEach-Object { $_.Key })
        return @($keys | Where-Object { $k = $_; @($allowed | Where-Object { $_.Key -eq $k }).Count -eq 0 } | Select-Object -Unique)
    }

    function findStaleUntraceable($messages, $allowed) {
        $keys = @($messages | Where-Object { $_.Untraceable } | ForEach-Object { $_.Key })
        return @($allowed | Where-Object { $a = $_; $keys -notcontains $a.Key } | ForEach-Object { $_.Key })
    }

    function findStaleExcluded($messages, $excluded) {
        $forms = @($messages | Where-Object { -not $_.Untraceable } | ForEach-Object { toExcludedForm $_.Text })
        $stale = @()
        foreach ($e in $excluded) {
            $hit = @($forms | Where-Object { testExcludedBy $e $_ })
            if ($hit.Count -eq 0) { $stale += (showExcludedEntry $e) }
        }
        return @($stale)
    }

    # 除外の一覧の項目が、この文言（置き換えた形）に当たるか。Text は全体が同じ、Pattern は正規表現に一致
    function testExcludedBy($entry, [string]$form) {
        if ($entry.Pattern) { return [regex]::IsMatch($form, $entry.Pattern) }
        return ((normalizeSpace $entry.Text) -eq $form)
    }

    function showExcludedEntry($entry) {
        if ($entry.Pattern) { return "Pattern: $($entry.Pattern)" }
        return $entry.Text
    }

    ${fnStringsCache} = @{}
    ${fnRefsCache} = @{}
    ${scripts} = loadScripts @(Get-ChildItem ${scriptsRoot} -Filter "*.ps1" -Recurse | ForEach-Object { $_.FullName })
    ${messagesAll} = @(findMessages ${scripts})
    ${docLinesAll} = @(getDocLines (getListDocPaths))
}

Describe "利用者に出すエラー・警告の文言と設計書の一覧" -Tag Meta {
    It "突き合わせる設計書（errors*.md・log-warnings*.md・gui の messages*.md）が読めて、文言を取り出せている" {
        @(getListDocPaths).Count | Should -BeGreaterOrEqual 5
        ${docLinesAll}.Count | Should -BeGreaterThan 100
        @(${messagesAll} | Where-Object { $_.Kind -eq "throw" }).Count | Should -BeGreaterThan 10
        @(${messagesAll} | Where-Object { $_.Kind -eq "writeIndexerLog" }).Count | Should -BeGreaterThan 10
        @(${messagesAll} | Where-Object { $_.Kind -eq "describeIngestError" }).Count | Should -BeGreaterThan 3
    }

    It "throw・writeIndexerLog（Yellow・Red）・describeIngestError の文言が、設計書の一覧にあるか、除外の一覧にある" {
        $unlisted = findUnlisted ${messagesAll} ${docLinesAll} ${messageExcluded}
        ($unlisted -join "`n") | Should -Be ""
    }

    It "追えない呼び出しが、追えない呼び出しの一覧にある" {
        $missing = findUnlistedUntraceable ${messagesAll} ${untraceableAllowed}
        ($missing -join "`n") | Should -Be ""
    }

    It "追えない呼び出しの一覧に、今のコードに無い項目が残っていない" {
        $stale = findStaleUntraceable ${messagesAll} ${untraceableAllowed}
        ($stale -join "`n") | Should -Be ""
    }

    It "除外の一覧に、今のコードに無い文言が残っていない" {
        $stale = findStaleExcluded ${messagesAll} ${messageExcluded}
        ($stale -join "`n") | Should -Be ""
    }

    It "除外の一覧と追えない呼び出しの一覧の項目に、理由がある" {
        $noReason = @(@(${messageExcluded}) + @(${untraceableAllowed}) | Where-Object { [string]::IsNullOrWhiteSpace($_.Reason) })
        $noReason.Count | Should -Be 0
    }
}

Describe "文言の取り出しと突き合わせの部品（小さな例）" -Tag Meta {
    BeforeAll {
        # ソースの文字列から、文言と追えない呼び出しを取り出す
        function findMessagesIn([string]$Source, [string]$Name = "x.ps1", [string[]]$OtherSources = @()) {
            $functions = @{}
            $files = @()
            $tokens = $null; $errors = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseInput($Source, [ref]$tokens, [ref]$errors)
            $files += (newScriptInfo $Name $ast $functions)
            $i = 0
            foreach ($o in $OtherSources) {
                $i++
                $ast2 = [System.Management.Automation.Language.Parser]::ParseInput($o, [ref]$tokens, [ref]$errors)
                $files += (newScriptInfo "other$i.ps1" $ast2 $functions)
            }
            return @(findMessages @{ Files = $files; Functions = $functions })
        }
        function textsOf($messages) {
            return @($messages | Where-Object { -not $_.Untraceable } | ForEach-Object { showText $_.Text })
        }
    }

    It "埋め込みの置き換え: <Name>" -TestCases @(
        @{ Name = "<N> の設計書に一致する"; Code = '"$n 件"'; Doc = "処理した <N> 件です"; Expected = $true }
        @{ Name = "具体的な値の設計書にも一致する"; Code = '"$n 件"'; Doc = "3 件"; Expected = $true }
        @{ Name = "文言が違えば一致しない"; Code = '"$n 件"'; Doc = "<N> 個"; Expected = $false }
        @{ Name = "-f の {0} を埋め込みとみなす"; Code = '("{0} 件" -f $n)'; Doc = "<N> 件"; Expected = $true }
        @{ Name = "`${x} と `$(…) も埋め込みとみなす"; Code = '"${a}と$($b.c)件"'; Doc = "<A>と<B>件"; Expected = $true }
        @{ Name = "空白の並びは 1 つにそろえる"; Code = '"a  b 日本語"'; Doc = "a b 日本語"; Expected = $true }
    ) {
        param ($Name, $Code, $Doc, $Expected)
        $m = @(findMessagesIn "function f { throw $Code }")
        $m.Count | Should -Be 1
        testListed $m[0].Text @(normalizeSpace ($Doc + " 日本語")) | Should -Be $Expected
    }

    It "日本語を含まない文字列は取り出さない" {
        @(findMessagesIn 'function f { throw "internal error" }').Count | Should -Be 0
    }

    It "取り出す形: <Name>" -TestCases @(
        @{ Name = "throw の文字列"; Code = 'function f { throw "あいう" }'; Expected = @("あいう") }
        @{ Name = "throw の -f の左"; Code = 'function f { throw ("あ{0}い" -f 1) }'; Expected = @("あ{}い") }
        @{ Name = "throw (New-Object 型 (文字列))"; Code = 'function f { throw (New-Object System.Exception("あいう")) }'; Expected = @("あいう") }
        @{ Name = "throw (New-Object 型 文字列)"; Code = 'function f { throw (New-Object System.Exception "あいう") }'; Expected = @("あいう") }
        @{ Name = "+ でつないだ文字列"; Code = 'function f { throw ("あ" + "い" + $x + "う") }'; Expected = @("あい{}う") }
        @{ Name = "writeIndexerLog の Yellow"; Code = 'function f { writeIndexerLog "あいう" "Yellow" }'; Expected = @("あいう") }
        @{ Name = "writeIndexerLog の Red"; Code = 'function f { writeIndexerLog "あいう" "Red" }'; Expected = @("あいう") }
        @{ Name = "writeIndexerLog の Green は見ない"; Code = 'function f { writeIndexerLog "あいう" "Green" }'; Expected = @() }
        @{ Name = "writeIndexerLog の色なしは見ない"; Code = 'function f { writeIndexerLog "あいう" }'; Expected = @() }
        @{ Name = "展開される文字列の変数をたどる"; Code = 'function f($p) { $m = "あ"; throw "${m}${p}" }'; Expected = @("あ{}") }
        @{ Name = "展開される文字列の変数が日本語の文字列を持つ（ファイルの直下）"; Code = '$n = "が起動しています"; function f($p) { throw "${p}${n}" }'; Expected = @("{}が起動しています") }
    ) {
        param ($Name, $Code, $Expected)
        $got = @(textsOf (findMessagesIn $Code))
        ($got -join "|") | Should -Be ($Expected -join "|")
    }

    It "変数の追い方: <Name>" -TestCases @(
        @{ Name = "同じ関数の中の代入"; Code = 'function f { $m = "あいう"; throw $m }'; Other = @(); Expected = @("あいう") }
        @{ Name = "if の枝ごとの代入をすべて取る"; Code = 'function f { if ($a) { $m = "あ" } else { $m = "い" }; throw $m }'; Other = @(); Expected = @("あ", "い") }
        @{ Name = "同じ関数の中が、同じファイルの直下より先"; Code = '$m = "ファイル"; function f { $m = "関数"; throw $m }'; Other = @(); Expected = @("関数") }
        @{ Name = "同じファイルの直下"; Code = '$m = "ファイル"; function f { throw $m }'; Other = @(); Expected = @("ファイル") }
        @{ Name = "同じファイルの直下が、ほかのファイルの直下より先"; Code = '$m = "自分"; function f { throw $m }'; Other = @('$m = "他"'); Expected = @("自分") }
        @{ Name = "ほかのファイルの直下"; Code = 'function f { throw ${m} }'; Other = @('$m = "他"'); Expected = @("他") }
        @{ Name = "`$script: の変数"; Code = '$script:m = "あいう"; function f { throw $script:m }'; Other = @(); Expected = @("あいう") }
        @{ Name = "`$x + 文字列"; Code = 'function f { $p = "あ"; throw ($p + "い") }'; Other = @(); Expected = @("あい") }
        @{ Name = "scripts/ の関数の呼び出しは、その関数の文字列"; Code = 'function f { writeIndexerLog (g 1) "Red" }'; Other = @('function g { return "あいう" }'); Expected = @("あいう") }
    ) {
        param ($Name, $Code, $Other, $Expected)
        $got = @(textsOf (findMessagesIn $Code "x.ps1" $Other))
        ($got -join "|") | Should -Be ($Expected -join "|")
    }

    It "追えない呼び出し: <Name>" -TestCases @(
        @{ Name = "引数のまま"; Code = 'function f($reason) { throw $reason }'; Expected = "x.ps1:f:`$reason" }
        @{ Name = "例外のメッセージの代入"; Code = 'function f { $m = $_.Exception.Message; writeIndexerLog $m "Red" }'; Expected = "x.ps1:f:`$m" }
        @{ Name = "日本語の無い関数の呼び出し"; Code = 'function f { throw (g 1) }'; Expected = "x.ps1:f:(g 1)" }
        @{ Name = "展開される文字列が引数だけ"; Code = 'function f($p) { throw "$p" }'; Expected = 'x.ps1:f:"$p"' }
        @{ Name = "知らない式の形（-join）"; Code = 'function f($p) { throw ($p -join "x") }'; Expected = 'x.ps1:f:($p -join "x")' }
    ) {
        param ($Name, $Code, $Expected)
        $m = @(findMessagesIn ($Code + "`nfunction g { return 1 }") | Where-Object { $_.Untraceable })
        $m.Count | Should -Be 1
        $m[0].Key | Should -Be $Expected
    }

    It "`$_ とメンバーの参照、再 throw は追えない呼び出しにしない" {
        @(findMessagesIn 'function f { try { 1 } catch { throw $_ } }').Count | Should -Be 0
        @(findMessagesIn 'function f { try { 1 } catch { throw $_.Exception } }').Count | Should -Be 0
        @(findMessagesIn 'function f { try { 1 } catch { throw } }').Count | Should -Be 0
    }

    It "describeIngestError の中の文字列を取り出す" {
        $src = 'function describeIngestError($e) { if ($e) { return "原因です" } return "ほかの原因" }'
        $m = @(findMessagesIn $src "indexer_state.ps1")
        (@($m | ForEach-Object { $_.Kind } | Select-Object -Unique) -join ",") | Should -Be "describeIngestError"
        $m.Count | Should -Be 2
    }

    It "一覧の残り: 除外の一覧・追えない呼び出しの一覧の、今のコードに無い項目を見つける" {
        $messages = @(findMessagesIn 'function f($r) { throw "あいう"; throw $r }')
        @(findStaleExcluded $messages @(@{ Text = "あいう"; Reason = "r" })).Count | Should -Be 0
        @(findStaleExcluded $messages @(@{ Text = "消えた文言"; Reason = "r" })).Count | Should -Be 1
        @(findStaleExcluded $messages @(@{ Pattern = "^あ.う$"; Reason = "r" })).Count | Should -Be 0
        @(findStaleExcluded $messages @(@{ Pattern = "^消えた"; Reason = "r" })).Count | Should -Be 1
        @(findStaleUntraceable $messages @(@{ Key = 'x.ps1:f:$r'; Reason = "r" })).Count | Should -Be 0
        @(findStaleUntraceable $messages @(@{ Key = 'x.ps1:f:$gone'; Reason = "r" })).Count | Should -Be 1
        @(findUnlistedUntraceable $messages @()).Count | Should -Be 1
    }

    It "一覧に無い文言を見つけ、除外の一覧にあれば見つけない" {
        $messages = @(findMessagesIn 'function f { throw "一覧に無い文言です" }')
        @(findUnlisted $messages @("別の文言") @()).Count | Should -Be 1
        @(findUnlisted $messages @("別の文言") @(@{ Text = "一覧に無い文言です"; Reason = "r" })).Count | Should -Be 0
        @(findUnlisted $messages @("前 一覧に無い文言です 後") @()).Count | Should -Be 0
        @(findUnlisted $messages @("別の文言") @(@{ Pattern = "^一覧に.+です$"; Reason = "r" })).Count | Should -Be 0
        @(findUnlisted $messages @("別の文言") @(@{ Pattern = "^別"; Reason = "r" })).Count | Should -Be 1
    }
}
