# 単一 .ps1 版を、ASCII だけの文字で書かれた形に書き換える（tools\new_single_script.ps1 が使う）。
# Windows PowerShell 5.1 は、BOM の無い UTF-8 のファイルを ANSI（日本語環境では Shift_JIS）で読むため、
# 日本語が混ざると文字化けして構文エラーになる。ASCII だけにすれば、どの文字コードで読まれても同じに読める。
#
# 正規表現ではなく、PowerShell 自身の字句・構文木（Parser）で書き換える。書き換えの決まり:
#   - 日本語を含むコメントは消す（ブロックコメントは、行数が変わらないよう改行だけ残す。改行が無ければ空白 1 つにする）。
#   - 単一引用符の文字列・ヒアストリングは、値が同じになる ([regex]::Unescape('..あ..')) の式に作り直す。
#   - 二重引用符の文字列・ヒアストリングは、$var や $(...) をそのまま残し、外側の日本語を $([regex]::Unescape('あ')) にする。
#   - 変数名・属性の引数・using・引用符の無い語に日本語があるものは、書き換えられないので止める（元のソースを直す）。
# 書き出す前に、値が同じになること・ASCII だけであること・構文エラーが 0 であることを確かめる。
# 作る処理（newAsciiScriptEdits）と確かめる処理（assertAsciiScriptResult）は分けてある。

# 文字が ASCII より外か
function hasNonAscii {
    param ([string]$text)
    return [regex]::IsMatch($text, '[^\x00-\x7F]')
}

# UTF-16 の 1 単位を \uXXXX の形にする
function unitEscape {
    param ([char]$c)
    return ('\u{0:X4}' -f [int]$c)
}

# 単一引用符の文字列を、値が同じ ASCII だけの式 ([regex]::Unescape('..あ..')) にする。
# Unescape は \uXXXX を文字に戻す（ASCII の \ は \\ にして残す）。構文木の節が少なく、読み込みが軽い
function newAsciiLiteralExpr {
    param ([string]$value)

    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $value.ToCharArray()) {
        if ([int]$c -gt 127) {
            [void]$sb.Append((unitEscape $c))
        } elseif ($c -eq '\') {
            [void]$sb.Append('\\')
        } elseif ($c -eq "'") {
            [void]$sb.Append("''")
        } else {
            [void]$sb.Append($c)
        }
    }
    return "([regex]::Unescape('" + $sb.ToString() + "'))"
}

# 二重引用符の文字列（ヒアストリング含む）の、外側の日本語だけを $([regex]::Unescape('..')) にした文字列を返す。
# $spans は残す範囲（$var・$(...) の位置。トークンの先頭からの相対）。そこに日本語があれば止める
function newAsciiExpandableText {
    param (
        [string]$raw,
        [object[]]$spans,
        [bool]$isHere,
        [string]$where
    )

    $sb = New-Object System.Text.StringBuilder
    $newSpans = New-Object System.Collections.Generic.List[object]
    $inserted = 0
    $afterSpan = $false
    $i = 0
    while ($i -lt $raw.Length) {
        $inSpan = $null
        foreach ($s in $spans) {
            if ($i -ge $s.Start -and $i -lt $s.End) { $inSpan = $s; break }
        }
        if ($inSpan) {
            $piece = $raw.Substring($i, $inSpan.End - $i)
            if ($null -ne $inSpan.Replacement) { $piece = $inSpan.Replacement }
            if (hasNonAscii $piece) {
                throw "$where に、書き換えられない場所（文字列の中の変数名など）に日本語があります。元のソースで、日本語を式の外に出してください: $piece"
            }
            [void]$newSpans.Add(@{ Start = $sb.Length; Length = $piece.Length })
            [void]$sb.Append($piece)
            $i = $inSpan.End
            $afterSpan = $true
            continue
        }
        $c = $raw[$i]
        if ([int]$c -le 127) {
            [void]$sb.Append($c)
            $i++
            $afterSpan = $false
            continue
        }
        # バッククォートで日本語を逃がしている書き方は、バッククォートを外す（逃がしても同じ文字）
        if ($sb.Length -gt 0 -and $sb[$sb.Length - 1] -eq '`' -and $i -ge 1 -and $raw[$i - 1] -eq '`') {
            $backticks = 0
            for ($k = $i - 1; $k -ge 0 -and $raw[$k] -eq '`'; $k--) { $backticks++ }
            if ($backticks % 2 -eq 1) { [void]$sb.Remove($sb.Length - 1, 1) }
        }
        # 文字としての $ の直後に $( を置くと、$$ という変数に読まれる。その $ はバッククォートで逃がす
        if (-not $afterSpan -and $sb.Length -gt 0 -and $sb[$sb.Length - 1] -eq '$') {
            $backticks = 0
            for ($k = $sb.Length - 2; $k -ge 0 -and $sb[$k] -eq '`'; $k--) { $backticks++ }
            if ($backticks % 2 -eq 0) { [void]$sb.Insert($sb.Length - 1, '`') }
        }
        # 続く日本語を 1 つの式にまとめる
        $units = New-Object System.Text.StringBuilder
        while ($true) {
            $c = $raw[$i]
            [void]$units.Append((unitEscape $c))
            # 二重引用符の文字列では、全角の二重引用符（“ ” „）は 2 つ重ねて 1 文字を表す
            if (-not $isHere -and ($c -eq [char]0x201C -or $c -eq [char]0x201D -or $c -eq [char]0x201E) -and
                ($i + 1) -lt $raw.Length -and $raw[$i + 1] -eq $c) {
                $i++
            }
            $i++
            if ($i -ge $raw.Length -or [int]$raw[$i] -le 127) { break }
            $atSpan = $false
            foreach ($s in $spans) {
                if ($i -ge $s.Start -and $i -lt $s.End) { $atSpan = $true; break }
            }
            if ($atSpan) { break }
        }
        [void]$sb.Append('$([regex]::Unescape(''' + $units.ToString() + '''))')
        $inserted++
        $afterSpan = $true
    }
    # 確かめ用: 残した範囲（変数・式）を同じ目印に置き換えた、書き換え前と後の文字列
    $newText = $sb.ToString()
    $newMask = $newText
    for ($k = $newSpans.Count - 1; $k -ge 0; $k--) {
        $newMask = $newMask.Remove($newSpans[$k].Start, $newSpans[$k].Length).Insert($newSpans[$k].Start, '$(0)')
    }
    $oldMask = $raw
    for ($k = @($spans).Count - 1; $k -ge 0; $k--) {
        $oldMask = $oldMask.Remove($spans[$k].Start, $spans[$k].End - $spans[$k].Start).Insert($spans[$k].Start, '$(0)')
    }
    return @{ Text = $newText; OldMask = $oldMask; NewMask = $newMask; Inserted = $inserted }
}

# 式の文字列を読んで、最初の文字列の構文木を返す（単独で読めない形は、かっこで包んで読む）。読めなければ $null
function getParsedStringNode {
    param ([string]$expr)

    foreach ($source in @($expr, ("(" + $expr + ")"))) {
        $t = $null
        $e = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$t, [ref]$e)
        if ($e.Count -gt 0) { continue }
        $node = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -or $n -is [System.Management.Automation.Language.ExpandableStringExpressionAst] }, $true)
        if ($node) { return $node }
    }
    return $null
}

# 文字列の中身を、変数の展開などをせずに比べるための値（ExpandableString の Value）を取り出す
function getParsedStringValue {
    param ([string]$expr)

    $node = getParsedStringNode $expr
    if (-not $node) { return $null }
    return [string]$node.Value
}

# 二重引用符の文字列の中の、展開される式（$var・$(...)）の文字列の一覧。読めなければ $null
function getParsedNestedTexts {
    param ([string]$expr)

    $node = getParsedStringNode $expr
    if (-not $node) { return $null }
    if ($node -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) {
        return , @($node.NestedExpressions | ForEach-Object { $_.Extent.Text })
    }
    return , @()
}

# 属性の引数（[ValidateSet('あ')]・[Parameter(HelpMessage = 'あ')]）の、定数の文字列の範囲。
# 属性の引数は定数でなければ動かないので、式（[regex]::Unescape(..)）には書き換えられない
function getAsciiAttributeArgumentRanges {
    param ($ast)

    $ranges = New-Object System.Collections.Generic.List[object]
    $attributes = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AttributeAst] }, $true)
    foreach ($attribute in $attributes) {
        $arguments = @($attribute.PositionalArguments) + @($attribute.NamedArguments | ForEach-Object { $_.Argument })
        foreach ($argument in $arguments) {
            if ($argument -is [System.Management.Automation.Language.ConstantExpressionAst] -or
                $argument -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) {
                [void]$ranges.Add(@{ Start = $argument.Extent.StartOffset; End = $argument.Extent.EndOffset })
            }
        }
    }
    return , $ranges
}

# 書き換えの内容（edits）と、あとで確かめる材料（checkSingle・checkDouble）を作る。ここでは書き換えた結果は作らない
function newAsciiScriptEdits {
    param (
        [string]$text,
        $tokens,
        $ast,
        [string]$label
    )

    # 二重引用符の文字列ごとの、残す範囲（変数・式）
    $nestedByStart = @{}
    $exprNodes = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.ExpandableStringExpressionAst] }, $true)
    foreach ($n in $exprNodes) {
        $nestedByStart[$n.Extent.StartOffset] = $n
    }
    $attributeRanges = getAsciiAttributeArgumentRanges $ast

    $edits = New-Object System.Collections.Generic.List[object]
    $checkSingle = New-Object System.Collections.Generic.List[object]
    $checkDouble = New-Object System.Collections.Generic.List[object]

    # 日本語の位置から、それを含むトークンだけを二分探索で拾う（全トークンの位置を取ると遅い）
    $targets = New-Object System.Collections.Generic.List[object]
    $lastEnd = -1
    foreach ($m in [regex]::Matches($text, '[^\x00-\x7F]+')) {
        $p = $m.Index
        if ($p -lt $lastEnd) { continue }
        $lo = 0
        $hi = $tokens.Count - 1
        while ($lo -lt $hi) {
            $mid = [int][Math]::Ceiling(($lo + $hi) / 2)
            if ($tokens[$mid].Extent.StartOffset -le $p) { $lo = $mid } else { $hi = $mid - 1 }
        }
        $found = $tokens[$lo]
        if ($found.Extent.StartOffset -gt $p -or $found.Extent.EndOffset -le $p) {
            throw "$label の $($found.Extent.StartLineNumber) 行目付近に、トークンの外の日本語があります。"
        }
        $lastEnd = $found.Extent.EndOffset
        [void]$targets.Add($found)
    }

    foreach ($tok in $targets) {
        $raw = $tok.Extent.Text
        if (-not (hasNonAscii $raw)) { continue }
        $start = $tok.Extent.StartOffset
        $len = $tok.Extent.EndOffset - $start
        $where = "$label の $($tok.Extent.StartLineNumber) 行目"
        $kind = $tok.Kind.ToString()

        if ($kind -eq 'Comment') {
            # 行数が変わらないよう改行だけ残す。行の途中のブロックコメントは、前後のトークンがつながらないよう空白 1 つにする
            $newlines = [regex]::Matches($raw, "\n").Count
            $commentText = " "
            if ($newlines -gt 0) { $commentText = "`r`n" * $newlines }
            [void]$edits.Add(@{ Start = $start; Length = $len; Text = $commentText })
            continue
        }
        if ($tok -is [System.Management.Automation.Language.StringToken]) {
            foreach ($range in $attributeRanges) {
                if ($start -ge $range.Start -and $start -lt $range.End) {
                    throw "$where の日本語は、属性の引数（定数でなければ動かない場所）にあるため、書き換えられません。元のソースで、日本語を使わない形にしてください: $raw"
                }
            }
        }
        if ($tok -is [System.Management.Automation.Language.StringLiteralToken]) {
            $repl = newAsciiLiteralExpr $tok.Value
            [void]$edits.Add(@{ Start = $start; Length = $len; Text = $repl })
            [void]$checkSingle.Add(@{ Expr = $repl; Expected = [string]$tok.Value; Where = $where })
            continue
        }
        if ($tok -is [System.Management.Automation.Language.StringExpandableToken]) {
            $isHere = $kind -like 'HereString*'
            $spans = @()
            $node = $nestedByStart[$start]
            if ($node) {
                foreach ($ne in $node.NestedExpressions) {
                    $span = @{ Start = $ne.Extent.StartOffset - $start; End = $ne.Extent.EndOffset - $start; Replacement = $null }
                    if ((hasNonAscii $ne.Extent.Text) -and $ne -is [System.Management.Automation.Language.SubExpressionAst]) {
                        $innerText = $ne.Extent.Text
                        $inner = $innerText.Substring(2, $innerText.Length - 3)
                        $span.Replacement = '$(' + (convertToAsciiScript $inner $where) + ')'
                    }
                    $spans += $span
                }
            }
            $made = newAsciiExpandableText $raw $spans $isHere $where
            [void]$edits.Add(@{ Start = $start; Length = $len; Text = $made.Text })
            [void]$checkDouble.Add(@{ Old = $made.OldMask; New = $made.NewMask; Inserted = $made.Inserted; Where = $where })
            continue
        }
        # $obj.日本語 の形のメンバー名は、$obj.([regex]::Unescape('..')) の形（名前を式で与える書き方）にする
        if (($kind -eq 'Identifier' -or $kind -eq 'Generic') -and $start -ge 1 -and $text[$start - 1] -eq '.' -and
            ($start -lt 2 -or $text[$start - 2] -ne '.')) {
            $repl = newAsciiLiteralExpr $raw
            [void]$edits.Add(@{ Start = $start; Length = $len; Text = $repl })
            [void]$checkSingle.Add(@{ Expr = $repl; Expected = [string]$raw; Where = $where })
            continue
        }
        # ハッシュテーブルの引用符の無いキー（@{ 日本語 = 1 }）は、キーを式で与える書き方 @{ ([regex]::Unescape('..')) = 1 } にする
        if ($kind -eq 'Identifier' -or $kind -eq 'Generic') {
            $after = $text.Substring($tok.Extent.EndOffset)
            if ($after -match '^[ \t]*=(?!=)') {
                $repl = newAsciiLiteralExpr $raw
                [void]$edits.Add(@{ Start = $start; Length = $len; Text = $repl })
                [void]$checkSingle.Add(@{ Expr = $repl; Expected = [string]$raw; Where = $where })
                continue
            }
        }
        throw "$where に、書き換えられない日本語があります（コメント・文字列以外: $kind）。元のソースで、引用符で囲んだ文字列にしてください: $raw"
    }

    return @{ Edits = $edits; CheckSingle = $checkSingle.ToArray(); CheckDouble = $checkDouble.ToArray() }
}

# 前から順に、置き換えない範囲と置き換えの文字列をつなぐ（edits はトークンの順＝位置の昇順）
function joinAsciiScriptEdits {
    param (
        [string]$text,
        $edits
    )

    $sb = New-Object System.Text.StringBuilder($text.Length)
    $pos = 0
    foreach ($e in $edits) {
        [void]$sb.Append($text, $pos, $e.Start - $pos)
        [void]$sb.Append($e.Text)
        $pos = $e.Start + $e.Length
    }
    [void]$sb.Append($text, $pos, $text.Length - $pos)
    return $sb.ToString()
}

# 書き換えた結果を確かめる。違えば例外にする（結果は返さない）
function assertAsciiScriptResult {
    param (
        [string]$result,
        $checkSingle,
        $checkDouble,
        [string]$label
    )

    if (hasNonAscii $result) {
        $lineNo = 0
        foreach ($line in ($result -split "`r?`n")) {
            $lineNo++
            if (hasNonAscii $line) { throw "$label の $lineNo 行目に、書き換えられなかった日本語が残りました: $line" }
        }
    }

    # 確かめ 1: 単一引用符の文字列は、作り直した式を評価して元の値と同じになる
    if (@($checkSingle).Count -gt 0) {
        $joined = ($checkSingle | ForEach-Object { $_.Expr }) -join ",`r`n"
        $values = & ([scriptblock]::Create("return ,@($joined)"))
        for ($k = 0; $k -lt @($checkSingle).Count; $k++) {
            if ([string]$values[$k] -ne $checkSingle[$k].Expected) {
                throw "$($checkSingle[$k].Where) の文字列が、書き換えで元と違う値になりました。"
            }
        }
    }

    # 確かめ 2: 二重引用符の文字列は、(a) 残した式（$var・$(...)）の数が前後で同じで、書き換えで足した
    # $([regex]::Unescape('..')) のほかに、展開される式が増えていない（残した式の中身そのものは比べない）、(b) 書き換えた文字列の値の $([regex]::Unescape('..')) を戻すと、元の値と同じになる
    foreach ($c in $checkDouble) {
        $oldValue = getParsedStringValue $c.Old
        $newValue = getParsedStringValue $c.New
        $oldNested = getParsedNestedTexts $c.Old
        $newNested = getParsedNestedTexts $c.New
        if ($null -eq $oldValue -or $null -eq $newValue -or $null -eq $oldNested -or $null -eq $newNested) {
            throw "$($c.Where) の文字列を、書き換えのあとに確かめられませんでした。"
        }
        $newOther = @($newNested | Where-Object { -not $_.StartsWith('$([regex]::Unescape(') })
        $newAdded = @($newNested).Count - $newOther.Count
        $same = ($newAdded -eq [int]$c.Inserted -and $newOther.Count -eq @($oldNested).Count)
        for ($k = 0; $same -and $k -lt $newOther.Count; $k++) {
            if ($newOther[$k] -ne @($oldNested)[$k]) { $same = $false }
        }
        if (-not $same) {
            throw "$($c.Where) の文字列の、展開される式（変数・`$(...)）が、書き換えで元と違う形になりました。"
        }
        $decoded = [regex]::Replace($newValue, '\$\(\[regex\]::Unescape\(''((?:\\u[0-9A-F]{4})+)''\)\)', { param($m) [regex]::Unescape($m.Groups[1].Value) })
        if ($decoded -ne $oldValue) {
            throw "$($c.Where) の文字列が、書き換えで元と違う値になりました。"
        }
    }

    # 確かめ 3: 構文エラーが 0
    $t2 = $null
    $e2 = $null
    [void][System.Management.Automation.Language.Parser]::ParseInput($result, [ref]$t2, [ref]$e2)
    if ($e2.Count -gt 0) {
        throw "$label を ASCII に書き換えたところ、構文エラーになりました: $($e2[0].Message)（$($e2[0].Extent.StartLineNumber) 行目: $($e2[0].Extent.Text)）"
    }
}

# スクリプトの文字列を、ASCII だけの文字で書かれた同じ動きのものにする
function convertToAsciiScript {
    param (
        [string]$text,
        [string]$label
    )

    if (-not (hasNonAscii $text)) { return $text }

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($text, [ref]$tokens, [ref]$errors)
    if ($errors.Count -gt 0) {
        throw "$label の読み取りに失敗しました: $($errors[0].Message)（$($errors[0].Extent.StartLineNumber) 行目）"
    }

    $made = newAsciiScriptEdits $text $tokens $ast $label
    $result = joinAsciiScriptEdits $text $made.Edits
    assertAsciiScriptResult $result $made.CheckSingle $made.CheckDouble $label
    return $result
}

# XAML の文字列を、ASCII だけの文字で書かれた同じ内容のものにする。日本語は &#xNNNN; にし、コメントは消す。
# 書き換えの前後で、コメントを除いた XML の内容が同じになることを確かめる
function convertToAsciiXaml {
    param (
        [string]$text,
        [string]$label
    )

    if (-not (hasNonAscii $text)) { return $text }

    $noComment = [regex]::Replace($text, '<!--.*?-->', {
            param($m)
            if (hasNonAscii $m.Value) { return "" }
            return $m.Value
        }, [System.Text.RegularExpressions.RegexOptions]::Singleline)

    $sb = New-Object System.Text.StringBuilder
    $chars = $noComment.ToCharArray()
    for ($i = 0; $i -lt $chars.Length; $i++) {
        $c = $chars[$i]
        if ([int]$c -le 127) {
            [void]$sb.Append($c)
        } elseif ([char]::IsHighSurrogate($c) -and ($i + 1) -lt $chars.Length -and [char]::IsLowSurrogate($chars[$i + 1])) {
            [void]$sb.AppendFormat('&#x{0:X};', [char]::ConvertToUtf32($c, $chars[$i + 1]))
            $i++
        } else {
            [void]$sb.AppendFormat('&#x{0:X4};', [int]$c)
        }
    }
    $result = $sb.ToString()

    # 確かめ: 前後を XML として読み、コメントを除いた内容（OuterXml）が同じ
    $before = New-Object System.Xml.XmlDocument
    $before.PreserveWhitespace = $true
    $before.LoadXml($text)
    $after = New-Object System.Xml.XmlDocument
    $after.PreserveWhitespace = $true
    $after.LoadXml($result)
    $removeComments = {
        param($doc)
        foreach ($n in @($doc.SelectNodes('//comment()'))) { [void]$n.ParentNode.RemoveChild($n) }
    }
    & $removeComments $before
    & $removeComments $after
    # 改行だけが残ったコメントのあとの空白の違いは、コメントを除いた時点で同じ位置に残る
    if ($before.OuterXml.Replace("`r`n", "`n") -ne $after.OuterXml.Replace("`r`n", "`n")) {
        throw "$label の XAML が、書き換えで元と違う内容になりました。"
    }
    return $result
}
