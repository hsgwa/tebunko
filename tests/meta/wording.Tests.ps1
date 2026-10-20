# 画面・進み具合・ログ・例外の文言に、内部の言葉（システムインデックス・集約ファイル・本文インデックス）と、対象を Office だけに見せる言葉（Office ファイル）を出さないことを確かめる。
# 利用者には、どれも「インデックス」と呼ぶ。対象は Excel・PowerPoint・Word・PDF・テキストなので、並べて書く（決まりは docs\design\gui\common.md）。内部の言葉はコメントとコードの名前でだけ使ってよい（コメントは対象にしない）。
# 対象: scripts\ のすべての .ps1 の文字列（AST の文字列の定数）と、すべての .xaml の属性の値・文字
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"

    $script:internalTerms = "集約ファイル|システムインデックス|本文インデックス|Office\s*ファイル"
    # 許す例外。文字列からこの型に当たる所を取り除いてから探す
    $script:allowedPatterns = @(
        # 前の版のシステムインデックスの txt の名前のパターン（indexer_plan.ps1 の tebunkoOwnFileNamePatterns。今の版が自分のファイルとして外す）
        "システムインデックス\*\.txt"
        # 画面の 2 か所（［インデックス管理］の詳細・［検索］の検索対象の欄）の「集約ファイル N 件」。画面の刷新のときに直すので、それまで許す
        '集約ファイル (?=\$\(\$summary)'
    )

    function findInternalTerms {
        # 文字列から、許す例外を除いて、内部の言葉に当たる所を返す
        param ([string]$text)

        $rest = $text
        foreach ($pattern in $script:allowedPatterns) {
            $rest = [regex]::Replace($rest, $pattern, "")
        }
        return @([regex]::Matches($rest, $script:internalTerms) | ForEach-Object { $_.Value })
    }

    function getScriptStrings {
        # .ps1 の文字列を "<相対パス>:<行>" と値の組で返す（同じ所の重複は除く）
        param ([string]$root)

        $seen = @{}
        foreach ($file in Get-ChildItem -LiteralPath $root -Recurse -Filter *.ps1) {
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$null)
            $nodes = $ast.FindAll({
                    param ($node)
                    $node -is [System.Management.Automation.Language.StringConstantExpressionAst] -or
                    $node -is [System.Management.Automation.Language.ExpandableStringExpressionAst]
                }, $true)
            foreach ($node in $nodes) {
                $value = [string]$node.Value
                $key = "$($file.FullName.Substring($root.Length + 1)):$($node.Extent.StartLineNumber): $value"
                if (!$seen.ContainsKey($key)) { $seen[$key] = $value }
            }
        }
        return $seen
    }
}

Describe "文言に出さない言葉" -Tag Meta {
    It "<Case>: <Expected>" -TestCases @(
        @{ Case = "内部の言葉があれば見つける（集約ファイル）"; Text = "集約ファイルに入れていません"; Expected = "集約ファイル" }
        @{ Case = "内部の言葉があれば見つける（システムインデックス）"; Text = "システムインデックスを作ります"; Expected = "システムインデックス" }
        @{ Case = "内部の言葉があれば見つける（本文インデックス）"; Text = "本文インデックスを読みます"; Expected = "本文インデックス" }
        @{ Case = "対象を Office だけに見せる言葉があれば見つける（Office ファイル）"; Text = "Office ファイルのあるフォルダ"; Expected = "Office ファイル" }
        @{ Case = "空白の無い書き方・全角空白も見つける（Officeファイル）"; Text = "Officeファイル"; Expected = "Officeファイル" }
        @{ Case = "全角空白の書き方も見つける"; Text = "Office　ファイル"; Expected = "Office　ファイル" }
        @{ Case = "「ファイル」「Office・テキスト」は見つけない"; Text = "ファイルのあるフォルダ（Office・テキスト）"; Expected = "" }
        @{ Case = "「インデックス」だけなら見つけない"; Text = "インデックスを作成しています"; Expected = "" }
        @{ Case = "許す例外（前の版の txt の名前のパターン）は見つけない"; Text = "システムインデックス*.txt"; Expected = "" }
        @{ Case = "パターンの形でない txt の名前は見つける"; Text = "システムインデックス_1.txt"; Expected = "システムインデックス" }
        @{ Case = "許す例外（画面の「集約ファイル N 件」）は見つけない"; Text = "集約ファイル `$(`$summary['Count'].ToString('N0')) 件"; Expected = "" }
        @{ Case = "許す例外と同じ文字列に別の内部の言葉があれば見つける"; Text = "システムインデックス*.txt と集約ファイル"; Expected = "集約ファイル" }
    ) {
        param ($Text, $Expected)
        (findInternalTerms $Text) -join "," | Should -Be $Expected
    }

    It "scripts の文字列に、内部の言葉が無い" {
        $found = New-Object System.Collections.Generic.List[string]
        $strings = getScriptStrings $scriptsDir
        foreach ($key in $strings.Keys) {
            if ((findInternalTerms $strings[$key]).Count -gt 0) { $found.Add($key) }
        }
        (@($found | Sort-Object) -join "`n") | Should -Be ""
    }

    It "xaml の属性の値・文字に、内部の言葉が無い" {
        $found = New-Object System.Collections.Generic.List[string]
        foreach ($file in Get-ChildItem -LiteralPath $scriptsDir -Recurse -Filter *.xaml) {
            $xml = [xml]([System.IO.File]::ReadAllText($file.FullName))
            $values = @($xml.SelectNodes('//@*') | ForEach-Object { $_.Value }) + @($xml.SelectNodes('//text()') | ForEach-Object { $_.Value })
            foreach ($value in $values) {
                if ((findInternalTerms $value).Count -gt 0) { $found.Add("$($file.Name): $value") }
            }
        }
        (@($found | Sort-Object) -join "`n") | Should -Be ""
    }
}
