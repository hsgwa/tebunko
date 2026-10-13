# フォルダ構成の決まりごとのテスト（文脈と層の分け方を保つ）。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    # 読み込み口からのたどり方は tools\script_rules.ps1 と共有する（check_release_package.ps1・
    # new_single_script.ps1 も同じたどり方を使う）
    . "$PSScriptRoot\..\..\tools\script_rules.ps1"
    # 関数の呼び出し・変数・型の使い方を AST で読む部品（解析は重いので、使うときに 1 回だけ行って使い回す）
    . "$PSScriptRoot\..\..\tools\script_arch.ps1"
    function getScriptArch {
        if (-not $script:scriptArch) { $script:scriptArch = newScriptArch $scriptsDir }
        return $script:scriptArch
    }

    # getSourcedFiles は getDotSourceTargets（存在しない先も返す）の、このテスト向けの薄い別名。
    # ui\ 配下の "$TebunkoDir\..." の形は、起動口 gui.ps1 から渡される tebunko\ 直下を指す
    function getSourcedFiles {
        param ([string]$path)
        return @(getDotSourceTargets $path "${scriptsDir}\tebunko" | Where-Object { Test-Path -LiteralPath $_ })
    }
}

Describe "依存の向き" -Tag Meta {
    # shared はどのツールからも使う部品。ツール（scripts\tebunko など）のフォルダを知っていてはいけない。
    # 製品の名前 tebunko は shared でも使う（前の版の一時フォルダ %TEMP%\tebunko の片付けなど）ため、ツールのフォルダを指す書き方だけを探す
    It "shared 配下にツールのフォルダが出てこない" {
        $toolNames = @(Get-ChildItem "${scriptsDir}" -Directory | Where-Object { $_.Name -ne "shared" } | ForEach-Object { [regex]::Escape($_.Name) })
        $pattern = "\.\.\\(" + ($toolNames -join "|") + ")\\|scripts[\\/](" + ($toolNames -join "|") + ")\b"
        $found = @(Get-ChildItem "${scriptsDir}\shared" -Recurse -Include *.ps1, *.xaml |
            Select-String -Pattern $pattern |
            ForEach-Object { "$($_.Filename):$($_.LineNumber)" })
        ($found -join ", ") | Should -Be ""
    }

    It "ツール同士は互いを読み込まない" {
        $tools = @(Get-ChildItem "${scriptsDir}" -Directory | Where-Object { $_.Name -ne "shared" })
        foreach ($tool in $tools) {
            $others = @($tools | Where-Object { $_.Name -ne $tool.Name } | ForEach-Object { $_.Name })
            if ($others.Count -eq 0) { continue }
            foreach ($file in (Get-ChildItem $tool.FullName -Recurse -Filter "*.ps1")) {
                foreach ($sourced in (getSourcedFiles $file.FullName)) {
                    foreach ($other in $others) {
                        ($sourced -like "*\scripts\$other\*") | Should -Be $false
                    }
                }
            }
        }
    }
}

Describe "読み込み漏れ" -Tag Meta {
    # 起動口からたどれないファイルは、足したのに読み込み忘れている
    It "すべての .ps1 が起動口からたどれる" {
        $entries = @("${scriptsDir}\tebunko\gui.ps1", "${scriptsDir}\tebunko\indexer.ps1")
        $seen = getReachableFiles $entries "${scriptsDir}\tebunko"
        $all = @(Get-ChildItem "${scriptsDir}" -Recurse -Filter "*.ps1" | ForEach-Object { $_.FullName })
        $missing = @($all | Where-Object { !$seen.Contains($_) } | ForEach-Object { Split-Path $_ -Leaf })
        ($missing -join ", ") | Should -Be ""
    }
}

Describe "判断層" -Tag Meta {
    # 判断層（入力は素の値、出力は素の値）は画面に触らない。触らないからテストが書ける
    It "判断層のファイルに画面への依存が無い" {
        $files = @(
            "${scriptsDir}\shared\core\text.ps1"
            "${scriptsDir}\tebunko\index\index_name.ps1"
            "${scriptsDir}\tebunko\index\content_index_format.ps1"
            "${scriptsDir}\tebunko\search\search_query.ps1"
            "${scriptsDir}\tebunko\search\search_gram.ps1"
            "${scriptsDir}\tebunko\indexer\indexer_decide.ps1"
        ) + @(Get-ChildItem "${scriptsDir}" -Recurse -Filter "*_view.ps1" | ForEach-Object { $_.FullName })
        foreach ($file in $files) {
            $text = [System.IO.File]::ReadAllText($file)
            $hit = [regex]::Matches($text, '\$ui\.|\$window|System\.Windows\.Media')
            "$(Split-Path $file -Leaf): $($hit.Count)" | Should -Be "$(Split-Path $file -Leaf): 0"
        }
    }

    # 上の文字列の検査と同じファイルを、構文木で見る。コメント・文字列の中の名前は数えず、型の式・引数の型は数える
    It "判断層のファイルが画面の変数・型に触らない（構文木）" {
        $files = @(
            "shared/core/text.ps1"
            "tebunko/index/index_name.ps1"
            "tebunko/index/content_index_format.ps1"
            "tebunko/search/search_query.ps1"
            "tebunko/search/search_gram.ps1"
            "tebunko/indexer/indexer_decide.ps1"
            "*_view.ps1"
        )
        $found = @(findTypeUses (getScriptArch) -In $files -Variable ui, window -TypeLike "System.Windows.*")
        ($found -join ", ") | Should -Be ""
    }
}

Describe "状態層と読み込み口" -Tag Meta {
    # 状態層（ファイル・COM を読み書きする層）は画面に触らない。フォルダで決めるので、足したファイルも対象になる
    It "状態層のファイルに画面への依存が無い" {
        $dirs = @(
            "${scriptsDir}\tebunko\core"
            "${scriptsDir}\tebunko\index"
            "${scriptsDir}\tebunko\indexer"
            "${scriptsDir}\tebunko\search"
            "${scriptsDir}\shared\core"
            "${scriptsDir}\shared\office"
        )
        $found = @(Get-ChildItem $dirs -Recurse -Filter "*.ps1" |
            Select-String -Pattern '\$ui\b|\$window\b|System\.Windows\b|PresentationFramework' |
            ForEach-Object { "$($_.Filename):$($_.LineNumber)" })
        ($found -join ", ") | Should -Be ""
    }

    # 上の文字列の検査と同じフォルダを、構文木で見る（コメント・文字列の中の名前は数えず、型の式・引数の型は数える）
    It "状態層のファイルが画面の変数・型に触らない（構文木）" {
        $dirs = @(
            "tebunko/core/*"
            "tebunko/index/*"
            "tebunko/indexer/*"
            "tebunko/search/*"
            "shared/core/*"
            "shared/office/*"
        )
        $found = @(findTypeUses (getScriptArch) -In $dirs -Variable ui, window -TypeLike "System.Windows.*")
        ($found -join ", ") | Should -Be ""
    }

    # 画面以外の読み込み口から ui/ をたどれると、インデックス作成などの画面の無い起動口が画面を読み込んでしまう
    It "画面以外の読み込み口から ui/ のファイルをたどれない" {
        $entries = @(
            "${scriptsDir}\shared\shared.ps1"
            "${scriptsDir}\tebunko\lib.ps1"
            "${scriptsDir}\tebunko\indexer\indexer_lib.ps1"
            "${scriptsDir}\tebunko\indexer.ps1"
        )
        $reached = getReachableFiles $entries "${scriptsDir}\tebunko"
        $ui = @($reached | Where-Object { $_ -like "*\ui\*" } | ForEach-Object { Split-Path $_ -Leaf })
        ($ui -join ", ") | Should -Be ""
    }
}

# 依存の向きを、ファイルの組で見る（関数の呼び出しと、別のファイルが最上位で作る変数の読み。tools\script_arch.ps1）。
# 全部のファイルが同じスコープに読み込まれるため、文字列を探す検査では、下の層が上の層の関数・変数を名前で使っても見つからない。
# 表の 1 行は「from のパスのファイルは、to のパスのファイルの関数・変数に依存しない」。
# 向きの決め方は docs/design/structure/source.md の「フォルダの分け方」にある
Describe "依存の向き（ファイル単位）" -Tag Meta {
    $rules = @(
        # shared はツールを知らない
        @{ from = "shared/*"; to = "tebunko/*"; except = "" }
        # 判断層・状態層は画面層に依存しない
        @{ from = "tebunko/core/*"; to = "*/ui/*"; except = "" }
        @{ from = "tebunko/index/*"; to = "*/ui/*"; except = "" }
        @{ from = "tebunko/indexer/*"; to = "*/ui/*"; except = "" }
        @{ from = "tebunko/search/*"; to = "*/ui/*"; except = "" }
        @{ from = "*_view.ps1"; to = "*/ui/*"; except = "*_view.ps1" }
        @{ from = "*/index_name.ps1"; to = "*/ui/*"; except = "" }
        @{ from = "*/search_query.ps1"; to = "*/ui/*"; except = "" }
        @{ from = "*/text.ps1"; to = "*/ui/*"; except = "" }
        # 共有の状態層も、共有の画面層に依存しない
        @{ from = "shared/core/*"; to = "shared/ui/*"; except = "" }
        @{ from = "shared/office/*"; to = "shared/ui/*"; except = "" }
        # 状態層の中のフォルダは、下から上へ依存しない（indexer と search は同じ高さで、互いに依存しない。部品の -From・-To は一方向なので 2 行で書く）
        @{ from = "tebunko/core/*"; to = "tebunko/index/*"; except = "" }
        @{ from = "tebunko/core/*"; to = "tebunko/indexer/*"; except = "" }
        @{ from = "tebunko/core/*"; to = "tebunko/search/*"; except = "" }
        @{ from = "tebunko/index/*"; to = "tebunko/indexer/*"; except = "" }
        @{ from = "tebunko/index/*"; to = "tebunko/search/*"; except = "" }
        @{ from = "tebunko/indexer/*"; to = "tebunko/search/*"; except = "" }
        @{ from = "tebunko/search/*"; to = "tebunko/indexer/*"; except = "" }
        @{ from = "shared/core/*"; to = "shared/office/*"; except = "" }
    )

    BeforeAll {
        # 今ある依存を、「from のファイル・to のファイル・名前」の 3 つで許す。ファイルの組だけで許すと、同じ組の別の依存も通ってしまうため。
        # 名前は関数の名前か、$ を付けた変数の名前。Reason は、なぜ今は許しているか。直したら、この一覧から消す
        # （どの行でも検出されなくなった例外は、下の It が失敗にする）。この一覧に足して通すことはしない
        function newAllowances {
            param ([string]$from, [string]$to, [string[]]$names, [string]$reason)

            foreach ($name in $names) {
                @{ From = $from; To = $to; Name = $name; Reason = $reason }
            }
        }

        $toolKnown = "共有の部品が、ツールの側で定義した関数・変数を使っている。決めた形ではない今ある依存"
        $upward = "状態層の中で、下のフォルダが上のフォルダの関数・変数を使っている。関数を下のフォルダへ移すまでの今ある依存"
        $script:allowances = @(
            # shared → tebunko
            newAllowances "shared/core/folder.ps1" "tebunko/search/source_map.ps1" @("joinSourcePath") $toolKnown
            newAllowances "shared/core/fs.ps1" "tebunko/core/settings.ps1" @('$appId') $toolKnown
            newAllowances "shared/office/office_reader.ps1" "tebunko/index/index_name.ps1" @("toIndexFileName") $toolKnown
            newAllowances "shared/ui/shell.ps1" "tebunko/ui/startup_error.ps1" @('$appTitle') $toolKnown
            # 判断層 → 画面層
            newAllowances "tebunko/ui/search/search_bar_view.ps1" "shared/ui/shell.ps1" @("formatTime") "判断層が、共有の画面層の関数を呼んでいる今ある依存"
            # core → index
            newAllowances "tebunko/core/settings.ps1" "tebunko/index/index_name.ps1" @("testSourceNameRecordable") $upward
            newAllowances "tebunko/core/workspace.ps1" "tebunko/index/system_index.ps1" @("updateSystemIndexState") $upward
            newAllowances "tebunko/core/workspace_mover.ps1" "tebunko/index/index_store.ps1" @("getIndexNameMap") $upward
            # index → indexer（状態ファイル・ログの読み書きが indexer にある）
            newAllowances "tebunko/index/content_index_store.ps1" "tebunko/indexer/indexer_state.ps1" @("writeIndexerLog") $upward
            newAllowances "tebunko/index/system_index.ps1" "tebunko/indexer/indexer_state.ps1" @("writeIndexerLog") $upward
            newAllowances "tebunko/index/index_archive.ps1" "tebunko/indexer/indexer_state.ps1" @("newStatusRow", "readStatusFile", "toStatusLine", "writeStatusFile") $upward
            # index → search（取り込み一覧・情報源フォルダのファイルの読み書きと、システムインデックスの gram の関数が search にある）
            newAllowances "tebunko/index/content_index_store.ps1" "tebunko/search/search_query.ps1" @("truncateHitLine") $upward
            newAllowances "tebunko/index/index_archive.ps1" "tebunko/search/source_map.ps1" @("getSourceFolderMap", "readSourceFolderFile") $upward
            newAllowances "tebunko/index/index_store.ps1" "tebunko/search/source_map.ps1" @("getSourceFolderMap", "readSourceFolderFile") $upward
            newAllowances "tebunko/index/index_catalog.ps1" "tebunko/search/source_map.ps1" @("writeSourceFolderFile") $upward
            newAllowances "tebunko/index/system_index.ps1" "tebunko/search/search_gram.ps1" @(
                '$systemIndexFileName', "addTextGrams", "convertFromSystemIndexState", "convertToGramText", "convertToSystemIndexState",
                "getGramPartCount", "getIndexNameOfRelPath", "getSystemIndexFileNames", "newSystemIndexState", "testSystemIndexPath"
            ) $upward
            # indexer → search
            newAllowances "tebunko/indexer/indexer_plan.ps1" "tebunko/search/search_gram.ps1" @('$systemIndexFileName') $upward
            newAllowances "tebunko/indexer/indexer_run.ps1" "tebunko/search/source_map.ps1" @("writeSourceFolderFile") $upward
        )

        # 行ごとの検出の結果。表の行ごとの It と、使われなくなった例外の It が、同じ解析を繰り返さないよう覚えておく
        $script:foundByRule = @{}
        function getFoundItems {
            param ([string]$from, [string]$to, [string]$except)

            $key = "$from|$to|$except"
            if (-not $script:foundByRule.ContainsKey($key)) {
                $script:foundByRule[$key] = @(findFileDependencyItems (getScriptArch) -From $from -To $to -ToExcept $except)
            }
            return , $script:foundByRule[$key]
        }
    }

    # 失敗の文には、ファイルの組と関数・変数の名前が出る（例: shared/core/folder.ps1 -> tebunko/search/source_map.ps1（joinSourcePath））
    It "<from> は <to> に依存しない" -TestCases $rules {
        param ($from, $to, $except)

        $left = @(removeAllowedDependencies (getFoundItems $from $to $except) $script:allowances)
        (@(formatFileDependencies $left) -join ", ") | Should -Be "" -Because "新しい依存は入れない。今ある依存を許すなら、このファイルの例外の一覧に理由つきで足す。向きの決まりを変えるなら docs/design/structure/source.md の「フォルダの分け方」も直す"
    }

    # 表の from・to が、どのファイルにも当たらないと、上の It は何も見ずに通ってしまう（フォルダやファイルの名前を変えたとき）
    It "<from> -> <to> の両方に、当たるファイルがある" -TestCases $rules {
        param ($from, $to, $except)

        $paths = @((getScriptArch).Files | ForEach-Object { $_.Path })
        @($paths | Where-Object { testPathLike $_ @($from) }).Count | Should -BeGreaterThan 0 -Because "from に当たるファイルが無い。表のパスを、名前を変えたフォルダ・ファイルに合わせる"
        @($paths | Where-Object { (testPathLike $_ @($to)) -and -not (testPathLike $_ @($except)) }).Count | Should -BeGreaterThan 0 -Because "to に当たるファイルが無い。表のパスを、名前を変えたフォルダ・ファイルに合わせる"
    }

    # 直した依存の例外が残ると、同じ依存を後から足しても通ってしまう。全部の行の結果の和と突き合わせる
    # （1 行ずつ突き合わせると、別の行で出る依存の例外を、使われないとしてしまう）
    It "使われなくなった例外が無い" -TestCases @(@{ rules = $rules }) {
        param ($rules)

        $all = @()
        foreach ($rule in $rules) { $all += (getFoundItems $rule.from $rule.to $rule.except) }
        $unused = @(findUnusedAllowances $script:allowances $all | ForEach-Object { "$($_.From) -> $($_.To)（$($_.Name)）" })
        ($unused -join ", ") | Should -Be "" -Because "直した依存の例外は、このファイルの例外の一覧から消す"
    }

    It "例外のそれぞれに理由がある" {
        $noReason = @($script:allowances | Where-Object { -not $_.Reason } | ForEach-Object { "$($_.From)（$($_.Name)）" })
        ($noReason -join ", ") | Should -Be ""
    }
}