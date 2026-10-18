# フォルダ構成の決まりごとのテスト（文脈と層の分け方を保つ）。
# 層に属するファイルの一覧（判断層・状態層）は tests/helpers/layer_paths.ps1 の 1 か所で決める
BeforeDiscovery {
    . "$PSScriptRoot\..\helpers\layer_paths.ps1"
}

BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\..\helpers\layer_paths.ps1"
    # 読み込み口からのたどり方は tools\script_rules.ps1 と共有する（check_release_package.ps1・
    # new_single_script.ps1 も同じたどり方を使う）
    . "$PSScriptRoot\..\..\tools\script_rules.ps1"
    # 関数の呼び出し・変数・型の使い方を AST で読む部品（解析は重いので、使うときに 1 回だけ行って使い回す）
    . "$PSScriptRoot\..\..\tools\script_arch.ps1"
    # 名前を文字列の並びで受け取る関数。別のスレッドへ渡す関数・変数の名前（位置引数の順）を、依存として数える（script_arch.ps1 は tebunko を知らないので、ここで教える）
    $nameListCommands = @{
        newWorkerState = @(
            @{ Kind = "Function"; Parameter = "functions" }
            @{ Kind = "Variable"; Parameter = "variables" }
        )
    }
    function getScriptArch {
        if (-not $script:scriptArch) { $script:scriptArch = newScriptArch $scriptsDir $nameListCommands }
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
        $files = @((getScriptArch).Files | Where-Object { testPathLike $_.Path $judgmentLayerPaths } | ForEach-Object { "${scriptsDir}\$($_.Path -replace '/', '\')" })
        foreach ($file in $files) {
            $text = [System.IO.File]::ReadAllText($file)
            $hit = [regex]::Matches($text, '\$ui\.|\$window|System\.Windows\.Media')
            "$(Split-Path $file -Leaf): $($hit.Count)" | Should -Be "$(Split-Path $file -Leaf): 0"
        }
    }

    # 上の文字列の検査と同じファイルを、構文木で見る。コメント・文字列の中の名前は数えず、型の式・引数の型は数える
    It "判断層のファイルが画面の変数・型に触らない（構文木）" {
        $found = @(findTypeUses (getScriptArch) -In $judgmentLayerPaths -Variable ui, window -TypeLike "System.Windows.*")
        ($found -join ", ") | Should -Be ""
    }

    # 一覧のパスがどのファイルにも当たらないと、上の It は何も見ずに通ってしまう（ファイルの名前を変えたとき）
    It "判断層の一覧の <path> に当たるファイルがある" -TestCases @($judgmentLayerPaths | ForEach-Object { @{ path = $_ } }) {
        param ($path)

        @((getScriptArch).Files | Where-Object { testPathLike $_.Path @($path) }).Count | Should -BeGreaterThan 0 -Because "tests/helpers/layer_paths.ps1 の判断層の一覧を、名前を変えたファイルに合わせる"
    }

    # 判断層の並びは、tests/helpers/layer_paths.ps1 を正とし、AGENTS.md と docs/design/structure/source.md の層の表に同じ並びで書く
    It "AGENTS.md と source.md の判断層の並びが、一覧と同じ" {
        $root = Split-Path $scriptsDir -Parent
        $expected = @($judgmentLayerPaths | ForEach-Object { ($_ -split '/')[-1] } | Sort-Object)
        $agents = @(Get-Content -LiteralPath "$root\AGENTS.md" -Encoding UTF8 | Where-Object { $_ -like '- 判断層（*' })
        $agents.Count | Should -Be 1 -Because "AGENTS.md の「ソースの分け方」に、「- 判断層（…）は」で始まる行が 1 つ要る"
        $source = @(Get-Content -LiteralPath "$root\docs\design\structure\source.md" -Encoding UTF8 | Where-Object { $_ -like '| 判断層（入力は素の値*' })
        $source.Count | Should -Be 1 -Because "source.md の層の表に、判断層の行が 1 つ要る"
        $agentsHead = ($agents[0] -split '）')[0]
        $sourceCell = ($source[0] -split '\|')[2]
        $sourceHead = $sourceCell.Substring(0, $sourceCell.IndexOf('*_view.ps1') + '*_view.ps1`'.Length)
        foreach ($head in @($agentsHead, $sourceHead)) {
            $names = @([regex]::Matches($head, '`([^`]+\.ps1)`') | ForEach-Object { $_.Groups[1].Value } | Sort-Object)
            ($names -join ", ") | Should -Be ($expected -join ", ") -Because "判断層の並びを tests/helpers/layer_paths.ps1 に合わせる"
        }
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
        $found = @(findTypeUses (getScriptArch) -In $stateLayerPaths -Variable ui, window -TypeLike "System.Windows.*")
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
        # 判断層のファイルは、tests/helpers/layer_paths.ps1 の一覧から 1 つずつ行にする（画面層・判断層でない状態層への依存を、下で足す）
        @{ from = "*_view.ps1"; to = "*/ui/*"; except = "*_view.ps1" }
        foreach ($path in $judgmentLayerPaths) {
            if ($path -ne "*_view.ps1") { @{ from = $path; to = "*/ui/*"; except = "" } }
        }
        # 判断層は、判断層でない状態層（ファイル・COM を読み書きする層）に依存しない。判断層どうしの依存はよい
        foreach ($path in $judgmentLayerPaths) {
            @{ from = $path; to = $stateLayerPaths -join ","; except = $judgmentLayerPaths -join "," }
        }
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
        # Kind は、判断層 → 判断層でない状態層の例外だけに付ける（下の It が確かめる）。
        #   境界: 判断層が状態層の関数を呼ぶ境界。受け渡しは tests/tebunko/ui/layer_boundary.Tests.ps1 が、その関数をスタブ（Mock）にして確かめる
        #   移すと消える: ファイル・COM に触らない値・関数が、状態層のファイルにある。判断層のファイルへ移すと依存ではなくなる
        function newAllowances {
            param ([string]$from, [string]$to, [string[]]$names, [string]$reason, [string]$kind = "")

            foreach ($name in $names) {
                @{ From = $from; To = $to; Name = $name; Reason = $reason; Kind = $kind }
            }
        }

        $script:allowances = @(
            # index → search（取り込み一覧・情報源フォルダのファイルの読み書きと、gram の関数が search にある）
            newAllowances "tebunko/index/content_index_store.ps1" "tebunko/search/search_query.ps1" @("truncateHitLine") "ヒットの行を切り詰める truncateHitLine を、定数（hitLineMaxChars・hitLineBeforeMatchChars）と一緒に core へ移すと消える（文字列だけを扱う）"
            newAllowances "tebunko/index/index_archive.ps1" "tebunko/search/source_map.ps1" @("getSourceFolderMap", "readSourceFolderFile") "取り込み一覧・情報源フォルダのファイルを読む関数を index へ移すと消える（ファイルの持ち主は index）"
            newAllowances "tebunko/index/index_store.ps1" "tebunko/search/source_map.ps1" @("getSourceFolderMap", "readSourceFolderFile") "取り込み一覧・情報源フォルダのファイルを読む関数を index へ移すと消える（ファイルの持ち主は index）"
            newAllowances "tebunko/index/index_catalog.ps1" "tebunko/search/source_map.ps1" @("writeSourceFolderFile") "情報源フォルダのファイルを書く writeSourceFolderFile を index へ移すと消える（ファイルの持ち主は index）"
            newAllowances "tebunko/index/gram_index_store.ps1" "tebunko/search/search_gram.ps1" @("addTextGrams") "2-gram を数える addTextGrams を index へ移すと消える（search_gram.ps1 の関数）"
            # indexer → search
            newAllowances "tebunko/indexer/indexer_run.ps1" "tebunko/search/source_map.ps1" @("writeSourceFolderFile") "情報源フォルダのファイルを書く writeSourceFolderFile を index へ移すと消える（ファイルの持ち主は index）"
            # index → indexer（クラス）
            newAllowances "tebunko/index/index_catalog.ps1" "tebunko/indexer/indexer_state.ps1" @("[StatusLedger]") "取り込み一覧の台帳のクラス StatusLedger を index へ移すと消える（台帳は index が使う）"
            # 判断層 → 画面層（クラス）
            newAllowances "tebunko/ui/search/result_list_view.ps1" "tebunko/ui/search/result_types.ps1" @("[FileGroup]", "[HitRow]") "結果の行のクラス FileGroup・HitRow を、画面に触らない判断層のファイルへ移すと消える"
            # 判断層 → 判断層でない状態層。今ある分。境界は layer_boundary.Tests.ps1 で受け渡しを確かめ、移すと消えるものは判断層のファイルへ移すと消える
            $moved = "ファイル・COM に触らない値・関数が状態層のファイルにある。判断層のファイルへ移すと消える"
            newAllowances "tebunko/index/content_index_format.ps1" "shared/core/pdf_file.ps1" @("testPdfExtension") $moved "移すと消える"
            newAllowances "tebunko/index/content_index_format.ps1" "shared/core/text_file.ps1" @("testTextExtension") $moved "移すと消える"
            newAllowances "tebunko/index/index_name.ps1" "shared/core/folder.ps1" @("getFolderLeafName") $moved "移すと消える"
            newAllowances "tebunko/index/index_name.ps1" "shared/core/pdf_file.ps1" @('$pdfExtensions', "testPdfExtension") $moved "移すと消える"
            newAllowances "tebunko/index/index_name.ps1" "shared/core/text_file.ps1" @('$textExtensions', "testTextExtension") $moved "移すと消える"
            newAllowances "tebunko/indexer/indexer_decide.ps1" "tebunko/core/paths.ps1" @('$stateDone', '$stateFailed', '$stateNew') $moved "移すと消える"
            newAllowances "tebunko/search/search_query.ps1" "shared/core/pdf_file.ps1" @('$pdfExtensions') $moved "移すと消える"
            newAllowances "tebunko/search/search_query.ps1" "shared/core/text_file.ps1" @('$textExtensions') $moved "移すと消える"
            newAllowances "tebunko/search/search_query.ps1" "tebunko/core/settings.ps1" @('$fileKindNames') $moved "移すと消える"
            newAllowances "tebunko/ui/index/index_archive_view.ps1" "shared/core/folder.ps1" @("normalizeFolderPath", "testFolderUnder", "testSameFolder") $moved "移すと消える"
            newAllowances "tebunko/ui/index/index_archive_view.ps1" "tebunko/index/index_archive_rules.ps1" @('$importCollisionOverwrite', '$importCollisionRename', "getExportFileName") $moved "移すと消える"
            newAllowances "tebunko/ui/index/index_archive_view.ps1" "tebunko/index/index_archive_rules.ps1" @("getIndexFolderConflict") "インデックスのフォルダの重なりの判定を、状態層の関数（index_archive_rules.ps1）に任せる境界" "境界"
            newAllowances "tebunko/ui/index/index_detail_view.ps1" "shared/core/folder.ps1" @("normalizeFolderPath") $moved "移すと消える"
            newAllowances "tebunko/ui/index/index_edit_view.ps1" "shared/core/folder.ps1" @("normalizeFolderPath") $moved "移すと消える"
            newAllowances "tebunko/ui/index/index_edit_view.ps1" "tebunko/index/index_archive_rules.ps1" @("getIndexFolderConflict") "インデックスのフォルダの重なりの判定を、状態層の関数（index_archive_rules.ps1）に任せる境界" "境界"
            newAllowances "tebunko/ui/index/index_list_view.ps1" "shared/core/fs.ps1" @('$pathStateFound', '$pathStateOther', '$pathStateUnreachable') $moved "移すと消える"
            newAllowances "tebunko/ui/indexing_view.ps1" "tebunko/core/paths.ps1" @('$planKindDropped', '$planKindIngest', '$planKindMissing', '$planKindUnchecked') $moved "移すと消える"
            newAllowances "tebunko/ui/indexing_view.ps1" "tebunko/core/settings.ps1" @('$cloudFilesAsk', '$cloudFilesDownload') $moved "移すと消える"
            newAllowances "tebunko/ui/leftover_view.ps1" "shared/office/office_process.ps1" @('$officeProcessNames') $moved "移すと消える"
            newAllowances "tebunko/ui/search/open_source_view.ps1" "tebunko/core/settings.ps1" @('$openModeReadOnly') $moved "移すと消える"
            newAllowances "tebunko/ui/search/result_list_view.ps1" "tebunko/core/settings.ps1" @('$openModeNew', '$openModeNormal', '$openModeReadOnly') $moved "移すと消える"
            newAllowances "tebunko/ui/search/search_bar_view.ps1" "tebunko/core/settings.ps1" @('$fileKindNames') $moved "移すと消える"
            newAllowances "tebunko/ui/settings/settings_view.ps1" "shared/core/folder.ps1" @("normalizeFolderPath", "testFolderUnder", "testLocalDrivePath", "testSameFolder") $moved "移すと消える"
            newAllowances "tebunko/ui/settings/settings_view.ps1" "tebunko/core/settings.ps1" @("testSettingsFileName") "設定ファイルの名前の判定を、状態層の関数（settings.ps1）に任せる境界" "境界"
            newAllowances "tebunko/ui/settings/settings_view.ps1" "tebunko/core/workspace.ps1" @("[Workspace]") "ワークスペースのフォルダ名を組み立てるクラス Workspace（ファイルに触らない）の、フォルダ名の部分を判断層のファイルへ移すと消える" "移すと消える"
        )

        # 行ごとの検出の結果。表の行ごとの It と、使われなくなった例外の It が、同じ解析を繰り返さないよう覚えておく
        $script:foundByRule = @{}
        function getFoundItems {
            param ([string]$from, [string]$to, [string]$except)

            $key = "$from|$to|$except"
            if (-not $script:foundByRule.ContainsKey($key)) {
                # 複数のパスは , でつないだ 1 つの文字列で書く（表の行の名前に出すため）
                $script:foundByRule[$key] = @(findFileDependencyItems (getScriptArch) -From ($from -split ',') -To ($to -split ',') -ToExcept ($except -split ','))
            }
            return , $script:foundByRule[$key]
        }
    }

    # 失敗の文には、ファイルの組と関数・変数の名前が出る（例: tebunko/index/index_archive.ps1 -> tebunko/search/source_map.ps1（getSourceFolderMap））
    It "<from> は <to> に依存しない" -TestCases $rules {
        param ($from, $to, $except)

        $left = @(removeAllowedDependencies (getFoundItems $from $to $except) $script:allowances)
        (@(formatFileDependencies $left) -join ", ") | Should -Be "" -Because "新しい依存は入れない。今ある依存を許すなら、このファイルの例外の一覧に理由つきで足す。向きの決まりを変えるなら docs/design/structure/source.md の「フォルダの分け方」も直す"
    }

    # 表の from・to が、どのファイルにも当たらないと、上の It は何も見ずに通ってしまう（フォルダやファイルの名前を変えたとき）
    It "<from> -> <to> の両方に、当たるファイルがある" -TestCases $rules {
        param ($from, $to, $except)

        $paths = @((getScriptArch).Files | ForEach-Object { $_.Path })
        @($paths | Where-Object { testPathLike $_ ($from -split ',') }).Count | Should -BeGreaterThan 0 -Because "from に当たるファイルが無い。表のパスを、名前を変えたフォルダ・ファイルに合わせる"
        @($paths | Where-Object { (testPathLike $_ ($to -split ',')) -and -not (testPathLike $_ ($except -split ',')) }).Count | Should -BeGreaterThan 0 -Because "to に当たるファイルが無い。表のパスを、名前を変えたフォルダ・ファイルに合わせる"
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

    # 判断層 → 判断層でない状態層の例外は、境界（受け渡しをテストで確かめる）か、移すと消えるもののどちらか
    It "判断層から状態層への例外に、種類（境界・移すと消える）がある" {
        $noKind = @($script:allowances | Where-Object {
                (testPathLike $_.From $judgmentLayerPaths) -and (testPathLike $_.To $stateLayerPaths) -and -not (testPathLike $_.To $judgmentLayerPaths) -and
                (@("境界", "移すと消える") -notcontains $_.Kind)
            } | ForEach-Object { "$($_.From) -> $($_.To)（$($_.Name)）" })
        ($noKind -join ", ") | Should -Be "" -Because "newAllowances の最後の引数に、境界か移すと消えるを書く"
    }

    # 境界の理由で載せたのに、スタブにして確かめていないと、理由だけがあって受け渡しが確かめられていない
    It "境界の例外の関数は、layer_boundary.Tests.ps1 の Mock の対象に出る" {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile("$here\tebunko\ui\layer_boundary.Tests.ps1", [ref]$null, [ref]$null)
        $mocked = New-Object System.Collections.Generic.HashSet[string]
        foreach ($command in $ast.FindAll({ param ($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq "Mock" }, $true)) {
            # Mock <名前> {…} と、Mock -CommandName <名前> {…}
            $first = $command.CommandElements[1]
            if ($first -is [System.Management.Automation.Language.CommandParameterAst] -and $first.ParameterName -eq "CommandName") {
                $first = if ($first.Argument) { $first.Argument } else { $command.CommandElements[2] }
            }
            if ($first -is [System.Management.Automation.Language.StringConstantExpressionAst]) { [void]$mocked.Add([string]$first.Value) }
        }
        $notMocked = @($script:allowances | Where-Object { $_.Kind -eq "境界" -and -not $mocked.Contains($_.Name) } | ForEach-Object { "$($_.From)（$($_.Name)）" })
        ($notMocked -join ", ") | Should -Be "" -Because "境界の関数は、layer_boundary.Tests.ps1 でスタブにして受け渡しを確かめる。変数は境界にできない（移すと消える）。確かめていない関数は、移すと消えるにするか、境界のテストを足す"
    }
}

# 構文木から呼ぶ先が決まらない書き方（名前の文字列で関数を引く・実行する形）は、向きの検査では依存として数えられない。
# 今ある所をこの閉じた一覧に載せ、新しく増えたら失敗にする（tools/script_arch.ps1 の findStringNameCalls が形ごとに数える）
Describe "名前を文字列で渡す呼び出し" -Tag Meta {
    BeforeAll {
        $stringCall = "(名前の文字列)"
        $script:stringCallAllowances = @(
            @{ From = "shared/core/worker_pool.ps1"; To = $stringCall; Name = "Get-Command"; Count = 1
                Reason = "別のスレッドへ、関数を名前で写す（呼ぶ関数が同じスレッドに読み込まれていることは、ワーカーのテストが確かめる）。名前を渡す側の文字列は、向きの検査に数えない" }
            @{ From = "shared/office/office_app.ps1"; To = $stringCall; Name = "Get-Command"; Count = 1
                Reason = "別のスレッドへ、関数を名前で写す（呼ぶ関数が同じスレッドに読み込まれていることは、ワーカーのテストが確かめる）。名前を渡す側の文字列は、向きの検査に数えない" }
            @{ From = "shared/core/worker_pool.ps1"; To = $stringCall; Name = ".AddScript"; Count = 1
                Reason = "別のスレッドで、ワーカーのスクリプト本文を実行する。本文は、呼び出し側が渡す値で、ここで関数の名前を決めない" }
            @{ From = "shared/office/office_app.ps1"; To = $stringCall; Name = ".AddScript"; Count = 1
                Reason = "別のスレッドで、COM を扱うスクリプト本文を実行する。本文は、読み込み済みの関数を呼ぶだけで、ここで関数の名前を決めない" }
            @{ From = "tebunko/indexer/ingest_dispatcher.ps1"; To = $stringCall; Name = ".AddScript"; Count = 1
                Reason = "取り込みのワーカーで、スクリプトブロックを実行する。ブロックの中の呼び出しは、構文木が数える" }
            @{ From = "tebunko/indexer/indexing_session.ps1"; To = $stringCall; Name = ".AddScript"; Count = 1
                Reason = "取り込みの実行を別のスレッドで動かす。渡すのはスクリプトブロックで、中の呼び出しは構文木が数える" }
            @{ From = "tebunko/search/search_service.ps1"; To = $stringCall; Name = ".AddScript"; Count = 1
                Reason = "検索の実行を別のスレッドで動かす。渡すのはスクリプトブロックで、中の呼び出しは構文木が数える" }
            @{ From = "shared/core/worker_pool.ps1"; To = $stringCall; Name = "Get-Variable の名前が定数でない"; Count = 1
                Reason = "別のスレッドへ写す変数の名前を、引数で受け取る。写す変数の一覧は、呼び出し側が渡す" }
            @{ From = "tebunko/core/parts.ps1"; To = $stringCall; Name = "Get-Variable の名前が定数でない"; Count = 1
                Reason = "別のスレッドへ渡す、配布版の変数（bundledParts・bundledScriptPath・bundledVersion）を、同じ関数の中で決まった 3 つの名前の繰り返しで引く" }
        )
    }

    It "一覧に無い、名前を文字列で渡す呼び出しが無い" {
        $left = @(removeAllowedDependencies @(findStringNameCalls (getScriptArch)) $script:stringCallAllowances | ForEach-Object { "$($_.From):$($_.Line) $($_.Name)" })
        ($left -join ", ") | Should -Be "" -Because "名前を文字列で渡す呼び出しは、向きの検査から見えない。増やさない。要るなら、このファイルの一覧に理由つきで足す"
    }

    It "使われなくなった行が無い" {
        $unused = @(findUnusedAllowances $script:stringCallAllowances @(findStringNameCalls (getScriptArch)) | ForEach-Object { "$($_.From)（$($_.Name)）" })
        ($unused -join ", ") | Should -Be "" -Because "なくなった呼び出しの行は、一覧から消す"
    }

    It "どの行にも理由がある" {
        $noReason = @($script:stringCallAllowances | Where-Object { -not $_.Reason } | ForEach-Object { "$($_.From)（$($_.Name)）" })
        ($noReason -join ", ") | Should -Be ""
    }

    # 載っているファイルに同じ形を足しても気づけるよう、ファイルと形ごとの数を行に持たせる
    It "ファイルと形ごとの数が、一覧の行の数と同じ" {
        $actual = @{}
        foreach ($call in @(findStringNameCalls (getScriptArch))) {
            $key = "$($call.From) $($call.Name)"
            $actual[$key] = 1 + [int]$actual[$key]
        }
        $differ = @($script:stringCallAllowances | Where-Object { [int]$actual["$($_.From) $($_.Name)"] -ne $_.Count } |
                ForEach-Object { "$($_.From)（$($_.Name)）: 一覧 $($_.Count)・実際 $([int]$actual["$($_.From) $($_.Name)"])" })
        ($differ -join ", ") | Should -Be "" -Because "数が変わったら、その場所を確かめて、一覧の数と理由を直す"
    }
}