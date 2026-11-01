# 処理の記録（shared\core\trace.ps1）の区切りが、決めた場所に入っていることを確かめる（docs\design\testing\perf.md「処理の記録（trace）」の入れ所の表）。
# 区切りを消す・関数を作り替えるときに、記録が黙って欠けないようにする。
# 区切りを入れる場所: File（scripts\tebunko 以下のパス）・Function（囲む関数。ファイルの先頭から外れた処理は ""）・
# Category（区分）・Marks（その関数の中に markTrace "<区分>" "<名前の一部>" があること）・Writes（その関数の中に writeTrace "<区分>" があること）
${tracePoints} = @(
    @{ File = "ui\splash.ps1"; Function = "stepSplash"; Category = "startup"; Marks = @("stepSplash"); Writes = $false }
    @{ File = "ui\shell\startup.ps1"; Function = "loadStartupData"; Category = "startup"; Marks = @("06 "); Writes = $false }
    @{ File = "ui\index_tree.ps1"; Function = "applyIndexTreeData"; Category = "startup"; Marks = @("07 "); Writes = $false }
    @{ File = "ui\search\search_session.ps1"; Function = "markSearchSentTrace"; Category = "startup"; Marks = @("08 ", "09 "); Writes = $false }
    @{ File = "ui\search\search_session.ps1"; Function = "pumpSearch"; Category = "startup"; Marks = @("10 "); Writes = $false }
    @{ File = "ui\search\search_session.ps1"; Function = "finishSearch"; Category = "startup"; Marks = @("11 "); Writes = $true }
    @{ File = "ui\search\search_session.ps1"; Function = "markSearchSentTrace"; Category = "search"; Marks = @("送り"); Writes = $false }
    @{ File = "ui\search\search_session.ps1"; Function = "pumpSearch"; Category = "search"; Marks = @("最初の行"); Writes = $false }
    @{ File = "ui\search\search_session.ps1"; Function = "finishSearch"; Category = "search"; Marks = @("総数"); Writes = $true }
    @{ File = "search\search_run.ps1"; Function = "invokeSearchRequest"; Category = "search"; Marks = @("受け取り", "終わり"); Writes = $true }
    @{ File = "search\search_run.ps1"; Function = "invokeSearchRequestBody"; Category = "search"; Marks = @("gram 索引で絞り終え", "列挙の終わり", "searchContentIndex の終わり"); Writes = $false }
    @{ File = "ui\preview.ps1"; Function = "showDetail"; Category = "preview"; Marks = @("行を選ぶ"); Writes = $false }
    @{ File = "ui\preview.ps1"; Function = "applyDetail"; Category = "preview"; Marks = @("本文が入る"); Writes = $true }
    @{ File = "ui\preview.ps1"; Function = "startReadMore"; Category = "preview"; Marks = @("続きを読む: 始め"); Writes = $false }
    @{ File = "ui\preview.ps1"; Function = "applyReadMore"; Category = "preview"; Marks = @("続きを読む: 表に入る"); Writes = $true }
    @{ File = "ui\index\index_archive.ps1"; Function = "startIndexArchiveJob"; Category = "archive"; Marks = @("始め", "裏の仕事の終わり", "読み直しの終わり"); Writes = $true }
    @{ File = "ui\index\index_archive.ps1"; Function = "updateImportProgress"; Category = "archive"; Marks = @("最初の進み具合"); Writes = $false }
    @{ File = "ui\settings\settings.ps1"; Function = "startWorkspaceMove"; Category = "workspace"; Marks = @("移動の始め"); Writes = $false }
    @{ File = "ui\settings\settings.ps1"; Function = "finishWorkspaceMove"; Category = "workspace"; Marks = @("裏の移動の終わり", "画面の読み直しの終わり"); Writes = $true }
    @{ File = "ui\indexing_tab.ps1"; Function = "startIndexing"; Category = "index"; Marks = @("開始の押下"); Writes = $false }
    @{ File = "ui\indexing_tab.ps1"; Function = "confirmIndexingTargets"; Category = "index"; Marks = @("確認の画面を閉じる"); Writes = $false }
    @{ File = "ui\indexing_tab.ps1"; Function = "continueStartIndexing"; Category = "index"; Marks = @("インデクサのプロセスを起動"); Writes = $false }
    @{ File = "ui\indexing_tab.ps1"; Function = "updateIndexingProgress"; Category = "index"; Marks = @("最初の進み具合"); Writes = $false }
    @{ File = "ui\indexing_tab.ps1"; Function = "finishIndexing"; Category = "index"; Marks = @("終わりを受ける", "読み直しの終わり"); Writes = $true }
)

# 関数でない場所（窓のイベントの処理）の区切り: File・Category・Mark（ファイルにこの markTrace があること）
${traceFilePoints} = @(
    @{ File = "ui\shell\window_events.ps1"; Category = "startup"; Mark = "04 " }
    @{ File = "ui\shell\window_events.ps1"; Category = "startup"; Mark = "05 " }
    @{ File = "gui.ps1"; Category = "startup"; Mark = "01 " }
)

BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"

    function getTraceFunctionText {
        param ([string]$File, [string]$Function)

        $path = Join-Path "${scriptsDir}\tebunko" $File
        $tokens = $null
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
        $found = $ast.FindAll({ param ($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $Function }, $true)
        if ($found.Count -eq 0) { return $null }
        return $found[0].Extent.Text
    }
}

Describe "処理の記録の区切りの入れ所" -Tag Meta {
    It "<File> の <Function> に、<Category> の区切り <Marks> がある" -TestCases $tracePoints {
        $text = getTraceFunctionText $File $Function
        $text | Should -Not -BeNullOrEmpty -Because "$Function が $File に無い（関数を移した・名前を変えたときは、この表と perf.md の入れ所の表を直す）"
        foreach ($mark in $Marks) {
            $text | Should -Match ('markTrace\s+"' + $Category + '"\s+"[^"]*' + [regex]::Escape($mark)) -Because "区切り「$mark」"
        }
        if ($Writes) {
            $text | Should -Match ('writeTrace\s+"' + $Category + '"')
        }
    }

    It "<File> に、<Category> の区切り「<Mark>」がある" -TestCases $traceFilePoints {
        $text = [System.IO.File]::ReadAllText((Join-Path "${scriptsDir}\tebunko" $File), [System.Text.Encoding]::UTF8)
        $text | Should -Match ('markTrace\s+"' + $Category + '"\s+"' + [regex]::Escape($Mark))
    }

    It "区分ごとに writeTrace が 1 か所以上ある（書かれない記録を作らない）" -TestCases @(
        @{ Category = "startup" }, @{ Category = "search" }, @{ Category = "preview" }, @{ Category = "archive" }, @{ Category = "workspace" }, @{ Category = "index" }
    ) {
        $hits = @(Get-ChildItem -LiteralPath "${scriptsDir}\tebunko" -Recurse -Filter "*.ps1" | Select-String -Pattern ('writeTrace\s+"' + $Category + '"'))
        $hits.Count | Should -BeGreaterThan 0
    }

    It "markTrace・writeTrace に渡す区分は、決めた 6 つだけ" {
        $known = @("startup", "search", "preview", "archive", "workspace", "index")
        $used = @(Get-ChildItem -LiteralPath "${scriptsDir}\tebunko" -Recurse -Filter "*.ps1" | Select-String -Pattern '(?:markTrace|writeTrace)\s+"([^"]+)"' |
            ForEach-Object { $_.Matches[0].Groups[1].Value } | Sort-Object -Unique)
        @($used | Where-Object { $known -notcontains $_ }) | Should -BeNullOrEmpty
    }
}
