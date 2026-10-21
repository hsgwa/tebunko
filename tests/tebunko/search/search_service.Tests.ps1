# 検索の要求と検索の司令のスレッド（tebunko\search\search_run.ps1 の newSearchRequest / invokeSearchRequest、
# tebunko\search\search_service.ps1 の SearchService）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    function waitRequest($request, [int]$seconds = 60) {
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        while (!$request.Finished -and $watch.Elapsed.TotalSeconds -lt $seconds) {
            Start-Sleep -Milliseconds 20
        }
        return $request.Finished
    }

    function takeHits($request) {
        $hits = New-Object System.Collections.Generic.List[object]
        $hit = $null
        while ($request.Queue.TryDequeue([ref]$hit)) {
            $hits.Add($hit)
        }
        return , $hits
    }
}

Describe "newSearchRequest" -Tag Unit {
    It "検索条件を入れ、進み具合を始めの値にする" {
        $request = newSearchRequest "単価" $true @("x") 100 @{ CaseSensitive = $true; FileKinds = @("excel"); IncludeShapes = $false } "C:\ws"
        $request.Word | Should -Be "単価"
        $request.Limit | Should -Be 100
        $request.CaseSensitive | Should -Be $true
        $request.FileFilter | Should -Be (newFileKindFilter @("excel"))
        $request.FileFilter | Should -Not -Be ""
        $request.IncludeShapes | Should -Be $false
        $request.IncludeComments | Should -Be $true
        $request.WorkDir | Should -Be "C:\ws"
        $request.GramUsed | Should -Be $false
        $request.Total | Should -Be -1
        $request.Finished | Should -Be $false
        $request.IsSynchronized | Should -Be $true
    }
}

Describe "invokeSearchRequest" -Tag Io {
    BeforeAll {
        $tsvRoot = Join-Path $TestDrive "request_tsv"
        newTsv "$tsvRoot\A社.xlsx\Sheet1.tsv" @("見積先：`t(株)山田商事", "単価`t105")
        newTsv "$tsvRoot\sub\文書.docx\ページ001.tsv" @("単価は別紙")
        $contentIndexRoot = Join-Path $TestDrive "request_content_index"
        [void](newContentIndexFiles $tsvRoot $contentIndexRoot)
    }

    It "ヒットを列に入れ、件数と終わったことを伝える" {
        $request = newSearchRequest "単価" $true @($contentIndexRoot) 0
        invokeSearchRequest $request
        $request.Finished | Should -Be $true
        $request.Error | Should -BeNullOrEmpty
        (takeHits $request).Count | Should -Be 2
        $request.Total | Should -Be 2
        $request.Done | Should -Be 2
    }

    It "壊れた本文インデックスのファイルが混ざっても、Error にせず、ほかのヒットを返し、飛ばしたファイルを Skipped に入れる" {
        $brokenRoot = Join-Path $TestDrive "request_broken_content_index"
        [void](newContentIndexFiles $tsvRoot $brokenRoot)
        [void][System.IO.Directory]::CreateDirectory("$brokenRoot\壊れ")
        writeContentIndexFile "$brokenRoot\壊れ\content_index.xlsx.001.tsv" "$([char]0x1E) 版=999`n単価`n"
        $request = newSearchRequest "単価" $true @($brokenRoot) 0
        invokeSearchRequest $request
        $request.Error | Should -BeNullOrEmpty
        (takeHits $request).Count | Should -Be 2
        @($request.Skipped).Count | Should -Be 1
        $request.Skipped[0] | Should -BeLike "*\壊れ\content_index.xlsx.001.tsv"
    }

    It "照合のプールを渡しても同じ結果になる" {
        $pool = newContentIndexWorkerPool 2
        try {
            $request = newSearchRequest "単価" $true @($contentIndexRoot) 0
            invokeSearchRequest $request $pool
            (takeHits $request).Count | Should -Be 2
        } finally {
            $pool.Close()
        }
    }

    It "Excel のヘッダー・フッターの行は、プールの別スレッドでも、図形・コメントを除いた検索でも場所つきでヒットする" {
        $hfTsvRoot = Join-Path $TestDrive "pool_hf_tsv"
        newTsv "$hfTsvRoot\B社.xlsx\$(toIndexFileName "4月[ヘッダー・フッター]")" @("社外秘")
        newTsv "$hfTsvRoot\B社.xlsx\$(toIndexFileName "4月[図形]")" @("A1`t社外秘の図")
        $hfContentIndexRoot = Join-Path $TestDrive "pool_hf_content_index"
        [void](newContentIndexFiles $hfTsvRoot $hfContentIndexRoot)

        $pool = newContentIndexWorkerPool 2
        try {
            $request = newSearchRequest "社外秘" $true @($hfContentIndexRoot) 0 @{ IncludeShapes = $false; IncludeComments = $false }
            invokeSearchRequest $request $pool
            $hits = @(takeHits $request)
            $hits.Count | Should -Be 1
            $hits[0].Location | Should -Be "4月[ヘッダー・フッター]"
        } finally {
            $pool.Close()
        }
    }

    It "照合のプールの別スレッドでも、テキストの長い行を切る（newContentIndexWorkerPool に渡す関数・変数が揃っている）" {
        $poolTsvRoot = Join-Path $TestDrive "pool_text_tsv"
        $longLine = ("あ" * 100000) + "対象語" + ("あ" * 100000)
        newTsv "$poolTsvRoot\長い.log\$(toIndexFileName "本文")" @($longLine)
        $poolContentIndexFiles = newContentIndexFiles $poolTsvRoot (Join-Path $TestDrive "pool_text_content_index")

        $pool = newContentIndexWorkerPool 2
        try {
            # searchContentIndex のタスク分けを経由せず、プールのスレッドに直接照合させる（content_index_search.ps1 と同じ呼び方）
            $search = newSearchRegex "対象語" $true
            $job = $pool.Submit(${contentIndexWorkerScript}.ToString(), @($poolContentIndexFiles, 0, $poolContentIndexFiles.Count, $search.Regex, -1, $search.TextRegex, $search.ScanMode, $null, $null, $null, $null))
            $output = $pool.Receive($job)
            $output[0].Timeout | Should -Be $false
            $hits = @($output[0].Hits)
            $hits.Count | Should -Be 1
            $hits[0].Line.Length | Should -BeLessOrEqual 1002
            $hits[0].Line | Should -Match "対象語"
        } finally {
            $pool.Close()
        }
    }

    It "始める前に取り消されていたら、検索せずに終える" {
        $request = newSearchRequest "単価" $true @($contentIndexRoot) 0
        $request.Stop = $true
        invokeSearchRequest $request
        $request.Cancelled | Should -Be $true
        $request.Finished | Should -Be $true
        $request.Total | Should -Be -1
    }

    It "検索できなかった理由を Error に入れる（例外は投げない）" {
        Mock getContentIndexFiles { throw "読めません" }
        $request = newSearchRequest "単価" $true @($contentIndexRoot) 0
        invokeSearchRequest $request
        $request.Error | Should -Be "読めません"
        $request.Finished | Should -Be $true
    }
}

Describe "invokeSearchRequest（どの段で中止されても、先の段に進まない）" -Tag Io {
    BeforeDiscovery {
        # invokeSearchRequest が順に呼ぶ段。段を足したら（呼ぶ関数を足したら）ここに行を足し、下の Mock にも足す
        $stages = @(
            @{ Stage = "getGramContentIndexFiles"; Later = @("getContentIndexFiles", "searchContentIndex") }
            @{ Stage = "getContentIndexFiles"; Later = @("searchContentIndex") }
            @{ Stage = "searchContentIndex"; Later = @() }
        )
    }

    BeforeAll {
        $tsvRoot = Join-Path $TestDrive "stage_tsv"
        newTsv "$tsvRoot\A社.xlsx\Sheet1.tsv" @("単価`t105")
        $contentIndexRoot = Join-Path $TestDrive "stage_content_index"
        [void](newContentIndexFiles $tsvRoot $contentIndexRoot)
    }

    It "<Stage> の中で Stop が立つと、Cancelled にして終え、後の段（<Later>）を呼ばず、ヒットを列に入れない" -ForEach $stages {
        $script:stopStage = $Stage
        $script:stopRequest = newSearchRequest "単価" $true @($contentIndexRoot) 0
        Mock getGramContentIndexFiles {
            if ($script:stopStage -eq "getGramContentIndexFiles") { $script:stopRequest.Stop = $true }
            return $null
        }
        Mock getContentIndexFiles {
            if ($script:stopStage -eq "getContentIndexFiles") { $script:stopRequest.Stop = $true }
            return @{ Folders = @(); ContentIndexFiles = @() }
        }
        Mock searchContentIndex {
            if ($script:stopStage -eq "searchContentIndex") { $script:stopRequest.Stop = $true }
            return @{ Hits = @(); Total = 0; Truncated = $false; Cancelled = $true; Skipped = @() }
        }

        invokeSearchRequest $script:stopRequest

        $script:stopRequest.Cancelled | Should -Be $true
        $script:stopRequest.Finished | Should -Be $true
        $script:stopRequest.Error | Should -BeNullOrEmpty
        $script:stopRequest.Queue.IsEmpty | Should -Be $true
        Should -Invoke -CommandName $Stage -Times 1 -Exactly -Scope It
        foreach ($name in $Later) {
            Should -Invoke -CommandName $name -Times 0 -Exactly -Scope It
        }
        # 止める合図は、要求そのものを渡す（作業スレッドにも渡せる）
        Should -Invoke -CommandName $Stage -Times 1 -Exactly -Scope It -ParameterFilter { [object]::ReferenceEquals($stopSignal, $script:stopRequest) }
    }

    It "invokeSearchRequest が呼ぶ検索・インデックスの関数は、すべて段の表にあるか、止める必要の無い短い呼び出しとして許されている（段を足したとき表に足し忘れない）" -ForEach @(@{ Stages = $stages }) {
        # 止める必要の無い短い呼び出しは、ここに名前と理由を書く
        $allowed = @{
            getSearchOverflowRoot = "溢れの置き場所の名前を組み立てるだけ"
            getSearchOverflowPidDir = "溢れの置き場所の名前を組み立てるだけ"
            getSearchOverflowDir = "溢れの置き場所の名前を組み立てるだけ"
            testSearchOverflowRoot = "検索の始めに 1 回、置き場所の根を確かめるだけ（検索の中身の段ではない）"
            removeSearchOverflow = "検索の始めに 1 回、前の溢れを消すだけ（検索の中身の段ではない）"
        }
        $parse = { param ($path) [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$null, [ref]$null) }
        $defined = New-Object System.Collections.Generic.HashSet[string]
        foreach ($dir in @("search", "index")) {
            foreach ($file in (Get-ChildItem -LiteralPath "${scriptsDir}\tebunko\$dir" -Filter "*.ps1")) {
                foreach ($function in (& $parse $file.FullName).FindAll({ param ($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
                    [void]$defined.Add($function.Name)
                }
            }
        }
        $body = (Get-Command invokeSearchRequest).ScriptBlock.Ast
        $called = @($body.FindAll({ param ($node) $node -is [System.Management.Automation.Language.CommandAst] }, $true) |
                ForEach-Object { $_.GetCommandName() } | Where-Object { $_ -and $defined.Contains($_) } | Sort-Object -Unique)
        $called.Count | Should -BeGreaterThan 0
        $stageNames = @($Stages | ForEach-Object { $_.Stage })
        $unexpected = @($called | Where-Object { $stageNames -notcontains $_ -and !$allowed.ContainsKey($_) })
        $unexpected | Should -Be @()
        # 表の段は、実際に呼ばれている
        @($stageNames | Where-Object { $called -notcontains $_ }) | Should -Be @()
    }
}

Describe "SearchService" -Tag Io {
    BeforeAll {
        $tsvRoot = Join-Path $TestDrive "service_tsv"
        newTsv "$tsvRoot\A社.xlsx\Sheet1.tsv" @("単価`t105", "単価`t200")
        newTsv "$tsvRoot\B社.xlsx\Sheet1.tsv" @("単価`t300")
        $contentIndexRoot = Join-Path $TestDrive "service_content_index"
        [void](newContentIndexFiles $tsvRoot $contentIndexRoot)
        $libLoad = getPartLoad lib
    }

    It "要求を順に実行し、同じスレッドを使い続ける" {
        $service = newSearchService (newTsvTextCache) 2
        try {
            $first = $service.Request((newSearchRequest "単価" $true @($contentIndexRoot) 0))
            waitRequest $first | Should -Be $true
            (takeHits $first).Count | Should -Be 3
            $thread = $service.PowerShell
            $second = $service.Request((newSearchRequest "200" $true @($contentIndexRoot) 0))
            waitRequest $second | Should -Be $true
            (takeHits $second).Count | Should -Be 1
            [object]::ReferenceEquals($thread, $service.PowerShell) | Should -Be $true
            $service.IsRunning() | Should -Be $true
            $service.GetFailure() | Should -Be ""
        } finally {
            $service.Close()
        }
    }

    It "次の要求を渡すと、前の要求を取り消す" {
        $service = newSearchService $null 1
        try {
            $first = newSearchRequest "単価" $true @($contentIndexRoot) 0
            [void]$service.Request($first)
            $second = $service.Request((newSearchRequest "単価" $true @($contentIndexRoot) 0))
            $first.Stop | Should -Be $true
            waitRequest $second | Should -Be $true
            (takeHits $second).Count | Should -Be 3
        } finally {
            $service.Close()
        }
    }

    It "閉じるとスレッドを止める。Close は何度呼んでもよい" {
        $service = newSearchService $null 1
        $service.IsRunning() | Should -Be $true
        $service.Close()
        $service.Close()
        $service.IsRunning() | Should -Be $false
    }

    It "Deferred で作ると始めず、Prepare で始める（何度呼んでも同じスレッド）" {
        $service = newSearchService $null 1 -Deferred
        try {
            $service.IsRunning() | Should -Be $false
            $service.Prepare()
            $service.IsRunning() | Should -Be $true
            $thread = $service.PowerShell
            $service.Prepare()
            [object]::ReferenceEquals($thread, $service.PowerShell) | Should -Be $true
        } finally {
            $service.Close()
        }
    }

    It "Deferred で作ったまま最初に検索しても（画面が出た直後でも）、検索の要求で始まって結果が返る" {
        $service = newSearchService (newTsvTextCache) 2 -Deferred
        try {
            $service.IsRunning() | Should -Be $false
            $request = $service.Request((newSearchRequest "単価" $true @($contentIndexRoot) 0))
            waitRequest $request | Should -Be $true
            (takeHits $request).Count | Should -Be 3
            $service.IsRunning() | Should -Be $true
        } finally {
            $service.Close()
        }
    }

    It "Deferred で作って始めずに閉じても、例外にならない" {
        $service = newSearchService $null 1 -Deferred
        { $service.Close() } | Should -Not -Throw
        { $service.Close() } | Should -Not -Throw
        $service.IsRunning() | Should -Be $false
    }

    It "ワークスペースを切り替えて検索すると、前のワークスペースの溢れを次の検索の始めに消し、閉じると今のものも消す" {
        $wsA = Join-Path $TestDrive "ws_a"
        $wsB = Join-Path $TestDrive "ws_b"
        $pidA = getSearchOverflowPidDir ([Workspace]::new($wsA).TmpRoot)
        $pidB = getSearchOverflowPidDir ([Workspace]::new($wsB).TmpRoot)
        $service = newSearchService $null 1
        try {
            $first = $service.Request((newSearchRequest "単価" $true @($contentIndexRoot) 1 @{} $wsA))
            waitRequest $first | Should -Be $true
            [System.IO.Directory]::Exists($pidA) | Should -Be $true
            $second = $service.Request((newSearchRequest "単価" $true @($contentIndexRoot) 1 @{} $wsB))
            waitRequest $second | Should -Be $true
            [System.IO.Directory]::Exists($pidA) | Should -Be $false
            [System.IO.Directory]::Exists($pidB) | Should -Be $true
        } finally {
            $service.Close()
        }
        [System.IO.Directory]::Exists($pidB) | Should -Be $false
    }

    It "司令のスレッドが止まっていたら理由を返し、次の要求で作り直す（zip 版。読み込む先が無い）" {
        $brokenLibLoad = @{
            State   = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
            Prelude = ". '$(Join-Path $TestDrive "無い.ps1")'"
        }
        $service = [SearchService]::new(${searchServiceScript}.ToString(), $brokenLibLoad, $null, 1, ${searchServiceCloseMilliseconds})
        try {
            $watch = [System.Diagnostics.Stopwatch]::StartNew()
            while ($service.IsRunning() -and $watch.Elapsed.TotalSeconds -lt 30) {
                Start-Sleep -Milliseconds 20
            }
            $service.IsRunning() | Should -Be $false
            $service.GetFailure() | Should -Not -BeNullOrEmpty
            $service.LibLoad = $libLoad
            $request = $service.Request((newSearchRequest "単価" $true @($contentIndexRoot) 0))
            waitRequest $request | Should -Be $true
            (takeHits $request).Count | Should -Be 3
        } finally {
            $service.Close()
        }
    }

    It "司令のスレッドが止まっていたら理由を返す（単一 .ps1 版。部品を読み込む関数が失敗する）" {
        $brokenState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
        $brokenEntry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry("importTebunkoPart", 'throw "部品を読み込めません"')
        $brokenState.Commands.Add($brokenEntry)
        $brokenLibLoad = @{ State = $brokenState; Prelude = ". importTebunkoPart" }
        $service = [SearchService]::new(${searchServiceScript}.ToString(), $brokenLibLoad, $null, 1, ${searchServiceCloseMilliseconds})
        try {
            $watch = [System.Diagnostics.Stopwatch]::StartNew()
            while ($service.IsRunning() -and $watch.Elapsed.TotalSeconds -lt 30) {
                Start-Sleep -Milliseconds 20
            }
            $service.IsRunning() | Should -Be $false
            $service.GetFailure() | Should -Not -BeNullOrEmpty
        } finally {
            $service.Close()
        }
    }

    It "スレッドの数を省くと、コア数から決める（getWorkerCount）" {
        $service = newSearchService
        try {
            $service.Workers | Should -Be (getWorkerCount)
        } finally {
            $service.Close()
        }
        $pool = newContentIndexWorkerPool
        try {
            $pool.Size | Should -Be (getWorkerCount)
        } finally {
            $pool.Close()
        }
    }

    It "閉じるときに司令のスレッドが止まらなければ、スレッドを止めて片づける" {
        # 要求の列を見ずに動き続ける司令のスクリプト（照合が長引いている、など）
        $service = [SearchService]::new('param ($requests, $cache, $workers) Start-Sleep -Seconds 60', $libLoad, $null, 1, 200)
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        $service.Close()
        $watch.Elapsed.TotalSeconds | Should -BeLessThan 30
        $service.IsRunning() | Should -Be $false
    }

    It "司令のスレッドが止めずにエラーだけを書いて終わったら、その内容を理由として返す" {
        $service = [SearchService]::new('param ($requests, $cache, $workers) Write-Error "司令のエラー"', $libLoad, $null, 1, 5000)
        try {
            $watch = [System.Diagnostics.Stopwatch]::StartNew()
            while ($service.IsRunning() -and $watch.Elapsed.TotalSeconds -lt 30) {
                Start-Sleep -Milliseconds 20
            }
            $service.GetFailure() | Should -Match "司令のエラー"
        } finally {
            $service.Close()
        }
    }
}

Describe "trimTsvTextCache" -Tag Unit {
    BeforeAll {
        function addEntry($cache, [string]$path, [int]$chars, [long]$generation) {
            $cache.Texts[$path] = [object[]]@(0L, 0L, ("x" * $chars), @(), $generation)
            $cache.Chars[0] += $chars
        }
    }

    It "上限の 9 割以下なら何も追い出さず、世代だけ進める" {
        $cache = newTsvTextCache 100
        addEntry $cache "a" 50 0
        trimTsvTextCache $cache | Should -Be 0
        $cache.Texts.Count | Should -Be 1
        $cache.Generation[0] | Should -Be 1
    }

    It "超えていたら、今の世代で使わなかったものを古い世代から追い出す" {
        $cache = newTsvTextCache 100
        $cache.Generation[0] = 3
        addEntry $cache "old1" 30 1
        addEntry $cache "old2" 30 2
        addEntry $cache "now" 35 3
        trimTsvTextCache $cache | Should -Be 1
        @($cache.Texts.Keys | Sort-Object) -join "," | Should -Be "now,old2"
        $cache.Chars[0] | Should -Be 65
        $cache.Generation[0] | Should -Be 4
    }

    It "今の世代で使ったものは、上限を超えていても残す" {
        $cache = newTsvTextCache 100
        addEntry $cache "now1" 60 0
        addEntry $cache "now2" 40 0
        trimTsvTextCache $cache | Should -Be 0
        $cache.Texts.Count | Should -Be 2
    }
}

Describe "読んだ内容の世代（searchContentIndex）" -Tag Io {
    BeforeAll {
        $tsvRoot = Join-Path $TestDrive "generation_tsv"
        newTsv "$tsvRoot\A社.xlsx\Sheet1.tsv" @("単価`t105")
        $contentIndexFiles = newContentIndexFiles $tsvRoot (Join-Path $TestDrive "generation_content_index")
    }

    It "入れたとき・使ったときの世代を残す" {
        $cache = newTsvTextCache
        [void](searchContentIndex "単価" $contentIndexFiles $true -cache $cache)
        $cache.Texts[$contentIndexFiles[0].Path][4] | Should -Be 0
        [void](trimTsvTextCache $cache)
        [void](searchContentIndex "単価" $contentIndexFiles $true -cache $cache)
        $cache.Texts[$contentIndexFiles[0].Path][4] | Should -Be 1
    }
}

Describe "検索の司令のスクリプト（searchServiceScript）" -Tag Io {
    BeforeAll {
        # 司令のスレッドで動くスクリプトを、このスレッドで直接動かして確かめる
        $tsvRoot = Join-Path $TestDrive "script_tsv"
        newTsv "$tsvRoot\A社.xlsx\Sheet1.tsv" @("単価`t105", "単価`t200")
        newTsv "$tsvRoot\B社.xlsx\Sheet1.tsv" @("単価`t300")
        $contentIndexRoot = Join-Path $TestDrive "script_content_index"
        [void](newContentIndexFiles $tsvRoot $contentIndexRoot)
    }

    It "要求の列の要求を順に実行し、検索のたびに読んだ内容を整理する。列が閉じられたら終わる" {
        $requests = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
        $first = newSearchRequest "単価" $true @($contentIndexRoot) 0
        $second = newSearchRequest "300" $true @($contentIndexRoot) 0
        $requests.Add($first)
        $requests.Add($second)
        $requests.CompleteAdding()
        $cache = newTsvTextCache

        & ${searchServiceScript} $requests $cache 2

        $first.Finished | Should -Be $true
        (takeHits $first).Count | Should -Be 3
        (takeHits $second).Count | Should -Be 1
        # 検索 1 回ごとに世代を進める（trimTsvTextCache）
        $cache.Generation[0] | Should -Be 2
    }
}

Describe "invokeSearchRequest（gram 索引）" -Tag Io {
    BeforeAll {
        $tsvRoot = Join-Path $TestDrive "gram_tsv"
        newTsv "$tsvRoot\A社.xlsx\Sheet1.tsv" @("単価`t105")
        newTsv "$tsvRoot\B社.xlsx\Sheet1.tsv" @("別の本文")
        $contentIndexRoot = Join-Path $TestDrive "gram_content_index"
        $contentIndexFiles = newContentIndexFiles $tsvRoot $contentIndexRoot
    }

    It "gram 索引で絞れたら GramUsed が立ち、絞った分だけを照合する" {
        $narrowed = @{ Folders = @(); ContentIndexFiles = @($contentIndexFiles | Select-Object -First 1) }
        Mock getGramContentIndexFiles { $narrowed }
        $request = newSearchRequest "単価" $true @($contentIndexRoot) 0 @{}
        invokeSearchRequest $request
        $request.Error | Should -BeNullOrEmpty
        $request.GramUsed | Should -Be $true
        $request.Total | Should -Be 1
    }

    It "gram 索引で絞れなかったら、すべてを照合する（GramUsed は立たない）" {
        Mock getGramContentIndexFiles { $null }
        $request = newSearchRequest "単価" $true @($contentIndexRoot) 0
        invokeSearchRequest $request
        $request.GramUsed | Should -Be $false
        $request.Total | Should -Be @($contentIndexFiles).Count
        (takeHits $request).Count | Should -Be 1
    }

    It "gram 索引で 0 件まで絞ったとき、GramUsed が立ち、集めた数は 0（画面は「インデックスが未作成」にしない）" {
        Mock getGramContentIndexFiles { @{ Folders = @(); ContentIndexFiles = @() } }
        $request = newSearchRequest "単価" $true @($contentIndexRoot) 0 @{}
        invokeSearchRequest $request
        $request.Error | Should -BeNullOrEmpty
        $request.GramUsed | Should -Be $true
        $request.IndexTotal | Should -Be 0
        $request.Total | Should -Be 0
        (takeHits $request).Count | Should -Be 0
    }

    It "gram 索引で絞った候補が、ファイルの種類の絞り込みで全部外れたとき、GramUsed が立つ（ヒットは 0 件）" {
        $narrowed = @{ Folders = @(); ContentIndexFiles = @($contentIndexFiles | Select-Object -First 1) }
        Mock getGramContentIndexFiles { $narrowed }
        $request = newSearchRequest "単価" $true @($contentIndexRoot) 0 @{ FileKinds = @("word") }
        invokeSearchRequest $request
        $request.GramUsed | Should -Be $true
        $request.IndexTotal | Should -Be 1
        (takeHits $request).Count | Should -Be 0
    }
}
