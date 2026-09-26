# 検索の要求と検索の司令のスレッド（tebunko\search\search_run.ps1 の newSearchRequest / invokeSearchRequest、
# tebunko\search\search_service.ps1 の SearchService）のテスト
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

Describe "newSearchRequest" -Tag Unit {
    It "検索条件を入れ、進み具合を始めの値にする" {
        $request = newSearchRequest "単価" $true @("x") 100 @{ CaseSensitive = $true; FileFilter = "*.xlsx"; IncludeShapes = $false } $true
        $request.Word | Should Be "単価"
        $request.Limit | Should Be 100
        $request.CaseSensitive | Should Be $true
        $request.FileFilter | Should Be "*.xlsx"
        $request.IncludeShapes | Should Be $false
        $request.IncludeComments | Should Be $true
        $request.UseFast | Should Be $true
        $request.Total | Should Be -1
        $request.Finished | Should Be $false
        $request.IsSynchronized | Should Be $true
    }
}

Describe "invokeSearchRequest" -Tag Io {
    $tsvRoot = Join-Path $TestDrive "request_tsv"
    newTsv "$tsvRoot\A社.xlsx\Sheet1.tsv" @("見積先：`t(株)山田商事", "単価`t105")
    newTsv "$tsvRoot\sub\文書.docx\ページ001.tsv" @("単価は別紙")
    $packRoot = Join-Path $TestDrive "request_pack"
    [void](newPackIndex $tsvRoot $packRoot)

    It "ヒットを列に入れ、件数と終わったことを伝える" {
        $request = newSearchRequest "単価" $true @($packRoot) 0
        invokeSearchRequest $request
        $request.Finished | Should Be $true
        $request.Error | Should BeNullOrEmpty
        (takeHits $request).Count | Should Be 2
        $request.Total | Should Be 2
        $request.Done | Should Be 2
        $request.FastUsed | Should Be $false
    }

    It "照合のプールを渡しても同じ結果になる" {
        $pool = newPackWorkerPool 2
        try {
            $request = newSearchRequest "単価" $true @($packRoot) 0
            invokeSearchRequest $request $pool
            (takeHits $request).Count | Should Be 2
        } finally {
            $pool.Close()
        }
    }

    It "始める前に取り消されていたら、検索せずに終える" {
        $request = newSearchRequest "単価" $true @($packRoot) 0
        $request.Stop = $true
        invokeSearchRequest $request
        $request.Cancelled | Should Be $true
        $request.Finished | Should Be $true
        $request.Total | Should Be -1
    }

    It "検索できなかった理由を Error に入れる（例外は投げない）" {
        Mock getIndexPackFiles { throw "読めません" }
        $request = newSearchRequest "単価" $true @($packRoot) 0
        invokeSearchRequest $request
        $request.Error | Should Be "読めません"
        $request.Finished | Should Be $true
    }
}

Describe "SearchService" -Tag Io {
    $tsvRoot = Join-Path $TestDrive "service_tsv"
    newTsv "$tsvRoot\A社.xlsx\Sheet1.tsv" @("単価`t105", "単価`t200")
    newTsv "$tsvRoot\B社.xlsx\Sheet1.tsv" @("単価`t300")
    $packRoot = Join-Path $TestDrive "service_pack"
    [void](newPackIndex $tsvRoot $packRoot)
    $libPath = "${scriptsDir}\tebunko\lib.ps1"

    It "要求を順に実行し、同じスレッドを使い続ける" {
        $service = newSearchService $libPath (newTsvTextCache) 2
        try {
            $first = $service.Request((newSearchRequest "単価" $true @($packRoot) 0))
            waitRequest $first | Should Be $true
            (takeHits $first).Count | Should Be 3
            $thread = $service.PowerShell
            $second = $service.Request((newSearchRequest "200" $true @($packRoot) 0))
            waitRequest $second | Should Be $true
            (takeHits $second).Count | Should Be 1
            [object]::ReferenceEquals($thread, $service.PowerShell) | Should Be $true
            $service.IsRunning() | Should Be $true
            $service.GetFailure() | Should Be ""
        } finally {
            $service.Close()
        }
    }

    It "次の要求を渡すと、前の要求を取り消す" {
        $service = newSearchService $libPath $null 1
        try {
            $first = newSearchRequest "単価" $true @($packRoot) 0
            [void]$service.Request($first)
            $second = $service.Request((newSearchRequest "単価" $true @($packRoot) 0))
            $first.Stop | Should Be $true
            waitRequest $second | Should Be $true
            (takeHits $second).Count | Should Be 3
        } finally {
            $service.Close()
        }
    }

    It "閉じるとスレッドを止める。Close は何度呼んでもよい" {
        $service = newSearchService $libPath $null 1
        $service.IsRunning() | Should Be $true
        $service.Close()
        $service.Close()
        $service.IsRunning() | Should Be $false
    }

    It "司令のスレッドが止まっていたら理由を返し、次の要求で作り直す" {
        $service = newSearchService (Join-Path $TestDrive "無い.ps1") $null 1
        try {
            $watch = [System.Diagnostics.Stopwatch]::StartNew()
            while ($service.IsRunning() -and $watch.Elapsed.TotalSeconds -lt 30) {
                Start-Sleep -Milliseconds 20
            }
            $service.IsRunning() | Should Be $false
            $service.GetFailure() | Should Not BeNullOrEmpty
            $service.LibPath = $libPath
            $request = $service.Request((newSearchRequest "単価" $true @($packRoot) 0))
            waitRequest $request | Should Be $true
            (takeHits $request).Count | Should Be 3
        } finally {
            $service.Close()
        }
    }

    It "スレッドの数を省くと、コア数から決める（getWorkerCount）" {
        $service = newSearchService $libPath
        try {
            $service.Workers | Should Be (getWorkerCount)
        } finally {
            $service.Close()
        }
        $pool = newPackWorkerPool
        try {
            $pool.Size | Should Be (getWorkerCount)
        } finally {
            $pool.Close()
        }
    }

    It "閉じるときに司令のスレッドが止まらなければ、スレッドを止めて片づける" {
        # 要求の列を見ずに動き続ける司令のスクリプト（照合が長引いている、など）
        $service = [SearchService]::new('param ($libPath, $requests, $cache, $workers) Start-Sleep -Seconds 60', $libPath, $null, 1, 200)
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        $service.Close()
        $watch.Elapsed.TotalSeconds | Should BeLessThan 30
        $service.IsRunning() | Should Be $false
    }

    It "司令のスレッドが止めずにエラーだけを書いて終わったら、その内容を理由として返す" {
        $service = [SearchService]::new('param ($libPath, $requests, $cache, $workers) Write-Error "司令のエラー"', $libPath, $null, 1, 5000)
        try {
            $watch = [System.Diagnostics.Stopwatch]::StartNew()
            while ($service.IsRunning() -and $watch.Elapsed.TotalSeconds -lt 30) {
                Start-Sleep -Milliseconds 20
            }
            $service.GetFailure() | Should Match "司令のエラー"
        } finally {
            $service.Close()
        }
    }
}

Describe "trimTsvTextCache" -Tag Unit {
    function addEntry($cache, [string]$path, [int]$chars, [long]$generation) {
        $cache.Texts[$path] = [object[]]@(0L, 0L, ("x" * $chars), @(), $generation)
        $cache.Chars[0] += $chars
    }

    It "上限の 9 割以下なら何も追い出さず、世代だけ進める" {
        $cache = newTsvTextCache 100
        addEntry $cache "a" 50 0
        trimTsvTextCache $cache | Should Be 0
        $cache.Texts.Count | Should Be 1
        $cache.Generation[0] | Should Be 1
    }

    It "超えていたら、今の世代で使わなかったものを古い世代から追い出す" {
        $cache = newTsvTextCache 100
        $cache.Generation[0] = 3
        addEntry $cache "old1" 30 1
        addEntry $cache "old2" 30 2
        addEntry $cache "now" 35 3
        trimTsvTextCache $cache | Should Be 1
        @($cache.Texts.Keys | Sort-Object) -join "," | Should Be "now,old2"
        $cache.Chars[0] | Should Be 65
        $cache.Generation[0] | Should Be 4
    }

    It "今の世代で使ったものは、上限を超えていても残す" {
        $cache = newTsvTextCache 100
        addEntry $cache "now1" 60 0
        addEntry $cache "now2" 40 0
        trimTsvTextCache $cache | Should Be 0
        $cache.Texts.Count | Should Be 2
    }
}

Describe "読んだ内容の世代（searchPackIndex）" -Tag Io {
    $tsvRoot = Join-Path $TestDrive "generation_tsv"
    newTsv "$tsvRoot\A社.xlsx\Sheet1.tsv" @("単価`t105")
    $packs = newPackIndex $tsvRoot (Join-Path $TestDrive "generation_pack")

    It "入れたとき・使ったときの世代を残す" {
        $cache = newTsvTextCache
        [void](searchPackIndex "単価" $packs $true -cache $cache)
        $cache.Texts[$packs[0].Path][4] | Should Be 0
        [void](trimTsvTextCache $cache)
        [void](searchPackIndex "単価" $packs $true -cache $cache)
        $cache.Texts[$packs[0].Path][4] | Should Be 1
    }
}

Describe "検索の司令のスクリプト（searchServiceScript）" -Tag Io {
    # 司令のスレッドで動くスクリプトを、このスレッドで直接動かして確かめる
    $tsvRoot = Join-Path $TestDrive "script_tsv"
    newTsv "$tsvRoot\A社.xlsx\Sheet1.tsv" @("単価`t105", "単価`t200")
    newTsv "$tsvRoot\B社.xlsx\Sheet1.tsv" @("単価`t300")
    $packRoot = Join-Path $TestDrive "script_pack"
    [void](newPackIndex $tsvRoot $packRoot)

    It "要求の列の要求を順に実行し、検索のたびに読んだ内容を整理する。列が閉じられたら終わる" {
        $requests = New-Object 'System.Collections.Concurrent.BlockingCollection[hashtable]'
        $first = newSearchRequest "単価" $true @($packRoot) 0
        $second = newSearchRequest "300" $true @($packRoot) 0
        $requests.Add($first)
        $requests.Add($second)
        $requests.CompleteAdding()
        $cache = newTsvTextCache

        & ${searchServiceScript} "${scriptsDir}\tebunko\lib.ps1" $requests $cache 2

        $first.Finished | Should Be $true
        (takeHits $first).Count | Should Be 3
        (takeHits $second).Count | Should Be 1
        # 検索 1 回ごとに世代を進める（trimTsvTextCache）
        $cache.Generation[0] | Should Be 2
    }
}

Describe "invokeSearchRequest（高速検索）" -Tag Io {
    $tsvRoot = Join-Path $TestDrive "fast_tsv"
    newTsv "$tsvRoot\A社.xlsx\Sheet1.tsv" @("単価`t105")
    $packRoot = Join-Path $TestDrive "fast_pack"
    $packs = newPackIndex $tsvRoot $packRoot

    It "Windows Search が使えれば、候補の集約ファイルだけを照合する" {
        Mock testWindowsSearch { $true }
        Mock getFastSearchPackFiles { @{ Folders = @(); Packs = $packs } }
        $request = newSearchRequest "単価" $true @($packRoot) 0 @{} $true
        invokeSearchRequest $request
        $request.FastAvailable | Should Be $true
        $request.FastUsed | Should Be $true
        (takeHits $request).Count | Should Be 1
        Assert-MockCalled getFastSearchPackFiles -Times 1 -Exactly -Scope It
    }
}
