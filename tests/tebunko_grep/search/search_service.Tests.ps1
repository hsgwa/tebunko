# 検索の要求と検索の司令のスレッド（tebunko_grep\search\search_run.ps1 の newSearchRequest / invokeSearchRequest、
# tebunko_grep\search\search_service.ps1 の SearchService）のテスト
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
    $libPath = "${scriptsDir}\tebunko_grep\lib.ps1"

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
}
