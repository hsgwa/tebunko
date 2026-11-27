# 検索結果の取り出し（tebunko\ui\search\search_session.ps1 の pumpSearch・takeSearchHits）を、DispatcherTimer に頼らずに
# 画面のディスパッチャーへ積んだ仕事（search_run.ps1 の requestSearchPump）だけで回せることのテスト。
# 関数だけをファイルから取り出し、画面の部品は呼ばれた操作を残す偽物にする（ファイルは画面を作る処理を含むため、そのままは読み込まない）。
# 終わりの 2 つの Describe は、本物のディスパッチャーで、WPF が優先度 Background の仕事を、タイマー（WM_TIMER）無しに回すことを確かめる。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    . "${scriptsDir}\shared\ui\types.ps1"
    . "${scriptsDir}\tebunko\ui\search\result_types.ps1"

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile("${scriptsDir}\tebunko\ui\search\search_session.ps1", [ref]$tokens, [ref]$errors)
    foreach ($name in @("pumpSearch", "takeSearchHits")) {
        $definition = $ast.Find({ param ($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
        if ($null -ne $definition) {
            . ([scriptblock]::Create($definition.Extent.Text))
        }
    }

    # 画面の部品の偽物（呼ばれた操作を $global:pumpLog に残す）
    function flushResults { $global:pumpLog.Add("flushResults"); if ($global:onFlush) { & $global:onFlush } }
    function showSearchProgress { param ($s) $global:pumpLog.Add("showSearchProgress") }
    function finishSearch { $global:pumpLog.Add("finishSearch"); $script:search = $null }
    function addFileGroupLocation { param ($group, $book, $location) }
    function newFileGroup {
        param ([string]$key, [string]$relDir, [string]$book)
        $group = [FileGroup]::new()
        $group.Book = $book
        $group.RelDir = $relDir
        $group.FullPath = $key
        $script:fileGroups[$key] = $group
        return $group
    }

    # 画面のディスパッチャーの代わり。積まれた仕事（引数の要求）を控え、テストが順に呼ぶ
    function newFakeWake {
        $wake = [pscustomobject]@{ Posted = New-Object System.Collections.Generic.Queue[object] }
        $wake | Add-Member -MemberType ScriptMethod -Name Post -Value {
            param ($callback, $state)
            $this.Posted.Enqueue($state)
        }
        return $wake
    }

    function newPumpRequest {
        param ($wake, [int]$count, [bool]$finished)
        $request = newSearchRequest "単価" $true @("x") 0
        $request.Wake = $wake
        $request.WakeCallback = [System.Threading.SendOrPostCallback] { param ($state) }
        addPumpHits $request $count
        $request.Total = 1
        $request.Done = 1
        $request.Finished = $finished
        return $request
    }

    function addPumpHits {
        param ($request, [int]$count, [int]$start = 1)
        for ($i = $start; $i -lt $start + $count; $i++) {
            $request.Queue.Enqueue([pscustomobject]@{ Root = "C:\共有"; RelPath = "営業部\content_index.xlsx.001.tsv"; RelDir = "営業部"; FileName = "見積.xlsx"
                Book = "見積.xlsx"; Location = "Sheet1"; LineNumber = $i; Line = "単価`t$i" })
        }
    }

    function startPumpSearch {
        param ($request)
        $script:search = @{ Shared = $request; Limit = 0; TraceFirstRow = $true; PumpTimer = 0; PumpPosted = 0 }
    }

    # 積まれた仕事を、無くなるまで（多くても limit 回）順に呼ぶ。呼んだ回数を返す
    function invokePosted {
        param ($wake, [int]$limit = 100000)
        $calls = 0
        while ($wake.Posted.Count -gt 0 -and $calls -lt $limit) {
            pumpSearch $wake.Posted.Dequeue()
            $calls++
        }
        return $calls
    }
}

Describe "pumpSearch（積まれた取り出し）" -Tag Unit {
    BeforeEach {
        $global:pumpLog = New-Object System.Collections.Generic.List[string]
        $global:onFlush = $null
        $script:fileGroups = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
        $script:dirtyGroups = New-Object 'System.Collections.Generic.HashSet[object]'
        $script:hitCount = 0
        $script:searchPumpClosed = $false
        ${searchPumpMilliseconds} = 1
        # Tick を一度も呼ばない偽のタイマー（Start・Stop は数えるだけ）
        $script:searchTimer = [pscustomobject]@{ Stops = 0 }
        $script:searchTimer | Add-Member -MemberType ScriptMethod -Name Stop -Value { $this.Stops++ }
        $script:searchTimer | Add-Member -MemberType ScriptMethod -Name Start -Value { }
        $script:searchService = [pscustomobject]@{}
        $script:searchService | Add-Member -MemberType ScriptMethod -Name IsRunning -Value { $true }
        $script:searchService | Add-Member -MemberType ScriptMethod -Name GetFailure -Value { "" }
    }
    AfterAll {
        Remove-Variable -Name pumpLog, onFlush -Scope Global -ErrorAction SilentlyContinue
    }

    It "タイマーが一度も動かなくても、積まれた取り出しだけで最後まで取り出して終える" {
        $wake = newFakeWake
        $request = newPumpRequest $wake 3000 $true
        startPumpSearch $request
        requestSearchPump $request
        $calls = invokePosted $wake
        $calls | Should -BeGreaterThan 1
        $request.Queue.IsEmpty | Should -Be $true
        $script:hitCount | Should -Be 3000
        @($global:pumpLog | Where-Object { $_ -eq "finishSearch" }).Count | Should -Be 1
        $wake.Posted.Count | Should -Be 0
        $script:search | Should -BeNullOrEmpty
    }

    It "予算を使い切って残りがあれば、自分をもう一度積む" {
        $wake = newFakeWake
        $request = newPumpRequest $wake 3000 $true
        startPumpSearch $request
        pumpSearch $request
        $request.Queue.IsEmpty | Should -Be $false
        $wake.Posted.Count | Should -Be 1
        [object]::ReferenceEquals($wake.Posted.Peek(), $request) | Should -Be $true
        $request.PumpQueued | Should -Be $true
        @($global:pumpLog | Where-Object { $_ -eq "finishSearch" }).Count | Should -Be 0
    }

    It "取り出しの入口で印を偽にしてから取り出す（積み漏れが無い）" {
        ${searchPumpMilliseconds} = 60000
        $wake = newFakeWake
        $request = newPumpRequest $wake 5 $false
        startPumpSearch $request
        # 司令が積んだ仕事を待っている（印が真）ところから始める
        requestSearchPump $request
        [void]$wake.Posted.Dequeue()
        # 取り出しの途中で、司令が 1 件足して積もうとする
        $global:onFlush = {
            $global:onFlush = $null
            addPumpHits $request 1 6
            requestSearchPump $request
        }
        pumpSearch $request
        $script:hitCount | Should -Be 5
        $wake.Posted.Count | Should -Be 1
        [void](invokePosted $wake)
        $script:hitCount | Should -Be 6
        $request.Queue.IsEmpty | Should -Be $true
    }

    It "前の検索の取り出しは何もしない" {
        $wake = newFakeWake
        $first = newPumpRequest $wake 10 $true
        startPumpSearch $first
        requestSearchPump $first
        $second = newPumpRequest $wake 10 $true
        startPumpSearch $second
        [void](invokePosted $wake 1)
        $second.Queue.Count | Should -Be 10
        $script:hitCount | Should -Be 0
        @($global:pumpLog).Count | Should -Be 0
        [object]::ReferenceEquals($script:search.Shared, $second) | Should -Be $true
    }

    It "窓を閉じたあとに積まれていた取り出しは何もしない" {
        $wake = newFakeWake
        $request = newPumpRequest $wake 10 $true
        startPumpSearch $request
        requestSearchPump $request
        $script:searchPumpClosed = $true
        [void](invokePosted $wake)
        $request.Queue.Count | Should -Be 10
        @($global:pumpLog).Count | Should -Be 0
        $script:searchTimer.Stops | Should -Be 0
        $wake.Posted.Count | Should -Be 0
    }

    It "タイマーからの取り出しと積まれた取り出しを、検索ごとに数える" {
        ${searchPumpMilliseconds} = 60000
        $wake = newFakeWake
        $request = newPumpRequest $wake 5 $false
        startPumpSearch $request
        $s = $script:search
        pumpSearch
        requestSearchPump $request
        [void](invokePosted $wake)
        $s.PumpTimer | Should -Be 1
        $s.PumpPosted | Should -Be 1
    }
}

Describe "本物のディスパッチャー（優先度 Background の積み直し）" -Tag Unit {
    BeforeAll {
        # 5 秒で抜けなければ、別のスレッドから DispatcherFrame を抜けさせる（DispatcherTimer は使わない）
        function startFrameGuard {
            param ($frame, [int]$seconds)
            $guard = [powershell]::Create()
            [void]$guard.AddScript({
                    param ($frame, $seconds)
                    Start-Sleep -Seconds $seconds
                    $frame.Continue = $false
                }).AddArgument($frame).AddArgument($seconds)
            return @{ PowerShell = $guard; Handle = $guard.BeginInvoke() }
        }
    }

    It "本物のディスパッチャーで、DispatcherTimer 無しに Background の取り出しが回る" {
        $dispatcher = [System.Windows.Threading.Dispatcher]::CurrentDispatcher
        $shared = @{ Count = 0; Frame = (New-Object System.Windows.Threading.DispatcherFrame) }
        $shared.Context = New-Object System.Windows.Threading.DispatcherSynchronizationContext($dispatcher, [System.Windows.Threading.DispatcherPriority]::Background)
        $shared.Callback = [System.Threading.SendOrPostCallback]({
                param ($state)
                $shared.Count++
                if ($shared.Count -ge 3) { $shared.Frame.Continue = $false } else { $shared.Context.Post($shared.Callback, $null) }
            }.GetNewClosure())
        $shared.Context.Post($shared.Callback, $null)
        $guard = startFrameGuard $shared.Frame 5
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            [System.Windows.Threading.Dispatcher]::PushFrame($shared.Frame)
        } finally {
            $guard.PowerShell.Stop()
            $guard.PowerShell.Dispose()
        }
        $shared.Count | Should -Be 3
        $watch.Elapsed.TotalSeconds | Should -BeLessThan 5
    }
}

Describe "本物のディスパッチャー（別のスレッドから Send の仕事を積み続ける中の積み直し）" -Tag Slow {
    BeforeAll {
        # 別のスレッド（ランスペース）が、優先度 Send の仕事を時間切れ 1 秒の同期の呼び出しで途切れずに積む（止まった回の UI オートメーションの形をまねる）。
        # 積む仕事は .NET のメソッドから作った delegate で、別のランスペースのスクリプトブロックを画面のスレッドで呼ばない。
        # 1 本目は、1 段目（stage1Seconds 秒）が過ぎたら、見張りのタイマーの Tick と毎秒の件数の最小を控え、形を作れていれば 2 段目の鎖を Post で積む
        # （鎖はテストのランスペースで作った SendOrPostCallback）。作れていなければ、鎖を積まずに抜けさせる。鎖が 10 秒で回り切らなければ抜けさせる
        $script:floodSender = {
            param ($shared, [int]$index)
            $counts = $shared.Counts[$index]
            $clock = $shared.Clock
            $dispatcher = $shared.Dispatcher
            $work = [Delegate]::CreateDelegate([Func[bool]], [System.Threading.Thread].GetMethod("Yield"))
            $timeout = [TimeSpan]::FromSeconds(1)
            $send = [System.Windows.Threading.DispatcherPriority]::Send
            while (!$shared.Started -and !$shared.Stop) { Start-Sleep -Milliseconds 1 }
            while (!$shared.Stop) {
                $result = $dispatcher.Invoke($send, $timeout, $work)
                $now = $clock.Elapsed.TotalSeconds
                if ($null -ne $result) {
                    $slot = [int][Math]::Floor($now)
                    if ($slot -lt $counts.Length) { $counts[$slot]++ }
                }
                if ($index -ne 0) { continue }
                if ($shared.Phase -eq 0 -and $now -ge $shared.Stage1Seconds) {
                    $shared.TicksAtStage1 = [int]$shared.Ticks
                    $min = [long]::MaxValue
                    for ($s = 0; $s -lt $shared.Stage1Seconds; $s++) {
                        $sum = 0
                        foreach ($c in $shared.Counts) { $sum += $c[$s] }
                        if ($sum -lt $min) { $min = $sum }
                    }
                    $shared.MinPerSecond = $min
                    $shared.ShapeMade = ($shared.TicksAtStage1 -eq 0 -and $min -ge $shared.Need)
                    if ($shared.Chain -and $shared.ShapeMade) {
                        $shared.PostedAt = $now
                        $shared.Phase = 2
                        $shared.Context.Post($shared.ChainCallback, $null)
                    } else {
                        $shared.Phase = 3
                        $shared.Frame.Continue = $false
                    }
                } elseif ($shared.Phase -eq 2 -and !$shared.Done -and $now -ge $shared.PostedAt + 10) {
                    $shared.Phase = 3
                    $shared.Frame.Continue = $false
                }
            }
        }

        # threads 本で Send の仕事を積み続け、テストのスレッドのディスパッチャーを回す。chain なら 2 段目で priority の鎖（50 回の積み直し）を回す
        function invokeSendFlood {
            param ([int]$threads, [System.Windows.Threading.DispatcherPriority]$priority, [bool]$chain)

            $dispatcher = [System.Windows.Threading.Dispatcher]::CurrentDispatcher
            $shared = [hashtable]::Synchronized(@{
                    Dispatcher = $dispatcher; Clock = (New-Object System.Diagnostics.Stopwatch); Frame = (New-Object System.Windows.Threading.DispatcherFrame)
                    Counts = @(1..$threads | ForEach-Object { , [long[]]::new(30) }); Started = $false; Stop = $false; Phase = 0; Stage1Seconds = 2; Need = 9000
                    Chain = $chain; Ticks = 0; TicksAtStage1 = -1; MinPerSecond = 0; ShapeMade = $false; PostedAt = 0.0; DoneAt = 0.0; Done = $false; ChainCount = 0
                })
            $shared.Context = New-Object System.Windows.Threading.DispatcherSynchronizationContext($dispatcher, $priority)
            $shared.ChainCallback = [System.Threading.SendOrPostCallback]({
                    param ($state)
                    $shared.ChainCount++
                    if ($shared.ChainCount -ge 50) {
                        $shared.DoneAt = $shared.Clock.Elapsed.TotalSeconds
                        $shared.Done = $true
                        $shared.Frame.Continue = $false
                    } else {
                        $shared.Context.Post($shared.ChainCallback, $null)
                    }
                }.GetNewClosure())
            # 見張りの DispatcherTimer（WM_TIMER が来たかを数えるだけ）
            $watchTimer = New-Object System.Windows.Threading.DispatcherTimer
            $watchTimer.Interval = [TimeSpan]::FromMilliseconds(50)
            $watchTimer.Add_Tick({ $shared.Ticks++ }.GetNewClosure())
            $senders = @(for ($i = 0; $i -lt $threads; $i++) {
                    $ps = [powershell]::Create()
                    [void]$ps.AddScript($script:floodSender).AddArgument($shared).AddArgument($i)
                    @{ PowerShell = $ps; Handle = $ps.BeginInvoke() }
                })
            try {
                [void]$dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Send, [Action]({ $shared.Clock.Start(); $shared.Started = $true }.GetNewClosure()))
                $watchTimer.Start()
                [System.Windows.Threading.Dispatcher]::PushFrame($shared.Frame)
            } finally {
                $watchTimer.Stop()
                $shared.Stop = $true
                foreach ($sender in $senders) {
                    try { [void]$sender.PowerShell.EndInvoke($sender.Handle) } catch { Write-Host "積むスレッドの終わり: $($_.Exception.Message)" }
                    foreach ($e in $sender.PowerShell.Streams.Error) { Write-Host "積むスレッドの例外: $e" }
                    $sender.PowerShell.Dispose()
                }
            }
            return @{
                Threads = $threads; ShapeMade = $shared.ShapeMade; TicksAtStage1 = $shared.TicksAtStage1; MinPerSecond = $shared.MinPerSecond
                Done = $shared.Done; ChainSeconds = $(if ($shared.Done) { $shared.DoneAt - $shared.PostedAt } else { -1 }); Ticks = $shared.Ticks; ChainCount = $shared.ChainCount
            }
        }

        function formatFlood {
            param ($r)
            return ("積むスレッド {0} 本・1 段目の Tick {1} 回・毎秒の最小 {2} 件・鎖 {3}/50 回・鎖の秒 {4:0.000}・全体の Tick {5} 回" -f
                $r.Threads, $r.TicksAtStage1, $r.MinPerSecond, $r.ChainCount, $r.ChainSeconds, $r.Ticks)
        }

        # 1 段目の形（Tick 0 回・毎秒 9,000 件以上）を作れる本数を決める（2・4・8 本）
        $script:floodThreads = 0
        $script:floodCalibration = New-Object System.Collections.Generic.List[string]
        foreach ($threads in @(2, 4, 8)) {
            $r = invokeSendFlood $threads ([System.Windows.Threading.DispatcherPriority]::Background) $false
            $script:floodCalibration.Add((formatFlood $r))
            Write-Host ("本数の決め: " + (formatFlood $r))
            if ($r.ShapeMade) {
                $script:floodThreads = $threads
                break
            }
        }
    }

    It "本物のディスパッチャーで、別のスレッドから Send の仕事を積み続けても、タイマー無しに Background の取り出しが回り切る（<Name>）" -TestCases @(
        @{ Name = "Background"; Required = $true }
        @{ Name = "Normal"; Required = $false }
    ) {
        if ($script:floodThreads -eq 0) {
            Set-ItResult -Inconclusive -Because ("止まった回の形を作れなかった: " + ($script:floodCalibration -join " / "))
            return
        }
        $r = invokeSendFlood $script:floodThreads ([System.Windows.Threading.DispatcherPriority]$Name) $true
        Write-Host ("{0}: {1}" -f $Name, (formatFlood $r))
        if (!$r.ShapeMade) {
            Set-ItResult -Inconclusive -Because ("止まった回の形を作れなかった: " + (formatFlood $r))
            return
        }
        if (!$Required) {
            # 比べのための行。落とさない（数字は上の出力に書いた）
            return
        }
        $r.TicksAtStage1 | Should -Be 0
        $r.Done | Should -Be $true -Because (formatFlood $r)
        $r.ChainSeconds | Should -BeLessThan 5
        $r.Ticks | Should -Be 0 -Because (formatFlood $r)
    }
}
