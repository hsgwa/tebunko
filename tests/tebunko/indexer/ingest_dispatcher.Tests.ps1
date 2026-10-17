# 取り込みの振り分け（tebunko\indexer\ingest_dispatcher.ps1 の IngestDispatcher）のテスト。
# 振り分けの流れ全体（indexer.ps1 を通して動かす）は indexer.Tests.ps1 にある。ここでは、クラスだけを直接動かして確かめる。
BeforeAll {
    $script:folderCount = 0
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\indexer\indexer_lib.ps1"

    function newDispatcherOptions {
        param ([hashtable]$extra = @{})

        $options = @{
            LaneExcel = ${laneExcel}; LaneWord = ${laneWord}; LanePowerPoint = ${lanePowerPoint}; LaneReader = ${laneReader}
            LookAhead = 200; WorkerScript = { }
            PostponedReasonInUse = "使っていた"; PostponedReasonRunning = "起動していた"
        }
        foreach ($key in $extra.Keys) {
            $options[$key] = $extra[$key]
        }
        return $options
    }

    # 名前（拡張子つき）の並びのファイルを $TestDrive の下に作り、振り分け役を返す（Start はしない）。
    # lanes を省くと、すべて読み取りのレーン。missing に入れた名前は、ファイルを作らない
    function newTestDispatcher {
        param ([string[]]$names, [string[]]$lanes = @(), [string[]]$missing = @(), [hashtable]$carried = @{}, [hashtable]$extraOptions = @{})

        $script:folderCount++
        $folder = Join-Path $TestDrive "資料置き場$($script:folderCount)"
        [System.IO.Directory]::CreateDirectory($folder) | Out-Null
        $targets = New-Object System.Collections.Generic.List[object]
        $laneList = New-Object System.Collections.Generic.List[string]
        $bookFolders = New-Object System.Collections.Generic.List[string]
        for ($i = 0; $i -lt $names.Count; $i++) {
            if ($missing -notcontains $names[$i]) {
                [System.IO.File]::WriteAllText("$folder\$($names[$i])", "x")
            }
            $targets.Add([pscustomobject]@{ 相対パス = "資料\$($names[$i])" })
            $laneList.Add($(if ($lanes.Count -gt $i) { $lanes[$i] } else { ${laneReader} }))
            $bookFolders.Add("資料\$($names[$i])")
        }
        return [IngestDispatcher]::new($targets, $laneList.ToArray(), $bookFolders.ToArray(), @{ "資料" = $folder }, $carried, (newDispatcherOptions $extraOptions))
    }

    # DispatchNext を、回の終わりまで繰り返して集める
    function takeAllSteps {
        param ($dispatcher)

        $dispatcher.BeginPass()
        $steps = New-Object System.Collections.Generic.List[object]
        while ($null -ne ($step = $dispatcher.DispatchNext())) {
            $steps.Add($step)
        }
        return , $steps.ToArray()
    }
}

Describe "IngestDispatcher（作り方）" -Tag Unit {
    It "options に足りないもの（<Key>）があれば、その名前を言って例外にする" -TestCases @(
        @{ Key = "LaneExcel" }
        @{ Key = "LookAhead" }
        @{ Key = "WorkerScript" }
        @{ Key = "PostponedReasonInUse" }
    ) {
        $options = newDispatcherOptions
        $options.Remove($Key)
        { [IngestDispatcher]::new(@(), [string[]]@(), [string[]]@(), @{}, @{}, $options) } | Should -Throw "*${Key}*"
    }

    It "Start は取り込みのスレッドの入れ物だけを作る（スレッドは、そのレーンの最初のファイルを渡すときまで始めない）" {
        $dispatcher = newTestDispatcher @("a.docx")
        $dispatcher.Pool | Should -BeNullOrEmpty
        $dispatcher.UsesThreads | Should -Be $false
        $dispatcher.Start(2, @{})
        try {
            $dispatcher.UsesThreads | Should -Be $true
            $dispatcher.Pool.Readers | Should -Be 2
            $dispatcher.Pool.Workers.Count | Should -Be 0
            $dispatcher.Pool.Started.Count | Should -Be 0
            $dispatcher.Pool.Queues.Count | Should -Be 4
        } finally {
            $dispatcher.Stop()
        }
    }
}

Describe "IngestDispatcher（振り分け）" -Tag Io {
    It "並びの順に、渡せるファイルを Dispatched で返す。レーンの取り込み中が上限（読み取りは 2 つのとき 4 件）なら、次の回まで渡さない" {
        $dispatcher = newTestDispatcher @("a.docx", "b.docx", "c.docx", "d.docx", "e.docx")
        $dispatcher.Start(2, @{})
        try {
            $steps = takeAllSteps $dispatcher
            @($steps | ForEach-Object { $_.Kind }) | Should -Be @("Dispatched", "Dispatched", "Dispatched", "Dispatched")
            @($steps | ForEach-Object { $_.Number }) | Should -Be @(1, 2, 3, 4)
            $steps[0].RelPath | Should -Be "資料\a.docx"
            $steps[0].Task.SourcePath | Should -Be (Join-Path $TestDrive "資料置き場$($script:folderCount)\a.docx")
            $dispatcher.InflightCount() | Should -Be 4
            $dispatcher.Undispatched | Should -Be 1
        } finally {
            $dispatcher.Stop()
        }
    }

    It "空いているレーンのファイルを先に渡す（並びの先頭のレーンが埋まっていても、先のファイルを待たせない）" {
        $dispatcher = newTestDispatcher @("a.xlsx", "b.xlsx", "c.xlsx", "d.docx") @(${laneExcel}, ${laneExcel}, ${laneExcel}, ${laneReader})
        $dispatcher.Start(1, @{})
        try {
            $steps = takeAllSteps $dispatcher
            # Excel のレーンは 2 件まで。3 件目を飛ばして、読み取りのレーンの 4 件目を渡す
            @($steps | ForEach-Object { $_.Number }) | Should -Be @(1, 2, 4)
            $dispatcher.Undispatched | Should -Be 1
            # 取り込みが 1 件終わると、次の回で飛ばしたファイルを渡す
            $dispatcher.Pool.Results.Add(@{ RelPath = "資料\a.xlsx"; Ok = $true })
            $received = $dispatcher.Receive()
            $received.Kind | Should -Be "Completed"
            $received.Entry.Number | Should -Be 1
            $received.Result.Ok | Should -Be $true
            $again = takeAllSteps $dispatcher
            @($again | ForEach-Object { $_.Number }) | Should -Be @(3)
            $dispatcher.Undispatched | Should -Be 0
        } finally {
            $dispatcher.Stop()
        }
    }

    It "スレッドを使わないとき（UsesThreads が偽）は、取り込み中が 1 件あれば次を渡さず、預けた結果を Receive が返して取り込み中から外す" {
        $dispatcher = newTestDispatcher @("a.docx", "b.docx")
        $steps = takeAllSteps $dispatcher
        $steps.Count | Should -Be 1
        $steps[0].Number | Should -Be 1
        $dispatcher.InflightCount() | Should -Be 1
        $dispatcher.SetInlineResult(@{ RelPath = "資料\a.docx"; Ok = $true })
        $received = $dispatcher.Receive()
        $received.Kind | Should -Be "Completed"
        $received.Result.RelPath | Should -Be "資料\a.docx"
        $dispatcher.InflightCount() | Should -Be 0
        $dispatcher.Receive() | Should -BeNullOrEmpty
        @(takeAllSteps $dispatcher | ForEach-Object { $_.Number }) | Should -Be @(2)
    }

    It "前回強制終了したファイル（carried）は、回数を 1 つ増やして取り込み中の記録に移す。記録に残るのは、まだ渡していない分だけ" {
        $dispatcher = newTestDispatcher @("a.docx", "b.docx") -carried @{ "資料\a.docx" = 1; "資料\b.docx" = 3 }
        $steps = takeAllSteps $dispatcher
        $steps.Count | Should -Be 1
        $entries = @($dispatcher.IngestingEntries() | Sort-Object { $_.RelPath })
        $entries.Count | Should -Be 2
        $entries[0].RelPath | Should -Be "資料\a.docx"
        $entries[0].Count | Should -Be 2
        $entries[1].RelPath | Should -Be "資料\b.docx"
        $entries[1].Count | Should -Be 3
    }

    It "取り込む直前に元のファイルが無くなっていたら Vanished で返し、渡さずに済ませる（取り込み中にはしない）" {
        $dispatcher = newTestDispatcher @("a.docx", "b.docx") -missing @("a.docx")
        $steps = takeAllSteps $dispatcher
        $steps[0].Kind | Should -Be "Vanished"
        $steps[0].RelPath | Should -Be "資料\a.docx"
        $steps[1].Kind | Should -Be "Dispatched"
        $dispatcher.Undispatched | Should -Be 0
        $dispatcher.InflightCount() | Should -Be 1
    }

    It "クロール対象フォルダごと見えなくなったら、FolderLost と残りの件数を残して回を終える" {
        $dispatcher = newTestDispatcher @("a.docx", "b.docx", "c.docx")
        Remove-Item -LiteralPath (Join-Path $TestDrive "資料置き場$($script:folderCount)") -Recurse -Force
        $steps = takeAllSteps $dispatcher
        $steps.Count | Should -Be 0
        $dispatcher.FolderLost | Should -Be (Join-Path $TestDrive "資料置き場$($script:folderCount)")
        $dispatcher.Remaining | Should -Be 3
        $dispatcher.StopIfRequested($true) | Should -Be $false
    }

    It "options の TestPath で、元のファイル・フォルダの有無を差し替えられる（ファイルが無くフォルダがあれば Vanished、フォルダも無ければ FolderLost）" -TestCases @(
        @{ FolderExists = $true; Kind = "Vanished"; Lost = "" }
        @{ FolderExists = $false; Kind = $null; Lost = "あり" }
    ) {
        $folderExists = $FolderExists
        $dispatcher = newTestDispatcher @("a.docx") -extraOptions @{ TestPath = { param ($path, $kind) $kind -eq "Folder" -and $folderExists }.GetNewClosure() }
        $steps = takeAllSteps $dispatcher
        if ($Kind) {
            $steps[0].Kind | Should -Be $Kind
        } else {
            $steps.Count | Should -Be 0
            $dispatcher.FolderLost | Should -Not -BeNullOrEmpty
        }
    }

    It "諦めたレーンのファイルは、渡さずに Abandoned で返す" {
        $dispatcher = newTestDispatcher @("a.ppt", "b.docx") @(${lanePowerPoint}, ${laneReader})
        $dispatcher.Start(1, @{})
        try {
            $dispatcher.Pool.LaneStates[${lanePowerPoint}].Abandoned = $true
            $steps = takeAllSteps $dispatcher
            $steps[0].Kind | Should -Be "Abandoned"
            $steps[0].Number | Should -Be 1
            $steps[0].Row.相対パス | Should -Be "資料\a.ppt"
            $steps[1].Kind | Should -Be "Dispatched"
            $dispatcher.InflightCount() | Should -Be 1
            $dispatcher.Undispatched | Should -Be 0
        } finally {
            $dispatcher.Stop()
        }
    }

    It "StopIfRequested は、中止を求められて渡していないファイルがあるときに、1 回だけ `$true を返し、以後は渡さない" {
        $dispatcher = newTestDispatcher @("a.docx", "b.docx")
        $dispatcher.StopIfRequested($false) | Should -Be $false
        $dispatcher.StopIfRequested($true) | Should -Be $true
        $dispatcher.Stopped | Should -Be $true
        $dispatcher.Remaining | Should -Be 2
        $dispatcher.StopIfRequested($true) | Should -Be $false
        (takeAllSteps $dispatcher).Count | Should -Be 0
    }
}

Describe "IngestDispatcher（諦めたレーン）" -Tag Io {
    It "TakeAbandonedLane は、スレッドを始めたレーンのうち、諦めて後回しを済ませていないものを 1 つずつ、取り込み中のファイルといっしょに返す" {
        $dispatcher = newTestDispatcher @("a.ppt", "b.docx", "c.ppt") @(${lanePowerPoint}, ${laneReader}, ${lanePowerPoint})
        $dispatcher.Start(1, @{})
        try {
            $dispatcher.TakeAbandonedLane() | Should -BeNullOrEmpty
            $steps = takeAllSteps $dispatcher
            $steps.Count | Should -Be 3
            # 始まったが諦めていないレーンは、返さない
            $dispatcher.Pool.Started[${lanePowerPoint}] = $true
            $dispatcher.Pool.Started[${laneReader}] = $true
            $dispatcher.TakeAbandonedLane() | Should -BeNullOrEmpty
            $dispatcher.InflightCount() | Should -Be 3
            # スレッドを始めていないレーンは、諦めていても返さない
            $dispatcher.Pool.LaneStates[${laneWord}].Abandoned = $true
            $dispatcher.TakeAbandonedLane() | Should -BeNullOrEmpty
            # 始めて、諦めたレーンは、取り込み中のファイルといっしょに返す
            $dispatcher.Pool.LaneStates[${lanePowerPoint}].Abandoned = $true
            $taken = $dispatcher.TakeAbandonedLane()
            $taken.Lane | Should -Be ${lanePowerPoint}
            @($taken.Entries | ForEach-Object { $_.Number } | Sort-Object) | Should -Be @(1, 3)
            $dispatcher.InflightCount() | Should -Be 1
            # 済ませたレーンは、もう返さない
            $dispatcher.TakeAbandonedLane() | Should -BeNullOrEmpty
        } finally {
            $dispatcher.Stop()
        }
    }

    It "Receive は、諦めたレーンがあれば、戻らない結果を待たずに `$null を返す。結果があればそちらを返す" {
        $dispatcher = newTestDispatcher @("a.docx") @(${lanePowerPoint})
        $dispatcher.Start(1, @{})
        try {
            [void](takeAllSteps $dispatcher)
            $dispatcher.Pool.Started[${lanePowerPoint}] = $true
            $dispatcher.Pool.LaneStates[${lanePowerPoint}].Abandoned = $true
            $dispatcher.Pool.Results.Add(@{ RelPath = "資料\a.docx"; Ok = $true })
            $dispatcher.Receive().Kind | Should -Be "Completed"
            $dispatcher.Receive() | Should -BeNullOrEmpty
        } finally {
            $dispatcher.Stop()
        }
    }

    It "後回しにしたファイルについて、諦めたレーンのスレッドが後から返した結果は捨てる（取り込み中に無いため）" {
        $dispatcher = newTestDispatcher @("a.ppt") @(${lanePowerPoint})
        $dispatcher.Start(1, @{})
        try {
            [void](takeAllSteps $dispatcher)
            $dispatcher.InflightCount() | Should -Be 1
            # 諦めたレーンのファイルを後回しにして、取り込み中から外す
            $dispatcher.Pool.Started[${lanePowerPoint}] = $true
            $dispatcher.Pool.LaneStates[${lanePowerPoint}].Abandoned = $true
            $taken = $dispatcher.TakeAbandonedLane()
            $taken.Entries.Count | Should -Be 1
            $dispatcher.InflightCount() | Should -Be 0
            # そのあとでスレッドが返した結果は、反映せずに捨てる
            $dispatcher.Pool.Results.Add(@{ RelPath = "資料\a.ppt"; Ok = $true })
            $dispatcher.Receive() | Should -BeNullOrEmpty
            $dispatcher.Pool.Results.Count | Should -Be 0
        } finally {
            $dispatcher.Stop()
        }
    }

    It "Postponed は、諦めたレーンの結果なら『使っていた』、そうでなければ『起動していた』の理由で返す（諦めた: <Abandoned>）" -TestCases @(
        @{ Abandoned = $true; Reason = "使っていた" }
        @{ Abandoned = $false; Reason = "起動していた" }
    ) {
        $dispatcher = newTestDispatcher @("a.ppt") @(${lanePowerPoint})
        $dispatcher.Start(1, @{})
        try {
            [void](takeAllSteps $dispatcher)
            $dispatcher.Pool.LaneStates[${lanePowerPoint}].Abandoned = $Abandoned
            $dispatcher.Pool.LaneStates[${lanePowerPoint}].Handled = $true
            $dispatcher.Pool.Results.Add(@{ RelPath = "資料\a.ppt"; Postponed = $true })
            $received = $dispatcher.Receive()
            $received.Kind | Should -Be "Postponed"
            $received.Reason | Should -Be $Reason
            $dispatcher.InflightCount() | Should -Be 0
        } finally {
            $dispatcher.Stop()
        }
    }

    It "回し直し先（Word のレーン）が諦めたレーンなら、渡さずに RerouteAbandoned で返して取り込み中から外す" {
        $dispatcher = newTestDispatcher @("a.docx")
        $dispatcher.Start(1, @{})
        try {
            [void](takeAllSteps $dispatcher)
            $dispatcher.Pool.LaneStates[${laneWord}].Abandoned = $true
            $dispatcher.Pool.Results.Add(@{ RelPath = "資料\a.docx"; Reroute = $true })
            $received = $dispatcher.Receive()
            $received.Kind | Should -Be "RerouteAbandoned"
            $received.Reason | Should -Be "使っていた"
            $received.Entry.Number | Should -Be 1
            $dispatcher.InflightCount() | Should -Be 0
            $dispatcher.Pool.Workers.Count | Should -Be 0
        } finally {
            $dispatcher.Stop()
        }
    }

    It "回し直し先（Word のレーン）が空いていれば、そのレーンの列に渡し直して取り込み中のまま、諦めたあとはそのレーンで後回しにされる" {
        $dispatcher = newTestDispatcher @("a.docx")
        $dispatcher.Start(1, @{})
        try {
            [void](takeAllSteps $dispatcher)
            # Word のレーンのスレッドは始めない（始めたことにして、列に渡すだけにする）
            $dispatcher.Pool.Started[${laneWord}] = $true
            $dispatcher.Pool.Results.Add(@{ RelPath = "資料\a.docx"; Reroute = $true })
            $dispatcher.Receive() | Should -BeNullOrEmpty
            $dispatcher.InflightCount() | Should -Be 1
            $dispatcher.Pool.Queues[${laneWord}].Count | Should -Be 1
            $dispatcher.Pool.LaneStates[${laneWord}].Abandoned = $true
            $taken = $dispatcher.TakeAbandonedLane()
            $taken.Lane | Should -Be ${laneWord}
            $taken.Entries[0].Number | Should -Be 1
        } finally {
            $dispatcher.Stop()
        }
    }

    It "Receive は、諦めたレーンのスレッドが正常に終わっていても失敗にしない。諦めていないレーンのスレッドが終わっていれば失敗にする（諦めた: <Abandoned>）" -TestCases @(
        @{ Abandoned = $true }
        @{ Abandoned = $false }
    ) {
        $runspaces = New-Object System.Collections.Generic.List[object]
        try {
            $dispatcher = newTestDispatcher @("a.ppt") @(${lanePowerPoint})
            $dispatcher.Start(1, @{})
            [void](takeAllSteps $dispatcher)
            $pool = $dispatcher.Pool
            $pool.Started[${lanePowerPoint}] = $true
            $pool.LaneStates[${lanePowerPoint}].Abandoned = $Abandoned
            $pool.LaneStates[${lanePowerPoint}].Handled = $true
            $runspace = [runspacefactory]::CreateRunspace()
            $runspace.Open()
            $ps = [powershell]::Create()
            $ps.Runspace = $runspace
            $runspaces.Add(@{ PowerShell = $ps; Runspace = $runspace })
            # すぐ終わるスレッド
            [void]$ps.AddScript({ 1 })
            $handle = $ps.BeginInvoke()
            [void]$handle.AsyncWaitHandle.WaitOne(10000)
            $worker = @{ PowerShell = $ps; Runspace = $runspace; Handle = $handle; Ended = $false; Lane = ${lanePowerPoint} }
            $pool.Workers.Add($worker)
            if ($Abandoned) {
                # 失敗にしないで待ち続けるので、別のスレッドが少しあとに結果を入れる
                $feeder = [powershell]::Create()
                $feederRunspace = [runspacefactory]::CreateRunspace()
                $feederRunspace.Open()
                $feeder.Runspace = $feederRunspace
                $runspaces.Add(@{ PowerShell = $feeder; Runspace = $feederRunspace })
                [void]$feeder.AddScript({ param ($results) Start-Sleep -Milliseconds 1200; $results.Add(@{ RelPath = "資料\a.ppt"; Ok = $true }) }).AddArgument($pool.Results)
                [void]$feeder.BeginInvoke()
                $dispatcher.Receive().Result.RelPath | Should -Be "資料\a.ppt"
                $worker.Ended | Should -Be $false
            } else {
                { $dispatcher.Receive() } | Should -Throw "*取り込みのスレッドが止まりました*"
            }
        } finally {
            foreach ($item in $runspaces) {
                try { $item.PowerShell.Dispose() } catch {}
                try { $item.Runspace.Dispose() } catch {}
            }
        }
    }

    It "Stop は、諦めたレーンの固まったスレッドが終わるのを待たずに戻る（止める依頼だけを出す）" {
        $dispatcher = newTestDispatcher @("a.ppt") @(${lanePowerPoint})
        $dispatcher.Start(1, @{})
        $pool = $dispatcher.Pool
        $pool.Started[${lanePowerPoint}] = $true
        $pool.LaneStates[${lanePowerPoint}].Abandoned = $true
        $gate = New-Object System.Threading.ManualResetEvent($false)
        $runspace = [runspacefactory]::CreateRunspace()
        $runspace.Open()
        $ps = [powershell]::Create()
        $ps.Runspace = $runspace
        [void]$ps.AddScript({ param ($gate) [void]$gate.WaitOne() }).AddArgument($gate)
        $worker = @{ PowerShell = $ps; Runspace = $runspace; Handle = $ps.BeginInvoke(); Ended = $false; Lane = ${lanePowerPoint} }
        $pool.Workers.Add($worker)
        try {
            $watch = [System.Diagnostics.Stopwatch]::StartNew()
            $dispatcher.Stop()
            $watch.Elapsed.TotalSeconds | Should -BeLessThan 20
            $worker.Ended | Should -Be $true
        } finally {
            [void]$gate.Set()
            [void]$worker.Handle.AsyncWaitHandle.WaitOne(10000)
            $ps.Dispose()
            $runspace.Dispose()
        }
    }
}
