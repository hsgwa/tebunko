# 取り込みの振り分け（tebunko\indexer\ingest_dispatcher.ps1 の IngestDispatcher）のテスト。
# 振り分けの流れ全体（indexer.ps1 を通して動かす）は indexer.Tests.ps1 にある。ここでは、クラスだけを直接動かして確かめる。
BeforeAll {
    $script:folderCount = 0
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\indexer\indexer_lib.ps1"

    function newDispatcherOptions {
        param ([hashtable]$extra = @{})

        $options = @{
            LookAhead = 200; WorkerScript = { }
            PostponedReasonInUse = "使っていた"; PostponedReasonRunning = "起動していた"
        }
        foreach ($key in $extra.Keys) {
            $options[$key] = $extra[$key]
        }
        return $options
    }

    # 名前（拡張子つき）の並びのファイルを $TestDrive の下に作り、振り分け役を返す（Start はしない）。
    # apps は、ファイルごとに最初に要る Office のアプリ。省くと、すべて Office が要らない（空）。missing に入れた名前は、ファイルを作らない
    function newTestDispatcher {
        param ([string[]]$names, [string[]]$apps = @(), [string[]]$missing = @(), [hashtable]$carried = @{}, [hashtable]$extraOptions = @{})

        $script:folderCount++
        $folder = Join-Path $TestDrive "資料置き場$($script:folderCount)"
        [System.IO.Directory]::CreateDirectory($folder) | Out-Null
        $targets = New-Object System.Collections.Generic.List[object]
        $appList = New-Object System.Collections.Generic.List[string]
        $bookFolders = New-Object System.Collections.Generic.List[string]
        for ($i = 0; $i -lt $names.Count; $i++) {
            if ($missing -notcontains $names[$i]) {
                [System.IO.File]::WriteAllText("$folder\$($names[$i])", "x")
            }
            $targets.Add([pscustomobject]@{ 相対パス = "資料\$($names[$i])" })
            $appList.Add($(if ($apps.Count -gt $i) { $apps[$i] } else { "" }))
            $bookFolders.Add("資料\$($names[$i])")
        }
        return [IngestDispatcher]::new($targets, $appList.ToArray(), $bookFolders.ToArray(), @{ "資料" = $folder }, $carried, (newDispatcherOptions $extraOptions))
    }

    # スレッドを始めたことにする（本物のスレッドは作らない。終わっている入れ物を置く）
    function startFakeWorker {
        param ($dispatcher, [int]$index)

        $ps = [pscustomobject]@{}
        $ps | Add-Member -MemberType ScriptMethod -Name Dispose -Value { }
        $dispatcher.Pool.Slots[$index].Worker = @{ PowerShell = $ps; Runspace = $ps; Handle = [pscustomobject]@{ IsCompleted = $true }; Ended = $true }
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
        @{ Key = "LookAhead" }
        @{ Key = "WorkerScript" }
        @{ Key = "PostponedReasonInUse" }
        @{ Key = "PostponedReasonRunning" }
    ) {
        $options = newDispatcherOptions
        $options.Remove($Key)
        { [IngestDispatcher]::new(@(), [string[]]@(), [string[]]@(), @{}, @{}, $options) } | Should -Throw "*${Key}*"
    }

    It "Start は取り込みのワーカーの入れ物だけを作る（スレッドは、そのワーカーに最初のファイルを渡すときまで始めない）" {
        $dispatcher = newTestDispatcher @("a.docx")
        $dispatcher.Pool | Should -BeNullOrEmpty
        $dispatcher.UsesThreads | Should -Be $false
        $dispatcher.Start(3, @{})
        try {
            $dispatcher.UsesThreads | Should -Be $true
            $dispatcher.Pool.Slots.Count | Should -Be 3
            @($dispatcher.Pool.Slots | Where-Object { $null -ne $_.Worker }).Count | Should -Be 0
            @($dispatcher.Pool.Slots | Where-Object { $_.Busy -ne 0 -or $_.Owns.Count -ne 0 }).Count | Should -Be 0
        } finally {
            $dispatcher.Stop()
        }
    }
}

Describe "IngestDispatcher（振り分け）" -Tag Io {
    It "並びの順に、渡せるファイルを Dispatched で返す。ワーカーごとの渡してある数が上限（2 つのワーカーなら 4 件）なら、次の回まで渡さない" {
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

    It "Office が要るファイルは持ち主のワーカーにだけ渡し、持ち主が満杯なら、Office が要らないファイルを先に渡す。取り込みが 1 件終わると、次の回で飛ばしたファイルを渡す" {
        $dispatcher = newTestDispatcher @("a.xlsx", "b.xlsx", "c.xlsx", "d.docx") @("Excel", "Excel", "Excel", "")
        $dispatcher.Start(2, @{})
        try {
            $steps = takeAllSteps $dispatcher
            # Excel の持ち主は 2 件まで。3 件目を飛ばして、Office を持たないワーカーに 4 件目を渡す
            @($steps | ForEach-Object { $_.Number }) | Should -Be @(1, 2, 4)
            @($steps | ForEach-Object { $_.Worker }) | Should -Be @(0, 0, 1)
            @($dispatcher.Pool.Slots[0].Owns) | Should -Be @("Excel")
            $dispatcher.Pool.Slots[1].Owns.Count | Should -Be 0
            $dispatcher.Undispatched | Should -Be 1
            $dispatcher.Pool.Slots[0].Busy | Should -Be 2
            $dispatcher.Pool.Slots[1].Busy | Should -Be 1
            $dispatcher.Pool.Results.Add(@{ RelPath = "資料\a.xlsx"; Ok = $true })
            $received = $dispatcher.Receive()
            $received.Kind | Should -Be "Completed"
            $received.Entry.Number | Should -Be 1
            $received.Result.Ok | Should -Be $true
            $dispatcher.Pool.Slots[0].Busy | Should -Be 1
            $again = takeAllSteps $dispatcher
            @($again | ForEach-Object { $_.Number }) | Should -Be @(3)
            $again[0].Worker | Should -Be 0
            $dispatcher.Undispatched | Should -Be 0
        } finally {
            $dispatcher.Stop()
        }
    }

    It "ワーカーに渡すファイルの Task は、Office が要るか（Office）を持つ。スレッドを使わないときは持たない" {
        $dispatcher = newTestDispatcher @("a.xlsx", "b.docx") @("Excel", "")
        $dispatcher.Start(2, @{})
        try {
            $steps = takeAllSteps $dispatcher
            $steps[0].Task.Office | Should -Be $true
            $steps[1].Task.Office | Should -Be $false
        } finally {
            $dispatcher.Stop()
        }
        $inline = newTestDispatcher @("a.xlsx") @("Excel")
        $step = (takeAllSteps $inline)[0]
        $step.Worker | Should -Be -1
        $step.Task.ContainsKey("Office") | Should -Be $false
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

    It "前回強制終了したファイル（carried）は、渡すときに回数を 1 つ増やす（GetInflightCount が今回の分まで含めた回数を返す。渡していない・取り込み中でないものは 0）" {
        $dispatcher = newTestDispatcher @("a.docx", "b.docx") -carried @{ "資料\a.docx" = 1; "資料\b.docx" = 3 }
        $dispatcher.GetInflightCount("資料\a.docx") | Should -Be 0
        $steps = takeAllSteps $dispatcher
        $steps.Count | Should -Be 1
        $dispatcher.GetInflightCount($steps[0].RelPath) | Should -BeGreaterThan 1
        $dispatcher.GetInflightCount("資料\無い.docx") | Should -Be 0
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

    It "元のファイルが無くなっていたファイルは、渡し先のワーカーの渡してある数にも、持ち主にも数えない" {
        $dispatcher = newTestDispatcher @("a.xlsx") @("Excel") -missing @("a.xlsx")
        $dispatcher.Start(2, @{})
        try {
            (takeAllSteps $dispatcher)[0].Kind | Should -Be "Vanished"
            @($dispatcher.Pool.Slots | Where-Object { $_.Busy -ne 0 -or $_.Owns.Count -ne 0 }).Count | Should -Be 0
        } finally {
            $dispatcher.Stop()
        }
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
        @{ FolderExists = $true; Kind = "Vanished" }
        @{ FolderExists = $false; Kind = $null }
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

    It "持ち主が諦めたワーカーのアプリのファイルは、渡さずに Abandoned で返す。Office が要らないファイルは、諦めていないワーカーに渡す" {
        $dispatcher = newTestDispatcher @("a.ppt", "b.docx") @("PowerPoint", "")
        $dispatcher.Start(2, @{})
        try {
            $dispatcher.Pool.Slots[0].Owns.Add("PowerPoint")
            $dispatcher.Pool.Slots[0].State.Abandoned = $true
            $steps = takeAllSteps $dispatcher
            $steps[0].Kind | Should -Be "Abandoned"
            $steps[0].Number | Should -Be 1
            $steps[0].Row.相対パス | Should -Be "資料\a.ppt"
            $steps[1].Kind | Should -Be "Dispatched"
            $steps[1].Worker | Should -Be 1
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

    It "中止を求められているとき、Receive は止められた結果を取り込み中から外して `"Stopped`" で返し、初めの 1 件だけ First を真にして残りを数える" -TestCases @(
        @{ Name = "止められた結果"; Result = @{ RelPath = "資料\a.docx"; Stopped = $true } }
        @{ Name = "別の経路へ回す結果"; Result = @{ RelPath = "資料\a.docx"; Reroute = $true } }
    ) {
        param($Name, $Result)
        $dispatcher = newTestDispatcher @("a.docx", "b.docx")
        [void](takeAllSteps $dispatcher)
        $dispatcher.SetInlineResult($Result)
        $received = $dispatcher.Receive($true)
        $received.Kind | Should -Be "Stopped"
        $received.First | Should -Be $true
        $received.RelPath | Should -Be "資料\a.docx"
        $dispatcher.InflightCount() | Should -Be 0
        $dispatcher.Stopped | Should -Be $true
        $dispatcher.Remaining | Should -Be 2
    }

    It "Receive が受けた結果は LastResult に残り、受ける結果が無ければ `$null に戻る（結果: <Name>）" -TestCases @(
        @{ Name = "取り込んだ"; Result = @{ RelPath = "資料\a.docx"; Ok = $true }; Stop = $false }
        @{ Name = "止められた"; Result = @{ RelPath = "資料\a.docx"; Stopped = $true }; Stop = $false }
        @{ Name = "回し直しを止めた"; Result = @{ RelPath = "資料\a.docx"; Reroute = $true }; Stop = $true }
    ) {
        param($Name, $Result, $Stop)
        $dispatcher = newTestDispatcher @("a.docx", "b.docx")
        [void](takeAllSteps $dispatcher)
        $dispatcher.SetInlineResult($Result)
        [void]$dispatcher.Receive($Stop)
        $dispatcher.LastResult.RelPath | Should -Be "資料\a.docx"
        # 次の Receive は、受ける結果が無いので、前の結果を引き継がない
        $dispatcher.Receive($false) | Should -BeNullOrEmpty
        $dispatcher.LastResult | Should -BeNullOrEmpty
    }
}

Describe "IngestDispatcher（諦めたワーカー）" -Tag Io {
    It "TakeAbandonedWorker は、スレッドを始めたワーカーのうち、諦めて後回しを済ませていないものを 1 つずつ、取り込み中のファイルといっしょに返す" {
        $dispatcher = newTestDispatcher @("a.ppt", "b.docx", "c.ppt") @("PowerPoint", "", "PowerPoint")
        $dispatcher.Start(3, @{})
        try {
            $dispatcher.TakeAbandonedWorker() | Should -BeNullOrEmpty
            $steps = takeAllSteps $dispatcher
            $steps.Count | Should -Be 3
            @($steps | ForEach-Object { $_.Worker }) | Should -Be @(0, 1, 0)
            # 始まったが諦めていないワーカーは、返さない
            startFakeWorker $dispatcher 0
            startFakeWorker $dispatcher 1
            $dispatcher.TakeAbandonedWorker() | Should -BeNullOrEmpty
            $dispatcher.InflightCount() | Should -Be 3
            # スレッドを始めていないワーカーは、諦めていても返さない
            $dispatcher.Pool.Slots[2].State.Abandoned = $true
            $dispatcher.TakeAbandonedWorker() | Should -BeNullOrEmpty
            # 始めて、諦めたワーカーは、取り込み中のファイルといっしょに返す
            $dispatcher.Pool.Slots[0].State.Abandoned = $true
            $taken = $dispatcher.TakeAbandonedWorker()
            $taken.Worker | Should -Be 0
            @($taken.Entries | ForEach-Object { $_.Number } | Sort-Object) | Should -Be @(1, 3)
            $dispatcher.InflightCount() | Should -Be 1
            $dispatcher.Pool.Slots[0].Busy | Should -Be 0
            # 済ませたワーカーは、もう返さない
            $dispatcher.TakeAbandonedWorker() | Should -BeNullOrEmpty
        } finally {
            $dispatcher.Stop()
        }
    }

    It "Receive は、諦めたワーカーがあれば、戻らない結果を待たずに `$null を返す。結果があればそちらを返す" {
        $dispatcher = newTestDispatcher @("a.docx") @("PowerPoint")
        $dispatcher.Start(1, @{})
        try {
            [void](takeAllSteps $dispatcher)
            startFakeWorker $dispatcher 0
            $dispatcher.Pool.Slots[0].State.Abandoned = $true
            $dispatcher.Pool.Results.Add(@{ RelPath = "資料\a.docx"; Ok = $true })
            $dispatcher.Receive().Kind | Should -Be "Completed"
            $dispatcher.Receive() | Should -BeNullOrEmpty
        } finally {
            $dispatcher.Stop()
        }
    }

    It "後回しにしたファイルについて、諦めたワーカーが後から返した結果は捨てる（取り込み中に無いため）" {
        $dispatcher = newTestDispatcher @("a.ppt") @("PowerPoint")
        $dispatcher.Start(1, @{})
        try {
            [void](takeAllSteps $dispatcher)
            $dispatcher.InflightCount() | Should -Be 1
            # 諦めたワーカーのファイルを後回しにして、取り込み中から外す
            startFakeWorker $dispatcher 0
            $dispatcher.Pool.Slots[0].State.Abandoned = $true
            $taken = $dispatcher.TakeAbandonedWorker()
            $taken.Entries.Count | Should -Be 1
            $dispatcher.InflightCount() | Should -Be 0
            # そのあとでワーカーが返した結果は、反映せずに捨てる
            $dispatcher.Pool.Results.Add(@{ RelPath = "資料\a.ppt"; Ok = $true })
            $dispatcher.Receive() | Should -BeNullOrEmpty
            # 捨てた結果は、時間にも数えない
            $dispatcher.LastResult | Should -BeNullOrEmpty
            $dispatcher.Pool.Results.Count | Should -Be 0
        } finally {
            $dispatcher.Stop()
        }
    }

    It "Postponed は、諦めたワーカーの結果なら『使っていた』、そうでなければ『起動していた』の理由で返す（諦めた: <Abandoned>）" -TestCases @(
        @{ Abandoned = $true; Reason = "使っていた" }
        @{ Abandoned = $false; Reason = "起動していた" }
    ) {
        $dispatcher = newTestDispatcher @("a.ppt") @("PowerPoint")
        $dispatcher.Start(1, @{})
        try {
            [void](takeAllSteps $dispatcher)
            $dispatcher.Pool.Slots[0].State.Abandoned = $Abandoned
            $dispatcher.Pool.Slots[0].State.Handled = $true
            $dispatcher.Pool.Results.Add(@{ RelPath = "資料\a.ppt"; Postponed = $true })
            $received = $dispatcher.Receive()
            $received.Kind | Should -Be "Postponed"
            $received.Reason | Should -Be $Reason
            $dispatcher.InflightCount() | Should -Be 0
        } finally {
            $dispatcher.Stop()
        }
    }

    It "回し直し先（PowerPoint）の持ち主が諦めたワーカーなら、渡さずに RerouteAbandoned で返して取り込み中から外す" {
        $dispatcher = newTestDispatcher @("a.pptx")
        $dispatcher.Start(2, @{})
        try {
            [void](takeAllSteps $dispatcher)
            $dispatcher.Pool.Slots[1].Owns.Add("PowerPoint")
            $dispatcher.Pool.Slots[1].State.Abandoned = $true
            $dispatcher.Pool.Results.Add(@{ RelPath = "資料\a.pptx"; Reroute = $true })
            $received = $dispatcher.Receive()
            $received.Kind | Should -Be "RerouteAbandoned"
            $received.Reason | Should -Be "使っていた"
            $received.Entry.Number | Should -Be 1
            # Receive が飲む回し直しの結果も、LastResult に残る（種類ごとの「回し直し」に数えるため）
            $dispatcher.LastResult.Reroute | Should -Be $true
            $dispatcher.InflightCount() | Should -Be 0
            @($dispatcher.Pool.Slots | Where-Object { $null -ne $_.Worker }).Count | Should -Be 0
            $dispatcher.Pool.Slots[0].Busy | Should -Be 0
        } finally {
            $dispatcher.Stop()
        }
    }

    It "回し直し先（Word）に持ち主がいなければ、空いているワーカーを持ち主にして列に渡し直して取り込み中のまま、諦めたあとはそのワーカーで後回しにされる" {
        $dispatcher = newTestDispatcher @("a.docx")
        $dispatcher.Start(2, @{})
        try {
            [void](takeAllSteps $dispatcher)
            # スレッドは始めない（始めたことにして、列に渡すだけにする）
            startFakeWorker $dispatcher 0
            startFakeWorker $dispatcher 1
            $dispatcher.Pool.Results.Add(@{ RelPath = "資料\a.docx"; Reroute = $true })
            $dispatcher.Receive() | Should -BeNullOrEmpty
            $dispatcher.InflightCount() | Should -Be 1
            $owners = @(0, 1 | Where-Object { $dispatcher.Pool.Slots[$_].Owns -contains "Word" })
            $owners.Count | Should -Be 1
            $owner = $owners[0]
            $dispatcher.Pool.Slots[$owner].Queue.Count | Should -Be 1
            $dispatcher.Pool.Slots[$owner].Busy | Should -Be 1
            $dispatcher.Pool.Slots[1 - $owner].Busy | Should -Be 0
            $dispatcher.Pool.Slots[$owner].State.Abandoned = $true
            $taken = $dispatcher.TakeAbandonedWorker()
            $taken.Worker | Should -Be $owner
            $taken.Entries[0].Number | Should -Be 1
        } finally {
            $dispatcher.Stop()
        }
    }

    It "回し直しは、持ち主が取り込み中でも（渡してある数が上限でも）持ち主の列に渡す" {
        $dispatcher = newTestDispatcher @("a.docx", "b.doc", "c.doc") @("", "Word", "Word")
        $dispatcher.Start(2, @{})
        try {
            $steps = takeAllSteps $dispatcher
            @($steps | ForEach-Object { $_.Worker }) | Should -Be @(0, 1, 1)
            startFakeWorker $dispatcher 0
            startFakeWorker $dispatcher 1
            $dispatcher.Pool.Results.Add(@{ RelPath = "資料\a.docx"; Reroute = $true })
            $dispatcher.Receive() | Should -BeNullOrEmpty
            $dispatcher.Pool.Slots[1].Busy | Should -Be 3
            $dispatcher.Pool.Slots[1].Queue.Count | Should -Be 1
            $dispatcher.Pool.Slots[0].Busy | Should -Be 0
        } finally {
            $dispatcher.Stop()
        }
    }

    It "Receive は、諦めたワーカーが正常に終わっていても失敗にしない。諦めていないワーカーが終わっていれば失敗にする（諦めた: <Abandoned>）" -TestCases @(
        @{ Abandoned = $true }
        @{ Abandoned = $false }
    ) {
        $runspaces = New-Object System.Collections.Generic.List[object]
        try {
            $dispatcher = newTestDispatcher @("a.ppt") @("PowerPoint")
            $dispatcher.Start(1, @{})
            [void](takeAllSteps $dispatcher)
            $pool = $dispatcher.Pool
            $pool.Slots[0].State.Abandoned = $Abandoned
            $pool.Slots[0].State.Handled = $true
            $runspace = [runspacefactory]::CreateRunspace()
            $runspace.Open()
            $ps = [powershell]::Create()
            $ps.Runspace = $runspace
            $runspaces.Add(@{ PowerShell = $ps; Runspace = $runspace })
            # すぐ終わるスレッド
            [void]$ps.AddScript({ 1 })
            $handle = $ps.BeginInvoke()
            [void]$handle.AsyncWaitHandle.WaitOne(10000)
            $worker = @{ PowerShell = $ps; Runspace = $runspace; Handle = $handle; Ended = $false }
            $pool.Slots[0].Worker = $worker
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

    It "Stop は、諦めたワーカーの固まったスレッドが終わるのを待たずに戻る（止める依頼だけを出す）" {
        $dispatcher = newTestDispatcher @("a.ppt") @("PowerPoint")
        $dispatcher.Start(1, @{})
        $pool = $dispatcher.Pool
        $pool.Slots[0].State.Abandoned = $true
        $gate = New-Object System.Threading.ManualResetEvent($false)
        $runspace = [runspacefactory]::CreateRunspace()
        $runspace.Open()
        $ps = [powershell]::Create()
        $ps.Runspace = $runspace
        [void]$ps.AddScript({ param ($gate) [void]$gate.WaitOne() }).AddArgument($gate)
        $worker = @{ PowerShell = $ps; Runspace = $runspace; Handle = $ps.BeginInvoke(); Ended = $false }
        $pool.Slots[0].Worker = $worker
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
