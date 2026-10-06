# このツールが起動した Excel・Word・PowerPoint の PID の記録と、実行中の Office の一覧・強制終了。
# 止めてよいのは「記録があり、記録と今のプロセスが合うもの」だけ。記録の無いもの（利用者・ほかのアプリのもの）は一覧に出ても止める対象にしない。

# ----------------------------------------------------------------------------
# 検索・Officeプロセス・画面（gui.ps1）で共有する処理
# ----------------------------------------------------------------------------

# 強制終了の対象: プロセス名 → 表示名
${officeProcessNames} = [ordered]@{ EXCEL = "Excel"; WINWORD = "Word"; POWERPNT = "PowerPoint" }

# 書きかけの記録（*.tmp）をこの時間より古ければ消す（書いている最中のものを消さないため）
${officeRecordTmpMaxAge} = [TimeSpan]::FromHours(1)

function getOfficeStartTicks {
    # プロセスの起動時刻（UTC の Ticks）。読めない（管理者として動いているなど）ときは $null
    param (
        $process
    )

    try {
        return [long]$process.StartTime.ToUniversalTime().Ticks
    } catch {
        return $null
    }
}

# ----------------------------------------------------------------------------
# 記録（<ワークスペース>\office_pids\<PC の鍵>\<PID>.txt）
# 1 行・タブ区切り: <プロセス名> <起動時刻の Ticks> <持ち主の PID> <持ち主の起動時刻の Ticks>
# ----------------------------------------------------------------------------

function addOfficeRecord {
    # 起動した Office の記録を 1 つ書く。書けたら $true。場所が空・書けない（共有に届かない・権限が無い）ときは $false（呼ぶ側は取り込みを続ける）。
    # 持ち主は、省けばこのプロセス（記録を書いたプロセス）。書きかけを読ませないよう、.tmp に書いてから名前を変える
    param (
        [string]$dir,
        [int]$id,
        [string]$processName,
        [long]$startTicks,
        [int]$ownerId = 0,
        [long]$ownerStartTicks = 0
    )

    if (!$dir -or $startTicks -le 0) {
        return $false
    }
    $tmp = Join-Path $dir "$id.tmp"
    try {
        if ($ownerId -le 0) {
            $self = [System.Diagnostics.Process]::GetCurrentProcess()
            $ownerId = $self.Id
            $ticks = getOfficeStartTicks $self
            $ownerStartTicks = $(if ($null -ne $ticks) { $ticks } else { 0 })
        }
        [void][System.IO.Directory]::CreateDirectory($dir)
        [System.IO.File]::WriteAllText($tmp, "$processName`t$startTicks`t$ownerId`t$ownerStartTicks", (New-Object System.Text.UTF8Encoding($false)))
        $final = Join-Path $dir "$id.txt"
        [System.IO.File]::Delete($final)
        [System.IO.File]::Move($tmp, $final)
        return $true
    } catch {
        try { [System.IO.File]::Delete($tmp) } catch {}
        return $false
    }
}

function removeOfficeRecord {
    # PID の記録を消す。ファイルが無い・消せないときも失敗にしない（消せなければ次の起動の確認で、PID が無いものとして消える）
    param (
        [string]$dir,
        [int]$id
    )

    if (!$dir) {
        return
    }
    try {
        [System.IO.File]::Delete((Join-Path $dir "$id.txt"))
    } catch {}
}

function readOfficeRecords {
    # 記録を読み、今のプロセスと照らし合わせて @{ Id; ProcessName; StartTicks; OwnerId; OwnerStartTicks; Owned } の配列を返す。
    # Owned: PID があり、名前が同じで、起動時刻が読めて同じ（記録を書いたときのプロセスのまま）。
    # 次のものは、PID が別のプロセスのものになった・無くなった記録なので消す: PID のプロセスが無い・名前が違う・起動時刻が読めて違う。
    # 起動時刻が読めないときは Owned にせず、消しもしない。形の違う記録は使わず、消さない。古い .tmp（書きかけの残り）は消す。
    # フォルダが無い・読めないときは空の配列
    param (
        [string]$dir
    )

    $result = New-Object System.Collections.Generic.List[object]
    if (!$dir) {
        return $result.ToArray()
    }
    try {
        foreach ($tmp in [System.IO.Directory]::GetFiles($dir, "*.tmp")) {
            try {
                if ([System.IO.Path]::GetExtension($tmp) -eq ".tmp" -and
                    ([datetime]::UtcNow - [System.IO.File]::GetLastWriteTimeUtc($tmp)) -gt ${officeRecordTmpMaxAge}) {
                    [System.IO.File]::Delete($tmp)
                }
            } catch {}
        }
        $files = [System.IO.Directory]::GetFiles($dir, "*.txt")
    } catch {
        return $result.ToArray()
    }
    foreach ($file in $files) {
        $id = 0
        $startTicks = [long]0
        $ownerId = 0
        $ownerTicks = [long]0
        try {
            if ([System.IO.Path]::GetExtension($file) -ne ".txt") { continue }
            $fields = ([System.IO.File]::ReadAllText($file).Trim() -split "`t")
            if (!([int]::TryParse([System.IO.Path]::GetFileNameWithoutExtension($file), [ref]$id) -and $id -gt 0 -and $fields.Count -eq 4 -and
                    ${officeProcessNames}.Contains($fields[0]) -and [long]::TryParse($fields[1], [ref]$startTicks) -and $startTicks -gt 0 -and
                    [int]::TryParse($fields[2], [ref]$ownerId) -and [long]::TryParse($fields[3], [ref]$ownerTicks))) {
                continue
            }
        } catch {
            continue
        }
        $process = Get-Process -Id $id -ErrorAction SilentlyContinue
        $owned = $false
        if ($null -eq $process -or $process.ProcessName -ne $fields[0]) {
            removeOfficeRecord $dir $id
            continue
        }
        $actual = getOfficeStartTicks $process
        if ($null -ne $actual) {
            if ($actual -ne $startTicks) {
                removeOfficeRecord $dir $id
                continue
            }
            $owned = $true
        }
        $result.Add([pscustomobject]@{
            Id = $id; ProcessName = $fields[0]; StartTicks = $startTicks; OwnerId = $ownerId; OwnerStartTicks = $ownerTicks; Owned = $owned
        })
    }
    return $result.ToArray()
}

function getOfficeRecordOwner {
    # 記録を書いたプロセス（持ち主）の今の状態: Self（この画面のプロセス）・Other（生きているほかのプロセス）・Gone（もういない）。
    # 持ち主の起動時刻が読めない・記録に無いときは、生きていれば安全な側（Other）にする
    param (
        $record,
        [int]$selfId,
        [long]$selfStartTicks
    )

    if ($record.OwnerId -eq $selfId -and $record.OwnerStartTicks -eq $selfStartTicks) {
        return "Self"
    }
    $owner = Get-Process -Id $record.OwnerId -ErrorAction SilentlyContinue
    if ($null -eq $owner) {
        return "Gone"
    }
    $actual = getOfficeStartTicks $owner
    if ($null -ne $actual -and $record.OwnerStartTicks -gt 0 -and $actual -ne $record.OwnerStartTicks) {
        return "Gone"
    }
    return "Other"
}

# ----------------------------------------------------------------------------
# 一覧と強制終了
# ----------------------------------------------------------------------------

function getOfficeProcesses {
    # 自分のセッションで実行中の Excel・Word・PowerPoint を返す。各要素:
    #   Id・ProcessName・AppName・StartTime（読めなければ $null）・
    #   HasWindow（MainWindowHandle が 0 でない）・
    #   Owned（recordDir の記録があり、記録と今のプロセスの PID・名前・起動時刻が合う。recordDir が空なら常に $false）・
    #   Owner（Owned のとき、記録を書いたプロセスの状態 Self・Other・Gone。Owned でなければ空）
    # 残り物（止めてよいもの）の選び出しは判断層（getLeftoverPrompt）が行う。ここは材料を返すだけ。
    # recordDir を読むため、画面のスレッドでは呼ばない（startJob の中で呼ぶ）。
    # selfId・selfStartTicks は、この画面のプロセスの PID と起動時刻（省けば今のプロセス。テストで差し替える）
    param (
        [string]$recordDir = "",
        [int]$selfId = 0,
        [long]$selfStartTicks = 0
    )

    $current = [System.Diagnostics.Process]::GetCurrentProcess()
    if ($selfId -le 0) {
        $selfId = $current.Id
        $ticks = getOfficeStartTicks $current
        $selfStartTicks = $(if ($null -ne $ticks) { $ticks } else { 0 })
    }
    $records = @{}
    foreach ($record in @(readOfficeRecords $recordDir)) {
        $records[[int]$record.Id] = $record
    }

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($process in @(Get-Process -Name @(${officeProcessNames}.Keys) -ErrorAction SilentlyContinue |
            Where-Object { $_.SessionId -eq $current.SessionId })) {
        $startTime = $null
        try {
            $startTime = $process.StartTime
        } catch {
            # 権限の無いプロセスは起動時刻を取得できない
        }
        $record = $records[[int]$process.Id]
        $owned = ($null -ne $record -and $record.Owned -and $record.ProcessName -eq $process.ProcessName)
        $result.Add([pscustomobject]@{
            Id          = $process.Id
            ProcessName = $process.ProcessName
            AppName     = ${officeProcessNames}[$process.ProcessName]
            StartTime   = $startTime
            HasWindow   = ($process.MainWindowHandle -ne [IntPtr]::Zero)
            Owned       = $owned
            Owner       = $(if ($owned) { getOfficeRecordOwner $record $selfId $selfStartTicks } else { "" })
        })
    }
    return $result.ToArray()
}

function stopOfficeProcesses {
    # 指定したプロセスを保存せずに終了し、@{ Id; Stopped; Status; Reason } の配列を返す。
    # targets は getOfficeProcesses の要素（Id・ProcessName・StartTime）。止める直前に、PID が今も同じプロセスか（名前・起動時刻）を照らし直す。
    #   Status: Stopped（止めた）・Changed（名前か起動時刻が違う・読めない・Office の名前でない。止めない）・
    #           Gone（もう無い。止めた扱いにも失敗にもしない）・Failed（止めようとして失敗）
    #   Reason: Changed のときは「PID <N> は確認の後に別のプロセスに変わったため、終了しませんでした。」、Failed のときは例外のメッセージ。それ以外は空
    # recordDir があれば、止めた・もう無いものの記録を消す。プロセスの強制終了は、ここの Stop-Process の 1 か所だけ（tests/meta/safety.Tests.ps1）
    param (
        [object[]]$targets,
        [string]$recordDir = ""
    )

    $results = New-Object System.Collections.Generic.List[object]
    foreach ($target in $targets) {
        $id = [int]$target.Id
        $process = Get-Process -Id $id -ErrorAction SilentlyContinue
        if ($null -eq $process) {
            removeOfficeRecord $recordDir $id
            $results.Add(@{ Id = $id; Stopped = $false; Status = "Gone"; Reason = "" })
            continue
        }
        $expected = $null
        try { $expected = [long]$target.StartTime.ToUniversalTime().Ticks } catch {}
        $actual = getOfficeStartTicks $process
        if (!${officeProcessNames}.Contains([string]$target.ProcessName) -or $process.ProcessName -ne $target.ProcessName -or
            $null -eq $expected -or $null -eq $actual -or $expected -ne $actual) {
            $results.Add(@{ Id = $id; Stopped = $false; Status = "Changed"; Reason = "PID $id は確認の後に別のプロセスに変わったため、終了しませんでした。" })
            continue
        }
        try {
            Stop-Process -Id $id -Force -ErrorAction Stop
            removeOfficeRecord $recordDir $id
            $results.Add(@{ Id = $id; Stopped = $true; Status = "Stopped"; Reason = "" })
        } catch {
            $results.Add(@{ Id = $id; Stopped = $false; Status = "Failed"; Reason = $_.Exception.Message })
        }
    }
    return $results.ToArray()
}
