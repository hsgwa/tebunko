# Officeアプリ（Excel・Word・PowerPoint）の起動・終了と、1ファイルの制限時間の監視。
# COM を使う処理だけをまとめる。使う側が dot-source する（shared.ps1 では読み込まない）。

# ----------------------------------------------------------------------------
# Officeアプリ（Excel・Word・PowerPoint）の起動・終了
# ----------------------------------------------------------------------------

# 起動中のアプリ: 名前 → @{ Com; Pid; Shared }
$script:apps = @{}

# ExitWait: Quit の後、終わるのを待つ時間（ミリ秒）。過ぎたら強制終了する（quitApp）。
#   Excel は抽出で取り出したCOMオブジェクトが解放されきらないため、インデクサが動いている間は Quit しても終わらず、
#   待ちの上限まで待ってから強制終了していた（実測: 毎回 5 秒待って強制終了）。待つだけ無駄なため短くする。
#   GC で解放を促して自分で終わらせる方法は、インデクサの終了が COM の解放待ちで約 60 秒止まることがあった（実測 25 回中 1〜2 回）ため採らない
# 利用者に渡すときの、アプリごとの違い（handOverForeignApp・restoreHandedOverApp）:
#   Items      : 開いているファイルの一覧を持つ、アプリのプロパティの名前（Workbooks・Documents・Presentations）
#   Close      : 自分のファイルを保存せずに閉じる処理（引数は 1 つのファイル）
#   Restore    : 渡すときに戻す設定。getApp が変える設定は、ここに全部ある（テストで確かめる）。Required は戻らないと渡し切れない設定、
#                Optional は戻ればよい設定、ShowWindow は窓を出すか（Visible を真にする行は restoreHandedOverApp の中の 1 か所だけ）
#   WindowCheck: ファイルの一覧を読めないとき、見える窓があるかで利用者のものかを判断するか
#                （Word は自分で窓を隠しているため窓では判断できない。読めなければ渡しも終了もせず、持ち続けて後で読み直す）
$appInfo = @{
    Excel      = @{
        ProgId = "Excel.Application"; Process = "EXCEL"; ExitWait = 1000
        Items = "Workbooks"; Close = { param ($item) $item.Close($false) }
        Restore = @{
            Required = @(@("DisplayAlerts", $true), @("UserControl", $true))
            Optional = @(@("EnableEvents", $true), @("ScreenUpdating", $true), @("AskToUpdateLinks", $true), @("AutomationSecurity", 1))
            ShowWindow = $true
        }
        WindowCheck = $true
    }
    Word       = @{
        ProgId = "Word.Application"; Process = "WINWORD"; ExitWait = 5000
        Items = "Documents"; Close = { param ($item) $item.Close(0) }  # wdDoNotSaveChanges
        Restore = @{
            Required = @(,@("DisplayAlerts", -1))  # wdAlertsAll
            Optional = @(,@("AutomationSecurity", 1))
            ShowWindow = $true
        }
        WindowCheck = $false
    }
    # SingleInstance: 1つのセッションに1つのプロセスしか持てないアプリ（PowerPointだけ）。
    # 既に起動している（利用者が開いている）ときは、接続せずに使わない（下の getApp）
    PowerPoint = @{
        ProgId = "PowerPoint.Application"; Process = "POWERPNT"; ExitWait = 5000; SingleInstance = $true
        Items = "Presentations"; Close = { param ($item) $item.Close() }
        Restore = @{
            Required = @(,@("DisplayAlerts", 2))  # ppAlertsAll
            Optional = @(,@("AutomationSecurity", 1))
            ShowWindow = $false
        }
        WindowCheck = $true
    }
}

# SingleInstance のアプリが、既に自分のセッションで起動している（利用者が使用中の）ときに投げる例外の文言。
# 呼ぶ側は "<アプリ名>${officeAppInUseMessage}" の形で使う。Reroute（extract_office.ps1 の officeRequiredMessage・
# OperationCanceledException）と取り違えないよう、型（InvalidOperationException）でも区別する
${officeAppInUseMessage} = " が起動しているため、更新に使用できません"

# 起動した Office のプロセスの優先度は下げない（Normal のまま）。利用者がダブルクリックしたファイルがインデックス作成の Excel・Word で開くことがあり、
# PowerPoint は 1 つのプロセスしか持てないため、利用者とプロセスを共有しないと確実には言えない。利用者の操作を遅くしないよう、
# 優先度を下げるのは、利用者と共有しないインデックス作成のスレッドだけにする（docs/design/structure/threads.md「スレッドの一覧」）
# 起動したアプリの PID を入れる入れ物（ConcurrentDictionary[int,string]。$null なら入れない）。
# 画面が閉じるときに、インデックス作成が起動した Office を PID で止めるために使う
$script:officePidSink = $null
# 起動した Office の PID の記録を置くフォルダ（<ワークスペース>\office_pids\<PC の鍵>。$null なら書かない）。
# shared はツールを知らないため、使う側（インデックス作成）が場所を決めて入れる（office_process.ps1 の addOfficeRecord・removeOfficeRecord）。
# 記録は、次に画面を起動したときに、残った Office を確認して止めるために使う（書けなくても取り込みは続ける。その Office は止める対象にならない）
$script:officeRecordDir = $null
# 自分が開くファイルの置き場（インデックス作成の作業領域の一時フォルダ。$null なら、開いているファイルをすべて利用者のものとみなす）。
# 使う側が入れる。置き場の外のファイルが Excel・Word・PowerPoint に開かれたら、利用者が開いたファイルとして扱う（getForeignWorkbookCount・handOverApp）
$script:officeOwnDir = $null
# Excel・Word・PowerPoint を利用者に渡そうとしたときに呼ぶ処理（スクリプトブロック。引数はアプリ名と、渡し切れたか（戻しきれたか。$null は、渡すか決められず持ち続ける）。呼ぶ変数が $null なら何もしない）。ログを書く使う側が入れる
$script:onOfficeHandOver = $null
# 渡そうとして戻しきれなかったアプリ（COM の参照を放さずに持ち続ける）
$script:officeKeptApps = @()
# 渡し切れずに持ち続けたアプリを仕上げ直して待つ、取り込みのスレッドの終わりの上限（秒）と間隔（ミリ秒）
$script:officeKeptWaitSeconds = 30
$script:officeKeptWaitStep = 500

function resolveLongName {
    # 8.3 の短い名前（TEST~1 など）を含むパスを長い名前にする。読めなければ $null
    param ([string]$path)

    try {
        return (Get-Item -LiteralPath $path -ErrorAction Stop).FullName
    } catch {
        return $null
    }
}

function resolveOwnDir {
    # 自分が開くブックの置き場（一時フォルダ）を、比べる形にする。短い名前を長くできなければ、そのまま使う。
    # 置き場が決まっていなければ $null（そのときは開いたブックをすべて利用者のものと数える）
    param ([string]$dir)

    if (-not $dir) {
        return $null
    }
    $long = resolveLongName $dir
    return $(if ($long) { $long } else { $dir })
}

function isUnderDir {
    # path が dir の下にあるか。両方を正規化し、dir は区切りで終わる形にして、大文字小文字を区別せずに前置で比べる
    # （tmp と tmp2 を取り違えない）。正規化できない（URL など）ときは、下に無いものとする
    param ([string]$path, [string]$dir)

    try {
        $full = [System.IO.Path]::GetFullPath($path)
        $base = [System.IO.Path]::GetFullPath($dir).TrimEnd('\') + '\'
        return $full.StartsWith($base, [System.StringComparison]::OrdinalIgnoreCase)
    } catch {
        return $false
    }
}

function getWorkbookSplit {
    # アプリで開いているファイル（Excel のブック・Word の文書・PowerPoint のプレゼンテーション）を、自分が開いたもの（置き場の下）と
    # 利用者が開いたもの（外）に分ける。@{ Own; Foreign }
    # 開いたファイルの FullName を控えて照合する方法は、失敗の途中で Open が戻らなかったファイルを控えられないため採らない。
    # 短い名前（~ を含む）は長い名前にしてから比べる。読めなければ、置き場の下と確かめられないので利用者のものと数える
    # （データを失わない側。自分のものは置き場の下にあり、読める）
    param ($com, [string]$name)

    $own = New-Object System.Collections.ArrayList
    $foreign = New-Object System.Collections.ArrayList
    $workbooks = $com.($appInfo[$name].Items)
    try {
        foreach ($book in @($workbooks)) {
            if ($null -eq $book) { continue }
            $full = [string]$book.FullName
            if ($full.Contains('~')) {
                $full = resolveLongName $full
            }
            if ($full -and $script:officeOwnDir -and (isUnderDir $full $script:officeOwnDir)) {
                [void]$own.Add($book)
            } else {
                [void]$foreign.Add($book)
            }
        }
    } finally {
        try { releaseComObject $workbooks } catch {}
    }
    return @{ Own = $own; Foreign = $foreign }
}

function releaseWorkbookSplit {
    # getWorkbookSplit で取り出したファイルの参照を放す（残すとアプリが終わらなくなる）
    param ($split)

    foreach ($book in @($split.Own) + @($split.Foreign)) { try { releaseComObject $book } catch {} }
}

function getForeignWorkbookCount {
    # アプリで利用者が開いたファイル（置き場の外）の数。読めない（COM の例外）ときは -1（分からない。handOverForeignApp が判断する）
    param ($com, [string]$name)

    try {
        $split = getWorkbookSplit $com $name
        $count = @($split.Foreign).Count
        # 取り出したファイルの参照を残すとアプリが終わらなくなるため、数えたら放す
        releaseWorkbookSplit $split
        return $count
    } catch {
        return -1
    }
}

function testProcessHasWindow {
    # 起動で控えた PID のプロセスに、見える窓があるか。PID だけを見る（窓を中身で探さない）。控えていなければ（0）偽
    param ([int]$processId, [string]$processName)

    if ($processId -le 0) {
        return $false
    }
    $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
    return [bool]($process -and $process.ProcessName -eq $processName -and $process.MainWindowHandle -ne 0)
}

function testProcessExists {
    # 指定した PID に、指定した名前のプロセスがあるか。PID と名前だけを見る
    param ([int]$processId, [string]$processName)

    $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
    return [bool]($process -and $process.ProcessName -eq $processName)
}

function getMonotonicMilliseconds {
    # 経過時間の計測用の、戻らない時計（ミリ秒）
    return [long]([System.Diagnostics.Stopwatch]::GetTimestamp() * 1000 / [System.Diagnostics.Stopwatch]::Frequency)
}

function getOwnSessionProcessIds {
    # 自分のセッションで動いている、指定した名前のプロセスのIDの一覧。
    # ほかのセッション（同じPCの別の利用者・別の作業フォルダの tebunko 等）のプロセスは、自分のものと取り違えないため数えない
    param (
        [string]$processName,
        [int]$sessionId
    )

    return @(Get-Process -Name $processName -ErrorAction SilentlyContinue |
        Where-Object { $_.SessionId -eq $sessionId } | ForEach-Object { $_.Id })
}

function getApp {
    # アプリのCOMオブジェクトを返す。起動していなければ起動する。
    # SingleInstance のアプリ（PowerPoint）が既に自分のセッションで起動しているときは、新しいプロセスができず
    # 利用者のアプリに接続することになるため、接続せずに例外を投げる（呼ぶ側は後回しにする）
    param (
        [string]$name
    )

    if (-not $script:apps.ContainsKey($name)) {
        $info = $appInfo[$name]
        $sessionId = [System.Diagnostics.Process]::GetCurrentProcess().SessionId

        if ($info.SingleInstance -and (getOwnSessionProcessIds $info.Process $sessionId).Count -gt 0) {
            throw (New-Object System.InvalidOperationException "${name}${officeAppInUseMessage}")
        }

        # 終了できなかった場合に強制終了するため、新しく起動したプロセスのIDを控えておく
        $before = getOwnSessionProcessIds $info.Process $sessionId
        $com = New-Object -ComObject $info.ProgId
        $after = getOwnSessionProcessIds $info.Process $sessionId
        $newIds = @($after | Where-Object { $before -notcontains $_ })

        if ($info.SingleInstance -and $newIds.Count -eq 0) {
            # 事前の確認から New-Object の間に、利用者が先にアプリを起動した（先に接続した）。
            # DisplayAlerts・AutomationSecurity などの設定を変える前に解放し、利用者のアプリに触らない
            try { releaseComObject $com } catch {}
            throw (New-Object System.InvalidOperationException "${name}${officeAppInUseMessage}")
        }
        if ($newIds.Count -eq 1 -and $script:officePidSink) {
            $script:officePidSink[[int]$newIds[0]] = $info.Process
        }
        if ($newIds.Count -eq 1 -and $script:officeRecordDir) {
            # 起動時刻が読めたときだけ記録する（読めなければ、残っても止める対象にならない）
            $started = Get-Process -Id ([int]$newIds[0]) -ErrorAction SilentlyContinue
            $startTicks = $(if ($started -and $started.ProcessName -eq $info.Process) { getOfficeStartTicks $started } else { $null })
            if ($null -ne $startTicks) {
                [void](addOfficeRecord $script:officeRecordDir ([int]$newIds[0]) $info.Process $startTicks)
            }
        }

        switch ($name) {
            "Excel" {
                $com.Visible = $false
                $com.DisplayAlerts = $false
                $com.EnableEvents = $false
                $com.ScreenUpdating = $false
                $com.AskToUpdateLinks = $false
            }
            "Word" {
                $com.Visible = $false
                $com.DisplayAlerts = 0  # wdAlertsNone
            }
            "PowerPoint" {
                # PowerPointはウィンドウを隠せないため、ファイルをウィンドウ無しで開く（Visible は変更しない）
                $com.DisplayAlerts = 1  # ppAlertsNone
            }
        }
        $com.AutomationSecurity = 3  # msoAutomationSecurityForceDisable（マクロ無効）

        # Excel・Word は起動中のアプリに接続することがある（新しいプロセスができない場合。PowerPoint は上で防いでいるため
        # ここには来ない）。その場合は利用者のものなので終了させない
        $script:apps[$name] = @{
            Com    = $com
            Pid    = $(if ($newIds.Count -eq 1) { $newIds[0] } else { 0 })
            Shared = ($newIds.Count -eq 0)
        }
        updateWatchedPids
    }
    return $script:apps[$name].Com
}

function restoreHandedOverApp {
    # 渡すアプリに対して、(0) 自分が開いたファイル（置き場の下と確かめられたもの）を保存せずに閉じる (1) 必須の設定（$appInfo の Restore.Required と、窓を出すアプリは Visible）
    # (2) 任意の設定を戻す。state に済んだものを控えるので、何度呼んでも済んだものはやり直さない。全部済んだら $true
    # tries: 必須の設定をやり直す回数（失敗したときだけ間を置く）。利用者のファイルは閉じない
    param ($state, [int]$tries = 1)

    $com = $state.Com
    $info = $appInfo[$state.Name]
    if (-not $state.BooksClosed) {
        try {
            $books = getWorkbookSplit $com $state.Name
            # 前の回の一覧の参照は、取り直した一覧に置き換える前に放す
            if ($state.Books) {
                releaseWorkbookSplit $state.Books
            }
            $closeFailed = $false
            foreach ($book in $books.Own) {
                try { & $info.Close $book } catch { $closeFailed = $true }
            }
            $state.Books = $books
            # 閉じられなかったものが 1 つでもあれば、立てない（次回に一覧を取り直す。閉じたものは一覧から消えるので二重には閉じない）
            $state.BooksClosed = (-not $closeFailed)
        } catch {}
    }
    # 窓・UserControl・DisplayAlerts が戻らないまま参照を放すと、Excel が保存の確認なしに終わりうるため、成功を確かめる
    $required = @($info.Restore.Required)
    if ($info.Restore.ShowWindow) { $required += , @("Visible", $true) }
    $optional = @($info.Restore.Optional)
    for ($try = 0; $try -lt $tries; $try++) {
        $left = @($required | Where-Object { -not $state.Done[$_[0]] })
        if ($left.Count -eq 0) { break }
        if ($try -gt 0) { Start-Sleep -Milliseconds 200 }
        foreach ($setting in $left) {
            try { $com.($setting[0]) = $setting[1]; $state.Done[$setting[0]] = $true } catch {}
        }
    }
    foreach ($setting in @($optional | Where-Object { -not $state.Done[$_[0]] })) {
        try { $com.($setting[0]) = $setting[1]; $state.Done[$setting[0]] = $true } catch {}
    }
    return [bool]($state.BooksClosed -and (@($required + $optional | Where-Object { -not $state.Done[$_[0]] }).Count -eq 0))
}

function releaseHandedOverApp {
    # このスレッドが持つ COM を解放しきる。WaitForPendingFinalizers は長く止まることがあるため使わない
    param ($state)

    try {
        if ($state.Books) {
            releaseWorkbookSplit $state.Books
        }
        releaseComObject $state.Com
        [GC]::Collect()
    } catch {}
}

function untrackApp {
    # アプリを (a) 見張りの対象から外す (b) 一括終了の対象から外す (c) 残り物の確認の記録から消す。誰も止めなくなる。それぞれ、ほかが失敗しても行う
    param ([string]$name, $app)

    try { $script:apps.Remove($name) } catch {}
    try { updateWatchedPids } catch {}
    try {
        if ($app.Pid -and $script:officePidSink) {
            $removed = $null
            [void]$script:officePidSink.TryRemove([int]$app.Pid, [ref]$removed)
        }
    } catch {}
    try {
        if ($app.Pid -and $script:officeRecordDir) {
            removeOfficeRecord $script:officeRecordDir ([int]$app.Pid)
        }
    } catch {}
}

function keepUndecidedApp {
    # ファイルの一覧が読めず、利用者のファイルがあるか分からないアプリ（窓で判断できない Word）を、終了させず・渡さず、持ち続ける。
    # 利用者のファイルがあるかもしれないので、誰も止めない（untrackApp）。retryKeptApps が一覧を読み直し、利用者のファイルがあれば渡し、
    # 自分のファイルだけなら終了させる
    param ([string]$name)

    $app = $script:apps[$name]
    untrackApp $name $app
    $script:officeKeptApps += , @{ Name = $name; Com = $app.Com; Pid = [int]$app.Pid; BooksClosed = $false; Books = $null; Done = @{}; Undecided = $true }
    if ($script:onOfficeHandOver) {
        # 渡したかどうかがまだ決まっていない（$null）ことを知らせる
        try { & $script:onOfficeHandOver $name $null } catch {}
    }
}

function quitApp {
    # アプリを終了させる。Quit し、COM の参照を放し、控えた PID のプロセス（名前も合うもの）が終わるのを待つ。
    # GC::WaitForPendingFinalizers() は COM の解放待ちで長時間（約 60 秒）止まることがあるため使わず、
    # 待ち時間（ExitWait）を過ぎても終わらなければ PID を指定して強制終了する（名前では止めない）。終わったと確かめられたら $true
    param ([string]$name, $com, [int]$processId)

    try { $com.Quit() } catch {}
    try { releaseComObject $com } catch {}
    if ($processId -le 0) {
        return $true
    }
    $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
    if ($null -eq $process -or $process.ProcessName -ne $appInfo[$name].Process) {
        return $true
    }
    $wait = $appInfo[$name].ExitWait
    if ($process.WaitForExit($wait)) {
        return $true
    }
    # 終了処理中のプロセスは Kill() が「アクセス拒否」で失敗することがあるが、そのまま終了するため無視する
    try { $process.Kill() } catch {}
    # 強制終了は非同期のため、同じ上限で終わるのを待つ。それでも残った場合は、次の getApp が
    # 利用者のものとみなして後回しにする（安全な側に倒れる）
    return [bool]$process.WaitForExit($wait)
}

function quitKeptApp {
    # 自分のファイルだけだと分かった、持ち続けたアプリを終了させる（quitApp）
    param ($state)

    [void](quitApp $state.Name $state.Com $state.Pid)
}

function handOverApp {
    # 起動したアプリ（Excel・Word・PowerPoint）を、利用者に渡す。Quit も強制終了もしない（利用者のファイルを閉じない）。渡したアプリはもう使わない。
    # 見張り・終了時の一括終了・次の起動の残り物の確認のどれからも外すので、誰も止めない。
    # Shared のアプリ（利用者のアプリに接続したもの）には当てない（止めず、設定も変えない）
    param ([string]$name)

    $app = $script:apps[$name]
    if ($null -eq $app -or $app.Shared) {
        return $false
    }
    untrackApp $name $app

    # 自分のファイルを閉じる・設定を戻す・窓を出す、を全部通して初めて「渡した」。通らなければ、持ち続けて後で設定し直す
    $state = @{ Name = $name; Com = $app.Com; Pid = [int]$app.Pid; BooksClosed = $false; Books = $null; Done = @{} }
    $ok = restoreHandedOverApp $state 3
    if ($ok) {
        releaseHandedOverApp $state
    } else {
        # 戻しきれなかったアプリは、参照を放すと終わるおそれがあるため、放さずに持ち続ける（止めも強制終了もしない）。
        # retryKeptApps と waitKeptApps が設定し直す。取り込みのスレッドが終わるまでに通らなければ、参照が切れる
        $script:officeKeptApps += , $state
    }

    if ($script:onOfficeHandOver) {
        try { & $script:onOfficeHandOver $name $ok } catch {}
    }
    return $ok
}

function retryKeptApps {
    # 渡し切れずに持ち続けているアプリをもう一度仕上げる（済んでいないものだけ）。全部通れば参照を放して「渡した」ログを書く。待たない。
    # 控えた PID のプロセスが（その名前で）もう無ければ、利用者が先に終えたので、参照だけ放して一覧から外す（PID と名前だけを見る。窓は探さない。止めない）
    if (@($script:officeKeptApps).Count -eq 0) {
        return
    }
    $remaining = @()
    foreach ($kept in @($script:officeKeptApps)) {
        if ($kept.Pid -gt 0 -and -not (testProcessExists ([int]$kept.Pid) $appInfo[$kept.Name].Process)) {
            releaseHandedOverApp $kept
            continue
        }
        if ($kept.Undecided) {
            # 利用者のファイルがあるか、一覧を読み直す。まだ読めなければ、そのまま持ち続ける
            $foreignCount = getForeignWorkbookCount $kept.Com $kept.Name
            if ($foreignCount -lt 0) {
                $remaining += , $kept
                continue
            }
            if ($foreignCount -eq 0) {
                quitKeptApp $kept
                continue
            }
            $kept.Undecided = $false
        }
        if (restoreHandedOverApp $kept 1) {
            releaseHandedOverApp $kept
            if ($script:onOfficeHandOver) {
                try { & $script:onOfficeHandOver $kept.Name $true } catch {}
            }
        } else {
            $remaining += , $kept
        }
    }
    $script:officeKeptApps = $remaining
}

function waitKeptApps {
    # 取り込みのスレッドの終わりに限り、持ち続けているアプリを、間を置いて上限まで仕上げ直す。持ち続けが無ければ待たない。
    # shouldStop が真を返したら（中止・画面を閉じる）次の刻みで抜ける。上限に届かなかったものは、そのまま持ち続け（スレッドが終わると参照が切れる）
    param ([scriptblock]$shouldStop = { $false })

    $started = getMonotonicMilliseconds
    while (@($script:officeKeptApps).Count -gt 0) {
        retryKeptApps
        if (@($script:officeKeptApps).Count -eq 0 -or ((getMonotonicMilliseconds) - $started) -ge ($script:officeKeptWaitSeconds * 1000) -or (& $shouldStop)) {
            break
        }
        Start-Sleep -Milliseconds $script:officeKeptWaitStep
    }
}

function handOverForeignApp {
    # 起動したアプリ（Excel・Word・PowerPoint）に利用者が開いたファイルがあれば、利用者に渡して $true を返す。無ければ何もせず $false
    param ([string]$name)

    # 前に戻しきれなかったアプリがあれば、ここで仕上げ直す
    retryKeptApps
    $app = $script:apps[$name]
    if ($null -eq $app -or $app.Shared) {
        return $false
    }
    $count = getForeignWorkbookCount $app.Com $name
    if ($count -eq 0) {
        return $false
    }
    # ファイルの一覧を読めない（COM が呼び出しを拒んだ。利用者が操作中のことがある）ときは、
    #   窓で判断できるアプリ（Excel・PowerPoint）: 見える窓があれば利用者が使っているとみなして渡し、無ければ何もしない（今までどおり終了させる）
    #   窓で判断できないアプリ（Word。自分で窓を隠している）: 渡しも終了もさせず、持ち続けて後で一覧を読み直す（空の窓を出さず、データも失わない）
    if ($count -lt 0) {
        if ($appInfo[$name].WindowCheck) {
            if (-not (testProcessHasWindow ([int]$app.Pid) $appInfo[$name].Process)) { return $false }
        } else {
            keepUndecidedApp $name
            return $true
        }
    }
    [void](handOverApp $name)
    return $true
}

function stopApp {
    param (
        [string]$name
    )

    $app = $script:apps[$name]
    if ($null -eq $app) {
        return
    }
    # インデックス作成中に利用者が同じアプリでファイルを開いた場合は、閉じずに利用者に渡す（自分のファイルだけなら、下で終了させる）
    if (handOverForeignApp $name) {
        return
    }
    $script:apps.Remove($name)
    updateWatchedPids

    # 利用者のアプリに接続したもの（Shared）は、終了させず、参照だけ放す。利用者に渡したものとして、記録も消す
    if ($app.Shared) {
        try { releaseComObject $app.Com } catch {}
        $exited = $true
    } else {
        $exited = quitApp $name $app.Com $app.Pid
    }
    # 記録は、プロセスが終わったと確かめてから消す。終わらなければ残す（次の起動の確認に任せる）
    if ($app.Pid -and $script:officeRecordDir -and $exited) {
        removeOfficeRecord $script:officeRecordDir ([int]$app.Pid)
    }
    if ($app.Pid -and $script:officePidSink) {
        $removed = $null
        [void]$script:officePidSink.TryRemove([int]$app.Pid, [ref]$removed)
    }
}

function stopAllApps {
    retryKeptApps
    # 1つのアプリの終了に失敗しても、残りのアプリは終了させる
    foreach ($name in @($script:apps.Keys)) {
        try {
            stopApp $name
        } catch {
            Write-Host "    ${name} の終了に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

function getAppName {
    # 拡張子から、抽出に使うアプリの名前を返す
    param (
        [string]$path
    )

    switch -Regex ([System.IO.Path]::GetExtension($path).ToLower()) {
        "^\.xls" { return "Excel" }
        "^\.doc" { return "Word" }
        "^\.ppt" { return "PowerPoint" }
    }
    return $null
}

function releaseComObject($object) {
    if ($null -ne $object) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($object)
    }
}

# ----------------------------------------------------------------------------
# 1ファイルの制限時間の監視
# ----------------------------------------------------------------------------

# 取り込み中のCOM呼び出しは応答が無いと戻らず、Ctrl+C も効かないため、別スレッドで制限時間を監視する。
# 制限時間を過ぎたら、自分で起動したOfficeアプリを強制終了する（COM呼び出しが例外で戻り、そのファイルは失敗になる）。
#   Deadline: 取り込み中のファイルの制限時刻（取り込み中でなければ MaxValue）
#   Pids    : 強制終了してよいプロセスID（自分で起動したOfficeアプリ）
#   TimedOut: 制限時間を過ぎて強制終了した
$script:watchdog = [hashtable]::Synchronized(@{ Deadline = [datetime]::MaxValue; Pids = @(); TimedOut = $false; Stop = $false })
$script:watchdogThread = $null

function updateWatchedPids {
    # 強制終了してよいプロセスIDを、起動中のアプリのうち自分で起動したものにする（利用者のアプリは終了させない）
    $script:watchdog.Pids = @($script:apps.Values | Where-Object { -not $_.Shared -and $_.Pid } | ForEach-Object { $_.Pid })
}

function startWatchdog {
    $ps = [PowerShell]::Create()
    [void]$ps.AddScript({
        param($watch, [string[]]$processNames)
        while (-not $watch.Stop) {
            Start-Sleep -Milliseconds 500
            if ([datetime]::Now -lt $watch.Deadline) {
                continue
            }
            $watch.Deadline = [datetime]::MaxValue
            $watch.TimedOut = $true
            foreach ($id in @($watch.Pids)) {
                try {
                    # 終了済みでIDが別のプロセスに再利用されている場合に備え、Officeアプリであることを確かめる
                    $process = [System.Diagnostics.Process]::GetProcessById($id)
                    if ($processNames -contains $process.ProcessName) {
                        $process.Kill()
                    }
                } catch {}
            }
        }
    }).AddArgument($script:watchdog).AddArgument([string[]]@($appInfo.Values | ForEach-Object { $_.Process }))
    $script:watchdogThread = @{ PowerShell = $ps; Handle = $ps.BeginInvoke() }
}

function stopWatchdog {
    if ($null -eq $script:watchdogThread) {
        return
    }
    $script:watchdog.Stop = $true
    try { [void]$script:watchdogThread.PowerShell.EndInvoke($script:watchdogThread.Handle) } catch {}
    $script:watchdogThread.PowerShell.Dispose()
    $script:watchdogThread = $null
}
