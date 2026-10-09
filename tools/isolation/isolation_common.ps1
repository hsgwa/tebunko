# 実機の確かめ・テストが、利用者の既定のワークスペースに触れないようにするための共通の関数。
# tools\run_isolated.ps1・tests\run.ps1・tests\gui\gui_helpers.ps1・tools\capture_screens.ps1 が読み込む。
# 既定のワークスペースは、環境変数 TEBUNKO_DEFAULT_WORKSPACE で差し替える（scripts\tebunko\core\settings.ps1 の getDefaultWorkDir）。
# このファイルは、環境変数を通さずに既定のワークスペースの場所を求める（利用者の既定のワークスペースかどうかを調べるため）。

function getRealDefaultWorkspace {
    # 利用者の既定のワークスペース。環境変数 TEBUNKO_DEFAULT_WORKSPACE を通さずに求める
    return Join-Path ([System.Environment]::GetFolderPath("UserProfile")) "Documents\tebunko_ws"
}

function testIsolationPathInside {
    # Path が Dir の中（Dir 自身を含む）か。大文字小文字は区別しない
    param ([string]$Path, [string]$Dir)

    if ([string]::IsNullOrWhiteSpace($Path) -or [string]::IsNullOrWhiteSpace($Dir)) { return $false }
    $full = [System.IO.Path]::GetFullPath($Path).TrimEnd("\") + "\"
    $base = [System.IO.Path]::GetFullPath($Dir).TrimEnd("\") + "\"
    return $full.StartsWith($base, [System.StringComparison]::OrdinalIgnoreCase)
}

function assertNotRealWorkspace {
    # 差し替えた既定の場所が、利用者の既定のワークスペースと同じか、その下を指していれば例外にする（既定のワークスペースに書く前に止める）
    param ([string]$Path, [string]$Real = (getRealDefaultWorkspace))

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "既定のワークスペースが決まっていません"
    }
    if ((testIsolationPathInside $Path $Real) -or (testIsolationPathInside $Real $Path)) {
        throw "差し替えた既定の場所が、利用者の既定のワークスペースを指しています: $Path"
    }
}

function newIsolatedFolder {
    # 一時フォルダの下に新しい使い捨てのフォルダを作って、そのパスを返す
    param ([string]$Prefix = "tebunko-isolated")

    $path = Join-Path ([System.IO.Path]::GetTempPath()) ($Prefix + "-" + [guid]::NewGuid().ToString("N").Substring(0, 8))
    [void][System.IO.Directory]::CreateDirectory($path)
    return $path
}

function getIsolationSnapshot {
    # フォルダ（またはファイル）の中身の写し。名前・大きさ・更新時刻（ファイルだけ。フォルダは名前だけ）を控える（中身は読まない）。
    # 返すのは「相対パス → 大きさ:更新時刻」の連想配列。無ければキー "" に "(無い)" を入れる
    param ([string]$Path)

    $map = @{}
    if (!(Test-Path -LiteralPath $Path)) {
        $map[""] = "(無い)"
        return $map
    }
    $item = Get-Item -LiteralPath $Path -Force
    if (!$item.PSIsContainer) {
        $map[""] = "$($item.Length):$($item.LastWriteTimeUtc.Ticks)"
        return $map
    }
    $map[""] = "(フォルダ)"
    $prefix = $item.FullName.TrimEnd("\")
    foreach ($child in @(Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue)) {
        # フォルダの更新時刻は控えない。NTFS は、一覧に出すフォルダの時刻を遅れて更新するため、
        # 中身が同じでも前後で値が違って見える（誤検知になる）。中の増減は、中の名前の違いに出る
        $map[$child.FullName.Substring($prefix.Length)] = if ($child.PSIsContainer) { "(フォルダ)" } else { "$($child.Length):$($child.LastWriteTimeUtc.Ticks)" }
    }
    return $map
}

function formatIsolationValue {
    # 写しの値（"大きさ:更新時刻の ticks"）を、人が読める形（大きさと更新時刻）にする
    param ($Value)

    $text = [string]$Value
    if ($text -match '^(\d*):(\d+)$') {
        $time = (New-Object DateTime ([int64]$Matches[2]), ([DateTimeKind]::Utc)).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss.fffffff")
        return "大きさ $($Matches[1]) 更新 $time"
    }
    return $text
}

function compareIsolationSnapshot {
    # 前後の写しの違いの一覧（無ければ空）。変わったもの・増えたものは更新時刻も出す
    param ($Before, $After, [string]$Label = "")

    $diffs = New-Object System.Collections.Generic.List[string]
    foreach ($key in @($After.Keys | Sort-Object)) {
        if (!$Before.ContainsKey($key)) {
            $diffs.Add("$Label 増えた: $key ($(formatIsolationValue $After[$key]))")
        } elseif ([string]$Before[$key] -ne [string]$After[$key]) {
            $diffs.Add("$Label 変わった: $key ($(formatIsolationValue $Before[$key]) -> $(formatIsolationValue $After[$key]))")
        }
    }
    foreach ($key in @($Before.Keys | Sort-Object)) {
        if (!$After.ContainsKey($key)) {
            $diffs.Add("$Label 消えた: $key")
        }
    }
    return @($diffs)
}

function newIsolationWatch {
    # 見張る場所（表示の名前 → パス）の前の写しを控える。終わりに getIsolationWatchDiffs で比べる
    param ([System.Collections.IDictionary]$Targets)

    $before = [ordered]@{}
    foreach ($label in $Targets.Keys) {
        $before[$label] = getIsolationSnapshot $Targets[$label]
    }
    return @{ Targets = $Targets; Before = $before }
}

function getIsolationWatchDiffs {
    # 見張った場所の、今と控えた写しの違いの一覧（無ければ空）
    param ($Watch)

    $diffs = New-Object System.Collections.Generic.List[string]
    foreach ($label in $Watch.Targets.Keys) {
        foreach ($d in @(compareIsolationSnapshot $Watch.Before[$label] (getIsolationSnapshot $Watch.Targets[$label]) $label)) {
            $diffs.Add($d)
        }
    }
    return @($diffs)
}

function getIsolationReport {
    # 違いの一覧から、終了コードと表示する行を作る（違いが無ければ 0 と空）。run.ps1・capture_screens.ps1（finishWorkspaceGuard 経由）・run_isolated.ps1 が使う。
    # 前後の比べは、このプロセスの外（利用者自身・別の作業ツリーのテスト）が書いた場合も落ちる。落ちたら、まず何が書いたかを疑う
    param ([string[]]$Diffs, [string]$Title = "前後で、利用者の既定のワークスペースなどに違いがありました")

    $diffs = @($Diffs | Where-Object { $_ })
    if ($diffs.Count -eq 0) {
        return @{ ExitCode = 0; Lines = @() }
    }
    $lines = @("${Title}:") + @($diffs | ForEach-Object { "  $_" }) +
        @("（このプロセスの外が書いた場合も、ここに出ます。ほかの作業ツリーのテスト・利用者自身の操作が重なっていなかったか確かめてください）")
    return @{ ExitCode = 1; Lines = $lines }
}

function startWorkspaceGuard {
    # 使い捨ての既定のワークスペースを作って TEBUNKO_DEFAULT_WORKSPACE に入れ（子のプロセスにも引き継がれる）、
    # 既定のワークスペース（Real。テストでは既定のワークスペースの代わりの場所）の写しを控える。
    # 入れた場所が既定のワークスペースの場所と同じか、その下なら、何も入れずに例外にする。終わったら stopWorkspaceGuard を呼ぶ
    param ([string]$Real = (getRealDefaultWorkspace), [string]$Prefix = "tebunko-test-ws")

    $isolated = newIsolatedFolder $Prefix
    try {
        assertNotRealWorkspace $isolated $Real
    } catch {
        Remove-Item -LiteralPath $isolated -Recurse -Force -ErrorAction SilentlyContinue
        throw
    }
    $guard = @{
        Real     = $Real
        Isolated = $isolated
        Previous = $env:TEBUNKO_DEFAULT_WORKSPACE
        Watch    = (newIsolationWatch ([ordered]@{ "既定のワークスペース" = $Real }))
    }
    $env:TEBUNKO_DEFAULT_WORKSPACE = $isolated
    return $guard
}

function stopWorkspaceGuard {
    # 環境変数を元に戻し、使い捨てのフォルダを消して、既定のワークスペースの前後の違いの一覧（無ければ空）を返す
    param ($Guard)

    $env:TEBUNKO_DEFAULT_WORKSPACE = $Guard.Previous
    Remove-Item -LiteralPath $Guard.Isolated -Recurse -Force -ErrorAction SilentlyContinue
    return @(getIsolationWatchDiffs $Guard.Watch)
}

function finishWorkspaceGuard {
    # stopWorkspaceGuard の結果を、終了コードと表示する行（getIsolationReport）にする
    param ($Guard)

    return getIsolationReport @(stopWorkspaceGuard $Guard) "前後で、利用者の既定のワークスペースに違いがありました"
}

function waitIsolationProcess {
    # 起動した画面のプロセスを待つ。Wait 秒（0 は終わるまで）たっても生きていれば、その PID とその子だけを止める（名前では探さない）。
    # TimedOut は、待ち時間のあとも生きていた（止めた）こと。ExitCode は、待ち時間の前に終わったときの終了コード
    param ($Process, [int]$Wait = 0)

    if ($Wait -le 0) {
        $Process.WaitForExit()
        return @{ TimedOut = $false; ExitCode = $Process.ExitCode }
    }
    if ($Process.WaitForExit($Wait * 1000)) {
        return @{ TimedOut = $false; ExitCode = $Process.ExitCode }
    }
    & taskkill.exe /PID $Process.Id /T /F | Out-Null
    [void]$Process.WaitForExit(5000)
    return @{ TimedOut = $true; ExitCode = $null }
}

function getIsolationLaunchVerdict {
    # waitIsolationProcess の結果と待ち秒数から、画面の起動の判定を返す（ExitCode: 0 か 1、Line: 出す文。無ければ空）。
    # 待ち時間のあとも生きていた（TimedOut）なら 0。待ち秒数があり、その前に 0 以外で終わったなら、画面が開かなかったとみなして 1
    param ($Waited, [int]$Wait, [int]$ProcessId)

    if ($Waited.TimedOut) {
        return @{ ExitCode = 0; Line = "$Wait 秒たったので、PID $ProcessId とその子を止めます" }
    }
    if ($Wait -gt 0 -and $Waited.ExitCode -ne 0) {
        return @{ ExitCode = 1; Line = "待ち時間の前に、画面のプロセス (PID $ProcessId) が終了コード $($Waited.ExitCode) で終わりました" }
    }
    return @{ ExitCode = 0; Line = "" }
}
