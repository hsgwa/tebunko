# 処理の記録（どこにどれぐらい時間がかかったかを、区分ごとに入り切りできる形で控えて書く）。
#
# 区切りの場所に markTrace を置き、区間の終わりで writeTrace が控えをファイルへ足す。
# 入れる区分は指定（initTrace に渡す文字列。区分名をカンマで並べる。* で全部。空なら全部切り）で決め、
# 指定の元（環境変数 TEBUNKO_TRACE）は呼び出し側が読む（gui.ps1 の先頭の 1 か所。検索の司令のスレッドへは文字列で渡す）。
# 切りの区分は何もしない。書けなくても例外を出さない（記録のせいで本来の処理を止めない）。
# 控えはスレッド（Runspace）ごと。書くファイルもスレッドごとに分け、同時に書いても混ざらないようにする。
# 経過は initTrace に渡した起点（UTC の Ticks）からのミリ秒。スレッドの間で同じ起点を渡せば、そろって読める（docs/design/testing/perf.md）。

# 2 回読み込まれても控えを消さないよう、無いときだけ作る（gui.ps1 が先に読み込む）
if ($null -eq (Get-Variable -Name traceState -Scope Script -ErrorAction SilentlyContinue)) {
    $script:traceState = @{
        # 入っている区分（* のときは All）
        Categories = @{}
        All = $false
        Spec = ""
        OriginTicks = [DateTime]::UtcNow.Ticks
        # 区分ごとの控え（@(Ticks, 区分, 名前, スレッド) の並び）
        Buffers = @{}
        # 書き終えて、以後は控えを積まない区分
        Closed = @{}
    }
}

function initTrace {
    param (
        [string]$Spec,
        # 経過の起点（UTC の Ticks）。0 なら今
        [long]$OriginTicks = 0
    )

    $state = $script:traceState
    $state.Categories = @{}
    $state.All = $false
    # 控えと -Once の印も作り直す（指定を替えたら、前の控えを持ち越さない）
    $state.Buffers = @{}
    $state.Closed = @{}
    $state.Spec = "$Spec"
    foreach ($name in @("$Spec".Split(","))) {
        $trimmed = $name.Trim()
        if ($trimmed -eq "*") {
            $state.All = $true
        } elseif ($trimmed -ne "") {
            $state.Categories[$trimmed] = $true
        }
    }
    if ($OriginTicks -gt 0) {
        $state.OriginTicks = $OriginTicks
    } else {
        $state.OriginTicks = [DateTime]::UtcNow.Ticks
    }
}

function getTraceSpec {
    # initTrace に渡した指定と起点。ほかのスレッドへ渡して、同じ指定・同じ起点で initTrace させるために使う
    return @{ Spec = $script:traceState.Spec; OriginTicks = $script:traceState.OriginTicks }
}

function testTraceActive {
    # どれかの区分が入っているか
    $state = $script:traceState
    return $state.All -or $state.Categories.Count -gt 0
}

function testTrace {
    param (
        [string]$Category
    )

    $state = $script:traceState
    return $state.All -or $state.Categories.ContainsKey($Category)
}

function markTrace {
    param (
        [string]$Category,
        [string]$Name,
        # 別のスレッドで起きた時刻（UTC の Ticks）を、その時刻で控えるとき。0 なら今
        [long]$AtTicks = 0
    )

    $state = $script:traceState
    if (-not ($state.All -or $state.Categories.ContainsKey($Category))) {
        return
    }
    if ($state.Closed.ContainsKey($Category)) {
        return
    }
    if (-not $state.Buffers.ContainsKey($Category)) {
        $state.Buffers[$Category] = New-Object System.Collections.Generic.List[object]
    }
    $ticks = if ($AtTicks -gt 0) { $AtTicks } else { [DateTime]::UtcNow.Ticks }
    $state.Buffers[$Category].Add(@($ticks, $Category, $Name, [System.Threading.Thread]::CurrentThread.ManagedThreadId))
}

function writeTrace {
    # 区分の控えだけをファイル（<書き出す場所>\<起点の日時>_<PID>_<スレッド>.tsv）の末尾に足し、その区分の控えを空にする。
    # -Once を付けた区分は、以後このプロセスでは控えを積まない（起動のように 1 回だけ書く区分）
    param (
        [string]$Category,
        [string]$TraceDir,
        [switch]$Once
    )

    try {
        $state = $script:traceState
        if ($Once) {
            $state.Closed[$Category] = $true
        }
        if (-not $state.Buffers.ContainsKey($Category)) {
            return
        }
        $items = $state.Buffers[$Category].ToArray()
        $state.Buffers.Remove($Category)
        if ($items.Count -eq 0 -or [string]::IsNullOrEmpty($TraceDir)) {
            return
        }
        if (testTraceNetworkPath $TraceDir) {
            return
        }
        $threadId = [System.Threading.Thread]::CurrentThread.ManagedThreadId
        $origin = [DateTime]::SpecifyKind([DateTime]::new([long]$state.OriginTicks), [DateTimeKind]::Utc)
        $file = Join-Path $TraceDir ("{0}_{1}_{2}.tsv" -f $origin.ToLocalTime().ToString("yyyyMMdd-HHmmss"), $PID, $threadId)
        $builder = New-Object System.Text.StringBuilder
        if (-not [System.IO.File]::Exists($file)) {
            [void]$builder.Append("時刻`t経過ミリ秒`t区分`t名前`tスレッド`r`n")
        }
        foreach ($item in $items) {
            $time = ([DateTime]::SpecifyKind([DateTime]::new([long]$item[0]), [DateTimeKind]::Utc)).ToLocalTime().ToString("HH:mm:ss.fffffff")
            $elapsed = ($item[0] - $state.OriginTicks) / 10000.0
            $name = "$($item[2])" -replace "[\t\r\n]", " "
            [void]$builder.Append(("{0}`t{1}`t{2}`t{3}`t{4}`r`n" -f $time, $elapsed.ToString("0.000", [System.Globalization.CultureInfo]::InvariantCulture), $item[1], $name, $item[3]))
        }
        [void][System.IO.Directory]::CreateDirectory($TraceDir)
        [System.IO.File]::AppendAllText($file, $builder.ToString(), (New-Object System.Text.UTF8Encoding($false)))
    } catch {
        $null = $_
    }
}

function testTraceNetworkPath {
    # ネットワークの場所には書かない（遅い・止まるため）。folder.ps1 の testNetworkPath があればそれで、無ければ \\ で始まるかで見る
    param (
        [string]$Path
    )

    if ($null -ne (Get-Command -Name testNetworkPath -CommandType Function -ErrorAction SilentlyContinue)) {
        return [bool](testNetworkPath $Path)
    }
    return $Path.StartsWith("\\") -and -not $Path.StartsWith("\\?\")
}
