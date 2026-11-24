# 検索結果が件数の上限（searchLimit）を超えたとき、表に入らない分（溢れ）の置き場所・書き手・読み手・後始末（状態層）。
# 溢れは $workspace.TmpRoot\<PC の鍵>\search_overflow\<PID>\<要求の番号>\ に、照合の作業ごとのファイルとして書き溜める。
#   <作業の番号>.head.tsv … 司令が、作業の返した行のうち表に入らなかった分を書く
#   <作業の番号>.tsv      … 作業が、数えるだけに切り替えた後のヒットを書く（searchContentIndexFiles）
# 全件の順は「表 → 作業の番号の順に（head → 作業のファイル）」で、1 スレッドで探した順と同じになる。
# 1 行は RelDir<TAB>Book<TAB>Location<TAB>LineNumber<TAB>Line（BOM なしの UTF-8）。
# 作る・消すときは、置き場所が TmpRoot\<PC の鍵>\search_overflow そのものであり、リンクで外へつながっていないことを確かめてから行う（testSearchOverflowRoot）。

# 溢れの置き場所の名前
${searchOverflowDirName} = "search_overflow"

function getSearchOverflowRoot {
    # 溢れを置く根（tmpRoot\<PC の鍵>\search_overflow）を返す。文字列を組み立てるだけで、ファイルには触らない。
    # 取り込みの作業フォルダ（<PC の鍵>\<PID>）と同じ親の下に置くが、名前が数字でないため、取り込みの片付け（removeStaleProcessDirs）は触らない
    param (
        [string]$tmpRoot
    )

    return Join-Path (Join-Path $tmpRoot (getMachineKey)) ${searchOverflowDirName}
}

function getSearchOverflowPidDir {
    # このプロセスの溢れのフォルダ（根\<PID>）。画面を閉じるときと、次の検索を始めるときに、中ごと消す
    param (
        [string]$tmpRoot
    )

    return Join-Path (getSearchOverflowRoot $tmpRoot) $PID
}

function getSearchOverflowDir {
    # 検索の要求 1 つの溢れのフォルダ（根\<PID>\<要求の番号>）。文字列を組み立てるだけ
    param (
        [string]$tmpRoot,
        [int]$requestNumber
    )

    return Join-Path (getSearchOverflowPidDir $tmpRoot) $requestNumber
}

function testSearchOverflowRoot {
    # root が、溢れの根として使ってよいものか。作る・消すのは、これが真のときだけ行う。
    #   絶対パス（ドライブ付き・UNC）で、途中に .. や余計な区切りを含まず、tmpRoot\<PC の鍵>\search_overflow と一致すること
    #   root・その親（<PC の鍵>）・その親（tmpRoot）が、リンクでないこと（まだ無ければ問題にしない）
    param (
        [string]$root,
        [string]$tmpRoot
    )

    if (!$root -or !$root.Trim() -or !$tmpRoot -or !$tmpRoot.Trim()) {
        return $false
    }
    $trimmed = $root.TrimEnd([char]92, [char]47)
    if ($trimmed -notmatch '^([A-Za-z]:\\|\\\\[^\\])' -or $tmpRoot -notmatch '^([A-Za-z]:\\|\\\\[^\\])') {
        return $false
    }
    try {
        $full = [System.IO.Path]::GetFullPath($trimmed)
        $expected = [System.IO.Path]::GetFullPath((getSearchOverflowRoot $tmpRoot))
    } catch {
        return $false
    }
    if (![string]::Equals($trimmed, $full, [System.StringComparison]::OrdinalIgnoreCase) -or
        ![string]::Equals($full, $expected.TrimEnd([char]92), [System.StringComparison]::OrdinalIgnoreCase)) {
        return $false
    }
    $keyDir = [System.IO.Path]::GetDirectoryName($full)
    foreach ($checked in @($full, $keyDir, [System.IO.Path]::GetDirectoryName($keyDir))) {
        if (testPathIsLink $checked) {
            return $false
        }
    }
    return $true
}

function removeSearchOverflow {
    # path（溢れの根の下のフォルダ。<PID>・<要求の番号> のどちらでも）を、中ごと消す。消したら（無いときも）$true。
    # 根（testSearchOverflowRoot）の確かめか、根から path までの途中のリンクの確かめ（removeDirectoryUnder）が偽なら、何も消さずに $false。
    # 使用中などで消せなかったときも $false（例外は投げない。残りは次の検索・起動時に消える）
    param (
        [string]$path,
        [string]$tmpRoot
    )

    $root = getSearchOverflowRoot $tmpRoot
    if (!(testSearchOverflowRoot $root $tmpRoot)) {
        return $false
    }
    if (![System.IO.Directory]::Exists((toLongPath $path))) {
        return $true
    }
    try {
        return [bool](removeDirectoryUnder $path $root 1)
    } catch {
        return $false
    }
}

function newSearchOverflowWriter {
    # 溢れのファイル（dir\name）を作って開き、BOM なしの UTF-8 の StreamWriter を返す。フォルダが無ければ作る。
    # 作れない・開けない（容量・権限）ときは $null。照合のスレッドからも呼ぶため、.NET だけで書く（コマンドレットを使わない）
    param (
        [string]$dir,
        [string]$name
    )

    try {
        [void][System.IO.Directory]::CreateDirectory($dir)
        $stream = [System.IO.FileStream]::new((Join-Path $dir $name), [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
        $writer = [System.IO.StreamWriter]::new($stream, [System.Text.UTF8Encoding]::new($false))
        # 行の区切りは LF だけにする（Line の中に CR が入っていても、読み手は LF だけで区切れる）
        $writer.NewLine = "`n"
        return $writer
    } catch {
        return $null
    }
}

function writeSearchOverflowHits {
    # 司令が、表に入らなかった行（searchContentIndex の Hits の要素）を、溢れの head のファイルに書き足す。
    # writer は newSearchOverflowWriter の結果（書き始めは呼ぶ側が決める）。書けなければ $false（writer は閉じない）
    param (
        $writer,
        $hits
    )

    try {
        foreach ($hit in $hits) {
            $writer.WriteLine("$($hit.RelDir)`t$($hit.Book)`t$($hit.Location)`t$($hit.LineNumber)`t$($hit.Line)")
        }
        return $true
    } catch {
        return $false
    }
}

function getSearchOverflowFiles {
    # dir の溢れのファイルを、全件の順（作業の番号の順に、head → 作業のファイル）に並べたパスの一覧を返す。フォルダが無ければ空
    param (
        [string]$dir
    )

    if (!$dir -or ![System.IO.Directory]::Exists($dir)) {
        return @()
    }
    $entries = foreach ($path in [System.IO.Directory]::GetFiles($dir, "*.tsv")) {
        if ([System.IO.Path]::GetFileName($path) -match '^(\d+)(\.head)?\.tsv$') {
            [pscustomobject]@{ Slot = [int]$Matches[1]; Order = $(if ($Matches[2]) { 0 } else { 1 }); Path = $path }
        }
    }
    return @($entries | Sort-Object Slot, Order | ForEach-Object { $_.Path })
}

function readSearchOverflowLines {
    # dir の溢れの全件を、全件の順（getSearchOverflowFiles）の行の文字列で返す（1 行は RelDir<TAB>Book<TAB>Location<TAB>LineNumber<TAB>Line）。
    # 全部をメモリに読むので、テスト・小さい量向け。書き出しは getSearchOverflowFiles を 1 ファイルずつ流して読む
    param (
        [string]$dir
    )

    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($path in (getSearchOverflowFiles $dir)) {
        # 行の区切りは LF だけ（ReadAllLines は CR でも区切るため使わない）
        $text = [System.IO.File]::ReadAllText($path, [System.Text.UTF8Encoding]::new($false))
        if ($text.Length -eq 0) { continue }
        $parts = $text.Split([char]10)
        $count = if ($parts[$parts.Length - 1] -eq "") { $parts.Length - 1 } else { $parts.Length }
        for ($i = 0; $i -lt $count; $i++) { $lines.Add($parts[$i]) }
    }
    return $lines.ToArray()
}

function removeStaleSearchOverflow {
    # 強制終了などで残った溢れを、起動時に消す。根の下の <PID> のフォルダのうち、動いていないプロセスのものだけを消し、消した数を返す。
    # 今のプロセスの <PID>・動いているプロセスのもの（同じ PC で別に起動している tebunko）・別の PC の鍵の下のものは残す。
    # 数字でない名前は触らない。根・途中にリンクがあるものは辿らない（removeSearchOverflow）。
    #   isAlive: PID を受けて、そのプロセスが動いているかを返すスクリプトブロック（テストで差し替える）
    param (
        [string]$tmpRoot,
        [scriptblock]$isAlive = { param ($processId) [bool](Get-Process -Id $processId -ErrorAction SilentlyContinue) }
    )

    $root = getSearchOverflowRoot $tmpRoot
    if (!(testSearchOverflowRoot $root $tmpRoot) -or ![System.IO.Directory]::Exists((toLongPath $root))) {
        return 0
    }
    $removed = 0
    foreach ($dir in @(Get-ChildItem -LiteralPath $root -Directory -Force | Where-Object { $_.Name -match "^\d{1,9}$" -and [int]$_.Name -ne $PID })) {
        if (testLinkItem $dir) {
            continue
        }
        if (& $isAlive ([int]$dir.Name)) {
            continue
        }
        if (removeSearchOverflow $dir.FullName $tmpRoot) {
            $removed++
        }
    }
    return $removed
}
