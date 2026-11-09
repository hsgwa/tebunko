# 検索の一時ファイルの置き場所・後始末（状態層）。上限を超えた分（溢れ）は、検索の間には書き溜めない（出力するときに照合し直す。search_export.ps1）。
# ここには、出力（csv・tsv）の本文の一時ファイルが入る（根\<PID>\export\）。
# 置き場所は $workspace.TmpRoot\<PC の鍵>\search_overflow\<PID>\ 。
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
