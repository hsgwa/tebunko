# 比較の作業フォルダの読み書きと後始末（状態層。ファイルを読み書きする）。
#
#   %TEMP%\tebunko\diff\<画面の PID>\<比較の番号>\
#       抽出要求.tsv  抽出結果.tsv  進捗.txt  優先.tsv  中止要求  比較エラー.txt
#       left\<番号>\<場所>.tsv  順番.txt
#       right\<番号>\<場所>.tsv 順番.txt
#       work\         … 抽出の作業フォルダ（抽出プロセスが使う。1 ファイルごとに空にする）

function newDiffJobDir {
    # 比較の作業フォルダを新しく作り、そのパスを返す
    param (
        [int]$ownerPid = $PID,
        [string]$root = ${diffTempRoot}
    )

    $base = Join-Path $root ([string]$ownerPid)
    $stamp = (Get-Date).ToString("yyyyMMddHHmmssfff")
    $dir = Join-Path $base $stamp
    $suffix = 1
    while (Test-Path -LiteralPath $dir) {
        $dir = Join-Path $base "${stamp}_$suffix"
        $suffix++
    }
    [System.IO.Directory]::CreateDirectory($dir) | Out-Null
    return $dir
}

function removeDiffJobDir {
    # 比較の作業フォルダを消す（ウイルス対策ソフト等が掴んでいることがあるため、少し待って数回試す）
    param (
        [string]$dir
    )

    if ($dir -and (Test-Path -LiteralPath (toLongPath $dir))) {
        removeDirectoryRetry $dir
    }
}

function removeStaleDiffDirs {
    # 画面が強制終了されて残った作業フォルダ（PID のプロセスが無いもの）を消す。消した数を返す
    param (
        [string]$root = ${diffTempRoot}
    )

    if (!(Test-Path -LiteralPath $root)) {
        return 0
    }
    $removed = 0
    foreach ($dir in @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue)) {
        $ownerPid = 0
        if (![int]::TryParse($dir.Name, [ref]$ownerPid)) {
            continue
        }
        if ($ownerPid -eq $PID) {
            continue
        }
        if ($null -ne (Get-Process -Id $ownerPid -ErrorAction SilentlyContinue)) {
            continue
        }
        try {
            removeDirectoryRetry $dir.FullName
            $removed++
        } catch { }
    }
    return $removed
}

function getExtractDir {
    # 1 ファイルの抽出結果（場所ごとの TSV）を置くフォルダ
    param (
        [string]$jobDir,
        [string]$side,
        [int]$id
    )

    return (Join-Path (Join-Path $jobDir $side) ([string]$id))
}

function writeExtractRequest {
    # 抽出要求を書く。items は @{ Id; Side; Path } の配列（この順に抽出する）
    param (
        [string]$jobDir,
        [object[]]$items
    )

    $lines = @($items | ForEach-Object { "$($_.Id)`t$($_.Side)`t$($_.Path)" })
    [System.IO.File]::WriteAllLines((Join-Path $jobDir ${diffRequestFileName}), [string[]]$lines, ${utf8Bom})
}

function readExtractRequest {
    # 抽出要求を読む（@{ Id; Side; Path } の配列）
    param (
        [string]$jobDir
    )

    $path = Join-Path $jobDir ${diffRequestFileName}
    $result = New-Object System.Collections.Generic.List[object]
    foreach ($line in [System.IO.File]::ReadAllLines($path, [System.Text.Encoding]::UTF8)) {
        $parts = $line.Split("`t")
        if ($parts.Count -lt 3) {
            continue
        }
        $result.Add(@{ Id = [int]$parts[0]; Side = $parts[1]; Path = $parts[2] })
    }
    return , $result.ToArray()
}

function addExtractResult {
    # 1 ファイルの抽出の結果を追記する
    param (
        [string]$jobDir,
        [int]$id,
        [string]$side,
        [string]$state,
        [int]$count,
        [string]$message = ""
    )

    $message = $message -replace "[\t\r\n]+", " "
    $line = "$id`t$side`t$state`t$count`t$message"
    [System.IO.File]::AppendAllText((Join-Path $jobDir ${diffResultFileName}), "$line`r`n", ${utf8Bom})
}

function readExtractResults {
    # 抽出の結果を読む（"番号|側" → @{ Id; Side; State; Count; Message }）。書き込み途中の行（列が足りない）は読まない
    param (
        [string]$jobDir
    )

    $result = @{}
    $path = Join-Path $jobDir ${diffResultFileName}
    if (!(Test-Path -LiteralPath $path)) {
        return $result
    }
    $text = readSharedText $path
    foreach ($line in ($text -split "`r?`n")) {
        $parts = $line.Split("`t")
        if ($parts.Count -lt 5) {
            continue
        }
        $id = 0
        if (![int]::TryParse($parts[0], [ref]$id)) {
            continue
        }
        $result["$id|$($parts[1])"] = @{ Id = $id; Side = $parts[1]; State = $parts[2]; Count = [int]$parts[3]; Message = $parts[4] }
    }
    return $result
}

function readSharedText {
    # ほかのプロセスが書いている途中のファイルも読めるように、共有を許して読む
    param (
        [string]$path
    )

    $stream = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        return $reader.ReadToEnd()
    } finally {
        $stream.Dispose()
    }
}

function writeDiffProgress {
    # 進み具合（済んだ数・全体の数・今のファイル）を書く
    param (
        [string]$jobDir,
        [int]$done,
        [int]$total,
        [string]$current
    )

    [System.IO.File]::WriteAllText((Join-Path $jobDir ${diffProgressFileName}), "$done`t$total`t$current", ${utf8Bom})
}

function readDiffProgress {
    # 進み具合を読む（@{ Done; Total; Current }。無ければ $null）
    param (
        [string]$jobDir
    )

    $path = Join-Path $jobDir ${diffProgressFileName}
    if (!(Test-Path -LiteralPath $path)) {
        return $null
    }
    try {
        $parts = (readSharedText $path).Split("`t")
    } catch {
        return $null
    }
    if ($parts.Count -lt 3) {
        return $null
    }
    return @{ Done = [int]$parts[0]; Total = [int]$parts[1]; Current = $parts[2] }
}

function writeDiffPriority {
    # 次に抽出してほしいファイル（@{ Id; Side } の配列。上から）を書く
    param (
        [string]$jobDir,
        [object[]]$items
    )

    # 読む側は、書き込み途中の行（形の合わない行）を読み飛ばすため、そのまま上書きする
    $lines = @($items | ForEach-Object { "$($_.Id)`t$($_.Side)" })
    [System.IO.File]::WriteAllLines((Join-Path $jobDir ${diffPriorityFileName}), [string[]]$lines, ${utf8Bom})
}

function readDiffPriority {
    # 次に抽出してほしいファイル（"番号|側" の配列）
    param (
        [string]$jobDir
    )

    $path = Join-Path $jobDir ${diffPriorityFileName}
    if (!(Test-Path -LiteralPath $path)) {
        return , @()
    }
    try {
        $text = readSharedText $path
    } catch {
        return , @()
    }
    return , @($text -split "`r?`n" | Where-Object { $_ -match "^\d+`t(left|right)$" } | ForEach-Object { $_.Replace("`t", "|") })
}

function requestDiffStop {
    # 抽出プロセスに中止を頼む
    param (
        [string]$jobDir
    )

    [System.IO.File]::WriteAllText((Join-Path $jobDir ${diffStopFileName}), "", ${utf8Bom})
}

function testDiffStopRequested {
    param (
        [string]$jobDir
    )

    return (Test-Path -LiteralPath (Join-Path $jobDir ${diffStopFileName}))
}

function saveExtractedUnits {
    # 抽出の作業フォルダ（tmpDir）にできた TSV を、抽出結果のフォルダ（destDir）へ移し、抽出した順を 順番.txt に書く
    param (
        [string]$tmpDir,
        [string]$destDir
    )

    # Join-Path は \\?\ の付いたパスを扱えないため、パスは [System.IO.Path]::Combine でつなぐ
    $longDest = toLongPath $destDir
    [System.IO.Directory]::CreateDirectory($longDest) | Out-Null
    $files = @(Get-ChildItem -LiteralPath (toLongPath $tmpDir) -Filter "*.tsv" -File | Sort-Object LastWriteTimeUtc, Name)
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($file in $files) {
        [System.IO.File]::Move($file.FullName, [System.IO.Path]::Combine($longDest, $file.Name))
        $names.Add($file.Name)
    }
    [System.IO.File]::WriteAllLines([System.IO.Path]::Combine($longDest, ${diffOrderFileName}), [string[]]$names.ToArray(), ${utf8Bom})
    return $names.Count
}

function readExtractedUnits {
    # 抽出結果のフォルダの TSV を、場所の名前 → 行（string[]）の順序付きの辞書にして返す（順番.txt の順。無ければ名前の順）
    param (
        [string]$dir
    )

    $units = [ordered]@{}
    $longDir = toLongPath $dir
    if (!(Test-Path -LiteralPath $longDir)) {
        return $units
    }
    $orderFile = [System.IO.Path]::Combine($longDir, ${diffOrderFileName})
    $names = if (Test-Path -LiteralPath $orderFile) {
        @([System.IO.File]::ReadAllLines($orderFile, [System.Text.Encoding]::UTF8) | Where-Object { $_ })
    } else {
        @(Get-ChildItem -LiteralPath $longDir -Filter "*.tsv" -File | Sort-Object Name | ForEach-Object { $_.Name })
    }
    foreach ($name in $names) {
        $path = [System.IO.Path]::Combine($longDir, $name)
        if (!(Test-Path -LiteralPath $path)) {
            continue
        }
        $place = decodeIndexPlace ([System.IO.Path]::GetFileNameWithoutExtension($name))
        $units[$place] = [string[]][System.IO.File]::ReadAllLines($path, [System.Text.Encoding]::UTF8)
    }
    return $units
}

function getFileHash {
    # ファイルの SHA-256（16 進数の文字列）。ほかのアプリが開いていても読めるよう、共有を許して読む
    param (
        [string]$path
    )

    $stream = New-Object System.IO.FileStream((toLongPath $path), [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $sha = [System.Security.Cryptography.SHA256]::Create()
        try {
            return [BitConverter]::ToString($sha.ComputeHash($stream)).Replace("-", "")
        } finally {
            $sha.Dispose()
        }
    } finally {
        $stream.Dispose()
    }
}

function getFolderFiles {
    # フォルダの Office ファイルを @{ RelPath; Path; Size; Time } の配列で返す（subfolders が false なら直下だけ）。
    # アクセスできないフォルダがあれば HasError
    param (
        [string]$folder,
        [bool]$subfolders = $true
    )

    $scan = findOfficeFiles $folder
    $longRoot = (toLongPath $scan.Root).TrimEnd("\") + "\"
    $files = New-Object System.Collections.Generic.List[object]
    foreach ($file in $scan.Files) {
        $fullName = $file.FullName
        $relative = if ($fullName.StartsWith($longRoot, [System.StringComparison]::OrdinalIgnoreCase)) { $fullName.Substring($longRoot.Length) } else { $file.Name }
        if (!$subfolders -and $relative.Contains("\")) {
            continue
        }
        $files.Add(@{ RelPath = $relative; Path = (fromLongPath $fullName); Size = $file.Length; Time = $file.LastWriteTime })
    }
    return @{ Files = $files.ToArray(); HasError = $scan.HasError }
}

function writeDiffReport {
    # 比較結果ファイルを書く（UTF-8 BOM 付き・CRLF・上書き）。書いたパスを返す
    param (
        [string[]]$lines,
        [string]$path = ${diffResultFile}
    )

    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
    [System.IO.File]::WriteAllLines($path, $lines, ${utf8Bom})
    return $path
}
