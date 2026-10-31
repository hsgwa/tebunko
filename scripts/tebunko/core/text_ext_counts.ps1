# 前の回に数えたテキストの拡張子ごとの件数の控え（状態層）。置き場は <ワークスペース>\store\text_ext_counts.json（Workspace.TextExtCountsFile）。
# インデックスの名前ごとに、拡張子 → 件数 と数えた日時を持つ。詳細の「取り込む拡張子」ダイアログが、前の回の件数を出すために読む。
# インデクサが、そのインデックスを含む回が終わり方 0 で終わったときにだけ書く。消えても・読めなくても困らない控え
# （読めなければ「控え無し」として扱い、次の回で書き直す）。本文インデックス・取り込み一覧の形式ではなく、エクスポートの zip にも入れない。

function readTextExtCounts {
    # 控えを インデックス名 → @{ Time（数えた日時の文字列）; Counts（拡張子 → 件数の Hashtable） } の Hashtable で返す。
    # 無い・読めない・形が違うものは、空（またはその名前の項目なし）にする
    param (
        [string]$path = $workspace.TextExtCountsFile
    )

    $result = @{}
    if ([string]::IsNullOrEmpty($path) -or -not [System.IO.File]::Exists($path)) {
        return $result
    }
    try {
        $data = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
    } catch {
        return $result
    }
    if ($data -isnot [System.Management.Automation.PSCustomObject]) {
        return $result
    }
    foreach ($entry in $data.PSObject.Properties) {
        $item = $entry.Value
        if ($item -isnot [System.Management.Automation.PSCustomObject] -or $item.Counts -isnot [System.Management.Automation.PSCustomObject]) {
            continue
        }
        $counts = @{}
        foreach ($count in $item.Counts.PSObject.Properties) {
            $number = 0
            if ($count.Name -is [string] -and [int]::TryParse([string]$count.Value, [ref]$number) -and $number -ge 0) {
                $counts[[string]$count.Name] = $number
            }
        }
        $result[[string]$entry.Name] = @{ Time = [string]$item.Time; Counts = $counts }
    }
    return $result
}

function writeTextExtCountsTable {
    # 控えの表（readTextExtCounts の形）を、一時のファイルに書いてから置き換える（途中で止まっても半端なファイルを残さない）
    param (
        [hashtable]$table,
        [string]$path
    )

    $root = [ordered]@{}
    foreach ($name in @($table.Keys | Sort-Object)) {
        $counts = [ordered]@{}
        foreach ($extension in @($table[$name].Counts.Keys | Sort-Object)) {
            $counts[$extension] = [int]$table[$name].Counts[$extension]
        }
        $root[$name] = [ordered]@{ Time = [string]$table[$name].Time; Counts = $counts }
    }
    $json = $root | ConvertTo-Json -Depth 4
    $dir = [System.IO.Path]::GetDirectoryName($path)
    [void][System.IO.Directory]::CreateDirectory($dir)
    $temp = "$path.tmp"
    [System.IO.File]::WriteAllText($temp, $json, (New-Object System.Text.UTF8Encoding($false)))
    if ([System.IO.File]::Exists($path)) {
        [System.IO.File]::Delete($path)
    }
    [System.IO.File]::Move($temp, $path)
}

function writeTextExtCounts {
    # そのインデックス（名前）の件数と数えた日時を書く。ほかの名前の項目は残す
    param (
        [string]$name,
        $counts,   # 拡張子 → 件数
        [string]$time,
        [string]$path = $workspace.TextExtCountsFile
    )

    $table = readTextExtCounts $path
    $copy = @{}
    foreach ($extension in @($counts.Keys)) { $copy[[string]$extension] = [int]$counts[$extension] }
    $table[$name] = @{ Time = $time; Counts = $copy }
    writeTextExtCountsTable $table $path
}

function renameTextExtCounts {
    # インデックスの名前を変えたとき、控えの項目のキーを付け替える。古い名前の項目が無ければ何もしない
    param (
        [string]$oldName,
        [string]$newName,
        [string]$path = $workspace.TextExtCountsFile
    )

    $table = readTextExtCounts $path
    if (-not $table.ContainsKey($oldName)) {
        return
    }
    $moved = $table[$oldName]
    $table.Remove($oldName)
    $table[$newName] = $moved
    writeTextExtCountsTable $table $path
}

function removeTextExtCounts {
    # インデックスを削除したとき、控えのその名前の項目を消す。項目が無ければ何もしない
    param (
        [string]$name,
        [string]$path = $workspace.TextExtCountsFile
    )

    $table = readTextExtCounts $path
    if (-not $table.ContainsKey($name)) {
        return
    }
    $table.Remove($name)
    writeTextExtCountsTable $table $path
}
