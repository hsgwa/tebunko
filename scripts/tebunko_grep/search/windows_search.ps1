# Windows Search への問い合わせ（OLE DB の Search.CollatorDSO）。読み取り（SELECT）だけを送る。
# 管理者権限は要らない。Windows Search のサービスが止まっている・ポリシーで無効なときは開けない（高速検索を使わない）。

function openWindowsSearch {
    # Windows Search への接続を開く。開けなければ $null
    try {
        $connection = New-Object System.Data.OleDb.OleDbConnection "Provider=Search.CollatorDSO;Extended Properties='Application=Windows';"
        $connection.Open()
        return $connection
    } catch {
        return $null
    }
}

function invokeWindowsSearch {
    # 問い合わせを実行し、行（object[]）の一覧を返す。SELECT 以外は送らない
    param (
        $connection,
        [string]$sql,
        [int]$timeoutSeconds = 10
    )

    if ($sql -notmatch "^\s*SELECT\s") {
        throw "Windows Search へは SELECT だけを送る: $sql"
    }
    $command = $connection.CreateCommand()
    $command.CommandText = $sql
    $command.CommandTimeout = $timeoutSeconds
    $rows = New-Object 'System.Collections.Generic.List[object[]]'
    $reader = $command.ExecuteReader()
    try {
        while ($reader.Read()) {
            $values = New-Object object[] $reader.FieldCount
            [void]$reader.GetValues($values)
            $rows.Add($values)
        }
    } finally {
        $reader.Dispose()
        $command.Dispose()
    }
    return , $rows
}

function testWindowsSearch {
    # 高速検索に使えるか: Windows Search を開けて、system_index が索引の対象か（フォルダ自体が SystemIndex に入っている）
    param (
        [string]$systemRoot = ${systemIndexDir}
    )

    if (![System.IO.Directory]::Exists((toLongPath $systemRoot))) {
        return $false
    }
    $connection = openWindowsSearch
    if ($null -eq $connection) {
        return $false
    }
    try {
        $url = convertToScopeUrl $systemRoot
        $rows = invokeWindowsSearch $connection "SELECT TOP 1 System.ItemUrl FROM SystemIndex WHERE SCOPE='$url'"
        return $rows.Count -gt 0
    } catch {
        return $false
    } finally {
        $connection.Dispose()
    }
}

function testTsvIndexedByWindowsSearch {
    # 本文インデックス（index の TSV）が Windows Search に索引されているか（利用者が対象から外していないか）。
    # 外していなくても結果は正しいが、txt の索引が遅くなるため、画面で案内を出すのに使う
    param (
        [string]$indexRoot = ${indexDir}
    )

    $connection = openWindowsSearch
    if ($null -eq $connection) {
        return $false
    }
    try {
        $rows = invokeWindowsSearch $connection "SELECT TOP 1 System.ItemUrl FROM SystemIndex WHERE SCOPE='$(convertToScopeUrl $indexRoot)' AND System.FileExtension = '.tsv'"
        return $rows.Count -gt 0
    } catch {
        return $false
    } finally {
        $connection.Dispose()
    }
}
