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

function getWindowsSearchState {
    # 高速検索に使えるかと、使えない理由。次のどれかの文字列を返す（画面の文言は index_view.ps1 の getFastSearchRowView）。
    #   NoFolder     : system_index のフォルダが無い（インデックスを作っていない・ワークスペースが違う）
    #   NoConnection : Windows Search を開けない、または問い合わせに失敗した（時間切れを含む）
    #   NotInScope   : system_index が 0 件で、ワークスペースの中も 1 件も索引されていない（索引の対象外とみなす）
    #   NotYet       : system_index が 0 件だが、ワークスペースの中のほかのものは索引されている（対象だが、まだ索引されていないとみなす）
    #   Ok           : system_index が 1 件以上索引されている（フォルダ自体が SystemIndex に入っている）
    # 管理者権限なしに索引の対象の一覧を読む方法が無いため、NotInScope と NotYet はワークスペースの中の様子で見分ける。
    #   workspaceDir: 空なら、systemRoot の 1 つ上（ワークスペース）
    #   connection  : 開いた接続。$null なら自分で開いて閉じる（渡したら呼び出し側が閉じる）
    param (
        [string]$systemRoot = $workspace.SystemIndexDir,
        [string]$workspaceDir = "",
        $connection = $null
    )

    if (![System.IO.Directory]::Exists((toLongPath $systemRoot))) {
        return "NoFolder"
    }
    if (!$workspaceDir) {
        $workspaceDir = [System.IO.Path]::GetDirectoryName($systemRoot.TrimEnd("\"))
    }
    $own = $null -eq $connection
    if ($own) {
        $connection = openWindowsSearch
        if ($null -eq $connection) {
            return "NoConnection"
        }
    }
    try {
        $rows = invokeWindowsSearch $connection "SELECT TOP 1 System.ItemUrl FROM SystemIndex WHERE SCOPE='$(convertToScopeUrl $systemRoot)'"
        if ($rows.Count -gt 0) {
            return "Ok"
        }
        $rows = invokeWindowsSearch $connection "SELECT TOP 1 System.ItemUrl FROM SystemIndex WHERE SCOPE='$(convertToScopeUrl $workspaceDir)'"
        return $(if ($rows.Count -gt 0) { "NotYet" } else { "NotInScope" })
    } catch {
        return "NoConnection"
    } finally {
        if ($own) {
            $connection.Dispose()
        }
    }
}

function testWindowsSearch {
    # 高速検索に使えるか: Windows Search を開けて、system_index が索引の対象か（getWindowsSearchState が Ok）
    param (
        [string]$systemRoot = $workspace.SystemIndexDir
    )

    return (getWindowsSearchState $systemRoot) -eq "Ok"
}

function testTsvIndexedByWindowsSearch {
    # 本文インデックス（index の TSV）が Windows Search に索引されているか（利用者が対象から外していないか）。
    # 外していなくても結果は正しいが、txt の索引が遅くなるため、画面で案内を出すのに使う
    #   connection: 開いた接続。$null なら自分で開いて閉じる（渡したら呼び出し側が閉じる）
    param (
        [string]$indexRoot = $workspace.IndexDir,
        $connection = $null
    )

    $own = $null -eq $connection
    if ($own) {
        $connection = openWindowsSearch
        if ($null -eq $connection) {
            return $false
        }
    }
    try {
        $rows = invokeWindowsSearch $connection "SELECT TOP 1 System.ItemUrl FROM SystemIndex WHERE SCOPE='$(convertToScopeUrl $indexRoot)' AND System.FileExtension = '.tsv'"
        return $rows.Count -gt 0
    } catch {
        return $false
    } finally {
        if ($own) {
            $connection.Dispose()
        }
    }
}
