# 検索結果をファイル（.xlsx・.csv・.tsv）に出力する書き手（状態層）。画面のスレッドでは書かず、検索の司令のスレッド（search_service.ps1）が、
# 書き出しの要求（newExportRequest）を invokeSearchExport で実行する。
#
# 書き出す中身: 表に出ている分（要求の ShownHits。画面が表の順に渡す）→ 件数の上限を超えて表に入らなかった分（溢れ）の順。
#   溢れは検索の間には書き溜めない。出力するときに、元の検索の条件と本文インデックスのファイルの一覧（Source.ContentIndexFiles）で、上限なしに照合し直して書く
#   （先頭から Source.Limit 件は表の分なので飛ばす。表は作業の順で先頭から Limit 件なので、照合し直した結果の先頭と一致する）。
#   溢れを含めるのは、元の検索が上限を超えていて（Truncated）、総数の時間切れ・検索の中止のどちらでもないときだけ。
#   それ以外は表の分だけ（「先頭だけ」）。溢れの行は元の検索の順のまま、絞り込み（FilterText）に合うものだけを書く。
# 1 件ずつ列（toResultCells。形式ごとに決まる）にして書き手に渡す。全部をメモリに読まない（照合し直すヒットは、作業ごとに受け取って書き、溜めない）。
#   .xlsx … 見出しは 5 列に決まっているので、<出力先>.new に直接ストリームで書く（xlsx_writer.ps1）。
#   .csv・.tsv … 見出しの列の数は全件を見ないと決まらないので、本文を一時ファイル（溢れの置き場所の export\）に書きながら最大のセルの数を数え、
#                終わってから見出し・本文を <出力先>.new に書く（csv_writer.ps1）。
#   照合し直す間に正規表現の照合が時間切れになったら、そこまでに書いた分は捨てて表の分だけで書き直す（理由 ExportTimeout）。
# どちらも最後に出力先と置き換える。中止・失敗のときは一時ファイルと .new を消し、出力先は前のままにする。

# 溢れの置き場所（根\<PID>）の下の、書き出しの一時ファイルのフォルダ。次の検索の始めと画面を閉じるときに、溢れごと消える
${searchExportDirName} = "export"

# 何行ごとに、書いた件数（Written）を更新するか（中止（Stop）は 1 行ごとに見る）
${searchExportCheckLines} = 1000

# .xlsx の見出し（ヒットした行は 1 セルに入れるので 5 列）
${searchExportXlsxHeaders} = @("ファイル名", "場所", "種別", "行", "該当行")

function newExportRequest {
    # 書き出しの要求を作る（画面が作り、SearchService.Export で渡す）。検索の要求と同じく Synchronized の hashtable。
    #   画面が書く: Source（出力する元の検索の要求）・OutputPath（出力先）・ShownHits（表に出ている分の素の値。RelDir・Book・Location・LineNumber・Line）・
    #              FilterText（絞り込み）・Word（検索ワード。状態の文言用）・Format（xlsx / csv / tsv）・
    #              TmpRoot（出力先の置き場 open_copy を持つ作業フォルダの tmp。空でなければ、置き場がリンクでないことを書く前に確かめる）・Stop（中止）
    #   司令が書く: Written（書いた件数。途中も増える）・Outcome（Done / Cancelled / Locked / Denied / Failed）・Scope（All / Overflow / HeadOnly）・
    #              HeadOnlyReason（先頭だけになった理由 CountTimeout / Cancelled / ExportTimeout）・IndexChanged（照合し直した本文インデックスのファイルが、検索のあとに変わっていた）・TruncatedCells（.xlsx で切ったセルの数）・
    #              DroppedRows（.xlsx の行の上限で書かなかった件数）・Error・Finished
    param (
        $source,
        [string]$outputPath,
        $shownHits,
        [string]$filterText,
        [string]$word,
        [string]$format = "xlsx",
        [string]$tmpRoot = ""
    )

    return [hashtable]::Synchronized(@{
        Kind = "Export"; Number = 0; WorkDir = $source.WorkDir; Word = $word
        Source = $source; OutputPath = $outputPath; ShownHits = $shownHits; FilterText = $filterText
        Format = $format; TmpRoot = $tmpRoot
        Stop = $false; Finished = $false; Written = 0; Outcome = ""; Scope = ""; HeadOnlyReason = ""; IndexChanged = $false; TruncatedCells = 0; DroppedRows = 0; Error = $null
    })
}

function getSearchExportScope {
    # 元の検索 source から、何を書き出すかを決める: @{ Scope; Reason }
    #   All     … 上限を超えていない（表がすべて）
    #   Overflow … 上限を超えていて、照合し直して溢れも書ける
    #   HeadOnly … 上限を超えているが溢れは書けない。Reason は CountTimeout（総数の時間切れ）/ Cancelled（検索を止めた）。
    #              照合し直す間の時間切れ（ExportTimeout）は、書いている途中で決まる（invokeSearchExport）
    param (
        $source
    )

    if (!$source.Truncated) {
        return @{ Scope = "All"; Reason = "" }
    }
    foreach ($reason in @("CountTimeout", "Cancelled")) {
        if ($source.$reason) {
            return @{ Scope = "HeadOnly"; Reason = $reason }
        }
    }
    return @{ Scope = "Overflow"; Reason = "" }
}

function getSearchExportIoOutcome {
    # IOException の種類を分ける。ほかのアプリが開いている（共有違反・ロック違反）なら Locked、それ以外（フォルダが無い・届かない共有・ディスクの空きなど）は Failed
    param (
        $exception
    )

    $code = $exception.GetBaseException().HResult -band 0xFFFF
    if ($code -eq 32 -or $code -eq 33) {
        return "Locked"
    }
    return "Failed"
}

function testSearchExportTarget {
    # 出力先に書けるかを、書き始める前に調べる。出力先のフォルダを作り、ファイルがあれば共有なしで開いて閉じる。
    # 返すもの: 空文字（書ける）/ Locked（ほかのアプリが開いている）/ Denied（読み取り専用・権限が無い。アプリを閉じても直らない）/ Failed（フォルダを作れない・届かないなど）
    param (
        [string]$path
    )

    try {
        [void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path))
        if ([System.IO.File]::Exists($path)) {
            [System.IO.File]::Open($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None).Dispose()
        }
        return ""
    } catch [System.UnauthorizedAccessException] {
        return "Denied"
    } catch [System.IO.IOException] {
        return getSearchExportIoOutcome $_.Exception
    }
}

function getSearchExportDelimiter {
    # 形式ごとの区切り（csv はコンマ、tsv はタブ）
    param (
        [string]$format
    )

    if ($format -eq "tsv") {
        return "`t"
    }
    return ","
}

function newSearchExportSink {
    # 書き手を開く。返すもの: 書いた件数・最大のセルの数・書き手を持つ state
    #   xlsx … newPath に直接書く手を開く（見出しは 5 列）。開かれていれば IOException、読み取り専用なら UnauthorizedAccessException を投げる
    #   csv・tsv … bodyPath に本文を書く手を開く
    param (
        [string]$format,
        [string]$bodyPath,
        [string]$newPath
    )

    $state = @{ Format = $format; Count = 0; Columns = 0; Seen = 0; Delimiter = ""; Xlsx = $null; Body = $null; TimedOut = $false }
    if ($format -eq "xlsx") {
        $state.Xlsx = newXlsxWriter $newPath ${searchExportXlsxHeaders}
        return $state
    }
    $state.Delimiter = getSearchExportDelimiter $format
    $stream = [System.IO.FileStream]::new((toLongPath $bodyPath), [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
    $state.Body = [System.IO.StreamWriter]::new($stream, [System.Text.UTF8Encoding]::new($false))
    $state.Body.NewLine = "`r`n"
    return $state
}

function addSearchExportRow {
    # 1 件（列の配列）を書き手に渡し、件数と、見出しに要る最大のセルの数（ファイル名・場所・種別・行番号の 4 列を除く）を state に反映する
    param (
        [hashtable]$state,
        [string[]]$cells
    )

    if ($state.Format -eq "xlsx") {
        if (addXlsxRow $state.Xlsx $cells) {
            $state.Count++
        }
        return
    }
    $state.Body.WriteLine((toCsvLine $cells $state.Delimiter))
    $state.Count++
    $columns = $cells.Count - 4
    if ($columns -gt $state.Columns) {
        $state.Columns = $columns
    }
}

function stepSearchExport {
    # 1 行（読んだ行・表の行）進むたびに呼ぶ。中止（Stop）は 1 行ごとに見る（1 行の組み立ては遅く、数百行ごとでは中止から止まるまで数秒かかるため）。
    # 書いた件数（Written）は一定の行ごとに要求に出す。中止されていたら $true
    param (
        [hashtable]$request,
        [hashtable]$state
    )

    if ($request.Stop) {
        return $true
    }
    $state.Seen++
    if (($state.Seen % ${searchExportCheckLines}) -eq 0) {
        $request.Written = $state.Count
    }
    return $false
}

function writeSearchExportShown {
    # 表に出ている分を書く。中止されたら $false
    param (
        [hashtable]$request,
        [hashtable]$state
    )

    foreach ($hit in $request.ShownHits) {
        if (stepSearchExport $request $state) {
            return $false
        }
        addSearchExportRow $state (toResultCells $hit.RelDir $hit.Book $hit.Location ([int]$hit.LineNumber) $hit.Line $state.Format)
    }
    return $true
}

function testSearchExportIndexChanged {
    # 元の検索で照合した本文インデックスのファイルの一覧（Path・Ticks・Size）が、今のファイルと違うか（無くなった・更新日時や大きさが違う）
    param (
        $contentIndexFiles
    )

    foreach ($file in $contentIndexFiles) {
        try {
            $info = [System.IO.FileInfo]::new((toLongPath $file.Path))
            if (!$info.Exists -or $info.LastWriteTimeUtc.Ticks -ne $file.Ticks -or $info.Length -ne $file.Size) {
                return $true
            }
        } catch {
            return $true
        }
    }
    return $false
}

function writeSearchExportOverflow {
    # 上限を超えて表に入らなかった分を、上限なしに照合し直して書く（表の分は、照合し直した結果の先頭から Source.Limit 件なので飛ばす）。
    # 作業ごとにヒットを受け取って書き、溜めない。中止されたら $false。照合の時間切れは state.TimedOut を立てて $true（呼ぶ側が表の分だけで書き直す）
    param (
        [hashtable]$request,
        [hashtable]$state
    )

    $source = $request.Source
    $request.IndexChanged = testSearchExportIndexChanged $source.ContentIndexFiles
    $exportRequest = $request
    $exportState = $state
    $exportCtx = @{ Skip = [int]$source.Limit; Cancelled = $false }
    try {
        $result = searchContentIndex $source.Word $source.ContentIndexFiles ([bool]$source.SimpleMatch) 0 -caseSensitive ([bool]$source.CaseSensitive) `
            -fileFilter ([string]$source.FileFilter) -includeShapes ([bool]$source.IncludeShapes) -includeComments ([bool]$source.IncludeComments) `
            -keepHits $false -stopSignal $request -onProgress {
            param ($done, $total, $newHits, $hitTotal)
            foreach ($hit in $newHits) {
                if ($exportCtx.Skip -gt 0) {
                    $exportCtx.Skip--
                    continue
                }
                if (stepSearchExport $exportRequest $exportState) {
                    $exportCtx.Cancelled = $true
                    return
                }
                if (!(testHitFilter $hit.Book $exportRequest.FilterText)) {
                    continue
                }
                addSearchExportRow $exportState (toResultCells $hit.RelDir $hit.Book $hit.Location ([int]$hit.LineNumber) $hit.Line $exportState.Format)
            }
        }
    } catch {
        if ($_.Exception.Message -eq ${contentIndexTimeoutMessage}) {
            $state.TimedOut = $true
            return $true
        }
        throw
    }
    return !($exportCtx.Cancelled -or $request.Stop -or $result.Cancelled)
}

function writeSearchExportBody {
    # 本文（表の分 → 溢れ）を state の書き手に書く。中止されたら $false
    param (
        [hashtable]$request,
        [hashtable]$scope,
        [hashtable]$state
    )

    if (!(writeSearchExportShown $request $state)) {
        return $false
    }
    if ($scope.Scope -eq "Overflow") {
        return (writeSearchExportOverflow $request $state)
    }
    return $true
}

function getSearchExportBodyPath {
    # csv・tsv の本文の一時ファイルの場所。溢れの置き場所が使えれば根\<PID>\export\ の下（後始末・リンクの確かめが溢れと同じに効く）、使えなければ出力先の横
    param (
        [hashtable]$request
    )

    $own = [Workspace]::new($request.WorkDir)
    if (testSearchOverflowRoot (getSearchOverflowRoot $own.TmpRoot) $own.TmpRoot) {
        $dir = Join-Path (getSearchOverflowPidDir $own.TmpRoot) ${searchExportDirName}
        if (newDirectoryUnder $dir (getSearchOverflowRoot $own.TmpRoot)) {
            return Join-Path $dir "body.tmp"
        }
    }
    return "$($request.OutputPath).part"
}

function closeSearchExportSink {
    # 書き手を閉じる。xlsx は完成させて @{ TruncatedCells; DroppedRows } を返す。csv・tsv は本文を閉じて、見出し・本文を newPath に書く
    param (
        [hashtable]$state,
        [string]$bodyPath,
        [string]$newPath
    )

    if ($state.Format -eq "xlsx") {
        $done = closeXlsxWriter $state.Xlsx
        $state.Xlsx = $null
        return @{ TruncatedCells = $done.TruncatedCells; DroppedRows = $done.DroppedRows }
    }
    $state.Body.Dispose()
    $state.Body = $null
    $headers = (toResultHeader $state.Columns).Split("`t")
    writeCsvFile $bodyPath $newPath $headers $state.Delimiter
    return @{ TruncatedCells = 0; DroppedRows = 0 }
}

function abortSearchExportSink {
    # 完成させずに書き手を閉じる（中止・失敗。閉じる途中の例外は握りつぶす）。ファイルを消すのは呼ぶ側
    param (
        $state
    )

    if ($null -eq $state) {
        return
    }
    if ($null -ne $state.Xlsx) {
        abortXlsxWriter $state.Xlsx
    }
    if ($null -ne $state.Body) {
        try {
            $state.Body.Dispose()
        } catch {
            Write-Verbose "本文の一時ファイルを閉じられませんでした"
        }
    }
}

function removeSearchExportTemp {
    # 書き出しの一時ファイル・.new を消す（無くてもよい。消せなくても例外は投げない。残りは次の書き出しの始め・次の検索で消える）
    param (
        [string[]]$paths
    )

    foreach ($path in $paths) {
        if (!$path) { continue }
        try {
            [System.IO.File]::Delete((toLongPath $path))
        } catch {
            Write-Verbose "一時ファイルを消せませんでした: $path"
        }
    }
}

function restartSearchExportHeadOnly {
    # 照合し直す間に時間が切れたとき、そこまでに書いた分は捨て、表の分だけで書き直す（理由 ExportTimeout）。新しい書き手の state を返す。中止されたら $null
    param (
        [hashtable]$request,
        [hashtable]$state,
        [string]$bodyPath,
        [string]$newPath
    )

    abortSearchExportSink $state
    removeSearchExportTemp @($bodyPath, $newPath)
    $request.Scope = "HeadOnly"
    $request.HeadOnlyReason = "ExportTimeout"
    $fresh = newSearchExportSink $request.Format $bodyPath $newPath
    if (!(writeSearchExportShown $request $fresh)) {
        abortSearchExportSink $fresh
        return $null
    }
    return $fresh
}

function invokeSearchExport {
    # 書き出しの要求（newExportRequest）を実行する。結果は要求の Outcome・Written・Scope・HeadOnlyReason・TruncatedCells・DroppedRows に入れ、最後に Finished を立てる。
    # 中止・失敗のときは、一時ファイルと .new を消し、出力先は前のままにする
    param (
        [hashtable]$request
    )

    $target = toLongPath $request.OutputPath
    $newPath = "${target}.new"
    $partPath = "${target}.part"
    $bodyPath = ""
    $state = $null
    try {
        if ($request.Stop) {
            $request.Outcome = "Cancelled"
            return
        }
        # 画面のスレッドでは確かめない（届かない場所でも画面が止まらないように）。出力先の置き場（open_copy）が、ワークスペースの tmp の中の本物のフォルダで、
        # リンク（ジャンクション）で外につながっていないことを、書く前に確かめる（［新規で開く］のコピーと同じ確かめ。出力先は <open_copy>\<12 桁>\<名前>）
        if ($request.TmpRoot) {
            $copyDir = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetDirectoryName($request.OutputPath))
            if (!(testOpenCopyDir $copyDir $request.TmpRoot)) {
                $request.Outcome = "Failed"
                $request.Error = "出力先の置き場が使えません（作業フォルダの tmp の中の open_copy ではないか、リンクになっています）"
                return
            }
        }
        $blocked = testSearchExportTarget $target
        if ($blocked) {
            $request.Outcome = $blocked
            if ($blocked -eq "Failed") {
                $request.Error = "出力先のフォルダを作れないか、届きません"
            }
            return
        }
        # 強制終了などで残った .new・.part（溢れの置き場所が使えないときの本文の一時ファイル）は、置き換える前なので、ここで消してよい
        removeSearchExportTemp @($newPath, $partPath)
        $scope = getSearchExportScope $request.Source
        $request.Scope = $scope.Scope
        $request.HeadOnlyReason = $scope.Reason
        if ($request.Format -ne "xlsx") {
            $bodyPath = getSearchExportBodyPath $request
        }
        $state = newSearchExportSink $request.Format $bodyPath $newPath
        if (!(writeSearchExportBody $request $scope $state)) { $request.Outcome = "Cancelled"; return }
        if ($state.TimedOut) {
            $state = restartSearchExportHeadOnly $request $state $bodyPath $newPath
            if ($null -eq $state) { $request.Outcome = "Cancelled"; return }
        }
        $closed = closeSearchExportSink $state $bodyPath $newPath
        if ([System.IO.File]::Exists($target)) {
            [System.IO.File]::Replace($newPath, $target, [NullString]::Value)
        } else {
            [System.IO.File]::Move($newPath, $target)
        }
        $request.Written = $state.Count
        $request.TruncatedCells = $closed.TruncatedCells
        $request.DroppedRows = $closed.DroppedRows
        $request.Outcome = "Done"
    } catch [System.UnauthorizedAccessException] {
        $request.Outcome = "Denied"
    } catch [System.IO.IOException] {
        $request.Outcome = getSearchExportIoOutcome $_.Exception
        if ($request.Outcome -eq "Failed") {
            $request.Error = $_.Exception.Message
        }
    } catch {
        $request.Error = $_.Exception.Message
        $request.Outcome = "Failed"
    } finally {
        abortSearchExportSink $state
        removeSearchExportTemp @($bodyPath, $newPath)
        $request.Finished = $true
    }
}
