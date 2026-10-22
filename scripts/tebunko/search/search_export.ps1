# 検索結果をファイル（search_results.txt）に出力する書き手（状態層）。画面のスレッドでは書かず、検索の司令のスレッド（search_service.ps1）が、
# 書き出しの要求（newExportRequest）を invokeSearchExport で実行する。
#
# 書き出す中身: 表に出ている分（要求の ShownHits。画面が表の順に渡す）→ 件数の上限を超えて表に入らなかった分（溢れ。search_overflow.ps1）の順。
#   溢れを含めるのは、元の検索が上限を超えていて（Truncated）、総数の時間切れ・溢れの書き損じ・検索の中止のどれでもないときだけ。
#   それ以外は表の分だけ（「先頭だけ」）。溢れの行は元の検索の順のまま、絞り込み（FilterText）に合うものだけを書く。
# 見出しの列の数は全件を見ないと決まらないので、本文を一時ファイル（溢れの置き場所の export\）に書きながら最大のセルの数を数え、
# 終わってから「件数の行・見出し・本文」を <出力先>.new に書き、出力先と置き換える。中止・失敗のときは一時ファイルと .new を消し、出力先は前のままにする。
# 全部をメモリに読まない（溢れは 1 ファイルずつ、区切りが LF だけの行を順に読む）。

# 溢れの置き場所（根\<PID>）の下の、書き出しの一時ファイルのフォルダ。次の検索の始めと画面を閉じるときに、溢れごと消える
${searchExportDirName} = "export"

# 何行ごとに中止（Stop）を見て、書いた件数（Written）を更新するか
${searchExportCheckLines} = 1000

function newExportRequest {
    # 書き出しの要求を作る（画面が作り、SearchService.Export で渡す）。検索の要求と同じく Synchronized の hashtable。
    #   画面が書く: Source（出力する元の検索の要求）・OutputPath（出力先）・ShownHits（表に出ている分の素の値。RelDir・Book・Location・LineNumber・Line）・
    #              FilterText（絞り込み）・Word（見出しの検索ワード）・Stop（中止）
    #   司令が書く: Written（書いた件数。途中も増える）・Outcome（Done / Cancelled / Locked / Denied / Failed）・Scope（All / Overflow / HeadOnly）・
    #              HeadOnlyReason（先頭だけになった理由 CountTimeout / OverflowFailed / Cancelled）・Error・Finished
    param (
        $source,
        [string]$outputPath,
        $shownHits,
        [string]$filterText,
        [string]$word
    )

    return [hashtable]::Synchronized(@{
        Kind = "Export"; Number = 0; WorkDir = $source.WorkDir; Word = $word
        Source = $source; OutputPath = $outputPath; ShownHits = $shownHits; FilterText = $filterText
        Stop = $false; Finished = $false; Written = 0; Outcome = ""; Scope = ""; HeadOnlyReason = ""; Error = $null
    })
}

function getSearchExportScope {
    # 元の検索 source から、何を書き出すかを決める: @{ Scope; Reason }
    #   All     … 上限を超えていない（表がすべて）
    #   Overflow … 上限を超えていて、溢れも書ける（溢れの置き場所 OverflowDir がある）
    #   HeadOnly … 上限を超えているが溢れは書けない。Reason は CountTimeout（総数の時間切れ）/ OverflowFailed（溢れを書けなかった・置き場所が消えている）/ Cancelled（検索を止めた）
    param (
        $source
    )

    if (!$source.Truncated) {
        return @{ Scope = "All"; Reason = "" }
    }
    foreach ($reason in @("CountTimeout", "OverflowFailed", "Cancelled")) {
        if ($source.$reason) {
            return @{ Scope = "HeadOnly"; Reason = $reason }
        }
    }
    # 司令のスレッドを作り直したあとなど、溢れの置き場所がもう無いときは、表の分だけ（書けなかったのと同じ扱い）
    if (!$source.OverflowDir -or ![System.IO.Directory]::Exists($source.OverflowDir)) {
        return @{ Scope = "HeadOnly"; Reason = "OverflowFailed" }
    }
    return @{ Scope = "Overflow"; Reason = "" }
}

function testSearchExportTarget {
    # 出力先に書けるかを、書き始める前に調べる。出力先のフォルダを作り、ファイルがあれば共有なしで開いて閉じる。
    # 返すもの: 空文字（書ける）/ Locked（ほかのアプリが開いている）/ Denied（読み取り専用・権限が無い。アプリを閉じても直らない）
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
        return "Locked"
    }
}


function addSearchExportLine {
    # 本文の一時ファイルに 1 行書き、件数と、見出しに要る最大のセルの数（ファイル名・場所・種別・行番号の 4 列を除く）を state に反映する
    param (
        $writer,
        [hashtable]$state,
        [string]$line
    )

    $writer.WriteLine($line)
    $state.Count++
    $cells = getSearchResultColumnCount $line
    if ($cells -gt $state.Columns) {
        $state.Columns = $cells
    }
}

function stepSearchExport {
    # 1 行（読んだ行・表の行）進むたびに呼ぶ。一定の行ごとに書いた件数を要求に出し、中止されていたら $true
    param (
        [hashtable]$request,
        [hashtable]$state
    )

    $state.Seen++
    if (($state.Seen % ${searchExportCheckLines}) -ne 0) {
        return $false
    }
    $request.Written = $state.Count
    return [bool]$request.Stop
}

function writeSearchExportShown {
    # 表に出ている分を書く。中止されたら $false
    param (
        [hashtable]$request,
        $writer,
        [hashtable]$state
    )

    foreach ($hit in $request.ShownHits) {
        if (stepSearchExport $request $state) {
            return $false
        }
        addSearchExportLine $writer $state (getSearchResultLine $hit.RelDir $hit.Book $hit.Location ([int]$hit.LineNumber) $hit.Line)
    }
    return $true
}

function writeSearchExportOverflowLine {
    # 溢れの 1 行（RelDir<TAB>Book<TAB>Location<TAB>LineNumber<TAB>Line）を、絞り込みに合えば書く。中止されたら $false
    param (
        [hashtable]$request,
        $writer,
        [hashtable]$state,
        [string]$text
    )

    if (stepSearchExport $request $state) {
        return $false
    }
    $parts = $text.Split([char[]]@([char]9), 5)
    if ($parts.Length -lt 5 -or !(testHitFilter $parts[1] $request.FilterText)) {
        return $true
    }
    addSearchExportLine $writer $state (getSearchResultLine $parts[0] $parts[1] $parts[2] ([int]$parts[3]) $parts[4])
    return $true
}

function writeSearchExportOverflowFile {
    # 溢れのファイル 1 つを、LF だけで区切った行として順に読んで書く（ファイルの全部をメモリに読まない）。中止されたら $false
    param (
        [hashtable]$request,
        $writer,
        [hashtable]$state,
        [string]$path
    )

    $reader = [System.IO.StreamReader]::new([System.IO.FileStream]::new($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite), [System.Text.UTF8Encoding]::new($false))
    try {
        $buffer = New-Object char[] 65536
        $rest = ""
        while (($count = $reader.Read($buffer, 0, $buffer.Length)) -gt 0) {
            $parts = ($rest + [string]::new($buffer, 0, $count)).Split([char]10)
            $rest = $parts[$parts.Length - 1]
            for ($i = 0; $i -lt $parts.Length - 1; $i++) {
                if (!(writeSearchExportOverflowLine $request $writer $state $parts[$i])) {
                    return $false
                }
            }
        }
        if ($rest -ne "") {
            return (writeSearchExportOverflowLine $request $writer $state $rest)
        }
        return $true
    } finally {
        $reader.Dispose()
    }
}

function writeSearchExportBody {
    # 本文（表の分 → 溢れ）を bodyPath の一時ファイルに書く。返すもの: @{ Count; Columns; Seen }。中止されたら $null
    param (
        [hashtable]$request,
        [hashtable]$scope,
        [string]$bodyPath
    )

    $state = @{ Count = 0; Columns = 0; Seen = 0 }
    $stream = [System.IO.FileStream]::new($bodyPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
    $writer = [System.IO.StreamWriter]::new($stream, [System.Text.UTF8Encoding]::new($false))
    $writer.NewLine = "`r`n"
    try {
        if (!(writeSearchExportShown $request $writer $state)) {
            return $null
        }
        if ($scope.Scope -eq "Overflow") {
            foreach ($path in (getSearchOverflowFiles $request.Source.OverflowDir)) {
                if (!(writeSearchExportOverflowFile $request $writer $state $path)) {
                    return $null
                }
            }
        }
        return $state
    } finally {
        $writer.Dispose()
    }
}

function getSearchExportBodyPath {
    # 本文の一時ファイルの場所。溢れの置き場所が使えれば根\<PID>\export\ の下（後始末・リンクの確かめが溢れと同じに効く）、使えなければ出力先の横
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
    return "$($request.OutputPath).body"
}

function writeSearchExportFile {
    # 件数の行・見出し・本文（bodyPath）を、BOM 付き UTF-8 で newPath に書く（今の search_results.txt と同じ形。最後に空行が 1 つ付く）
    param (
        [string]$bodyPath,
        [string]$newPath,
        [string]$word,
        [hashtable]$state
    )

    $stream = [System.IO.FileStream]::new($newPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    $writer = [System.IO.StreamWriter]::new($stream, [System.Text.UTF8Encoding]::new($true))
    try {
        $writer.WriteLine((getSearchResultCountLine $word $state.Count))
        $writer.WriteLine((toResultHeader $state.Columns))
        $writer.Flush()
        $body = [System.IO.FileStream]::new($bodyPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        try {
            $body.CopyTo($stream)
        } finally {
            $body.Dispose()
        }
        $writer.WriteLine("")
    } finally {
        $writer.Dispose()
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
            [System.IO.File]::Delete($path)
        } catch {
            Write-Verbose "一時ファイルを消せませんでした: $path"
        }
    }
}

function invokeSearchExport {
    # 書き出しの要求（newExportRequest）を実行する。結果は要求の Outcome・Written・Scope・HeadOnlyReason に入れ、最後に Finished を立てる。
    # 中止・失敗のときは、一時ファイルと .new を消し、出力先は前のままにする
    param (
        [hashtable]$request
    )

    $target = $request.OutputPath
    $newPath = "${target}.new"
    $bodyPath = ""
    try {
        if ($request.Stop) {
            $request.Outcome = "Cancelled"
            return
        }
        $blocked = testSearchExportTarget $target
        if ($blocked) {
            $request.Outcome = $blocked
            return
        }
        # 強制終了などで残った .new・.body（溢れの置き場所が使えないときの本文の一時ファイル）は、置き換える前なので、ここで消してよい
        removeSearchExportTemp @($newPath, "${target}.body")
        $scope = getSearchExportScope $request.Source
        $request.Scope = $scope.Scope
        $request.HeadOnlyReason = $scope.Reason
        $bodyPath = getSearchExportBodyPath $request
        $state = writeSearchExportBody $request $scope $bodyPath
        if ($null -eq $state) {
            $request.Outcome = "Cancelled"
            return
        }
        writeSearchExportFile $bodyPath $newPath $request.Word $state
        if ([System.IO.File]::Exists($target)) {
            [System.IO.File]::Replace($newPath, $target, [NullString]::Value)
        } else {
            [System.IO.File]::Move($newPath, $target)
        }
        $request.Written = $state.Count
        $request.Outcome = "Done"
    } catch [System.UnauthorizedAccessException] {
        $request.Outcome = "Denied"
    } catch [System.IO.IOException] {
        $request.Outcome = "Locked"
    } catch {
        $request.Error = $_.Exception.Message
        $request.Outcome = "Failed"
    } finally {
        removeSearchExportTemp @($bodyPath, $newPath)
        $request.Finished = $true
    }
}
