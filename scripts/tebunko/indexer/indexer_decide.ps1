# 1 ファイルを取り込み直すかどうかの判断（差分取り込みの要）。
# ファイルにも画面にも触らないため、そのままテストできる（tests\tebunko\indexer\indexer_decide.Tests.ps1）。

# 読み取る内容（抽出版）。読み取る場所を増やしたら、その形式の版を上げる。
# 前の版で取り込んだファイルは、更新が無くても取り込み直す（取り込み一覧の「抽出版」。空は 1）
#   2: Excel（.xlsx / .xlsm）の図形・コメントを読む。
#      Word・PowerPoint のコメント・SmartArt・グラフを読み、Word のテキストボックスを本文から図形に分ける
#      （Word・PowerPoint は旧形式も新形式に変換してから読むため、.doc / .ppt も上げる）
#   3: Excel のグラフ・SmartArt（表示のグラフシートを含む）を読む。
#      グラフの読み方を、タイトル・軸ラベル・系列名だけにし、項目名は読まないようにする
#      （.xlsx / .xlsm はグラフ・SmartArt を読むようになった分で 3。.docx / .docm / .pptx / .pptm は
#      項目名を読まなくなった分だけ上げる。.doc / .ppt はこのアイテムでは上げない）
#   4: Excel（.xlsx / .xlsm）のヘッダー・フッター（表示のワークシート・グラフシート）を読む。
#   4（.docx / .docm / .doc / .pptx / .pptm）: Word・PowerPoint に埋め込んだ Office のファイル（Word 文書・Excel ブック・
#      PowerPoint）の中の文字を、"ページNNN[埋め込みN]"・"スライドNNN[埋め込みN]" として読む
#      （.doc は Word が新形式に変換した後も埋め込みが残るため上げる。.ppt は変換後に埋め込みが残らないため上げない）。
${extractVersions} = @{
    ".xlsx" = 4; ".xlsm" = 4
    ".docx" = 4; ".docm" = 4; ".doc" = 4
    ".pptx" = 4; ".pptm" = 4; ".ppt" = 2
}

function getExtractVersion {
    # ファイルの形式（拡張子）の今の抽出版を返す
    param (
        [string]$path
    )

    $version = ${extractVersions}[[System.IO.Path]::GetExtension($path).ToLowerInvariant()]
    return $(if ($null -eq $version) { 1 } else { $version })
}

function testExtractOutdated {
    # 取り込み一覧の行が、今の抽出版より前の版で取り込んだものか
    param (
        $row
    )

    $version = 1
    if ($row.抽出版) {
        [void][int]::TryParse([string]$row.抽出版, [ref]$version)
    }
    return ($version -lt (getExtractVersion $row.相対パス))
}

function getIngestDecision {
    # 前回の取り込み一覧の行と、いまのファイルの更新日時・サイズから、取り込むかどうかと、その理由を返す。
    #   Ingest: 取り込むか / Reason: done（取り込み済み）・failed（前回失敗。再取り込みするかは呼び出し元が決める）・
    #            new（一覧に無い）・updated（更新された）・pending（前回未完了）・lost（インデックスが無い・壊れている）・
    #            outdated（前の抽出版で取り込んだ）
    param (
        $old,                 # 前回の取り込み一覧の行（無ければ $null）
        [string]$updated,     # いまのファイルの更新日時（formatFileTime）
        [string]$size,        # いまのファイルのサイズ（文字列）
        [bool]$indexComplete  # インデックス（TSV）がそろっているか（状態が「済」のときだけ意味がある）
    )

    # 前回と更新日時・サイズが同じで、前回「未取り込み」で終わっていなければ、同じファイルとみなす
    $sameFile = ($old -and $old.更新日時 -eq $updated -and $old.サイズ -eq $size -and $old.状態 -ne ${stateNew})
    # 取り込み済みでも、インデックス（TSV）が無くなっていれば取り込み直す。
    # 一覧だけを見ると「済」のままになり、検索しても出てこない状態が続くため
    $lostIndex = ($sameFile -and $old.状態 -eq ${stateDone} -and -not $indexComplete)

    if ($sameFile -and -not $lostIndex) {
        if ($old.状態 -eq ${stateFailed}) {
            return @{ Ingest = $false; Reason = "failed" }
        }
        if (testExtractOutdated $old) {
            return @{ Ingest = $true; Reason = "outdated" }
        }
        return @{ Ingest = $false; Reason = "done" }
    }
    if ($lostIndex) {
        return @{ Ingest = $true; Reason = "lost" }
    }
    if ($null -eq $old) {
        return @{ Ingest = $true; Reason = "new" }
    }
    if ($old.状態 -eq ${stateNew}) {
        return @{ Ingest = $true; Reason = "pending" }
    }
    return @{ Ingest = $true; Reason = "updated" }
}

# 利用者が使っている PowerPoint のために取り込めなかったファイルの理由（取り込み一覧の「エラー」列と、更新に失敗したファイルの一覧に出す。
# 状態は「未取り込み」のまま残すので、閉じてからの次の更新で、再取り込みを選ばなくても取り込む）
${postponedReasonRunning} = "PowerPoint が起動していたため取り込めませんでした。PowerPoint を閉じてから、もう一度更新してください。"
${postponedReasonInUse}   = "PowerPoint を使っている方がいたため取り込めませんでした。PowerPoint を保存して閉じてから、もう一度更新してください。"

# 取り込むファイルが最初に要る Office のアプリ（docs/design/indexing/parallel.md）。Office を使わずに読めるファイルは空文字を返す
function getIngestApp {
    # Excel はセルの表示値を読むため、すべて Excel。Word・PowerPoint は、旧形式（.doc / .ppt）だけ Office が要り、新形式は Office を使わずに直接読む
    # （中身が旧形式・パスワード付きと分かったら、読み取りが Office が要ると返し、getRerouteApp のアプリに回し直す）。
    # テキストの拡張子（.txt 等）と PDF（.pdf）は、どの分岐にも当たらないため Office を使わずに読む
    param (
        [string]$relPath
    )

    # 種類は getOfficeFileKind（取り込みの対象の拡張子の表）で決め、旧形式かどうかだけをここで見る
    $kind = getOfficeFileKind $relPath
    if ($kind -eq "Excel") { return "Excel" }
    $extension = [System.IO.Path]::GetExtension($relPath).ToLowerInvariant()
    if ($kind -eq "Word" -and $extension -eq ".doc") { return "Word" }
    if ($kind -eq "PowerPoint" -and $extension -eq ".ppt") { return "PowerPoint" }
    return ""
}

function getRerouteApp {
    # Office を使わずに読めず、Office が要ると返したファイルの、回し直し先の Office のアプリ
    param (
        [string]$relPath
    )

    if ((getOfficeFileKind $relPath) -eq "PowerPoint") {
        return "PowerPoint"
    }
    return "Word"
}

function selectIngestWorker {
    # 取り込みのワーカーのうち、ファイル 1 つを渡す先を決める（取り込みのワーカーは 1 種類。IngestDispatcher が呼ぶ）。
    # Office の COM は作ったスレッドからしか呼べないため、Office のアプリは「初めてそのアプリのファイルを受けたワーカー」が持ち（Owns）、
    # そのアプリのファイルはそのワーカーにだけ渡す。Office のプロセスはアプリごとに最大 1 つのままにする。
    #   Office が要るファイル … 持ち主へ。いなければ、Office を持たないワーカーのうち空いているものを新しい持ち主にする。
    #                         Office を持たないワーカーが 1 つも無いときだけ、空いているワーカーが 2 つ目のアプリを持つ
    #   Office が要らないファイル … Office を持たないワーカーのうち空いているものを先に。無ければ、空いているワーカーへ
    # workers: ワーカーごとの @{ Busy（渡してある数）; Owns（持っているアプリの名前の並び）; Abandoned（諦めた）} の並び。
    #   まだ始めていないワーカーは Busy 0・Owns 空で並べる。app: ファイルが要る Office のアプリ（要らなければ空）。
    #   capacity: 1 ワーカーに同時に渡しておく数
    # 戻り値 @{ Kind; Index }。Kind: Worker（Index のワーカーに渡す）／Wait（空きが無い。今は渡さない）／
    #   Abandoned（渡し先が全部、制限時間を過ぎても利用者が使っている PowerPoint を止めずに諦めたワーカー。後回しにする）
    param (
        $workers,
        [string]$app,
        [int]$capacity
    )

    $free = {
        param ($list)
        $best = -1
        foreach ($i in $list) {
            if ([int]$workers[$i].Busy -lt $capacity -and ($best -lt 0 -or [int]$workers[$i].Busy -lt [int]$workers[$best].Busy)) {
                $best = $i
            }
        }
        return $best
    }
    $alive = @(for ($i = 0; $i -lt $workers.Count; $i++) { if (!$workers[$i].Abandoned) { $i } })
    if ($alive.Count -eq 0) {
        return @{ Kind = "Abandoned"; Index = -1 }
    }
    $bare = @($alive | Where-Object { @($workers[$_].Owns).Count -eq 0 })
    if ($app) {
        for ($i = 0; $i -lt $workers.Count; $i++) {
            if (@($workers[$i].Owns) -contains $app) {
                if ($workers[$i].Abandoned) {
                    return @{ Kind = "Abandoned"; Index = $i }
                }
                if ([int]$workers[$i].Busy -lt $capacity) {
                    return @{ Kind = "Worker"; Index = $i }
                }
                return @{ Kind = "Wait"; Index = -1 }
            }
        }
        $pool = if ($bare.Count -gt 0) { $bare } else { $alive }
        $index = & $free $pool
    } else {
        $index = & $free $bare
        if ($index -lt 0) {
            $index = & $free $alive
        }
    }
    if ($index -lt 0) {
        return @{ Kind = "Wait"; Index = -1 }
    }
    return @{ Kind = "Worker"; Index = $index }
}

function splitCloudItems {
    # クラウドにだけあるファイルの一覧を、取り込むものと、ダウンロードせずに残すものに分ける。
    #   取り込む: 「ダウンロードして取り込む」が選ばれていて、取り込み対象のもの（Failed が $false）。
    #            前回失敗したもの（Failed が $true）は、「前回失敗したものも取り込み直す」も選ばれているときだけ
    #   items の各要素は createTargetList が返す Cloud の要素（Failed: 前回失敗したファイルか）
    param (
        $items,
        [bool]$retryFailed,
        [bool]$includeCloud
    )

    $included = New-Object System.Collections.Generic.List[object]
    $skipped = New-Object System.Collections.Generic.List[object]
    # List[object] を @() に渡すと PowerShell 5.1 で落ちることがあるため、そのまま回す
    foreach ($item in $items) {
        if ($includeCloud -and ((-not $item.Failed) -or $retryFailed)) {
            $included.Add($item)
        }
        else {
            $skipped.Add($item)
        }
    }
    return @{ Included = $included; Skipped = $skipped }
}

# ----------------------------------------------------------------------------
# 取り込みのジャーナル（store\journal\ingest.tsv。司令だけが 1 行ずつ追記する。読み書きは ingest_journal.ps1）の行の形と読み方
# ----------------------------------------------------------------------------
# 記録の種類（1 行 = 1 記録。タブ区切り。取り込みの結果は取り込み一覧にだけ置き、ここには置かない）:
#   版<TAB>版番号            先頭の 1 行。違えば読まない
#   回<TAB>日時<TAB>行数     インデックス作成の 1 回の始め。行数は、その回の始めに書き直した後の取り込み一覧のデータ行の数
#                            （これより後ろに追記された行を「この回の行」とする。日時では決めない）
#   書                       取り込み一覧を書き直す前の印（書き直しの前後で行数がずれるため、印があれば行数で「この回の行」を決めない）
#   始<TAB>相対パス<TAB>回数  取り込みのワーカーに渡した（回数は続けて終わらなかった回数。この渡しを含む）
#   持<TAB>相対パス<TAB>回数  止まった仕事の回数の持ち越し（この回にまだ渡していない）
#   出<TAB>フォルダ          そのフォルダ（インデックスのフォルダからの相対）を本文インデックスに書き出した
#   終                       取り込み一覧を書き直し、全部を書き出し終えた
# 知らない種類・余った列は読み飛ばす。

function convertToIngestJournalLine {
    # ジャーナルの 1 記録（種類と列）を 1 行の文字列にする。列のタブ・改行は 1 行を保つためスペースにする
    param (
        [string]$kind,
        [object[]]$fields = @()
    )

    $parts = New-Object System.Collections.Generic.List[string]
    $parts.Add($kind)
    foreach ($field in $fields) {
        $parts.Add((([string]$field) -replace "[\t\r\n]+", " "))
    }
    return [string]::Join("`t", $parts.ToArray())
}

function convertFromIngestJournal {
    # 前回のジャーナルの行と、そのときの取り込み一覧の相対パス（ファイルの順。同じ相対パスの追記も数える）から、
    # 始めの復旧の内容を決める。戻り値: @{ Valid; Finished; Aligned; Folders; Resume }
    #   Valid   : 版が合い、回の記録がある（無い・違うなら、読まない＝何も復旧しない）
    #   Finished: 回の後に「終」がある（前回はきれいに終わった。Folders・Resume は空）
    #   Aligned : 取り込み一覧の行数で「この回の行」を決められた（書の印が無く、行数が合う）
    #   Folders : 先に書き出すフォルダ（始めたが、出が無い）
    #   Resume  : 取り込み中に止まった仕事と、持ち越した仕事の @{ RelPath; Count } の並び（回数を上げずに返す。上限は呼び出し側が見る）。
    #             始の相対パスは、同じ回に何度も出たら最後の行。この回の行（取り込み一覧の、回の行数より後ろ）に相対パスがあれば済んでいる。
    #             持だけで始が無いものは回数のまま。書の印があるときは、行を見ずに始の仕事を全部返す
    param (
        [string[]]$lines,
        [string[]]$statusRelPaths,
        [int]$version
    )

    $result = @{ Valid = $false; Finished = $false; Aligned = $true; Folders = [string[]]@(); Resume = @() }
    $records = New-Object System.Collections.Generic.List[string[]]
    foreach ($line in @($lines)) {
        if ($null -ne $line -and $line.Trim() -ne "") {
            $records.Add($line.Split("`t"))
        }
    }
    if ($records.Count -eq 0 -or $records[0][0] -cne "版" -or $records[0].Count -lt 2 -or $records[0][1] -ne [string]$version) {
        return $result
    }
    $round = -1
    for ($i = 0; $i -lt $records.Count; $i++) {
        if ($records[$i][0] -ceq "回") {
            $round = $i
        }
    }
    if ($round -lt 0) {
        return $result
    }
    $result.Valid = $true

    $rowCount = 0
    if ($records[$round].Count -lt 3 -or -not [int]::TryParse($records[$round][2], [ref]$rowCount) -or $rowCount -lt 0) {
        $result.Aligned = $false
    }
    $started = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([System.StringComparer]::OrdinalIgnoreCase)
    $held = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([System.StringComparer]::OrdinalIgnoreCase)
    $published = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $checkedNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::Ordinal)
    for ($i = $round + 1; $i -lt $records.Count; $i++) {
        $fields = $records[$i]
        $count = 0
        switch -CaseSensitive ($fields[0]) {
            "始" {
                if ($fields.Count -ge 3 -and [int]::TryParse($fields[2], [ref]$count) -and $count -ge 1 -and !(testIndexRelPath $fields[1] $false $checkedNames)) {
                    $started[$fields[1]] = $count
                }
            }
            "持" {
                if ($fields.Count -ge 3 -and [int]::TryParse($fields[2], [ref]$count) -and $count -ge 1 -and !(testIndexRelPath $fields[1] $false $checkedNames)) {
                    $held[$fields[1]] = $count
                }
            }
            "出" {
                if ($fields.Count -ge 2 -and $fields[1] -ne "") {
                    [void]$published.Add($fields[1])
                }
            }
            "書" { $result.Aligned = $false }
            "終" { $result.Finished = $true }
        }
    }

    $folders = New-Object System.Collections.Generic.List[string]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($relPath in $started.Keys) {
        $folder = [System.IO.Path]::GetDirectoryName($relPath)
        if (!$published.Contains($folder) -and $seen.Add($folder)) {
            $folders.Add($folder)
        }
    }
    $result.Folders = $folders.ToArray()
    if ($result.Finished) {
        return $result
    }

    $paths = @($statusRelPaths)
    if ($result.Aligned -and $rowCount -gt $paths.Count) {
        $result.Aligned = $false
    }
    $roundRows = $null
    if ($result.Aligned) {
        $roundRows = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        for ($i = $rowCount; $i -lt $paths.Count; $i++) {
            [void]$roundRows.Add($paths[$i])
        }
    }
    $resume = New-Object System.Collections.Generic.List[hashtable]
    foreach ($relPath in $started.Keys) {
        if ($null -eq $roundRows -or !$roundRows.Contains($relPath)) {
            $resume.Add(@{ RelPath = $relPath; Count = $started[$relPath] })
        }
    }
    foreach ($relPath in $held.Keys) {
        if (!$started.ContainsKey($relPath)) {
            $resume.Add(@{ RelPath = $relPath; Count = $held[$relPath] })
        }
    }
    $result.Resume = $resume.ToArray()
    return $result
}
