# インデックス（work\index 配下）の作成・集計・改名・削除。

function testStatusRowFailed {
    # 取り込み一覧の行を「失敗」として数えるか。状態が「失敗」の行と、「未取り込み」で理由（エラー）が書いてある行
    # （利用者の PowerPoint のために取り込めなかったファイル。次の更新で取り込むため状態は「未取り込み」のまま）
    param (
        $row
    )

    $state = [string]$row.状態
    return ($state -eq ${stateFailed}) -or ($state -eq ${stateNew} -and [string]$row.エラー -ne "")
}

function getIndexFileEntries {
    # あるインデックスに取り込んだファイルの一覧（取り込み一覧のそのインデックスの行）を、画面の表に出す形で返す。画面の裏の仕事で動く（数万行を読むため）。
    # 戻り値: @{ Rows（@{ RelPath（インデックス名を除いたフォルダからの相対パス）; UpdatedText; Kind（Done / Failed / Pending）; StateText; Reason; SourcePath（取り込み一覧に記録した元のフォルダと相対パスから組んだ元のファイルのパス。フォルダの記録が無ければ空）} の配列。相対パスの順）; Done; Failed; Pending }
    # 失敗の数え方は testStatusRowFailed と同じ。状態の文言は labels（Kind → 文言。判断層の getIndexFileStateLabels）から取る
    param (
        [string]$indexName,
        [hashtable]$labels,
        [string]$path = $workspace.StatusFile
    )

    $result = @{ Rows = @(); Done = 0; Failed = 0; Pending = 0 }
    if (!$indexName) {
        return $result
    }
    $prefix = $indexName + "\"
    $items = New-Object System.Collections.Generic.List[object]
    $keys = New-Object System.Collections.Generic.List[string]
    $status = readStatusFile $path
    $base = ""
    foreach ($folder in $status.Folders) {
        if ([string]::Equals([string]$folder.Name, $indexName, [System.StringComparison]::OrdinalIgnoreCase)) {
            $base = ([string]$folder.Path).TrimEnd([char]92)
            break
        }
    }
    foreach ($row in $status.Rows.Values) {
        $relPath = [string]$row.相対パス
        if (!$relPath.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }
        $kind = if (testStatusRowFailed $row) { "Failed" } elseif ($row.状態 -eq ${stateDone}) { "Done" } else { "Pending" }
        $result[$kind]++
        $rest = $relPath.Substring($prefix.Length)
        $keys.Add($rest)
        $items.Add([pscustomobject]@{
            RelPath = $rest; UpdatedText = $(if ($row.更新日時.Length -ge 16) { $row.更新日時.Substring(0, 16) } else { [string]$row.更新日時 }); Kind = $kind; StateText = [string]$labels[$kind]; SourcePath = $(if ($base) { $base + "\" + $rest } else { "" })
            Reason = $(if ($kind -eq "Failed") { [string]$row.エラー } else { "" })
        })
    }
    # 型を明示する（引数の型が曖昧だと、並べ替えが別の（並べ替えた結果を返さない）呼び方になる）
    [string[]]$keyArray = $keys.ToArray()
    [object[]]$array = $items.ToArray()
    [Array]::Sort([Array]$keyArray, [Array]$array, [System.Collections.IComparer][System.StringComparer]::OrdinalIgnoreCase)
    $result.Rows = $array
    return $result
}

function getWorkspaceIndexFolders {
    # ワークスペースに残っているインデックスを、名前と前回の元のフォルダの組（@{ Name; Path }）で集めて返す。
    # 取り込み一覧のクロール対象フォルダの行は名前と Path を持つ。インデックスのフォルダ（store\content\・古い content_index\）の直下にだけあって
    # 取り込み一覧に行の無いフォルダは Path を空にする（どのフォルダにも付けられない名前として扱う）。
    # 設定から外れたあとも残してあるインデックスを含む。画面の追加・名前の変更・インポートが、その名前を別のフォルダに付けないために使う
    param (
        $ws
    )

    $folders = New-Object System.Collections.Generic.List[object]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($folder in (readStatusFile $ws.StatusFile).Folders.ToArray()) {
        if ($folder.Name) {
            $folders.Add([pscustomobject]@{ Name = [string]$folder.Name; Path = [string]$folder.Path })
            [void]$seen.Add([string]$folder.Name)
        }
    }
    # 作り直す前の形式のインデックス（古い置き場）の名前も数える
    foreach ($root in @($ws.IndexDir, $ws.OldIndexDir)) {
        if (Test-Path -LiteralPath (toLongPath $root) -PathType Container) {
            foreach ($dir in [System.IO.Directory]::GetDirectories((toLongPath $root))) {
                $name = [System.IO.Path]::GetFileName($dir)
                if ($seen.Add($name)) {
                    $folders.Add([pscustomobject]@{ Name = $name; Path = "" })
                }
            }
        }
    }
    # 1 件ずつ返す（受け取る側は @(...) で包んで配列にする。, を付けて返すと配列が 1 要素の入れ子になる）
    return $folders.ToArray()
}

function getIndexStats {
    # 取り込み一覧の行をインデックス名ごとに集計する（画面のインデックス一覧に出す件数・最終更新）:
    #   インデックス名（大文字・小文字を区別しない）→ @{ Total; Done; Pending; Failed; LastIngested（"yyyy/MM/dd HH:mm:ss"。無ければ空） }
    # rows は readStatusFile の Rows（相対パス → 行）。取り込み一覧を読み直さずに済むよう、読み込み済みの行を受け取る
    param (
        $rows
    )

    $stats = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    if ($null -eq $rows) {
        return , $stats
    }
    foreach ($entry in $rows.GetEnumerator()) {
        $name = (splitIndexRelPath $entry.Key).Name
        if ($name -eq "") {
            continue
        }
        if (!$stats.ContainsKey($name)) {
            $stats[$name] = @{ Total = 0; Done = 0; Pending = 0; Failed = 0; LastIngested = "" }
        }
        $stat = $stats[$name]
        $stat.Total++
        $state = [string]$entry.Value.状態
        if ($state -eq ${stateDone}) {
            $stat.Done++
        } elseif (testStatusRowFailed $entry.Value) {
            $stat.Failed++
        } elseif ($state -eq ${stateNew}) {
            $stat.Pending++
        }
        # 取り込み日時は "yyyy/MM/dd HH:mm:ss" のため、文字列のまま比べて新しい方を残せる
        $ingested = [string]$entry.Value.取り込み日時
        if ($ingested -gt $stat.LastIngested) {
            $stat.LastIngested = $ingested
        }
    }
    return , $stats
}

function removeGramIndexOfWorkspace {
    # インデックスのフォルダ（dir = Workspace.IndexDir）と同じワークスペースの gram 索引から、インデックス name の分を消す。
    # 消せなくてもインデックスの操作は続ける（インデックスが無いフォルダの索引は、次のインデックス作成で消える）
    param (
        [string]$name,
        [string]$dir
    )

    $owner = [Workspace]::new((getWorkspaceDirOfIndexDir $dir))
    try {
        [void](removeGramIndexOf $name $owner.GramDir)
    } catch {
    }
}

function getSearchIndexes {
    # インデックスの一覧（getSearchIndexData の Indexes）。引数は getSearchIndexData と同じ
    param (
        [string]$dir = $workspace.IndexDir,
        [string]$statusPath = $workspace.StatusFile,
        [string]$settingsPath = ${settingsFile}
    )

    return @((getSearchIndexData $dir $statusPath $settingsPath).Indexes)
}

function getSearchIndexData {
    # @{ Root（dir のフルパス）; Sources（インデックス名 → 元のフォルダ。getSourceFolderMap の結果）; Indexes } を返す。
    # Indexes はインデックスの一覧（work\index 直下のフォルダ 1 つがインデックス 1 つ）で、
    # @{ Name（インデックス名）; Path（インデックスのフォルダのフルパス）; SourcePath（元のフォルダ。分からなければ ""） } の配列で返す。
    # 並びは［インデックス管理］の一覧（targetFolders）と同じにし、その一覧に無いもの
    # （別の場所・PC から work\index にコピーしたインデックスなど）は名前順で後ろに付ける
    param (
        [string]$dir = $workspace.IndexDir,
        [string]$statusPath = $workspace.StatusFile,
        [string]$settingsPath = ${settingsFile}
    )

    if (!(Test-Path -LiteralPath $dir -PathType Container)) {
        return @{ Root = $dir; Indexes = @(); Sources = @{} }
    }
    $root = (Resolve-Path -LiteralPath $dir).ProviderPath.TrimEnd("\")
    $sources = getSourceFolderMap $root $statusPath $settingsPath

    # ［インデックス管理］の一覧の順番（インデックス名 → 何番目か）
    $order = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($folder in @(getTargetFolders $settingsPath | Where-Object { $_.Name })) {
        if (!$order.ContainsKey($folder.Name)) {
            $order[$folder.Name] = $order.Count
        }
    }

    $indexes = New-Object System.Collections.Generic.List[object]
    foreach ($sub in @(Get-ChildItem -LiteralPath (toLongPath $root) -Directory -ErrorAction SilentlyContinue)) {
        $name = $sub.Name
        $path = (fromLongPath $sub.FullName)
        if (!$sources.ContainsKey($name)) {
            # 取り込み一覧にも設定にも無いインデックス（別の場所・PC からコピーしたものなど）は、
            # そのフォルダの中の source_folder.txt から元のフォルダを読む
            $own = readSourceFolderFile $path
            if ($own.ContainsKey($name)) {
                $sources[$name] = $own[$name]
            }
        }
        $indexes.Add([pscustomobject]@{
            Name       = $name
            Path       = $path
            SourcePath = $(if ($sources.ContainsKey($name)) { $sources[$name] } else { "" })
            Order      = $(if ($order.ContainsKey($name)) { $order[$name] } else { [int]::MaxValue })
        })
    }
    $sorted = @($indexes | Sort-Object Order, Name | ForEach-Object {
        [pscustomobject]@{ Name = $_.Name; Path = $_.Path; SourcePath = $_.SourcePath }
    })
    return @{ Root = $root; Indexes = $sorted; Sources = $sources }
}


# インデックスのフォルダが壊れている（0 バイトのTSVがある）ことを表す件数。testIndexComplete は取り込み直す
${indexBrokenCount} = -1


function getContentIndexKeyOf {
    # getIndexTsvCounts の本文インデックスのファイルのキー（「<インデックスのフォルダからのフォルダ>\content_index.<拡張子>」）を返す
    param (
        [string]$relDir,
        [string]$extension
    )

    $name = "${contentIndexFileNamePrefix}.$extension"
    if ($relDir) { return "$relDir\$name" }
    return $name
}

function getIndexTsvCounts {
    # インデックスのフォルダの中のフォルダごとのTSVの数を返す（取り込み一覧の「済」と、インデックスの実体が合っているかの確認に使う）:
    #   インデックスのフォルダからの相対パス（大文字・小文字を区別しない）→ そのフォルダの直下のTSVの数
    #   本文インデックスのファイル（content_index.<拡張子>.<番号>.tsv）は、「<フォルダ>\content_index.<拡張子>」 → 1（どれかが 0 バイトなら ${indexBrokenCount}）
    # 元のファイル1つにつき1フォルダ（<ファイル名.xlsx>\<場所>.tsv）のため、キーは取り込み一覧の相対パスと同じになる。
    # 0 バイトのTSVがあるフォルダは ${indexBrokenCount}（-1）にする。
    # 空のシート・ページは保存しない（prettyTsv / writeUnits）ため、0 バイトのTSVは書き込みの途中で
    # 電源が落ちた場合などに限られ、そのままでは検索しても中身が出てこない。
    # 列挙できないとき（アクセス権が無い等）は $null を返す（呼び出し元は確認しない）
    param (
        [string]$dir = $workspace.IndexDir
    )

    $counts = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([System.StringComparer]::OrdinalIgnoreCase)
    $root = toLongPath ([string]$dir).TrimEnd("\")
    if (!$root -or ![System.IO.Directory]::Exists($root)) {
        return , $counts
    }

    # 1ファイルずつ調べると件数の分だけ時間がかかるため、フォルダ・TSVをそれぞれ1回ずつ列挙して数える
    # （大きさは列挙のときに分かるため、1件ずつ調べる必要は無い）
    $prefix = $root.Length + 1
    try {
        # 内容が空のファイル（TSV 0 件）はフォルダだけが残るため、フォルダも 0 件として数える
        foreach ($sub in [System.IO.Directory]::EnumerateDirectories($root, "*", [System.IO.SearchOption]::AllDirectories)) {
            $counts[$sub.Substring($prefix)] = 0
        }
        foreach ($file in (New-Object System.IO.DirectoryInfo($root)).EnumerateFiles("*.tsv", [System.IO.SearchOption]::AllDirectories)) {
            $contentIndexInfo = readContentIndexFileName $file.Name
            if ($null -ne $contentIndexInfo) {
                # 本文インデックスのファイル（content_index.<拡張子>.<番号>.tsv）は、番号を除いた「<フォルダ>\content_index.<拡張子>」をキーにする
                # （testIndexComplete が拡張子ごとに見る）。どれか 1 つでも 0 バイトなら壊れているとする
                $contentIndexDir = if ($file.DirectoryName.Length -ge $prefix) { $file.DirectoryName.Substring($prefix) } else { "" }
                $key = getContentIndexKeyOf $contentIndexDir $contentIndexInfo.Extension
                $current = 0
                if ($file.Length -eq 0 -or ($counts.TryGetValue($key, [ref]$current) -and $current -eq ${indexBrokenCount})) {
                    $counts[$key] = ${indexBrokenCount}
                } else {
                    $counts[$key] = 1
                }
                continue
            }
            $parent = [System.IO.Path]::GetDirectoryName($file.FullName)
            if ($parent.Length -lt $prefix) {
                continue  # インデックスのフォルダの直下のTSVは、どのファイルのものか分からないため数えない
            }
            $key = $parent.Substring($prefix)
            $count = 0
            if (!$counts.TryGetValue($key, [ref]$count)) {
                $count = 0
            }
            if ($file.Length -eq 0 -or $count -eq ${indexBrokenCount}) {
                $counts[$key] = ${indexBrokenCount}
            } else {
                $counts[$key] = $count + 1
            }
        }
    } catch {
        return $null
    }
    return , $counts
}

function testIndexComplete {
    # 取り込み一覧の行（状態が「済」）に対して、インデックスの実体がそろっているかを返す。
    # 利用者が work\content_index のフォルダ・ファイルを直接削除した場合に、「済」のまま検索できなくなるのを防ぐ。
    # 元のファイルの中身は、フォルダの本文インデックスのファイル（content_index.<拡張子>.tsv）か、本文インデックスのファイルに入れる前の TSV
    # （<ファイル名.xlsx>\<場所>.tsv。インデックス作成が途中で止まったとき）のどちらかにある
    #   row    : 取り込み一覧の行（TSV数 を使う）
    #   relPath: 取り込み一覧の相対パス（= インデックスのフォルダからの相対パス）
    #   counts : getIndexTsvCounts の結果（$null なら確認せず、そろっているものとして扱う）
    param (
        $row,
        [string]$relPath,
        $counts
    )

    if ($null -eq $counts -or $null -eq $row) {
        return $true
    }
    # 本文インデックスのファイル: そのフォルダ・その拡張子の本文インデックスのファイルがあれば、そろっているものとする
    # （本文インデックスのファイルの中の元のファイルごとの場所の数は、中身を読まないと分からないため数えない）
    $contentIndexRel = getContentIndexKeyOf ([string][System.IO.Path]::GetDirectoryName($relPath)) (getContentIndexExtension $relPath)
    $contentIndexCount = 0
    if ($counts.TryGetValue($contentIndexRel, [ref]$contentIndexCount) -and $contentIndexCount -ne ${indexBrokenCount}) {
        return $true
    }
    $expected = 0
    if (-not [int]::TryParse([string]$row.TSV数, [ref]$expected)) {
        return $true  # TSVの数を記録していない行は確認できない
    }
    $actual = 0
    if (-not $counts.TryGetValue($relPath, [ref]$actual)) {
        return $false  # 本文インデックスのファイルも、元のファイルごとのフォルダも無い
    }
    if ($actual -eq ${indexBrokenCount}) {
        return $false  # 0 バイトのTSVがある（書き込みの途中で電源が落ちた場合など）
    }
    # 余分なTSVがあっても（利用者が置いた等）取り込み直さない。足りない場合だけ作り直す
    return ($actual -ge $expected)
}

function publishIndexFiles {
    # 取り込んで作ったTSV（fromDir の直下）を、その元のファイルのインデックスのフォルダ（bookDir）に入れる。
    # 作りかけのインデックスを残さないよう、いったん stagingDir に集めてから bookDir ごと入れ替える。
    # 途中で強制終了されても、bookDir は「前回のまま」か「今回の分がそろった状態」のどちらかになる
    # （1件ずつ bookDir へ移すと、途中で止まったときに一部のシートだけのインデックスが残り、検索で気付けない）。
    # 以前のインデックスはフォルダごと置き換える（シートの削除・名前変更に追従するため）。
    # 消す・作る・移すのは、stagingDir は publishRoot の下、bookDir は indexRoot の下で、途中にリンクが無いときだけ
    # （確かめが偽なら何も動かさず、例外にする）
    param (
        [string]$fromDir,
        [string]$bookDir,
        [string]$stagingDir,
        [string]$indexRoot,
        [string]$publishRoot
    )

    $outside = "置き場の外か、途中にリンクがあるため、インデックスを置けません: "
    if (!(removeDirectoryUnder $stagingDir $publishRoot) -or !(newDirectoryUnder $stagingDir $publishRoot)) {
        throw "${outside}${stagingDir}"
    }
    foreach ($file in @(Get-ChildItem -LiteralPath (toLongPath $fromDir) -Filter "*.tsv" -File)) {
        [System.IO.File]::Move($file.FullName, (toLongPath (Join-Path $stagingDir $file.Name)))
    }

    if (!(removeDirectoryUnder $bookDir $indexRoot) -or !(newDirectoryUnder ([System.IO.Path]::GetDirectoryName($bookDir)) $indexRoot)) {
        throw "${outside}${bookDir}"
    }
    try {
        # 元と先で根が違う（publishRoot と indexRoot）ため moveDirectoryUnder でなく、それぞれの根で確かめてから移す
        if (!(testRemovablePath $stagingDir $publishRoot) -or !(testRemovablePath $bookDir $indexRoot)) {
            throw "${outside}${bookDir}"
        }
        [System.IO.Directory]::Move((toLongPath $stagingDir), (toLongPath $bookDir))
    } catch [System.IO.IOException] {
        # work を別のドライブへのリンクにしている場合など、フォルダごとは移せないときは1件ずつ移す
        if (!(newDirectoryUnder $bookDir $indexRoot)) {
            throw "${outside}${bookDir}"
        }
        foreach ($file in @(Get-ChildItem -LiteralPath (toLongPath $stagingDir) -Filter "*.tsv" -File)) {
            [System.IO.File]::Move($file.FullName, (toLongPath (Join-Path $bookDir $file.Name)))
        }
        [void](removeDirectoryUnder $stagingDir $publishRoot)
    }
}
