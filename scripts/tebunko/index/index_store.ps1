# インデックス（work\index 配下）の作成・集計・改名・削除。

function getIndexNameMap {
    # 取り込み一覧に記録したインデックス名 → クロール対象フォルダのパス（大文字・小文字を区別しない）。
    # クロール対象フォルダの行は先頭にあるため、見出し行まで読んで打ち切る（取り込み一覧が大きくても時間がかからないように）
    param (
        [string]$path = $workspace.StatusFile
    )

    $map = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
    if (!(Test-Path -LiteralPath $path)) {
        return , $map
    }
    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $stream = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
    $reader = New-Object System.IO.StreamReader($stream, ${utf8Bom})
    try {
        $header = ${statusColumns} -join "`t"
        while ($null -ne ($line = $reader.ReadLine())) {
            if ($line -eq $header) {
                break
            }
            $fields = $line.Split("`t")
            if ($fields[0] -eq ${statusFolderKey} -and $fields.Count -eq 3 -and $fields[2]) {
                $map[$fields[2]] = $fields[1]
            }
        }
    } finally {
        $reader.Dispose()
    }
    return , $map
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
        } elseif ($state -eq ${stateNew}) {
            $stat.Pending++
        } elseif ($state -eq ${stateFailed}) {
            $stat.Failed++
        }
        # 取り込み日時は "yyyy/MM/dd HH:mm:ss" のため、文字列のまま比べて新しい方を残せる
        $ingested = [string]$entry.Value.取り込み日時
        if ($ingested -gt $stat.LastIngested) {
            $stat.LastIngested = $ingested
        }
    }
    return , $stats
}

function renameIndex {
    # インデックス名を変える（画面の［編集…］）。インデックスのフォルダ（work\index\<名前>）を改名し、
    # 取り込み一覧の記録も書き換えるため、名前を変えてもインデックスは作り直さない。
    # インデックス作成中は呼ばない（画面はインデックス作成中この操作を無効にする）
    param (
        [string]$oldName,
        [string]$newName,
        [string]$dir = $workspace.IndexDir,
        [string]$statusPath = $workspace.StatusFile
    )

    # 大文字・小文字だけの変更も改名するため、同じ名前かは大文字・小文字を区別して比べる（-ceq）
    if ($oldName -eq "" -or $newName -eq "" -or $oldName -ceq $newName) {
        return
    }

    $from = Join-Path $dir $oldName
    $to   = Join-Path $dir $newName
    if (Test-Path -LiteralPath $from -PathType Container) {
        if ([string]::Equals($oldName, $newName, [System.StringComparison]::OrdinalIgnoreCase)) {
            # 大文字・小文字だけを変える場合は、そのままでは改名できないため一時名を経由する
            $tmp = Join-Path $dir "${oldName}_rename_${PID}"
            [System.IO.Directory]::Move((toLongPath $from), (toLongPath $tmp))
            [System.IO.Directory]::Move((toLongPath $tmp), (toLongPath $to))
        } else {
            if (Test-Path -LiteralPath $to) {
                throw "「${newName}」のフォルダが既にあるため、名前を変えられません: ${to}"
            }
            [System.IO.Directory]::Move((toLongPath $from), (toLongPath $to))
        }
    }
    renameStatusIndexName $oldName $newName $statusPath
    # 高速検索の システムインデックス（同じワークスペースの system_index）は古い名前の分を消す。次のインデックス作成で作り直す
    removeSystemIndexOfWorkspace $oldName $dir
}

function removeIndex {
    # インデックスを削除する（画面の［削除］）。インデックスのフォルダ（work\index\<名前>）と、取り込み一覧の記録を削除する。
    # インデックス作成中は呼ばない（画面はインデックス作成中この操作を無効にする）
    param (
        [string]$name,
        [string]$dir = $workspace.IndexDir,
        [string]$statusPath = $workspace.StatusFile
    )

    if ($name -eq "") {
        return
    }
    $target = Join-Path $dir $name
    if (Test-Path -LiteralPath $target -PathType Container) {
        # 中に長いパス（260文字超）のTSVがあっても削除できるよう \\?\ 付きで削除する。
        # 画面は別スレッド（$ErrorActionPreference が既定の Continue）で呼ぶため、消せなければ例外にして
        # 取り込み一覧の記録を残す（-ErrorAction Stop が無いと、フォルダが残ったまま記録だけ消える）
        Remove-Item -LiteralPath (toLongPath $target) -Recurse -Force -ErrorAction Stop
    }
    removeStatusIndexName $name $statusPath
    removeSystemIndexOfWorkspace $name $dir
}

function removeSystemIndexOfWorkspace {
    # インデックスのフォルダ（dir = <ワークスペース>\index）と同じワークスペースの system_index から、インデックス name の分を消す。
    # 消せなくても（状態ファイルがほかに開かれている等）インデックスの操作は続ける。次のインデックス作成で整理される
    param (
        [string]$name,
        [string]$dir
    )

    $owner = [Workspace]::new([System.IO.Path]::GetDirectoryName($dir.TrimEnd("\")))
    try {
        [void](removeSystemIndexOf $name $owner.SystemIndexDir $owner.SystemIndexStateFile)
    } catch {
    }
}

function getSearchIndexes {
    # インデックスの一覧（work\index 直下のフォルダ 1 つがインデックス 1 つ）を
    # @{ Name（インデックス名）; Path（インデックスのフォルダのフルパス）; SourcePath（元のフォルダ。分からなければ ""） } の配列で返す。
    # 並びは［1 インデックス管理］の一覧（targetFolders）と同じにし、その一覧に無いもの
    # （別の場所・PC から work\index にコピーしたインデックスなど）は名前順で後ろに付ける
    param (
        [string]$dir = $workspace.IndexDir,
        [string]$statusPath = $workspace.StatusFile,
        [string]$settingsPath = ${settingsFile}
    )

    if (!(Test-Path -LiteralPath $dir -PathType Container)) {
        return @()
    }
    $root = (Resolve-Path -LiteralPath $dir).ProviderPath.TrimEnd("\")
    $sources = getSourceFolderMap $root $statusPath $settingsPath

    # ［1 インデックス管理］の一覧の順番（インデックス名 → 何番目か）
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
            # そのフォルダの中の 元のフォルダ.txt から元のフォルダを読む
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
    return @($indexes | Sort-Object Order, Name | ForEach-Object {
        [pscustomobject]@{ Name = $_.Name; Path = $_.Path; SourcePath = $_.SourcePath }
    })
}


# インデックスのフォルダが壊れている（0 バイトのTSVがある）ことを表す件数。testIndexComplete は取り込み直す
${indexBrokenCount} = -1


function getPackKeyOf {
    # getIndexTsvCounts の集約ファイルのキー（「<インデックスのフォルダからのフォルダ>\content.<拡張子>」）を返す
    param (
        [string]$relDir,
        [string]$extension
    )

    $name = "content.$extension"
    if ($relDir) { return "$relDir\$name" }
    return $name
}

function getIndexTsvCounts {
    # インデックスのフォルダの中のフォルダごとのTSVの数を返す（取り込み一覧の「済」と、インデックスの実体が合っているかの確認に使う）:
    #   インデックスのフォルダからの相対パス（大文字・小文字を区別しない）→ そのフォルダの直下のTSVの数
    #   集約ファイル（content.<拡張子>.<番号>.tsv）は、「<フォルダ>\content.<拡張子>」 → 1（どれかが 0 バイトなら ${indexBrokenCount}）
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
            $packInfo = readPackFileName $file.Name
            if ($null -ne $packInfo) {
                # 集約ファイル（content.<拡張子>.<番号>.tsv）は、番号を除いた「<フォルダ>\content.<拡張子>」をキーにする
                # （testIndexComplete が拡張子ごとに見る）。どれか 1 つでも 0 バイトなら壊れているとする
                $packDir = if ($file.DirectoryName.Length -ge $prefix) { $file.DirectoryName.Substring($prefix) } else { "" }
                $key = getPackKeyOf $packDir $packInfo.Extension
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
    # 利用者が work\index のフォルダ・ファイルを直接削除した場合に、「済」のまま検索できなくなるのを防ぐ。
    # 元のファイルの中身は、フォルダの集約ファイル（content.<拡張子>.tsv）か、集約ファイルに入れる前の TSV
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
    # 集約ファイル: そのフォルダ・その拡張子の集約ファイルがあれば、そろっているものとする
    # （集約ファイルの中の元のファイルごとの場所の数は、中身を読まないと分からないため数えない）
    $packRel = getPackKeyOf ([string][System.IO.Path]::GetDirectoryName($relPath)) (getPackExtension $relPath)
    $packCount = 0
    if ($counts.TryGetValue($packRel, [ref]$packCount) -and $packCount -ne ${indexBrokenCount}) {
        return $true
    }
    $expected = 0
    if (-not [int]::TryParse([string]$row.TSV数, [ref]$expected)) {
        return $true  # TSVの数を記録していない行は確認できない
    }
    $actual = 0
    if (-not $counts.TryGetValue($relPath, [ref]$actual)) {
        return $false  # 集約ファイルも、元のファイルごとのフォルダも無い
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
    # 以前のインデックスはフォルダごと置き換える（シートの削除・名前変更に追従するため）
    param (
        [string]$fromDir,
        [string]$bookDir,
        [string]$stagingDir
    )

    removeDirectoryRetry $stagingDir
    [System.IO.Directory]::CreateDirectory((toLongPath $stagingDir)) | Out-Null
    foreach ($file in @(Get-ChildItem -LiteralPath (toLongPath $fromDir) -Filter "*.tsv" -File)) {
        [System.IO.File]::Move($file.FullName, (toLongPath (Join-Path $stagingDir $file.Name)))
    }

    removeDirectoryRetry $bookDir
    [System.IO.Directory]::CreateDirectory((toLongPath ([System.IO.Path]::GetDirectoryName($bookDir)))) | Out-Null
    try {
        [System.IO.Directory]::Move((toLongPath $stagingDir), (toLongPath $bookDir))
    } catch [System.IO.IOException] {
        # work を別のドライブへのリンクにしている場合など、フォルダごとは移せないときは1件ずつ移す
        [System.IO.Directory]::CreateDirectory((toLongPath $bookDir)) | Out-Null
        foreach ($file in @(Get-ChildItem -LiteralPath (toLongPath $stagingDir) -Filter "*.tsv" -File)) {
            [System.IO.File]::Move($file.FullName, (toLongPath (Join-Path $bookDir $file.Name)))
        }
        removeDirectoryRetry $stagingDir
    }
}
