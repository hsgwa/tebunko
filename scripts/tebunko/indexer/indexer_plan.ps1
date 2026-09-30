# クロール（どのファイルを取り込むかを数え、画面の返事を待つ）。

# 取り込み対象の拡張子（Office は shared\office\office_files.ps1、テキストは shared\core\text_file.ps1）
$targetExtensions = @(${officeExtensions}) + @(${textExtensions})

# tebunko が作ったファイルを、名前だけで見分けるパターン（テキストの拡張子のものだけに当たる。どこにあっても外す）。
#   ・今の版の本文インデックス（content_index.xlsx.001.tsv 等。packFileNamePattern と同じ組み立て）
#   ・前の版（名前をそろえる前）の集約ファイル（content.xlsx.001.tsv 等）
#   ・システムインデックス（今の版 system_index*.txt・前の版 システムインデックス*.txt。search_gram.ps1 の systemIndexFileName・
#     workspace.ps1 の legacySystemIndexPattern と同じ組み立て）
${tebunkoOwnFileNamePatterns} = @(
    "${packFileNamePrefix}.*.tsv",
    "content.*.tsv",
    "$([System.IO.Path]::GetFileNameWithoutExtension(${systemIndexFileName}))*.txt",
    ${legacySystemIndexPattern}
)

function testTebunkoOwnFileName {
    # ファイル名（拡張子を含む）が、上のパターンのどれかに当たるか
    param (
        [string]$name
    )

    foreach ($pattern in ${tebunkoOwnFileNamePatterns}) {
        if ($name -like $pattern) {
            return $true
        }
    }
    return $false
}

# ----------------------------------------------------------------------------
# 取り込み対象
# ----------------------------------------------------------------------------

function getBookDir {
    # 取り込み対象のファイルの相対パスから、そのファイルのインデックスを入れるフォルダを返す。
    # 相対パスの最後は元のファイル名のため、インデックスフォルダと相対パスをつなぐとフォルダ名になる
    #   例: "営業\2024\A社.xlsx" → "work\index\営業\2024\A社.xlsx"（この中に "明細.tsv" 等を入れる）
    param (
        [string]$relPath
    )

    return (Join-Path $workspace.IndexDir $relPath)
}

function removeBookDir {
    # そのファイルのインデックスのフォルダを削除する（元のファイルが無くなったとき・取り込み直すとき）。
    # ウイルス対策ソフト・エクスプローラーが一時的に掴んでいることがあるため、少し待って数回試す
    param (
        [string]$bookDir
    )

    removeDirectoryRetry $bookDir
}

function getTebunkoExcludeDirs {
    # 除外するフォルダ（tebunko が作ったワークスペース）の一覧を返す。
    #   ・今のワークスペース（$workspace.Entries()。content_index・前の版の index・system_index・取り込み一覧 等）
    #   ・スキャンで見つけたファイルの中に取り込み一覧（$workspace.StatusFile と同じ名前）があれば、
    #     そのフォルダをほかのワークスペースとみなし、その Entries() も外す
    #     （以前の既定の場所・切り替える前のワークスペース・ほかの人のワークスペースを、テキストの拡張子（.tsv 等）で拾わないため）
    param (
        [object[]]$scannedFiles
    )

    $statusFileName = [System.IO.Path]::GetFileName($workspace.StatusFile)
    $dirs = New-Object System.Collections.Generic.List[string]
    $dirs.AddRange([string[]]$workspace.Entries())
    foreach ($file in $scannedFiles) {
        if ($file.Name -eq $statusFileName) {
            $otherDir = [System.IO.Path]::GetDirectoryName((fromLongPath $file.FullName))
            $dirs.AddRange([string[]]([Workspace]::new($otherDir).Entries()))
        }
    }
    return @($dirs | Where-Object { $_ } | Select-Object -Unique)
}

function testUnderAnyDir {
    # \\?\ 付きのフルパス（またはそのもの）が、dirs（\\?\ の付かない通常のパス）のどれかの下・そのものかを返す。
    #   dirPrefixes: dirs を toLongPath して末尾の \ を外したもの（呼び出し側で 1 回だけ作り、ファイルごとに作り直さない）
    param (
        [string]$fullName,
        [string[]]$dirPrefixes
    )

    foreach ($prefix in $dirPrefixes) {
        if ($fullName.Equals($prefix, [System.StringComparison]::OrdinalIgnoreCase) -or
            $fullName.StartsWith("$prefix\", [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }
    return $false
}

function findTargetFiles {
    # クロール対象フォルダ配下の、取り込み対象（Office・テキストの拡張子）のファイルを検索し、
    # @{ Root; Files; HasError（アクセスできないフォルダがあった） } を返す。
    # Root は実際に列挙したフォルダ（\\?\ の付かない通常のパス）。相対パスはこの Root から求める
    # （設定に書かれたパスは、末尾の \ ・ドライブ文字と UNC パスなど書き方が違うことがあるため、文字数で切り出さない）。
    # tebunko が作ったファイル（getTebunkoExcludeDirs のフォルダの下・名前で分かるファイル）は対象に含めない
    param (
        [string]$targetFolder
    )

    # \\?\ を付けないと、パスが約248文字を超えるフォルダの中を検索できない（アクセスできないフォルダ扱いになる）。
    # 見つかったファイルの FullName は \\?\ 付きになる（fromLongPath で戻す）
    $scanErrors = $null
    $root = (Resolve-Path -LiteralPath $targetFolder).ProviderPath
    $scanned = @(Get-ChildItem -LiteralPath (toLongPath $root) -Recurse -File -ErrorAction SilentlyContinue -ErrorVariable scanErrors |
        Where-Object { ($targetExtensions -contains $_.Extension.ToLower()) -and -not $_.Name.StartsWith('~$') })

    # 除外するフォルダの \\?\ 付きの前方一致の文字列を先に作る（ファイルごとにドライブの割り当てをたどる処理は呼ばない）。
    # このクロール対象フォルダの外にあるものは、どのファイルにも当たらないため先に落とす（比べる数を減らす）
    $longRootPrefix = (toLongPath $root).TrimEnd("\")
    $excludeDirPrefixes = @(getTebunkoExcludeDirs $scanned | ForEach-Object { (toLongPath $_).TrimEnd("\") } | Where-Object {
        $_.Equals($longRootPrefix, [System.StringComparison]::OrdinalIgnoreCase) -or $_.StartsWith("$longRootPrefix\", [System.StringComparison]::OrdinalIgnoreCase)
    })
    $files = @($scanned | Where-Object {
        if ($excludeDirPrefixes.Count -gt 0 -and (testUnderAnyDir $_.FullName $excludeDirPrefixes)) { return $false }
        # 名前で分かる tebunko のファイル（content_index.*.tsv 等）は、テキストの拡張子にしか当たらないため、それだけ調べる
        if ((testTextExtension $_.Name) -and (testTebunkoOwnFileName $_.Name)) { return $false }
        return $true
    })

    return @{ Root = $root; Files = $files; HasError = (@($scanErrors).Count -gt 0) }
}

function createTargetList {
    # クロール対象フォルダを1つ検索して取り込み一覧の行を作り直し、次を返す。
    #   Rows   : 全ファイルの行 / Targets: 取り込む行 / Failed: 前回失敗し、更新の無い行
    #   Plan   : 画面の確認に出す件数（newIngestPlanRow。取り込み予定.tsv の1行）
    #   Removed: 元のファイルが無くなったファイルの相対パス（呼び出し元が集約ファイルから外す）
    # 行の相対パスは "インデックス名\フォルダからの相対パス"（= work\index からの相対パス）とする。
    # ・前回の一覧と更新日時・サイズが同じで取り込み済み（済）のファイルは取り込まない
    # ・取り込み済みでも、インデックス（TSV）が無くなっていれば取り込み直す（利用者が work\index を直接削除した場合など）
    # ・元ファイルが無くなったファイルは、インデックスを削除して一覧から除く（アクセスできないフォルダがあった場合は除かない）
    param (
        $folder,   # @{ Path; Name }
        $previous, # readStatusFile の Rows（相対パス → 行）
        $counts    # getIndexTsvCounts の結果（インデックスの実体。$null なら確認しない）
    )

    $prefix = "$($folder.Name)\"
    $scan = findTargetFiles $folder.Path
    $rows = New-Object System.Collections.Generic.List[object]
    $targets = New-Object System.Collections.Generic.List[object]
    $failed = New-Object System.Collections.Generic.List[object]
    $found = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $count = @{ Done = 0; New = 0; Updated = 0; Pending = 0; Lost = 0 }

    # 列挙したファイルの FullName は、\\?\ 付きの Root で始まる。ファイルが多いと1件ずつの関数呼び出しだけで
    # 時間がかかるため、その場合は先頭を切り落として相対パスにする（それ以外は getPathUnderFolder で求める）
    $longRoot = (toLongPath $scan.Root).TrimEnd("\") + "\"
    foreach ($file in $scan.Files) {
        $fullName = $file.FullName
        if ($fullName.Length -gt $longRoot.Length -and $fullName.StartsWith($longRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
            $relative = $fullName.Substring($longRoot.Length)
        } else {
            $relative = getPathUnderFolder (fromLongPath $fullName) $scan.Root
        }
        if ($null -eq $relative -or $relative -eq "") {
            $relative = $file.Name  # 通常は起こらない（$scan.Root の下を列挙している）
        }
        $relPath = $prefix + $relative
        [void]$found.Add($relPath)
        $updated = formatFileTime $file.LastWriteTime
        $size = [string]$file.Length

        $old = $null
        [void]$previous.TryGetValue($relPath, [ref]$old)

        # 取り込むかどうかの判断は indexer_decide.ps1（テストしやすいように分けてある）
        # インデックス（TSV）の有無は、前回と同じファイルで「済」のときだけ調べる（取り込み直すものには要らない）
        $indexComplete = $true
        if ($old -and $old.状態 -eq ${stateDone} -and $old.更新日時 -eq $updated -and $old.サイズ -eq $size) {
            $indexComplete = [bool](testIndexComplete $old $relPath $counts)
        }
        $decision = getIngestDecision $old $updated $size $indexComplete

        if (-not $decision.Ingest) {
            # 更新なし。失敗したファイルを再取り込みするかは呼び出し元で決める
            $row = $old
            if ($decision.Reason -eq "failed") {
                $failed.Add($row)
            } else {
                $count.Done++
            }
        } else {
            $row = newStatusRow $relPath $updated $size ${stateNew}
            $targets.Add($row)
            switch ($decision.Reason) {
                "lost"    { $count.Lost++ }
                "new"     { $count.New++ }
                "pending" { $count.Pending++ }
                default   { $count.Updated++ }  # updated と outdated（前の抽出版で取り込んだ。画面では更新ありと同じに扱う）
            }
        }
        $rows.Add($row)
    }

    $removed = New-Object System.Collections.Generic.List[string]
    foreach ($relPath in @($previous.Keys)) {
        if (!$relPath.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase) -or $found.Contains($relPath)) {
            continue
        }
        if ($scan.HasError) {
            $rows.Add($previous[$relPath])
            continue
        }
        removeBookDir (getBookDir $relPath)
        $removed.Add($relPath)
    }

    $detail = "取り込み済み {0} 件 / 新規 {1} 件 / 更新あり {2} 件 / 前回未完了 {3} 件 / 前回失敗 {4} 件" -f
        $count.Done, $count.New, $count.Updated, $count.Pending, $failed.Count
    if ($count.Lost -gt 0) {
        # インデックスを直接削除された場合など。ふだんは 0 件のため、あるときだけ表示する
        $detail += " / インデックスが無い・壊れている $($count.Lost) 件"
    }
    writeIndexerLog ("  [{0}] 対象ファイル {1} 件（{2}）" -f $folder.Name, $scan.Files.Count, $detail)
    if ($count.Lost -gt 0) {
        writeIndexerLog "    インデックス（TSV）が無くなった・壊れている $($count.Lost) 件は取り込み直します。（インデックスを直接削除した・0 バイトのTSVが残っている）" "Yellow"
    }
    if ($removed.Count -gt 0) {
        writeIndexerLog "    元ファイルが無くなった $($removed.Count) 件は、インデックスから除きます。"
    }
    if ($scan.HasError) {
        writeIndexerLog "    アクセスできないフォルダがあったため、元ファイルが無くなったかどうかの確認は行いませんでした。" "Yellow"
    }

    $plan = newIngestPlanRow $folder.Name $folder.Path ${planKindIngest} $scan.Files.Count $targets.Count `
        $count.New $count.Updated $count.Pending $count.Lost $failed.Count
    return @{ Rows = $rows; Targets = $targets; Failed = $failed; Plan = $plan; Removed = $removed.ToArray() }
}
