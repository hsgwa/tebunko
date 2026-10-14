# クロール（どのファイルを取り込むかを数え、画面の返事を待つ）。

# 取り込み対象の拡張子（Office は shared\office\office_file_kind_view.ps1、テキストは shared\core\text_file.ps1、PDF は shared\core\pdf_file.ps1）
$targetExtensions = @(${officeExtensions}) + @(${textExtensions}) + @(${pdfExtensions})

# tebunko が作ったファイルを、名前だけで見分けるパターン（テキストの拡張子のものだけに当たる。どこにあっても外す）。
#   ・今の版の本文インデックス（content_index.xlsx.001.tsv 等。contentIndexFileNamePattern と同じ組み立て）
#   ・前の版（名前をそろえる前）の本文インデックスのファイル（content.xlsx.001.tsv 等）
#   ・システムインデックス（今の版 system_index*.txt・前の版 システムインデックス*.txt。search_gram.ps1 の systemIndexFileName・
#     workspace.ps1 の legacySystemIndexPattern と同じ組み立て）
${tebunkoOwnFileNamePatterns} = @(
    "${contentIndexFileNamePrefix}.*.tsv",
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

function selectOnlyNames {
    # 選んだインデックス名（onlyNames）を、名前を割り当て済みのクロール対象フォルダ（folders）に当てはめる。
    # 戻り値: @{ Selected（今回更新する名前の集合。大文字・小文字を区別しない）; Skipped（更新できない名前の @{ Name; Reason } の配列） }
    # 更新できないのは、設定に無い名前と、チェックが付いていない名前（チェックを付けるのは画面側）
    param (
        [object[]]$folders,
        [string[]]$onlyNames
    )

    $selected = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $skipped = New-Object System.Collections.Generic.List[object]
    foreach ($name in @($onlyNames | Where-Object { $_ })) {
        $folder = @($folders | Where-Object { $_.Name -eq $name }) | Select-Object -First 1
        if (!$folder) {
            $skipped.Add([pscustomobject]@{ Name = $name; Reason = "設定にありません" })
        } elseif (!$folder.Enabled) {
            $skipped.Add([pscustomobject]@{ Name = $name; Reason = "チェックが付いていません" })
        } else {
            [void]$selected.Add($name)
        }
    }
    return @{ Selected = $selected; Skipped = $skipped.ToArray() }
}

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
    # 相対パスは取り込み一覧（共有のワークスペースで書き換えられうる）から来るため、置き場（IndexDir）の下でなければ消さない
    # （例外にすると、書き換えられた 1 行で毎回インデックス作成が止まるため、ログに出して続ける）
    param (
        [string]$bookDir
    )

    if (!(testPathUnder $bookDir $workspace.IndexDir)) {
        writeIndexerLog "インデックスの置き場の外を指すパスのため、削除しませんでした: ${bookDir}" "Yellow"
        return
    }
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
    #   Cloud  : 取り込む（または前回失敗した）ファイルのうち、クラウドにだけあるもの（OneDrive 等のオンデマンド）。
    #            開くとダウンロードされるため Targets・Failed には入れず、呼び出し元が取り込むかを決める。
    #            各要素は @{ RelPath; Row（取り込むときの行）; Old; Failed; Length; PendingOld }。
    #            Rows には前回の行をそのまま入れる（前回の行が無い・前回未完了のときは入れない）
    #   Plan   : 画面の確認に出す件数（newIngestPlanRow。取り込み予定.tsv の1行）
    #   Removed: 元のファイルが無くなったファイルの相対パス（呼び出し元が本文インデックスのファイルから外す）
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
    $cloud = New-Object System.Collections.Generic.List[object]
    # 取り込み予定に出す、クラウドにだけあるファイルの件数と容量（取り込みの対象と、前回失敗を分けて、下のループで足す）
    $cloudCount = 0
    $cloudFailedCount = 0
    [long]$cloudBytes = 0
    [long]$cloudFailedBytes = 0
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

        # 読む必要があるファイル（取り込む・前回失敗）だけ、クラウドにだけあるかを調べる。
        # 属性は一覧が返したものを見るだけで、ファイルは開かない（ダウンロードは起きない）
        if (($decision.Ingest -or $decision.Reason -eq "failed") -and (testCloudOnlyAttributes ([int]$file.Attributes))) {
            $isFailed = (-not $decision.Ingest)
            $pendingOld = [bool]($old -and $old.状態 -eq ${stateNew})
            if ($isFailed) {
                $cloudFailedCount++
                $cloudFailedBytes += [long]$file.Length
            } else {
                $cloudCount++
                $cloudBytes += [long]$file.Length
            }
            $cloud.Add(@{
                RelPath    = $relPath
                Row        = $(if ($isFailed) { $old } else { newStatusRow $relPath $updated $size ${stateNew} })
                Old        = $old
                Failed     = $isFailed
                Length     = [long]$file.Length
                PendingOld = $pendingOld
            })
            # 取り込まないときは前回の行を残す（前回の行が無い・前回未完了なら、行を作らない）
            if ($old -and -not $pendingOld) {
                $rows.Add($old)
            }
            continue
        }

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
    if ($cloud.Count -gt 0) {
        $detail += " / クラウドにだけある $($cloud.Count) 件"
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
        $count.New $count.Updated $count.Pending $count.Lost $failed.Count `
        $cloudCount $cloudFailedCount $cloudBytes $cloudFailedBytes
    return @{ Rows = $rows; Targets = $targets; Failed = $failed; Cloud = $cloud; Plan = $plan; Removed = $removed.ToArray() }
}

# ----------------------------------------------------------------------------
# 取り込む順番を決める（IngestPlanner）
# ----------------------------------------------------------------------------

# インデックス作成の司令（invokeIndexerBody）の前半。クロール対象フォルダにインデックス名を割り当て、クロールして、
# 取り込む順番（Targets）・画面に出す件数（Plan）・取り込み一覧に書く行（Rows）を組み立てる。
# 取り込み直すかの判断は indexer_decide.ps1（判断層）の関数を、createTargetList 経由でそのまま呼ぶ。
# 進み具合・画面の確認・取り込み一覧の読み書きは持たない（IndexingReporter・StatusLedger を司令がつなぐ）。
# クラスの本体では $script: 変数を読まない決まり（docs/design/structure/classes.md。テストと本番で読める範囲が違うため）なので、
# 計画の種類・状態の値は kinds（Unchecked・Missing・Dropped・StateFailed の 4 つ。足りなければ例外）、
# 名前の保存は、司令が渡す save（割り当てた名前を設定に保存し、保存した一覧を返すスクリプトブロック。IndexCatalog.StoreAssignedNames）に任せる。
# メソッドは次の順に呼ぶ: AssignNames → SaveNames → SelectOnly（選んだときだけ）→ Crawl →
#   KeepForCancel（取りやめるとき）か ApplyFailed → ApplyInterrupted → FolderByName。Crawl は AssignNames より前なら例外にする。
# メソッドの中のローカル変数は、プロパティと大文字・小文字だけの違いにしない
class IngestPlanner {
    # 名前を割り当てたクロール対象フォルダ（@{ Path; Name; Enabled } の並び）
    [object[]]$Folders
    # 前回の取り込み一覧の行（相対パス → 行）
    $Previous
    # 設定から外れたインデックス（findDroppedIndexes の結果）
    [object[]]$Dropped
    # 選んだインデックスだけを更新するときの当てはめ（selectOnlyNames の結果。選ばなければ $null）
    $Only
    # 全ファイルの行・取り込む行・前回失敗し更新の無い行・クラウドにだけあるファイル・画面の確認に出す件数
    [System.Collections.Generic.List[object]]$Rows
    [System.Collections.Generic.List[object]]$Targets
    [System.Collections.Generic.List[object]]$Failed
    # クラウドにだけあるファイル（createTargetList の Cloud。取り込むかは確認の後に司令が決める）
    [System.Collections.Generic.List[object]]$Cloud
    [System.Collections.Generic.List[object]]$Plan
    # 元のファイルが無くなったファイルの相対パス（司令が本文インデックスのファイルから外す）
    [System.Collections.Generic.List[string]]$RemovedPaths
    # 前回、取り込み中に強制終了して最後に回したファイル（相対パス → 回数）
    $Carried
    hidden $Kinds

    IngestPlanner([hashtable]$kinds) {
        foreach ($key in @("Unchecked", "Missing", "Dropped", "StateFailed")) {
            if (-not $kinds.ContainsKey($key) -or $null -eq $kinds[$key]) {
                throw "IngestPlanner: kinds に $key がありません"
            }
        }
        $this.Kinds = $kinds
        $this.Rows = New-Object System.Collections.Generic.List[object]
        $this.Targets = New-Object System.Collections.Generic.List[object]
        $this.Failed = New-Object System.Collections.Generic.List[object]
        $this.Cloud = New-Object System.Collections.Generic.List[object]
        $this.Plan = New-Object System.Collections.Generic.List[object]
        $this.RemovedPaths = New-Object System.Collections.Generic.List[string]
        $this.Carried = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([System.StringComparer]::OrdinalIgnoreCase)
    }

    [void] AssignNames($targetFolders, $status) {
        # クロール対象フォルダごとにインデックス名（work\content_index 直下のフォルダ名）を決める。前回と同じフォルダは同じ名前を使う。
        # status: 取り込み一覧（StatusLedger の ReadStatus。@{ Folders; Rows }）。
        # 設定から外れたインデックスは、ここでは消さない（確認ダイアログに「削除予定」として出す）
        $this.Folders = @(assignIndexNames $targetFolders $status.Folders)
        $this.Previous = $status.Rows
        # 名前が今の決まりに合わない前の版のインデックス（status.LegacyFolders）も、設定に無いので外れたインデックスとして出す
        $previousFolders = New-Object System.Collections.Generic.List[object]
        foreach ($folder in $status.Folders) { $previousFolders.Add($folder) }
        foreach ($folder in $status.LegacyFolders) { if ($folder) { $previousFolders.Add($folder) } }
        $this.Dropped = @(findDroppedIndexes $this.Folders $previousFolders.ToArray())
    }

    [void] SaveNames($targetFolders, [scriptblock]$save) {
        # 新しく割り当てたインデックス名を設定に保存する（インデックスの「名前」と「置き場所」を設定で分けて持つため。
        # 名前が設定にあれば、フォルダを移してパスを書き換えても同じインデックスとして扱える）
        if (@($targetFolders | Where-Object { -not $_.Name }).Count -eq 0) {
            return
        }
        # 保存した一覧（読み直して名前だけを足したもの）の名前で続ける。この間に画面で外した・名前を付けられなかったフォルダは、今回は取り込まない
        $saved = @(& $save $this.Folders)
        $this.Folders = @($saved | Where-Object {
            $savedPath = $_.Path
            $_.Name -and @($targetFolders | Where-Object { $_.Path -eq $savedPath }).Count -gt 0
        })
        writeIndexerLog "インデックス名を設定に保存しました: $((@($this.Folders | ForEach-Object { $_.Name }) -join '、'))"
    }

    [object[]] SelectOnly([string[]]$onlyNames) {
        # 選んだインデックスだけを更新するとき。選ばなかったインデックスは、この回では一切触らない
        # （取り込み一覧の行は前回のまま・確認の表に出さない・失敗の一覧に加えない）。
        # 更新できない名前（@{ Name; Reason } の並び）を返す。更新できるものが 1 つも無いかは Only.Selected.Count で見る
        $this.Only = selectOnlyNames $this.Folders $onlyNames
        foreach ($skip in $this.Only.Skipped) {
            writeIndexerLog "  [$($skip.Name)] 更新できません（$($skip.Reason)）" "Yellow"
        }
        return @($this.Only.Skipped)
    }

    [void] Crawl($indexCounts, [scriptblock]$onFolder) {
        # クロール対象フォルダを順に検索して、Rows・Targets・Failed・Plan・RemovedPaths を組み立てる。
        # onFolder: フォルダを検索する直前に呼ぶ（司令が進み具合を画面に出す。$null なら何もしない）
        if ($null -eq $this.Folders) {
            throw "IngestPlanner: Crawl の前に AssignNames を呼んでください"
        }
        foreach ($folder in $this.Folders) {
            if ($null -ne $this.Only -and !$this.Only.Selected.Contains([string]$folder.Name)) {
                # 選ばなかったインデックス（下で前回の結果をそのまま残す）
            } elseif (-not $folder.Enabled) {
                writeIndexerLog "  [$($folder.Name)] $($folder.Path) … チェックなしのため取り込みません（インデックスはそのまま残します）"
                $this.Plan.Add((newIngestPlanRow $folder.Name $folder.Path $this.Kinds.Unchecked))
            } elseif (!(Test-Path -LiteralPath $folder.Path -PathType Container)) {
                writeIndexerLog "  [$($folder.Name)] $($folder.Path) … フォルダが見つからないため取り込みません" "Yellow"
                $this.Plan.Add((newIngestPlanRow $folder.Name $folder.Path $this.Kinds.Missing))
            } else {
                writeIndexerLog "  [$($folder.Name)] $($folder.Path)"
                if ($null -ne $onFolder) {
                    $null = & $onFolder $folder
                }
                $list = createTargetList $folder $this.Previous $indexCounts
                $this.Rows.AddRange($list.Rows)
                $this.Targets.AddRange($list.Targets)
                $this.Failed.AddRange($list.Failed)
                $this.Cloud.AddRange($list.Cloud)
                $this.Plan.Add($list.Plan)
                $this.RemovedPaths.AddRange([string[]]@($list.Removed))
                continue
            }

            # 取り込まないフォルダは前回の結果をそのまま残す
            $prefix = "$($folder.Name)\"
            foreach ($key in @($this.Previous.Keys)) {
                if ($key.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $this.Rows.Add($this.Previous[$key])
                }
            }
        }

        foreach ($item in $this.Dropped) {
            $this.Plan.Add((newIngestPlanRow $item.Name $item.Path $this.Kinds.Dropped))
        }
    }

    [object[]] KeepForCancel() {
        # 画面で取りやめたときに書く取り込み一覧の行。1件も取り込んでいないため、取り込み対象にした行は前回の記録のまま
        # （一覧に無かったファイルは記録しない）にする。「未取り込み」で記録すると、次回［すべて更新］が［続きから再開］になり、
        # 中断したように見えるため
        $targetPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($row in $this.Targets) {
            [void]$targetPaths.Add($row.相対パス)
        }
        $keep = New-Object System.Collections.Generic.List[object]
        # クラウドにだけあって前回未完了のファイルは、取り込み一覧に前回の行を残す（取りやめでは何も消さない）
        foreach ($item in $this.Cloud) {
            if ($item.PendingOld) {
                $keep.Add($item.Old)
            }
        }
        foreach ($row in $this.Rows) {
            if (!$targetPaths.Contains($row.相対パス)) {
                $keep.Add($row)
                continue
            }
            $old = $null
            if ($this.Previous.TryGetValue($row.相対パス, [ref]$old)) {
                $keep.Add($old)
            }
        }
        return $keep.ToArray()
    }

    [void] ApplyFailed([bool]$retry) {
        # 前回取り込みに失敗し、その後更新されていないファイルを、取り込み直す（retry）か、スキップするかを決める
        if ($this.Failed.Count -eq 0) {
            return
        }
        writeIndexerLog ""
        if ($retry) {
            writeIndexerLog "前回取り込みに失敗し、その後更新されていないファイル $($this.Failed.Count) 件も再取り込みします。"
            $this.Targets.AddRange($this.Failed)
        } else {
            writeIndexerLog "前回取り込みに失敗し、その後更新されていないファイル $($this.Failed.Count) 件はスキップします。（パスワード付きなど）"
        }
    }

    [void] ApplyInterrupted([object[]]$entries, [int]$interruptLimit) {
        # 前回、取り込み中に強制終了した（ウィンドウを閉じた・PCが停止した等）ファイルは、同じファイルで止まり続けないよう最後に回す。
        # 続けて interruptLimit 回強制終了したファイルは、応答しなくなるファイルとみなして失敗とする。
        # 最後に回したファイルは Carried に残す（取り込みを始めるまで記録を残し、その前にまた止まっても回数が分かるように）。
        # entries: StatusLedger の ReadIngestingFiles（@{ RelPath; Count } の並び）
        foreach ($entry in $entries) {
            $hit = @($this.Targets | Where-Object { $_.相対パス -eq $entry.RelPath }) | Select-Object -First 1
            if (!$hit) {
                continue
            }
            [void]$this.Targets.Remove($hit)
            if ($entry.Count -ge $interruptLimit) {
                $hit.状態 = $this.Kinds.StateFailed
                $hit.TSV数 = ""
                $hit.取り込み日時 = formatFileTime (Get-Date)
                $hit.エラー = "更新中に $($entry.Count) 回続けて強制終了されたため、更新を中止しました（Officeアプリが応答しなくなる可能性があります）"
                writeIndexerLog ""
                writeIndexerLog "取り込み中に $($entry.Count) 回続けて強制終了したファイルは、失敗としてスキップします: $($hit.相対パス)" "Yellow"
            } else {
                $this.Targets.Add($hit)
                $this.Carried[$hit.相対パス] = $entry.Count
                writeIndexerLog ""
                writeIndexerLog "前回、取り込み中に強制終了したファイルは最後に取り込みます: $($hit.相対パス)" "Yellow"
            }
        }
    }

    [object] FolderByName() {
        # インデックス名 → クロール対象フォルダ
        $byName = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($folder in $this.Folders) {
            $byName[$folder.Name] = $folder.Path.TrimEnd("\")
        }
        return $byName
    }
}
