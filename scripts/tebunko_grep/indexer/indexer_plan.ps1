# クロール（どのファイルを取り込むかを数え、画面の返事を待つ）。

$targetExtensions = ${officeExtensions}  # 取り込み対象の拡張子（shared\office\office_files.ps1。画面のフォルダ選択でも同じ一覧を使う）

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

    return (Join-Path $indexDir $relPath)
}

function getIndexFiles {
    # そのファイルのインデックス（フォルダの中のTSV）を返す
    param (
        [string]$bookDir
    )

    # 長いパス（260文字超）でも見つかるよう \\?\ 付きで調べる（返すファイルの FullName も \\?\ 付き）
    if (!(Test-Path -LiteralPath (toLongPath $bookDir) -PathType Container)) {
        return @()
    }
    return @(Get-ChildItem -LiteralPath (toLongPath $bookDir) -Filter "*.tsv" -File)
}

function removeBookDir {
    # そのファイルのインデックスのフォルダを削除する（元のファイルが無くなったとき・取り込み直すとき）。
    # ウイルス対策ソフト・エクスプローラーが一時的に掴んでいることがあるため、少し待って数回試す
    param (
        [string]$bookDir
    )

    removeDirectoryRetry $bookDir
}

function findOfficeFiles {
    # クロール対象フォルダ配下のOfficeファイルを検索し、@{ Root; Files; HasError（アクセスできないフォルダがあった） } を返す。
    # Root は実際に列挙したフォルダ（\\?\ の付かない通常のパス）。相対パスはこの Root から求める
    # （設定に書かれたパスは、末尾の \ ・ドライブ文字と UNC パスなど書き方が違うことがあるため、文字数で切り出さない）
    param (
        [string]$targetFolder
    )

    # \\?\ を付けないと、パスが約248文字を超えるフォルダの中を検索できない（アクセスできないフォルダ扱いになる）。
    # 見つかったファイルの FullName は \\?\ 付きになる（fromLongPath で戻す）
    $scanErrors = $null
    $root = (Resolve-Path -LiteralPath $targetFolder).ProviderPath
    $files = @(Get-ChildItem -LiteralPath (toLongPath $root) -Recurse -File -ErrorAction SilentlyContinue -ErrorVariable scanErrors |
        Where-Object { ($targetExtensions -contains $_.Extension.ToLower()) -and -not $_.Name.StartsWith('~$') })

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
    # ・一覧に無いファイル（初回など）は、インデックス（TSV）が元ファイルより新しければ取り込み済みとする
    # ・元ファイルが無くなったファイルは、インデックスを削除して一覧から除く（アクセスできないフォルダがあった場合は除かない）
    param (
        $folder,   # @{ Path; Name }
        $previous, # readStatusFile の Rows（相対パス → 行）
        $counts    # getIndexTsvCounts の結果（インデックスの実体。$null なら確認しない）
    )

    $prefix = "$($folder.Name)\"
    $scan = findOfficeFiles $folder.Path
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
            $latest = $null
            $indexFiles = @()
            # インデックスのフォルダが無いことが counts で分かっていれば、ディスクを調べない（初回は全ファイルが一覧に無いため）
            if ($null -eq $old -and ($null -eq $counts -or $counts.ContainsKey($relPath))) {
                $indexFiles = getIndexFiles (getBookDir $relPath)
                foreach ($indexFile in $indexFiles) {
                    if ($null -eq $latest -or $indexFile.LastWriteTime -ge $latest.LastWriteTime) {
                        $latest = $indexFile
                    }
                }
            }
            if ($latest -and $latest.LastWriteTime -ge $file.LastWriteTime) {
                $tsvCount = $indexFiles.Count
                $row = newStatusRow $relPath $updated $size ${stateDone} $tsvCount (formatFileTime $latest.LastWriteTime)
                $count.Done++
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
    Write-Host ("  [{0}] Officeファイル {1} 件（{2}）" -f $folder.Name, $scan.Files.Count, $detail)
    if ($count.Lost -gt 0) {
        Write-Host "    インデックス（TSV）が無くなった・壊れている $($count.Lost) 件は取り込み直します。（インデックスを直接削除した・0 バイトのTSVが残っている）" -ForegroundColor Yellow
    }
    if ($removed.Count -gt 0) {
        Write-Host "    元ファイルが無くなった $($removed.Count) 件は、インデックスから除きます。"
    }
    if ($scan.HasError) {
        Write-Host "    アクセスできないフォルダがあったため、元ファイルが無くなったかどうかの確認は行いませんでした。" -ForegroundColor Yellow
    }

    $plan = newIngestPlanRow $folder.Name $folder.Path ${planKindIngest} $scan.Files.Count $targets.Count `
        $count.New $count.Updated $count.Pending $count.Lost $failed.Count
    return @{ Rows = $rows; Targets = $targets; Failed = $failed; Plan = $plan; Removed = $removed.ToArray() }
}

function waitForIndexingApproval {
    # 取り込み対象の件数を画面に渡し（取り込み予定.tsv）、［インデックス作成を開始］（インデックス作成開始要求）か［キャンセル］（インデックス作成中止要求）の返事を待つ。
    #   取り込む → @{ RetryFailed } / 取りやめ → $null
    # 画面を閉じた・落ちた場合に待ち続けないよう、$approvalTimeoutMinutes で打ち切って取りやめる
    param (
        $plan,               # newIngestPlanRow の配列（インデックスごと）
        [int]$targetCount,   # 取り込み対象の合計（画面の進み具合に出す）
        [int]$failedCount    # 前回失敗の合計（画面で再取り込みするかを選ぶ）
    )

    removeIndexingStartRequest
    writeIngestPlan $plan
    writeIndexingProgress ${indexingPhaseConfirm} 0 $targetCount $failedCount "取り込む内容を画面で確認しています…"
    Write-Host ""
    Write-Host "取り込み対象を画面に表示しました。［インデックス作成を開始］が押されるまで待ちます。（${approvalTimeoutMinutes} 分待っても返事が無ければ取りやめます）"

    $limit = (Get-Date).AddMinutes($approvalTimeoutMinutes)
    while ($true) {
        if (Test-Path -LiteralPath ${stopRequestFile}) {
            # 画面で［キャンセル］［中止］を押した
            Remove-Item -LiteralPath ${stopRequestFile} -Force
            removeIngestPlan
            return $null
        }
        $answer = readIndexingStartRequest
        if ($answer) {
            removeIndexingStartRequest
            removeIngestPlan
            return $answer
        }
        if ((Get-Date) -gt $limit) {
            Write-Host "画面からの返事が ${approvalTimeoutMinutes} 分ありませんでした。インデックス作成を取りやめます。" -ForegroundColor Yellow
            removeIngestPlan
            return $null
        }
        Start-Sleep -Milliseconds 300
    }
}
