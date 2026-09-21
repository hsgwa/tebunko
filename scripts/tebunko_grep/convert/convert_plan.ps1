# 変換対象の決定（どのファイルを変換するかを数え、画面の返事を待つ）。

$targetExtensions = ${officeExtensions}  # 変換対象の拡張子（shared\office\office_files.ps1。画面のフォルダ選択でも同じ一覧を使う）

# ----------------------------------------------------------------------------
# 変換対象
# ----------------------------------------------------------------------------

function getBookDir {
    # 変換対象ファイルの相対パスから、そのファイルの変換結果を入れるフォルダを返す。
    # 相対パスの最後は元のファイル名のため、インデックスフォルダと相対パスをつなぐとフォルダ名になる
    #   例: "営業\2024\A社.xlsx" → "work\index\営業\2024\A社.xlsx"（この中に "明細.tsv" 等を入れる）
    param (
        [string]$relPath
    )

    return (Join-Path $indexDir $relPath)
}

function getIndexFiles {
    # そのファイルの変換結果（フォルダの中のTSV）を返す
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
    # そのファイルの変換結果のフォルダを削除する（元のファイルが無くなったとき・変換し直すとき）。
    # ウイルス対策ソフト・エクスプローラーが一時的に掴んでいることがあるため、少し待って数回試す
    param (
        [string]$bookDir
    )

    removeDirectoryRetry $bookDir
}

function findOfficeFiles {
    # 変換対象フォルダ配下のOfficeファイルを検索し、@{ Root; Files; HasError（アクセスできないフォルダがあった） } を返す。
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
    # 変換対象フォルダを1つ検索して変換一覧の行を作り直し、次を返す。
    #   Rows   : 全ファイルの行 / Targets: 変換する行 / Failed: 前回失敗し、更新の無い行
    #   Plan   : 画面の確認に出す件数（newConvertPlanRow。変換予定.tsv の1行）
    # 行の相対パスは "インデックス名\フォルダからの相対パス"（= work\index からの相対パス）とする。
    # ・前回の一覧と更新日時・サイズが同じで変換済み（済）のファイルは変換しない
    # ・変換済みでも、インデックス（TSV）が無くなっていれば変換し直す（利用者が work\index を直接削除した場合など）
    # ・一覧に無いファイル（初回など）は、変換結果（TSV）が元ファイルより新しければ変換済みとする
    # ・元ファイルが無くなったファイルは、変換結果を削除して一覧から除く（アクセスできないフォルダがあった場合は除かない）
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

    foreach ($file in $scan.Files) {
        $relative = getPathUnderFolder (fromLongPath $file.FullName) $scan.Root
        if ($null -eq $relative -or $relative -eq "") {
            $relative = $file.Name  # 通常は起こらない（$scan.Root の下を列挙している）
        }
        $relPath = $prefix + $relative
        [void]$found.Add($relPath)
        $updated = formatFileTime $file.LastWriteTime
        $size = [string]$file.Length

        $old = $null
        [void]$previous.TryGetValue($relPath, [ref]$old)

        # 変換するかどうかの判断は convert_decide.ps1（テストしやすいように分けてある）
        # 変換結果（TSV）の有無は、前回と同じファイルで「済」のときだけ調べる（変換し直すものには要らない）
        $indexComplete = $true
        if ($old -and $old.状態 -eq ${stateDone} -and $old.更新日時 -eq $updated -and $old.サイズ -eq $size) {
            $indexComplete = [bool](testIndexComplete $old $relPath $counts)
        }
        $decision = getConvertDecision $old $updated $size $indexComplete

        if (-not $decision.Convert) {
            # 更新なし。失敗したファイルを再変換するかは呼び出し元で決める
            $row = $old
            if ($decision.Reason -eq "failed") {
                $failed.Add($row)
            } else {
                $count.Done++
            }
        } else {
            $latest = $null
            if ($null -eq $old) {
                $latest = getIndexFiles (getBookDir $relPath) | Sort-Object LastWriteTime | Select-Object -Last 1
            }
            if ($latest -and $latest.LastWriteTime -ge $file.LastWriteTime) {
                $tsvCount = @(getIndexFiles (getBookDir $relPath)).Count
                $row = newStatusRow $relPath $updated $size ${stateDone} $tsvCount (formatFileTime $latest.LastWriteTime)
                $count.Done++
            } else {
                $row = newStatusRow $relPath $updated $size ${stateNew}
                $targets.Add($row)
                switch ($decision.Reason) {
                    "lost"    { $count.Lost++ }
                    "new"     { $count.New++ }
                    "pending" { $count.Pending++ }
                    default   { $count.Updated++ }
                }
            }
        }
        $rows.Add($row)
    }

    $removed = 0
    foreach ($relPath in @($previous.Keys)) {
        if (!$relPath.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase) -or $found.Contains($relPath)) {
            continue
        }
        if ($scan.HasError) {
            $rows.Add($previous[$relPath])
            continue
        }
        removeBookDir (getBookDir $relPath)
        $removed++
    }

    $detail = "変換済み {0} 件 / 新規 {1} 件 / 更新あり {2} 件 / 前回未完了 {3} 件 / 前回失敗 {4} 件" -f
        $count.Done, $count.New, $count.Updated, $count.Pending, $failed.Count
    if ($count.Lost -gt 0) {
        # インデックスを直接削除された場合など。ふだんは 0 件のため、あるときだけ表示する
        $detail += " / 変換結果が無い・壊れている $($count.Lost) 件"
    }
    Write-Host ("  [{0}] Officeファイル {1} 件（{2}）" -f $folder.Name, $scan.Files.Count, $detail)
    if ($count.Lost -gt 0) {
        Write-Host "    変換結果（TSV）が無くなった・壊れている $($count.Lost) 件は変換し直します。（インデックスを直接削除した・0 バイトのTSVが残っている）" -ForegroundColor Yellow
    }
    if ($removed -gt 0) {
        Write-Host "    元ファイルが無くなった ${removed} 件の変換結果（TSV）を削除しました。"
    }
    if ($scan.HasError) {
        Write-Host "    アクセスできないフォルダがあったため、元ファイルが無くなったかどうかの確認は行いませんでした。" -ForegroundColor Yellow
    }

    $plan = newConvertPlanRow $folder.Name $folder.Path ${planKindConvert} $scan.Files.Count $targets.Count `
        $count.New $count.Updated $count.Pending $count.Lost $failed.Count
    return @{ Rows = $rows; Targets = $targets; Failed = $failed; Plan = $plan }
}

function waitForConvertApproval {
    # 変換対象の件数を画面に渡し（変換予定.tsv）、［変換を開始］（変換開始要求）か［キャンセル］（変換中止要求）の返事を待つ。
    #   変換する → @{ RetryFailed } / 取りやめ → $null
    # 画面を閉じた・落ちた場合に待ち続けないよう、$approvalTimeoutMinutes で打ち切って取りやめる
    param (
        $plan,               # newConvertPlanRow の配列（インデックスごと）
        [int]$targetCount,   # 変換対象の合計（画面の進み具合に出す）
        [int]$failedCount    # 前回失敗の合計（画面で再変換するかを選ぶ）
    )

    removeConvertStartRequest
    writeConvertPlan $plan
    writeConvertProgress ${convertPhaseConfirm} 0 $targetCount $failedCount "変換する内容を画面で確認しています…"
    Write-Host ""
    Write-Host "変換対象を画面に表示しました。［変換を開始］が押されるまで待ちます。（${approvalTimeoutMinutes} 分待っても返事が無ければ取りやめます）"

    $limit = (Get-Date).AddMinutes($approvalTimeoutMinutes)
    while ($true) {
        if (Test-Path -LiteralPath ${stopRequestFile}) {
            # 画面で［キャンセル］［中止］を押した
            Remove-Item -LiteralPath ${stopRequestFile} -Force
            removeConvertPlan
            return $null
        }
        $answer = readConvertStartRequest
        if ($answer) {
            removeConvertStartRequest
            removeConvertPlan
            return $answer
        }
        if ((Get-Date) -gt $limit) {
            Write-Host "画面からの返事が ${approvalTimeoutMinutes} 分ありませんでした。変換を取りやめます。" -ForegroundColor Yellow
            removeConvertPlan
            return $null
        }
        Start-Sleep -Milliseconds 300
    }
}
