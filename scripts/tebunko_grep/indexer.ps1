# Office → TSV 変換
#
# 画面で設定した変換対象フォルダ（setting.config。チェックなしのフォルダは変換しない）配下の
# Excel・Word・PowerPoint ファイルをシート・ページ・スライドごとにTSVへ変換し、work\index に保存する。
#   work\index\<インデックス名（フォルダごと）>\<相対フォルダ>\<ファイル名>.<拡張子>\<場所>.tsv
#   （元のファイル名はフォルダ名にする。以前の形式（…\<ファイル名>_<場所>.tsv）は変換時にこの形へ移す）
#
# ・Excel は Excel で変換する（セルの表示値を得るため）
# ・Word・PowerPoint（.docx / .pptx 等）は、ファイルを直接読む（Word・PowerPointは使わない）
# ・旧形式（.doc / .ppt）は、Word・PowerPointで新形式に変換してから読む
# ・元のファイルは占有しないよう、作業フォルダにコピーしてからコピーを開く・読む（変換中もほかの人が編集・保存できる）
# ・全ファイルの更新日時・サイズ・状態を work\変換一覧.tsv に記録し、
#   前回から更新されたファイル・未変換のファイルだけを変換する
# ・1ファイルの変換が制限時間を超えたら、Officeアプリを強制終了してそのファイルを失敗とし、次のファイルへ進む
# ・変換中に強制終了した場合、次回はそのファイルを最後に回す（続けて強制終了したら失敗とする）
# ・変換に失敗したファイルは一覧の状態を「失敗」とし、更新されない限り次回以降はスキップする（-RetryFailed で再変換する）
#
# 画面（gui.ps1）からウィンドウ無しで起動する。入力は求めない。
#   -RetryFailed    : 前回失敗し、その後更新されていないファイルも再変換する
#   -ConfirmTargets : 変換対象を数えた後、いったん止まって画面の返事を待つ（画面から起動したときに使う）。
#                     インデックスごとの件数を work\変換予定.tsv に書き、画面が work\変換開始要求 を作成したら変換を始める
#   中止            : 画面が work\変換中止要求 を作成すると、ファイルの切れ目で中止する（次回は未処理のファイルから再開する）
#   終了コード      : 0 = 完了（ファイルごとの失敗は変換一覧に記録）/ 1 = 続けられないエラー（work\変換エラー.txt）/ 2 = 中止（確認で取りやめた場合を含む）
#   表示内容は work\変換ログ.txt に記録する

param (
    [switch]$RetryFailed,
    [switch]$ConfirmTargets
)

. "$PSScriptRoot\lib.ps1"
. "$PSScriptRoot\..\shared\office\office_reader.ps1"
. "$PSScriptRoot\..\shared\office\office_app.ps1"
. "$PSScriptRoot\indexer\indexer_plan.ps1"
. "$PSScriptRoot\indexer\extract_office.ps1"
. "$PSScriptRoot\indexer\index_migrate.ps1"

$ErrorActionPreference = "Stop"

trap {
    # 続けられないエラー。画面に伝えるため、メッセージをファイルに書いて終了する
    $message = $_.Exception.Message
    Write-Host "＜エラー＞"
    Write-Host $message -ForegroundColor Red
    try {
        [System.IO.Directory]::CreateDirectory(${workDir}) | Out-Null
        [System.IO.File]::WriteAllText(${indexingErrorFile}, $message, ${utf8Bom})
    } catch {}
    try { Stop-Transcript | Out-Null } catch {}
    exit 1
}

# 前回の実行で残った中止要求・エラーは使わない。表示内容はログに記録する
[System.IO.Directory]::CreateDirectory(${workDir}) | Out-Null
foreach ($oldFile in @(${stopRequestFile}, ${indexingErrorFile}, ${ingestPlanFile}, ${indexingStartRequestFile})) {
    if (Test-Path -LiteralPath $oldFile) {
        Remove-Item -LiteralPath $oldFile -Force
    }
}
# 前回の進み具合も消す。残っていると、画面が起動直後に前回の最後の1行（「仕上げ」など）を読んでしまう
removeIndexingProgress
try { Start-Transcript -LiteralPath ${indexingLogFile} -Force | Out-Null } catch {}

# 同じ work（同じ配置フォルダ）に対して変換を2つ動かすと、変換一覧・インデックスが食い違うため1つだけ動かす。
# 画面は実行中の変換を見つけて進み具合を表示するため、ここに来るのは画面を使わずに起動した場合。
# ミューテックスはプロセスが終われば解放されるため、強制終了されても残らない
$indexerMutex = newAppMutex "indexer"
if (!$indexerMutex.Acquired) {
    throw "ほかの変換が実行中です。変換が終わってから実行してください。"
}

# 進み具合は1行のファイル（変換進捗.txt）に書く。画面はこれを読んで表示する
# （変換一覧は数万行になるため、画面が毎秒読み直すと、その間ずっと画面が固まる）
writeIndexingProgress ${indexingPhaseCrawl} 0 0 0 "変換の準備をしています…"

$restartInterval = 100  # Officeアプリを再起動する間隔（ファイル数）。メモリ肥大化対策（再起動 1 回で起動し直す約 2 秒かかるため、間隔を詰めすぎない）
$fileTimeoutMinutes = 10  # 1ファイルの変換の制限時間（分）。超えたらOfficeアプリを強制終了し、そのファイルは失敗とする
$approvalTimeoutMinutes = 60  # -ConfirmTargets で画面の返事を待つ制限時間（分）。画面が落ちた場合に待ち続けないよう打ち切る
$interruptLimit = 2       # 変換中に続けて強制終了した回数がこれに達したファイルは、失敗として以降スキップする
$failureListLimit = 50    # 終了時に失敗したファイルと原因を表示する最大件数（残りは変換一覧で確認する）

# ----------------------------------------------------------------------------
# メイン処理
# ----------------------------------------------------------------------------

$targetFolders = @(getTargetFolders)
if ($targetFolders.Count -eq 0) {
    throw "変換対象フォルダがありません。画面で変換対象フォルダを追加してください。"
}
if (@($targetFolders | Where-Object { $_.Enabled }).Count -eq 0) {
    throw "チェックの付いた変換対象フォルダがありません。画面で変換するフォルダにチェックを付けてください。"
}

[System.IO.Directory]::CreateDirectory($indexDir) | Out-Null
removeStaleTmpDirs
[System.IO.Directory]::CreateDirectory($tmpDir) | Out-Null
[System.IO.Directory]::CreateDirectory(${publishDir}) | Out-Null

# 以前の版の途中状態ファイル（変換一覧.tsv に置き換えた）は使わないため削除する
foreach ($name in @("変換対象一覧.txt", "変換失敗一覧.txt")) {
    $oldFile = Join-Path $workDir $name
    if (Test-Path -LiteralPath $oldFile) {
        Remove-Item -LiteralPath $oldFile -Force
    }
}

# 変換対象フォルダごとにインデックス名（work\index 直下のフォルダ名）を決める。前回と同じフォルダは同じ名前を使う
$statusExists = Test-Path -LiteralPath $statusFile
$status = readStatusFile
$folders = @(assignIndexNames $targetFolders $status.Folders)
$previous = moveLegacyIndex $folders $status $statusExists
removeDroppedFolders $folders $status.Folders
migrateFlatIndex

# 新しく割り当てたインデックス名を設定に保存する（インデックスの「名前」と「置き場所」を設定で分けて持つため。
# 名前が設定にあれば、フォルダを移してパスを書き換えても同じインデックスとして扱える）
if (@($targetFolders | Where-Object { -not $_.Name }).Count -gt 0) {
    writeTargetFolders $folders
    Write-Host "インデックス名を設定に保存しました: $((@($folders | ForEach-Object { $_.Name }) -join '、'))"
}

Write-Host "出力先フォルダ: ${indexDir}"
Write-Host ""
Write-Host "変換対象のファイルを検索しています..."
# 変換一覧の「済」に対してインデックス（TSV）が残っているかを調べるため、今あるTSVの数を数えておく
# （利用者が work\index を直接削除した場合に、「済」のまま検索できなくなるのを防ぐ）
writeIndexingProgress ${indexingPhaseCrawl} 0 0 0 "変換済みのインデックスを確認しています…"
$indexCounts = getIndexTsvCounts
if ($null -eq $indexCounts) {
    Write-Host "  インデックスのフォルダを調べられないため、変換結果が残っているかの確認は行いません。" -ForegroundColor Yellow
}
$rows = New-Object System.Collections.Generic.List[object]
$targets = New-Object System.Collections.Generic.List[object]
$failed = New-Object System.Collections.Generic.List[object]
$plan = New-Object System.Collections.Generic.List[object]   # 画面の確認に出す、インデックスごとの件数
foreach ($folder in $folders) {
    if (-not $folder.Enabled) {
        Write-Host "  [$($folder.Name)] $($folder.Path) … チェックなしのため変換しません（インデックスはそのまま残します）"
        $plan.Add((newIngestPlanRow $folder.Name $folder.Path ${planKindUnchecked}))
    } elseif (!(Test-Path -LiteralPath $folder.Path -PathType Container)) {
        Write-Host "  [$($folder.Name)] $($folder.Path) … フォルダが見つからないため変換しません" -ForegroundColor Yellow
        $plan.Add((newIngestPlanRow $folder.Name $folder.Path ${planKindMissing}))
    } else {
        Write-Host "  [$($folder.Name)] $($folder.Path)"
        # 大きいフォルダ・ネットワーク越しでは時間がかかるため、どのフォルダを見ているかを画面に伝える
        writeIndexingProgress ${indexingPhaseCrawl} 0 0 0 "[$($folder.Name)] のOfficeファイルを探しています… $($folder.Path)"
        $list = createTargetList $folder $previous $indexCounts
        $rows.AddRange($list.Rows)
        $targets.AddRange($list.Targets)
        $failed.AddRange($list.Failed)
        $plan.Add($list.Plan)
        continue
    }

    # 変換しないフォルダは前回の結果をそのまま残す
    $prefix = "$($folder.Name)\"
    foreach ($key in @($previous.Keys)) {
        if ($key.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            $rows.Add($previous[$key])
        }
    }
}

# 画面から起動した場合は、数えた件数を画面に出して、変換するかどうかの返事を待つ。
# 「更新不要」かどうかも、この件数を見て画面が知らせる（変換対象が 0 件でも、前回失敗の再変換を選べる）
$retryTargets = [bool]$RetryFailed
if ($ConfirmTargets) {
    $answer = waitForIndexingApproval $plan.ToArray() $targets.Count $failed.Count
    if ($null -eq $answer) {
        # 取りやめ。1件も変換していないため、変換対象にした行は前回の記録のまま（一覧に無かったファイルは記録しない）にする。
        # 「未変換」で記録すると、次回［変換を開始］が［続きから再開］になり、中断したように見えるため。
        # 変換一覧自体は書き直す（以前の形式からの移行・無くなったファイルの削除を反映する必要があるため）
        $targetPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($row in $targets) {
            [void]$targetPaths.Add($row.相対パス)
        }
        $keep = New-Object System.Collections.Generic.List[object]
        foreach ($row in $rows) {
            if (!$targetPaths.Contains($row.相対パス)) {
                $keep.Add($row)
                continue
            }
            $old = $null
            if ($previous.TryGetValue($row.相対パス, [ref]$old)) {
                $keep.Add($old)
            }
        }
        Write-Host ""
        Write-Host "画面で取りやめたため、変換しません。（変換一覧は前回のままです）" -ForegroundColor Yellow
        writeStatusFile $folders $keep
        writeSourceFolderFile $folders
        writeIndexingProgress ${indexingPhaseFinish} 0 0 0 "変換を取りやめました"
        removeTmpDir
        try { Stop-Transcript | Out-Null } catch {}
        exit 2
    }
    $retryTargets = $answer.RetryFailed
}

if ($failed.Count -gt 0) {
    Write-Host ""
    if ($retryTargets) {
        Write-Host "前回変換に失敗し、その後更新されていないファイル $($failed.Count) 件も再変換します。"
        $targets.AddRange($failed)
    } else {
        Write-Host "前回変換に失敗し、その後更新されていないファイル $($failed.Count) 件はスキップします。（パスワード付きなど）"
    }
}

# 前回、変換中に強制終了した（ウィンドウを閉じた・PCが停止した等）ファイルは、同じファイルで止まり続けないよう最後に回す。
# 続けて $interruptLimit 回強制終了したファイルは、応答しなくなるファイルとみなして失敗とする
$interrupted = readIngestingFile
if ($interrupted) {
    $row = @($targets | Where-Object { $_.相対パス -eq $interrupted.RelPath }) | Select-Object -First 1
    if (!$row) {
        $interrupted = $null
        removeIngestingFile
    } elseif ($interrupted.Count -ge $interruptLimit) {
        [void]$targets.Remove($row)
        $row.状態 = ${stateFailed}
        $row.TSV数 = ""
        $row.変換日時 = formatFileTime (Get-Date)
        $row.エラー = "変換中に $($interrupted.Count) 回続けて強制終了されたため、変換を中止しました（Officeアプリが応答しなくなる可能性があります）"
        Write-Host ""
        Write-Host "変換中に $($interrupted.Count) 回続けて強制終了したファイルは、失敗としてスキップします: $($row.相対パス)" -ForegroundColor Yellow
        $interrupted = $null
        removeIngestingFile
    } else {
        [void]$targets.Remove($row)
        $targets.Add($row)
        Write-Host ""
        Write-Host "前回、変換中に強制終了したファイルは最後に変換します: $($row.相対パス)" -ForegroundColor Yellow
    }
}
writeIndexingProgress ${indexingPhaseCrawl} 0 $targets.Count 0 "変換一覧を書き出しています…"
writeStatusFile $folders $rows
# インデックスのフォルダごと別の場所・PCへコピーしても元のファイルの場所が分かるよう、インデックス名と変換対象フォルダの対応を置く
writeSourceFolderFile $folders

# インデックス名 → 変換対象フォルダ
$folderByName = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($folder in $folders) {
    $folderByName[$folder.Name] = $folder.Path.TrimEnd("\")
}

if ($targets.Count -eq 0) {
    Write-Host ""
    Write-Host "変換が必要なファイルはありません。（一覧: $(Split-Path $statusFile -Leaf)）" -ForegroundColor Green
    writeIndexingProgress ${indexingPhaseFinish} 0 0 0 "変換が必要なファイルはありませんでした"
    removeTmpDir
    try { Stop-Transcript | Out-Null } catch {}
    exit 0
}

Write-Host ""
Write-Host "$($targets.Count) 件のファイルを変換します。（1ファイルの変換に ${fileTimeoutMinutes} 分以上かかった場合は、そのファイルを失敗として次のファイルへ進みます）"

$total = $targets.Count
$successCount = 0
$stopped = $false
$failures = New-Object System.Collections.Generic.List[object]  # 今回失敗したファイル: @{ RelPath; Message }
# 変換の直前に元のファイルが無くなっていたファイルの相対パス。変換一覧から除く
$droppedRows = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$folderLost = ""  # 変換中に見えなくなった変換対象フォルダ（見つかったら中止する）
$remaining = 0

try {
    startWatchdog

    for ($i = 0; $i -lt $total; $i++) {
        # 画面から中止を求められたら、次のファイルに進まずに終える（残りは「未変換」のまま、次回変換する）
        if (Test-Path -LiteralPath ${stopRequestFile}) {
            Remove-Item -LiteralPath ${stopRequestFile} -Force
            $stopped = $true
            $remaining = $total - $i
            Write-Host ""
            Write-Host "中止の要求を受けたため、変換を中止します。（残り ${remaining} 件は次回変換します）" -ForegroundColor Yellow
            break
        }

        $row = $targets[$i]
        $relPath = $row.相対パス
        $parts = splitIndexRelPath $relPath
        $sourceFolder = $folderByName[$parts.Name]
        $sourcePath = Join-Path $sourceFolder $parts.Rest
        Write-Host ("[{0}/{1}] {2}" -f ($i + 1), $total, $relPath)
        # 画面はこの1行から進み具合を作る（変換一覧は読まない）
        writeIndexingProgress ${indexingPhaseIngest} $i ($total - $i) $failures.Count $relPath

        # 変換対象を調べてから変換するまでの間に、元のファイルが移動・削除されることがある。
        # 「失敗」として記録すると、再変換を選ぶまで残ってしまうため、無くなったファイルは一覧・インデックスから除く
        if (![System.IO.File]::Exists((toLongPath $sourcePath))) {
            if (!(Test-Path -LiteralPath $sourceFolder -PathType Container)) {
                # 変換対象フォルダごと見えなくなった（ネットワークの切断・USBメモリの取り外し等）。
                # 残りのファイルをすべて失敗にしないよう、「未変換」のまま中止する（次回、続きから変換できる）。
                # finally（Officeアプリの終了・変換一覧の書き直し）を通すため、例外にせず抜ける
                $folderLost = $sourceFolder
                $remaining = $total - $i
                break
            }
            Write-Host "    元のファイルが無くなったため、変換せずに一覧から除きます。（移動・削除・名前変更された）" -ForegroundColor Yellow
            removeBookDir (getBookDir $relPath)
            [void]$droppedRows.Add($relPath)
            continue
        }

        # 強制終了したときに、どのファイルの変換中に止まったか次回分かるよう記録する
        $startCount = 1
        if ($interrupted -and $interrupted.RelPath -eq $relPath) {
            $startCount = $interrupted.Count + 1
        }
        writeIngestingFile $relPath $startCount

        try {
            clearTmpDir
            $script:watchdog.TimedOut = $false
            $script:watchdog.Deadline = (Get-Date).AddMinutes($fileTimeoutMinutes)
            try {
                $tsvCount = ingestFile $sourcePath
            } finally {
                $script:watchdog.Deadline = [datetime]::MaxValue
            }
            publishTsv (getBookDir $relPath)
            Write-Host "    TSV ${tsvCount} 件を作成しました。"
            $row.状態 = ${stateDone}
            $row.TSV数 = [string]$tsvCount
            $row.エラー = ""
            $successCount++
        } catch {
            $message = describeIngestError $_.Exception
            if ($script:watchdog.TimedOut) {
                $message = "${fileTimeoutMinutes} 分以内に変換が終わらなかったため中止しました（Officeアプリを強制終了しました）"
            }
            Write-Host "    変換に失敗しました: ${message}" -ForegroundColor Red
            $row.状態 = ${stateFailed}
            $row.TSV数 = ""
            $row.エラー = $message
            $failures.Add(@{ RelPath = $relPath; Message = $message })

            # アプリが不安定になっている可能性があるため終了する（次に必要になったときに起動し直す）
            try { stopApp (getAppName $relPath) } catch {}
        }

        # 中断しても結果が残るよう、1件ごとに変換一覧へ追記する（最後に1ファイル1行にまとめ直す）
        $row.変換日時 = formatFileTime (Get-Date)
        addStatusRow $row
        removeIngestingFile

        if ($script:watchdog.TimedOut -or (($i + 1) % $restartInterval) -eq 0) {
            # 制限時間を過ぎて強制終了したアプリは使えないため、すべて終了して次に必要になったときに起動し直す
            stopAllApps
        }
    }
} finally {
    # 中止・続けられないエラーの場合もここは実行される。
    # 後片付けも数十秒かかることがあるため、何をしているかを画面に伝える（進み具合の数はそのまま残す）
    # 画面は「成功 = 処理済み - 失敗」と出すため、変換しなかった（元ファイルが無くなった）分は数に入れない
    $processed = $successCount + $failures.Count
    writeIndexingProgress ${indexingPhaseFinish} $processed 0 $failures.Count "Officeアプリを終了しています…"
    stopWatchdog
    stopAllApps
    removeTmpDir
    removeIngestingFile
    writeIndexingProgress ${indexingPhaseFinish} $processed 0 $failures.Count "変換一覧を書き直しています…"
    # 変換の直前に無くなっていたファイルの行は除く（次回の検索でも見つからず、インデックスも削除済み）
    writeStatusFile $folders @($rows | Where-Object { $_ -and !$droppedRows.Contains([string]$_.相対パス) })
    # 初めて変換したインデックスは、最初に書き出した時点ではまだフォルダが無いため、ここでもう一度書く
    # （work\index\<インデックス名>\元のフォルダ.txt。インデックス 1 個だけをコピーしても元のファイルの場所が分かる）
    writeSourceFolderFile $folders
    # 画面が終わり方（成功・失敗の件数）を読めるよう、進み具合は消さずに最後の状態を残す
    writeIndexingProgress ${indexingPhaseFinish} $processed $remaining $failures.Count ""
}

if ($folderLost) {
    # 変換対象フォルダが見えなくなった場合は、続けられないエラーとして画面に知らせる（残りは未変換のまま）
    $message = "変換対象フォルダが見つからなくなったため、変換を中止しました: ${folderLost}" +
        "（残り ${remaining} 件は未変換のまま残しました。フォルダを使えるようにしてから、もう一度変換してください）"
    Write-Host ""
    Write-Host $message -ForegroundColor Red
    [System.IO.File]::WriteAllText(${indexingErrorFile}, $message, ${utf8Bom})
    try { Stop-Transcript | Out-Null } catch {}
    exit 1
}

Write-Host ""
if ($stopped) {
    Write-Host "TSV変換を中止しました。（成功: ${successCount} 件 / 失敗: $($failures.Count) 件）" -ForegroundColor Yellow
} else {
    Write-Host "TSV変換が完了しました。（成功: ${successCount} 件 / 失敗: $($failures.Count) 件）" -ForegroundColor Green
}
if ($droppedRows.Count -gt 0) {
    Write-Host "変換の直前に元のファイルが無くなった $($droppedRows.Count) 件は、変換せずに一覧から除きました。" -ForegroundColor Yellow
}
Write-Host "各ファイルの状態・更新日時は $(Split-Path $statusFile -Leaf) で確認できます。"
if ($failures.Count -gt 0) {
    # 変換中の表示は流れて見えなくなるため、失敗したファイルと原因を最後にまとめて表示する
    Write-Host ""
    Write-Host "＜変換に失敗したファイルと原因＞" -ForegroundColor Yellow
    foreach ($failure in @($failures | Select-Object -First $failureListLimit)) {
        Write-Host "  $($failure.RelPath)"
        Write-Host "    → $($failure.Message)" -ForegroundColor Red
    }
    if ($failures.Count -gt $failureListLimit) {
        Write-Host "  ほか $($failures.Count - $failureListLimit) 件（一覧の「状態」が「${stateFailed}」の行。原因は「エラー」列）"
    }
    Write-Host ""
    Write-Host "失敗したファイルは、画面で「失敗分も再変換する」を選んで変換すると再変換します。" -ForegroundColor Yellow
}

try { Stop-Transcript | Out-Null } catch {}
if ($stopped) {
    exit 2
}
exit 0
