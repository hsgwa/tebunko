# Office → TSV インデックス作成
#
# 画面で設定したクロール対象フォルダ（setting.config。チェックなしのフォルダは取り込まない）配下の
# Excel・Word・PowerPoint ファイルをシート・ページ・スライドごとにTSVへ書き出し、work\index に保存する。
#   work\index\<インデックス名（フォルダごと）>\<相対フォルダ>\<ファイル名>.<拡張子>\<場所>.tsv
#   （元のファイル名はフォルダ名にする。以前の形式（…\<ファイル名>_<場所>.tsv）は取り込み時にこの形へ移す）
#
# ・Excel は Excel で抽出する（セルの表示値を得るため）
# ・Word・PowerPoint（.docx / .pptx 等）は、ファイルを直接読む（Word・PowerPointは使わない）
# ・旧形式（.doc / .ppt）は、Word・PowerPointで新形式に変換してから読む
# ・元のファイルは占有しないよう、作業フォルダにコピーしてからコピーを開く・読む（インデックス作成中もほかの人が編集・保存できる）
# ・全ファイルの更新日時・サイズ・状態を work\取り込み一覧.tsv に記録し、
#   前回から更新されたファイル・未取り込みのファイルだけを取り込む
# ・1ファイルの取り込みが制限時間を超えたら、Officeアプリを強制終了してそのファイルを失敗とし、次のファイルへ進む
# ・取り込み中に強制終了した場合、次回はそのファイルを最後に回す（続けて強制終了したら失敗とする）
# ・取り込みに失敗したファイルは一覧の状態を「失敗」とし、更新されない限り次回以降はスキップする（-RetryFailed で再取り込みする）
#
# 画面（gui.ps1）からウィンドウ無しで起動する。入力は求めない。
#   -RetryFailed    : 前回失敗し、その後更新されていないファイルも再取り込みする
#   -ConfirmTargets : 取り込み対象を数えた後、いったん止まって画面の返事を待つ（画面から起動したときに使う）。
#                     インデックスごとの件数を work\取り込み予定.tsv に書き、画面が work\インデックス作成開始要求 を作成したらインデックス作成を始める
#   中止            : 画面が work\インデックス作成中止要求 を作成すると、ファイルの切れ目で中止する（次回は未処理のファイルから再開する）
#   終了コード      : 0 = 完了（ファイルごとの失敗は取り込み一覧に記録）/ 1 = 続けられないエラー（work\インデックス作成エラー.txt）/ 2 = 中止（確認で取りやめた場合を含む）
#   表示内容は work\インデックス作成ログ.txt に記録する

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

# 既定のワークスペースにほかのファイルが置いてあれば、インデックスのファイルと混ざるため作成しない。
# 空でないフォルダにエラーのファイルを書かないよう、例外（下の trap）にせずに終える（画面は起動する前に同じ確認をする）
$workspaceBlock = getWorkspaceBlockMessage
if ($workspaceBlock) {
    Write-Host $workspaceBlock -ForegroundColor Red
    exit 1
}

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

# 同じ work に対してインデックス作成を2つ動かすと、取り込み一覧・インデックスが食い違うため1つだけ動かす。
# work の置き場所は設定で変えられ、別の配置フォルダのツールが同じ work を使うこともあるため、鍵は work のパスから作る。
# 画面は実行中のインデックス作成を見つけて進み具合を表示するため、ここに来るのは画面を使わずに起動した場合。
# ミューテックスはプロセスが終われば解放されるため、強制終了されても残らない。
# 実行中のインデックス作成が画面とやり取りしているファイル（中止要求・取り込み予定・開始要求・進み具合）とログを
# 消したり上書きしたりしないよう、下の後片付けより先に確かめる
$indexerMutex = newAppMutex "indexer" ${workDir}
if (!$indexerMutex.Acquired) {
    throw "ほかのインデックス作成が実行中です。インデックス作成が終わってから実行してください。"
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

# 進み具合は1行のファイル（インデックス作成進捗.txt）に書く。画面はこれを読んで表示する
# （取り込み一覧は数万行になるため、画面が毎秒読み直すと、その間ずっと画面が固まる）
writeIndexingProgress ${indexingPhaseCrawl} 0 0 0 "インデックス作成の準備をしています…"

$restartInterval = 100  # Officeアプリを再起動する間隔（ファイル数）。メモリ肥大化対策（再起動 1 回で起動し直す約 2 秒かかるため、間隔を詰めすぎない）
$fileTimeoutMinutes = 10  # 1ファイルの取り込みの制限時間（分）。超えたらOfficeアプリを強制終了し、そのファイルは失敗とする
$approvalTimeoutMinutes = 60  # -ConfirmTargets で画面の返事を待つ制限時間（分）。画面が落ちた場合に待ち続けないよう打ち切る
$interruptLimit = 2       # 取り込み中に続けて強制終了した回数がこれに達したファイルは、失敗として以降スキップする
$failureListLimit = 50    # 終了時に失敗したファイルと原因を表示する最大件数（残りは取り込み一覧で確認する）

# ----------------------------------------------------------------------------
# メイン処理
# ----------------------------------------------------------------------------

$targetFolders = @(getTargetFolders)
if ($targetFolders.Count -eq 0) {
    throw "クロール対象フォルダがありません。画面でクロール対象フォルダを追加してください。"
}
if (@($targetFolders | Where-Object { $_.Enabled }).Count -eq 0) {
    throw "チェックの付いたクロール対象フォルダがありません。画面で取り込むフォルダにチェックを付けてください。"
}

[System.IO.Directory]::CreateDirectory($indexDir) | Out-Null
removeStaleTmpDirs
[System.IO.Directory]::CreateDirectory($tmpDir) | Out-Null
[System.IO.Directory]::CreateDirectory(${publishDir}) | Out-Null

# 以前の版の途中状態ファイル（取り込み一覧.tsv に置き換えた）は使わないため削除する
foreach ($name in @("変換対象一覧.txt", "変換失敗一覧.txt")) {
    $oldFile = Join-Path $workDir $name
    if (Test-Path -LiteralPath $oldFile) {
        Remove-Item -LiteralPath $oldFile -Force
    }
}

# クロール対象フォルダごとにインデックス名（work\index 直下のフォルダ名）を決める。前回と同じフォルダは同じ名前を使う
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
Write-Host "クロールしています..."
# 取り込み一覧の「済」に対してインデックス（TSV）が残っているかを調べるため、今あるTSVの数を数えておく
# （利用者が work\index を直接削除した場合に、「済」のまま検索できなくなるのを防ぐ）
writeIndexingProgress ${indexingPhaseCrawl} 0 0 0 "取り込み済みのインデックスを確認しています…"
$indexCounts = getIndexTsvCounts
if ($null -eq $indexCounts) {
    Write-Host "  インデックスのフォルダを調べられないため、インデックスが残っているかの確認は行いません。" -ForegroundColor Yellow
}
$rows = New-Object System.Collections.Generic.List[object]
$targets = New-Object System.Collections.Generic.List[object]
$failed = New-Object System.Collections.Generic.List[object]
$plan = New-Object System.Collections.Generic.List[object]   # 画面の確認に出す、インデックスごとの件数
foreach ($folder in $folders) {
    if (-not $folder.Enabled) {
        Write-Host "  [$($folder.Name)] $($folder.Path) … チェックなしのため取り込みません（インデックスはそのまま残します）"
        $plan.Add((newIngestPlanRow $folder.Name $folder.Path ${planKindUnchecked}))
    } elseif (!(Test-Path -LiteralPath $folder.Path -PathType Container)) {
        Write-Host "  [$($folder.Name)] $($folder.Path) … フォルダが見つからないため取り込みません" -ForegroundColor Yellow
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

    # 取り込まないフォルダは前回の結果をそのまま残す
    $prefix = "$($folder.Name)\"
    foreach ($key in @($previous.Keys)) {
        if ($key.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            $rows.Add($previous[$key])
        }
    }
}

# 画面から起動した場合は、数えた件数を画面に出して、取り込むかどうかの返事を待つ。
# 「更新不要」かどうかも、この件数を見て画面が知らせる（取り込み対象が 0 件でも、前回失敗の再取り込みを選べる）
$retryTargets = [bool]$RetryFailed
if ($ConfirmTargets) {
    $answer = waitForIndexingApproval $plan.ToArray() $targets.Count $failed.Count
    if ($null -eq $answer) {
        # 取りやめ。1件も取り込んでいないため、取り込み対象にした行は前回の記録のまま（一覧に無かったファイルは記録しない）にする。
        # 「未取り込み」で記録すると、次回［インデックス作成を開始］が［続きから再開］になり、中断したように見えるため。
        # 取り込み一覧自体は書き直す（以前の形式からの移行・無くなったファイルの削除を反映する必要があるため）
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
        Write-Host "画面で取りやめたため、取り込みません。（取り込み一覧は前回のままです）" -ForegroundColor Yellow
        writeStatusFile $folders $keep
        writeSourceFolderFile $folders
        writeIndexingProgress ${indexingPhaseFinish} 0 0 0 "インデックス作成を取りやめました"
        removeTmpDir
        try { Stop-Transcript | Out-Null } catch {}
        exit 2
    }
    $retryTargets = $answer.RetryFailed
}

if ($failed.Count -gt 0) {
    Write-Host ""
    if ($retryTargets) {
        Write-Host "前回取り込みに失敗し、その後更新されていないファイル $($failed.Count) 件も再取り込みします。"
        $targets.AddRange($failed)
    } else {
        Write-Host "前回取り込みに失敗し、その後更新されていないファイル $($failed.Count) 件はスキップします。（パスワード付きなど）"
    }
}

# 前回、取り込み中に強制終了した（ウィンドウを閉じた・PCが停止した等）ファイルは、同じファイルで止まり続けないよう最後に回す。
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
        $row.取り込み日時 = formatFileTime (Get-Date)
        $row.エラー = "取り込み中に $($interrupted.Count) 回続けて強制終了されたため、取り込みを中止しました（Officeアプリが応答しなくなる可能性があります）"
        Write-Host ""
        Write-Host "取り込み中に $($interrupted.Count) 回続けて強制終了したファイルは、失敗としてスキップします: $($row.相対パス)" -ForegroundColor Yellow
        $interrupted = $null
        removeIngestingFile
    } else {
        [void]$targets.Remove($row)
        $targets.Add($row)
        Write-Host ""
        Write-Host "前回、取り込み中に強制終了したファイルは最後に取り込みます: $($row.相対パス)" -ForegroundColor Yellow
    }
}
writeIndexingProgress ${indexingPhaseCrawl} 0 $targets.Count 0 "取り込み一覧を書き出しています…"
writeStatusFile $folders $rows
# インデックスのフォルダごと別の場所・PCへコピーしても元のファイルの場所が分かるよう、インデックス名とクロール対象フォルダの対応を置く
writeSourceFolderFile $folders

# インデックス名 → クロール対象フォルダ
$folderByName = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($folder in $folders) {
    $folderByName[$folder.Name] = $folder.Path.TrimEnd("\")
}

if ($targets.Count -eq 0) {
    Write-Host ""
    Write-Host "取り込みが必要なファイルはありません。（一覧: $(Split-Path $statusFile -Leaf)）" -ForegroundColor Green
    # 取り込むファイルが無くても、システムインデックスがまだ無いフォルダ（この版に上げた直後など）は作る
    writeIndexingProgress ${indexingPhaseFinish} 0 0 0 "システムインデックス（高速検索用）を確かめています…"
    try {
        [void](updateSystemIndexes)
    } catch {
        Write-Host "システムインデックスを作れませんでした（次のインデックス作成で作り直します）: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    writeIndexingProgress ${indexingPhaseFinish} 0 0 0 "取り込みが必要なファイルはありませんでした"
    removeTmpDir
    try { Stop-Transcript | Out-Null } catch {}
    exit 0
}

Write-Host ""
Write-Host "$($targets.Count) 件のファイルを取り込みます。（1ファイルの取り込みに ${fileTimeoutMinutes} 分以上かかった場合は、そのファイルを失敗として次のファイルへ進みます）"

$total = $targets.Count
$successCount = 0
$stopped = $false
$failures = New-Object System.Collections.Generic.List[object]  # 今回失敗したファイル: @{ RelPath; Message }
# 取り込みの直前に元のファイルが無くなっていたファイルの相対パス。取り込み一覧から除く
$droppedRows = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$folderLost = ""  # インデックス作成中に見えなくなったクロール対象フォルダ（見つかったら中止する）
$remaining = 0

try {
    startWatchdog

    for ($i = 0; $i -lt $total; $i++) {
        # 画面から中止を求められたら、次のファイルに進まずに終える（残りは「未取り込み」のまま、次回取り込みする）
        if (Test-Path -LiteralPath ${stopRequestFile}) {
            Remove-Item -LiteralPath ${stopRequestFile} -Force
            $stopped = $true
            $remaining = $total - $i
            Write-Host ""
            Write-Host "中止の要求を受けたため、インデックス作成を中止します。（残り ${remaining} 件は次回取り込みします）" -ForegroundColor Yellow
            break
        }

        $row = $targets[$i]
        $relPath = $row.相対パス
        $parts = splitIndexRelPath $relPath
        $sourceFolder = $folderByName[$parts.Name]
        $sourcePath = Join-Path $sourceFolder $parts.Rest
        Write-Host ("[{0}/{1}] {2}" -f ($i + 1), $total, $relPath)
        # 画面はこの1行から進み具合を作る（取り込み一覧は読まない）
        writeIndexingProgress ${indexingPhaseIngest} $i ($total - $i) $failures.Count $relPath

        # 取り込み対象を調べてから取り込むまでの間に、元のファイルが移動・削除されることがある。
        # 「失敗」として記録すると、再取り込みを選ぶまで残ってしまうため、無くなったファイルは一覧・インデックスから除く
        if (![System.IO.File]::Exists((toLongPath $sourcePath))) {
            if (!(Test-Path -LiteralPath $sourceFolder -PathType Container)) {
                # クロール対象フォルダごと見えなくなった（ネットワークの切断・USBメモリの取り外し等）。
                # 残りのファイルをすべて失敗にしないよう、「未取り込み」のまま中止する（次回、続きから取り込める）。
                # finally（Officeアプリの終了・取り込み一覧の書き直し）を通すため、例外にせず抜ける
                $folderLost = $sourceFolder
                $remaining = $total - $i
                break
            }
            Write-Host "    元のファイルが無くなったため、取り込まずに一覧から除きます。（移動・削除・名前変更された）" -ForegroundColor Yellow
            removeBookDir (getBookDir $relPath)
            [void]$droppedRows.Add($relPath)
            continue
        }

        # 強制終了したときに、どのファイルの取り込み中に止まったか次回分かるよう記録する
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
            # 高速検索: このフォルダの システムインデックスを作り直すまで、検索ではこのフォルダを必ず照合させる
            if (!(markSystemIndexChanged @([System.IO.Path]::GetDirectoryName($relPath)))) {
                Write-Host "    システムインデックスの状態を書き込めませんでした（インデックス作成の終わりに作り直します）。" -ForegroundColor Yellow
            }
            Write-Host "    TSV ${tsvCount} 件を作成しました。"
            $row.状態 = ${stateDone}
            $row.TSV数 = [string]$tsvCount
            $row.エラー = ""
            $row.抽出版 = [string](getExtractVersion $relPath)
            $successCount++
        } catch {
            $message = describeIngestError $_.Exception
            if ($script:watchdog.TimedOut) {
                $message = "${fileTimeoutMinutes} 分以内に取り込みが終わらなかったため中止しました（Officeアプリを強制終了しました）"
            }
            Write-Host "    取り込みに失敗しました: ${message}" -ForegroundColor Red
            $row.状態 = ${stateFailed}
            $row.TSV数 = ""
            $row.エラー = $message
            $row.抽出版 = ""
            $failures.Add(@{ RelPath = $relPath; Message = $message })

            # アプリが不安定になっている可能性があるため終了する（次に必要になったときに起動し直す）
            try { stopApp (getAppName $relPath) } catch {}
        }

        # 中断しても結果が残るよう、1件ごとに取り込み一覧へ追記する（最後に1ファイル1行にまとめ直す）
        $row.取り込み日時 = formatFileTime (Get-Date)
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
    # 画面は「成功 = 処理済み - 失敗」と出すため、取り込まなかった（元ファイルが無くなった）分は数に入れない
    $processed = $successCount + $failures.Count
    writeIndexingProgress ${indexingPhaseFinish} $processed 0 $failures.Count "Officeアプリを終了しています…"
    stopWatchdog
    stopAllApps
    removeTmpDir
    removeIngestingFile
    writeIndexingProgress ${indexingPhaseFinish} $processed 0 $failures.Count "取り込み一覧を書き直しています…"
    # 取り込みの直前に無くなっていたファイルの行は除く（次回の検索でも見つからず、インデックスも削除済み）
    writeStatusFile $folders @($rows | Where-Object { $_ -and !$droppedRows.Contains([string]$_.相対パス) })
    # 初めて取り込んだインデックスは、最初に書き出した時点ではまだフォルダが無いため、ここでもう一度書く
    # （work\index\<インデックス名>\元のフォルダ.txt。インデックス 1 個だけをコピーしても元のファイルの場所が分かる）
    writeSourceFolderFile $folders
    # 高速検索用の システムインデックスを作り直す。中止したとき・フォルダが見えなくなったときは作らない
    # （作り直していないフォルダは反映待ちのままのため、検索ではそのフォルダを照合する）
    if (!$stopped -and !$folderLost) {
        writeIndexingProgress ${indexingPhaseFinish} $processed 0 $failures.Count "システムインデックス（高速検索用）を作っています…"
        try {
            [void](updateSystemIndexes)
        } catch {
            Write-Host "システムインデックスを作れませんでした（次のインデックス作成で作り直します）: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    # 画面が終わり方（成功・失敗の件数）を読めるよう、進み具合は消さずに最後の状態を残す
    writeIndexingProgress ${indexingPhaseFinish} $processed $remaining $failures.Count ""
}

if ($folderLost) {
    # クロール対象フォルダが見えなくなった場合は、続けられないエラーとして画面に知らせる（残りは未取り込みのまま）
    $message = "クロール対象フォルダが見つからなくなったため、インデックス作成を中止しました: ${folderLost}" +
        "（残り ${remaining} 件は未取り込みのまま残しました。フォルダを使えるようにしてから、もう一度取り込んでください）"
    Write-Host ""
    Write-Host $message -ForegroundColor Red
    [System.IO.File]::WriteAllText(${indexingErrorFile}, $message, ${utf8Bom})
    try { Stop-Transcript | Out-Null } catch {}
    exit 1
}

Write-Host ""
if ($stopped) {
    Write-Host "インデックス作成を中止しました。（成功: ${successCount} 件 / 失敗: $($failures.Count) 件）" -ForegroundColor Yellow
} else {
    Write-Host "インデックス作成が完了しました。（成功: ${successCount} 件 / 失敗: $($failures.Count) 件）" -ForegroundColor Green
}
if ($droppedRows.Count -gt 0) {
    Write-Host "取り込みの直前に元のファイルが無くなった $($droppedRows.Count) 件は、取り込まずに一覧から除きました。" -ForegroundColor Yellow
}
Write-Host "各ファイルの状態・更新日時は $(Split-Path $statusFile -Leaf) で確認できます。"
if ($failures.Count -gt 0) {
    # インデックス作成中の表示は流れて見えなくなるため、失敗したファイルと原因を最後にまとめて表示する
    Write-Host ""
    Write-Host "＜取り込みに失敗したファイルと原因＞" -ForegroundColor Yellow
    foreach ($failure in @($failures | Select-Object -First $failureListLimit)) {
        Write-Host "  $($failure.RelPath)"
        Write-Host "    → $($failure.Message)" -ForegroundColor Red
    }
    if ($failures.Count -gt $failureListLimit) {
        Write-Host "  ほか $($failures.Count - $failureListLimit) 件（一覧の「状態」が「${stateFailed}」の行。原因は「エラー」列）"
    }
    Write-Host ""
    Write-Host "失敗したファイルは、画面で「失敗分も再取り込みする」を選んで取り込むと再取り込みします。" -ForegroundColor Yellow
}

try { Stop-Transcript | Out-Null } catch {}
if ($stopped) {
    exit 2
}
exit 0
