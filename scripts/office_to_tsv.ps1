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
# 画面（config_gui.ps1）からウィンドウ無しで起動する。入力は求めない。
#   -RetryFailed : 前回失敗し、その後更新されていないファイルも再変換する
#   中止        : 画面が work\変換中止要求 を作成すると、ファイルの切れ目で中止する（次回は未処理のファイルから再開する）
#   終了コード  : 0 = 完了（ファイルごとの失敗は変換一覧に記録）/ 1 = 続けられないエラー（work\変換エラー.txt）/ 2 = 中止
#   表示内容は work\変換ログ.txt に記録する

param (
    [switch]$RetryFailed
)

. "$PSScriptRoot\common.ps1"
. "$PSScriptRoot\office_reader.ps1"

$ErrorActionPreference = "Stop"

trap {
    # 続けられないエラー。画面に伝えるため、メッセージをファイルに書いて終了する
    $message = $_.Exception.Message
    Write-Host "＜エラー＞"
    Write-Host $message -ForegroundColor Red
    try {
        [System.IO.Directory]::CreateDirectory(${workDir}) | Out-Null
        [System.IO.File]::WriteAllText(${convertErrorFile}, $message, ${utf8Bom})
    } catch {}
    try { Stop-Transcript | Out-Null } catch {}
    exit 1
}

# 前回の実行で残った中止要求・エラーは使わない。表示内容はログに記録する
[System.IO.Directory]::CreateDirectory(${workDir}) | Out-Null
foreach ($oldFile in @(${stopRequestFile}, ${convertErrorFile})) {
    if (Test-Path -LiteralPath $oldFile) {
        Remove-Item -LiteralPath $oldFile -Force
    }
}
try { Start-Transcript -LiteralPath ${convertLogFile} -Force | Out-Null } catch {}

# 同じ work（同じ配置フォルダ）に対して変換を2つ動かすと、変換一覧・インデックスが食い違うため1つだけ動かす。
# 画面は実行中の変換を見つけて進み具合を表示するため、ここに来るのは画面を使わずに起動した場合。
# ミューテックスはプロセスが終われば解放されるため、強制終了されても残らない
$convertMutex = newAppMutex "convert"
if (!$convertMutex.Acquired) {
    throw "ほかの変換が実行中です。変換が終わってから実行してください。"
}

# 進み具合は1行のファイル（変換進捗.txt）に書く。画面はこれを読んで表示する
# （変換一覧は数万行になるため、画面が毎秒読み直すと、その間ずっと画面が固まる）
writeConvertProgress ${convertPhaseScan} 0 0 0 "変換の準備をしています…"

$targetExtensions = ${officeExtensions}  # 変換対象の拡張子（common.ps1。画面のフォルダ選択でも同じ一覧を使う）
$restartInterval = 50  # Officeアプリを再起動する間隔（ファイル数）。メモリ肥大化対策
$fileTimeoutMinutes = 10  # 1ファイルの変換の制限時間（分）。超えたらOfficeアプリを強制終了し、そのファイルは失敗とする
$interruptLimit = 2       # 変換中に続けて強制終了した回数がこれに達したファイルは、失敗として以降スキップする
$excelMaxPath = 218       # Excelで開けるパスの長さの目安（古い版の上限）。作業フォルダのコピーのパスがこれ以上なら短い名前にする
$excelExtraCells = 1000000  # 使用範囲がデータの範囲よりこのセル数以上広いシートは、データの範囲だけを一時シートにコピーしてから書き出す
$failureListLimit = 50    # 終了時に失敗したファイルと原因を表示する最大件数（残りは変換一覧で確認する）

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
    # 変換対象フォルダを1つ検索して変換一覧の行を作り直し、@{ Rows（全ファイル）; Targets（変換する）; Failed（前回失敗し、更新の無い） } を返す。
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

        $sameFile = ($old -and $old.更新日時 -eq $updated -and $old.サイズ -eq $size -and $old.状態 -ne ${stateNew})
        # 変換済みでも、インデックス（TSV）が無くなっていれば変換し直す。
        # 一覧だけを見ると「済」のままになり、検索しても出てこない状態が続くため
        $lostIndex = ($sameFile -and $old.状態 -eq ${stateDone} -and -not (testIndexComplete $old $relPath $counts))

        if ($sameFile -and -not $lostIndex) {
            # 更新なし。失敗したファイルを再変換するかは呼び出し元で決める
            $row = $old
            if ($row.状態 -eq ${stateFailed}) {
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
                if ($lostIndex) {
                    $count.Lost++
                } elseif ($null -eq $old) {
                    $count.New++
                } elseif ($old.状態 -eq ${stateNew}) {
                    $count.Pending++
                } else {
                    $count.Updated++
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

    return @{ Rows = $rows; Targets = $targets; Failed = $failed }
}

# ----------------------------------------------------------------------------
# Officeアプリ（Excel・Word・PowerPoint）の起動・終了
# ----------------------------------------------------------------------------

# 起動中のアプリ: 名前 → @{ Com; Pid; Shared }
$script:apps = @{}

$appInfo = @{
    Excel      = @{ ProgId = "Excel.Application";      Process = "EXCEL" }
    Word       = @{ ProgId = "Word.Application";       Process = "WINWORD" }
    PowerPoint = @{ ProgId = "PowerPoint.Application"; Process = "POWERPNT" }
}

function getApp {
    # アプリのCOMオブジェクトを返す。起動していなければ起動する
    param (
        [string]$name
    )

    if (-not $script:apps.ContainsKey($name)) {
        $info = $appInfo[$name]

        # 終了できなかった場合に強制終了するため、新しく起動したプロセスのIDを控えておく
        $before = @(Get-Process -Name $info.Process -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
        $com = New-Object -ComObject $info.ProgId
        $after = @(Get-Process -Name $info.Process -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
        $newIds = @($after | Where-Object { $before -notcontains $_ })

        switch ($name) {
            "Excel" {
                $com.Visible = $false
                $com.DisplayAlerts = $false
                $com.EnableEvents = $false
                $com.ScreenUpdating = $false
                $com.AskToUpdateLinks = $false
            }
            "Word" {
                $com.Visible = $false
                $com.DisplayAlerts = 0  # wdAlertsNone
            }
            "PowerPoint" {
                # PowerPointはウィンドウを隠せないため、ファイルをウィンドウ無しで開く（Visible は変更しない）
                $com.DisplayAlerts = 1  # ppAlertsNone
            }
        }
        $com.AutomationSecurity = 3  # msoAutomationSecurityForceDisable（マクロ無効）

        # PowerPoint等は起動中のアプリに接続することがある。その場合は利用者のものなので終了させない
        $script:apps[$name] = @{
            Com    = $com
            Pid    = $(if ($newIds.Count -eq 1) { $newIds[0] } else { 0 })
            Shared = ($newIds.Count -eq 0)
        }
        updateWatchedPids
    }
    return $script:apps[$name].Com
}

function stopApp {
    param (
        [string]$name
    )

    $app = $script:apps[$name]
    if ($null -eq $app) {
        return
    }
    $script:apps.Remove($name)
    updateWatchedPids

    # 変換中に利用者が同じアプリでファイルを開いた場合は、終了させない
    $inUse = $app.Shared
    if (-not $inUse) {
        try {
            switch ($name) {
                "Word"       { $inUse = ($app.Com.Documents.Count -gt 0) }
                "PowerPoint" { $inUse = ($app.Com.Presentations.Count -gt 0) }
            }
        } catch {}
    }

    if (-not $inUse) {
        try { $app.Com.Quit() } catch {}
    }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app.Com) } catch {}

    # GC::WaitForPendingFinalizers() はCOMの解放待ちで長時間（約60秒）止まることがあるため使わず、
    # 終了しなかったアプリはプロセスIDを指定して強制終了する
    if (-not $inUse -and $app.Pid) {
        $process = Get-Process -Id $app.Pid -ErrorAction SilentlyContinue
        if ($process -and -not $process.WaitForExit(5000)) {
            # 終了処理中のプロセスは Kill() が「アクセス拒否」で失敗することがあるが、そのまま終了するため無視する
            try { $process.Kill() } catch {}
        }
    }
}

function stopAllApps {
    # 1つのアプリの終了に失敗しても、残りのアプリは終了させる
    foreach ($name in @($script:apps.Keys)) {
        try {
            stopApp $name
        } catch {
            Write-Host "    ${name} の終了に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

function getAppName {
    # 拡張子から、変換に使うアプリの名前を返す
    param (
        [string]$path
    )

    switch -Regex ([System.IO.Path]::GetExtension($path).ToLower()) {
        "^\.xls" { return "Excel" }
        "^\.doc" { return "Word" }
        "^\.ppt" { return "PowerPoint" }
    }
    return $null
}

function releaseComObject($object) {
    if ($null -ne $object) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($object)
    }
}

# ----------------------------------------------------------------------------
# 1ファイルの制限時間の監視
# ----------------------------------------------------------------------------

# 変換中のCOM呼び出しは応答が無いと戻らず、Ctrl+C も効かないため、別スレッドで制限時間を監視する。
# 制限時間を過ぎたら、自分で起動したOfficeアプリを強制終了する（COM呼び出しが例外で戻り、そのファイルは失敗になる）。
#   Deadline: 変換中のファイルの制限時刻（変換中でなければ MaxValue）
#   Pids    : 強制終了してよいプロセスID（自分で起動したOfficeアプリ）
#   TimedOut: 制限時間を過ぎて強制終了した
$script:watchdog = [hashtable]::Synchronized(@{ Deadline = [datetime]::MaxValue; Pids = @(); TimedOut = $false; Stop = $false })
$script:watchdogThread = $null

function updateWatchedPids {
    # 強制終了してよいプロセスIDを、起動中のアプリのうち自分で起動したものにする（利用者のアプリは終了させない）
    $script:watchdog.Pids = @($script:apps.Values | Where-Object { -not $_.Shared -and $_.Pid } | ForEach-Object { $_.Pid })
}

function startWatchdog {
    $ps = [PowerShell]::Create()
    [void]$ps.AddScript({
        param($watch, [string[]]$processNames)
        while (-not $watch.Stop) {
            Start-Sleep -Milliseconds 500
            if ([datetime]::Now -lt $watch.Deadline) {
                continue
            }
            $watch.Deadline = [datetime]::MaxValue
            $watch.TimedOut = $true
            foreach ($id in @($watch.Pids)) {
                try {
                    # 終了済みでIDが別のプロセスに再利用されている場合に備え、Officeアプリであることを確かめる
                    $process = [System.Diagnostics.Process]::GetProcessById($id)
                    if ($processNames -contains $process.ProcessName) {
                        $process.Kill()
                    }
                } catch {}
            }
        }
    }).AddArgument($script:watchdog).AddArgument([string[]]@($appInfo.Values | ForEach-Object { $_.Process }))
    $script:watchdogThread = @{ PowerShell = $ps; Handle = $ps.BeginInvoke() }
}

function stopWatchdog {
    if ($null -eq $script:watchdogThread) {
        return
    }
    $script:watchdog.Stop = $true
    try { [void]$script:watchdogThread.PowerShell.EndInvoke($script:watchdogThread.Handle) } catch {}
    $script:watchdogThread.PowerShell.Dispose()
    $script:watchdogThread = $null
}

# ----------------------------------------------------------------------------
# 変換
# ----------------------------------------------------------------------------

function copyDataRangeToTempSheet {
    # 使用範囲（UsedRange）が実際のデータよりずっと広いシートのために、データのある範囲だけを
    # 同じブックの一時シートへ同じ位置でコピーし、その一時シートを返す。縮める必要が無い・できない場合は $null を返す。
    #
    # 最終行・右端のセルに書式だけが残っていると使用範囲がシート全体になり、Excelのテキスト保存が
    # 空セルのタブだけで数GBを書き出して制限時間も超える（実測: 10分で6GB、終わらない）。
    # 元のシートの行・列は削除しない（そこを参照する数式が #REF! になり、ほかのシートの表示値まで変わるため）。
    param (
        $workbook,
        $worksheet
    )

    $usedRange = $worksheet.UsedRange
    try {
        $firstRow = $usedRange.Row
        $firstColumn = $usedRange.Column
        $usedLastRow = $firstRow + $usedRange.Rows.Count - 1
        $usedLastColumn = $firstColumn + $usedRange.Columns.Count - 1
    } finally {
        releaseComObject $usedRange
    }

    # 値・数式のある最後の行・列を探す（書式だけのセルには当たらない）。
    # 引数: What, After, LookIn（-4123 = xlFormulas）, LookAt, SearchOrder（1 = xlByRows / 2 = xlByColumns）, SearchDirection（2 = xlPrevious）
    $topLeft = $worksheet.Range("A1")
    $cells = $worksheet.Cells
    try {
        $found = $cells.Find("*", $topLeft, -4123, [Type]::Missing, 1, 2)
        $dataLastRow = if ($found) { $found.Row } else { 0 }
        $found = $cells.Find("*", $topLeft, -4123, [Type]::Missing, 2, 2)
        $dataLastColumn = if ($found) { $found.Column } else { 0 }
    } finally {
        releaseComObject $cells
        releaseComObject $topLeft
    }

    # 値のあるセルが無い（書式だけのシート）場合は、そのまま書き出して空のTSVとして扱う
    if ($dataLastRow -lt $firstRow -or $dataLastColumn -lt $firstColumn) {
        return $null
    }

    # セル数は int を超えるため double で数える（シート全体は約 172 億セル）
    $usedCells = [double]($usedLastRow - $firstRow + 1) * ($usedLastColumn - $firstColumn + 1)
    $dataCells = [double]($dataLastRow - $firstRow + 1) * ($dataLastColumn - $firstColumn + 1)
    if (($usedCells - $dataCells) -lt $excelExtraCells) {
        return $null
    }

    # 数式の参照先がずれないよう、コピー先は元と同じ位置（行・列）にする
    $temp = $null
    try {
        $temp = $workbook.Worksheets.Add()
        $source = $worksheet.Range($worksheet.Cells.Item($firstRow, $firstColumn), $worksheet.Cells.Item($dataLastRow, $dataLastColumn))
        try {
            [void]$source.Copy($temp.Cells.Item($firstRow, $firstColumn))
        } finally {
            releaseComObject $source
        }
        return $temp
    } catch {
        # ブックの構成が保護されている場合など。元のシートをそのまま書き出す（時間切れで失敗することがある）
        Write-Host "    $($worksheet.Name) の使用範囲を縮められませんでした: $($_.Exception.Message)" -ForegroundColor Yellow
        if ($temp) {
            $temp.Delete()
            releaseComObject $temp
        }
        return $null
    }
}

function convertWorkbook {
    # Excelファイルをシートごとに作業フォルダへTSV出力し、出力したシート数を返す
    param (
        [string]$sourcePath
    )

    $bookName = [System.IO.Path]::GetFileName($sourcePath)

    # 元のファイルを占有しないよう、作業フォルダにコピーしてからコピーを開く
    # （Excelで開いている間、元のファイルを利用者が上書き保存・移動できなくなるのを防ぐ。長いパスのファイルも開ける）。
    # ファイル名を参照する数式（CELL("filename") 等）の表示値が変わらないよう、コピーは元と同じファイル名にする。
    # 作業フォルダ＋ファイル名が長すぎてExcelで開けない場合だけ、短い名前にする。
    # コピーはブックを閉じた後に削除する（開けずに例外になった場合は、次のファイルの変換前・終了時に作業フォルダごと空にする）
    $copyPath = Join-Path $tmpDir $bookName
    if ($copyPath.Length -ge $excelMaxPath) {
        $copyPath = Join-Path $tmpDir ("source" + [System.IO.Path]::GetExtension($sourcePath))
    }
    copyFileShared $sourcePath $copyPath
    $openPath = $copyPath

    # 読み取り専用・リンク更新なしで開く。
    # パスワード付きのファイルは、ダイアログを出さずにエラーとするためダミーのパスワードを渡す
    $workbooks = (getApp "Excel").Workbooks
    try {
        $wb = $workbooks.Open($openPath, 0, $true, [Type]::Missing, "dummy", "dummy", $true)
    } finally {
        releaseComObject $workbooks
    }

    $sheets = New-Object System.Collections.Generic.List[object]
    try {
        $worksheets = $wb.Worksheets
        foreach ($ws in @($worksheets)) {
            try {
                # 非表示シートは除外（-1 = xlSheetVisible。0 = 非表示、2 = 完全に非表示）
                if ($ws.Visible -ne -1) {
                    continue
                }

                # Excelは [ ] を含むパスに保存できないため、一時ファイル名で保存し、整形時にリネームする
                $tmpPath = Join-Path $tmpDir ("sheet{0}.tmp" -f $ws.Index)
                $tsvPath = Join-Path $tmpDir (toIndexFileName $ws.Name)

                # 使用範囲が実際のデータよりずっと広いシートは、データの範囲だけを一時シートにコピーしてから書き出す
                $temp = copyDataRangeToTempSheet $wb $ws
                $target = if ($temp) { $temp } else { $ws }
                try {
                    # Excelは使用範囲（UsedRange）の左上のセルから出力するため、整形時にA1からの位置に戻せるよう控えておく。
                    # 一時シートは書式だけのセルが無く、元のシートより使用範囲が狭いことがあるため、書き出すシートから取る
                    $usedRange = $target.UsedRange
                    try {
                        $firstRow = $usedRange.Row
                        $firstColumn = $usedRange.Column
                    } finally {
                        releaseComObject $usedRange
                    }

                    [void]$target.Activate()
                    [void]$target.SaveAs($tmpPath, 42)  # 42 = xlUnicodeText
                } finally {
                    if ($temp) {
                        # 一時シートは、次のシートの Index がずれないようすぐ削除する（ブックは保存せずに閉じる）
                        $temp.Delete()
                        releaseComObject $temp
                    }
                }
                $sheets.Add(@($tmpPath, $tsvPath, $firstRow, $firstColumn))
            } finally {
                releaseComObject $ws
            }
        }
        releaseComObject $worksheets
    } finally {
        # 保存したファイルはブックを閉じるまでロックされている
        $wb.Close($false)
        releaseComObject $wb
        Remove-Item -LiteralPath $copyPath -Force
    }

    $count = 0
    foreach ($sheet in $sheets) {
        if (prettyTsv $sheet[0] $sheet[1] $sheet[2] $sheet[3]) {
            $count++
        }
        Remove-Item -LiteralPath $sheet[0] -Force
    }
    return $count
}

function convertWithWord {
    # Wordで開き、.docx 形式で保存する（旧形式 .doc 等を読めるようにするため）
    param (
        [string]$sourcePath,
        [string]$destPath
    )

    # 読み取り専用で開く。パスワード付きのファイルは、ダイアログを出さずにエラーとするためダミーのパスワードを渡す
    #   引数: FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate,
    #         Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible
    $documents = (getApp "Word").Documents
    try {
        $doc = $documents.Open($sourcePath, $false, $true, $false, "dummy", "dummy", $false, "dummy", "dummy", [Type]::Missing, [Type]::Missing, $false)
    } finally {
        releaseComObject $documents
    }

    try {
        # 保存時のページ区切り（ページ番号の目安）を正しく記録させるため、ページ割りを確定させてから保存する
        $doc.Repaginate()
        [void]$doc.SaveAs2($destPath, 12)  # 12 = wdFormatXMLDocument（.docx）
    } finally {
        $doc.Close(0)
        releaseComObject $doc
    }
}

function convertWithPowerPoint {
    # PowerPointで開き、.pptx 形式で保存する（旧形式 .ppt 等を読めるようにするため）
    param (
        [string]$sourcePath,
        [string]$destPath
    )

    # 読み取り専用・ウィンドウ無しで開く。
    # ファイル名の後ろに "::<パスワード>::" を付けると、パスワード付きのファイルはダイアログを出さずにエラーになる
    $presentations = (getApp "PowerPoint").Presentations
    try {
        $pres = $presentations.Open("${sourcePath}::dummy::", -1, 0, 0)  # ReadOnly, Untitled = False, WithWindow = False
    } finally {
        releaseComObject $presentations
    }

    try {
        [void]$pres.SaveAs($destPath, 24)  # 24 = ppSaveAsOpenXMLPresentation（.pptx）
    } finally {
        $pres.Close()
        releaseComObject $pres
    }
}

function convertDocument {
    # Word・PowerPointのファイルを場所（ページ・スライド）ごとに作業フォルダへTSV出力し、出力した数を返す
    param (
        [string]$sourcePath
    )

    $isWord = ((getAppName $sourcePath) -eq "Word")

    # 元のファイルを占有しないよう、作業フォルダにコピーしてからコピーを読む（読んでいる間も、利用者が上書き保存・移動できる）。
    # 新形式（ZIP）はコピーをそのまま読む。
    # 旧形式・パスワード付き・拡張子と中身が異なるファイルは、Word・PowerPointで新形式に変換してから読む
    $copyPath = Join-Path $tmpDir ("source" + [System.IO.Path]::GetExtension($sourcePath))
    $readPath = $copyPath
    $workFiles = @($copyPath)
    try {
        copyFileShared $sourcePath $copyPath
        if (!(isZipFile $copyPath)) {
            # PowerPointは、プレゼンテーションではないファイル（中身がテキスト等）もアウトラインとして開き、
            # 文字化けした内容になるため、旧形式（複合ドキュメント形式）でなければ変換しない。
            # Wordはテキスト・HTML・RTFも正しく読めるため、そのまま Word で開く
            if (!$isWord -and !(isCompoundFile $copyPath)) {
                throw "ファイルが壊れているか、PowerPointのファイルではありません（新形式（ZIP）でも旧形式でもない内容です）。"
            }

            # Word・PowerPointは拡張子と中身が異なるファイル（中身が .doc の .docx 等）を開けないため、
            # コピーに旧形式の拡張子を付け直してから開く
            if ($isWord) {
                $legacyPath = Join-Path $tmpDir "source.doc"
                $readPath = Join-Path $tmpDir "converted.docx"
            } else {
                $legacyPath = Join-Path $tmpDir "source.ppt"
                $readPath = Join-Path $tmpDir "converted.pptx"
            }
            if ($legacyPath -ne $copyPath) {
                [System.IO.File]::Move($copyPath, $legacyPath)
                $copyPath = $legacyPath
            }
            $workFiles = @($copyPath, $readPath)

            if ($isWord) {
                convertWithWord $copyPath $readPath
            } else {
                convertWithPowerPoint $copyPath $readPath
            }
        }

        if ($isWord) {
            $units = readDocxUnits $readPath
        } else {
            $units = readPptxUnits $readPath
        }
    } finally {
        foreach ($path in $workFiles) {
            if (Test-Path -LiteralPath $path) {
                Remove-Item -LiteralPath $path -Force
            }
        }
    }

    return (writeUnits $units $tmpDir)
}

function convertFile {
    # 1ファイルを変換し、作成したTSVの数を返す
    param (
        [string]$sourcePath
    )

    if ((getAppName $sourcePath) -eq "Excel") {
        return (convertWorkbook $sourcePath)
    }
    return (convertDocument $sourcePath)
}

function publishTsv {
    # 作業フォルダのTSVを、そのファイルの変換結果のフォルダへ移動する。
    # 途中で強制終了されても一部のシートだけのインデックスが残らないよう、
    # 出力用のフォルダ（work\変換出力\<PID>）に集めてからフォルダごと入れ替える（publishIndexFiles）
    param (
        [string]$bookDir
    )

    publishIndexFiles $tmpDir $bookDir (Join-Path ${publishDir} ([System.IO.Path]::GetFileName($bookDir)))
}

function clearTmpDir {
    Get-ChildItem -LiteralPath (toLongPath $tmpDir) -File | Remove-Item -Force
}

function removeTmpDir {
    # 作業フォルダ（%TEMP%\win_grep\<PID>）と出力用のフォルダ（work\変換出力\<PID>）を削除する。終了時に呼ぶ
    foreach ($dir in @(${tmpDir}, ${publishDir})) {
        try {
            removeDirectoryRetry $dir
        } catch {
            Write-Host "    作業フォルダを削除できませんでした: ${dir}" -ForegroundColor Yellow
        }
    }
}

function removeStaleTmpDirs {
    # 強制終了などで残った、ほかの（終了済みの）プロセスの作業フォルダ
    # （%TEMP%\win_grep\<PID>・work\変換出力\<PID>）を削除する
    foreach ($parent in @((Split-Path ${tmpDir} -Parent), (Split-Path ${publishDir} -Parent))) {
        removeStaleProcessDirs $parent
    }
}

function removeStaleProcessDirs {
    # プロセスIDの名前のフォルダのうち、そのプロセスが既に終わっているものを削除する
    param (
        [string]$parent
    )

    if (!(Test-Path -LiteralPath $parent)) {
        return
    }
    foreach ($dir in @(Get-ChildItem -LiteralPath $parent -Directory | Where-Object { $_.Name -match "^\d{1,9}$" -and [int]$_.Name -ne $PID })) {
        if (Get-Process -Id ([int]$dir.Name) -ErrorAction SilentlyContinue) {
            continue  # 実行中の変換（またはPIDを再利用した別のプロセス）のものは残す
        }
        try {
            Remove-Item -LiteralPath (toLongPath $dir.FullName) -Recurse -Force
        } catch {
            # 使用中などで削除できなければ、次回に回す
        }
    }
}

function moveLegacyIndex {
    # 以前の形式（work\index 直下に変換対象フォルダ1つ分のインデックスがある）を、そのフォルダのインデックス名の下へ移す。
    # 前回の変換一覧の行（相対パス → 行）を、移した後の相対パス（"インデックス名\…"）で返す
    param (
        [object[]]$folders,  # assignIndexNames の結果
        $status,             # readStatusFile の結果
        [bool]$statusExists
    )

    $legacyPath = $null
    $legacyFolder = @($status.Folders | Where-Object { -not $_.Name }) | Select-Object -First 1
    if ($legacyFolder) {
        $legacyPath = normalizeFolderPath $legacyFolder.Path
    } elseif (-not $statusExists) {
        # 変換一覧が無い（変換一覧を使う前の版）: work\index 直下が各フォルダのインデックス名のフォルダだけでなければ、
        # 以前の形式で、変換対象フォルダの1件目のフォルダのインデックスとみなす
        $names = @($folders | ForEach-Object { $_.Name })
        $others = @(Get-ChildItem -LiteralPath $indexDir -Force | Where-Object {
            ($_.PSIsContainer -and $names -notcontains $_.Name) -or (-not $_.PSIsContainer -and $_.Name -ne ${sourceFolderFileName})
        })
        if ($others.Count -gt 0) {
            $legacyPath = $folders[0].Path
        }
    }
    if (!$legacyPath) {
        return , $status.Rows
    }

    $rows = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    $folder = @($folders | Where-Object { $_.Path -eq $legacyPath }) | Select-Object -First 1
    if (!$folder) {
        Write-Host "work\index 直下に、変換対象から外したフォルダ（${legacyPath}）の以前の形式のインデックスがあります。不要なら削除してください。" -ForegroundColor Yellow
        return , $rows
    }

    # work\index を丸ごと work\index\<インデックス名> に移す（同じ名前のサブフォルダがあっても衝突しないよう、いったん別名にする）
    $movingDir = "${indexDir}_移行中"
    [System.IO.Directory]::Move($indexDir, $movingDir)
    [System.IO.Directory]::CreateDirectory($indexDir) | Out-Null
    [System.IO.Directory]::Move($movingDir, (Join-Path $indexDir $folder.Name))
    Write-Host "以前の形式のインデックスを work\index\$($folder.Name) に移しました。（$($folder.Path) のインデックス）"

    foreach ($row in $status.Rows.Values) {
        $row.相対パス = "$($folder.Name)\$($row.相対パス)"
        $rows[$row.相対パス] = $row
    }
    return , $rows
}

function removeDroppedFolders {
    # 変換対象フォルダから削除されたフォルダのインデックス（work\index\<インデックス名>）を削除する。
    # チェックを外しただけのフォルダは削除しない
    param (
        [object[]]$folders,          # assignIndexNames の結果
        [object[]]$previousFolders   # readStatusFile の Folders
    )

    # インデックス名で比べる。フォルダを移動して登録し直した場合は、同じ名前を引き継ぐため削除しない（assignIndexNames）
    $current = @($folders | ForEach-Object { $_.Name })
    foreach ($previous in @($previousFolders | Where-Object { $_.Name -and $current -notcontains $_.Name })) {
        $dir = Join-Path $indexDir $previous.Name
        if (Test-Path -LiteralPath $dir) {
            # 中に長いパス（260文字超）のTSVがあっても削除できるよう \\?\ 付きで削除する
            Remove-Item -LiteralPath (toLongPath $dir) -Recurse -Force
        }
        Write-Host "変換対象から削除されたフォルダ（$($previous.Path)）のインデックスを削除しました。"
    }
}

function migrateFlatIndex {
    # 以前の形式のTSV（<ファイル名>_<場所>.tsv）を、今の形式（<ファイル名>\<場所>.tsv）へ移す。
    # 変換はし直さず、名前を変えるだけ（更新日時もそのまま）。
    # 今の形式のTSVの名前は場所だけ（_ は符号化されている）のため、_ を含む名前が以前の形式
    $moved = 0
    $failed = 0
    foreach ($file in @(Get-ChildItem -LiteralPath (toLongPath $indexDir) -Filter "*.tsv" -File -Recurse -ErrorAction SilentlyContinue)) {
        if ($file.Name.IndexOf("_") -lt 0) {
            continue
        }
        $name = splitIndexFileName $file.Name
        if ($name.book -eq $file.Name) {
            continue  # ファイル名と場所に分けられないものは触らない
        }

        $bookDir = Join-Path ([System.IO.Path]::GetDirectoryName((fromLongPath $file.FullName))) $name.book
        try {
            $dest = Join-Path $bookDir (toIndexFileName $name.sheet)
            [System.IO.Directory]::CreateDirectory((toLongPath $bookDir)) | Out-Null
            if (Test-Path -LiteralPath (toLongPath $dest)) {
                Remove-Item -LiteralPath (toLongPath $dest) -Force
            }
            [System.IO.File]::Move($file.FullName, (toLongPath $dest))
            $moved++
        } catch {
            $failed++
        }
    }

    if ($moved -gt 0) {
        Write-Host "以前の形式のTSV ${moved} 件を、元のファイル名のフォルダへ移しました。（変換し直しません）"
    }
    if ($failed -gt 0) {
        Write-Host "以前の形式のTSV ${failed} 件は移せませんでした。該当のファイルは次の変換で作り直します。" -ForegroundColor Yellow
    }
}

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
writeConvertProgress ${convertPhaseScan} 0 0 0 "変換済みのインデックスを確認しています…"
$indexCounts = getIndexTsvCounts
if ($null -eq $indexCounts) {
    Write-Host "  インデックスのフォルダを調べられないため、変換結果が残っているかの確認は行いません。" -ForegroundColor Yellow
}
$rows = New-Object System.Collections.Generic.List[object]
$targets = New-Object System.Collections.Generic.List[object]
$failed = New-Object System.Collections.Generic.List[object]
foreach ($folder in $folders) {
    if (-not $folder.Enabled) {
        Write-Host "  [$($folder.Name)] $($folder.Path) … チェックなしのため変換しません（インデックスはそのまま残します）"
    } elseif (!(Test-Path -LiteralPath $folder.Path -PathType Container)) {
        Write-Host "  [$($folder.Name)] $($folder.Path) … フォルダが見つからないため変換しません" -ForegroundColor Yellow
    } else {
        Write-Host "  [$($folder.Name)] $($folder.Path)"
        # 大きいフォルダ・ネットワーク越しでは時間がかかるため、どのフォルダを見ているかを画面に伝える
        writeConvertProgress ${convertPhaseScan} 0 0 0 "[$($folder.Name)] のOfficeファイルを探しています… $($folder.Path)"
        $list = createTargetList $folder $previous $indexCounts
        $rows.AddRange($list.Rows)
        $targets.AddRange($list.Targets)
        $failed.AddRange($list.Failed)
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

if ($failed.Count -gt 0) {
    Write-Host ""
    if ($RetryFailed) {
        Write-Host "前回変換に失敗し、その後更新されていないファイル $($failed.Count) 件も再変換します。"
        $targets.AddRange($failed)
    } else {
        Write-Host "前回変換に失敗し、その後更新されていないファイル $($failed.Count) 件はスキップします。（パスワード付きなど）"
    }
}

# 前回、変換中に強制終了した（ウィンドウを閉じた・PCが停止した等）ファイルは、同じファイルで止まり続けないよう最後に回す。
# 続けて $interruptLimit 回強制終了したファイルは、応答しなくなるファイルとみなして失敗とする
$interrupted = readConvertingFile
if ($interrupted) {
    $row = @($targets | Where-Object { $_.相対パス -eq $interrupted.RelPath }) | Select-Object -First 1
    if (!$row) {
        $interrupted = $null
        removeConvertingFile
    } elseif ($interrupted.Count -ge $interruptLimit) {
        [void]$targets.Remove($row)
        $row.状態 = ${stateFailed}
        $row.TSV数 = ""
        $row.変換日時 = formatFileTime (Get-Date)
        $row.エラー = "変換中に $($interrupted.Count) 回続けて強制終了されたため、変換を中止しました（Officeアプリが応答しなくなる可能性があります）"
        Write-Host ""
        Write-Host "変換中に $($interrupted.Count) 回続けて強制終了したファイルは、失敗としてスキップします: $($row.相対パス)" -ForegroundColor Yellow
        $interrupted = $null
        removeConvertingFile
    } else {
        [void]$targets.Remove($row)
        $targets.Add($row)
        Write-Host ""
        Write-Host "前回、変換中に強制終了したファイルは最後に変換します: $($row.相対パス)" -ForegroundColor Yellow
    }
}
writeConvertProgress ${convertPhaseScan} 0 $targets.Count 0 "変換一覧を書き出しています…"
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
    writeConvertProgress ${convertPhaseFinish} 0 0 0 "変換が必要なファイルはありませんでした"
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
        writeConvertProgress ${convertPhaseRun} $i ($total - $i) $failures.Count $relPath

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
        writeConvertingFile $relPath $startCount

        try {
            clearTmpDir
            $script:watchdog.TimedOut = $false
            $script:watchdog.Deadline = (Get-Date).AddMinutes($fileTimeoutMinutes)
            try {
                $tsvCount = convertFile $sourcePath
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
            $message = describeConvertError $_.Exception
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
        removeConvertingFile

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
    writeConvertProgress ${convertPhaseFinish} $processed 0 $failures.Count "Officeアプリを終了しています…"
    stopWatchdog
    stopAllApps
    removeTmpDir
    removeConvertingFile
    writeConvertProgress ${convertPhaseFinish} $processed 0 $failures.Count "変換一覧を書き直しています…"
    # 変換の直前に無くなっていたファイルの行は除く（次回の検索でも見つからず、インデックスも削除済み）
    writeStatusFile $folders @($rows | Where-Object { $_ -and !$droppedRows.Contains([string]$_.相対パス) })
    # 初めて変換したインデックスは、最初に書き出した時点ではまだフォルダが無いため、ここでもう一度書く
    # （work\index\<インデックス名>\元のフォルダ.txt。インデックス 1 個だけをコピーしても元のファイルの場所が分かる）
    writeSourceFolderFile $folders
    # 画面が終わり方（成功・失敗の件数）を読めるよう、進み具合は消さずに最後の状態を残す
    writeConvertProgress ${convertPhaseFinish} $processed $remaining $failures.Count ""
}

if ($folderLost) {
    # 変換対象フォルダが見えなくなった場合は、続けられないエラーとして画面に知らせる（残りは未変換のまま）
    $message = "変換対象フォルダが見つからなくなったため、変換を中止しました: ${folderLost}" +
        "（残り ${remaining} 件は未変換のまま残しました。フォルダを使えるようにしてから、もう一度変換してください）"
    Write-Host ""
    Write-Host $message -ForegroundColor Red
    [System.IO.File]::WriteAllText(${convertErrorFile}, $message, ${utf8Bom})
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
