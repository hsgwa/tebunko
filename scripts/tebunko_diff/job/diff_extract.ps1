# 比較の抽出の本体（状態層）。抽出プロセス（..\differ.ps1）が読み込んで invokeDiffer を呼ぶ。
# テストから同じプロセスの中で呼べるよう、起動口と分けてある。
# 使う側は ..\lib.ps1 と、shared\office の office_reader.ps1・office_app.ps1・office_extract.ps1 を読み込んでおくこと

function invokeDiffer {
    # 作業フォルダ（jobDir）の 抽出要求.tsv に書かれたファイルを順に抽出し、終了コード（0 = 完了、2 = 中止）を返す。
    #   ・原本は作業フォルダへコピーし、読み取り専用・マクロ無効で開く（office_extract.ps1）
    #   ・1 ファイルの抽出が ${diffFileTimeoutMinutes} 分を超えたら、自分で起動した Office を強制終了し、そのファイルを失敗にして次へ進む
    #   ・1 ファイル終わるたびに 抽出結果.tsv に 1 行追記し、優先.tsv を読んで、画面が選んだファイルを次に抽出する
    #   ・抽出要求は、画面が後から書き足すことがある（フォルダ同士で、片方にしか無いファイルを選んだとき）。
    #     ファイルの切れ目ごとに読み直し、まだ結果の無いものを足す
    #   ・画面が 中止要求 を作ると、ファイルの切れ目で止まる
    # 続けられないエラー（作業フォルダが無い等）は例外にする
    param (
        [string]$jobDir
    )

    if (!(Test-Path -LiteralPath $jobDir -PathType Container)) {
        throw "比較の作業フォルダが見つかりません: $jobDir"
    }

    # 抽出の作業フォルダ（office_extract.ps1 はここにコピーを置き、TSV を書き出す）
    $script:tmpDir = Join-Path $jobDir "work"
    [System.IO.Directory]::CreateDirectory($script:tmpDir) | Out-Null

    $pending = New-Object System.Collections.Generic.List[object]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($key in @((readExtractResults $jobDir).Keys)) { [void]$seen.Add($key) }
    $doneCount = 0
    $stopped = $false

    startWatchdog
    try {
        addDifferRequests $jobDir $pending $seen
        while ($pending.Count -gt 0) {
            if (testDiffStopRequested $jobDir) {
                $stopped = $true
                break
            }
            $total = $doneCount + $pending.Count
            $item = takeDifferItem $jobDir $pending
            writeDiffProgress $jobDir $doneCount $total $item.Path
            extractDifferItem $jobDir $item
            $doneCount++
            if ($script:watchdog.TimedOut) {
                # 制限時間を過ぎて強制終了したアプリは使えないため、すべて終了して次に必要になったときに起動し直す
                stopAllApps
            }
            addDifferRequests $jobDir $pending $seen
        }
    } finally {
        stopWatchdog
        stopAllApps
        try { removeDirectoryRetry $script:tmpDir } catch {}
    }

    writeDiffProgress $jobDir $doneCount ($doneCount + $pending.Count) ""
    if ($stopped) {
        return 2
    }
    return 0
}

function addDifferRequests {
    # 抽出要求を読み直し、まだ見ていない（結果の無い）ものを pending に足す
    param (
        [string]$jobDir,
        $pending,
        $seen
    )

    foreach ($item in (readExtractRequest $jobDir)) {
        if ($seen.Add("$($item.Id)|$($item.Side)")) {
            $pending.Add($item)
        }
    }
}

function takeDifferItem {
    # 次に抽出するものを pending から取り出す。画面が選んだファイル（優先.tsv の上から）があれば、それを先にする
    param (
        [string]$jobDir,
        $pending
    )

    $index = 0
    foreach ($key in (readDiffPriority $jobDir)) {
        $found = -1
        for ($i = 0; $i -lt $pending.Count; $i++) {
            if ("$($pending[$i].Id)|$($pending[$i].Side)" -eq $key) { $found = $i; break }
        }
        if ($found -ge 0) {
            $index = $found
            break
        }
    }
    $item = $pending[$index]
    $pending.RemoveAt($index)
    return $item
}

function extractDifferItem {
    # 1 ファイルを抽出し、結果を 抽出結果.tsv に追記する（失敗も結果として書く）
    param (
        [string]$jobDir,
        $item
    )

    # 前のファイルの作業ファイルが残っていれば消す（強制終了した Office が掴んでいたもの等）
    Get-ChildItem -LiteralPath (toLongPath $script:tmpDir) -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
    $script:watchdog.TimedOut = $false
    try {
        if (!(Test-Path -LiteralPath (toLongPath $item.Path) -PathType Leaf)) {
            throw (New-Object System.IO.FileNotFoundException "ファイルが見つかりません: $($item.Path)")
        }
        $script:watchdog.Deadline = (Get-Date).AddMinutes(${diffFileTimeoutMinutes})
        try {
            [void](extractOfficeFile $item.Path)
        } finally {
            $script:watchdog.Deadline = [datetime]::MaxValue
        }
        $count = saveExtractedUnits $script:tmpDir (getExtractDir $jobDir $item.Side $item.Id)
        addExtractResult $jobDir $item.Id $item.Side ${diffStateDone} $count
    } catch {
        $message = describeIngestError $_.Exception
        if ($script:watchdog.TimedOut) {
            $message = "${diffFileTimeoutMinutes} 分以内に読み取れなかったため中止しました（Officeアプリを強制終了しました）"
        }
        addExtractResult $jobDir $item.Id $item.Side ${diffStateFailed} 0 $message
        # アプリが不安定になっている可能性があるため終了する（次に必要になったときに起動し直す）
        try { stopApp (getAppName $item.Path) } catch {}
    }
}
