# 比較の抽出プロセス（画面がウィンドウ無しで起動する。入力は求めない）
#
#   powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -WindowStyle Hidden -File differ.ps1 -JobDir <比較の作業フォルダ>
#
# 比較の作業フォルダ（job\diff_job.ps1）の 抽出要求.tsv に書かれたファイルを順に抽出し、
# 場所ごとの TSV を <作業フォルダ>\<側>\<番号>\ に置く。抽出はインデクサと同じ（shared\office\office_extract.ps1）。
#   ・原本は作業フォルダへコピーし、読み取り専用・マクロ無効で開く
#   ・1 ファイルの抽出が ${diffFileTimeoutMinutes} 分を超えたら、自分で起動した Office を強制終了し、そのファイルを失敗にして次へ進む
#   ・1 ファイル終わるたびに 抽出結果.tsv に 1 行追記し、優先.tsv を読んで、画面が選んだファイルを次に抽出する
#   ・画面が 中止要求 を作ると、ファイルの切れ目で止まる
#   終了コード: 0 = 完了（ファイルごとの失敗は 抽出結果.tsv）/ 1 = 続けられないエラー（比較エラー.txt）/ 2 = 中止

param (
    [Parameter(Mandatory = $true)]
    [string]$JobDir
)

. "$PSScriptRoot\lib.ps1"
. "$PSScriptRoot\..\shared\office\office_reader.ps1"
. "$PSScriptRoot\..\shared\office\office_app.ps1"
. "$PSScriptRoot\..\shared\office\office_extract.ps1"

$ErrorActionPreference = "Stop"

trap {
    # 続けられないエラー。画面に伝えるため、メッセージをファイルに書いて終了する
    try {
        [System.IO.File]::WriteAllText((Join-Path $JobDir ${diffErrorFileName}), $_.Exception.Message, ${utf8Bom})
    } catch {}
    try { stopWatchdog } catch {}
    try { stopAllApps } catch {}
    exit 1
}

if (!(Test-Path -LiteralPath $JobDir -PathType Container)) {
    throw "比較の作業フォルダが見つかりません: $JobDir"
}

# 抽出の作業フォルダ（office_extract.ps1 はここにコピーを置き、TSV を書き出す）
${tmpDir} = Join-Path $JobDir "work"
[System.IO.Directory]::CreateDirectory(${tmpDir}) | Out-Null

$items = readExtractRequest $JobDir
$pending = New-Object System.Collections.Generic.List[object]
foreach ($item in $items) { $pending.Add($item) }
$total = $pending.Count
$doneCount = 0
$stopped = $false

startWatchdog
try {
    while ($pending.Count -gt 0) {
        if (testDiffStopRequested $JobDir) {
            $stopped = $true
            break
        }

        # 画面が選んだファイル（優先.tsv の上から）があれば、それを先に抽出する
        $index = 0
        foreach ($key in (readDiffPriority $JobDir)) {
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
        writeDiffProgress $JobDir $doneCount $total $item.Path

        # 前のファイルの作業ファイルが残っていれば消す（強制終了した Office が掴んでいたもの等）
        Get-ChildItem -LiteralPath (toLongPath ${tmpDir}) -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
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
            $count = saveExtractedUnits ${tmpDir} (getExtractDir $JobDir $item.Side $item.Id)
            addExtractResult $JobDir $item.Id $item.Side ${diffStateDone} $count
        } catch {
            $message = describeIngestError $_.Exception
            if ($script:watchdog.TimedOut) {
                $message = "${diffFileTimeoutMinutes} 分以内に読み取れなかったため中止しました（Officeアプリを強制終了しました）"
            }
            addExtractResult $JobDir $item.Id $item.Side ${diffStateFailed} 0 $message
            # アプリが不安定になっている可能性があるため終了する（次に必要になったときに起動し直す）
            try { stopApp (getAppName $item.Path) } catch {}
        }
        $doneCount++
        if ($script:watchdog.TimedOut) {
            # 制限時間を過ぎて強制終了したアプリは使えないため、すべて終了して次に必要になったときに起動し直す
            stopAllApps
        }
    }
} finally {
    stopWatchdog
    stopAllApps
    try { removeDirectoryRetry ${tmpDir} } catch {}
}

writeDiffProgress $JobDir $doneCount $total ""
if ($stopped) {
    exit 2
}
exit 0
