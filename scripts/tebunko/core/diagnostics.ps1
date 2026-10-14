# tebunko のログの先頭に書く見出しと、インデックス作成のログの書き口（状態層）。shared/ がツールの名前を知らないよう、名前はここで渡す。

function getToolDiagnosticHeader {
    # 見出しの行（配列）。版は配布物の VERSION.txt（無ければ開発版。.git のファイルから分かるタグとコミットを添える）
    $version = readVersionFile "${rootDir}\VERSION.txt"
    $devInfo = if ($null -eq $version) { readDevVersion ${rootDir} } else { $null }
    return formatDiagnosticHeader "tebunko" $version (getEnvironmentInfo) ([System.DateTimeOffset]::Now) $devInfo
}

# インデックス作成のログ（invokeIndexer が開く TextWriter。取り込みのスレッドでは 1 ファイル分を貯める StringWriter）
$script:indexerLog = $null
# ログをコンソールにも出すか（画面を使わずに indexer.ps1 を実行したとき）
$script:indexerEcho = $false

function writeIndexerLog {
    # インデックス作成の表示内容をログ（indexing_log.txt）に書く。color はコンソールに出すときの色
    param (
        [string]$text = "",
        [string]$color = ""
    )

    if ($null -ne $script:indexerLog) {
        try {
            $script:indexerLog.WriteLine($text)
        } catch {
            # ログのためだけの書き込みのため、書けなくてもインデックス作成は続ける
        }
    }
    if ($script:indexerEcho) {
        if ($color) {
            Write-Host $text -ForegroundColor $color
        } else {
            Write-Host $text
        }
    }
}
