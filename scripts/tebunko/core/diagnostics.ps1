# tebunko のログの先頭に書く見出し（状態層）。shared/ がツールの名前を知らないよう、名前はここで渡す。

function getToolDiagnosticHeader {
    # 見出しの行（配列）。版は配布物の VERSION.txt（無ければ開発版。.git のファイルから分かるタグとコミットを添える）
    $version = readVersionFile "${rootDir}\VERSION.txt"
    $devInfo = if ($null -eq $version) { readDevVersion ${rootDir} } else { $null }
    return formatDiagnosticHeader "tebunko" $version (getEnvironmentInfo) ([System.DateTimeOffset]::Now) $devInfo
}
