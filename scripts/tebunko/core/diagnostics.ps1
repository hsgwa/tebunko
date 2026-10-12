# tebunko のログの先頭に書く見出し（状態層）。shared/ がツールの名前を知らないよう、名前はここで渡す。

function getToolDiagnosticHeader {
    # 見出しの行（配列）。版は配布物の VERSION.txt（無ければ開発版）
    return formatDiagnosticHeader "tebunko" (readVersionFile "${rootDir}\VERSION.txt") (getEnvironmentInfo) ([System.DateTimeOffset]::Now)
}
