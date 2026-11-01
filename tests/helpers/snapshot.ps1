# 抽出結果の期待値（snapshot）と比べる助け関数。使うテストが dot-source する（読み込み口には入れない）。
#   . "$PSScriptRoot\..\..\helpers\snapshot.ps1"
# 期待値は tests\testdata\snapshot\に UTF-8（BOM なし）・LF の .txt で置く。
# 環境変数 TEBUNKO_UPDATE_SNAPSHOT=1 のときは比べずに期待値を書き直す（形式の版を上げる・取り出しを変える PR が、差分を PR に載せるために使う）。

function assertSnapshot {
    # actual を expectedPath の期待値と比べる。改行は LF にそろえて比べ、違えば最初に違う行を入れた例外にする
    param (
        [string]$actual,
        [string]$expectedPath
    )

    $actualText = $actual.Replace("`r`n", "`n")
    if ($env:TEBUNKO_UPDATE_SNAPSHOT -eq "1") {
        [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($expectedPath)) | Out-Null
        [System.IO.File]::WriteAllText($expectedPath, $actualText, (New-Object System.Text.UTF8Encoding($false)))
        return
    }
    if (![System.IO.File]::Exists($expectedPath)) {
        throw "期待値のファイルがありません: $expectedPath（TEBUNKO_UPDATE_SNAPSHOT=1 で作れます）"
    }
    $expectedText = [System.IO.File]::ReadAllText($expectedPath, (New-Object System.Text.UTF8Encoding($false))).Replace("`r`n", "`n")
    if ($actualText -ceq $expectedText) {
        return
    }
    $actualLines = $actualText.Split("`n")
    $expectedLines = $expectedText.Split("`n")
    $limit = [Math]::Max($actualLines.Count, $expectedLines.Count)
    for ($i = 0; $i -lt $limit; $i++) {
        $a = if ($i -lt $actualLines.Count) { $actualLines[$i] } else { "(行なし)" }
        $e = if ($i -lt $expectedLines.Count) { $expectedLines[$i] } else { "(行なし)" }
        if ($a -cne $e) {
            throw "期待値と違います（$expectedPath の $($i + 1) 行目）。期待値: [$e] 実際: [$a]"
        }
    }
}
