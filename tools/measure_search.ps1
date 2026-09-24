# 検索のどこに時間がかかっているかを、段階ごとに測る（画面の検索と同じ関数を、同じ順に呼ぶ）。
#
#   .\tools\measure_search.ps1                          work\index のすべてを、存在しないワードで検索して測る
#   .\tools\measure_search.ps1 -Words 見積,山田商事     ふだん使うワードでも測る
#   .\tools\measure_search.ps1 -SkipPack                1 ファイルにまとめたときの読み込み（6 章）を測らない
#
# 測る段階:
#   1. 環境            … PowerShell・CPU・インデックスのドライブの種類・Defender のリアルタイム保護
#   2. 画面と同じ検索  … 画面の検索スレッドと同じ流れ（スレッドの用意と lib.ps1 の読み込み → 列挙 → 検索）を 2 回
#   3. 準備の内訳      … スレッドの用意と lib.ps1 の読み込み・列挙（.NET だけ／getIndexTsvFiles）・newTsvFiles・並列検索のスレッドの用意
#   4. 読み込みの内訳  … TSV を開いて閉じるだけ・開いて全文を読む（1 スレッド。1 回目と 2 回目）
#   5. 照合の内訳      … 読んだ内容を使い回して検索（ファイルを開かない。1 スレッドと並列）
#   6. まとめた場合    … 同じ内容を 1 ファイルにまとめて %TEMP% に書き、読む（ファイルの数による差を見る。終わったら消す）
#
# 最初に TSV を開く段階（2 の 1 回目）が、Windows の起動直後・インデックス作成の直後にいちばん近い。
# PC を再起動した直後に 1 回、続けてもう 1 回実行すると、ファイルを開く検査（Defender など）の影響が分かる。
# 結果には時間と件数だけを出し、パス・ファイル名は出さない（そのまま共有できる）。インデックスは書き換えない。
param (
    [string[]]$Words = @(),
    [string]$Folder,
    [switch]$SkipPack,
    [string]$OutFile
)

$ErrorActionPreference = "Stop"
$repoDir = Split-Path $PSScriptRoot -Parent
$libFile = Join-Path $repoDir "scripts\tebunko_grep\lib.ps1"
. $libFile
if (!$Folder) {
    $Folder = ${indexDir}
}
if (!$OutFile) {
    $OutFile = Join-Path ${workDir} "検索の計測.txt"
}
# 存在しないワード（ヒットの処理を含まない、TSV を読んで照合するだけの時間）を先頭にする
$Words = @("tebunko計測用の存在しない語") + @($Words | Where-Object { $_ })

$report = New-Object System.Collections.Generic.List[string]
function out {
    param ([string]$text)
    $report.Add($text)
    Write-Host $text
}
function ms {
    param ([System.Diagnostics.Stopwatch]$watch)
    return $watch.Elapsed.TotalMilliseconds.ToString("#,0").PadLeft(8) + " ms"
}
$watch = New-Object System.Diagnostics.Stopwatch

# ---- 1. 環境 ----
out "# 検索の計測 $((Get-Date).ToString('yyyy-MM-dd HH:mm'))"
out ""
out "## 1. 環境"
out "PowerShell $($PSVersionTable.PSVersion) / CLR $([Environment]::Version) / 論理コア $([Environment]::ProcessorCount)"
$fullFolder = (Resolve-Path -LiteralPath $Folder).ProviderPath
$driveKind = if ($fullFolder.StartsWith("\\")) { "ネットワーク（UNC）" } else {
    try { [string]([System.IO.DriveInfo]::new([System.IO.Path]::GetPathRoot($fullFolder)).DriveType) } catch { "不明" }
}
out "インデックスのドライブ: $driveKind"
try {
    $mp = Get-MpComputerStatus -ErrorAction Stop
    out "Defender: リアルタイム保護 $($mp.RealTimeProtectionEnabled) / 読み取り時の検査 $($mp.OnAccessProtectionEnabled)"
} catch {
    out "Defender: 状態を取得できません（別のウイルス対策ソフト、または取得の権限なし）"
}

# ---- 2. 画面と同じ検索 ----
# 画面（search_tab.ps1 の searchScript）と同じく、新しいスレッドで lib.ps1 を読み込み、列挙・検索する。
# 読んだ内容の使い回し（cache）は画面と同じく 1 つを持ち続ける（1 回目は空、2 回目は 1 回目に読んだ内容を使う）
$cache = newTsvTextCache
$guiScript = {
    param ($libPath, $folder, $word, $cache)
    $total = [System.Diagnostics.Stopwatch]::StartNew()
    $step = [System.Diagnostics.Stopwatch]::StartNew()
    . $libPath
    $load = $step.Elapsed.TotalMilliseconds
    $step.Restart()
    $index = getIndexTsvFiles @($folder)
    $list = $step.Elapsed.TotalMilliseconds
    $step.Restart()
    $result = searchIndex $word $index.Files $true 10000 50 -cache $cache
    @{ Load = $load; List = $list; Search = $step.Elapsed.TotalMilliseconds; Total = $total.Elapsed.TotalMilliseconds
       Files = $index.Files.Count; Hits = $result.Hits.Count; Truncated = $result.Truncated }
}
out ""
out "## 2. 画面と同じ検索（スレッドの用意から検索の終わりまで。画面への表示は含まない）"
out "| 回 | ワード | TSV | ヒット | 合計 | うち lib.ps1 の読み込み | うち列挙 | うち検索 |"
out "|---|---|---|---|---|---|---|---|"
foreach ($round in 1..2) {
    foreach ($word in $Words) {
        $ps = [powershell]::Create()
        try {
            [void]$ps.AddScript($guiScript.ToString()).AddArgument($libFile).AddArgument($fullFolder).AddArgument($word).AddArgument($cache)
            $watch.Restart()
            $r = $ps.Invoke()[0]
            $watch.Stop()
        } finally {
            $ps.Dispose()
        }
        $label = if ($word -eq $Words[0]) { "（存在しない語）" } else { "ワード $([array]::IndexOf($Words, $word))" }
        $hits = if ($r.Truncated) { "$($r.Hits)+" } else { "$($r.Hits)" }
        out ("| {0} | {1} | {2:N0} | {3} | {4} | {5:N0} ms | {6:N0} ms | {7:N0} ms |" -f $round, $label, $r.Files, $hits, (ms $watch).Trim(), $r.Load, $r.List, $r.Search)
    }
}

# ---- 3. 準備の内訳 ----
out ""
out "## 3. 準備の内訳"
foreach ($i in 1..3) {
    $ps = [powershell]::Create()
    try {
        [void]$ps.AddScript(". '$($libFile.Replace("'", "''"))'")
        $watch.Restart()
        [void]$ps.Invoke()
        $watch.Stop()
    } finally {
        $ps.Dispose()
    }
    out "新しいスレッドの用意と lib.ps1 の読み込み（$i 回目）: $(ms $watch)"
}
$longFolder = toLongPath $fullFolder
$watch.Restart()
$rawCount = 0
foreach ($path in [System.IO.Directory]::EnumerateFiles($longFolder, "*.tsv", [System.IO.SearchOption]::AllDirectories)) {
    $rawCount++
}
$watch.Stop()
out "列挙（.NET だけ・1 スレッド・$($rawCount.ToString('N0')) 件）: $(ms $watch)"
$watch.Restart()
$found = findTsvFilesParallel $longFolder
$watch.Stop()
out "列挙（findTsvFilesParallel。更新日時・サイズ付き）: $(ms $watch)"
$watch.Restart()
$index = getIndexTsvFiles @($fullFolder)
$watch.Stop()
out "列挙（getIndexTsvFiles。上に加えてパスの組み立て・並べ替え）: $(ms $watch)"
$watch.Restart()
$files = newTsvFiles $index.Files
$watch.Stop()
out "newTsvFiles（TSV ごとに元のファイル名・場所を求める）: $(ms $watch)"
$workers = [Math]::Min([Environment]::ProcessorCount, ${searchWorkerMax})
$watch.Restart()
$pool = newSearchWorkerPool $workers
$watch.Stop()
out "並列検索のスレッドの用意（$workers スレッド）: $(ms $watch)"
# 実際に各スレッドが動き出すまで（関数の読み込みはスレッドごとに初めて使うときに起きる）
$watch.Restart()
$jobs = foreach ($i in 1..$workers) {
    $ps = [powershell]::Create()
    $ps.RunspacePool = $pool
    [void]$ps.AddScript({ 1 })
    @{ PowerShell = $ps; Handle = $ps.BeginInvoke() }
}
foreach ($job in $jobs) {
    [void]$job.PowerShell.EndInvoke($job.Handle)
    $job.PowerShell.Dispose()
}
$watch.Stop()
out "並列検索のスレッドが初めて動くまで: $(ms $watch)"
$watch.Restart()
foreach ($i in 1..200) {
    $ps = [powershell]::Create()
    $ps.RunspacePool = $pool
    [void]$ps.AddScript(${searchWorkerScript}).AddArgument($files).AddArgument(0).AddArgument(0).AddArgument([regex]"x").AddArgument(-1).AddArgument($null).AddArgument("scan").AddArgument($null)
    [void]$ps.EndInvoke($ps.BeginInvoke())
    $ps.Dispose()
}
$watch.Stop()
out "並列検索の受け渡し 200 回（中身は空。TSV 1 万件を 50 件ずつ検索するときの回数）: $(ms $watch)"
$pool.Dispose()

$bytes = 0L
$maxBytes = 0L
foreach ($f in $files) {
    $bytes += $f.Size
    $maxBytes = [Math]::Max($maxBytes, [long]$f.Size)
}
out ("TSV: {0:N0} 件 / 合計 {1:N1} MB / 平均 {2:N1} KB / 最大 {3:N1} MB / 入れ物のフォルダ {4:N0} 件" -f $files.Count, ($bytes / 1MB), ($bytes / [Math]::Max(1, $files.Count) / 1KB), ($maxBytes / 1MB), @($files | ForEach-Object { [System.IO.Path]::GetDirectoryName($_.Path) } | Sort-Object -Unique).Count)

# ---- 4. 読み込みの内訳 ----
out ""
out "## 4. 読み込みの内訳（1 スレッド。画面の検索は最大 $workers スレッドで並行して読む）"
$share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
foreach ($pass in 1..2) {
    $watch.Restart()
    foreach ($f in $files) {
        try {
            $stream = [System.IO.FileStream]::new($f.Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
            $stream.Dispose()
        } catch {
        }
    }
    $watch.Stop()
    out ("開いて閉じるだけ（{0} 回目）: {1}（1 件 {2:N0} マイクロ秒）" -f $pass, (ms $watch), ($watch.Elapsed.TotalMilliseconds * 1000 / [Math]::Max(1, $files.Count)))
}
# 読んだ内容は 5・6 章で使う。インデックスが大きくてもメモリを使いすぎないよう、残すのは textMaxChars 文字まで
$textMaxChars = 256000000
$texts = New-Object System.Collections.Generic.List[string]
foreach ($pass in 1..2) {
    $texts.Clear()
    $kept = 0L
    $watch.Restart()
    foreach ($f in $files) {
        try {
            $stream = [System.IO.FileStream]::new($f.Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
            $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8, $true)
            $text = $reader.ReadToEnd()
            if ($kept + $text.Length -le $textMaxChars) {
                $texts.Add($text)
                $kept += $text.Length
            }
            $reader.Dispose()
        } catch {
        }
    }
    $watch.Stop()
    out ("開いて全文を読む（{0} 回目）: {1}（{2:N1} MB/秒）" -f $pass, (ms $watch), ($bytes / 1MB / [Math]::Max(0.001, $watch.Elapsed.TotalSeconds)))
}

# ---- 5. 照合の内訳 ----
out ""
out "## 5. 照合の内訳（読んだ内容を使い回し、ファイルを開かない）"
$warm = newTsvTextCache $textMaxChars
[void](searchTsvFiles $files 0 $files.Count ([regex]"x") 0 ([regex]"x") "filter" $warm)
foreach ($word in $Words) {
    $search = newSearchRegex $word $true $false
    $label = if ($word -eq $Words[0]) { "（存在しない語）" } else { "ワード $([array]::IndexOf($Words, $word))" }
    $watch.Restart()
    $hits = searchTsvFiles $files 0 $files.Count $search.Regex 10000 $search.TextRegex $search.ScanMode $warm
    $watch.Stop()
    out "1 スレッド $label（ヒット $($hits.Count)）: $(ms $watch)"
    $watch.Restart()
    $result = searchIndex $word $index.Files $true 10000 50 -cache $warm
    $watch.Stop()
    out "並列（画面と同じ searchIndex）$label（ヒット $($result.Hits.Count)）: $(ms $watch)"
}
$chars = 0L
foreach ($text in $texts) {
    $chars += $text.Length
}
$search = newSearchRegex $Words[0] $true $false
$watch.Restart()
foreach ($text in $texts) {
    [void]$search.TextRegex.IsMatch($text)
}
$watch.Stop()
out ("正規表現だけ（TSV ごとに IsMatch。{0:N0} 件・{1:N0} 万文字）: {2}" -f $texts.Count, ($chars / 10000), (ms $watch))
if ($texts.Count -lt $files.Count) {
    out "（インデックスが大きいため、5 章の一部と 6 章は先頭の $($texts.Count.ToString('N0')) 件だけで測った）"
}

# ---- 6. まとめた場合 ----
if (!$SkipPack) {
    out ""
    out "## 6. 同じ内容を 1 ファイルにまとめた場合（%TEMP% に書いて読み、消す）"
    $packDir = Join-Path ([System.IO.Path]::GetTempPath()) "tebunko_grep_measure_$PID"
    [void][System.IO.Directory]::CreateDirectory($packDir)
    $packFile = Join-Path $packDir "pack.txt"
    try {
        $watch.Restart()
        $writer = [System.IO.StreamWriter]::new($packFile, $false, [System.Text.UTF8Encoding]::new($false))
        foreach ($text in $texts) {
            $writer.Write($text)
            $writer.Write("`n")
        }
        $writer.Dispose()
        $watch.Stop()
        out "書き出し（$(([System.IO.FileInfo]::new($packFile).Length / 1MB).ToString('N1')) MB）: $(ms $watch)"
        foreach ($pass in 1..2) {
            $watch.Restart()
            $all = [System.IO.File]::ReadAllText($packFile)
            $watch.Stop()
            out "1 ファイルを読む（$pass 回目）: $(ms $watch)"
        }
        $watch.Restart()
        [void]$search.TextRegex.IsMatch($all)
        $watch.Stop()
        out "まとめた全文に 1 回だけ照合（存在しない語）: $(ms $watch)"
        $all = $null
    } finally {
        Remove-Item -LiteralPath $packDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

$texts.Clear()
[System.IO.File]::WriteAllLines($OutFile, $report, [System.Text.UTF8Encoding]::new($true))
Write-Host ""
Write-Host "結果を $OutFile に書きました。"
