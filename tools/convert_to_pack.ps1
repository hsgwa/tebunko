# 今の形式のインデックス（場所ごとの TSV）を、検索用のまとめファイル（フォルダごと・拡張子ごとの content.xlsx.001.tsv など）に変換する（PoC 用）。
# 元のインデックスは読むだけで書き換えない。変換先に同じフォルダのまとめファイルがあれば飛ばすため、止めても続きから変換できる。
#
#   .\tools\convert_to_pack.ps1 -Source <元の index フォルダ> -Dest <変換先の index フォルダ>
#   -Minutes 9   … この時間を過ぎたら新しいフォルダを始めない（1 回の実行を短く切る）
#   -Workers 4   … 並行して変換するスレッドの数
param (
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$Dest,
    [double]$Minutes = 9,
    [int]$Workers = 4
)

$ErrorActionPreference = "Stop"
$lib = Join-Path (Split-Path $PSScriptRoot -Parent) "scripts\tebunko_grep\lib.ps1"
. $lib
$Source = (Resolve-Path -LiteralPath $Source).ProviderPath.TrimEnd("\")
[void][System.IO.Directory]::CreateDirectory($Dest)
$Dest = (Resolve-Path -LiteralPath $Dest).ProviderPath.TrimEnd("\")

# 直下に元のファイルのフォルダ（<ファイル名.xlsx>）があるフォルダを集める
$watch = [System.Diagnostics.Stopwatch]::StartNew()
$folders = New-Object System.Collections.Generic.List[string]
$bookDir = [regex]::new(${indexBookDirPattern}, "IgnoreCase")
foreach ($dir in @($Source) + @([System.IO.Directory]::GetDirectories((toLongPath $Source), "*", "AllDirectories") | ForEach-Object { fromLongPath $_ })) {
    if ($bookDir.IsMatch([System.IO.Path]::GetFileName($dir))) { continue }
    $hasBook = $false
    foreach ($sub in [System.IO.Directory]::EnumerateDirectories((toLongPath $dir))) {
        if ($bookDir.IsMatch([System.IO.Path]::GetFileName($sub))) { $hasBook = $true; break }
    }
    if ($hasBook) { $folders.Add($dir) }
}
$todo = @($folders | Where-Object {
    $rel = if ($_.Length -gt $Source.Length) { $_.Substring($Source.Length + 1) } else { "" }
    $target = if ($rel) { "$Dest\$rel" } else { $Dest }
    $longTarget = toLongPath $target
    !([System.IO.Directory]::Exists($longTarget) -and [System.IO.Directory]::GetFiles($longTarget, ${packFilePattern}).Count -gt 0)
})
Write-Host ("フォルダ {0:N0} 件（うち変換済み {1:N0} 件）。列挙 {2:N1} 秒" -f $folders.Count, ($folders.Count - $todo.Count), $watch.Elapsed.TotalSeconds)

$deadline = [DateTime]::Now.AddMinutes($Minutes)
$pool = [runspacefactory]::CreateRunspacePool(1, $Workers)
$pool.Open()
$jobs = foreach ($k in 0..($Workers - 1)) {
    $ps = [powershell]::Create()
    $ps.RunspacePool = $pool
    [void]$ps.AddScript({
        param ($lib, $list, $k, $step, $source, $dest, $deadline)
        . $lib
        $done = 0; $tsv = 0; $chars = 0L
        for ($i = $k; $i -lt $list.Count; $i += $step) {
            if ([DateTime]::Now -gt $deadline) { break }
            $folder = $list[$i]
            $rel = if ($folder.Length -gt $source.Length) { $folder.Substring($source.Length + 1) } else { "" }
            $target = if ($rel) { "$dest\$rel" } else { $dest }
            $r = convertIndexFolderToPack $folder $target
            $done++; $tsv += $r.Tsv; $chars += $r.Chars
        }
        @{ Done = $done; Tsv = $tsv; Chars = $chars }
    }).AddArgument($lib).AddArgument($todo).AddArgument($k).AddArgument($Workers).AddArgument($Source).AddArgument($Dest).AddArgument($deadline)
    @{ P = $ps; H = $ps.BeginInvoke() }
}
$done = 0; $tsv = 0; $chars = 0L
foreach ($j in $jobs) {
    $r = $j.P.EndInvoke($j.H)[0]
    $done += $r.Done; $tsv += $r.Tsv; $chars += $r.Chars
    $j.P.Dispose()
}
$pool.Dispose()
Write-Host ("変換: フォルダ {0:N0} 件・TSV {1:N0} 件・{2:N1} MB（UTF-16） / {3:N1} 秒。残り {4:N0} 件" -f $done, $tsv, ($chars * 2 / 1MB), $watch.Elapsed.TotalSeconds, ($todo.Count - $done))
