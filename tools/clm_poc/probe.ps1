# 制限言語モード（CLM）で、制限モードに使う部品が動くかを確かめる（Issue #71 の作業 1 の試作）。
#
# このスクリプト自身も CLM で動く書き方だけで書く（Add-Type・class・コア以外の .NET の型を使わない）。
# 結果は 1 行 1 項目で標準出力に出し、GitHub Actions ではジョブの概要（GITHUB_STEP_SUMMARY）にも表で書く。
#
# 使い方:
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\clm_poc\probe.ps1 -Root <リポジトリ> -Label <見出し>
param(
    [Parameter(Mandatory = $true)][string]$Root,
    [string]$Label = 'probe'
)

$ErrorActionPreference = 'Stop'
$tar = Join-Path $env:windir 'System32\tar.exe'
$results = @()

# 1 項目を実行して結果を貯める（CLM では [pscustomobject] に変換できないのでハッシュテーブル）。$check が例外を投げたら失敗として、そのメッセージを残す。
function probe([string]$name, [scriptblock]$check) {
    $start = Get-Date
    try {
        $detail = & $check
        $ok = 'ok'
    } catch {
        $detail = $_.Exception.Message
        $ok = 'NG'
    }
    $ms = [int]((Get-Date) - $start).TotalMilliseconds
    $text = (@($detail) -join ' ') -replace '[\r\n|]+', ' '
    if ($text.Length -gt 200) { $text = $text.Substring(0, 200) + '...' }
    $script:results += @{ Name = $name; Result = $ok; Ms = $ms; Detail = $text }
}

$work = Join-Path $env:TEMP ('clm_probe_' + (Get-Date -Format 'yyyyMMddHHmmssfff'))
New-Item -ItemType Directory -Path $work | Out-Null
$book = Join-Path $Root 'tests\testdata\office\Excel\基本.xlsx'
$docx = Get-ChildItem -LiteralPath (Join-Path $Root 'tests\testdata\office\Word') -Recurse -Filter '*.docx' | Select-Object -First 1

# --- 言語モードと実行環境 ---
probe 'LanguageMode（-File で起動したこのスクリプト）' { $ExecutionContext.SessionState.LanguageMode }
probe 'PowerShell の版' { $PSVersionTable.PSVersion.ToString() }
probe 'スクリプトの場所' { $PSCommandPath }
probe 'Get-ExecutionPolicy -List' { (Get-ExecutionPolicy -List | ForEach-Object { '{0}={1}' -f $_.Scope, $_.ExecutionPolicy }) -join ', ' }

# --- tar.exe（ZIP の読み書き） ---
probe 'tar.exe で xlsx から XML を取り出す（ASCII のパス）' {
    $ascii = Join-Path $work 'book.xlsx'
    Copy-Item -LiteralPath $book -Destination $ascii
    $out = Join-Path $work 'x1'
    New-Item -ItemType Directory -Path $out | Out-Null
    & $tar -xf $ascii -C $out 'xl/workbook.xml' 'xl/worksheets' 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "tar の終了コード $LASTEXITCODE" }
    'sheets=' + @(Get-ChildItem -LiteralPath (Join-Path $out 'xl\worksheets') -Filter '*.xml').Count
}
probe 'tar.exe に CP932 に無い文字のパスを引数で渡す（失敗する見込み）' {
    $uni = Join-Path $work '𠮷野家😀.xlsx'
    Copy-Item -LiteralPath $book -Destination $uni
    $out = Join-Path $work 'x2'
    New-Item -ItemType Directory -Path $out | Out-Null
    $msg = & $tar -xf $uni -C $out 'xl/workbook.xml' 2>&1
    "exit=$LASTEXITCODE " + (@($msg) -join ' ')
}
probe 'tar.exe に cmd のリダイレクトで渡す（CP932 に無い文字のパス）' {
    $uni = Join-Path $work '𠮷野家😀.xlsx'
    $out = Join-Path $work 'x3'
    New-Item -ItemType Directory -Path $out | Out-Null
    $msg = cmd.exe /d /c "`"$tar`" -xf - -C `"$out`" xl/workbook.xml < `"$uni`"" 2>&1
    if ($LASTEXITCODE -ne 0) { throw "exit=$LASTEXITCODE " + (@($msg) -join ' ') }
    [xml]$xml = Get-Content -LiteralPath (Join-Path $out 'xl\workbook.xml') -Encoding UTF8 -Raw
    'sheets=' + @($xml.workbook.sheets.sheet).Count
}
probe 'tar.exe で ZIP を作る（cmd で作業フォルダを移ってから相対の名前で）' {
    $dir = Join-Path $work '結果𠮷'
    New-Item -ItemType Directory -Path (Join-Path $dir 'xl') | Out-Null
    Set-Content -LiteralPath (Join-Path $dir 'xl\a.xml') -Value '<a/>' -Encoding UTF8
    $msg = cmd.exe /d /c "cd /d `"$dir`" && `"$tar`" --format zip -cf out.xlsx xl" 2>&1
    if ($LASTEXITCODE -ne 0) { throw "exit=$LASTEXITCODE " + (@($msg) -join ' ') }
    'bytes=' + (Get-Item -LiteralPath (Join-Path $dir 'out.xlsx')).Length
}
if ($docx) {
    probe 'tar.exe で docx から word/document.xml を取り出す' {
        $out = Join-Path $work 'x4'
        New-Item -ItemType Directory -Path $out | Out-Null
        $msg = cmd.exe /d /c "`"$tar`" -xf - -C `"$out`" word/document.xml < `"$($docx.FullName)`"" 2>&1
        if ($LASTEXITCODE -ne 0) { throw "exit=$LASTEXITCODE " + (@($msg) -join ' ') }
        'chars=' + (Get-Content -LiteralPath (Join-Path $out 'word\document.xml') -Encoding UTF8 -Raw).Length
    }
}

# --- tar.exe が使えないときの代わり（Microsoft が署名したモジュール） ---
probe 'Expand-Archive' {
    $zip = Join-Path $work 'book.zip'
    Copy-Item -LiteralPath $book -Destination $zip
    Expand-Archive -LiteralPath $zip -DestinationPath (Join-Path $work 'x5')
    'sheets=' + @(Get-ChildItem -LiteralPath (Join-Path $work 'x5\xl\worksheets') -Filter '*.xml').Count
}
probe 'Compress-Archive' {
    $zip = Join-Path $work 'made.zip'
    Compress-Archive -Path (Join-Path $work 'x5\xl') -DestinationPath $zip
    'bytes=' + (Get-Item -LiteralPath $zip).Length
}

# --- 検索に使う部品 ---
probe 'New-Object regex（オプションとタイムアウト付き）' {
    $re = New-Object regex -ArgumentList @('B', 'IgnoreCase', [timespan]::FromSeconds(2))
    'match=' + $re.IsMatch('abc') + ' timeout=' + $re.MatchTimeout
}
probe 'Select-String' {
    $tsv = Join-Path $work 'a.tsv'
    Set-Content -LiteralPath $tsv -Value @("1`tりんご", "2`tみかん", "3`tりんご飴") -Encoding UTF8
    'hits=' + @(Select-String -LiteralPath $tsv -Pattern 'りんご' -Encoding UTF8).Count
}
probe '[regex]::Replace にスクリプトブロックを渡す（CLM では壊れる見込み）' {
    [regex]::Replace('abc', 'b', { param($m) 'X' })
}
probe 'ループの中の関数呼び出し 1 万回' {
    function noop($x) { $x }
    $t = Measure-Command { for ($i = 0; $i -lt 10000; $i++) { noop $i | Out-Null } }
    ('{0:0.000} ms/回' -f ($t.TotalMilliseconds / 10000))
}
probe 'Start-Job' {
    $job = Start-Job -ScriptBlock { 1 + 1 }
    $null = Wait-Job -Job $job -Timeout 120
    $r = Receive-Job -Job $job
    Remove-Job -Job $job -Force
    'result=' + $r
}

# --- COM・.NET（制限の確認） ---
probe 'COM: Scripting.FileSystemObject（許可されている見込み）' { (New-Object -ComObject Scripting.FileSystemObject).GetTempName() }
probe 'COM: Shell.Application（止まる見込み）' { $null = New-Object -ComObject Shell.Application; 'created' }
probe 'Add-Type（止まる見込み）' { Add-Type -AssemblyName PresentationFramework; 'loaded' }
probe '[IO.Path]::GetTempPath()（止まる見込み）' { [IO.Path]::GetTempPath() }
probe 'New-Object System.Text.UTF8Encoding（止まる見込み）' { $null = New-Object System.Text.UTF8Encoding $false; 'created' }

# --- 今のスクリプトを読み込めるか（今は止まる見込み。作業 3 で直す対象を知るため） ---
$files = @(
    'scripts\shared\core\paths.ps1',
    'scripts\shared\core\fs.ps1',
    'scripts\shared\core\text.ps1',
    'scripts\tebunko_grep\core\paths_grep.ps1',
    'scripts\tebunko_grep\core\settings_grep.ps1',
    'scripts\tebunko_grep\index\index_name.ps1',
    'scripts\tebunko_grep\search\search_query.ps1',
    'scripts\tebunko_grep\lib.ps1'
)
foreach ($f in $files) {
    $path = Join-Path $Root $f
    probe "dot-source $f" { & { $ErrorActionPreference = 'Stop'; . $path }; 'loaded' }
}

Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue

# --- 結果 ---
foreach ($r in $results) { "[{0}] {1} ({2} ms): {3}" -f $r.Result, $r.Name, $r.Ms, $r.Detail }
if ($env:GITHUB_STEP_SUMMARY) {
    $lines = @("### $Label", '', '| 項目 | 結果 | ms | 詳細 |', '|---|---|---|---|')
    foreach ($r in $results) { $lines += '| {0} | {1} | {2} | {3} |' -f $r.Name, $r.Result, $r.Ms, $r.Detail }
    $lines += ''
    Add-Content -LiteralPath $env:GITHUB_STEP_SUMMARY -Value $lines -Encoding UTF8
}
