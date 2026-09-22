# 本物の制限言語モード（CLM）で、制限モードに使う部品が思ったとおりに動くかを確かめる（CI の clm.yml から呼ぶ）。
#
# 項目ごとに「動く（ok）」か「止まる（NG）」かの見込みを書き、見込みと違った項目の数を終了コードにする。
# このスクリプト自身も CLM で動く書き方だけで書く（Add-Type・class・コア以外の .NET の型・[pscustomobject] を使わない）。
# 結果は 1 行 1 項目で標準出力に出し、GitHub Actions ではジョブの概要（GITHUB_STEP_SUMMARY）にも表で書く。
#
# 使い方（WDAC のポリシーを当てた後に、新しい powershell.exe で）:
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\clm\check.ps1 -Root <リポジトリ>
# 手元の模擬の CLM では確かめられない（Microsoft が署名したモジュールまで止まるなど、本物より厳しい）。
param(
    [Parameter(Mandatory = $true)][string]$Root,
    [string]$Label = 'clm'
)

$ErrorActionPreference = 'Stop'
$tar = Join-Path $env:windir 'System32\tar.exe'
$results = @()

# 1 項目を実行して、見込み（ok = 例外なく終わる、NG = 例外で止まる）と比べる。
# CLM では [pscustomobject] に変換できないので、結果はハッシュテーブルで貯める
function expect([string]$name, [string]$expected, [scriptblock]$check) {
    $start = Get-Date
    try {
        $detail = & $check
        $actual = 'ok'
    } catch {
        $detail = $_.Exception.Message
        $actual = 'NG'
    }
    $ms = [int]((Get-Date) - $start).TotalMilliseconds
    $text = (@($detail) -join ' ') -replace '[\r\n|]+', ' '
    if ($text.Length -gt 200) { $text = $text.Substring(0, 200) + '...' }
    $script:results += @{ Name = $name; Expected = $expected; Actual = $actual; Ms = $ms; Detail = $text }
}

$work = Join-Path $env:TEMP ('clm_check_' + (Get-Date -Format 'yyyyMMddHHmmssfff'))
New-Item -ItemType Directory -Path $work | Out-Null
$book = Join-Path $Root 'tests\testdata\office\Excel\基本.xlsx'
$docx = @(Get-ChildItem -LiteralPath (Join-Path $Root 'tests\testdata\office\Word') -Recurse -Filter '*.docx')[0]
# tar.exe は引数を ANSI のコードページ（日本語版の Windows は CP932、英語版は CP1252）で受け取るため、
# そこに無い文字を含むパスは引数では開けない。英語版のランナーでは日本語の名前も開けないため、引数で渡すのは ASCII のパスだけにする
$asciiBook = Join-Path $work 'book.xlsx'
Copy-Item -LiteralPath $book -Destination $asciiBook
$unicodeBook = Join-Path $work '𠮷野家😀.xlsx'
Copy-Item -LiteralPath $book -Destination $unicodeBook

# --- 前提: 本当に CLM になっているか ---
expect '言語モードが ConstrainedLanguage' ok {
    $mode = [string]$ExecutionContext.SessionState.LanguageMode
    if ($mode -ne 'ConstrainedLanguage') { throw "言語モードが $mode（ポリシーが効いていない）" }
    $mode
}

# --- ZIP の読み書き（tar.exe） ---
expect 'tar.exe で xlsx から XML を取り出す（ASCII のパスを引数で渡す）' ok {
    $out = Join-Path $work 'x1'
    New-Item -ItemType Directory -Path $out | Out-Null
    & $tar -xf $asciiBook -C $out 'xl/workbook.xml' 'xl/worksheets' 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "tar の終了コード $LASTEXITCODE" }
    'sheets=' + @(Get-ChildItem -LiteralPath (Join-Path $out 'xl\worksheets') -Filter '*.xml').Count
}
expect 'tar.exe に CP932 に無い文字（𠮷・絵文字）のパスを引数で渡すと開けない' NG {
    $out = Join-Path $work 'x2'
    New-Item -ItemType Directory -Path $out | Out-Null
    $msg = & $tar -xf $unicodeBook -C $out 'xl/workbook.xml' 2>&1
    if ($LASTEXITCODE -ne 0) { throw "exit=$LASTEXITCODE " + (@($msg) -join ' ') }
    'opened'
}
# 展開先も -C で渡さず、cmd で移ってから展開する（利用者名が日本語だと TEMP のパスも日本語になるため）
expect 'tar.exe に cmd のリダイレクトで渡し、日本語の名前のフォルダに展開する（CP932 に無い文字のパス）' ok {
    $out = Join-Path $work '展開先𠮷'
    New-Item -ItemType Directory -Path $out | Out-Null
    $msg = cmd.exe /d /c "cd /d `"$out`" && `"$tar`" -xf - xl/workbook.xml < `"$unicodeBook`"" 2>&1
    if ($LASTEXITCODE -ne 0) { throw "exit=$LASTEXITCODE " + (@($msg) -join ' ') }
    [xml]$xml = Get-Content -LiteralPath (Join-Path $out 'xl\workbook.xml') -Encoding UTF8 -Raw
    'sheets=' + @($xml.workbook.sheets.sheet).Count
}
expect 'tar.exe で ZIP を作る（cmd で作業フォルダに移り、相対の名前で）' ok {
    $dir = Join-Path $work '結果𠮷'
    New-Item -ItemType Directory -Path (Join-Path $dir 'xl') | Out-Null
    Set-Content -LiteralPath (Join-Path $dir 'xl\a.xml') -Value '<a/>' -Encoding UTF8
    $msg = cmd.exe /d /c "cd /d `"$dir`" && `"$tar`" --format zip -cf out.xlsx xl" 2>&1
    if ($LASTEXITCODE -ne 0) { throw "exit=$LASTEXITCODE " + (@($msg) -join ' ') }
    $bytes = (Get-Item -LiteralPath (Join-Path $dir 'out.xlsx')).Length
    if ($bytes -le 0) { throw '空のファイル' }
    "bytes=$bytes"
}
expect 'tar.exe で docx から word/document.xml を取り出す' ok {
    $out = Join-Path $work 'x4'
    New-Item -ItemType Directory -Path $out | Out-Null
    $msg = cmd.exe /d /c "cd /d `"$out`" && `"$tar`" -xf - word/document.xml < `"$($docx.FullName)`"" 2>&1
    if ($LASTEXITCODE -ne 0) { throw "exit=$LASTEXITCODE " + (@($msg) -join ' ') }
    'chars=' + (Get-Content -LiteralPath (Join-Path $out 'word\document.xml') -Encoding UTF8 -Raw).Length
}

# --- tar.exe が使えないときの代わり（Microsoft が署名したモジュールは信頼されて動く） ---
expect 'Expand-Archive' ok {
    $zip = Join-Path $work 'book.zip'
    Copy-Item -LiteralPath $book -Destination $zip
    Expand-Archive -LiteralPath $zip -DestinationPath (Join-Path $work 'x5')
    'sheets=' + @(Get-ChildItem -LiteralPath (Join-Path $work 'x5\xl\worksheets') -Filter '*.xml').Count
}
expect 'Compress-Archive' ok {
    $zip = Join-Path $work 'made.zip'
    Compress-Archive -Path (Join-Path $work 'x5\xl') -DestinationPath $zip
    'bytes=' + (Get-Item -LiteralPath $zip).Length
}

# --- 検索に使う部品 ---
expect 'New-Object regex（オプションとタイムアウト付き）' ok {
    $re = New-Object regex -ArgumentList @('B', 'IgnoreCase', [timespan]::FromSeconds(2))
    if (!$re.IsMatch('abc')) { throw '一致しない' }
    'timeout=' + $re.MatchTimeout
}
expect 'Select-String' ok {
    $tsv = Join-Path $work 'a.tsv'
    Set-Content -LiteralPath $tsv -Value @("1`tりんご", "2`tみかん", "3`tりんご飴") -Encoding UTF8
    $hits = @(Select-String -LiteralPath $tsv -Pattern 'りんご' -Encoding UTF8).Count
    if ($hits -ne 2) { throw "hits=$hits" }
    "hits=$hits"
}
expect '[regex]::Replace にスクリプトブロックを渡すと、ブロックの文字列に置き換わる' ok {
    $replaced = [regex]::Replace('abc', 'b', { param($m) 'X' })
    if ($replaced -eq 'aXc') { throw 'スクリプトブロックが呼ばれた（CLM でない？）' }
    $replaced
}
expect 'Start-Job' ok {
    $job = Start-Job -ScriptBlock { 1 + 1 }
    $null = Wait-Job -Job $job -Timeout 120
    $r = Receive-Job -Job $job
    Remove-Job -Job $job -Force
    if ($r -ne 2) { throw "result=$r" }
    "result=$r"
}
expect 'ループの中の関数呼び出し 1 万回（参考）' ok {
    function noop($x) { $x }
    $t = Measure-Command { for ($i = 0; $i -lt 10000; $i++) { noop $i | Out-Null } }
    ('{0:0.000} ms/回' -f ($t.TotalMilliseconds / 10000))
}

# --- CLM で止まるもの ---
expect 'COM: Scripting.FileSystemObject（許可リストにある）' ok { (New-Object -ComObject Scripting.FileSystemObject).GetTempName() }
expect 'COM: Shell.Application' NG { $null = New-Object -ComObject Shell.Application; 'created' }
expect 'Add-Type' NG { Add-Type -AssemblyName PresentationFramework; 'loaded' }
expect '[IO.Path]::GetTempPath()' NG { [IO.Path]::GetTempPath() }
expect 'New-Object System.Text.UTF8Encoding' NG { $null = New-Object System.Text.UTF8Encoding $false; 'created' }
expect '[pscustomobject] への変換' NG { $null = [pscustomobject]@{ a = 1 }; 'created' }

# --- tebunko_grep の制限モード ---
expect 'start.ps1 が制限モードで起動する旨を出す（終了コード 11）' ok {
    $start = Join-Path $Root 'scripts\tebunko_grep\start.ps1'
    $output = & powershell -NoProfile -ExecutionPolicy Bypass -File $start -CheckOnly 2>&1
    $code = $LASTEXITCODE
    if ($code -ne 11) { throw "exit=$code " + (@($output) -join ' ') }
    "exit=$code"
}

expect '制限モードの検索と、結果のブック（tar.exe で xlsx を作る）' ok {
    . (Join-Path $Root 'scripts\tebunko_grep\restricted\lib_restricted.ps1')
    $index = Join-Path $work 'index'
    $tsv = Join-Path $index '見積\a.xlsx\売上.tsv'
    New-Item -ItemType Directory -Path (Split-Path $tsv -Parent) -Force | Out-Null
    Set-Content -LiteralPath $tsv -Value "りんご`t100`r`nみかん`r`nりんご飴`r`n" -Encoding UTF8 -NoNewline
    $files = getRestrictedTsvFiles @(@{ Name = '見積'; Path = (Join-Path $index '見積') })
    $result = searchRestricted 'りんご' $files $true
    if ($result.Hits.Count -ne 2) { throw "hits=$($result.Hits.Count)" }
    $regex = (newSearchRegex 'りんご' $true $false).Regex
    $book = saveResultBook $result $regex @(, @('検索ワード', 'りんご')) (Join-Path $work '検索結果')
    # ブックの名前は日本語のため、引数ではなく標準入力から渡す
    $list = cmd.exe /d /c "`"$tar`" -tf - < `"$book`"" 2>&1
    if (@($list) -notcontains 'xl/worksheets/sheet1.xml') { throw 'ブックの中に sheet1.xml が無い: ' + (@($list) -join ' ') }
    'hits=2 ' + (Get-Item -LiteralPath $book).Length + ' bytes'
}

Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue

# --- 結果 ---
$failed = 0
foreach ($r in $results) {
    $mark = if ($r.Actual -eq $r.Expected) { 'pass' } else { $failed++; 'FAIL' }
    '[{0}] {1}（見込み {2}・結果 {3}、{4} ms）: {5}' -f $mark, $r.Name, $r.Expected, $r.Actual, $r.Ms, $r.Detail
}
if ($env:GITHUB_STEP_SUMMARY) {
    $lines = @("### $Label", '', '| 項目 | 見込み | 結果 | ms | 詳細 |', '|---|---|---|---|---|')
    foreach ($r in $results) { $lines += '| {0} | {1} | {2} | {3} | {4} |' -f $r.Name, $r.Expected, $r.Actual, $r.Ms, $r.Detail }
    $lines += ''
    Add-Content -LiteralPath $env:GITHUB_STEP_SUMMARY -Value $lines -Encoding UTF8
}
"見込みと違った項目: $failed"
exit $failed
