# 場所ごとに 1 件ずつ散らばる語のデータ（tools\perf\new_scattered_data.ps1）、前と後の ref を交互に測る道具（tools\perf\measure_compare.ps1）、
# PSCustomObject 1 万件を作る時間の計測（tools\perf\measure_hit_objects.ps1）と、それらを流す perf.yml の入力と段を確かめる。
# スクリプトは同じプロセスで動かす（Windows PowerShell 5.1 は -File で起動すると、param の既定値の $PSScriptRoot が空になるため）
BeforeAll {
    $script:repo = (Resolve-Path "$PSScriptRoot\..\..").ProviderPath
    $script:newData = Join-Path $script:repo "tools\perf\new_scattered_data.ps1"
    $script:compare = Join-Path $script:repo "tools\perf\measure_compare.ps1"
    $script:hitObjects = Join-Path $script:repo "tools\perf\measure_hit_objects.ps1"
    $script:utf8 = New-Object System.Text.UTF8Encoding($false)

    function script:readJson([string]$path) { [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8) | ConvertFrom-Json }

    function script:newScattered([string]$dest, [string]$words, [int]$files, [int]$places) {
        & $script:newData -Tool $script:repo -Dest $dest -Words $words -Files $files -Places $places | Out-Null
    }

    # ファイルの一覧（相対パス）と中身を 1 つの文字列にする
    function script:dumpTree([string]$root) {
        $lines = foreach ($f in @([System.IO.Directory]::GetFiles($root, "*", "AllDirectories") | Sort-Object)) {
            $f.Substring($root.Length) + "`n" + [System.IO.File]::ReadAllText($f, [System.Text.Encoding]::UTF8)
        }
        return ($lines -join "`n---`n")
    }
}

Describe "new_scattered_data.ps1" -Tag Io {
    BeforeAll {
        # 検索の It が使う newContentIndexFiles・searchContentIndex（利用者の既定のワークスペースは使わない）
        . "$PSScriptRoot\..\helpers\load.ps1"
        $script:dest = Join-Path $TestDrive "data\tsv"
        $script:words = Join-Path $TestDrive "data\words.tsv"
        newScattered $script:dest $script:words 2 3
    }

    It "new_scattered_data はファイル × 場所ぶんの TSV を作り、どの TSV も語を含む行がちょうど 1 つ" {
        $tsv = @([System.IO.Directory]::GetFiles($script:dest, "*.tsv", "AllDirectories"))
        $tsv.Count | Should -Be 6
        $word = ([System.IO.File]::ReadAllLines($script:words, [System.Text.Encoding]::UTF8)[1].Split("`t"))[1]
        foreach ($f in $tsv) {
            $lines = @([System.IO.File]::ReadAllLines($f, [System.Text.Encoding]::UTF8))
            $lines.Count | Should -Be 3
            @($lines | Where-Object { $_.Contains($word) }).Count | Should -Be 1
        }
        @($tsv | ForEach-Object { [System.IO.Path]::GetDirectoryName($_) } | Select-Object -Unique).Count | Should -Be 2
    }

    It "new_scattered_data は語の表を words.tsv と同じ列で書く" {
        $lines = @([System.IO.File]::ReadAllLines($script:words, [System.Text.Encoding]::UTF8))
        $lines.Count | Should -Be 2
        $lines[0] | Should -BeExactly "名前`t語`t正規表現`t件数"
        $f = $lines[1].Split("`t")
        $f.Count | Should -Be 4
        $f[0] | Should -Not -BeNullOrEmpty
        $f[1] | Should -Not -BeNullOrEmpty
        $f[2] | Should -BeExactly "false"
        $f[3] | Should -BeExactly "6"
    }

    It "new_scattered_data は同じ引数から同じ中身を作る" {
        $again = Join-Path $TestDrive "again\tsv"
        $againWords = Join-Path $TestDrive "again\words.tsv"
        newScattered $again $againWords 2 3
        dumpTree $again | Should -BeExactly (dumpTree $script:dest)
        [System.IO.File]::ReadAllText($againWords) | Should -BeExactly ([System.IO.File]::ReadAllText($script:words))
    }

    It "new_scattered_data の既定は 9,500 件で、検索の上限 1 万件を超えない" {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:newData, [ref]$null, [ref]$null)
        $defaults = @{}
        foreach ($p in $ast.ParamBlock.Parameters) { if ($p.DefaultValue) { $defaults[$p.Name.VariablePath.UserPath] = [int]$p.DefaultValue.SafeGetValue() } }
        $defaults.Files | Should -Be 19
        $defaults.Places | Should -Be 500
        $defaults.Lines | Should -Be 1
        $defaults.Files * $defaults.Places * $defaults.Lines | Should -Be 9500
        $defaults.Files * $defaults.Places * $defaults.Lines | Should -BeLessThan 10000
    }

    It "作ったデータを検索すると、場所ごとに 1 件ずつのヒットになる" {
        $root = Join-Path $TestDrive "search\tsv"
        $words = Join-Path $TestDrive "search\words.tsv"
        newScattered $root $words 2 3
        $word = ([System.IO.File]::ReadAllLines($words, [System.Text.Encoding]::UTF8)[1].Split("`t"))[1]
        $files = newContentIndexFiles $root (Join-Path $TestDrive "search\content")
        $hits = @((searchContentIndex $word $files $true).Hits)
        $hits.Count | Should -Be 6
        @($hits | ForEach-Object { "$($_.RelDir)|$($_.Book)|$($_.Location)" } | Select-Object -Unique).Count | Should -Be 6
    }
}

Describe "measure_compare.ps1" -Tag Io {
    BeforeAll {
        # measure_perf.ps1 の代わりの偽の道具。受けた引数を記録し、-Out の <組>\<base|target>\<データ> から決まった TotalMs を result.json に書く。
        # TSV を見つけたら pack に変えたことにして消す（measure_perf.ps1 と同じく、2 組目からは TSV が無い）
        $script:fake = Join-Path $TestDrive "fake_measure_perf.ps1"
        $fakeText = @'
param ([string]$Index, [string]$Work, [string]$Tool, [string]$Words, [int]$Count, [string]$Out, [string]$Label, [string]$RunId, [string]$Ref, [string]$Sha, [string]$Scale, [int]$SearchLimit)
$ErrorActionPreference = "Stop"
$tsv = @([System.IO.Directory]::GetFiles($Index, "*.tsv", "AllDirectories"))
foreach ($f in $tsv) { [System.IO.File]::Delete($f) }
$log = [ordered]@{ Index = $Index; Work = $Work; Tool = $Tool; Words = $Words; Out = $Out; Count = $Count; HadTsv = ($tsv.Count -gt 0) }
[System.IO.File]::AppendAllText($env:PERF_FAKE_LOG, (($log | ConvertTo-Json -Compress) + "`n"), (New-Object System.Text.UTF8Encoding($false)))
$parts = $Out.TrimEnd("\").Split("\")
$data = $parts[-1]; $side = $parts[-2]; $round = [int]$parts[-3]
$table = [System.IO.File]::ReadAllText($env:PERF_FAKE_TABLE, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
$ms = [double]$table."$side|$round"
$search = foreach ($line in @([System.IO.File]::ReadAllLines($Words, [System.Text.Encoding]::UTF8) | Select-Object -Skip 1)) {
    $f = $line.Split("`t")
    [ordered]@{ Name = $f[0]; Hits = 6; TotalMs = [ordered]@{ Min = $ms; Median = $ms; Mean = $ms; Max = $ms; Count = $Count }; Gc0 = 3; Gc1 = 2; Gc2 = 1 }
}
[void][System.IO.Directory]::CreateDirectory($Out)
[System.IO.File]::WriteAllText((Join-Path $Out "result.json"), ([ordered]@{ Schema = 6; Search = @($search) } | ConvertTo-Json -Depth 6), (New-Object System.Text.UTF8Encoding($false)))
'@
        [System.IO.File]::WriteAllText($script:fake, $fakeText, (New-Object System.Text.UTF8Encoding($true)))
        $script:table = Join-Path $TestDrive "fake_table.json"
        $tableValues = [ordered]@{ "base|1" = 100; "base|2" = 110; "base|3" = 90; "target|1" = 120; "target|2" = 130; "target|3" = 110 }
        [System.IO.File]::WriteAllText($script:table, ($tableValues | ConvertTo-Json), $script:utf8)
        $env:PERF_FAKE_TABLE = $script:table

        # 元のデータ（perfdata の形の TSV と語の表、散らばる語の TSV と語の表）
        $script:src = Join-Path $TestDrive "src"
        foreach ($name in @("p1.xlsx", "p2.xlsx")) {
            $dir = Join-Path $script:src "perfdata\部署0\$name"
            [void][System.IO.Directory]::CreateDirectory($dir)
            [System.IO.File]::WriteAllText((Join-Path $dir "シート1.tsv"), "1`tの`r`n", $script:utf8)
        }
        $script:perfWords = Join-Path $script:src "perf_words.tsv"
        [System.IO.File]::WriteAllText($script:perfWords, "名前`t語`t正規表現`t件数`r`n大量`tの`tfalse`t10000+`r`n", $script:utf8)
        $script:scatWords = Join-Path $script:src "scattered_words.tsv"
        newScattered (Join-Path $script:src "scattered") $script:scatWords 2 3

        function script:runCompare([string]$name, [int]$rounds, [bool]$scattered) {
            $log = Join-Path $TestDrive "$name.log"
            $env:PERF_FAKE_LOG = $log
            $out = Join-Path $TestDrive "$name\result"
            $scat = @{}
            if ($scattered) { $scat = @{ ScatteredIndex = (Join-Path $script:src "scattered"); ScatteredWords = $script:scatWords } }
            & $script:compare -MeasureScript $script:fake -BaseTool $script:repo -TargetTool $script:repo -BaseRef "前" -TargetRef "後" `
                -PerfdataIndex (Join-Path $script:src "perfdata") -PerfdataWords $script:perfWords @scat -Rounds $rounds -Count 2 `
                -Work (Join-Path $TestDrive "$name\work") -Out $out | Out-Null
            return @{
                Calls = @([System.IO.File]::ReadAllLines($log, [System.Text.Encoding]::UTF8) | ForEach-Object { $_ | ConvertFrom-Json })
                Json = readJson (Join-Path $out "compare.json")
                Md = [System.IO.File]::ReadAllText((Join-Path $out "compare.md"), [System.Text.Encoding]::UTF8)
            }
        }
        function script:sideOf($call) { $p = $call.Out.TrimEnd("\").Split("\"); return "$($p[-2]) $($p[-1])" }

        $script:plain = runCompare "plain" 3 $false
        $script:withScattered = runCompare "scattered" 2 $true
    }

    It "measure_compare は前と後を交互に、組の数だけ測り、組ごとの中央値の中央値と揺れと比を出す" {
        @($script:plain.Calls | ForEach-Object { sideOf $_ }) -join "," | Should -BeExactly (@(1..3 | ForEach-Object { "base perfdata", "target perfdata" }) -join ",")
        $w = @($script:plain.Json.Words)
        $w.Count | Should -Be 1
        $w[0].Data | Should -BeExactly "perfdata"
        $w[0].Name | Should -BeExactly "大量"
        $w[0].BaseMs | Should -Be 100
        $w[0].TargetMs | Should -Be 120
        $w[0].Ratio | Should -Be 1.2
        $w[0].BaseSpread | Should -Be 0.2
        @($w[0].BaseRounds) -join "," | Should -BeExactly "100,110,90"
        @($w[0].TargetRounds) -join "," | Should -BeExactly "120,130,110"
        $w[0].HitsMatch | Should -BeTrue
        $script:plain.Json.Rounds | Should -Be 3
        $script:plain.Md | Should -Match "大量"
        $script:plain.Md | Should -Match "前"
    }

    It "measure_compare は 1 組目だけインデックス作成を測り、ref ごとに別の置き場を使う" {
        $calls = $script:plain.Calls
        $base = @($calls | Where-Object { (sideOf $_) -eq "base perfdata" })
        $target = @($calls | Where-Object { (sideOf $_) -eq "target perfdata" })
        @($base | ForEach-Object { $_.HadTsv }) -join "," | Should -BeExactly "True,False,False"
        @($target | ForEach-Object { $_.HadTsv }) -join "," | Should -BeExactly "True,False,False"
        @($base.Index | Select-Object -Unique).Count | Should -Be 1
        @($base.Work | Select-Object -Unique).Count | Should -Be 1
        @($target.Index | Select-Object -Unique).Count | Should -Be 1
        $base[0].Index | Should -Not -Be $target[0].Index
        $base[0].Work | Should -Not -Be $target[0].Work
        # 元のデータは消さない（写しを測る）
        @([System.IO.Directory]::GetFiles((Join-Path $script:src "perfdata"), "*.tsv", "AllDirectories")).Count | Should -Be 2
    }

    It "measure_compare は scattered のとき、1 組で perfdata と散らばる語を置き場と語の表を分けて 4 回起動し、起動ごとに別の -Out に書く" {
        $calls = $script:withScattered.Calls
        $calls.Count | Should -Be 8
        $order = @("base perfdata", "base scattered", "target perfdata", "target scattered")
        @($calls | ForEach-Object { sideOf $_ }) -join "," | Should -BeExactly (($order + $order) -join ",")
        $scatIndex = @($calls | Where-Object { (sideOf $_) -like "* scattered" } | ForEach-Object { $_.Index })
        foreach ($c in @($calls | Where-Object { (sideOf $_) -like "* perfdata" })) {
            $c.Words | Should -Be $script:perfWords
            $scatIndex | Should -Not -Contain $c.Index
        }
        $perfIndex = @($calls | Where-Object { (sideOf $_) -like "* perfdata" } | ForEach-Object { $_.Index })
        foreach ($c in @($calls | Where-Object { (sideOf $_) -like "* scattered" })) {
            $c.Words | Should -Be $script:scatWords
            $perfIndex | Should -Not -Contain $c.Index
        }
        @($calls.Out | Select-Object -Unique).Count | Should -Be 8
        $data = @($script:withScattered.Json.Words | ForEach-Object { $_.Data })
        $data | Should -Contain "perfdata"
        $data | Should -Contain "scattered"
    }

    It "measure_compare は前の ref が無ければ、後だけを測り、比を出さない" {
        $env:PERF_FAKE_LOG = Join-Path $TestDrive "only.log"
        $out = Join-Path $TestDrive "only\result"
        & $script:compare -MeasureScript $script:fake -TargetTool $script:repo -TargetRef "後" `
            -ScatteredIndex (Join-Path $script:src "scattered") -ScatteredWords $script:scatWords -Rounds 1 -Count 2 `
            -Work (Join-Path $TestDrive "only\work") -Out $out | Out-Null
        $calls = @([System.IO.File]::ReadAllLines($env:PERF_FAKE_LOG, [System.Text.Encoding]::UTF8) | ForEach-Object { $_ | ConvertFrom-Json })
        @($calls | ForEach-Object { sideOf $_ }) -join "," | Should -BeExactly "target scattered"
        $w = @((readJson (Join-Path $out "compare.json")).Words)
        $w[0].TargetMs | Should -Be 120
        $w[0].BaseMs | Should -BeNullOrEmpty
        $w[0].Ratio | Should -BeNullOrEmpty
    }

    AfterAll {
        Remove-Item Env:\PERF_FAKE_LOG, Env:\PERF_FAKE_TABLE -ErrorAction SilentlyContinue
    }
}

Describe "measure_hit_objects.ps1" -Tag Io {
    It "measure_hit_objects は 8 項目の PSCustomObject を件数ぶん作る時間を回数ぶん測る" {
        $out = Join-Path $TestDrive "hit"
        & $script:hitObjects -Hits 100 -Repeat 3 -Out $out | Out-Null
        $json = readJson (Join-Path $out "hit_objects.json")
        $json.Hits | Should -Be 100
        $json.Repeat | Should -Be 3
        @($json.Properties) -join "," | Should -BeExactly "Root,RelPath,RelDir,FileName,Book,Location,LineNumber,Line"
        $json.Created | Should -Be 100
        foreach ($k in @("Min", "Median", "Mean", "Max")) { $json.Ms.$k | Should -Not -BeNullOrEmpty }
        $json.Ms.Count | Should -Be 3
        $md = [System.IO.File]::ReadAllText((Join-Path $out "hit_objects.md"), [System.Text.Encoding]::UTF8)
        $md | Should -Match "(?m)^\| 100 \| 3 \|"
    }
}

Describe "perf.yml の比べと散らばる語の入力" -Tag Unit {
    BeforeAll {
        $script:yml = [System.IO.File]::ReadAllText((Join-Path $script:repo ".github\workflows\perf.yml"), [System.Text.Encoding]::UTF8) -replace "`r`n", "`n"
        # 入力 1 つの定義（次の入力か permissions の手前まで）
        function script:inputBlock([string]$name) {
            $m = [regex]::Match($script:yml, "(?ms)^      ${name}:\n(.*?)(?=^      [a-z_]+:\n|^permissions:)")
            $m.Success | Should -BeTrue -Because "入力 $name がある"
            return $m.Groups[1].Value
        }
        # 段 1 つ（次の段の手前まで）
        function script:steps() {
            $jobs = $script:yml.Substring($script:yml.IndexOf("`njobs:"))
            return @([regex]::Split($jobs, "(?m)^(?=      - name: )") | Where-Object { $_ -match "^      - name: " })
        }
        function script:stepNamed([string]$name) {
            $found = @(steps | Where-Object { $_ -match "^      - name: $([regex]::Escape($name))\n" })
            $found.Count | Should -Be 1 -Because "段 $name がある"
            return $found[0]
        }
        $script:addedSteps = @("前の ref の計測対象のコードをチェックアウトする", "散らばる語のデータを生成する", "前と後を比べて測る", "PSCustomObject 1 万件を作る時間を測る")
    }

    It "perf.yml は base_ref が空なら今の計測だけを流し、足した段は base_ref か scattered があるときだけ流す" {
        $base = inputBlock "base_ref"
        $base | Should -Match "(?m)^        default: `"`"$"
        $base | Should -Match "(?m)^        required: false$"
        $rounds = inputBlock "rounds"
        $rounds | Should -Match '(?m)^        options: \["3", "5"\]$'
        $rounds | Should -Match '(?m)^        default: "5"$'
        $scattered = inputBlock "scattered"
        $scattered | Should -Match '(?m)^        options: \["0", "1"\]$'
        $scattered | Should -Match '(?m)^        default: "0"$'

        (stepNamed "計測する") | Should -Match "(?m)^        if: \$\{\{ inputs\.base_ref == '' \}\}$"
        foreach ($name in $script:addedSteps) {
            $step = stepNamed $name
            $if = [regex]::Match($step, "(?m)^        if: (.*)$")
            $if.Success | Should -BeTrue -Because "$name に if がある"
            $if.Groups[1].Value | Should -Match "inputs\.(base_ref|scattered)"
            $if.Groups[1].Value | Should -Not -Match "[^\x00-\x7F]"
            $run = [regex]::Match($step, "(?ms)^        run: \|\n(.*)$")
            if ($run.Success) { $run.Groups[1].Value | Should -Not -Match "[^\x00-\x7F]" -Because "$name の run は ASCII だけ" }
        }
        $script:yml | Should -Match '(?m)^  group: perf-\$\{\{ inputs\.ref \}\}-\$\{\{ inputs\.base_ref \}\}$'
        # 結果の Summary には、比べと PSCustomObject の節があれば足す
        (stepNamed "結果を Summary に書き出す") | Should -Match "compare\.md"
        (stepNamed "結果を Summary に書き出す") | Should -Match "hit_objects\.md"
    }

    It "perf.yml は base_ref があるとき、取り込みのデータを作る段を飛ばす" {
        $if = [regex]::Match((stepNamed "取り込みのデータを生成する"), "(?m)^        if: (.*)$")
        $if.Success | Should -BeTrue
        $if.Groups[1].Value | Should -Match "inputs\.ingest != '0'"
        $if.Groups[1].Value | Should -Match "&& inputs\.base_ref == ''"
    }
}
