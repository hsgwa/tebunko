# 性能の計測（tools\measure_perf.ps1）のテスト
$measure = "$PSScriptRoot\..\..\tools\measure_perf.ps1"

function writeTsv([string]$path, [string[]]$lines) {
    [void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path))
    [System.IO.File]::WriteAllText($path, (($lines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
}

Describe "measure_perf.ps1" -Tag Io {
    # 取り込みの一時置き場の形（<フォルダ>\<ファイル名.xlsx>\<場所>.tsv）の小さなインデックスと、検索する語の表
    $index = Join-Path $TestDrive "work\index"
    writeTsv "$index\営業\見積.xlsx\シート1.tsv" @("1`t見積書`t山田", "2`tりんご`t佐藤")
    writeTsv "$index\営業\見積.xlsx\シート2.tsv" @("3`tみかん")
    writeTsv "$index\総務\名簿.xlsx\シート1.tsv" @("10`t佐藤`t総務", "11`t鈴木`t総務")
    $words = Join-Path $TestDrive "words.tsv"
    writeTsv $words @("名前`t語`t正規表現`t件数", "無い`t該当無し`tfalse`t0", "まれ`t見積書`tfalse`t1", "数字`t^\d{2}\t`ttrue`t2")
    $out = Join-Path $TestDrive "work\result"
    & $measure -Index $index -Work (Join-Path $TestDrive "work") -Words $words -AddWords "佐藤" -Repeat 1 -SampleMs 50 -Label "テスト" 6>$null
    $result = [System.IO.File]::ReadAllText("$out\result.json") | ConvertFrom-Json

    It "pack を作り、フォルダ・ブック・TSV の数を数える" {
        $result.Pack.Folders | Should Be 2
        $result.Pack.Books | Should Be 2
        $result.Pack.Tsv | Should Be 3
        $result.Pack.Packs | Should Be 2
        # pack にまとめたので、ブックごとの TSV は残らない
        @([System.IO.Directory]::GetFiles($index, "シート*.tsv", "AllDirectories")).Count | Should Be 0
    }

    It "語ごとに初回と 2 回目を検索し、ヒットの件数を返す" {
        $hits = @{}
        foreach ($s in $result.Search) { $hits["$($s.Name) $($s.Round)"] = $s.Hits }
        $hits["無い 初回"] | Should Be "0"
        $hits["まれ 初回"] | Should Be "1"
        $hits["まれ 2 回目"] | Should Be "1"
        $hits["数字 初回"] | Should Be "2"
        $hits["追加 1 初回"] | Should Be "2"
        @($result.SearchRuns).Count | Should Be 8
    }

    It "段階ごとのリソースと、推移の記録を書く" {
        $names = @($result.Resources | ForEach-Object { $_.Phase })
        $names -contains "pack の作成" | Should Be $true
        $names -contains "検索 1 初回" | Should Be $true
        $result.PeakWorkingSetMB | Should BeGreaterThan 0
        @([System.IO.File]::ReadAllLines("$out\resource.csv")).Count | Should BeGreaterThan 1
    }

    It "summary.md に表とグラフを書き、パスを書かない" {
        $md = [System.IO.File]::ReadAllText("$out\summary.md")
        $md | Should Match "# tebunko の性能（テスト）"
        $md | Should Match "xychart-beta"
        $md.Contains($TestDrive) | Should Be $false
    }

    It "pack だけのインデックスは、作成を測らずに検索だけを測る" {
        $again = Join-Path $TestDrive "again"
        & $measure -Index $index -Work $again -Words $words -Repeat 1 -SampleMs 50 6>$null
        $second = [System.IO.File]::ReadAllText("$again\result\result.json") | ConvertFrom-Json
        $second.Pack | Should BeNullOrEmpty
        @($second.Search).Count | Should Be 6
    }
}
