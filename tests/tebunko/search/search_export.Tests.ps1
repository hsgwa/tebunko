# 結果のファイル出力（tebunko\search\search_export.ps1）のテスト。
# 件数の上限で表を小さくした検索の「表 + 溢れ」を出力したファイルが、上限なしの結果を出力したファイルと同じになることを確かめる
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    function script:newExportWorkspace {
        param ([string]$name)
        $work = Join-Path $TestDrive $name
        [void][System.IO.Directory]::CreateDirectory($work)
        return [Workspace]::new($work)
    }

    function script:newSource {
        # 検索の要求の代わり（invokeSearchExport が読む欄だけ）
        param ($own, $result, [string]$overflowDir)
        return @{
            WorkDir = $own.Dir; OverflowDir = $overflowDir
            Truncated = [bool]$result.Truncated; CountTimeout = $false; OverflowFailed = $false; Cancelled = $false
        }
    }

    function script:getReferenceRows {
        # 上限なしの結果を、形式ごとの列（toResultCells）にした期待値。見出しの行から始まる
        param ([string]$format, $hits)
        $rows = @($hits | ForEach-Object { , (toResultCells $_.RelDir $_.Book $_.Location ([int]$_.LineNumber) $_.Line $format) })
        if ($format -eq "xlsx") {
            $header = @("ファイル名", "場所", "種別", "行", "該当行")
        } else {
            $columns = 0
            foreach ($row in $rows) { $columns = [math]::Max($columns, $row.Count - 4) }
            $header = (toResultHeader $columns).Split("`t")
        }
        return @(, $header) + $rows
    }

    function script:readExportRows {
        # 書いたファイルを行（セルの配列）に読み戻す。csv・tsv は囲みの要らない値だけのテストデータなので、区切りで分けるだけ
        param ([string]$path, [string]$format)
        if ($format -eq "xlsx") {
            $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
            try {
                $reader = New-Object System.IO.StreamReader($zip.GetEntry("xl/worksheets/sheet1.xml").Open(), [System.Text.Encoding]::UTF8)
                $xml = [xml]$reader.ReadToEnd()
                $reader.Dispose()
            } finally {
                $zip.Dispose()
            }
            return @($xml.worksheet.sheetData.row | ForEach-Object { , @($_.c | ForEach-Object { $_.is.t.InnerText }) })
        }
        $delimiter = if ($format -eq "tsv") { "`t" } else { "," }
        $text = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8).TrimStart([char]0xFEFF)
        $lines = $text.Split(@("`r`n"), [System.StringSplitOptions]::None)
        return @($lines[0..($lines.Count - 2)] | ForEach-Object { , $_.Split($delimiter) })
    }

    function script:assertSameRows {
        param ($actual, $expected)
        $actual.Count | Should -Be $expected.Count
        for ($i = 0; $i -lt $expected.Count; $i++) {
            ($actual[$i] -join [char]1) | Should -Be ($expected[$i] -join [char]1)
        }
    }
}

Describe "結果のファイル出力" -Tag Io {
    BeforeAll {
        $tsvRoot = Join-Path $TestDrive "ex_tsv"
        for ($i = 0; $i -lt 6; $i++) {
            $lines = @("多い $i", "x", "多い 多い $i", "同じ行に多い 多い 多い")
            if ($i -ge 4) { $lines += "後ろだけ $i" }
            newTsv "$tsvRoot\d$i\book$i.xlsx\S.tsv" $lines
        }
        $contentIndexFiles = newContentIndexFiles $tsvRoot (Join-Path $TestDrive "ex_content_index")
    }

    Context "上限を超えた分も書く" {
        It "上限 <Limit>・<Mode>・<Filter>・<Format>: 表 + 溢れの出力が、上限なしの出力と同じになる" -ForEach @(
            @{ Limit = 1; Mode = "single"; Filter = ""; Format = "xlsx" }
            @{ Limit = 5; Mode = "parallel"; Filter = ""; Format = "csv" }
            @{ Limit = 7; Mode = "parallel"; Filter = ""; Format = "tsv" }
            @{ Limit = 12; Mode = "single"; Filter = ""; Format = "xlsx" }
            @{ Limit = 100; Mode = "single"; Filter = ""; Format = "csv" }
            @{ Limit = 5; Mode = "parallel"; Filter = "book3"; Format = "tsv" }
            @{ Limit = 5; Mode = "single"; Filter = "BOOK4"; Format = "xlsx" }
            @{ Limit = 5; Mode = "single"; Filter = "存在しない"; Format = "csv" }
        ) {
            $own = newExportWorkspace "w-$Limit-$Mode-$Filter-$Format"
            $dir = getSearchOverflowDir $own.TmpRoot 1
            $workers = if ($Mode -eq "single") { 1 } else { 3 }
            $all = searchContentIndex "多い" $contentIndexFiles $true 0 -workerCount $workers -taskBytes 1
            $expectedHits = @($all.Hits | Where-Object { testHitFilter $_.Book $Filter })
            $result = searchContentIndex "多い" $contentIndexFiles $true $Limit -workerCount $workers -taskBytes 1 -checkEvery 1 -overflowDir $dir
            $shown = @($result.Hits | Where-Object { testHitFilter $_.Book $Filter })
            $out = Join-Path $own.Dir "out.$Format"
            $request = newExportRequest (newSource $own $result $dir) $out $shown $Filter "多い" $Format $true

            invokeSearchExport $request

            $request.Outcome | Should -Be "Done"
            $request.Finished | Should -BeTrue
            $request.Written | Should -Be $expectedHits.Count
            $request.Scope | Should -Be $(if ($result.Truncated) { "Overflow" } else { "All" })
            assertSameRows (readExportRows $out $Format) (getReferenceRows $Format $expectedHits)
            [System.IO.File]::Exists("$out.new") | Should -BeFalse
            @(Get-ChildItem -LiteralPath (Join-Path (getSearchOverflowPidDir $own.TmpRoot) "export") -ErrorAction SilentlyContinue).Count | Should -Be 0
        }

        It "<Format>: ファイルの先頭は BOM 付きの UTF-8 で、行の区切りは CRLF" -ForEach @(@{ Format = "csv" }, @{ Format = "tsv" }) {
            $own = newExportWorkspace "w-bom-$Format"
            $dir = getSearchOverflowDir $own.TmpRoot 1
            $result = searchContentIndex "多い" $contentIndexFiles $true 3 -workerCount 1 -overflowDir $dir
            $out = Join-Path $own.Dir "out.$Format"
            $request = newExportRequest (newSource $own $result $dir) $out @($result.Hits) "" "多い" $Format $true
            invokeSearchExport $request
            $bytes = [System.IO.File]::ReadAllBytes($out)
            $bytes[0..2] | Should -Be @(0xEF, 0xBB, 0xBF)
            $text = [System.Text.Encoding]::UTF8.GetString($bytes)
            ([regex]::Matches($text, "(?<!\r)\n")).Count | Should -Be 0
        }
    }

    Context "先頭だけを書く" {
        It "<Reason>: 溢れは書かず、表の分だけを出力して理由を残す" -ForEach @(
            @{ Reason = "CountTimeout" }
            @{ Reason = "OverflowFailed" }
            @{ Reason = "Cancelled" }
        ) {
            $own = newExportWorkspace "w-head-$Reason"
            $dir = getSearchOverflowDir $own.TmpRoot 1
            $result = searchContentIndex "多い" $contentIndexFiles $true 3 -workerCount 1 -overflowDir $dir
            $source = newSource $own $result $dir
            $source[$Reason] = $true
            $out = Join-Path $own.Dir "out.csv"
            $request = newExportRequest $source $out @($result.Hits) "" "多い" "csv" $true

            invokeSearchExport $request

            $request.Outcome | Should -Be "Done"
            $request.Scope | Should -Be "HeadOnly"
            $request.HeadOnlyReason | Should -Be $Reason
            $request.Written | Should -Be 3
            assertSameRows (readExportRows $out "csv") (getReferenceRows "csv" $result.Hits)
        }
    }

    Context "中止・失敗" {
        BeforeEach {
            $own = newExportWorkspace "w-stop-$([guid]::NewGuid().ToString('N'))"
            $dir = getSearchOverflowDir $own.TmpRoot 1
            $result = searchContentIndex "多い" $contentIndexFiles $true 3 -workerCount 1 -overflowDir $dir
            $out = Join-Path $own.Dir "out.csv"
            [System.IO.File]::WriteAllText($out, "前の結果", ${utf8Bom})
        }

        It "始める前に中止されていたら、書かずに前のファイルを残す" {
            $request = newExportRequest (newSource $own $result $dir) $out @($result.Hits) "" "多い" "csv" $true
            $request.Stop = $true
            invokeSearchExport $request
            $request.Outcome | Should -Be "Cancelled"
            $request.Finished | Should -BeTrue
            [System.IO.File]::ReadAllText($out, [System.Text.Encoding]::UTF8) | Should -Be "前の結果"
        }

        It "書いている途中で中止されたら、一時ファイルと .new を消し、前のファイルを残す" {
            $many = 1..3000 | ForEach-Object { [pscustomobject]@{ RelDir = ""; Book = "b.txt"; Location = "p"; LineNumber = $_; Line = "行 $_" } }
            $request = newExportRequest (newSource $own $result $dir) $out $many "" "多い" "csv" $true
            # 出力先の確かめを通ったあと（書き始める前）に中止を立てる。書いている途中の確かめ（1000 行ごと）で止まる
            $script:exportRequest = $request
            Mock testSearchExportTarget { $script:exportRequest.Stop = $true; return "" }
            invokeSearchExport $request
            $request.Outcome | Should -Be "Cancelled"
            [System.IO.File]::ReadAllText($out, [System.Text.Encoding]::UTF8) | Should -Be "前の結果"
            [System.IO.File]::Exists("$out.new") | Should -BeFalse
            @(Get-ChildItem -LiteralPath (Join-Path (getSearchOverflowPidDir $own.TmpRoot) "export") -ErrorAction SilentlyContinue).Count | Should -Be 0
        }

        It "ほかのアプリが開いているときは Locked にして、書き始めない" {
            $request = newExportRequest (newSource $own $result $dir) $out @($result.Hits) "" "多い" "csv" $true
            $lock = [System.IO.File]::Open($out, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
            try {
                invokeSearchExport $request
            } finally {
                $lock.Dispose()
            }
            $request.Outcome | Should -Be "Locked"
            $request.Finished | Should -BeTrue
            [System.IO.File]::Exists("$out.new") | Should -BeFalse
            [System.IO.File]::ReadAllText($out, [System.Text.Encoding]::UTF8) | Should -Be "前の結果"
        }

        It "読み取り専用のときは Denied にして、書き始めない" {
            $request = newExportRequest (newSource $own $result $dir) $out @($result.Hits) "" "多い" "csv" $true
            [System.IO.File]::SetAttributes($out, [System.IO.FileAttributes]::ReadOnly)
            try {
                invokeSearchExport $request
            } finally {
                [System.IO.File]::SetAttributes($out, [System.IO.FileAttributes]::Normal)
            }
            $request.Outcome | Should -Be "Denied"
            [System.IO.File]::ReadAllText($out, [System.Text.Encoding]::UTF8) | Should -Be "前の結果"
        }

        It "前の出力が止められて残った .new は、次の出力の始めに消える" {
            [System.IO.File]::WriteAllText("$out.new", "古い", ${utf8Bom})
            $request = newExportRequest (newSource $own $result $dir) $out @($result.Hits) "" "多い" "csv" $true
            invokeSearchExport $request
            $request.Error | Should -BeNullOrEmpty
            $request.Outcome | Should -Be "Done"
            [System.IO.File]::Exists("$out.new") | Should -BeFalse
            [System.IO.File]::ReadAllText($out, [System.Text.Encoding]::UTF8) | Should -Not -Be "前の結果"
        }

        It "前の出力が強制終了で残した .part も、次の出力の始めに消える" {
            [System.IO.File]::WriteAllText("$out.part", "古い", ${utf8Bom})
            $request = newExportRequest (newSource $own $result $dir) $out @($result.Hits) "" "多い" "csv" $true
            invokeSearchExport $request
            $request.Outcome | Should -Be "Done"
            [System.IO.File]::Exists("$out.part") | Should -BeFalse
        }

        It "溢れの置き場所が使えないときは、本文の一時ファイルを出力先の横に作り、終わりに消す" {
            Mock testSearchOverflowRoot { return $false }
            $request = newExportRequest (newSource $own $result $dir) $out @($result.Hits) "" "多い" "csv" $true
            invokeSearchExport $request
            $request.Outcome | Should -Be "Done"
            [System.IO.File]::Exists("$out.part") | Should -BeFalse
            [System.IO.File]::ReadAllText($out, [System.Text.Encoding]::UTF8) | Should -Not -Be "前の結果"
        }

        It "書いている途中で想定外の例外が出たら Failed にして、理由を残し、前のファイルを残す" {
            Mock writeSearchExportBody { throw "壊れた" }
            $request = newExportRequest (newSource $own $result $dir) $out @($result.Hits) "" "多い" "csv" $true
            invokeSearchExport $request
            $request.Outcome | Should -Be "Failed"
            $request.Error | Should -Be "壊れた"
            $request.Finished | Should -BeTrue
            [System.IO.File]::ReadAllText($out, [System.Text.Encoding]::UTF8) | Should -Be "前の結果"
        }

        It "溢れの置き場所が消えているときは、表の分だけを書き、理由は OverflowFailed" {
            $source = newSource $own $result ""
            $source.Truncated = $true
            $request = newExportRequest $source $out @($result.Hits) "" "多い" "csv" $true
            invokeSearchExport $request
            $request.Outcome | Should -Be "Done"
            $request.Scope | Should -Be "HeadOnly"
            $request.HeadOnlyReason | Should -Be "OverflowFailed"
            $request.Written | Should -Be 3
        }

        It "上書きを許していないときは、出力先が既にあれば書かずに Exists を返し、前のファイルを残す" {
            $request = newExportRequest (newSource $own $result $dir) $out @($result.Hits) "" "多い" "csv" $false
            invokeSearchExport $request
            $request.Outcome | Should -Be "Exists"
            $request.Finished | Should -BeTrue
            [System.IO.File]::ReadAllText($out, [System.Text.Encoding]::UTF8) | Should -Be "前の結果"
            [System.IO.File]::Exists("$out.new") | Should -BeFalse
        }

        It "上書きを許していても、出力先がまだ無ければ書ける" {
            $fresh = Join-Path $own.Dir "fresh.csv"
            $request = newExportRequest (newSource $own $result $dir) $fresh @($result.Hits) "" "多い" "csv" $false
            invokeSearchExport $request
            $request.Outcome | Should -Be "Done"
        }

        It "xlsx: ほかのアプリが開いているときは Locked にして、前のファイルを残す" {
            $xlsx = Join-Path $own.Dir "out.xlsx"
            [System.IO.File]::WriteAllText($xlsx, "前の結果", ${utf8Bom})
            $request = newExportRequest (newSource $own $result $dir) $xlsx @($result.Hits) "" "多い" "xlsx" $true
            $lock = [System.IO.File]::Open($xlsx, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
            try {
                invokeSearchExport $request
            } finally {
                $lock.Dispose()
            }
            $request.Outcome | Should -Be "Locked"
            [System.IO.File]::Exists("$xlsx.new") | Should -BeFalse
            [System.IO.File]::ReadAllText($xlsx, [System.Text.Encoding]::UTF8) | Should -Be "前の結果"
        }

        It "xlsx: 書いている途中で中止されたら、.new を消し、前のファイルを残す" {
            $xlsx = Join-Path $own.Dir "stop.xlsx"
            [System.IO.File]::WriteAllText($xlsx, "前の結果", ${utf8Bom})
            $many = 1..3000 | ForEach-Object { [pscustomobject]@{ RelDir = ""; Book = "b.txt"; Location = "p"; LineNumber = $_; Line = "行 $_" } }
            $request = newExportRequest (newSource $own $result $dir) $xlsx $many "" "多い" "xlsx" $true
            $script:exportRequest = $request
            Mock testSearchExportTarget { $script:exportRequest.Stop = $true; return "" }
            invokeSearchExport $request
            $request.Outcome | Should -Be "Cancelled"
            [System.IO.File]::ReadAllText($xlsx, [System.Text.Encoding]::UTF8) | Should -Be "前の結果"
            [System.IO.File]::Exists("$xlsx.new") | Should -BeFalse
        }

        It "xlsx: 長すぎるセルを切った数を TruncatedCells に返す" {
            $xlsx = Join-Path $own.Dir "long.xlsx"
            $long = [pscustomobject]@{ RelDir = ""; Book = "b.txt"; Location = "p"; LineNumber = 1; Line = ("あ" * 40000) }
            $short = [pscustomobject]@{ RelDir = ""; Book = "b.txt"; Location = "p"; LineNumber = 2; Line = "短い" }
            $source = newSource $own $result $dir
            $source.Truncated = $false
            $request = newExportRequest $source $xlsx @($long, $short) "" "多い" "xlsx" $true
            invokeSearchExport $request
            $request.Outcome | Should -Be "Done"
            $request.Written | Should -Be 2
            $request.TruncatedCells | Should -Be 1
            $request.DroppedRows | Should -Be 0
        }

        It "出力先がまだ無くても書ける" {
            [System.IO.File]::Delete($out)
            $request = newExportRequest (newSource $own $result $dir) $out @($result.Hits) "" "多い" "csv" $false
            invokeSearchExport $request
            $request.Outcome | Should -Be "Done"
            [System.IO.File]::Exists($out) | Should -BeTrue
        }
    }
}

Describe "getSearchExportScope" -Tag Unit {
    It "<Name>" -ForEach @(
        @{ Name = "上限を超えていない"; Source = @{ Truncated = $false }; Scope = "All"; Reason = "" }
        @{ Name = "上限を超え、溢れが書ける"; Source = @{ Truncated = $true; CountTimeout = $false; OverflowFailed = $false; Cancelled = $false; OverflowDir = $env:TEMP }; Scope = "Overflow"; Reason = "" }
        @{ Name = "上限を超え、溢れの置き場所が消えている"; Source = @{ Truncated = $true; CountTimeout = $false; OverflowFailed = $false; Cancelled = $false; OverflowDir = "$env:TEMP\tebunko-無い-dir" }; Scope = "HeadOnly"; Reason = "OverflowFailed" }
        @{ Name = "上限を超え、溢れの置き場所が空"; Source = @{ Truncated = $true; CountTimeout = $false; OverflowFailed = $false; Cancelled = $false; OverflowDir = "" }; Scope = "HeadOnly"; Reason = "OverflowFailed" }
        @{ Name = "総数の時間切れ"; Source = @{ Truncated = $true; CountTimeout = $true; OverflowFailed = $false; Cancelled = $false }; Scope = "HeadOnly"; Reason = "CountTimeout" }
        @{ Name = "溢れを書けなかった"; Source = @{ Truncated = $true; CountTimeout = $false; OverflowFailed = $true; Cancelled = $false }; Scope = "HeadOnly"; Reason = "OverflowFailed" }
        @{ Name = "検索を止めた"; Source = @{ Truncated = $true; CountTimeout = $false; OverflowFailed = $false; Cancelled = $true }; Scope = "HeadOnly"; Reason = "Cancelled" }
        @{ Name = "理由が重なったら時間切れを先に"; Source = @{ Truncated = $true; CountTimeout = $true; OverflowFailed = $true; Cancelled = $true }; Scope = "HeadOnly"; Reason = "CountTimeout" }
    ) {
        $actual = getSearchExportScope $Source
        $actual.Scope | Should -Be $Scope
        $actual.Reason | Should -Be $Reason
    }
}
