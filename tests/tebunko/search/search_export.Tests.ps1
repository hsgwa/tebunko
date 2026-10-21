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

    function script:getReferenceText {
        # 上限なしの結果を、今までの書き方（writeSearchResult）で書いたファイルの中身
        param ([string]$word, $hits)
        $path = Join-Path $TestDrive ("ref-" + [guid]::NewGuid().ToString("N") + ".txt")
        $writer = New-Object System.IO.StreamWriter($path, $false, ${utf8Bom})
        try { writeSearchResult $writer $word @($hits) } finally { $writer.Close() }
        return [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
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
        It "上限 <Limit>・<Mode>・<Filter>: 表 + 溢れの出力が、上限なしの出力と同じになる" -ForEach @(
            @{ Limit = 1; Mode = "single"; Filter = "" }
            @{ Limit = 5; Mode = "parallel"; Filter = "" }
            @{ Limit = 7; Mode = "parallel"; Filter = "" }
            @{ Limit = 12; Mode = "single"; Filter = "" }
            @{ Limit = 100; Mode = "single"; Filter = "" }
            @{ Limit = 5; Mode = "parallel"; Filter = "book3" }
            @{ Limit = 5; Mode = "single"; Filter = "BOOK4" }
            @{ Limit = 5; Mode = "single"; Filter = "存在しない" }
        ) {
            $own = newExportWorkspace "w-$Limit-$Mode-$Filter"
            $dir = getSearchOverflowDir $own.TmpRoot 1
            $workers = if ($Mode -eq "single") { 1 } else { 3 }
            $all = searchContentIndex "多い" $contentIndexFiles $true 0 -workerCount $workers -taskBytes 1
            $expectedHits = @($all.Hits | Where-Object { testHitFilter $_.Book $Filter })
            $result = searchContentIndex "多い" $contentIndexFiles $true $Limit -workerCount $workers -taskBytes 1 -checkEvery 1 -overflowDir $dir
            $shown = @($result.Hits | Where-Object { testHitFilter $_.Book $Filter })
            $out = Join-Path $own.Dir "search_results.txt"
            $request = newExportRequest (newSource $own $result $dir) $out $shown $Filter "多い"

            invokeSearchExport $request

            $request.Outcome | Should -Be "Done"
            $request.Finished | Should -BeTrue
            $request.Written | Should -Be $expectedHits.Count
            $request.Scope | Should -Be $(if ($result.Truncated) { "Overflow" } else { "All" })
            [System.IO.File]::ReadAllText($out, [System.Text.Encoding]::UTF8) | Should -Be (getReferenceText "多い" $expectedHits)
            [System.IO.File]::Exists("$out.new") | Should -BeFalse
            @(Get-ChildItem -LiteralPath (Join-Path (getSearchOverflowPidDir $own.TmpRoot) "export") -ErrorAction SilentlyContinue).Count | Should -Be 0
        }

        It "ファイルの先頭は BOM 付きの UTF-8 で、行の区切りは CRLF" {
            $own = newExportWorkspace "w-bom"
            $dir = getSearchOverflowDir $own.TmpRoot 1
            $result = searchContentIndex "多い" $contentIndexFiles $true 3 -workerCount 1 -overflowDir $dir
            $out = Join-Path $own.Dir "search_results.txt"
            $request = newExportRequest (newSource $own $result $dir) $out @($result.Hits) "" "多い"
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
            $out = Join-Path $own.Dir "search_results.txt"
            $request = newExportRequest $source $out @($result.Hits) "" "多い"

            invokeSearchExport $request

            $request.Outcome | Should -Be "Done"
            $request.Scope | Should -Be "HeadOnly"
            $request.HeadOnlyReason | Should -Be $Reason
            $request.Written | Should -Be 3
            [System.IO.File]::ReadAllText($out, [System.Text.Encoding]::UTF8) | Should -Be (getReferenceText "多い" $result.Hits)
        }
    }

    Context "中止・失敗" {
        BeforeEach {
            $own = newExportWorkspace "w-stop-$([guid]::NewGuid().ToString('N'))"
            $dir = getSearchOverflowDir $own.TmpRoot 1
            $result = searchContentIndex "多い" $contentIndexFiles $true 3 -workerCount 1 -overflowDir $dir
            $out = Join-Path $own.Dir "search_results.txt"
            [System.IO.File]::WriteAllText($out, "前の結果", ${utf8Bom})
        }

        It "始める前に中止されていたら、書かずに前のファイルを残す" {
            $request = newExportRequest (newSource $own $result $dir) $out @($result.Hits) "" "多い"
            $request.Stop = $true
            invokeSearchExport $request
            $request.Outcome | Should -Be "Cancelled"
            $request.Finished | Should -BeTrue
            [System.IO.File]::ReadAllText($out, [System.Text.Encoding]::UTF8) | Should -Be "前の結果"
        }

        It "書いている途中で中止されたら、一時ファイルと .new を消し、前のファイルを残す" {
            $many = 1..3000 | ForEach-Object { [pscustomobject]@{ RelDir = ""; Book = "b.txt"; Location = "p"; LineNumber = $_; Line = "行 $_" } }
            $request = newExportRequest (newSource $own $result $dir) $out $many "" "多い"
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
            $request = newExportRequest (newSource $own $result $dir) $out @($result.Hits) "" "多い"
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
            $request = newExportRequest (newSource $own $result $dir) $out @($result.Hits) "" "多い"
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
            $request = newExportRequest (newSource $own $result $dir) $out @($result.Hits) "" "多い"
            invokeSearchExport $request
            $request.Error | Should -BeNullOrEmpty
            $request.Outcome | Should -Be "Done"
            [System.IO.File]::Exists("$out.new") | Should -BeFalse
            [System.IO.File]::ReadAllText($out, [System.Text.Encoding]::UTF8) | Should -Not -Be "前の結果"
        }

        It "出力先がまだ無くても書ける" {
            [System.IO.File]::Delete($out)
            $request = newExportRequest (newSource $own $result $dir) $out @($result.Hits) "" "多い"
            invokeSearchExport $request
            $request.Outcome | Should -Be "Done"
            [System.IO.File]::Exists($out) | Should -BeTrue
        }
    }
}

Describe "getSearchExportScope" -Tag Unit {
    It "<Name>" -ForEach @(
        @{ Name = "上限を超えていない"; Source = @{ Truncated = $false }; Scope = "All"; Reason = "" }
        @{ Name = "上限を超え、溢れが書ける"; Source = @{ Truncated = $true; CountTimeout = $false; OverflowFailed = $false; Cancelled = $false }; Scope = "Overflow"; Reason = "" }
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
