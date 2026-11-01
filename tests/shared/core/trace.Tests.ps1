# 処理の記録（shared\core\trace.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    # 書いたファイル（TSV）の本文の行（見出しを除く）を、ファイル名の順に返す
    function readTraceRows {
        param ([string]$Dir)

        if (-not (Test-Path -LiteralPath $Dir)) {
            return @()
        }
        $rows = @()
        foreach ($file in @(Get-ChildItem -LiteralPath $Dir -Filter "*.tsv" | Sort-Object Name)) {
            $rows += @([System.IO.File]::ReadAllLines($file.FullName, [System.Text.Encoding]::UTF8) | Select-Object -Skip 1)
        }
        return $rows
    }
}

Describe "initTrace・testTrace（指定の読み方）" -Tag Unit {
    It "指定 '<Spec>' で、<Category> は <Expected>" -TestCases @(
        @{ Spec = "startup,search"; Category = "startup"; Expected = $true }
        @{ Spec = "startup,search"; Category = "search"; Expected = $true }
        @{ Spec = "startup,search"; Category = "preview"; Expected = $false }
        @{ Spec = "*"; Category = "preview"; Expected = $true }
        @{ Spec = " startup , search "; Category = "search"; Expected = $true }
        @{ Spec = ""; Category = "startup"; Expected = $false }
        @{ Spec = $null; Category = "startup"; Expected = $false }
        @{ Spec = ",,"; Category = "startup"; Expected = $false }
    ) {
        initTrace $Spec
        testTrace $Category | Should -Be $Expected
    }
}

Describe "markTrace・writeTrace" -Tag Io {
    BeforeEach {
        $traceDir = Join-Path $TestDrive ("trace_" + [guid]::NewGuid().ToString("N"))
        initTrace "startup,search"
    }

    It "切りの区分では、markTrace は何も控えず、writeTrace はファイルを作らない" {
        markTrace "preview" "選ぶ"
        writeTrace "preview" $traceDir

        Test-Path -LiteralPath $traceDir | Should -Be $false
    }

    It "入りの区分は、控えた順に TSV へ足す（時刻・経過ミリ秒・区分・名前・スレッド）" {
        markTrace "search" "送り"
        markTrace "search" "最初の行"
        writeTrace "search" $traceDir

        $file = @(Get-ChildItem -LiteralPath $traceDir -Filter "*.tsv")
        $file.Count | Should -Be 1
        $file[0].Name | Should -Match ('^\d{8}-\d{6}_' + $PID + '_\d+\.tsv$')
        [System.IO.File]::ReadAllLines($file[0].FullName, [System.Text.Encoding]::UTF8)[0] | Should -Be "時刻`t経過ミリ秒`t区分`t名前`tスレッド"
        $rows = @(readTraceRows $traceDir)
        $rows.Count | Should -Be 2
        $cells = $rows[0] -split "`t"
        $cells.Count | Should -Be 5
        $cells[0] | Should -Match '^\d{2}:\d{2}:\d{2}\.\d{7}$'
        [double]$cells[1] | Should -BeGreaterOrEqual 0
        $cells[2] | Should -Be "search"
        $cells[3] | Should -Be "送り"
        ($rows[1] -split "`t")[3] | Should -Be "最初の行"
        [double]($rows[1] -split "`t")[1] | Should -BeGreaterOrEqual ([double]$cells[1])
    }

    It "2 回に分けて書くと、同じファイルの末尾に足し、見出しは 1 回だけ" {
        markTrace "search" "1 回目"
        writeTrace "search" $traceDir
        markTrace "search" "2 回目"
        writeTrace "search" $traceDir

        @(Get-ChildItem -LiteralPath $traceDir -Filter "*.tsv").Count | Should -Be 1
        $rows = @(readTraceRows $traceDir)
        $rows.Count | Should -Be 2
        ($rows[1] -split "`t")[3] | Should -Be "2 回目"
    }

    It "writeTrace は渡した区分だけを書いて消し、ほかの区分の控えを残す" {
        markTrace "startup" "起動"
        markTrace "search" "送り"
        writeTrace "search" $traceDir
        @(readTraceRows $traceDir).Count | Should -Be 1

        writeTrace "startup" $traceDir
        $rows = @(readTraceRows $traceDir)
        $rows.Count | Should -Be 2
        ($rows[1] -split "`t")[2] | Should -Be "startup"

        # 書いて消した控えは、もう一度書いても増えない
        writeTrace "search" $traceDir
        @(readTraceRows $traceDir).Count | Should -Be 2
    }

    It "名前の中のタブ・改行は空白にして、1 行に収める" {
        markTrace "search" "a`tb`r`nc"
        writeTrace "search" $traceDir

        $rows = @(readTraceRows $traceDir)
        $rows.Count | Should -Be 1
        ($rows[0] -split "`t")[3] | Should -Be "a b  c"
    }

    It "-Once で書いた区分は、以後 markTrace しても控えが増えない（2 回目の検索の形）" {
        markTrace "startup" "起動"
        writeTrace "startup" $traceDir -Once
        markTrace "startup" "2 回目の送り"
        writeTrace "startup" $traceDir -Once

        $rows = @(readTraceRows $traceDir)
        $rows.Count | Should -Be 1
        ($rows[0] -split "`t")[3] | Should -Be "起動"
    }

    It "2 回読み込んでも、控えが消えない" {
        markTrace "search" "送り"
        . "$PSScriptRoot\..\..\..\scripts\shared\core\trace.ps1"
        writeTrace "search" $traceDir

        @(readTraceRows $traceDir).Count | Should -Be 1
    }

    It "ネットワークのパスには書かない" {
        markTrace "search" "送り"
        { writeTrace "search" "\\server\share\trace" } | Should -Not -Throw
        # 控えは捨てられ、あとで別の場所に書いても出てこない
        writeTrace "search" $traceDir
        Test-Path -LiteralPath $traceDir | Should -Be $false
    }

    It "書けない場所（ファイルがフォルダの位置にある）でも、例外を出さない" {
        $blocker = Join-Path $TestDrive "blocker"
        [System.IO.File]::WriteAllText($blocker, "x")
        markTrace "search" "送り"

        { writeTrace "search" (Join-Path $blocker "trace") } | Should -Not -Throw
    }
}

Describe "スレッドの間での起点" -Tag Io {
    It "同じ起点を渡した別のスレッドの経過がそろい、スレッドごとに別のファイルへ書く" {
        $traceDir = Join-Path $TestDrive "threads"
        $origin = [DateTime]::UtcNow.Ticks - 10000000  # 1 秒前
        initTrace "search" $origin
        markTrace "search" "画面"

        $tracePs1 = (Resolve-Path "$PSScriptRoot\..\..\..\scripts\shared\core\trace.ps1").Path
        $ps = [powershell]::Create()
        try {
            $null = $ps.AddScript({
                param ($tracePs1, $origin, $traceDir)
                . $tracePs1
                initTrace "search" $origin
                markTrace "search" "司令"
                writeTrace "search" $traceDir
            }).AddArgument($tracePs1).AddArgument($origin).AddArgument($traceDir)
            $null = $ps.Invoke()
        } finally {
            $ps.Dispose()
        }
        writeTrace "search" $traceDir

        @(Get-ChildItem -LiteralPath $traceDir -Filter "*.tsv").Count | Should -Be 2
        $rows = @(readTraceRows $traceDir)
        $rows.Count | Should -Be 2
        foreach ($row in $rows) {
            $elapsed = [double]($row -split "`t")[1]
            # 起点が 1 秒前なので、どちらも 1 秒より後で、そろっている
            $elapsed | Should -BeGreaterThan 999
            $elapsed | Should -BeLessThan 60000
        }
    }
}

Describe "getTraceSpec・testTraceActive（ほかのスレッドへ渡す指定）" -Tag Unit {
    It "initTrace に渡した指定と起点を返し、どれかの区分が入っていれば testTraceActive が真になる" -TestCases @(
        @{ Spec = "startup,search"; Active = $true }
        @{ Spec = "*"; Active = $true }
        @{ Spec = ""; Active = $false }
        @{ Spec = " , "; Active = $false }
    ) {
        $origin = [DateTime]::UtcNow.Ticks - 5000
        initTrace $Spec $origin

        (getTraceSpec).Spec | Should -Be $Spec
        (getTraceSpec).OriginTicks | Should -Be $origin
        testTraceActive | Should -Be $Active
    }
}

Describe "testTraceNetworkPath" -Tag Unit {
    It "<Path> は、ネットワークの場所か: <Expected>" -TestCases @(
        @{ Path = "\\server\share\trace"; Expected = $true }
        @{ Path = "\\?\UNC\server\share\trace"; Expected = $true }
        @{ Path = "C:\work\trace"; Expected = $false }
        @{ Path = "\\?\C:\work\trace"; Expected = $false }
    ) {
        testTraceNetworkPath $Path | Should -Be $Expected
    }
}
