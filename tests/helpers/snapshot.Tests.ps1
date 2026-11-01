# assertSnapshot（tests\helpers\snapshot.ps1）の比べ方・書き直しのテスト。
BeforeAll {
    . "$PSScriptRoot\snapshot.ps1"
    $utf8 = New-Object System.Text.UTF8Encoding($false)
}

Describe "assertSnapshot" -Tag Unit {
    BeforeEach {
        $path = Join-Path $TestDrive "snap\期待値.txt"
        [System.IO.Directory]::CreateDirectory((Split-Path $path)) | Out-Null
        [System.IO.File]::WriteAllText($path, "1 行目`n2 行目`n", $utf8)
        $saved = $env:TEBUNKO_UPDATE_SNAPSHOT
        $env:TEBUNKO_UPDATE_SNAPSHOT = $null
    }
    AfterEach {
        $env:TEBUNKO_UPDATE_SNAPSHOT = $saved
    }

    It "<name>" -TestCases @(
        @{ name = "同じ文字列なら通る"; actual = "1 行目`n2 行目`n" }
        @{ name = "改行が CRLF でも、LF にそろえて通る"; actual = "1 行目`r`n2 行目`r`n" }
    ) {
        param ($name, $actual)
        { assertSnapshot $actual $path } | Should -Not -Throw
    }

    It "<name>" -TestCases @(
        @{ name = "違う行は、行の番号と両方の中身を入れて落ちる"; actual = "1 行目`n別の行`n"; message = "2 行目.*期待値: \[2 行目\] 実際: \[別の行\]" }
        @{ name = "行が足りなければ、足りない行の番号で落ちる"; actual = "1 行目`n"; message = "2 行目.*期待値: \[2 行目\] 実際: \[\]" }
    ) {
        param ($name, $actual, $message)
        $caught = $null
        try { assertSnapshot $actual $path } catch { $caught = $_.Exception.Message }
        $caught | Should -Match $message
    }

    It "期待値のファイルが無ければ落ちる" {
        { assertSnapshot "x" (Join-Path $TestDrive "無い.txt") } | Should -Throw -ExpectedMessage "*期待値のファイルがありません*"
    }

    It "TEBUNKO_UPDATE_SNAPSHOT=1 なら比べずに、UTF-8（BOM なし）・LF で書き直し、書き直した後は通る" {
        $env:TEBUNKO_UPDATE_SNAPSHOT = "1"
        $newPath = Join-Path $TestDrive "新しい\作り直し.txt"
        assertSnapshot "別`r`n中身`r`n" $path
        assertSnapshot "新規`r`n" $newPath
        $bytes = [System.IO.File]::ReadAllBytes($path)
        $utf8.GetString($bytes) | Should -BeExactly "別`n中身`n"
        ($bytes[0] -eq 0xEF) | Should -BeFalse
        [System.IO.File]::ReadAllText($newPath, $utf8) | Should -BeExactly "新規`n"
        $env:TEBUNKO_UPDATE_SNAPSHOT = $null
        { assertSnapshot "別`n中身`n" $path } | Should -Not -Throw
    }
}
