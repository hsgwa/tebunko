# 前の回に数えたテキストの拡張子ごとの件数の控え（tebunko\core\text_ext_counts.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "text_ext_counts（件数の控え）" -Tag Io {
    BeforeEach {
        $path = Join-Path $TestDrive "store\text_ext_counts_$([guid]::NewGuid()).json"
    }

    It "無いファイルは空" {
        (readTextExtCounts $path).Count | Should -Be 0
    }

    It "書いて読み直すと同じ。ほかの名前の項目は残る" {
        writeTextExtCounts "営業" @{ ".txt" = 12; ".log" = 3 } "2024/10/14 15:30:00" $path
        writeTextExtCounts "経理" @{ ".md" = 5 } "2024/10/15 09:00:00" $path
        writeTextExtCounts "営業" @{ ".txt" = 20 } "2024/10/16 10:00:00" $path
        $read = readTextExtCounts $path
        $read.Count | Should -Be 2
        $read["営業"].Time | Should -Be "2024/10/16 10:00:00"
        $read["営業"].Counts[".txt"] | Should -Be 20
        $read["営業"].Counts.ContainsKey(".log") | Should -BeFalse
        $read["経理"].Counts[".md"] | Should -Be 5
        Test-Path -LiteralPath "$path.tmp" | Should -BeFalse
    }

    It "<name>は控え無し（空）として読む" -TestCases @(
        @{ name = "壊れた JSON"; text = "{ これは JSON ではない" }
        @{ name = "配列"; text = '[1, 2]' }
        @{ name = "空のファイル"; text = '' }
    ) {
        param ($name, $text)
        [void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path))
        [System.IO.File]::WriteAllText($path, $text)
        (readTextExtCounts $path).Count | Should -Be 0
    }

    It "形の違う項目・負の件数は捨て、読める項目だけ返す" {
        [void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path))
        [System.IO.File]::WriteAllText($path, '{ "営業": { "Time": "t", "Counts": { ".txt": 4, ".log": -1, ".md": "x" } }, "経理": 5, "総務": { "Time": "t" } }')
        $read = readTextExtCounts $path
        @($read.Keys) | Should -Be @("営業")
        $read["営業"].Counts.Count | Should -Be 1
        $read["営業"].Counts[".txt"] | Should -Be 4
    }

    It "改名で項目のキーを付け替え、削除で消す。無い名前は何もしない" {
        writeTextExtCounts "営業" @{ ".txt" = 1 } "t1" $path
        writeTextExtCounts "経理" @{ ".md" = 2 } "t2" $path
        renameTextExtCounts "営業" "新営業" $path
        $read = readTextExtCounts $path
        $read.ContainsKey("営業") | Should -BeFalse
        $read["新営業"].Counts[".txt"] | Should -Be 1
        removeTextExtCounts "新営業" $path
        @((readTextExtCounts $path).Keys) | Should -Be @("経理")
        $before = [System.IO.File]::ReadAllText($path)
        renameTextExtCounts "無い" "別" $path
        removeTextExtCounts "無い" $path
        [System.IO.File]::ReadAllText($path) | Should -Be $before
    }
}
