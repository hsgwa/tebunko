# 本文インデックスのファイルの書き出し（tebunko\index\content_index_store.ps1 の publishIndexFolders）のテスト。
# 書き出し待ちのフォルダの 1 つが失敗しても、ほかのフォルダの書き出しと状態ファイルが続くことを確かめる
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    function script:newPublishTree {
        # index\A・B・C の 3 つのフォルダに、TSV（元のファイルごとのフォルダ）を置く。書き出し待ちの表（フォルダ → 無くなった元のファイル名）を返す
        param ([string]$root)
        $pending = New-Object 'System.Collections.Generic.Dictionary[string,object]'
        foreach ($name in "A", "B", "C") {
            newTsv "$root\index\$name\新規.xlsx\$(toIndexFileName "S")" @("$name の単価")
            $pending["$root\index\$name"] = @()
        }
        return $pending
    }

    function script:publishTree {
        param ($pending, [string]$root)
        return publishIndexFolders $pending "$root\index" "$root\system_index" "$root\state.tsv"
    }

    function script:getFileHash {
        param ([string]$path)
        return (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    }
}

Describe "publishIndexFolders: 1 つのフォルダが失敗しても残りを続ける" -Tag Io {
    BeforeEach {
        Mock writeIndexerLog {}
    }

    It "前の本文インデックスのファイルが壊れたフォルダ（<kind>）は何も書き換えずに飛ばし、ほかのフォルダは書き出して状態ファイルに入れる" -TestCases @(
        @{ kind = "版の違い" },
        @{ kind = "版なし" },
        @{ kind = "ファイル名の行に改行が無い" },
        @{ kind = "上限より大きい" },
        @{ kind = "開けない" }
    ) {
        param ($kind)
        $root = Join-Path $TestDrive ("broken_" + [guid]::NewGuid().ToString("N"))
        $pending = newPublishTree $root
        $mark = [string]${contentIndexMark}
        $broken = "$root\index\B\content_index.xlsx.001.tsv"
        $text = switch ($kind) {
            "版の違い" { "$mark 版=1`n" }
            "版なし" { "壊れた中身`n" }
            "ファイル名の行に改行が無い" { "$mark 版=${contentIndexVersion}`n$mark ファイル名=古い.xlsx" }
            default { "$mark 版=${contentIndexVersion}`n" }
        }
        writeContentIndexFile $broken $text
        if ($kind -eq "上限より大きい") {
            $stream = [System.IO.File]::OpenWrite($broken)
            try { $stream.SetLength(${contentIndexSearchMaxBytes} + 1) } finally { $stream.Dispose() }
        }
        if ($kind -eq "開けない") {
            Mock readContentIndexText { throw "開けません" } -ParameterFilter { $path -like "*\index\B\*" }
        }
        $before = getFileHash $broken

        $result = publishTree $pending $root

        $result.Count | Should -Be 2
        @($result.Skipped).Count | Should -Be 1
        $result.Skipped[0].Rel | Should -Be "B"
        $result.Skipped[0].Stage | Should -Be "本文インデックス"
        foreach ($name in "A", "C") {
            [System.IO.File]::Exists("$root\index\$name\content_index.xlsx.001.tsv") | Should -Be $true
            [System.IO.File]::Exists("$root\system_index\$name\${systemIndexFileName}") | Should -Be $true
            @([System.IO.Directory]::GetDirectories("$root\index\$name")).Count | Should -Be 0
        }
        $state = readSystemIndexState "$root\state.tsv"
        $state.Pending.ContainsKey("A\${systemIndexFileName}") | Should -Be $true
        $state.Pending.ContainsKey("C\${systemIndexFileName}") | Should -Be $true
        $state.Pending.ContainsKey("B\${systemIndexFileName}") | Should -Be $false
        getFileHash $broken | Should -Be $before
        [System.IO.Directory]::Exists("$root\index\B\新規.xlsx") | Should -Be $true
        Should -Invoke writeIndexerLog -Times 1 -Exactly -Scope It -ParameterFilter { $text -like "*本文インデックスにまとめられなかったフォルダがあります（B）: 本文インデックス: *インデックスを作り直すと直ります" -and $color -eq "Yellow" }
    }

    It "システムインデックスの書き出しで落ちたフォルダは、本文インデックスを新しいまま残し、状態ファイルに入れずに飛ばす" {
        $root = Join-Path $TestDrive "systemfail"
        $pending = newPublishTree $root
        Mock writeSystemIndexFolder { throw "書けません" } -ParameterFilter { $folder -like "*\index\B" }

        $result = publishTree $pending $root

        $result.Count | Should -Be 2
        @($result.Skipped).Count | Should -Be 1
        $result.Skipped[0].Rel | Should -Be "B"
        $result.Skipped[0].Stage | Should -Be "システムインデックス"
        (readContentIndexText "$root\index\B\content_index.xlsx.001.tsv") | Should -BeLike "*B の単価*"
        [System.IO.Directory]::Exists("$root\index\B\新規.xlsx") | Should -Be $false
        $state = readSystemIndexState "$root\state.tsv"
        $state.Pending.ContainsKey("B\${systemIndexFileName}") | Should -Be $false
        foreach ($name in "A", "C") {
            [System.IO.File]::Exists("$root\system_index\$name\${systemIndexFileName}") | Should -Be $true
            $state.Pending.ContainsKey("$name\${systemIndexFileName}") | Should -Be $true
        }
        Should -Invoke writeIndexerLog -Times 1 -Exactly -Scope It -ParameterFilter { $text -like "*（B）: システムインデックス: *" }
    }
}

Describe "convertFolderToContentIndex: 前の本文インデックスのファイルの大きさ" -Tag Io {
    It "上限（contentIndexSearchMaxBytes）より 1 バイト大きい前のファイルは、読まずに例外にする" {
        $dir = Join-Path $TestDrive "bigfile\営業"
        newTsv "$dir\A.xlsx\$(toIndexFileName "S")" @("A の単価")
        $old = "$dir\content_index.xlsx.001.tsv"
        writeContentIndexFile $old (convertToContentIndexText @(@{ Name = "B.xlsx"; Places = @(@{ Place = "S"; Text = "前の中身" }) }))
        Mock readContentIndexText { throw "読んではいけない" }
        $stream = [System.IO.File]::OpenWrite($old)
        try { $stream.SetLength(${contentIndexSearchMaxBytes} + 1) } finally { $stream.Dispose() }

        { convertFolderToContentIndex $dir $dir } | Should -Throw "*大きすぎます*"
        Should -Invoke readContentIndexText -Times 0 -Exactly -Scope It
    }
}
