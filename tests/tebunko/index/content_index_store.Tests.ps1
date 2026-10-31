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
        return publishIndexFolders $pending "$root\index"
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
        if ($kind -eq "開けない") {
            Mock readContentIndexText { throw "開けません" } -ParameterFilter { $path -like "*\index\B\*" }
        }
        $before = getFileHash $broken

        $result = publishTree $pending $root

        $result.Count | Should -Be 2
        @($result.Skipped).Count | Should -Be 1
        $result.Skipped[0].Rel | Should -Be "B"
        $result.Skipped[0].Stage | Should -Be "インデックスの書き出し"
        @($result.Published | Sort-Object) | Should -Be @("$root\index\A", "$root\index\C")
        foreach ($name in "A", "C") {
            [System.IO.File]::Exists("$root\index\$name\content_index.xlsx.001.tsv") | Should -Be $true
            @([System.IO.Directory]::GetDirectories("$root\index\$name")).Count | Should -Be 0
        }
        getFileHash $broken | Should -Be $before
        [System.IO.Directory]::Exists("$root\index\B\新規.xlsx") | Should -Be $true
        Should -Invoke writeIndexerLog -Times 1 -Exactly -Scope It -ParameterFilter { $text -like "*インデックスにまとめられなかったフォルダがあります（B）: インデックスの書き出し: *インデックスを作り直すと直ります" -and $color -eq "Yellow" }
    }
}

Describe "検索の上限（contentIndexSearchMaxBytes）を超える本文インデックスのファイルを作らない・直す" -Tag Io {
    BeforeAll {
        . "${scriptsDir}\tebunko\indexer\indexer_lib.ps1"

        function script:newBigBookTsv {
            # 上限（4KB に下げて使う）を超える元のファイルの TSV を作る
            param ([string]$dir)
            newTsv "$dir\$(toIndexFileName "S")" @(1..200 | ForEach-Object { "巨大な行の単価 $_ " + ("あ" * 20) })
        }

        function script:newOversizeContentIndexFile {
            # A.xlsx・B.xlsx を入れた、上限（4KB）を超える前の本文インデックスのファイルを dir に作る
            param ([string]$dir)
            $big = (1..200 | ForEach-Object { "巨大な行の単価 $_ " + ("あ" * 20) }) -join "`n"
            $books = @(
                @{ Name = "A.xlsx"; Places = @(@{ Place = "S"; Text = "小さい単価`n" }) }
                @{ Name = "B.xlsx"; Places = @(@{ Place = "S"; Text = "$big`n" }) }
            )
            [System.IO.Directory]::CreateDirectory($dir) | Out-Null
            writeContentIndexFile "$dir\content_index.xlsx.001.tsv" (convertToContentIndexText $books)
        }
    }

    BeforeEach {
        Mock writeIndexerLog {}
    }

    It "(a) 上限を超える元のファイルは置き場へ移さず取り込みの失敗にし、同じフォルダの小さな元のファイルは検索に出る" {
        ${contentIndexSearchMaxBytes} = 4KB
        $workspace = newTestWorkspace @{} "$TestDrive\fit_ws"
        $tmpDir = "$TestDrive\fit_ws\tmp"
        $pending = New-Object 'System.Collections.Generic.Dictionary[string,object]'
        $pending["$($workspace.IndexDir)\営業"] = @()
        newTsv "$tmpDir\$(toIndexFileName "S")" @("小さい単価")
        publishTsv (getBookDir "営業\A.xlsx")
        Remove-Item -LiteralPath "$tmpDir\*" -Force
        newBigBookTsv $tmpDir
        $caught = $null
        try { publishTsv (getBookDir "営業\B.xlsx") } catch { $caught = $_.Exception.Message }
        $caught | Should -Be "ファイルの文字が多すぎるため更新できません（上限 64 MB）。"
        [System.IO.Directory]::Exists("$($workspace.IndexDir)\営業\B.xlsx") | Should -Be $false

        $result = publishIndexFolders $pending $workspace.IndexDir
        @($result.Skipped).Count | Should -Be 0
        $found = getContentIndexFiles @($workspace.IndexDir)
        foreach ($file in $found.ContentIndexFiles) { $file.Size | Should -BeLessOrEqual ${contentIndexSearchMaxBytes} }
        $search = searchContentIndex "単価" $found.ContentIndexFiles $true
        @($search.Hits | Where-Object { $_.Book -eq "A.xlsx" }).Count | Should -Be 1
        @($search.Skipped).Count | Should -Be 0
    }

    It "(b) 上限を超える前の本文インデックスのファイルがあるフォルダは、書き出しが失敗せず、ファイルを外して中の元のファイルを取り込み直しにする" {
        ${contentIndexSearchMaxBytes} = 4KB
        $root = "$TestDrive\fit_b"
        newOversizeContentIndexFile "$root\index\営業"
        newTsv "$root\index\営業\C.xlsx\$(toIndexFileName "S")" @("新しい単価")
        $pending = New-Object 'System.Collections.Generic.Dictionary[string,object]'
        $pending["$root\index\営業"] = @()

        $result = publishIndexFolders $pending "$root\index"

        @($result.Skipped).Count | Should -Be 0
        $result.Count | Should -Be 1
        @($result.Requeued | Sort-Object) | Should -Be @("営業\A.xlsx", "営業\B.xlsx")
        $files = @(findContentIndexFiles "$root\index")
        $files.Count | Should -Be 1
        $files[0].Size | Should -BeLessOrEqual ${contentIndexSearchMaxBytes}
        $search = searchContentIndex "新しい単価" (getContentIndexFiles @("$root\index")).ContentIndexFiles $true
        @($search.Hits).Count | Should -Be 1
        Should -Invoke writeIndexerLog -Times 1 -Exactly -Scope It -ParameterFilter { $text -like "*大きすぎるインデックスのファイルを外しました（営業\content_index.xlsx.001.tsv）。中の 2 件は次の更新で取り込み直します" -and $color -eq "Yellow" }
    }

    It "(b) この回に入れ替える元のファイルは、上限超えのファイルを外しても取り込み直しにしない" {
        ${contentIndexSearchMaxBytes} = 4KB
        $root = "$TestDrive\fit_b2"
        newOversizeContentIndexFile "$root\index\営業"
        newTsv "$root\index\営業\A.xlsx\$(toIndexFileName "S")" @("入れ替えた単価")
        $pending = New-Object 'System.Collections.Generic.Dictionary[string,object]'
        $pending["$root\index\営業"] = @("B.xlsx")

        $result = publishIndexFolders $pending "$root\index"

        @($result.Requeued) | Should -Be @()
        $search = searchContentIndex "入れ替えた単価" (getContentIndexFiles @("$root\index")).ContentIndexFiles $true
        @($search.Hits | Where-Object { $_.Book -eq "A.xlsx" }).Count | Should -Be 1
    }

    It "(c) 新しく入れる元のファイルが 1 つで上限を超えるときは、入れずに取り込み直しにし、前のコピーも外す" {
        ${contentIndexSearchMaxBytes} = 4KB
        $dir = Join-Path $TestDrive "fit_c\営業"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        writeContentIndexFile "$dir\content_index.xlsx.001.tsv" (convertToContentIndexText @(@{ Name = "B.xlsx"; Places = @(@{ Place = "S"; Text = "前の中身`n" }) }))
        newBigBookTsv "$dir\B.xlsx"
        newTsv "$dir\A.xlsx\$(toIndexFileName "S")" @("小さい単価")

        $result = convertFolderToContentIndex $dir $dir @() $true

        @($result.Requeued) | Should -Be @("B.xlsx")
        $result.Books | Should -Be 1
        $text = readContentIndexText "$dir\content_index.xlsx.001.tsv"
        $text | Should -BeLike "*小さい単価*"
        $text | Should -Not -BeLike "*前の中身*"
        $text | Should -Not -BeLike "*巨大な行*"
        (measureContentIndexPartBytes @(@{ Block = $text })) | Should -BeLessThan 4KB
    }

    It "上限を超える前のファイルは、丸ごとは読まない" {
        $dir = Join-Path $TestDrive "bigfile\営業"
        newTsv "$dir\A.xlsx\$(toIndexFileName "S")" @("A の単価")
        $old = "$dir\content_index.xlsx.001.tsv"
        writeContentIndexFile $old (convertToContentIndexText @(@{ Name = "B.xlsx"; Places = @(@{ Place = "S"; Text = ("前の中身`n" * 300) }) }))
        Mock readContentIndexText { throw "読んではいけない" }
        ${contentIndexSearchMaxBytes} = 1KB

        $result = convertFolderToContentIndex $dir $dir

        @($result.Requeued) | Should -Be @("B.xlsx")
        Should -Invoke readContentIndexText -Times 0 -Exactly -Scope It
    }

    It "readPreviousContentIndexParts / getOversizeRequeueNames: 上限以下のファイルは読んで Parts に入れ、上限を超えるファイルは名前だけを読んで、keep に無い元のファイルを取り込み直し・ファイルを消す対象にする" {
        ${contentIndexSearchMaxBytes} = 4KB
        $dir = "$TestDrive\fit_prev"
        newOversizeContentIndexFile $dir
        writeContentIndexFile "$dir\content_index.docx.001.tsv" (convertToContentIndexText @(@{ Name = "D.docx"; Places = @(@{ Place = "ページ001"; Text = "x`n" }) }))
        $keep = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        [void]$keep.Add("a.xlsx")

        $previous = readPreviousContentIndexParts $dir $keep

        @($previous.Parts).Count | Should -Be 1
        @($previous.Parts)[0].Extension | Should -Be "docx"
        @($previous.Requeue) | Should -Be @("B.xlsx")
        @($previous.Remove) | Should -Be @("content_index.xlsx.001.tsv")
        @($previous.Dropped).Count | Should -Be 1
        $previous.Dropped[0].Count | Should -Be 1
        (getOversizeRequeueNames "$dir\content_index.xlsx.001.tsv" $keep) | Should -Be @("B.xlsx")
    }

    It "readNewContentIndexBooks: 1 つで上限を超える元のファイルは TooLarge に分け、小さいものはまとまりにして TSV の数を数える" {
        ${contentIndexSearchMaxBytes} = 4KB
        $dir = "$TestDrive\fit_new"
        newTsv "$dir\A.xlsx\$(toIndexFileName "S")" @("小さい単価")
        newBigBookTsv "$dir\B.xlsx"

        $new = readNewContentIndexBooks (getIndexFolderBooks $dir)

        @($new.Books | ForEach-Object { $_.Name }) | Should -Be @("A.xlsx")
        @($new.TooLarge) | Should -Be @("B.xlsx")
        $new.Tsv | Should -Be 2
    }

    It "findOversizeContentIndexFolders: 上限を超える本文インデックスのファイルがあるフォルダだけを返す" {
        ${contentIndexSearchMaxBytes} = 4KB
        $root = "$TestDrive\fit_find"
        newOversizeContentIndexFile "$root\営業"
        [System.IO.Directory]::CreateDirectory("$root\総務") | Out-Null
        writeContentIndexFile "$root\総務\content_index.xlsx.001.tsv" (convertToContentIndexText @(@{ Name = "A.xlsx"; Places = @(@{ Place = "S"; Text = "x`n" }) }))
        @(findOversizeContentIndexFolders $root) | Should -Be @("$root\営業")
    }

    It "assertContentIndexBookFits: 大きさの数え方は書き出す本文インデックスのファイルと同じ（上限ちょうどは通り、1 バイト少ないと落ちる）" {
        $dir = "$TestDrive\fit_exact"
        newTsv "$dir\$(toIndexFileName "S")" @("単価 1", "単価 2")
        $text = [System.IO.File]::ReadAllText("$dir\$(toIndexFileName "S")")
        $exact = measureContentIndexPartBytes @(@{ Block = (convertBookToContentIndexBlock @{ Name = "A.xlsx"; Places = @(@{ Place = "S"; Text = $text }) }) })
        { assertContentIndexBookFits $dir "A.xlsx" $exact } | Should -Not -Throw
        { assertContentIndexBookFits $dir "A.xlsx" ($exact - 2) } | Should -Throw "ファイルの文字が多すぎるため更新できません（上限 64 MB）。"
    }
}

Describe "publishIndexFolders: 書き終えたフォルダ（Published）" -Tag Io {
    BeforeEach {
        Mock writeIndexerLog {}
    }

    It "置き場の外を指すリンクのフォルダは飛ばし、Published に入れない" {
        $root = Join-Path $TestDrive "linked"
        $outside = Join-Path $TestDrive "linked_outside"
        newTsv "$root\index\A\新規.xlsx\$(toIndexFileName "S")" @("A の単価")
        newTsv "$outside\新規.xlsx\$(toIndexFileName "S")" @("外の単価")
        if (!(newTestJunction "$root\index\L" $outside)) {
            Set-ItResult -Skipped -Because "ジャンクションを作れない"
            return
        }
        $pending = New-Object 'System.Collections.Generic.Dictionary[string,object]'
        $pending["$root\index\A"] = @()
        $pending["$root\index\L"] = @()

        $result = publishTree $pending $root

        @($result.Published) | Should -Be @("$root\index\A")
    }

    It "書き出すフォルダが無ければ、Published は空" {
        $root = Join-Path $TestDrive "nothing"
        $pending = New-Object 'System.Collections.Generic.Dictionary[string,object]'

        $result = publishTree $pending $root

        @($result.Published).Count | Should -Be 0
    }
}
Describe "readContentIndexBookNamesStreaming" -Tag Io {
    BeforeAll {
        function script:writeUtf16 {
            param ([string]$path, [string]$text)
            [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
            [System.IO.File]::WriteAllText($path, $text, [System.Text.Encoding]::Unicode)
        }
        $script:rs = [string][char]0x1E
    }

    It "複数の元のファイルの名前を順に返す（版の行・場所の行・本文は名前にしない。% で符号化した値は戻す）" {
        $text = "$rs 版=2`n$rs ファイル名=A社.xlsx`n$rs シート=Sheet1`n本文`n$rs ファイル名=見積%25B.xlsx`r`n本文2`n"
        writeUtf16 "$TestDrive\names\ok.tsv" $text

        $names = readContentIndexBookNamesStreaming "$TestDrive\names\ok.tsv"

        ($names -join "|") | Should -Be "A社.xlsx|見積%B.xlsx"
    }

    It "版の行だけなら空の配列" {
        writeUtf16 "$TestDrive\names\only.tsv" "$rs 版=2`n"
        $names = readContentIndexBookNamesStreaming "$TestDrive\names\only.tsv"
        @($names).Count | Should -Be 0
    }

    It "行の長さの上限を超える行は飛ばし、その行の続きを次の名前と取り違えない（読み込みの区切りをまたぐ長い行も同じ）" {
        $long = "あ" * 100000
        $text = "$rs ファイル名=前.xlsx`n$rs ファイル名=$long`n$long`n$rs ファイル名=後.xlsx`n"
        writeUtf16 "$TestDrive\names\long.tsv" $text

        $names = readContentIndexBookNamesStreaming "$TestDrive\names\long.tsv" 1000

        ($names -join "|") | Should -Be "前.xlsx|後.xlsx"
    }

    It "最後の行に改行が無くても読む。読めないファイルは空の配列" {
        writeUtf16 "$TestDrive\names\last.tsv" "$rs ファイル名=最後.xlsx"
        ((readContentIndexBookNamesStreaming "$TestDrive\names\last.tsv") -join "|") | Should -Be "最後.xlsx"
        $none = readContentIndexBookNamesStreaming "$TestDrive\names\none.tsv"
        @($none).Count | Should -Be 0
    }
}
