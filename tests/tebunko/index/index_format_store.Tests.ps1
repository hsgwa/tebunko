# インデックスの格納形式の目録（store\format.tsv）の読み書きと、前の版の置き場の片付け（tebunko\index\index_format_store.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    function script:newFormatWorkspace {
        # テスト用のワークスペース（Dir は $TestDrive の下の名前）
        param ([string]$name)

        $dir = Join-Path $TestDrive $name
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        return [Workspace]::new($dir)
    }

    function script:newOldIndex {
        # 古い置き場（content_index\<名前>）に、本文インデックスのファイルを 1 つ置く
        param ($ws, [string]$name)

        $dir = Join-Path $ws.OldIndexDir $name
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        [System.IO.File]::WriteAllText("$dir\content_index.docx.001.tsv", "古い中身")
        [System.IO.File]::WriteAllText("$dir\source_folder.txt", "C:\data")
    }
}

Describe "readIndexFormat / writeIndexFormat" -Tag Io {
    It "書いた目録を、版・アプリの版・インデックスの行のとおりに読み戻す" {
        $ws = newFormatWorkspace "format_roundtrip"
        writeIndexFormat ([ordered]@{ "営業" = "作り終えた"; "技術" = "作成中" }) $ws.FormatFile "0.5.0"

        $catalog = readIndexFormat $ws.FormatFile
        $catalog.Status | Should -Be "Ok"
        $catalog.Version | Should -Be ${indexFormatVersion}
        $catalog.AppVersion | Should -Be "0.5.0"
        $catalog.Indexes["営業"] | Should -Be "作り終えた"
        $catalog.Indexes["技術"] | Should -Be "作成中"
        $bytes = [System.IO.File]::ReadAllBytes($ws.FormatFile)
        $bytes[0..2] | Should -Be @(0xEF, 0xBB, 0xBF)
    }

    It "知らない行・知らない列は読み飛ばす" {
        $ws = newFormatWorkspace "format_unknown"
        [System.IO.Directory]::CreateDirectory($ws.StoreDir) | Out-Null
        $text = "版=${indexFormatVersion}`r`nアプリの版=0.5.0`r`n将来のキー=値`r`n営業`t作り終えた`t将来の列`r`n`r`n"
        [System.IO.File]::WriteAllText($ws.FormatFile, $text, (New-Object System.Text.UTF8Encoding($true)))

        $catalog = readIndexFormat $ws.FormatFile
        $catalog.Status | Should -Be "Ok"
        $catalog.Indexes["営業"] | Should -Be "作り終えた"
        $catalog.Indexes.Count | Should -Be 1
    }

    It "<name>: 状態は <expected>" -TestCases @(
        @{ name = "ファイルが無い"; text = $null; expected = "Missing" }
        @{ name = "版の行が無い"; text = "営業`t作り終えた`r`n"; expected = "Broken" }
        @{ name = "版が数字でない"; text = "版=あ`r`n"; expected = "Broken" }
        @{ name = "空のファイル"; text = ""; expected = "Broken" }
        @{ name = "今より新しい版"; text = "版=99`r`n営業`t作り終えた`r`n"; expected = "Newer" }
        @{ name = "今より古い版"; text = "版=1`r`n"; expected = "Older" }
    ) {
        param ($name, $text, $expected)
        $ws = newFormatWorkspace "format_status_$($name.GetHashCode())"
        if ($null -ne $text) {
            [System.IO.Directory]::CreateDirectory($ws.StoreDir) | Out-Null
            [System.IO.File]::WriteAllText($ws.FormatFile, $text, (New-Object System.Text.UTF8Encoding($true)))
        }
        (readIndexFormat $ws.FormatFile).Status | Should -Be $expected
    }
}

Describe "updateIndexFormat と行の操作" -Tag Io {
    It "行を足し、状態を替え、改名し、消す" {
        $ws = newFormatWorkspace "format_rows"
        setIndexFormatRow "営業" "作成中" $ws.FormatFile | Should -Be $true
        setIndexFormatRow "技術" "作成中" $ws.FormatFile | Should -Be $true
        setIndexFormatRow "営業" "作り終えた" $ws.FormatFile | Should -Be $true
        renameIndexFormatRow "技術" "開発" $ws.FormatFile | Should -Be $true
        removeIndexFormatRow "営業" $ws.FormatFile | Should -Be $true

        $catalog = readIndexFormat $ws.FormatFile
        @($catalog.Indexes.Keys) | Should -Be @("開発")
        $catalog.Indexes["開発"] | Should -Be "作成中"
        Test-Path -LiteralPath "$($ws.FormatFile).lock" | Should -Be $false
    }

    It "目録が無いときの改名・削除は何もしない（目録を作らない）" {
        $ws = newFormatWorkspace "format_rows_none"
        renameIndexFormatRow "営業" "営2" $ws.FormatFile | Should -Be $true
        removeIndexFormatRow "営業" $ws.FormatFile | Should -Be $true
        Test-Path -LiteralPath $ws.FormatFile | Should -Be $false
    }

    It "今より新しい版の目録は書き換えない" {
        $ws = newFormatWorkspace "format_newer"
        [System.IO.Directory]::CreateDirectory($ws.StoreDir) | Out-Null
        $text = "版=99`r`n営業`t作り終えた`r`n"
        [System.IO.File]::WriteAllText($ws.FormatFile, $text, (New-Object System.Text.UTF8Encoding($true)))

        setIndexFormatRow "技術" "作成中" $ws.FormatFile | Should -Be $false
        [System.IO.File]::ReadAllText($ws.FormatFile) | Should -Be $text
    }

    It "markIndexFormatBuilding は、行の無いものだけ「作成中」で足し、あるものは変えない" {
        $ws = newFormatWorkspace "format_mark"
        setIndexFormatRow "営業" "作り終えた" $ws.FormatFile | Should -Be $true

        markIndexFormatBuilding @("営業", "技術") $ws | Should -Be $true

        $catalog = readIndexFormat $ws.FormatFile
        $catalog.Indexes["営業"] | Should -Be "作り終えた"
        $catalog.Indexes["技術"] | Should -Be "作成中"
    }
}

Describe "repairIndexFormat / getWorkspaceIndexFormat" -Tag Io {
    It "壊れた目録は、今の版の本文インデックスがあるインデックスの行を「作成中」にして書き直す" {
        $ws = newFormatWorkspace "format_repair"
        $good = Join-Path $ws.IndexDir "営業"
        $bad = Join-Path $ws.IndexDir "旧版"
        [System.IO.Directory]::CreateDirectory($good) | Out-Null
        [System.IO.Directory]::CreateDirectory($bad) | Out-Null
        $mark = [string][char]0x1E
        [System.IO.File]::WriteAllText("$good\content_index.docx.001.tsv", "$mark 版=${indexFormatVersion}`n", [System.Text.Encoding]::Unicode)
        [System.IO.File]::WriteAllText("$bad\content_index.docx.001.tsv", "$mark 版=1`n", [System.Text.Encoding]::Unicode)
        [System.IO.File]::WriteAllText($ws.FormatFile, "こわれた", (New-Object System.Text.UTF8Encoding($true)))

        $state = getWorkspaceIndexFormat $ws

        $state.State | Should -Be "Broken"
        $catalog = readIndexFormat $ws.FormatFile
        $catalog.Status | Should -Be "Ok"
        @($catalog.Indexes.Keys) | Should -Be @("営業")
        $catalog.Indexes["営業"] | Should -Be "作成中"
    }

    It "古い置き場にだけ本文インデックスがあれば、そのインデックスは Old で、目録は書かない" {
        $ws = newFormatWorkspace "format_old_state"
        newOldIndex $ws "営業"

        $state = getWorkspaceIndexFormat $ws

        $state.State | Should -Be "Old"
        $state.Indexes["営業"].State | Should -Be "Old"
        $state.Indexes["営業"].Searchable | Should -Be $false
        Test-Path -LiteralPath $ws.FormatFile | Should -Be $false
    }

    It "findOldIndexNames は、本文インデックスの無いフォルダ（source_folder.txt だけ）を数えない" {
        $ws = newFormatWorkspace "format_old_names"
        newOldIndex $ws "営業"
        [System.IO.Directory]::CreateDirectory((Join-Path $ws.OldIndexDir "空")) | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $ws.OldIndexDir "空\source_folder.txt"), "C:\data")

        @(findOldIndexNames $ws) | Should -Be @("営業")
    }
}

Describe "renameOldIndexFolder" -Tag Io {
    It "古い置き場の同じ名前のフォルダの名前も変える。大文字小文字だけの改名もできる" {
        $ws = newFormatWorkspace "format_old_rename"
        newOldIndex $ws "abc"

        renameOldIndexFolder "abc" "ABC" $ws
        @([System.IO.Directory]::GetDirectories($ws.OldIndexDir) | ForEach-Object { [System.IO.Path]::GetFileName($_) }) | Should -BeExactly @("ABC")

        renameOldIndexFolder "ABC" "営業" $ws
        Test-Path -LiteralPath (Join-Path $ws.OldIndexDir "営業\content_index.docx.001.tsv") | Should -Be $true
    }

    It "新しい名前のフォルダが既にあれば動かさず、無いものは何もしない" {
        $ws = newFormatWorkspace "format_old_rename_keep"
        newOldIndex $ws "営業"
        newOldIndex $ws "技術"

        renameOldIndexFolder "営業" "技術" $ws
        renameOldIndexFolder "無い" "新" $ws

        Test-Path -LiteralPath (Join-Path $ws.OldIndexDir "営業") | Should -Be $true
        Test-Path -LiteralPath (Join-Path $ws.OldIndexDir "新") | Should -Be $false
    }
}

Describe "removeOldIndexOf / completeIndexFormat" -Tag Io {
    It "作り終えたインデックスの古い分だけを消す。ほかのインデックスが古い置き場に残れば、置き場そのものは残す" {
        $ws = newFormatWorkspace "format_remove_some"
        newOldIndex $ws "営業"
        newOldIndex $ws "技術"
        [System.IO.File]::WriteAllText($ws.OldStatusFile, "古い取り込み一覧")

        $result = removeOldIndexOf "営業" $ws

        $result.Failed.Count | Should -Be 0
        $result.Removed | Should -Be 1
        Test-Path -LiteralPath (Join-Path $ws.OldIndexDir "営業") | Should -Be $false
        Test-Path -LiteralPath (Join-Path $ws.OldIndexDir "技術\content_index.docx.001.tsv") | Should -Be $true
        Test-Path -LiteralPath $ws.OldStatusFile | Should -Be $true
    }

    It "古い置き場に本文インデックスが無くなれば、配列にある古い置き場（content_index\ と ingest_status.tsv）だけを消す" {
        $ws = newFormatWorkspace "format_remove_all"
        newOldIndex $ws "営業"
        [System.IO.File]::WriteAllText($ws.OldStatusFile, "古い取り込み一覧")
        $other = Join-Path $ws.Dir "tmp\残す.txt"
        [System.IO.Directory]::CreateDirectory((Split-Path $other)) | Out-Null
        [System.IO.File]::WriteAllText($other, "別のもの")
        setIndexFormatRow "営業" "作り終えた" $ws.FormatFile | Should -Be $true

        $result = removeOldIndexOf "営業" $ws

        $result.Failed.Count | Should -Be 0
        Test-Path -LiteralPath $ws.OldIndexDir | Should -Be $false
        Test-Path -LiteralPath $ws.OldStatusFile | Should -Be $false
        # 配列に無いものは触らない
        Test-Path -LiteralPath $other | Should -Be $true
        Test-Path -LiteralPath $ws.FormatFile | Should -Be $true
    }

    It "ジャンクションの先を消さない（リンクだけを外す）" {
        $ws = newFormatWorkspace "format_remove_link"
        [System.IO.Directory]::CreateDirectory($ws.OldIndexDir) | Out-Null
        $outside = Join-Path $TestDrive "format_remove_link_outside"
        [System.IO.Directory]::CreateDirectory($outside) | Out-Null
        [System.IO.File]::WriteAllText("$outside\大事.txt", "消してはいけない")
        if (!(newTestJunction (Join-Path $ws.OldIndexDir "営業") $outside)) {
            Set-ItResult -Skipped -Because "ジャンクションを作れません: $($script:linkError)"
            return
        }

        removeOldIndexOf "営業" $ws | Out-Null

        Test-Path -LiteralPath "$outside\大事.txt" | Should -Be $true
        Test-Path -LiteralPath (Join-Path $ws.OldIndexDir "営業") | Should -Be $false
    }

    It "置き場の外を指す名前（<name>）では何も消さない" -TestCases @(
        @{ name = ".." }
        @{ name = "." }
        @{ name = "..\外" }
        @{ name = "" }
    ) {
        param ($name)
        $ws = newFormatWorkspace "format_remove_bad_name"
        newOldIndex $ws "営業"
        $outside = Join-Path $ws.Dir "外"
        [System.IO.Directory]::CreateDirectory($outside) | Out-Null
        [System.IO.File]::WriteAllText("$outside\大事.txt", "消してはいけない")

        $result = removeOldIndexOf $name $ws

        $result.Removed | Should -Be 0
        $result.Failed.Count | Should -Be 1
        Test-Path -LiteralPath "$outside\大事.txt" | Should -Be $true
        Test-Path -LiteralPath (Join-Path $ws.OldIndexDir "営業\content_index.docx.001.tsv") | Should -Be $true
    }

    It "消せなければ残して結果で返し、例外にしない" {
        $ws = newFormatWorkspace "format_remove_locked"
        newOldIndex $ws "営業"
        $locked = Join-Path $ws.OldIndexDir "営業\content_index.docx.001.tsv"
        $stream = [System.IO.File]::Open($locked, "Open", "Read", "None")
        try {
            $result = removeOldIndexOf "営業" $ws
        } finally {
            $stream.Dispose()
        }

        $result.Failed.Count | Should -BeGreaterThan 0
        Test-Path -LiteralPath $locked | Should -Be $true
    }

    It "completeIndexFormat は、目録の行を「作り終えた」にしてから古い分を消す" {
        $ws = newFormatWorkspace "format_complete"
        newOldIndex $ws "営業"
        setIndexFormatRow "営業" "作成中" $ws.FormatFile | Should -Be $true

        $result = completeIndexFormat @("営業") $ws

        $result.Updated | Should -Be $true
        $result.Failed.Count | Should -Be 0
        (readIndexFormat $ws.FormatFile).Indexes["営業"] | Should -Be "作り終えた"
        Test-Path -LiteralPath $ws.OldIndexDir | Should -Be $false
    }

    It "目録を書けない（今より新しい版）ときは、古い分を消さない" {
        $ws = newFormatWorkspace "format_complete_newer"
        newOldIndex $ws "営業"
        [System.IO.Directory]::CreateDirectory($ws.StoreDir) | Out-Null
        [System.IO.File]::WriteAllText($ws.FormatFile, "版=99`r`n", (New-Object System.Text.UTF8Encoding($true)))

        $result = completeIndexFormat @("営業") $ws

        $result.Updated | Should -Be $false
        Test-Path -LiteralPath (Join-Path $ws.OldIndexDir "営業\content_index.docx.001.tsv") | Should -Be $true
    }

    It "消す古い置き場は、配列 oldIndexEntries の 2 つだけ" {
        ${oldIndexEntries} | Should -Be @("OldIndexDir", "OldStatusFile")
    }
}
