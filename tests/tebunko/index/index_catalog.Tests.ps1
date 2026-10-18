# インデックスの改名・削除・名前の割り当てと保存（tebunko\index\index_catalog.ps1 の IndexCatalog）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    function script:newCatalogWorkspace {
        # テスト用のワークスペース（dir）と、その IndexCatalog を作る。設定ファイルは dir\setting.config
        param ([string]$dir)

        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        return @{
            Dir      = $dir
            Index    = "$dir\store\content"
            Status   = "$dir\store\ingest_status.tsv"
            Settings = "$dir\setting.config"
            Catalog  = (newIndexCatalog $dir "$dir\setting.config")
        }
    }
}

Describe "IndexCatalog.RenameIndex" -Tag Io {
    It "インデックスのフォルダと取り込み一覧の記録の名前を変え、中身はそのまま残す" {
        $ws = newCatalogWorkspace "$TestDrive\rename"
        $dir = $ws.Index
        $path = $ws.Status
        New-Item -ItemType Directory -Path "$dir\営業\a.xlsx" -Force | Out-Null
        Set-Content -LiteralPath "$dir\営業\a.xlsx\Sheet1.tsv" -Value "本文" -Encoding UTF8
        New-Item -ItemType Directory -Path "$dir\技術" -Force | Out-Null
        writeStatusFile @(
            [pscustomobject]@{ Path = "C:\data"; Name = "営業" },
            [pscustomobject]@{ Path = "D:\tech"; Name = "技術" }
        ) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00"),
            (newStatusRow "技術\d.pptx" "2025/01/10 12:34:56" "1" $stateNew)
        ) $path

        $ws.Catalog.RenameIndex("営業", "営業部")

        Test-Path "$dir\営業" | Should -Be $false
        Get-Content -LiteralPath "$dir\営業部\a.xlsx\Sheet1.tsv" | Should -Be "本文"

        $status = readStatusFile $path
        $status.Folders[0].Name | Should -Be "営業部"
        $status.Folders[0].Path | Should -Be "C:\data"
        $status.Rows.ContainsKey("営業部\a.xlsx") | Should -Be $true
        $status.Rows["営業部\a.xlsx"].状態 | Should -Be $stateDone
        # ほかのインデックスはそのまま
        $status.Folders[1].Name | Should -Be "技術"
        $status.Rows.ContainsKey("技術\d.pptx") | Should -Be $true
    }

    It "インデックスのフォルダがまだ無くても、取り込み一覧の記録は変える" {
        $ws = newCatalogWorkspace "$TestDrive\rename2"
        New-Item -ItemType Directory -Path $ws.Index -Force | Out-Null
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "営業" }) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateNew)
        ) $ws.Status

        $ws.Catalog.RenameIndex("営業", "営業部")

        (readStatusFile $ws.Status).Rows.ContainsKey("営業部\a.xlsx") | Should -Be $true
    }

    It "予約名に拡張子の付いた名前（<name>）のインデックスも、フォルダを動かしてから取り込み一覧の記録を変える" -TestCases @(
        @{ name = "aux.backup" }
        @{ name = "CON.txt" }
    ) {
        $ws = newCatalogWorkspace "$TestDrive\rename_reserved_$name"
        $long = toLongPath "$($ws.Index)\$name\a.xlsx"
        [System.IO.Directory]::CreateDirectory($long) | Out-Null
        [System.IO.File]::WriteAllText("$long\Sheet1.tsv", "本文")
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = $name }) @(
            (newStatusRow "$name\a.xlsx" "2025/01/10 12:34:56" "1" $stateNew)
        ) $ws.Status

        $ws.Catalog.RenameIndex($name, "営業部")

        [System.IO.Directory]::Exists((toLongPath "$($ws.Index)\$name")) | Should -Be $false
        [System.IO.File]::ReadAllText((toLongPath "$($ws.Index)\営業部\a.xlsx\Sheet1.tsv")) | Should -Be "本文"
        (readStatusFile $ws.Status).Rows.ContainsKey("営業部\a.xlsx") | Should -Be $true
    }

    It "同じ名前のフォルダが既にあれば例外にし、取り込み一覧の記録は変えない" {
        $ws = newCatalogWorkspace "$TestDrive\rename3"
        New-Item -ItemType Directory -Path "$($ws.Index)\営業" -Force | Out-Null
        New-Item -ItemType Directory -Path "$($ws.Index)\技術" -Force | Out-Null
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "営業" }) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateNew)
        ) $ws.Status

        { $ws.Catalog.RenameIndex("営業", "技術") } | Should -Throw

        Test-Path -LiteralPath "$($ws.Index)\営業" | Should -Be $true
        (readStatusFile $ws.Status).Rows.ContainsKey("営業\a.xlsx") | Should -Be $true
    }

    It "新しい名前がファイルとして既にあっても例外にし、フォルダも取り込み一覧の記録も変えない" {
        $ws = newCatalogWorkspace "$TestDrive\rename3file"
        New-Item -ItemType Directory -Path "$($ws.Index)\営業" -Force | Out-Null
        [System.IO.File]::WriteAllText("$($ws.Index)\技術", "x")
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "営業" }) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateNew)
        ) $ws.Status

        { $ws.Catalog.RenameIndex("営業", "技術") } | Should -Throw "*既にあるため*"

        [System.IO.Directory]::Exists("$($ws.Index)\営業") | Should -Be $true
        [System.IO.File]::Exists("$($ws.Index)\技術") | Should -Be $true
        (readStatusFile $ws.Status).Rows.ContainsKey("営業\a.xlsx") | Should -Be $true
    }

    It "大文字・小文字だけを変えられる（フォルダ・記録とも新しい書き方になる）" {
        $ws = newCatalogWorkspace "$TestDrive\rename4"
        $dir = $ws.Index
        writeListFile "$dir\sales\a.xlsx\Sheet1.tsv" @("本文")
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "sales" }) @(
            (newStatusRow "sales\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00")
        ) $ws.Status

        $ws.Catalog.RenameIndex("sales", "Sales")

        @(Get-ChildItem -LiteralPath $dir -Directory | ForEach-Object { $_.Name }) -join "," | Should -BeExactly "Sales"
        Get-Content -LiteralPath "$dir\Sales\a.xlsx\Sheet1.tsv" -Encoding UTF8 | Should -Be "本文"
        $status = readStatusFile $ws.Status
        $status.Folders[0].Name | Should -BeExactly "Sales"
        $status.Rows["Sales\a.xlsx"].相対パス | Should -BeExactly "Sales\a.xlsx"
    }

    It "名前が空・同じなら何もしない" {
        $ws = newCatalogWorkspace "$TestDrive\rename5"
        New-Item -ItemType Directory -Path "$($ws.Index)\営業" -Force | Out-Null
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "営業" }) @() $ws.Status

        $ws.Catalog.RenameIndex("", "技術")
        $ws.Catalog.RenameIndex("営業", "")
        $ws.Catalog.RenameIndex("営業", "営業")

        Test-Path -LiteralPath "$($ws.Index)\営業" | Should -Be $true
        (readStatusFile $ws.Status).Folders[0].Name | Should -Be "営業"
    }
}

Describe "IndexCatalog.RenameIndex（使えない名前）" -Tag Io {
    It "移す元・先が置き場の外へ出る・使えない名前なら、何も動かさずに例外にする: <from> -> <to>" -TestCases @(
        @{ from = "営業"; to = "..\外" }
        @{ from = "営業"; to = "..\..\外\営業" }
        @{ from = "営業"; to = ".." }
        @{ from = "営業"; to = "経理\中" }
        @{ from = "営業"; to = "営業." }
        @{ from = "営業"; to = "CON" }
        @{ from = "営業"; to = " 経理" }
        @{ from = "..\外"; to = "経理" }
        @{ from = "営業\..\経理"; to = "新" }
        @{ from = ".."; to = "新" }
        @{ from = "営業."; to = "新" }
        @{ from = "営業 "; to = "新" }
    ) {
        param ($from, $to)
        $ws = newCatalogWorkspace "$TestDrive\rename_bad"
        $dir = $ws.Index
        writeListFile "$dir\営業\a.xlsx\Sheet1.tsv" @("営業の本文")
        writeListFile "$dir\経理\b.xlsx\Sheet1.tsv" @("経理の本文")
        writeListFile "$($ws.Dir)\外\c.tsv" @("外の本文")
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "営業" }) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00")
        ) $ws.Status

        { $ws.Catalog.RenameIndex($from, $to) } | Should -Throw

        Test-Path -LiteralPath "$dir\営業\a.xlsx\Sheet1.tsv" | Should -Be $true
        Test-Path -LiteralPath "$dir\経理\b.xlsx\Sheet1.tsv" | Should -Be $true
        Test-Path -LiteralPath "$($ws.Dir)\外\c.tsv" | Should -Be $true
        @(Get-ChildItem -LiteralPath $dir -Directory | ForEach-Object { $_.Name } | Sort-Object) -join "," | Should -Be "営業,経理"
        @(Get-ChildItem -LiteralPath $ws.Dir -Directory | ForEach-Object { $_.Name } | Sort-Object) -join "," | Should -Be "store,外"
        $status = readStatusFile $ws.Status
        $status.Folders[0].Name | Should -Be "営業"
        $status.Rows.ContainsKey("営業\a.xlsx") | Should -Be $true
    }
}

Describe "IndexCatalog（既にある予約名のインデックス）" -Tag Io {
    # aux.backup のように、前の版で付いた予約名のフォルダ。Test-Path では無いように見えるため、\\?\ で作り、Directory.Exists で確かめる
    It "RenameIndex は、予約名（<name>）の名前から新しい名前へ改名できる" -TestCases @(
        @{ name = "aux.backup" }
        @{ name = "CON.txt" }
    ) {
        param ($name)
        $ws = newCatalogWorkspace "$TestDrive\reserved_rename_$($name.Substring(0, 3))"
        $old = Join-Path $ws.Index $name
        [System.IO.Directory]::CreateDirectory((toLongPath (Join-Path $old "a.xlsx"))) | Out-Null
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = $name }) @(
            (newStatusRow "$name\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00")
        ) $ws.Status

        $ws.Catalog.RenameIndex($name, "新")

        [System.IO.Directory]::Exists((toLongPath $old)) | Should -Be $false
        Test-Path -LiteralPath "$($ws.Index)\新\a.xlsx" | Should -Be $true
        $status = readStatusFile $ws.Status
        $status.Folders[0].Name | Should -Be "新"
        $status.Rows.ContainsKey("新\a.xlsx") | Should -Be $true
    }

    It "RemoveIndex は、予約名（<name>）のフォルダを消し、取り込み一覧の記録も消す" -TestCases @(
        @{ name = "aux.backup" }
        @{ name = "CON.txt" }
    ) {
        param ($name)
        $ws = newCatalogWorkspace "$TestDrive\reserved_remove_$($name.Substring(0, 3))"
        $target = Join-Path $ws.Index $name
        [System.IO.Directory]::CreateDirectory((toLongPath (Join-Path $target "a.xlsx"))) | Out-Null
        New-Item -ItemType Directory -Path "$($ws.Index)\技術" -Force | Out-Null
        writeStatusFile @(
            [pscustomobject]@{ Path = "C:\data"; Name = $name },
            [pscustomobject]@{ Path = "D:\tech"; Name = "技術" }
        ) @(
            (newStatusRow "$name\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00"),
            (newStatusRow "技術\d.pptx" "2025/01/10 12:34:56" "1" $stateNew)
        ) $ws.Status

        $ws.Catalog.RemoveIndex($name)

        [System.IO.Directory]::Exists((toLongPath $target)) | Should -Be $false
        Test-Path -LiteralPath "$($ws.Index)\技術" | Should -Be $true
        $status = readStatusFile $ws.Status
        $status.Folders.Count | Should -Be 1
        $status.Folders[0].Name | Should -Be "技術"
    }

    It "新しい名前には予約名を使えない（RenameIndex は何も動かさずに例外にする）" {
        $ws = newCatalogWorkspace "$TestDrive\reserved_new"
        New-Item -ItemType Directory -Path "$($ws.Index)\営業" -Force | Out-Null
        { $ws.Catalog.RenameIndex("営業", "aux.backup") } | Should -Throw
        Test-Path -LiteralPath "$($ws.Index)\営業" | Should -Be $true
        [System.IO.Directory]::Exists((toLongPath (Join-Path $ws.Index "aux.backup"))) | Should -Be $false
    }
}

Describe "IndexCatalog.RemoveIndex" -Tag Io {
    It "インデックスのフォルダと取り込み一覧の記録を削除し、ほかのインデックスは残す" {
        $ws = newCatalogWorkspace "$TestDrive\remove"
        $dir = $ws.Index
        New-Item -ItemType Directory -Path "$dir\営業\a.xlsx" -Force | Out-Null
        Set-Content -LiteralPath "$dir\営業\a.xlsx\Sheet1.tsv" -Value "本文" -Encoding UTF8
        New-Item -ItemType Directory -Path "$dir\技術" -Force | Out-Null
        writeStatusFile @(
            [pscustomobject]@{ Path = "C:\data"; Name = "営業" },
            [pscustomobject]@{ Path = "D:\tech"; Name = "技術" }
        ) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00"),
            (newStatusRow "技術\d.pptx" "2025/01/10 12:34:56" "1" $stateNew)
        ) $ws.Status

        $ws.Catalog.RemoveIndex("営業")

        Test-Path "$dir\営業" | Should -Be $false
        Test-Path "$dir\技術" | Should -Be $true

        $status = readStatusFile $ws.Status
        $status.Folders.Count | Should -Be 1
        $status.Folders[0].Name | Should -Be "技術"
        $status.Rows.Count | Should -Be 1
        $status.Rows.ContainsKey("技術\d.pptx") | Should -Be $true
    }

    It "予約名に拡張子の付いた名前（<name>）のインデックスも、フォルダと取り込み一覧の記録を削除する" -TestCases @(
        @{ name = "aux.backup" }
        @{ name = "CON.txt" }
    ) {
        $ws = newCatalogWorkspace "$TestDrive\remove_reserved_$name"
        $long = toLongPath "$($ws.Index)\$name\a.xlsx"
        [System.IO.Directory]::CreateDirectory($long) | Out-Null
        [System.IO.File]::WriteAllText("$long\Sheet1.tsv", "本文")
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = $name }) @(
            (newStatusRow "$name\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00")
        ) $ws.Status

        $ws.Catalog.RemoveIndex($name)

        [System.IO.Directory]::Exists((toLongPath "$($ws.Index)\$name")) | Should -Be $false
        (readStatusFile $ws.Status).Folders.Count | Should -Be 0
    }

    It "名前が空なら何もしない" {
        $ws = newCatalogWorkspace "$TestDrive\remove2"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "営業" }) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateNew)
        ) $ws.Status

        $ws.Catalog.RemoveIndex("")

        (readStatusFile $ws.Status).Rows.Count | Should -Be 1
    }

    It "置き場の外を指す名前は、外のフォルダを消さずに例外にし、取り込み一覧の記録は残す" -TestCases @(
        @{ badName = "..\外" }
        @{ badName = "営業\..\..\外" }
        @{ badName = ".." }
        @{ badName = "営業\..\経理" }   # 置き場の中を指すが、別のインデックスのフォルダになる
        @{ badName = "営業\中" }
        @{ badName = "営業/中" }
        @{ badName = "営業." }   # 最後の . は、有無の確かめでは「営業」を見るため、決まりに合わない名前は消さない
        @{ badName = "営業 " }
    ) {
        param ($badName)
        $ws = newCatalogWorkspace "$TestDrive\remove4"
        New-Item -ItemType Directory -Path "$($ws.Index)\営業\中" -Force | Out-Null
        New-Item -ItemType Directory -Path "$($ws.Index)\経理" -Force | Out-Null
        writeListFile "$($ws.Dir)\外\a.tsv" @("本文")
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "営業" }) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateNew)
        ) $ws.Status

        { $ws.Catalog.RemoveIndex($badName) } | Should -Throw

        Test-Path -LiteralPath "$($ws.Dir)\外\a.tsv" | Should -Be $true
        Test-Path -LiteralPath "$($ws.Index)\営業\中" | Should -Be $true
        Test-Path -LiteralPath "$($ws.Index)\経理" | Should -Be $true
        Test-Path -LiteralPath $ws.Dir | Should -Be $true
        (readStatusFile $ws.Status).Rows.Count | Should -Be 1
    }

    It "フォルダを消せなければ（TSV を開いている等）例外にし、取り込み一覧の記録は残す" {
        # 画面は別スレッド（$ErrorActionPreference が既定の Continue）で呼ぶため、消せなかったことを例外で知らせる必要がある
        $ws = newCatalogWorkspace "$TestDrive\remove3"
        $dir = $ws.Index
        writeListFile "$dir\営業\a.xlsx\Sheet1.tsv" @("本文")
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "営業" }) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00")
        ) $ws.Status

        $stream = [System.IO.File]::Open("$dir\営業\a.xlsx\Sheet1.tsv", "Open", "Read", "None")
        try {
            & {
                $ErrorActionPreference = "Continue"
                { $ws.Catalog.RemoveIndex("営業") } | Should -Throw
            }
        } finally {
            $stream.Dispose()
        }

        $status = readStatusFile $ws.Status
        $status.Folders.Count | Should -Be 1
        $status.Rows.ContainsKey("営業\a.xlsx") | Should -Be $true
    }

    It "フォルダがあるかどうか確かめられなければ（届かない・読めない）、例外にし、取り込み一覧の記録は残す" {
        $ws = newCatalogWorkspace "$TestDrive\remove_unknown"
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "営業" }) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00")
        ) $ws.Status
        Mock getPathState { @{ State = "Other"; IsDirectory = $false; Message = "アクセスが拒否されました" } }

        { $ws.Catalog.RemoveIndex("営業") } | Should -Throw "*フォルダがあるかどうか確かめられません*"

        (readStatusFile $ws.Status).Rows.ContainsKey("営業\a.xlsx") | Should -Be $true
    }

    It "インデックスのフォルダの中にあるジャンクションの先は消さず、ジャンクションだけを外す" {
        $ws = newCatalogWorkspace "$TestDrive\remove_junction"
        $outside = "$TestDrive\remove_junction_outside"
        writeListFile "$outside\外.tsv" @("外")
        writeListFile "$($ws.Index)\営業\a.xlsx\Sheet1.tsv" @("本文")
        if (!(newTestJunction "$($ws.Index)\営業\リンク" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "営業" }) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00")
        ) $ws.Status

        $ws.Catalog.RemoveIndex("営業")

        Test-Path -LiteralPath "$($ws.Index)\営業" | Should -Be $false
        readListFile "$outside\外.tsv" | Should -Be @("外")
        (readStatusFile $ws.Status).Rows.Count | Should -Be 0
    }

    It "インデックスのフォルダそのものがジャンクションなら、ジャンクションだけを外し、先は消さない" {
        $ws = newCatalogWorkspace "$TestDrive\remove_rootlink"
        $outside = "$TestDrive\remove_rootlink_outside"
        writeListFile "$outside\外.tsv" @("外")
        [System.IO.Directory]::CreateDirectory($ws.Index) | Out-Null
        if (!(newTestJunction "$($ws.Index)\営業" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "営業" }) @() $ws.Status

        $ws.Catalog.RemoveIndex("営業")

        Test-Path -LiteralPath "$($ws.Index)\営業" | Should -Be $false
        readListFile "$outside\外.tsv" | Should -Be @("外")
    }

    It "古い置き場（content_index）そのものがジャンクションでも、削除でき、取り込み一覧の記録も消える" {
        $ws = newCatalogWorkspace "$TestDrive\remove_oldrootlink"
        $outside = "$TestDrive\remove_oldrootlink_outside"
        writeListFile "$outside\営業\content_index.docx.001.tsv" @("古い中身")
        writeListFile "$outside\技術\content_index.docx.001.tsv" @("古い中身")
        writeListFile "$($ws.Index)\営業\a.xlsx\Sheet1.tsv" @("本文")
        if (!(newTestJunction "$($ws.Dir)\content_index" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "営業" }) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00")
        ) $ws.Status

        $ws.Catalog.RemoveIndex("営業")

        Test-Path -LiteralPath "$($ws.Index)\営業" | Should -Be $false
        Test-Path -LiteralPath "$outside\営業" | Should -Be $false
        Test-Path -LiteralPath "$outside\技術" | Should -Be $true
        (readStatusFile $ws.Status).Rows.Count | Should -Be 0
    }
}

Describe "IndexCatalog.RenameIndex（リンク）" -Tag Io {
    It "インデックスのフォルダがジャンクションなら、ジャンクションの名前だけを変え、先は動かさない" {
        # 名前は 1 階層なので、途中にリンクが挟まる形は作れない。リンクになりうるのは名前のフォルダそのものだけ
        $ws = newCatalogWorkspace "$TestDrive\rename_link"
        $outside = "$TestDrive\rename_link_outside"
        writeListFile "$outside\外.tsv" @("外")
        [System.IO.Directory]::CreateDirectory($ws.Index) | Out-Null
        if (!(newTestJunction "$($ws.Index)\営業" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        $ws.Catalog.RenameIndex("営業", "営業部")

        Test-Path -LiteralPath "$($ws.Index)\営業" | Should -Be $false
        Test-Path -LiteralPath "$($ws.Index)\営業部" | Should -Be $true
        readListFile "$outside\外.tsv" | Should -Be @("外")
    }

    It "古い置き場（content_index）そのものがジャンクションでも、改名でき、古い置き場の同じ名前も変わる" {
        $ws = newCatalogWorkspace "$TestDrive\rename_oldrootlink"
        $outside = "$TestDrive\rename_oldrootlink_outside"
        writeListFile "$outside\営業\content_index.docx.001.tsv" @("古い中身")
        writeListFile "$($ws.Index)\営業\a.xlsx\Sheet1.tsv" @("本文")
        if (!(newTestJunction "$($ws.Dir)\content_index" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        $ws.Catalog.RenameIndex("営業", "営業部")

        Test-Path -LiteralPath "$($ws.Index)\営業部" | Should -Be $true
        Test-Path -LiteralPath "$outside\営業部\content_index.docx.001.tsv" | Should -Be $true
        Test-Path -LiteralPath "$outside\営業" | Should -Be $false
    }
}

Describe "IndexCatalog.RemoveIndexes" -Tag Io {
    It "選んだインデックスをまとめて削除し、1 つ失敗しても残りを続けて、名前ごとの結果を返す" {
        $ws = newCatalogWorkspace "$TestDrive\removes"
        $dir = $ws.Index
        writeListFile "$dir\営業\a.xlsx\Sheet1.tsv" @("本文")
        writeListFile "$dir\技術\b.xlsx\Sheet1.tsv" @("本文")
        writeListFile "$dir\総務\c.xlsx\Sheet1.tsv" @("本文")
        writeStatusFile @(
            [pscustomobject]@{ Path = "C:\data"; Name = "営業" },
            [pscustomobject]@{ Path = "D:\tech"; Name = "技術" },
            [pscustomobject]@{ Path = "E:\gen"; Name = "総務" }
        ) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00"),
            (newStatusRow "技術\b.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00"),
            (newStatusRow "総務\c.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00")
        ) $ws.Status

        # 営業の TSV を開いたままにして消せなくする
        $stream = [System.IO.File]::Open("$dir\営業\a.xlsx\Sheet1.tsv", "Open", "Read", "None")
        try {
            $results = & {
                $ErrorActionPreference = "Continue"
                $ws.Catalog.RemoveIndexes([string[]]@("営業", "技術", ""))
            }
        } finally {
            $stream.Dispose()
        }

        @($results).Count | Should -Be 3
        $results[0].Name | Should -Be "営業"
        $results[0].Ok | Should -Be $false
        $results[0].Reason | Should -Not -BeNullOrEmpty
        $results[1].Name | Should -Be "技術"
        $results[1].Ok | Should -Be $true
        $results[1].Reason | Should -Be ""
        $results[2].Ok | Should -Be $false
        $results[2].Reason | Should -Match "空"
        Test-Path "$dir\技術" | Should -Be $false
        Test-Path "$dir\総務" | Should -Be $true
        # 消せなかった営業の記録は残る
        $status = readStatusFile $ws.Status
        $status.Rows.ContainsKey("営業\a.xlsx") | Should -Be $true
        $status.Rows.ContainsKey("技術\b.xlsx") | Should -Be $false
    }
}

Describe "IndexCatalog.RenameIndex / RemoveIndex の外したフォルダの記録（searchExcludes）" -Tag Io {
    BeforeEach {
        # It ごとに作り直す（前の It の改名が残らないように）
        Remove-Item -LiteralPath "$TestDrive\excl" -Recurse -Force -ErrorAction SilentlyContinue
        $script:ws = newCatalogWorkspace "$TestDrive\excl"
        $script:dir = $script:ws.Index
        $script:settings = $script:ws.Settings
        New-Item -ItemType Directory -Path "$($script:dir)\Sales", "$($script:dir)\Sales2", "$($script:dir)\技術" -Force | Out-Null
        writeSearchExcludes @(
            [pscustomobject]@{ Path = "$($script:dir)\Sales\見積"; Subfolders = $true },
            [pscustomobject]@{ Path = "$($script:dir)\Sales"; Subfolders = $false },
            [pscustomobject]@{ Path = "$($script:dir)\Sales2\見積"; Subfolders = $true },
            [pscustomobject]@{ Path = "$($script:dir)\技術\図面"; Subfolders = $true }) $script:settings
    }

    It "名前を変えると、旧名の下の記録を消し、ほかのインデックス（頭が同じ名前も）の記録は残す: <old> → <new>" -TestCases @(
        @{ old = "Sales"; new = "営業" },
        @{ old = "Sales"; new = "sales" }
    ) {
        param ($old, $new)
        $script:ws.Catalog.RenameIndex($old, $new)

        $paths = @(readSearchExcludes $script:settings | ForEach-Object { $_.Path })
        $paths | Should -Be @("$($script:dir)\Sales2\見積", "$($script:dir)\技術\図面")
    }

    It "新名の下に記録が残っていても消す" {
        writeSearchExcludes @([pscustomobject]@{ Path = "$($script:dir)\営業\古い"; Subfolders = $true }) $script:settings
        $script:ws.Catalog.RenameIndex("Sales", "営業")
        @(readSearchExcludes $script:settings).Count | Should -Be 0
    }

    It "インデックスを削除すると、その下の記録を消し、ほかのインデックスの記録は残す" {
        $script:ws.Catalog.RemoveIndex("Sales")

        $paths = @(readSearchExcludes $script:settings | ForEach-Object { $_.Path })
        $paths | Should -Be @("$($script:dir)\Sales2\見積", "$($script:dir)\技術\図面")
    }

    It "記録を消せなくても、名前の変更・削除は例外にならない" {
        # 設定ファイルの場所をフォルダにして、読み書きを失敗させる
        $bad = "$TestDrive\excl\badsetting"
        New-Item -ItemType Directory -Path $bad -Force | Out-Null
        $catalog = newIndexCatalog $script:ws.Dir $bad
        { $catalog.RenameIndex("Sales", "営業") } | Should -Not -Throw
        { $catalog.RemoveIndex("営業") } | Should -Not -Throw
        Test-Path -LiteralPath "$($script:dir)\営業" | Should -Be $false
    }
}

Describe "IndexCatalog 5 つの保存先（targetFolders・取り込み一覧・元のフォルダ.txt・searchExcludes・system_index）" -Tag Io {
    BeforeEach {
        Remove-Item -LiteralPath "$TestDrive\five" -Recurse -Force -ErrorAction SilentlyContinue
        $script:ws = newCatalogWorkspace "$TestDrive\five"
        $script:dir = $script:ws.Index
        # 営業（A）と営業部（A を頭に含む名前。#125 のように、頭が同じ名前の記録を巻き込まない）
        foreach ($name in @("営業", "営業部")) {
            writeListFile "$($script:dir)\$name\2024\A社.xlsx\明細.tsv" @("a")
        }
        $script:folders = @(
            [pscustomobject]@{ Name = "営業"; Path = "C:\data\営業"; Enabled = $true },
            [pscustomobject]@{ Name = "営業部"; Path = "C:\data\営業部"; Enabled = $true })
        writeTargetFolders $script:folders $script:ws.Settings
        $script:ws.Catalog.WriteSourceFile($script:folders)
        writeStatusFile $script:folders @(
            (newStatusRow "営業\2024\A社.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00"),
            (newStatusRow "営業部\2024\A社.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00")
        ) $script:ws.Status
        writeSearchExcludes @(
            [pscustomobject]@{ Path = "$($script:dir)\営業\2024"; Subfolders = $true },
            [pscustomobject]@{ Path = "$($script:dir)\営業部\2024"; Subfolders = $true }) $script:ws.Settings
        [void](updateSystemIndexes $script:dir "$($script:ws.Dir)\system_index" "$($script:ws.Dir)\system_index_state.tsv" { $false })
        # gram 索引（インデックス名ごとのフォルダ）
        foreach ($name in @("営業", "営業部")) {
            [System.IO.Directory]::CreateDirectory("$($script:ws.Dir)\store\gram\$name") | Out-Null
            [System.IO.File]::WriteAllText("$($script:ws.Dir)\store\gram\$name\toc.tsv", "tebunko-gram 1`r`nnext 1`r`n")
        }
    }

    It "名前を変えると 5 つとも新しい名前になり、頭が同じ別のインデックスはそのまま残る" {
        $catalog = $script:ws.Catalog
        $catalog.RenameIndex("営業", "第一営業")
        # 画面は、設定の一覧と元のフォルダの記録を、改名のあとに書き直す
        $renamed = @(
            [pscustomobject]@{ Name = "第一営業"; Path = "C:\data\営業"; Enabled = $true },
            [pscustomobject]@{ Name = "営業部"; Path = "C:\data\営業部"; Enabled = $true })
        $catalog.SaveTargets($renamed)
        $catalog.WriteSourceFile($renamed)

        # 1. 設定の targetFolders
        @(getTargetFolders $script:ws.Settings | ForEach-Object { $_.Name }) -join "," | Should -Be "第一営業,営業部"
        # 2. 取り込み一覧
        $status = readStatusFile $script:ws.Status
        @($status.Folders | ForEach-Object { $_.Name }) -join "," | Should -Be "第一営業,営業部"
        $status.Rows.ContainsKey("第一営業\2024\A社.xlsx") | Should -Be $true
        $status.Rows.ContainsKey("営業部\2024\A社.xlsx") | Should -Be $true
        # 3. 元のフォルダの記録（フォルダと一緒に移る）
        Test-Path -LiteralPath "$($script:dir)\第一営業\2024\A社.xlsx\明細.tsv" | Should -Be $true
        Test-Path -LiteralPath "$($script:dir)\営業" | Should -Be $false
        (Get-ChildItem -LiteralPath "$($script:dir)\第一営業" -Filter "source_folder.txt" -File).Count | Should -Be 1
        # 4. searchExcludes（旧名の下は消え、営業部の下は残る）
        @(readSearchExcludes $script:ws.Settings | ForEach-Object { $_.Path }) | Should -Be @("$($script:dir)\営業部\2024")
        # 5. system_index（旧名の分は消え、営業部の分は残る）
        [System.IO.Directory]::Exists("$($script:ws.Dir)\system_index\営業") | Should -Be $false
        [System.IO.Directory]::Exists("$($script:ws.Dir)\system_index\営業部") | Should -Be $true
        # 6. gram 索引（旧名の分は消え、営業部の分は残る）
        [System.IO.Directory]::Exists("$($script:ws.Dir)\store\gram\営業") | Should -Be $false
        [System.IO.Directory]::Exists("$($script:ws.Dir)\store\gram\営業部") | Should -Be $true
    }

    It "削除すると 5 つとも消え、頭が同じ別のインデックスはそのまま残る" {
        $catalog = $script:ws.Catalog
        $catalog.RemoveIndex("営業")
        $rest = @([pscustomobject]@{ Name = "営業部"; Path = "C:\data\営業部"; Enabled = $true })
        $catalog.SaveTargets($rest)
        $catalog.WriteSourceFile($rest)

        @(getTargetFolders $script:ws.Settings | ForEach-Object { $_.Name }) -join "," | Should -Be "営業部"
        $status = readStatusFile $script:ws.Status
        @($status.Folders | ForEach-Object { $_.Name }) -join "," | Should -Be "営業部"
        $status.Rows.ContainsKey("営業\2024\A社.xlsx") | Should -Be $false
        $status.Rows.ContainsKey("営業部\2024\A社.xlsx") | Should -Be $true
        Test-Path -LiteralPath "$($script:dir)\営業" | Should -Be $false
        Test-Path -LiteralPath "$($script:dir)\営業部\2024\A社.xlsx\明細.tsv" | Should -Be $true
        @(readSearchExcludes $script:ws.Settings | ForEach-Object { $_.Path }) | Should -Be @("$($script:dir)\営業部\2024")
        [System.IO.Directory]::Exists("$($script:ws.Dir)\system_index\営業") | Should -Be $false
        [System.IO.Directory]::Exists("$($script:ws.Dir)\system_index\営業部") | Should -Be $true
        [System.IO.Directory]::Exists("$($script:ws.Dir)\store\gram\営業") | Should -Be $false
        [System.IO.Directory]::Exists("$($script:ws.Dir)\store\gram\営業部") | Should -Be $true
    }
}

Describe "IndexCatalog の名前の割り当てと保存" -Tag Io {
    It "AssignAndSaveNames は名前の無い一覧に名前を割り当てて、設定に保存する" {
        $ws = newCatalogWorkspace "$TestDrive\assign"
        $folders = @([pscustomobject]@{ Name = ""; Path = "C:\data\見積"; Enabled = $true })
        writeTargetFolders $folders $ws.Settings

        $saved = @($ws.Catalog.AssignAndSaveNames($folders))

        $saved.Count | Should -Be 1
        $saved[0].Name | Should -Not -BeNullOrEmpty
        (getTargetFolders $ws.Settings)[0].Name | Should -Be $saved[0].Name
    }

    It "SaveTargets・WriteSourceFile は、設定の一覧と元のフォルダの記録を書く" {
        $ws = newCatalogWorkspace "$TestDrive\save"
        New-Item -ItemType Directory -Path "$($ws.Index)\営業" -Force | Out-Null
        $folders = @([pscustomobject]@{ Name = "営業"; Path = "C:\data\営業"; Enabled = $true })

        $ws.Catalog.SaveTargets($folders)
        $ws.Catalog.WriteSourceFile($folders)

        (getTargetFolders $ws.Settings)[0].Path | Should -Be "C:\data\営業"
        (getSearchIndexes $ws.Index $ws.Status $ws.Settings)[0].SourcePath | Should -Be "C:\data\営業"
    }
}

Describe "IndexCatalog.RenameIndexWithSource" -Tag Io {
    It "改名したあと、元のフォルダの記録（source_folder.txt）を一覧に合わせて書き直し、理由は空で返す" {
        $ws = newCatalogWorkspace "$TestDrive\rws1"
        New-Item -ItemType Directory -Path "$($ws.Index)\営業" -Force | Out-Null
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "営業" }) @() $ws.Status

        $result = $ws.Catalog.RenameIndexWithSource("営業", "営業部", @([pscustomobject]@{ Name = "営業部"; Path = "C:\data" }))

        $result.SourceFolderError | Should -Be ""
        Test-Path "$($ws.Index)\営業" | Should -Be $false
        (Get-Content -LiteralPath "$($ws.Index)\営業部\${sourceFolderFileName}") -join "|" | Should -BeLike "*営業部*C:\data*"
    }

    It "記録を書けなくても改名は成功にし、書けなかった理由（書けなかったファイルの名前を含む元のメッセージ）を返す" {
        $ws = newCatalogWorkspace "$TestDrive\rws2"
        New-Item -ItemType Directory -Path "$($ws.Index)\営業" -Force | Out-Null
        # 記録のファイル名の場所にフォルダがあると、書けない
        New-Item -ItemType Directory -Path "$($ws.Index)\営業\${sourceFolderFileName}" -Force | Out-Null
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "営業" }) @() $ws.Status

        $result = $ws.Catalog.RenameIndexWithSource("営業", "営業部", @([pscustomobject]@{ Name = "営業部"; Path = "C:\data" }))

        Test-Path "$($ws.Index)\営業部" | Should -Be $true
        $result.SourceFolderError | Should -Not -BeNullOrEmpty
        $result.SourceFolderError | Should -BeLike "*${sourceFolderFileName}*"
    }

    It "改名に失敗したら、例外をそのまま出す（同じ名前のフォルダが既にある）" {
        $ws = newCatalogWorkspace "$TestDrive\rws3"
        New-Item -ItemType Directory -Path "$($ws.Index)\営業", "$($ws.Index)\営業部" -Force | Out-Null

        { $ws.Catalog.RenameIndexWithSource("営業", "営業部", @()) } | Should -Throw
    }
}