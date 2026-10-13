# インデックスの改名・削除・名前の割り当てと保存（tebunko\index\index_catalog.ps1 の IndexCatalog）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    function script:newCatalogWorkspace {
        # テスト用のワークスペース（dir）と、その IndexCatalog を作る。設定ファイルは dir\setting.config
        param ([string]$dir)

        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        return @{
            Dir      = $dir
            Index    = "$dir\content_index"
            Status   = "$dir\ingest_status.tsv"
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
    ) {
        param ($badName)
        $ws = newCatalogWorkspace "$TestDrive\remove4"
        New-Item -ItemType Directory -Path "$($ws.Index)\営業" -Force | Out-Null
        writeListFile "$($ws.Dir)\外\a.tsv" @("本文")
        writeStatusFile @([pscustomobject]@{ Path = "C:\data"; Name = "営業" }) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateNew)
        ) $ws.Status

        { $ws.Catalog.RemoveIndex($badName) } | Should -Throw

        Test-Path -LiteralPath "$($ws.Dir)\外\a.tsv" | Should -Be $true
        Test-Path -LiteralPath "$($ws.Index)\営業" | Should -Be $true
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