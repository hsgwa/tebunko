# クロール（tebunko\indexer\indexer_plan.ps1 の createTargetList）のテスト。
BeforeDiscovery {
    # -TestCases の表が使う一覧の状態（$stateDone など）
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\indexer\indexer_plan.ps1"

    function newPrevious {
        param ($rows = @())
        $map = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($row in $rows) { $map[$row.相対パス] = $row }
        return , $map
    }

    function newCounts {
        param ([hashtable]$pairs = @{})
        $counts = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($key in $pairs.Keys) { $counts[$key] = $pairs[$key] }
        return , $counts
    }
}

Describe "createTargetList" -Tag Io {
    BeforeAll {
        $source = Join-Path $TestDrive "src"
        [System.IO.Directory]::CreateDirectory($source) | Out-Null
        $file = Join-Path $source "a.xlsx"
        [System.IO.File]::WriteAllText($file, "dummy")
        $info = Get-Item -LiteralPath $file
        $updated = formatFileTime $info.LastWriteTime
        $size = [string]$info.Length
        $folder = @{ Path = $source; Name = "売上" }
        # removeBookDir が実際のインデックスを見ないよう、テスト用のフォルダに向ける
        ${indexDir} = Join-Path $TestDrive "index"
        $workspace = newTestWorkspace @{ IndexDir = ${indexDir} }
    }

    # 一覧の行（state が $null なら一覧に無い。modified は元のファイルが更新されたか、version は抽出版）と
    # TSV の数え上げ（tsv が $null なら数えない・0 なら数え上げに無い）→ 取り込み対象・失敗の数と、取り込み予定のどの件数に数えるか
    It "<name>" -TestCases @(
        @{ name = "一覧に無いファイルは取り込み対象になる（新規）"; state = $null; modified = $false; version = "4"; tsv = $null; targets = 1; failed = 0; field = "新規" }
        @{ name = "取り込み済みで更新が無ければ取り込まない"; state = $stateDone; modified = $false; version = "4"; tsv = 1; targets = 0; failed = 0; field = "" }
        @{ name = "前の抽出版で取り込んだファイルは、更新が無くても取り込み直す（更新ありに数える）"; state = $stateDone; modified = $false; version = ""; tsv = 1; targets = 1; failed = 0; field = "更新あり" }
        @{ name = "取り込み済みでも TSV が無ければ取り込み直す（インデックスなし）"; state = $stateDone; modified = $false; version = "4"; tsv = 0; targets = 1; failed = 0; field = "インデックスなし" }
        @{ name = "更新されていれば取り込み対象になる（更新あり）"; state = $stateDone; modified = $true; version = "4"; tsv = 1; targets = 1; failed = 0; field = "更新あり" }
        @{ name = "前回失敗して更新が無ければ、取り込み対象ではなく失敗として返す"; state = $stateFailed; modified = $false; version = "4"; tsv = 1; targets = 0; failed = 1; field = "前回失敗" }
        @{ name = "前回「未取り込み」で終わっていれば取り込み対象になる（前回未完了）"; state = $stateNew; modified = $false; version = "4"; tsv = 1; targets = 1; failed = 0; field = "前回未完了" }
    ) {
        param ($name, $state, $modified, $version, $tsv, $targets, $failed, $field)
        $rows = @()
        if ($null -ne $state) {
            $rowUpdated = if ($modified) { "2000/01/01 00:00:00" } else { $updated }
            $rows = @((newStatusRow "売上\a.xlsx" $rowUpdated $size $state 1 $updated "" $version))
        }
        $counts = $null
        if ($tsv -eq 0) {
            $counts = newCounts
        } elseif ($tsv) {
            $counts = newCounts @{ "売上\a.xlsx" = $tsv }
        }
        $result = createTargetList $folder (newPrevious $rows) $counts
        $result.Rows.Count | Should -Be 1
        $result.Targets.Count | Should -Be $targets
        $result.Failed.Count | Should -Be $failed
        if ($field) {
            $result.Plan.$field | Should -Be 1
        }
        $result.Plan.ファイル数 | Should -Be 1
        $result.Plan.取り込み対象 | Should -Be $targets
    }
}

Describe "createTargetList（サブフォルダ・無くなったファイル）" -Tag Io {
    BeforeAll {
        $source = Join-Path $TestDrive "src2"
        [System.IO.Directory]::CreateDirectory("$source\2024") | Out-Null
        $file = Join-Path $source "2024\b.docx"
        [System.IO.File]::WriteAllText($file, "dummy")
        (Get-Item -LiteralPath $file).LastWriteTime = [datetime]"2024/04/01 09:00:00"
        $folder = @{ Path = $source; Name = "経理" }
        ${indexDir} = Join-Path $TestDrive "index2"
        $workspace = newTestWorkspace @{ IndexDir = ${indexDir} }
        $bookDir = Join-Path ${indexDir} "経理\2024\b.docx"

        function newIndexTsv([string]$name, [datetime]$time) {
            [System.IO.Directory]::CreateDirectory($bookDir) | Out-Null
            $path = Join-Path $bookDir $name
            [System.IO.File]::WriteAllText($path, "1`t見積")
            (Get-Item -LiteralPath $path).LastWriteTime = $time
        }
    }

    AfterEach {
        if (Test-Path -LiteralPath ${indexDir}) { Remove-Item -LiteralPath ${indexDir} -Recurse -Force }
    }

    It "サブフォルダのファイルは、インデックス名とフォルダからの相対パスでつなぐ" {
        $result = createTargetList $folder (newPrevious) $null
        $result.Rows[0].相対パス | Should -Be "経理\2024\b.docx"
    }

    It "一覧に無いファイルは、元のファイルより新しいインデックスのフォルダが残っていても取り込み対象にする" {
        newIndexTsv "1ページ.tsv" ([datetime]"2024/04/02 09:00:00")
        $result = createTargetList $folder (newPrevious) $null
        $result.Targets.Count | Should -Be 1
        $result.Plan.新規 | Should -Be 1
    }

    It "一覧にあって元のファイルが無くなったものは、インデックスを消して一覧から除く" {
        $gone = Join-Path ${indexDir} "経理\消えた.xlsx"
        [System.IO.Directory]::CreateDirectory($gone) | Out-Null
        [System.IO.File]::WriteAllText("$gone\Sheet1.tsv", "1`t見積")
        $previous = newPrevious @((newStatusRow "経理\消えた.xlsx" "2024/01/01 00:00:00" "10" ${stateDone} 1 "2024/01/01 00:00:00"))
        $result = createTargetList $folder $previous $null
        @($result.Rows | Where-Object { $_.相対パス -eq "経理\消えた.xlsx" }).Count | Should -Be 0
        Test-Path -LiteralPath $gone | Should -Be $false
    }

    It "ほかのインデックスの行には触らない（名前が前方一致するインデックスも別のものとして扱う）" {
        $other = Join-Path ${indexDir} "経理2\c.xlsx"
        [System.IO.Directory]::CreateDirectory($other) | Out-Null
        $previous = newPrevious @((newStatusRow "経理2\c.xlsx" "2024/01/01 00:00:00" "10" ${stateDone} 1 "2024/01/01 00:00:00"))
        $result = createTargetList $folder $previous $null
        @($result.Rows | Where-Object { $_.相対パス -eq "経理2\c.xlsx" }).Count | Should -Be 0
        Test-Path -LiteralPath $other | Should -Be $true
    }

    It "アクセスできないフォルダがあったときは、見つからなかったファイルの行とインデックスを残す" {
        Mock findTargetFiles { @{ Root = $source; Files = @(Get-Item -LiteralPath $file); HasError = $true } }
        $gone = Join-Path ${indexDir} "経理\読めない\d.xlsx"
        [System.IO.Directory]::CreateDirectory($gone) | Out-Null
        $previous = newPrevious @((newStatusRow "経理\読めない\d.xlsx" "2024/01/01 00:00:00" "10" ${stateDone} 1 "2024/01/01 00:00:00"))
        $result = createTargetList $folder $previous $null
        @($result.Rows | Where-Object { $_.相対パス -eq "経理\読めない\d.xlsx" }).Count | Should -Be 1
        Test-Path -LiteralPath $gone | Should -Be $true
        # \\?\ の付かないパスで列挙されたファイルも、フォルダからの相対パスにする
        $result.Rows[0].相対パス | Should -Be "経理\2024\b.docx"
    }
}

Describe "findTargetFiles" -Tag Io {
    BeforeAll {
        $source = Join-Path $TestDrive "scan"
        [System.IO.Directory]::CreateDirectory("$source\下\さらに下") | Out-Null
        foreach ($name in @("a.xlsx", "下\b.DOCX", "下\さらに下\c.pptm", "d.txt", "e.pdf", ('~$' + "a.xlsx"), "f.xls", "g.ppt", "h.py", "i.cpp", "j.exe", "k.ts")) {
            [System.IO.File]::WriteAllText((Join-Path $source $name), "dummy")
        }
        $workspace = newTestWorkspace @{} (Join-Path $TestDrive "ws-out-of-tree")
    }

    It "サブフォルダも含めて Office・テキストの拡張子のファイルだけを返す（.py・.cpp も対象、.exe・.ts は対象外、大文字の拡張子も含め、~$ で始まるロックファイルは除く）" {
        $scan = findTargetFiles $source
        @($scan.Files | ForEach-Object { $_.Name } | Sort-Object) -join "," | Should -Be "a.xlsx,b.DOCX,c.pptm,d.txt,f.xls,g.ppt,h.py,i.cpp"
        $scan.HasError | Should -Be $false
    }

    It "末尾に \ を付けたフォルダでも同じフォルダを列挙する" {
        $scan = findTargetFiles "$source\"
        $scan.Root.TrimEnd("\") | Should -Be $source
        $scan.Files.Count | Should -Be 8
    }
}

Describe "findTargetFiles（tebunko が作ったものの除外）" -Tag Io {
    BeforeEach {
        # It ごとに別のフォルダにする（前の It で作ったファイルが残って混ざらないように）
        $source = Join-Path $TestDrive "scan2-$(New-Guid)"
        [System.IO.Directory]::CreateDirectory($source) | Out-Null

        function newFile([string]$relPath, [string]$content = "dummy") {
            $path = Join-Path $source $relPath
            [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
            [System.IO.File]::WriteAllText($path, $content)
        }
    }

    It "今のワークスペース（クロール対象フォルダの中）の本文インデックス・システムインデックス・取り込み一覧は外れ、利用者が置いた .xlsx・.txt は取り込む" {
        $wsDir = Join-Path $source "ws"
        $workspace = newTestWorkspace @{} $wsDir
        newFile "ws\content_index\sheet.tsv"
        newFile "ws\system_index\system_index.txt"
        newFile "ws\ingest_status.tsv"
        newFile "ws\my.xlsx"
        newFile "ws\my.txt"

        $scan = findTargetFiles $source
        @($scan.Files | ForEach-Object { $_.Name } | Sort-Object) -join "," | Should -Be "my.txt,my.xlsx"
    }

    It "ingest_status.tsv のある別のフォルダ（ほかのワークスペース）の content_index の下は外れる" {
        $workspace = newTestWorkspace @{} (Join-Path $TestDrive "ws-out-of-tree2")
        newFile "other-ws\ingest_status.tsv"
        newFile "other-ws\content_index\sheet.tsv"
        newFile "other-ws\my-file.txt"

        $scan = findTargetFiles $source
        @($scan.Files | ForEach-Object { $_.Name } | Sort-Object) | Should -Contain "my-file.txt"
        @($scan.Files | ForEach-Object { $_.Name }) | Should -Not -Contain "sheet.tsv"
    }

    It "名前で分かる本文インデックス・前の版の集約ファイル・システムインデックスは、どこにあっても外れる" {
        $workspace = newTestWorkspace @{} (Join-Path $TestDrive "ws-out-of-tree3")
        newFile "content_index.txt.001.tsv"
        newFile "content.xlsx.001.tsv"
        newFile "システムインデックス.txt"
        newFile "system_index.001.txt"
        newFile "keep.txt"

        $scan = findTargetFiles $source
        @($scan.Files | ForEach-Object { $_.Name }) | Should -Not -Contain "content_index.txt.001.tsv"
        @($scan.Files | ForEach-Object { $_.Name }) | Should -Not -Contain "content.xlsx.001.tsv"
        @($scan.Files | ForEach-Object { $_.Name }) | Should -Not -Contain "システムインデックス.txt"
        @($scan.Files | ForEach-Object { $_.Name }) | Should -Not -Contain "system_index.001.txt"
        @($scan.Files | ForEach-Object { $_.Name }) | Should -Contain "keep.txt"
    }

    It "~`$ で始まる一時ファイルは除外する" {
        $workspace = newTestWorkspace @{} (Join-Path $TestDrive "ws-out-of-tree4")
        newFile ('~$locked.txt')
        $scan = findTargetFiles $source
        @($scan.Files | ForEach-Object { $_.Name }) | Should -Not -Contain '~$locked.txt'
    }
}

Describe "selectOnlyNames" -Tag Unit {
    BeforeAll {
        $folders = @(
            [pscustomobject]@{ Name = "営業"; Enabled = $true }
            [pscustomobject]@{ Name = "技術"; Enabled = $false }
        )
    }

    It "<name>" -TestCases @(
        @{ name = "チェックの付いた名前は選ばれる"; names = @("営業"); selected = @("営業"); skipped = @() }
        @{ name = "チェックが付いていない名前は、理由付きで外す"; names = @("営業", "技術"); selected = @("営業"); skipped = @("技術:チェックが付いていません") }
        @{ name = "設定に無い名前は、理由付きで外す"; names = @("総務"); selected = @(); skipped = @("総務:設定にありません") }
        @{ name = "空の名前は無いものとして扱う"; names = @("", "営業"); selected = @("営業"); skipped = @() }
    ) {
        param ($name, $names, $selected, $skipped)
        $result = selectOnlyNames $folders $names
        @($result.Selected | Sort-Object) | Should -Be @($selected)
        @($result.Skipped | ForEach-Object { "$($_.Name):$($_.Reason)" }) | Should -Be @($skipped)
    }
}

Describe "createTargetList（クラウドにだけあるファイル）" -Tag Io {
    BeforeAll {
        $source = Join-Path $TestDrive "cloud_src"
        [System.IO.Directory]::CreateDirectory($source) | Out-Null
        $file = Join-Path $source "c.xlsx"
        [System.IO.File]::WriteAllText($file, "dummy-cloud")
        # クラウドにだけあるファイルは、OFFLINE 属性で代わりに作る（ダウンロードは起きない）
        [System.IO.File]::SetAttributes($file, [System.IO.FileAttributes]::Offline)
        $info = Get-Item -LiteralPath $file
        $updated = formatFileTime $info.LastWriteTime
        $size = [string]$info.Length
        $folder = @{ Path = $source; Name = "雲" }
        ${indexDir} = Join-Path $TestDrive "cloud_index"
        $workspace = newTestWorkspace @{ IndexDir = ${indexDir} }
        $bookDir = Join-Path ${indexDir} "雲\c.xlsx"
    }

    AfterAll {
        [System.IO.File]::SetAttributes($file, [System.IO.FileAttributes]::Normal)
    }

    It "新規でクラウドにだけあれば、取り込み対象にも一覧にも入れず、クラウドの一覧に入れる" {
        $result = createTargetList $folder (newPrevious) $null
        $result.Targets.Count | Should -Be 0
        $result.Failed.Count | Should -Be 0
        $result.Rows.Count | Should -Be 0
        $result.Cloud.Count | Should -Be 1
        $result.Cloud[0].RelPath | Should -Be "雲\c.xlsx"
        $result.Cloud[0].Failed | Should -Be $false
        $result.Cloud[0].Row.状態 | Should -Be ${stateNew}
        $result.Plan.取り込み対象 | Should -Be 0
        $result.Plan.新規 | Should -Be 0
        $result.Plan.クラウド | Should -Be 1
        $result.Plan.クラウド失敗 | Should -Be 0
        $result.Plan.クラウド容量 | Should -Be ([long]$size)
    }

    It "取り込み済みで更新が無ければ、あとでクラウドにだけになっても「済」のままにする（クラウドの数に入れない）" {
        $previous = newPrevious @((newStatusRow "雲\c.xlsx" $updated $size ${stateDone} 1 $updated "" "99"))
        $result = createTargetList $folder $previous $null
        $result.Cloud.Count | Should -Be 0
        $result.Rows.Count | Should -Be 1
        $result.Rows[0].状態 | Should -Be ${stateDone}
        $result.Plan.クラウド | Should -Be 0
    }

    It "更新されていてクラウドにだけあれば、前回の行をそのまま残し、取り込み対象にしない" {
        $oldRow = newStatusRow "雲\c.xlsx" "2000/01/01 00:00:00" "5" ${stateDone} 1 "2000/01/01 00:00:00" "" "99"
        $result = createTargetList $folder (newPrevious @($oldRow)) $null
        $result.Targets.Count | Should -Be 0
        $result.Rows.Count | Should -Be 1
        $result.Rows[0].更新日時 | Should -Be "2000/01/01 00:00:00"
        $result.Cloud.Count | Should -Be 1
        $result.Cloud[0].PendingOld | Should -Be $false
    }

    It "前回失敗し、更新が無く、クラウドにだけあれば、失敗には入れず、クラウドの失敗として数える" {
        $failedRow = newStatusRow "雲\c.xlsx" $updated $size ${stateFailed} 0 $updated "読めない" "99"
        $result = createTargetList $folder (newPrevious @($failedRow)) $null
        $result.Failed.Count | Should -Be 0
        $result.Targets.Count | Should -Be 0
        $result.Cloud.Count | Should -Be 1
        $result.Cloud[0].Failed | Should -Be $true
        $result.Cloud[0].Row.状態 | Should -Be ${stateFailed}
        $result.Rows.Count | Should -Be 1
        $result.Plan.前回失敗 | Should -Be 0
        $result.Plan.クラウド | Should -Be 0
        $result.Plan.クラウド失敗 | Should -Be 1
        $result.Plan.クラウド失敗容量 | Should -Be ([long]$size)
    }

    It "前回未完了でクラウドにだけあれば、一覧に行を作らない（前回の取り込み途中のインデックスは、取り込みを決める側が消す）" {
        $pendingRow = newStatusRow "雲\c.xlsx" $updated $size ${stateNew}
        $result = createTargetList $folder (newPrevious @($pendingRow)) $null
        $result.Rows.Count | Should -Be 0
        $result.Targets.Count | Should -Be 0
        $result.Cloud.Count | Should -Be 1
        $result.Cloud[0].PendingOld | Should -Be $true
    }
}
