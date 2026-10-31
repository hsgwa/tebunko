# 取り込む順番を決める処理（tebunko\indexer\indexer_plan.ps1 の createTargetList・selectOnlyNames・IngestPlanner）のテスト。
BeforeDiscovery {
    # -TestCases の表が使う一覧の状態（$stateDone など）
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\indexer\indexer_plan.ps1"
    . "${scriptsDir}\tebunko\indexer\index_migrate.ps1"  # findDroppedIndexes（IngestPlanner）

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

    It "一覧にあって元のファイルが無くなったものは、一覧から除いて Removed に返し、インデックスのフォルダは消さない" {
        $gone = Join-Path ${indexDir} "経理\消えた.xlsx"
        [System.IO.Directory]::CreateDirectory($gone) | Out-Null
        [System.IO.File]::WriteAllText("$gone\Sheet1.tsv", "1`t見積")
        $previous = newPrevious @((newStatusRow "経理\消えた.xlsx" "2024/01/01 00:00:00" "10" ${stateDone} 1 "2024/01/01 00:00:00"))
        $result = createTargetList $folder $previous $null
        @($result.Rows | Where-Object { $_.相対パス -eq "経理\消えた.xlsx" }).Count | Should -Be 0
        @($result.Removed) | Should -Be @("経理\消えた.xlsx")
        Test-Path -LiteralPath $gone | Should -Be $true
    }

    It "中止の合図（shouldStop）が、一覧化のどの時点で立っても、止まったなら Stopped と空の Removed を返し、インデックスは消さない" -TestCases @(
        @{ from = 1 }
        @{ from = 2 }
        @{ from = 3 }
        @{ from = 4 }
        @{ from = 5 }
    ) {
        param ($from)
        $gone = Join-Path ${indexDir} "経理\消えた.xlsx"
        [System.IO.Directory]::CreateDirectory($gone) | Out-Null
        [System.IO.File]::WriteAllText("$gone\Sheet1.tsv", "1`t見積")
        $previous = newPrevious @((newStatusRow "経理\消えた.xlsx" "2024/01/01 00:00:00" "10" ${stateDone} 1 "2024/01/01 00:00:00"))
        $global:planStopCalls = 0
        $global:planStopFrom = $from
        try {
            $result = createTargetList $folder $previous $null { $global:planStopCalls++; $global:planStopCalls -ge $global:planStopFrom }
        } finally {
            Remove-Variable -Name planStopCalls, planStopFrom -Scope Global -ErrorAction SilentlyContinue
        }
        # 数えるだけで何も消さない（止まらずに最後まで進んだ回は、無くなったファイルとして Removed に返す）
        Test-Path -LiteralPath $gone | Should -Be $true
        if ($result.Stopped) {
            @($result.Removed).Count | Should -Be 0
        } else {
            $result.Removed.Count | Should -Be 1
        }
        # 最初から立っていれば必ず止まる
        if ($from -eq 1) { $result.Stopped | Should -Be $true }
    }

    It "中止の合図が立っていなければ、Stopped は偽" {
        (createTargetList $folder (newPrevious) $null { $false }).Stopped | Should -Be $false
        (createTargetList $folder (newPrevious) $null).Stopped | Should -Be $false
    }

    It "ほかのインデックスの行には触らない（名前が前方一致するインデックスも別のものとして扱う）" {
        $other = Join-Path ${indexDir} "経理2\c.xlsx"
        [System.IO.Directory]::CreateDirectory($other) | Out-Null
        $previous = newPrevious @((newStatusRow "経理2\c.xlsx" "2024/01/01 00:00:00" "10" ${stateDone} 1 "2024/01/01 00:00:00"))
        $result = createTargetList $folder $previous $null
        @($result.Rows | Where-Object { $_.相対パス -eq "経理2\c.xlsx" }).Count | Should -Be 0
        Test-Path -LiteralPath $other | Should -Be $true
    }

    It "置き場の外を指す相対パスの行があっても、数えは例外にならず、removeBookDir は外のフォルダを消さずにログに出す" -TestCases @(
        @{ rel = "経理\..\..\外\消えた.xlsx"; outsideName = "外" }
        @{ rel = "経理\..\..\index2外\消えた.xlsx"; outsideName = "index2外" }   # 置き場の名前で始まるだけの別のフォルダ
        @{ rel = "経理\..\.."; outsideName = "外" }
    ) {
        param ($rel, $outsideName)
        Mock writeIndexerLog {}
        $outside = Join-Path $TestDrive $outsideName
        [System.IO.Directory]::CreateDirectory("$outside\消えた.xlsx") | Out-Null
        [System.IO.File]::WriteAllText("$outside\消えた.xlsx\Sheet1.tsv", "1`t見積")
        [System.IO.File]::WriteAllText("$TestDrive\残る.txt", "a")
        $previous = newPrevious @((newStatusRow $rel "2024/01/01 00:00:00" "10" ${stateDone} 1 "2024/01/01 00:00:00"))
        $result = createTargetList $folder $previous $null
        removeBookDir (getBookDir $result.Removed[0])
        Test-Path -LiteralPath "$outside\消えた.xlsx\Sheet1.tsv" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\残る.txt" | Should -Be $true
        Should -Invoke writeIndexerLog -Times 1 -Exactly -Scope It -ParameterFilter { $text -like "*削除しませんでした*" -and $color -eq "Yellow" }
    }

    It "removeBookDir は、置き場の外のフォルダを消さずにログに出し、中のフォルダは消す" {
        Mock writeIndexerLog {}
        $outside = Join-Path $TestDrive "外2"
        [System.IO.Directory]::CreateDirectory($outside) | Out-Null
        [System.IO.File]::WriteAllText("$outside\a.tsv", "a")
        $inside = Join-Path ${indexDir} "経理\中.xlsx"
        [System.IO.Directory]::CreateDirectory($inside) | Out-Null

        removeBookDir (Join-Path ${indexDir} "..\外2")
        removeBookDir ${indexDir}

        Test-Path -LiteralPath "$outside\a.tsv" | Should -Be $true
        Test-Path -LiteralPath $inside | Should -Be $true
        Should -Invoke writeIndexerLog -Times 2 -Exactly -Scope It -ParameterFilter { $text -like "*削除しませんでした*" -and $color -eq "Yellow" }

        removeBookDir $inside
        Test-Path -LiteralPath $inside | Should -Be $false
    }

    It "removeBookDir は、途中の階層がジャンクションなら、置き場の下に見えても先を消さずにログに出す" {
        Mock writeIndexerLog {}
        $outside = Join-Path $TestDrive "リンクの先"
        [System.IO.File]::WriteAllText((Join-Path $outside "経理\中.xlsx\a.tsv" | ForEach-Object { [System.IO.Directory]::CreateDirectory((Split-Path $_ -Parent)) | Out-Null; $_ }), "a")
        [System.IO.Directory]::CreateDirectory(${indexDir}) | Out-Null
        if (!(newTestJunction (Join-Path ${indexDir} "営業") $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        removeBookDir (Join-Path ${indexDir} "営業\経理\中.xlsx")

        Test-Path -LiteralPath "$outside\経理\中.xlsx\a.tsv" | Should -Be $true
        Should -Invoke writeIndexerLog -Times 1 -Exactly -Scope It -ParameterFilter { $text -like "*削除しませんでした*" -and $color -eq "Yellow" }
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

    It "サブフォルダも含めて Office・テキスト・PDF の拡張子のファイルだけを返す（.py・.cpp・.pdf も対象、.exe・.ts は対象外、大文字の拡張子も含め、~$ で始まるロックファイルは除く）" {
        $scan = findTargetFiles $source
        @($scan.Files | ForEach-Object { $_.Name } | Sort-Object) -join "," | Should -Be "a.xlsx,b.DOCX,c.pptm,d.txt,e.pdf,f.xls,g.ppt,h.py,i.cpp"
        $scan.HasError | Should -Be $false
    }

    It "末尾に \ を付けたフォルダでも同じフォルダを列挙する" {
        $scan = findTargetFiles "$source\"
        $scan.Root.TrimEnd("\") | Should -Be $source
        $scan.Files.Count | Should -Be 9
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

    It "新しい形式のほかのワークスペース（<目印>）は、ワークスペース直下のファイルも store の中も外れる" -TestCases @(
        @{ 目印 = "store\ingest_status.tsv"; 内容 = "dummy" }
        @{ 目印 = "store\format.tsv"; 内容 = "版=2`r`n" }
    ) {
        param ($目印, $内容)
        $workspace = newTestWorkspace @{} (Join-Path $TestDrive "ws-out-of-tree4")
        newFile "other-ws\$目印" $内容
        newFile "other-ws\store\content\営業\source_folder.txt"
        newFile "other-ws\store\content\営業\sheet.tsv"
        newFile "other-ws\search_results.txt"
        newFile "other-ws\indexing_log.txt"
        newFile "other-ws\system_index_state.tsv"
        newFile "other-ws\my-file.txt"

        $scan = findTargetFiles $source
        $names = @($scan.Files | ForEach-Object { $_.Name })
        $names | Should -Contain "my-file.txt"
        # 前の版が書いた search_results.txt は、今の版は書かないので、利用者のファイルとして扱う
        $names | Should -Contain "search_results.txt"
        foreach ($excluded in "source_folder.txt", "sheet.tsv", "indexing_log.txt", "system_index_state.tsv") {
            $names | Should -Not -Contain $excluded
        }
    }

    It "利用者の format.tsv（<場所>）は、ワークスペースの目印にならず、同じ場所の index・tmp・store の中のファイルを取り込む" -TestCases @(
        @{ 場所 = "X直下"; 目録 = "X\format.tsv"; 内容 = "版=2`r`n" }
        @{ 場所 = "X直下（中身は任意）"; 目録 = "X\format.tsv"; 内容 = "a`tb`r`n" }
        @{ 場所 = "store の中の目録でないもの"; 目録 = "X\store\format.tsv"; 内容 = "a`tb`r`n" }
        @{ 場所 = "store の中の壊れたもの"; 目録 = "X\store\format.tsv"; 内容 = "" }
    ) {
        param ($場所, $目録, $内容)
        $workspace = newTestWorkspace @{} (Join-Path $TestDrive "ws-out-of-tree5")
        newFile $目録 $内容
        newFile "X\index\a.txt"
        newFile "X\tmp\b.txt"
        newFile "X\store\c.txt"

        $names = @((findTargetFiles $source).Files | ForEach-Object { $_.Name })
        foreach ($kept in "a.txt", "b.txt", "c.txt") {
            $names | Should -Contain $kept
        }
    }

    It "読めない（開いたままで共有できない）X\store\format.tsv があっても、例外にせず、同じ場所のファイルを取り込む" {
        $workspace = newTestWorkspace @{} (Join-Path $TestDrive "ws-out-of-tree6")
        newFile "X\store\format.tsv" "版=2`r`n"
        newFile "X\index\a.txt"
        $stream = New-Object System.IO.FileStream((Join-Path $source "X\store\format.tsv"), [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        try {
            $names = @((findTargetFiles $source).Files | ForEach-Object { $_.Name })
        } finally {
            $stream.Dispose()
        }

        $names | Should -Contain "a.txt"
    }

    It "名前で分かる本文インデックス・前の版の本文インデックスのファイル・システムインデックスは、どこにあっても外れる" {
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

Describe "findTargetFiles・createTargetList（リンク）" -Tag Io {
    BeforeEach {
        $base = Join-Path $TestDrive "link-$(New-Guid)"
        $source = Join-Path $base "対象"
        $outside = Join-Path $base "外"
        [System.IO.Directory]::CreateDirectory($source) | Out-Null
        [System.IO.Directory]::CreateDirectory($outside) | Out-Null
        [System.IO.File]::WriteAllText("$source\a.txt", "dummy")
        [System.IO.File]::WriteAllText("$outside\外.txt", "dummy")
        $workspace = newTestWorkspace @{} (Join-Path $TestDrive "ws-out-of-tree-link-$(New-Guid)")
        ${indexDir} = Join-Path $TestDrive "index-link-$(New-Guid)"
        $workspace = newTestWorkspace @{ IndexDir = ${indexDir} }
        $script:logged = New-Object System.Collections.Generic.List[string]
        Mock writeIndexerLog { $script:logged.Add([string]$text) }
    }

    It "中のジャンクションの先のファイルは Files に入れず、Links に入れる" {
        if (!(newTestJunction "$source\リンク" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        $scan = findTargetFiles $source
        @($scan.Files | ForEach-Object { $_.Name }) | Should -Be @("a.txt")
        @($scan.Links) | Should -Be @("$source\リンク")
        $scan.HasError | Should -Be $false
    }

    It "親を指すジャンクション（輪）があっても終わり、Files に同じファイルが重ならない" {
        if (!(newTestJunction "$source\輪" $base)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        $scan = findTargetFiles $source
        @($scan.Files | ForEach-Object { $_.Name }) | Should -Be @("a.txt")
        @($scan.Links) | Should -Be @("$source\輪")
    }

    It "ファイルのシンボリックリンクは Files に入れず、Links に入れる（作れなければ飛ばす）" {
        if (!(newTestSymlink "$source\リンク.txt" "$outside\外.txt" -file)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        $scan = findTargetFiles $source
        @($scan.Files | ForEach-Object { $_.Name }) | Should -Be @("a.txt")
        @($scan.Links) | Should -Be @("$source\リンク.txt")
    }

    It "リンクかどうか確かめられないフォルダ（クラウドのオンデマンドのフォルダなど）は、これまでどおり中も取り込む" {
        if (!(newTestJunction "$source\不明" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        Mock getLinkKind { return "Unknown" }

        $scan = findTargetFiles $source
        @($scan.Files | ForEach-Object { $_.Name } | Sort-Object) | Should -Be @("a.txt", "外.txt")
        @($scan.Links).Count | Should -Be 0
    }

    It "クロール対象フォルダそのものがジャンクションなら、これまでどおり先の中を読み、中のジャンクションだけを Links に入れる" {
        if (!(newTestJunction "$base\根" $outside) -or !(newTestJunction "$outside\内リンク" $source)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        $scan = findTargetFiles "$base\根"
        @($scan.Files | ForEach-Object { $_.Name }) | Should -Be @("外.txt")
        $scan.HasError | Should -Be $false
        @($scan.Links) | Should -Be @("$base\根\内リンク")
    }

    It "createTargetList は、中のリンクの先に前回の行があるとき、元のファイルが無くなったものとして Removed に入れる" {
        if (!(newTestJunction "$source\リンク" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        $previous = newPrevious @(
            (newStatusRow "売上\リンク\外.txt" "2024/01/01 00:00:00" "10" ${stateDone} 1 "2024/01/01 00:00:00")
        )

        $result = createTargetList @{ Path = $source; Name = "売上" } $previous $null
        @($result.Removed) | Should -Be @("売上\リンク\外.txt")
        @($result.Rows | Where-Object { $_.相対パス -eq "売上\リンク\外.txt" }).Count | Should -Be 0
    }

    It "createTargetList は、飛ばしたリンクを 1 件 1 行でログに出す" {
        if (!(newTestJunction "$source\リンク1" $outside) -or !(newTestJunction "$source\リンク2" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        $result = createTargetList @{ Path = $source; Name = "売上" } (newPrevious) $null
        $result.Rows.Count | Should -Be 1
        @($script:logged | Where-Object { $_ -like "*リンク*のため取り込みませんでした: *" }).Count | Should -Be 2
        @($script:logged | Where-Object { $_ -like "*のため取り込みませんでした: $source\リンク1" }).Count | Should -Be 1
        @($script:logged | Where-Object { $_ -like "*のため取り込みませんでした: $source\リンク2" }).Count | Should -Be 1
    }

    It "createTargetList は、クロール対象フォルダがリンクのとき、先の中のファイルを取り込み対象にし、リンクを理由にしたログを出さない" {
        if (!(newTestJunction "$base\根" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        $result = createTargetList @{ Path = "$base\根"; Name = "売上" } (newPrevious) $null
        @($result.Targets | ForEach-Object { $_.相対パス }) | Should -Be @("売上\外.txt")
        @($result.Removed).Count | Should -Be 0
        @($script:logged | Where-Object { $_ -like "*リンク*のため取り込みませんでした*" }).Count | Should -Be 0
        @($script:logged | Where-Object { $_ -like "*アクセスできないフォルダ*" }).Count | Should -Be 0
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

Describe "IngestPlanner" -Tag Io {
    BeforeAll {
        $kinds = @{ Unchecked = ${planKindUnchecked}; Missing = ${planKindMissing}; Dropped = ${planKindDropped}; StateFailed = ${stateFailed} }
        # 名前を保存する設定ファイル（$TestDrive の下。実機の設定には書かない）
        $settingsFile = Join-Path $TestDrive "planner-setting.config"
        ${indexDir} = Join-Path $TestDrive "planner-index"
        $workspace = newTestWorkspace @{ IndexDir = ${indexDir} } (Join-Path $TestDrive "planner-ws")

        $source = Join-Path $TestDrive "planner-src"
        [System.IO.Directory]::CreateDirectory($source) | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $source "a.xlsx"), "dummy")
        $info = Get-Item -LiteralPath (Join-Path $source "a.xlsx")
        $updated = formatFileTime $info.LastWriteTime
        $size = [string]$info.Length

        function newSaver {
            # 割り当てた名前を設定ファイルに保存する処理（本物は indexer_run.ps1 が IndexCatalog を通して渡す）
            param ($file)
            $catalog = [IndexCatalog]::new($workspace, $file, [StatusLedger]::new($workspace))
            return { param($assigned) $catalog.StoreAssignedNames($assigned) }.GetNewClosure()
        }

        function newPlanner {
            param ($folders, $previousRows = @(), $previousFolders = @(), $file = $settingsFile)
            $planner = [IngestPlanner]::new($kinds)
            $status = @{ Folders = $previousFolders; Rows = (newPrevious $previousRows) }
            $planner.AssignNames($folders, $status)
            return $planner
        }
    }

    It "名前の無いフォルダに名前を割り当て、前回の一覧に無くなったインデックスを外れたものとして見つける" {
        $planner = newPlanner @([pscustomobject]@{ Path = $source; Name = ""; Enabled = $true }) @() @(@{ Path = "C:\古い"; Name = "古い" })
        $planner.Folders.Count | Should -Be 1
        $planner.Folders[0].Name | Should -Not -BeNullOrEmpty
        $planner.Folders[0].Path | Should -Be $source
        @($planner.Dropped | ForEach-Object { $_.Name }) | Should -Be @("古い")
    }

    It "名前が今の決まりに合わない前の版のインデックス（LegacyFolders）も、外れたものとして見つける。LegacyFolders が無い状態でも動く" {
        $planner = [IngestPlanner]::new($kinds)
        $status = @{ Folders = @(@{ Path = "C:\古い"; Name = "古い" }); Rows = (newPrevious @()); LegacyFolders = @(@{ Path = "C:\営業"; Name = "営業." }) }
        $planner.AssignNames(@([pscustomobject]@{ Path = $source; Name = "売上"; Enabled = $true }), $status)
        @($planner.Dropped | ForEach-Object { $_.Name }) | Should -Be @("古い", "営業.")

        $planner = [IngestPlanner]::new($kinds)
        $planner.AssignNames(@([pscustomobject]@{ Path = $source; Name = "売上"; Enabled = $true }), @{ Folders = @(); Rows = (newPrevious @()) })
        $planner.Dropped.Count | Should -Be 0
    }

    It "全部に名前があれば、設定ファイルを書かない" {
        $named = @([pscustomobject]@{ Path = $source; Name = "売上"; Enabled = $true })
        $file = Join-Path $TestDrive "保存なし\setting.config"
        $planner = newPlanner $named @() @() $file
        $planner.SaveNames($named, (newSaver $file))
        Test-Path -LiteralPath $file | Should -Be $false
        @($planner.Folders | ForEach-Object { $_.Name }) | Should -Be @("売上")
    }

    Context "名前の無いフォルダがあるとき、割り当てた名前を設定ファイルに保存する" {
        BeforeEach {
            $file = Join-Path $TestDrive "保存\$([guid]::NewGuid())\setting.config"
            writeTargetFolders @(
                [pscustomobject]@{ Name = ""; Path = "C:\営業"; Enabled = $true }
                [pscustomobject]@{ Name = ""; Path = "C:\経理"; Enabled = $true }
            ) $file
            $read = @(getTargetFolders $file)
            $planner = newPlanner $read @() @() $file
        }

        It "保存した名前を設定ファイルに書き、Folders をその一覧にする" {
            $planner.SaveNames($read, (newSaver $file))
            $saved = @(getTargetFolders $file)
            @($saved | ForEach-Object { $_.Name }) | Should -Not -Contain ""
            @($planner.Folders | ForEach-Object { $_.Name }) | Should -Be @($saved | ForEach-Object { $_.Name })
            @($planner.Folders | ForEach-Object { $_.Path }) | Should -Be @("C:\営業", "C:\経理")
        }

        It "割り当てから保存までの間に設定から外したフォルダは、Folders から落とす" {
            writeTargetFolders @([pscustomobject]@{ Name = ""; Path = "C:\営業"; Enabled = $true }) $file
            $planner.SaveNames($read, (newSaver $file))
            @($planner.Folders | ForEach-Object { $_.Path }) | Should -Be @("C:\営業")
        }

        It "名前を付けられなかったフォルダ（同じ名前を、ほかが先に使った）は、Folders から落とす" {
            $taken = $planner.Folders[0].Name
            writeTargetFolders @(
                [pscustomobject]@{ Name = ""; Path = "C:\営業"; Enabled = $true }
                [pscustomobject]@{ Name = $taken; Path = "C:\総務"; Enabled = $true }
            ) $file
            $planner.SaveNames($read, (newSaver $file))
            @($planner.Folders | ForEach-Object { $_.Path }) | Should -Not -Contain "C:\営業"
        }
    }

    It "kinds に足りない値があれば、作るときに例外にする" -TestCases @(
        @{ missing = "Unchecked" }
        @{ missing = "Missing" }
        @{ missing = "Dropped" }
        @{ missing = "StateFailed" }
    ) {
        param ($missing)
        $partial = @{} + $kinds
        $partial.Remove($missing)
        { [IngestPlanner]::new($partial) } | Should -Throw "*$missing*"
    }

    It "AssignNames の前に Crawl を呼ぶと例外にする" {
        $planner = [IngestPlanner]::new($kinds)
        { $planner.Crawl($null, $null) } | Should -Throw "*AssignNames*"
    }

    It "進み具合を出す処理の例外は、包まずにそのまま伝える" {
        $folders = @([pscustomobject]@{ Path = $source; Name = "売上"; Enabled = $true })
        $planner = newPlanner $folders
        { $planner.Crawl((newCounts), { param ($target) throw "進み具合の失敗" }) } | Should -Throw "進み具合の失敗"
    }

    It "元のファイルが無くなったファイルは、RemovedPaths に積む" {
        $folders = @([pscustomobject]@{ Path = $source; Name = "売上"; Enabled = $true })
        $gone = newStatusRow "売上\消えた.xlsx" "2000/01/01 00:00:00" "1" ${stateDone}
        $planner = newPlanner $folders @($gone)
        $planner.Crawl((newCounts @{ "売上\消えた.xlsx" = 1 }), $null)
        @($planner.RemovedPaths) | Should -Be @("売上\消えた.xlsx")
    }

    It "作るときのハッシュテーブルの TextExcludes（インデックス名 → 取り込まない拡張子）を使う。ほかのインデックスには効かない。無ければ全部取り込む" {
        $dirA = Join-Path $TestDrive "planner-ext-a"
        $dirB = Join-Path $TestDrive "planner-ext-b"
        foreach ($dir in @($dirA, $dirB)) {
            [System.IO.Directory]::CreateDirectory($dir) | Out-Null
            [System.IO.File]::WriteAllText((Join-Path $dir "a.log"), "ログ")
            [System.IO.File]::WriteAllText((Join-Path $dir "b.txt"), "メモ")
        }
        $folders = @(
            [pscustomobject]@{ Path = $dirA; Name = "営業"; Enabled = $true }
            [pscustomobject]@{ Path = $dirB; Name = "経理"; Enabled = $true }
        )
        $withKinds = @{} + $kinds
        $withKinds.TextExcludes = @{ "営業" = @(".log") }
        $planner = [IngestPlanner]::new($withKinds)
        $planner.AssignNames($folders, @{ Folders = @(); Rows = (newPrevious @()) })
        $planner.Crawl((newCounts), $null)
        @($planner.Targets | ForEach-Object { $_.相対パス } | Sort-Object) | Should -Be @("営業\b.txt", "経理\a.log", "経理\b.txt")
        $none = newPlanner $folders @()
        $none.Crawl((newCounts), $null)
        @($none.Targets).Count | Should -Be 4
    }

    It "TextExcludes を先に入れたときは、設定を読まずにそれを使う（同じ名前の複数のフォルダは同じ一覧になる）" {
        $dirA = Join-Path $TestDrive "planner-ext-a"
        $dirB = Join-Path $TestDrive "planner-ext-b"
        $folders = @(
            [pscustomobject]@{ Path = $dirA; Name = "営業"; Enabled = $true }
            [pscustomobject]@{ Path = $dirB; Name = "営業"; Enabled = $true }
        )
        $planner = newPlanner $folders @()
        $planner.TextExcludes = @{ "営業" = @(".txt") }
        $planner.Crawl((newCounts), $null)
        @($planner.Targets | ForEach-Object { [System.IO.Path]::GetFileName($_.相対パス) }) | Should -Be @("a.log", "a.log")
    }

    It "インデックス名ごとの拡張子ごとの件数（取り込まない拡張子の分も。同じ名前の複数のフォルダは合計）を持つ" {
        $dirA = Join-Path $TestDrive "planner-ext-a"
        $folders = @([pscustomobject]@{ Path = $dirA; Name = "営業"; Enabled = $true })
        $folders += [pscustomobject]@{ Path = (Join-Path $TestDrive "planner-ext-b"); Name = "営業"; Enabled = $true }
        $planner = newPlanner $folders @()
        $planner.TextExcludes = @{ "営業" = @(".txt") }
        $planner.Crawl((newCounts), $null)
        $counts = $planner.TextCountsByIndex["営業"]
        $counts[".txt"] | Should -Be 2
        $counts[".log"] | Should -Be 2
        $counts.Count | Should -Be 2
        $planner.TextCountsByIndex.Count | Should -Be 1
    }

    It "取りやめるときは、元のファイルが無くなったファイルの前回の行も残す（インデックスは消していないので、次の回にまた無くなったファイルとして見つける）" {
        $folders = @([pscustomobject]@{ Path = $source; Name = "売上"; Enabled = $true })
        $gone = newStatusRow "売上\消えた.xlsx" "2000/01/01 00:00:00" "1" ${stateDone}
        $planner = newPlanner $folders @($gone)
        $planner.Crawl((newCounts @{ "売上\消えた.xlsx" = 1 }), $null)
        $keep = @($planner.KeepForCancel())
        @($keep | ForEach-Object { $_.相対パス }) | Should -Be @("売上\消えた.xlsx")
        [object]::ReferenceEquals($keep[0], $gone) | Should -Be $true
        @($planner.Rows | ForEach-Object { $_.相対パス }) | Should -Not -Contain "売上\消えた.xlsx"
    }

    It "クロールで、取り込み対象・チェックなし・フォルダなし・外れたインデックスを画面に出す件数に分ける" {
        $folders = @(
            [pscustomobject]@{ Path = $source; Name = "売上"; Enabled = $true }
            [pscustomobject]@{ Path = (Join-Path $TestDrive "x"); Name = "外し"; Enabled = $false }
            [pscustomobject]@{ Path = (Join-Path $TestDrive "無い"); Name = "無い"; Enabled = $true }
        )
        $kept = newStatusRow "外し\z.xlsx" "2000/01/01 00:00:00" "1" ${stateDone}
        $planner = newPlanner $folders @($kept) @(@{ Path = "C:\古い"; Name = "古い" })
        $script:seen = @()
        $planner.Crawl((newCounts), { param ($target) $script:seen += $target.Name })

        @($planner.Plan | ForEach-Object { $_.区分 }) | Should -Be @(${planKindIngest}, ${planKindUnchecked}, ${planKindMissing}, ${planKindDropped})
        $planner.Targets.Count | Should -Be 1
        $planner.Targets[0].相対パス | Should -Be "売上\a.xlsx"
        # チェックを外したフォルダは、前回の行をそのまま残す
        @($planner.Rows | ForEach-Object { $_.相対パス }) | Should -Be @("売上\a.xlsx", "外し\z.xlsx")
        $script:seen | Should -Be @("売上")
    }

    It "選んだインデックスだけを更新するとき、選ばなかったインデックスは前回の行を残して確認の表にも出さない" {
        $folders = @(
            [pscustomobject]@{ Path = $source; Name = "売上"; Enabled = $true }
            [pscustomobject]@{ Path = (Join-Path $TestDrive "x"); Name = "別"; Enabled = $true }
        )
        $other = newStatusRow "別\z.xlsx" "2000/01/01 00:00:00" "1" ${stateDone}
        $planner = newPlanner $folders @($other)
        $skipped = @($planner.SelectOnly([string[]]@("売上", "総務")))
        @($skipped | ForEach-Object { "$($_.Name):$($_.Reason)" }) | Should -Be @("総務:設定にありません")
        $planner.Only.Selected.Count | Should -Be 1
        $planner.Crawl($null, $null)
        @($planner.Plan | ForEach-Object { $_.インデックス名 }) | Should -Be @("売上")
        @($planner.Rows | ForEach-Object { $_.相対パス }) | Should -Be @("売上\a.xlsx", "別\z.xlsx")
    }

    It "取りやめるときは、取り込み対象の行を前回の記録のまま（一覧に無かった行は書かない）にする" {
        $folders = @([pscustomobject]@{ Path = $source; Name = "売上"; Enabled = $true })
        $old = newStatusRow "売上\a.xlsx" "2000/01/01 00:00:00" $size ${stateDone}
        $withPrevious = newPlanner $folders @($old)
        $withPrevious.Crawl((newCounts @{ "売上\a.xlsx" = 1 }), $null)
        $keep = @($withPrevious.KeepForCancel())
        $keep.Count | Should -Be 1
        [object]::ReferenceEquals($keep[0], $old) | Should -Be $true

        $firstTime = newPlanner $folders
        $firstTime.Crawl($null, $null)
        @($firstTime.KeepForCancel()).Count | Should -Be 0
    }

    It "クロールの途中で中止の合図が立ったら、残りのフォルダは前回の行を残し、Stopped を立てる" {
        $source2 = Join-Path $TestDrive "planner-src2"
        [System.IO.Directory]::CreateDirectory($source2) | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $source2 "b.xlsx"), "dummy")
        $folders = @(
            [pscustomobject]@{ Path = $source; Name = "売上"; Enabled = $true }
            [pscustomobject]@{ Path = $source2; Name = "経理"; Enabled = $true }
        )
        $old = newStatusRow "経理\z.xlsx" "2000/01/01 00:00:00" "1" ${stateDone}
        $planner = newPlanner $folders @($old)
        # 1 つ目のフォルダを出し終えたら中止の合図を立てる
        $script:stopNow = $false
        $planner.Crawl((newCounts), { param ($target) $script:stopNow = $true }, { $script:stopNow })

        $planner.Stopped | Should -Be $true
        # 2 つ目のフォルダは見ず、前回の行をそのまま残す
        @($planner.Rows | ForEach-Object { $_.相対パス }) | Should -Contain "経理\z.xlsx"
        @($planner.Rows | ForEach-Object { $_.相対パス }) | Should -Not -Contain "経理\b.xlsx"
        @($planner.RemovedPaths).Count | Should -Be 0
    }

    It "中止の合図が立っていなければ Stopped は偽。最初から立っていれば、どのフォルダも見ない" {
        $folders = @([pscustomobject]@{ Path = $source; Name = "売上"; Enabled = $true })
        $planner = newPlanner $folders
        $planner.Crawl($null, $null, { $false })
        $planner.Stopped | Should -Be $false
        $planner.Targets.Count | Should -Be 1

        $old = newStatusRow "売上\a.xlsx" "2000/01/01 00:00:00" $size ${stateDone}
        $planner = newPlanner $folders @($old)
        $planner.Crawl($null, $null, { $true })
        $planner.Stopped | Should -Be $true
        $planner.Targets.Count | Should -Be 0
        @($planner.Rows | ForEach-Object { $_.相対パス }) | Should -Be @("売上\a.xlsx")
    }

    It "前回の失敗は、再取り込みを選べば取り込み対象に足し、選ばなければ足さない" -TestCases @(
        @{ retry = $true; count = 1 }
        @{ retry = $false; count = 0 }
    ) {
        param ($retry, $count)
        $folders = @([pscustomobject]@{ Path = $source; Name = "売上"; Enabled = $true })
        $failedRow = newStatusRow "売上\a.xlsx" $updated $size ${stateFailed} "" "" "エラー" "4"
        $planner = newPlanner $folders @($failedRow)
        $planner.Crawl((newCounts @{ "売上\a.xlsx" = 1 }), $null)
        $planner.Failed.Count | Should -Be 1
        $planner.Targets.Count | Should -Be 0
        $planner.ApplyFailed($retry)
        $planner.Targets.Count | Should -Be $count
    }

    It "強制終了した回数が上限に達していなければ最後に回し、達していれば失敗にして外す" {
        $planner = [IngestPlanner]::new($kinds)
        $first = newStatusRow "売上\1.xlsx"
        $second = newStatusRow "売上\2.xlsx"
        $third = newStatusRow "売上\3.xlsx"
        $planner.Targets.AddRange([object[]]@($first, $second, $third))
        $planner.ApplyInterrupted(@(@{ RelPath = "売上\1.xlsx"; Count = 1 }, @{ RelPath = "売上\2.xlsx"; Count = 2 }, @{ RelPath = "売上\無い.xlsx"; Count = 1 }), 2)

        @($planner.Targets | ForEach-Object { $_.相対パス }) | Should -Be @("売上\3.xlsx", "売上\1.xlsx")
        $planner.Carried.Count | Should -Be 1
        $planner.Carried["売上\1.xlsx"] | Should -Be 1
        $second.状態 | Should -Be ${stateFailed}
        $second.エラー | Should -BeLike "*2 回続けて強制終了*"
    }

    It "インデックス名からクロール対象フォルダを引ける（末尾の \ は外す）" {
        $planner = newPlanner @([pscustomobject]@{ Path = "C:\共有\営業部\"; Name = "営業"; Enabled = $true })
        $byName = $planner.FolderByName()
        $byName["営業"] | Should -Be "C:\共有\営業部"
        $byName["営業".ToUpper()] | Should -Be "C:\共有\営業部"
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
        [System.IO.Directory]::CreateDirectory($bookDir) | Out-Null
        $result = createTargetList $folder (newPrevious @($pendingRow)) $null
        # ここでは消さない（消すのは取り込みを決める側。取りやめたときは残す）
        [System.IO.Directory]::Exists($bookDir) | Should -Be $true
        $result.Rows.Count | Should -Be 0
        $result.Targets.Count | Should -Be 0
        $result.Cloud.Count | Should -Be 1
        $result.Cloud[0].PendingOld | Should -Be $true
    }
}

Describe "IngestPlanner（クラウドにだけあるファイル）" -Tag Io {
    BeforeAll {
        $kinds = @{ Unchecked = ${planKindUnchecked}; Missing = ${planKindMissing}; Dropped = ${planKindDropped}; StateFailed = ${stateFailed} }
        # Crawl は設定（取り込まない拡張子）を読むため、$TestDrive の下の（無い）設定ファイルを指す
        $settingsFile = Join-Path $TestDrive "planner-cloud-setting.config"
        $source = Join-Path $TestDrive "planner_cloud_src"
        [System.IO.Directory]::CreateDirectory($source) | Out-Null
        $file = Join-Path $source "c.xlsx"
        [System.IO.File]::WriteAllText($file, "dummy-cloud")
        # クラウドにだけあるファイルは、OFFLINE 属性で代わりに作る（ダウンロードは起きない）
        [System.IO.File]::SetAttributes($file, [System.IO.FileAttributes]::Offline)
        $info = Get-Item -LiteralPath $file
        $updated = formatFileTime $info.LastWriteTime
        $size = [string]$info.Length
        ${indexDir} = Join-Path $TestDrive "planner_cloud_index"
        $workspace = newTestWorkspace @{ IndexDir = ${indexDir} }

        function newCloudPlanner {
            param ($previousRows)
            $planner = [IngestPlanner]::new($kinds)
            $planner.AssignNames(@([pscustomobject]@{ Path = $source; Name = "雲"; Enabled = $true }), @{ Folders = @(); Rows = (newPrevious $previousRows) })
            $planner.Crawl($null, $null)
            return $planner
        }
    }

    AfterAll {
        [System.IO.File]::SetAttributes($file, [System.IO.FileAttributes]::Normal)
    }

    It "Crawl が、クラウドにだけあるファイルを Cloud に集め、取り込み対象には入れない" {
        $planner = newCloudPlanner @()
        $planner.Cloud.Count | Should -Be 1
        $planner.Cloud[0].RelPath | Should -Be "雲\c.xlsx"
        $planner.Targets.Count | Should -Be 0
    }

    It "取りやめるとき、クラウドにだけあって前回未完了のファイルは、前回の行を残す" {
        $pendingRow = newStatusRow "雲\c.xlsx" $updated $size ${stateNew}
        $planner = newCloudPlanner @($pendingRow)
        $planner.Cloud[0].PendingOld | Should -Be $true
        $keep = @($planner.KeepForCancel())
        $keep.Count | Should -Be 1
        [object]::ReferenceEquals($keep[0], $pendingRow) | Should -Be $true
    }
}

Describe "getTargetExtensionList / 取り込まない拡張子（createTargetList・findTargetFiles）" -Tag Io {
    BeforeAll {
        $source = Join-Path $TestDrive "src_ext"
        [System.IO.Directory]::CreateDirectory($source) | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $source "a.log"), "ログ")
        [System.IO.File]::WriteAllText((Join-Path $source "b.txt"), "メモ")
        [System.IO.File]::WriteAllText((Join-Path $source "c.xlsx"), "dummy")
        $folder = @{ Path = $source; Name = "拡張子" }
        ${indexDir} = Join-Path $TestDrive "index_ext"
        $workspace = newTestWorkspace @{ IndexDir = ${indexDir} }
    }

    BeforeEach {
        Mock writeIndexerLog {}
    }

    AfterEach {
        if (Test-Path -LiteralPath ${indexDir}) { Remove-Item -LiteralPath ${indexDir} -Recurse -Force }
    }

    It "getTargetExtensionList は、Office・PDF とテキスト（取り込まないものを除く）を返す" {
        @(getTargetExtensionList).Count | Should -Be 86
        $some = @(getTargetExtensionList @(".log"))
        $some.Count | Should -Be 85
        $some | Should -Not -Contain ".log"
        $some | Should -Contain ".txt"
        $some | Should -Contain ".xlsx"
        $some | Should -Contain ".pdf"
    }

    It "一覧を渡さなければ今までどおり（.log も対象）、.log を除いた一覧を渡すと .log を返さない" {
        ((@((findTargetFiles $source).Files | ForEach-Object { $_.Name }) | Sort-Object) -join ",") | Should -Be "a.log,b.txt,c.xlsx"
        $result = findTargetFiles $source $null (getTargetExtensionList @(".log"))
        ((@($result.Files | ForEach-Object { $_.Name }) | Sort-Object) -join ",") | Should -Be "b.txt,c.xlsx"
    }

    It "前に取り込んだ .log は、取り込まないにすると Removed に入り（インデックスはまだ消さない）、ログが分かれる" {
        $logBook = Join-Path ${indexDir} "拡張子\a.log"
        [System.IO.Directory]::CreateDirectory($logBook) | Out-Null
        [System.IO.File]::WriteAllText("$logBook\1.tsv", "1`t見積")
        $previous = newPrevious @((newStatusRow "拡張子\a.log" "2024/01/01 00:00:00" "10" ${stateDone} 1 "2024/01/01 00:00:00"))
        $result = createTargetList $folder $previous $null $null (getTargetExtensionList @(".log"))
        @($result.Removed) | Should -Be @("拡張子\a.log")
        @($result.Rows | Where-Object { $_.相対パス -eq "拡張子\a.log" }).Count | Should -Be 0
        # ここでは消さない（取り込みを始めると決まったあとに呼び出し元が消す）
        Test-Path -LiteralPath $logBook | Should -Be $true
        Should -Invoke writeIndexerLog -Times 1 -Exactly -ParameterFilter { $text -like "*対象外にした拡張子のファイル 1 件は、インデックスから除きます*" }
        Should -Invoke writeIndexerLog -Times 0 -Exactly -ParameterFilter { $text -like "*元ファイルが無くなった*" }
    }

    It "元のファイルが無くなったものは、これまでの文言のまま（取り込まないにした件数に数えない）" {
        $previous = newPrevious @((newStatusRow "拡張子\消えた.txt" "2024/01/01 00:00:00" "10" ${stateDone} 1 "2024/01/01 00:00:00"))
        $result = createTargetList $folder $previous $null $null (getTargetExtensionList @(".log"))
        @($result.Removed) | Should -Be @("拡張子\消えた.txt")
        Should -Invoke writeIndexerLog -Times 1 -Exactly -ParameterFilter { $text -like "*元ファイルが無くなった 1 件は、インデックスから除きます*" }
        Should -Invoke writeIndexerLog -Times 0 -Exactly -ParameterFilter { $text -like "*取り込まないにした拡張子*" }
    }

    It "拡張子ごとの件数は、取り込まない拡張子の分も数え、Office・PDF も数え、tebunko が作ったファイルは数えない" {
        $countSource = Join-Path $TestDrive "src_counts"
        [System.IO.Directory]::CreateDirectory($countSource) | Out-Null
        foreach ($name in @("a.log", "b.LOG", "c.txt", "d.xlsx", "e.pdf", "f.zip", "content_index.xlsx.001.tsv", "system_index_state.tsv")) {
            [System.IO.File]::WriteAllText((Join-Path $countSource $name), "x")
        }
        $result = findTargetFiles $countSource $null (getTargetExtensionList @(".log"))
        $result.TextCounts[".log"] | Should -Be 2
        $result.TextCounts[".txt"] | Should -Be 1
        $result.TextCounts[".xlsx"] | Should -Be 1
        $result.TextCounts[".pdf"] | Should -Be 1
        $result.TextCounts.ContainsKey(".zip") | Should -BeFalse
        # tebunko が作った TSV は、拡張子 .tsv の件数に入らない
        $result.TextCounts.ContainsKey(".tsv") | Should -BeFalse
        # 取り込む一覧で絞るので、取り込まない .log は Files に出ない
        (@($result.Files | ForEach-Object { $_.Name } | Sort-Object) -join ",") | Should -Be "c.txt,d.xlsx,e.pdf"
    }

    It "取り込むに戻す（一覧に .log を入れる）と、.log は新規として数える" {
        $result = createTargetList $folder (newPrevious) $null $null (getTargetExtensionList @())
        @($result.Targets | Where-Object { $_.相対パス -eq "拡張子\a.log" }).Count | Should -Be 1
        @($result.Removed).Count | Should -Be 0
    }
}

Describe "writeRemovedTargetLog" -Tag Unit {
    BeforeEach {
        Mock writeIndexerLog {}
    }

    It "<name>" -TestCases @(
        @{ name = "外したものが 0 件なら何も出さない"; removed = @(); extensions = @(".txt"); excluded = 0; gone = 0 }
        @{ name = "拡張子の一覧が無ければ、すべて「無くなった」"; removed = @("a\x.log", "a\y.txt"); extensions = $null; excluded = 0; gone = 2 }
        @{ name = "一覧に無い拡張子 1 件は「取り込まないにした」"; removed = @("a\x.log"); extensions = @(".txt"); excluded = 1; gone = 0 }
        @{ name = "一覧にある拡張子 1 件は「無くなった」"; removed = @("a\y.txt"); extensions = @(".txt"); excluded = 0; gone = 1 }
        @{ name = "複数件は、それぞれの件数で分ける（大文字の拡張子も一覧と同じに見る）"; removed = @("a\1.log", "a\2.LOG", "a\3.txt", "a\4.TXT", "a\5.txt"); extensions = @(".txt"); excluded = 2; gone = 3 }
    ) {
        param ($name, $removed, $extensions, $excluded, $gone)
        writeRemovedTargetLog $removed $extensions
        $excludedTimes = [int]($excluded -gt 0)
        $goneTimes = [int]($gone -gt 0)
        Should -Invoke writeIndexerLog -Times $excludedTimes -Exactly -ParameterFilter { $text -eq "    対象外にした拡張子のファイル $excluded 件は、インデックスから除きます。" }
        Should -Invoke writeIndexerLog -Times $goneTimes -Exactly -ParameterFilter { $text -eq "    元ファイルが無くなった $gone 件は、インデックスから除きます。" }
        Should -Invoke writeIndexerLog -Times ($excludedTimes + $goneTimes) -Exactly
    }
}

Describe "getSubfolderRemovedCounts（直下だけに変えたとき、サブフォルダ側で外れる件数）" -Tag Unit {
    It "<label>" -TestCases @(
        @{ label = "直下だけに変えた（前回は再帰）フォルダは、サブフォルダ側の相対パスだけ数える（直下の無くなったファイルは数えない）"
           folders = @([pscustomobject]@{ Name = "経理"; Subfolders = $false; IngestedSubfolders = $true })
           removed = @("経理\2024\下.docx", "経理\2024\奥\深.docx", "経理\直下.docx"); expected = @{ "経理" = 2 } }
        @{ label = "前回の値が無い（再帰とみなす）ときも数える"
           folders = @([pscustomobject]@{ Name = "経理"; Subfolders = $false; IngestedSubfolders = $null })
           removed = @("経理\a\b.docx"); expected = @{ "経理" = 1 } }
        @{ label = "もともと直下だけ（前回も false）のフォルダは数えない"
           folders = @([pscustomobject]@{ Name = "経理"; Subfolders = $false; IngestedSubfolders = $false })
           removed = @("経理\a\b.docx"); expected = @{} }
        @{ label = "変えていない（含める）フォルダは数えない"
           folders = @([pscustomobject]@{ Name = "経理"; Subfolders = $true; IngestedSubfolders = $true })
           removed = @("経理\a\b.docx"); expected = @{} }
        @{ label = "名前が前方一致するほかのインデックスのパスは数えない"
           folders = @([pscustomobject]@{ Name = "経理"; Subfolders = $false; IngestedSubfolders = $true })
           removed = @("経理2\a\b.docx"); expected = @{ "経理" = 0 } }
    ) {
        param ($label, $folders, $removed, $expected)
        $counts = getSubfolderRemovedCounts $folders $removed
        $counts.Count | Should -Be $expected.Count
        foreach ($key in $expected.Keys) { $counts[$key] | Should -Be $expected[$key] }
    }
}

Describe "IngestPlanner の Plan の SubfolderRemovedCount" -Tag Io {
    BeforeAll {
        $kinds = @{ Unchecked = ${planKindUnchecked}; Missing = ${planKindMissing}; Dropped = ${planKindDropped}; StateFailed = ${stateFailed} }
        ${indexDir} = Join-Path $TestDrive "index_removed_count"
        $workspace = newTestWorkspace @{ IndexDir = ${indexDir} }
        $source = Join-Path $TestDrive "removed_count_src"
        [System.IO.Directory]::CreateDirectory("$source\下") | Out-Null
        [System.IO.File]::WriteAllText("$source\直下.docx", "dummy")
        [System.IO.File]::WriteAllText("$source\下\奥.docx", "dummy")
        function newCountPlanner([bool]$subfolders, $ingested) {
            $planner = [IngestPlanner]::new($kinds)
            $info = Get-Item -LiteralPath "$source\直下.docx"
            $status = @{ Folders = @(); Rows = (newPrevious @(
                (newStatusRow "経理\下\消えた.docx" "2000/01/01 00:00:00" "1" ${stateDone}),
                (newStatusRow "経理\下\奥\消えた2.docx" "2000/01/01 00:00:00" "1" ${stateDone}),
                (newStatusRow "経理\消えた直下.docx" "2000/01/01 00:00:00" "1" ${stateDone}))) }
            $planner.AssignNames(@([pscustomobject]@{ Name = "経理"; Path = $source; Enabled = $true; Subfolders = $subfolders; IngestedSubfolders = $ingested }), $status)
            $planner.Crawl((newCounts @{ "経理\下\消えた.docx" = 1; "経理\下\奥\消えた2.docx" = 1; "経理\消えた直下.docx" = 1 }), $null)
            return $planner
        }
    }

    It "直下だけに変えたフォルダの Plan の行に、サブフォルダ側の前回の行の数が入る（直下で本当に無くなったファイルは数えない）" {
        $planner = newCountPlanner $false $true
        @($planner.RemovedPaths).Count | Should -Be 3
        @($planner.Plan)[0].SubfolderRemovedCount | Should -Be 2
    }

    It "変えていないフォルダ（含める）の行は 0 になる" {
        $planner = newCountPlanner $true $true
        @($planner.Plan)[0].SubfolderRemovedCount | Should -Be 0
    }

    It "もともと直下だけ（前回も false）のフォルダの行は 0 になる" {
        $planner = newCountPlanner $false $false
        @($planner.Plan)[0].SubfolderRemovedCount | Should -Be 0
    }
}

Describe "createTargetList（直下だけ）" -Tag Io {
    BeforeAll {
        $source = Join-Path $TestDrive "top"
        [System.IO.Directory]::CreateDirectory("$source\2024") | Out-Null
        [System.IO.File]::WriteAllText("$source\直下.docx", "dummy")
        [System.IO.File]::WriteAllText("$source\2024\下.docx", "dummy")
        ${indexDir} = Join-Path $TestDrive "index_top"
        $workspace = newTestWorkspace @{ IndexDir = ${indexDir} }
        $settings = Join-Path $TestDrive "top_setting.json"
        # 設定から読み、名前を割り当てた形（インデクサが createTargetList に渡す形）で組む
        function newTopFolder([bool]$subfolders) {
            writeTargetFolders @([pscustomobject]@{ Name = "経理"; Path = $source; Enabled = $true; Subfolders = $subfolders }) $settings
            return @(assignIndexNames @(getTargetFolders $settings) @())[0]
        }
    }

    It "findTargetFiles -TopOnly は、サブフォルダのファイルを入れない" {
        @((findTargetFiles $source).Files | ForEach-Object { $_.Name } | Sort-Object) | Should -Be @("下.docx", "直下.docx")
        @((findTargetFiles $source -TopOnly).Files | ForEach-Object { $_.Name }) | Should -Be @("直下.docx")
    }

    It "Subfolders が <label> のとき、Files に入るのは <count> 件" -TestCases @(
        @{ label = "true"; value = $true; count = 2 }
        @{ label = "false"; value = $false; count = 1 }
    ) {
        param ($label, $value, $count)
        $result = createTargetList (newTopFolder $value) (newPrevious) $null
        $result.Rows.Count | Should -Be $count
        $result.Targets.Count | Should -Be $count
    }

    It "直下だけに変えると、前回取り込んだサブフォルダ側のファイルは Removed になる" {
        $before = createTargetList (newTopFolder $true) (newPrevious) $null
        $previous = newPrevious $before.Rows
        $result = createTargetList (newTopFolder $false) $previous $null
        @($result.Removed) | Should -Be @("経理\2024\下.docx")
        @($result.Rows | ForEach-Object { $_.相対パス }) | Should -Be @("経理\直下.docx")
    }
}

Describe "getTextExcludeRemovedCounts" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "取り込まない拡張子のファイルだけをインデックス名ごとに数える（大文字の拡張子も）"
           folders = @([pscustomobject]@{ Name = "営業"; Subfolders = $true; IngestedSubfolders = $true }, [pscustomobject]@{ Name = "経理"; Subfolders = $true; IngestedSubfolders = $true })
           removed = @("営業\a.log", "営業\b.LOG", "営業\c.txt", "経理\d.log"); excludes = @{ "営業" = @(".log"); "経理" = @() }; expected = @{ "営業" = 2 } }
        @{ name = "サブフォルダを含めないに変えたフォルダでは、サブフォルダ側のファイルは数えない（サブフォルダ側の 1 行に数える）"
           folders = @([pscustomobject]@{ Name = "営業"; Subfolders = $false; IngestedSubfolders = $true })
           removed = @("営業\a.log", "営業\sub\x.log"); excludes = @{ "営業" = @(".log") }; expected = @{ "営業" = 1 } }
        @{ name = "前から直下だけだったフォルダは、サブフォルダ側も数える"
           folders = @([pscustomobject]@{ Name = "営業"; Subfolders = $false; IngestedSubfolders = $false })
           removed = @("営業\a.log", "営業\sub\x.log"); excludes = @{ "営業" = @(".log") }; expected = @{ "営業" = 2 } }
        @{ name = "外すものが無ければ空"
           folders = @([pscustomobject]@{ Name = "営業"; Subfolders = $true; IngestedSubfolders = $true })
           removed = @("営業\a.txt"); excludes = @{ "営業" = @(".log") }; expected = @{} }
    ) {
        param ($name, $folders, $removed, $excludes, $expected)
        $result = getTextExcludeRemovedCounts $folders $removed $excludes
        $result.Count | Should -Be $expected.Count
        foreach ($key in $expected.Keys) { $result[$key] | Should -Be $expected[$key] }
    }
}

Describe "IngestPlanner の Plan の TextExtensionCounts・TextExcludeRemovedCount と Channel の CountedFiles" -Tag Io {
    BeforeAll {
        $kinds = @{ Unchecked = ${planKindUnchecked}; Missing = ${planKindMissing}; Dropped = ${planKindDropped}; StateFailed = ${stateFailed} }
        ${indexDir} = Join-Path $TestDrive "index_text_ext_plan"
        $workspace = newTestWorkspace @{ IndexDir = ${indexDir} }
        $dirA = Join-Path $TestDrive "textplan_a"
        $dirB = Join-Path $TestDrive "textplan_b"
        foreach ($dir in @($dirA, $dirB)) {
            [System.IO.Directory]::CreateDirectory($dir) | Out-Null
            [System.IO.File]::WriteAllText((Join-Path $dir "a.log"), "ログ")
            [System.IO.File]::WriteAllText((Join-Path $dir "b.txt"), "メモ")
        }
        function newTextPlanner($kindsArg, $previousRows = @()) {
            $planner = [IngestPlanner]::new($kindsArg)
            $status = @{ Folders = @(); Rows = (newPrevious $previousRows) }
            $planner.AssignNames(@(
                [pscustomobject]@{ Name = "営業"; Path = $dirA; Enabled = $true; Subfolders = $true; IngestedSubfolders = $true },
                [pscustomobject]@{ Name = "経理"; Path = $dirB; Enabled = $true; Subfolders = $true; IngestedSubfolders = $true }), $status)
            $planner.TextExcludes = @{ "営業" = @(".log") }
            return $planner
        }
    }

    It "営業だけ .log を取り込まないにすると、営業の行に .log の件数と、前に取り込んだ .log の外す件数が入り、経理には残る（元のファイルが無くなったものは外す件数に入らない）" {
        $planner = newTextPlanner $kinds @(
            (newStatusRow "営業\a.log" "2000/01/01 00:00:00" "1" ${stateDone}),
            (newStatusRow "営業\消えた.txt" "2000/01/01 00:00:00" "1" ${stateDone}),
            (newStatusRow "経理\a.log" "2000/01/01 00:00:00" "1" ${stateDone}))
        $planner.Crawl((newCounts @{ "営業\a.log" = 1; "営業\消えた.txt" = 1; "経理\a.log" = 1 }), $null)
        $rows = @($planner.Plan)
        $rows[0].インデックス名 | Should -Be "営業"
        $rows[0].TextExcludeRemovedCount | Should -Be 1
        $rows[0].TextExtensionCounts[".log"] | Should -Be 1
        $rows[0].TextExtensionCounts[".txt"] | Should -Be 1
        $rows[1].TextExcludeRemovedCount | Should -Be 0
        @($planner.Targets | ForEach-Object { $_.相対パス } | Sort-Object) | Should -Be @("営業\b.txt", "経理\a.log", "経理\b.txt")
    }

    It "Channel があれば、列挙したファイルの数を CountedFiles に書く（全フォルダの合計）。Channel が無くても動く" {
        $channel = [hashtable]::Synchronized(@{ CountedFiles = 0 })
        $planner = newTextPlanner @{ Unchecked = ${planKindUnchecked}; Missing = ${planKindMissing}; Dropped = ${planKindDropped}; StateFailed = ${stateFailed}; Channel = $channel }
        $planner.Crawl((newCounts), $null)
        $channel.CountedFiles | Should -Be 4
        $none = newTextPlanner $kinds
        { $none.Crawl((newCounts), $null) } | Should -Not -Throw
    }
}

Describe "Office・PDF を取り込まない拡張子にする（IngestPlanner の Crawl・findTargetFiles）" -Tag Io {
    BeforeAll {
        . "${scriptsDir}\shared\office\office_app.ps1"  # getApp（Office を起動する入口。Mock して呼ばれないことを確かめる）
        ${indexDir} = Join-Path $TestDrive "index_office_ext"
        $workspace = newTestWorkspace @{ IndexDir = ${indexDir} }
        $dirA = Join-Path $TestDrive "officeext_a"
        $dirB = Join-Path $TestDrive "officeext_b"
        foreach ($dir in @($dirA, $dirB)) {
            [System.IO.Directory]::CreateDirectory($dir) | Out-Null
            foreach ($name in @("a.xlsx", "b.pdf", "c.txt")) {
                [System.IO.File]::WriteAllText((Join-Path $dir $name), "x")
            }
        }
        function newOfficePlanner($excludes, $previousRows = @()) {
            $kinds = @{ Unchecked = ${planKindUnchecked}; Missing = ${planKindMissing}; Dropped = ${planKindDropped}; StateFailed = ${stateFailed} }
            $planner = [IngestPlanner]::new($kinds)
            $status = @{ Folders = @(); Rows = (newPrevious $previousRows) }
            $planner.AssignNames(@(
                [pscustomobject]@{ Name = "営業"; Path = $dirA; Enabled = $true; Subfolders = $true; IngestedSubfolders = $true },
                [pscustomobject]@{ Name = "経理"; Path = $dirB; Enabled = $true; Subfolders = $true; IngestedSubfolders = $true }), $status)
            $planner.TextExcludes = $excludes
            return $planner
        }
    }

    BeforeEach {
        Mock writeIndexerLog {}
    }

    It "A だけ .xlsx と .pdf を対象外にすると、A の findTargetFiles は返さず B は返す" {
        $a = findTargetFiles $dirA $null (getTargetExtensionList @(".xlsx", ".pdf"))
        $b = findTargetFiles $dirB $null (getTargetExtensionList @())
        (@($a.Files | ForEach-Object { $_.Name }) -join ",") | Should -Be "c.txt"
        (@($b.Files | ForEach-Object { $_.Name } | Sort-Object) -join ",") | Should -Be "a.xlsx,b.pdf,c.txt"
        # 対象外にしたものの件数も数える
        $a.TextCounts[".xlsx"] | Should -Be 1
        $a.TextCounts[".pdf"] | Should -Be 1
    }

    It "前に取り込んだ .xlsx・.pdf は Removed と TextExcludeRemovedCount に入り、他のインデックスは外れず、戻すと新規として数える" {
        $previousRows = @(
            (newStatusRow "営業\a.xlsx" "2000/01/01 00:00:00" "1" ${stateDone}),
            (newStatusRow "営業\b.pdf" "2000/01/01 00:00:00" "1" ${stateDone}),
            (newStatusRow "経理\a.xlsx" "2000/01/01 00:00:00" "1" ${stateDone}))
        $planner = newOfficePlanner @{ "営業" = @(".xlsx", ".pdf") } $previousRows
        $planner.Crawl((newCounts @{ "営業\a.xlsx" = 1; "営業\b.pdf" = 1; "経理\a.xlsx" = 1 }), $null)
        $rows = @($planner.Plan)
        $rows[0].インデックス名 | Should -Be "営業"
        $rows[0].TextExcludeRemovedCount | Should -Be 2
        $rows[0].TextExtensionCounts[".xlsx"] | Should -Be 1
        $rows[0].TextExtensionCounts[".pdf"] | Should -Be 1
        $rows[1].TextExcludeRemovedCount | Should -Be 0
        @($planner.RemovedPaths | Sort-Object) | Should -Be @("営業\a.xlsx", "営業\b.pdf")

        # 戻す（対象外なし）と、前に取り込んだ行が無ければ新規として対象に入り、外す件数は 0
        $back = newOfficePlanner @{} @()
        $back.Crawl((newCounts), $null)
        @($back.Targets | Where-Object { $_.相対パス -eq "営業\a.xlsx" }).Count | Should -Be 1
        @($back.Targets | Where-Object { $_.相対パス -eq "営業\b.pdf" }).Count | Should -Be 1
        @($back.Plan)[0].TextExcludeRemovedCount | Should -Be 0
    }

    It "Office 10 個をすべて対象外にしたインデックスだけの回では、Office の要るファイルが取り込みの対象に入らず、Office を起動しない（.xlsx があっても）" {
        Mock getApp {}
        $planner = newOfficePlanner @{ "営業" = @(${officeExtensions}); "経理" = @(${officeExtensions}) } @()
        $planner.Crawl((newCounts), $null)
        @($planner.Targets | Where-Object { $_.相対パス -match '\.xlsx$' }).Count | Should -Be 0
        @($planner.Targets | Where-Object { (getIngestApp $_.相対パス) -ne "" }).Count | Should -Be 0
        (@($planner.Targets | ForEach-Object { $_.相対パス } | Sort-Object) -join ",") | Should -Be "営業\b.pdf,営業\c.txt,経理\b.pdf,経理\c.txt"
        Should -Invoke getApp -Times 0 -Exactly
    }
}
