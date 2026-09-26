# クロール（tebunko\indexer\indexer_plan.ps1 の createTargetList）のテスト。
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

Describe "createTargetList" -Tag Io {
    $source = Join-Path $TestDrive "src"
    [System.IO.Directory]::CreateDirectory($source) | Out-Null
    $file = Join-Path $source "a.xlsx"
    [System.IO.File]::WriteAllText($file, "dummy")
    $info = Get-Item -LiteralPath $file
    $updated = formatFileTime $info.LastWriteTime
    $size = [string]$info.Length
    $folder = @{ Path = $source; Name = "売上" }
    # getIndexFiles / removeBookDir が実際のインデックスを見ないよう、テスト用のフォルダに向ける
    ${indexDir} = Join-Path $TestDrive "index"

    It "一覧に無いファイルは取り込み対象になる（新規）" {
        $result = createTargetList $folder (newPrevious) $null
        $result.Targets.Count | Should Be 1
        $result.Targets[0].相対パス | Should Be "売上\a.xlsx"
        $result.Plan.新規 | Should Be 1
        $result.Plan.ファイル数 | Should Be 1
        $result.Plan.取り込み対象 | Should Be 1
    }

    It "取り込み済みで更新が無ければ取り込まない" {
        $previous = newPrevious @((newStatusRow "売上\a.xlsx" $updated $size ${stateDone} 1 $updated "" "2"))
        $result = createTargetList $folder $previous (newCounts @{ "売上\a.xlsx" = 1 })
        $result.Targets.Count | Should Be 0
        $result.Plan.取り込み対象 | Should Be 0
        $result.Rows.Count | Should Be 1
    }

    It "前の抽出版で取り込んだファイルは、更新が無くても取り込み直す（更新ありに数える）" {
        $previous = newPrevious @((newStatusRow "売上\a.xlsx" $updated $size ${stateDone} 1 $updated))
        $result = createTargetList $folder $previous (newCounts @{ "売上\a.xlsx" = 1 })
        $result.Targets.Count | Should Be 1
        $result.Plan.更新あり | Should Be 1
    }

    It "取り込み済みでも TSV が無ければ取り込み直す（インデックスなし）" {
        $previous = newPrevious @((newStatusRow "売上\a.xlsx" $updated $size ${stateDone} 1 $updated))
        $result = createTargetList $folder $previous (newCounts)
        $result.Targets.Count | Should Be 1
        $result.Plan.インデックスなし | Should Be 1
    }

    It "更新されていれば取り込み対象になる（更新あり）" {
        $previous = newPrevious @((newStatusRow "売上\a.xlsx" "2000/01/01 00:00:00" $size ${stateDone} 1 $updated))
        $result = createTargetList $folder $previous (newCounts @{ "売上\a.xlsx" = 1 })
        $result.Targets.Count | Should Be 1
        $result.Plan.更新あり | Should Be 1
    }

    It "前回失敗して更新が無ければ、取り込み対象ではなく失敗として返す" {
        $previous = newPrevious @((newStatusRow "売上\a.xlsx" $updated $size ${stateFailed} 0 $updated "開けませんでした"))
        $result = createTargetList $folder $previous (newCounts @{ "売上\a.xlsx" = 1 })
        $result.Targets.Count | Should Be 0
        $result.Failed.Count | Should Be 1
        $result.Plan.前回失敗 | Should Be 1
    }

    It "前回「未取り込み」で終わっていれば取り込み対象になる（前回未完了）" {
        $previous = newPrevious @((newStatusRow "売上\a.xlsx" $updated $size ${stateNew}))
        $result = createTargetList $folder $previous (newCounts @{ "売上\a.xlsx" = 1 })
        $result.Targets.Count | Should Be 1
        $result.Plan.前回未完了 | Should Be 1
    }
}

Describe "createTargetList（インデックスが先にあるファイル・無くなったファイル）" -Tag Io {
    $source = Join-Path $TestDrive "src2"
    [System.IO.Directory]::CreateDirectory("$source\2024") | Out-Null
    $file = Join-Path $source "2024\b.docx"
    [System.IO.File]::WriteAllText($file, "dummy")
    (Get-Item -LiteralPath $file).LastWriteTime = [datetime]"2024/04/01 09:00:00"
    $folder = @{ Path = $source; Name = "経理" }
    ${indexDir} = Join-Path $TestDrive "index2"
    $bookDir = Join-Path ${indexDir} "経理\2024\b.docx"

    function newIndexTsv([string]$name, [datetime]$time) {
        [System.IO.Directory]::CreateDirectory($bookDir) | Out-Null
        $path = Join-Path $bookDir $name
        [System.IO.File]::WriteAllText($path, "1`t見積")
        (Get-Item -LiteralPath $path).LastWriteTime = $time
    }

    AfterEach {
        if (Test-Path -LiteralPath ${indexDir}) { Remove-Item -LiteralPath ${indexDir} -Recurse -Force }
    }

    It "サブフォルダのファイルは、インデックス名とフォルダからの相対パスでつなぐ" {
        $result = createTargetList $folder (newPrevious) $null
        $result.Rows[0].相対パス | Should Be "経理\2024\b.docx"
    }

    It "一覧に無くても、元のファイルより新しいインデックスがあれば取り込み済みとする" {
        newIndexTsv "1ページ.tsv" ([datetime]"2024/04/02 09:00:00")
        newIndexTsv "2ページ.tsv" ([datetime]"2024/04/03 09:00:00")
        $result = createTargetList $folder (newPrevious) $null
        $result.Targets.Count | Should Be 0
        $result.Rows[0].状態 | Should Be ${stateDone}
        $result.Rows[0].TSV数 | Should Be "2"
        # 取り込み日時は、いちばん新しいTSVの更新日時にする
        $result.Rows[0].取り込み日時 | Should Be (formatFileTime ([datetime]"2024/04/03 09:00:00"))
    }

    It "インデックスが元のファイルより古ければ取り込み対象にする" {
        newIndexTsv "1ページ.tsv" ([datetime]"2024/03/01 09:00:00")
        $result = createTargetList $folder (newPrevious) $null
        $result.Targets.Count | Should Be 1
        $result.Plan.新規 | Should Be 1
    }

    It "インデックスの数え上げ（counts）に無いファイルは、ディスクを調べずに取り込み対象にする" {
        newIndexTsv "1ページ.tsv" ([datetime]"2024/04/02 09:00:00")
        $result = createTargetList $folder (newPrevious) (newCounts)
        $result.Targets.Count | Should Be 1
    }

    It "一覧にあって元のファイルが無くなったものは、インデックスを消して一覧から除く" {
        $gone = Join-Path ${indexDir} "経理\消えた.xlsx"
        [System.IO.Directory]::CreateDirectory($gone) | Out-Null
        [System.IO.File]::WriteAllText("$gone\Sheet1.tsv", "1`t見積")
        $previous = newPrevious @((newStatusRow "経理\消えた.xlsx" "2024/01/01 00:00:00" "10" ${stateDone} 1 "2024/01/01 00:00:00"))
        $result = createTargetList $folder $previous $null
        @($result.Rows | Where-Object { $_.相対パス -eq "経理\消えた.xlsx" }).Count | Should Be 0
        Test-Path -LiteralPath $gone | Should Be $false
    }

    It "ほかのインデックスの行には触らない（名前が前方一致するインデックスも別のものとして扱う）" {
        $other = Join-Path ${indexDir} "経理2\c.xlsx"
        [System.IO.Directory]::CreateDirectory($other) | Out-Null
        $previous = newPrevious @((newStatusRow "経理2\c.xlsx" "2024/01/01 00:00:00" "10" ${stateDone} 1 "2024/01/01 00:00:00"))
        $result = createTargetList $folder $previous $null
        @($result.Rows | Where-Object { $_.相対パス -eq "経理2\c.xlsx" }).Count | Should Be 0
        Test-Path -LiteralPath $other | Should Be $true
    }

    It "アクセスできないフォルダがあったときは、見つからなかったファイルの行とインデックスを残す" {
        Mock findOfficeFiles { @{ Root = $source; Files = @(Get-Item -LiteralPath $file); HasError = $true } }
        $gone = Join-Path ${indexDir} "経理\読めない\d.xlsx"
        [System.IO.Directory]::CreateDirectory($gone) | Out-Null
        $previous = newPrevious @((newStatusRow "経理\読めない\d.xlsx" "2024/01/01 00:00:00" "10" ${stateDone} 1 "2024/01/01 00:00:00"))
        $result = createTargetList $folder $previous $null
        @($result.Rows | Where-Object { $_.相対パス -eq "経理\読めない\d.xlsx" }).Count | Should Be 1
        Test-Path -LiteralPath $gone | Should Be $true
        # \\?\ の付かないパスで列挙されたファイルも、フォルダからの相対パスにする
        $result.Rows[0].相対パス | Should Be "経理\2024\b.docx"
    }
}

Describe "findOfficeFiles" -Tag Io {
    $source = Join-Path $TestDrive "scan"
    [System.IO.Directory]::CreateDirectory("$source\下\さらに下") | Out-Null
    foreach ($name in @("a.xlsx", "下\b.DOCX", "下\さらに下\c.pptm", "d.txt", "e.pdf", ('~$' + "a.xlsx"), "f.xls", "g.ppt")) {
        [System.IO.File]::WriteAllText((Join-Path $source $name), "dummy")
    }

    It "サブフォルダも含めて Office の拡張子のファイルだけを返す（大文字の拡張子も含め、~$ で始まるロックファイルは除く）" {
        $scan = findOfficeFiles $source
        @($scan.Files | ForEach-Object { $_.Name } | Sort-Object) -join "," | Should Be "a.xlsx,b.DOCX,c.pptm,f.xls,g.ppt"
        $scan.HasError | Should Be $false
    }

    It "末尾に \ を付けたフォルダでも同じフォルダを列挙する" {
        $scan = findOfficeFiles "$source\"
        $scan.Root.TrimEnd("\") | Should Be $source
        $scan.Files.Count | Should Be 5
    }

    It "Office のファイルが無いフォルダは 0 件" {
        $empty = Join-Path $TestDrive "空"
        [System.IO.Directory]::CreateDirectory($empty) | Out-Null
        $scan = findOfficeFiles $empty
        @($scan.Files).Count | Should Be 0
        $scan.HasError | Should Be $false
    }
}

Describe "getIndexFiles" -Tag Io {
    It "フォルダの中の TSV だけを返す" {
        $dir = Join-Path $TestDrive "book.xlsx"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        [System.IO.File]::WriteAllText("$dir\Sheet1.tsv", "")
        [System.IO.File]::WriteAllText("$dir\メモ.txt", "")
        @(getIndexFiles $dir | ForEach-Object { $_.Name }) -join "," | Should Be "Sheet1.tsv"
    }

    It "フォルダが無ければ空" {
        @(getIndexFiles (Join-Path $TestDrive "無い.xlsx")).Count | Should Be 0
    }
}

Describe "getBookDir" -Tag Unit {
    It "インデックスのフォルダに相対パスをつなぐ" {
        ${indexDir} = "C:\tool\work\index"
        getBookDir "営業\2024\A社.xlsx" | Should Be "C:\tool\work\index\営業\2024\A社.xlsx"
    }
}

Describe "waitForIndexingApproval" -Tag Io {
    # 画面とのやり取りのファイルをテスト用のフォルダに向ける（関数は呼び出し元の変数を見る）
    $stopRequestFile          = Join-Path $TestDrive "インデックス作成中止要求"
    $ingestPlanFile           = Join-Path $TestDrive "取り込み予定.tsv"
    $indexingProgressFile     = Join-Path $TestDrive "インデックス作成進捗.txt"
    $indexingStartRequestFile = Join-Path $TestDrive "インデックス作成開始要求"
    $approvalTimeoutMinutes   = 60
    $plan = @((newIngestPlanRow "営業" "C:\共有\営業部" ${planKindIngest} 3 2 1 1 0 0 1))

    AfterEach {
        foreach ($path in @($stopRequestFile, $ingestPlanFile, $indexingProgressFile, $indexingStartRequestFile)) {
            if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
        }
    }

    It "取り込み予定と確認待ちの進み具合を書いてから待ち、画面が開始を選んだら返事を返す" {
        Mock Start-Sleep {
            # 待っている間に書かれている内容を確かめてから、画面が開始を選ぶ
            $script:seenPlan = @(readIngestPlan $ingestPlanFile)
            $script:seenProgress = readIndexingProgress $indexingProgressFile
            writeIndexingStartRequest $false $indexingStartRequestFile
        }
        $answer = waitForIndexingApproval $plan 2 1
        $answer.RetryFailed | Should Be $false
        $script:seenPlan.Count | Should Be 1
        $script:seenPlan[0].インデックス名 | Should Be "営業"
        $script:seenProgress.Phase | Should Be ${indexingPhaseConfirm}
        $script:seenProgress.Remaining | Should Be 2
        $script:seenProgress.Failed | Should Be 1
        # 返事を読んだら、開始要求と取り込み予定は消す
        Test-Path -LiteralPath $indexingStartRequestFile | Should Be $false
        Test-Path -LiteralPath $ingestPlanFile | Should Be $false
    }

    It "画面が失敗分の再取り込みを選んだら RetryFailed を返す" {
        Mock Start-Sleep { writeIndexingStartRequest $true $indexingStartRequestFile }
        (waitForIndexingApproval $plan 2 1).RetryFailed | Should Be $true
    }

    It "中止要求があれば `$null を返し、中止要求と取り込み予定を消す" {
        Mock Start-Sleep { [System.IO.File]::WriteAllText($stopRequestFile, "") }
        waitForIndexingApproval $plan 2 1 | Should BeNullOrEmpty
        Test-Path -LiteralPath $stopRequestFile | Should Be $false
        Test-Path -LiteralPath $ingestPlanFile | Should Be $false
    }

    It "前回残った開始要求は使わない（待ち始める前に消す）" {
        writeIndexingStartRequest $true $indexingStartRequestFile
        Mock Start-Sleep { [System.IO.File]::WriteAllText($stopRequestFile, "") }
        waitForIndexingApproval $plan 2 1 | Should BeNullOrEmpty
    }

    It "制限時間を過ぎても返事が無ければ `$null を返す" {
        # 制限時間を負にして、待ち始めた時点で制限時刻を過ぎているようにする
        $approvalTimeoutMinutes = -1
        Mock Start-Sleep { }
        waitForIndexingApproval $plan 2 1 | Should BeNullOrEmpty
        Test-Path -LiteralPath $ingestPlanFile | Should Be $false
    }
}
