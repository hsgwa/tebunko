# 比較の作業フォルダ（tebunko_diff\job\diff_job.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\tebunko_diff\lib.ps1"

Describe "作業フォルダ" -Tag Io {
    It "画面の PID の下に、比較ごとのフォルダを作る" {
        $dir = newDiffJobDir 12345 "$TestDrive\root"
        Test-Path -LiteralPath $dir | Should Be $true
        (Split-Path (Split-Path $dir) -Leaf) | Should Be "12345"
        $second = newDiffJobDir 12345 "$TestDrive\root"
        $second | Should Not Be $dir
    }

    It "画面が強制終了されて残ったフォルダ（プロセスの無い PID）を消し、動いている画面のものは残す" {
        $root = "$TestDrive\stale"
        $dead = newDiffJobDir 999999 $root
        $alive = newDiffJobDir $PID $root
        removeStaleDiffDirs $root | Should Be 1
        Test-Path -LiteralPath $dead | Should Be $false
        Test-Path -LiteralPath $alive | Should Be $true
    }

    It "抽出要求・抽出結果・進み具合・優先・中止を読み書きする" {
        $dir = newDiffJobDir $PID "$TestDrive\io"
        writeExtractRequest $dir @(@{ Id = 1; Side = "left"; Path = "C:\共有\a.xlsx" }, @{ Id = 1; Side = "right"; Path = "C:\共有\b.xlsx" })
        $request = readExtractRequest $dir
        $request.Count | Should Be 2
        $request[1].Path | Should Be "C:\共有\b.xlsx"

        addExtractResult $dir 1 "left" ${diffStateDone} 3
        addExtractResult $dir 1 "right" ${diffStateFailed} 0 "パスワード`r`nが設定されています"
        $results = readExtractResults $dir
        $results["1|left"].Count | Should Be 3
        $results["1|right"].State | Should Be ${diffStateFailed}
        $results["1|right"].Message | Should Be "パスワード が設定されています"

        writeDiffProgress $dir 1 2 "C:\共有\a.xlsx"
        (readDiffProgress $dir).Done | Should Be 1

        writeDiffPriority $dir @(@{ Id = 4; Side = "right" })
        writeDiffPriority $dir @(@{ Id = 2; Side = "left" }, @{ Id = 2; Side = "right" })
        (readDiffPriority $dir) -join "," | Should Be "2|left,2|right"

        testDiffStopRequested $dir | Should Be $false
        requestDiffStop $dir
        testDiffStopRequested $dir | Should Be $true
    }

    It "書き込み途中の結果の行（列が足りない）は読まない" {
        $dir = newDiffJobDir $PID "$TestDrive\partial"
        [System.IO.File]::WriteAllText((Join-Path $dir ${diffResultFileName}), "1`tleft`t済`t2`t`r`n2`tri")
        $results = readExtractResults $dir
        @($results.Keys).Count | Should Be 1
    }
}

Describe "抽出した TSV" -Tag Io {
    It "作業フォルダの TSV を移し、抽出した順に読み戻す" {
        $tmp = "$TestDrive\work"
        [System.IO.Directory]::CreateDirectory($tmp) | Out-Null
        foreach ($name in @("明細", "表紙", "明細[図形]")) {
            $path = Join-Path $tmp (toIndexFileName $name)
            [System.IO.File]::WriteAllLines($path, [string[]]@("$name の 1 行目"), ${utf8Bom})
            Start-Sleep -Milliseconds 20
        }
        $dest = "$TestDrive\job\left\1"
        saveExtractedUnits $tmp $dest | Should Be 3
        @(Get-ChildItem -LiteralPath $tmp -Filter "*.tsv").Count | Should Be 0
        $units = readExtractedUnits $dest
        (@($units.Keys) -join ",") | Should Be "明細,表紙,明細[図形]"
        $units["明細[図形]"][0] | Should Be "明細[図形] の 1 行目"
    }

    It "順番.txt が無ければ名前の順に読む。フォルダが無ければ空" {
        $dir = "$TestDrive\noorder"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        [System.IO.File]::WriteAllLines((Join-Path $dir "b.tsv"), [string[]]@("2"))
        [System.IO.File]::WriteAllLines((Join-Path $dir "a.tsv"), [string[]]@("1"))
        (@((readExtractedUnits $dir).Keys) -join ",") | Should Be "a,b"
        @((readExtractedUnits "$TestDrive\無い").Keys).Count | Should Be 0
    }
}

Describe "フォルダのファイル" -Tag Io {
    $root = "$TestDrive\folder"
    foreach ($rel in @("a.xlsx", "sub\b.docx", "sub\~`$b.docx", "c.txt")) {
        $path = Join-Path $root $rel
        [System.IO.Directory]::CreateDirectory((Split-Path $path)) | Out-Null
        [System.IO.File]::WriteAllText($path, $rel)
    }

    It "Office ファイルだけを、相対パス・サイズ付きで返す（Office の一時ファイルは除く）" {
        $scan = getFolderFiles $root $true
        (@($scan.Files | ForEach-Object { $_.RelPath } | Sort-Object) -join ",") | Should Be "a.xlsx,sub\b.docx"
        ($scan.Files | Where-Object { $_.RelPath -eq "a.xlsx" }).Size | Should Be 6
    }

    It "サブフォルダを比べないときは直下だけ" {
        $scan = getFolderFiles $root $false
        (@($scan.Files | ForEach-Object { $_.RelPath }) -join ",") | Should Be "a.xlsx"
    }

    It "ハッシュは中身が同じなら同じ" {
        $x = Join-Path $root "x.bin"
        $y = Join-Path $root "y.bin"
        [System.IO.File]::WriteAllText($x, "同じ")
        [System.IO.File]::WriteAllText($y, "同じ")
        getFileHash $x | Should Be (getFileHash $y)
    }

    It "比較結果ファイルを書く" {
        $path = writeDiffReport @("1 行目", "2 行目") "$TestDrive\out\比較結果.txt"
        [System.IO.File]::ReadAllLines($path)[1] | Should Be "2 行目"
    }
}
