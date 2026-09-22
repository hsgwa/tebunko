# 比較の抽出プロセス（tebunko_diff\differ.ps1）のテスト。本物の Excel・Word・PowerPoint を使う（タグ Office）。
# 抽出プロセスは画面と同じく、別の powershell.exe で起動する
. "$PSScriptRoot\..\helpers\load.ps1"
. "${scriptsDir}\tebunko_diff\lib.ps1"

$differ = "${scriptsDir}\tebunko_diff\differ.ps1"
$office = "${testDataDir}\office"

function runDiffer {
    # 抽出プロセスを起動し、終わるまで待って終了コードを返す
    param ([string]$jobDir)

    $process = Start-Process -FilePath "powershell.exe" -WindowStyle Hidden -PassThru `
        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$differ`" -JobDir `"$jobDir`""
    $null = $process.Handle
    if (!$process.WaitForExit(600000)) {
        $process.Kill()
        throw "抽出プロセスが終わりませんでした"
    }
    return $process.ExitCode
}

Describe "differ.ps1" -Tag Office {
    It "抽出して比べられる（Excel・Word・PowerPoint）。見つからないファイルは失敗にする" {
        $job = newDiffJobDir $PID "$TestDrive\diff"
        writeExtractRequest $job @(
            @{ Id = 1; Side = "left"; Path = "$office\Excel\2024\見積\A社.xlsx" }
            @{ Id = 1; Side = "right"; Path = "$office\Excel\2025\見積\A社.xlsx" }
            @{ Id = 2; Side = "left"; Path = "$office\Word\基本.docx" }
            @{ Id = 2; Side = "right"; Path = "$office\Word\基本.docx" }
            @{ Id = 3; Side = "left"; Path = "$office\PowerPoint\基本.pptx" }
            @{ Id = 3; Side = "right"; Path = "$office\PowerPoint\非表示スライド.pptx" }
            @{ Id = 4; Side = "left"; Path = "$office\無い.xlsx" }
        )
        runDiffer $job | Should Be 0

        $results = readExtractResults $job
        @($results.Keys).Count | Should Be 7
        $results["4|left"].State | Should Be ${diffStateFailed}
        $results["4|left"].Message | Should Match "ファイルが見つかりません"

        # Excel: 2024 年度と 2025 年度の見積の違いが出る
        $excel = compareOfficeUnits "Excel" (readExtractedUnits (getExtractDir $job "left" 1)) (readExtractedUnits (getExtractDir $job "right" 1))
        @($excel.Places | Where-Object { $_.Status -ne "same" }).Count -gt 0 | Should Be $true
        ((@($excel.Places | ForEach-Object { $_.Rows } | ForEach-Object { $_.RightText }) -join "`n") -match "2025") | Should Be $true

        # Word: 同じファイルは違いが無い
        $word = compareOfficeUnits "Word" (readExtractedUnits (getExtractDir $job "left" 2)) (readExtractedUnits (getExtractDir $job "right" 2))
        ($word.Inserts + $word.Deletes + $word.Changes) | Should Be 0

        # PowerPoint: スライドの見出しの行ができる
        $ppt = compareOfficeUnits "PowerPoint" (readExtractedUnits (getExtractDir $job "left" 3)) (readExtractedUnits (getExtractDir $job "right" 3))
        @($ppt.Places[0].Rows | Where-Object { $_.Type -eq "Header" }).Count -gt 0 | Should Be $true

        # 作業フォルダ（work）は消えている
        Test-Path -LiteralPath (Join-Path $job "work") | Should Be $false
        (readDiffProgress $job).Done | Should Be 7
    }

    It "優先されたファイルを先に抽出する" {
        $job = newDiffJobDir $PID "$TestDrive\priority"
        writeExtractRequest $job @(
            @{ Id = 1; Side = "left"; Path = "$office\Word\基本.docx" }
            @{ Id = 2; Side = "left"; Path = "$office\Word\表.docx" }
        )
        writeDiffPriority $job @(@{ Id = 2; Side = "left" })
        runDiffer $job | Should Be 0
        $first = @([System.IO.File]::ReadAllLines((Join-Path $job ${diffResultFileName})))[0]
        $first | Should Match "^2`tleft`t"
    }

    It "中止を頼まれていれば、抽出せずに終了コード 2 で終わる" {
        $job = newDiffJobDir $PID "$TestDrive\stop"
        writeExtractRequest $job @(@{ Id = 1; Side = "left"; Path = "$office\Word\基本.docx" })
        requestDiffStop $job
        runDiffer $job | Should Be 2
        Test-Path -LiteralPath (Join-Path $job ${diffResultFileName}) | Should Be $false
    }

    It "作業フォルダが無ければ、続けられないエラー（終了コード 1）" {
        runDiffer "$TestDrive\無いフォルダ" | Should Be 1
    }
}
