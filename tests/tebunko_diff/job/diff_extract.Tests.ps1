# 比較の抽出の本体（tebunko_diff\job\diff_extract.ps1 の invokeDiffer）のテスト。
# Word・PowerPoint の新形式（.docx・.pptx）は Office を使わずに読むため、Office が無くても動く（Excel は Office タグの differ.Tests.ps1）
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\tebunko_diff\lib.ps1"
. "${scriptsDir}\shared\office\office_reader.ps1"
. "${scriptsDir}\shared\office\office_app.ps1"
. "${scriptsDir}\shared\office\office_extract.ps1"
. "${scriptsDir}\tebunko_diff\job\diff_extract.ps1"

$office = "${testDataDir}\office"

Describe "invokeDiffer" -Tag Io {
    It "抽出して、場所ごとの TSV と結果を書く。見つからないファイルは失敗にする" {
        $job = newDiffJobDir $PID "$TestDrive\run"
        writeExtractRequest $job @(
            @{ Id = 1; Side = "left"; Path = "$office\Word\基本.docx" }
            @{ Id = 1; Side = "right"; Path = "$office\PowerPoint\基本.pptx" }
            @{ Id = 2; Side = "left"; Path = "$office\無い.docx" }
        )
        invokeDiffer $job | Should Be 0

        $results = readExtractResults $job
        $results["1|left"].State | Should Be ${diffStateDone}
        $results["1|right"].State | Should Be ${diffStateDone}
        $results["2|left"].State | Should Be ${diffStateFailed}
        $results["2|left"].Message | Should Match "ファイルが見つかりません"
        @((readExtractedUnits (getExtractDir $job "left" 1)).Keys).Count -gt 0 | Should Be $true
        @((readExtractedUnits (getExtractDir $job "right" 1)).Keys) -contains "スライド001" | Should Be $true
        Test-Path -LiteralPath (Join-Path $job "work") | Should Be $false
        $progress = readDiffProgress $job
        $progress.Done | Should Be 3
        $progress.Current | Should Be ""
    }

    It "優先されたものを先に抽出し、結果のあるものは抽出し直さない" {
        $job = newDiffJobDir $PID "$TestDrive\priority"
        writeExtractRequest $job @(
            @{ Id = 1; Side = "left"; Path = "$office\Word\基本.docx" }
            @{ Id = 2; Side = "left"; Path = "$office\Word\表.docx" }
            @{ Id = 3; Side = "left"; Path = "$office\Word\空文書.docx" }
        )
        addExtractResult $job 3 "left" ${diffStateDone} 0
        writeDiffPriority $job @(@{ Id = 9; Side = "left" }, @{ Id = 2; Side = "left" })
        invokeDiffer $job | Should Be 0
        $lines = @([System.IO.File]::ReadAllLines((Join-Path $job ${diffResultFileName})))
        $lines.Count | Should Be 3
        $lines[1] | Should Match "^2`tleft`t"
        $lines[2] | Should Match "^1`tleft`t"
    }

    It "中止を頼まれていれば、抽出せずに 2 を返す" {
        $job = newDiffJobDir $PID "$TestDrive\stop"
        writeExtractRequest $job @(@{ Id = 1; Side = "left"; Path = "$office\Word\基本.docx" })
        requestDiffStop $job
        invokeDiffer $job | Should Be 2
        Test-Path -LiteralPath (Join-Path $job ${diffResultFileName}) | Should Be $false
    }

    It "作業フォルダが無ければ例外にする" {
        { invokeDiffer "$TestDrive\無いフォルダ" } | Should Throw "比較の作業フォルダが見つかりません"
    }
}
