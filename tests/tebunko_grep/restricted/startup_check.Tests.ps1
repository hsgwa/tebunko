# 起動時の確認の事実集め（tebunko_grep\restricted\startup_check.ps1）のテスト。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\tebunko_grep\restricted\startup_check.ps1"

Describe "testFolderWritable" -Tag Io {
    It "書き込めるフォルダなら true で、確認用のファイルを残さない" {
        $dir = Join-Path $TestDrive "work"
        New-Item -ItemType Directory -Path $dir | Out-Null
        testFolderWritable $dir | Should Be $true
        @(Get-ChildItem -LiteralPath $dir).Count | Should Be 0
    }

    It "フォルダが無ければ作る" {
        $dir = Join-Path $TestDrive "新しい work"
        testFolderWritable $dir | Should Be $true
        Test-Path -LiteralPath $dir | Should Be $true
    }

    It "作れない場所なら false" {
        # 同じ名前のファイルがあると、フォルダを作れない
        $file = Join-Path $TestDrive "ファイル"
        Set-Content -LiteralPath $file -Value ""
        testFolderWritable (Join-Path $file "work") | Should Be $false
    }
}

Describe "getStartupFacts" -Tag Io {
    It "テストを動かしている PowerShell（FullLanguage）の事実を集める" {
        $dir = Join-Path $TestDrive "work"
        $facts = getStartupFacts $dir
        $facts.LanguageMode | Should Be "FullLanguage"
        $facts.WpfError | Should Be ""
        $facts.WorkDir | Should Be $dir
        $facts.CanWriteWork | Should Be $true
        $facts.HasExcel | Should Be (testExcelInstalled)
    }
}
