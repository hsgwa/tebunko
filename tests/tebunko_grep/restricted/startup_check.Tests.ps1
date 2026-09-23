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

Describe "testExcelInstalled" -Tag Io {
    # Excel の入っている PC でも入っていない PC でも同じ道筋を通るよう、探す場所を差し替えて確かめる
    It "Office の置き場所に EXCEL.EXE があれば true（クイック実行版・MSI 版のどちらも）" {
        foreach ($rel in @("Microsoft Office\root\Office16", "Microsoft Office\Office16")) {
            $root = Join-Path $TestDrive ($rel -replace '[\\ ]', "_")
            $dir = Join-Path $root $rel
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            Set-Content -LiteralPath (Join-Path $dir "EXCEL.EXE") -Value ""
            testExcelInstalled @($root) | Should Be $true
            # 2 つ目の場所にあっても見つける
            testExcelInstalled @((Join-Path $TestDrive "無い"), $root) | Should Be $true
        }
    }

    It "どこにも無ければ false（空の場所は飛ばす）" {
        testExcelInstalled @((Join-Path $TestDrive "無い"), "", $null) | Should Be $false
        testExcelInstalled @() | Should Be $false
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

Describe "readStartupModeSetting" -Tag Io {
    It "setting.config の startupMode を読む" {
        $path = Join-Path $TestDrive "setting.config"
        Set-Content -LiteralPath $path -Value '{ "startupMode": "clm", "useRegex": true }' -Encoding UTF8
        readStartupModeSetting $path | Should Be "clm"
    }

    It "ファイルが無い・空・壊れている・書かれていないときは空文字列（起動は続ける）" {
        readStartupModeSetting (Join-Path $TestDrive "無い.config") | Should Be ""
        $empty = Join-Path $TestDrive "空.config"
        Set-Content -LiteralPath $empty -Value "" -Encoding UTF8
        readStartupModeSetting $empty | Should Be ""
        $broken = Join-Path $TestDrive "壊れた.config"
        Set-Content -LiteralPath $broken -Value "{ これは JSON ではない" -Encoding UTF8
        readStartupModeSetting $broken | Should Be ""
        $other = Join-Path $TestDrive "ほか.config"
        Set-Content -LiteralPath $other -Value '{ "useRegex": true }' -Encoding UTF8
        readStartupModeSetting $other | Should Be ""
        $nulls = Join-Path $TestDrive "null.config"
        Set-Content -LiteralPath $nulls -Value "null" -Encoding UTF8
        readStartupModeSetting $nulls | Should Be ""
    }

    It "getStartupFacts は、設定ファイルを渡したときだけ読む" {
        $dir = Join-Path $TestDrive "work2"
        $path = Join-Path $TestDrive "facts.config"
        Set-Content -LiteralPath $path -Value '{ "startupMode": "restricted" }' -Encoding UTF8
        (getStartupFacts $dir $path).Setting | Should Be "restricted"
        (getStartupFacts $dir).Setting | Should Be ""
    }
}
