# 制限モードの読み込み口（tebunko_grep\restricted\lib_restricted.ps1）と、制限言語モードで決まるパスのテスト。
. "$PSScriptRoot\..\..\helpers\load.ps1"

Describe "lib_restricted.ps1" -Tag Unit {
    It "制限モードが使う関数と、パスの定義がそろう" {
        . "${scriptsDir}\tebunko_grep\restricted\lib_restricted.ps1"
        foreach ($name in @("readSettings", "readSearchOption", "newSearchRegex", "newFileFilter", "splitIndexTsvPath", "describePlace", "resolveSourcePath", "splitTsvCells", "normalizeFolderPath")) {
            (Get-Command $name -CommandType Function -ErrorAction SilentlyContinue) | Should Not BeNullOrEmpty
        }
        $indexDir | Should Be "$rootDir\work\index"
        $settingsFile | Should Be "$rootDir\setting.config"
    }
}

Describe "paths_grep.ps1（制限言語モードの TEMP）" -Tag Unit {
    It "GetTempPath の代わりに、環境変数 TMP・TEMP・USERPROFILE の順で TEMP を決める" {
        $saved = $env:TMP
        try {
            $env:TMP = "C:\tmp_for_test"
            $fullLanguage = $false
            . "${scriptsDir}\tebunko_grep\core\paths_grep.ps1"
            $tempRoot | Should Be "C:\tmp_for_test"
            $tmpDir | Should Be "C:\tmp_for_test\tebunko_grep\$PID"
            $env:TMP = ""
            . "${scriptsDir}\tebunko_grep\core\paths_grep.ps1"
            $tempRoot | Should Be $env:TEMP
        } finally {
            $env:TMP = $saved
        }
    }
}
