# ファイルの文字コードと改行のテスト。
. "$PSScriptRoot\..\helpers\load.ps1"

Describe "文字コードと改行" -Tag Meta {
    # PowerShell 5.1 は BOM なしのファイルをシステム既定のコードページ（CP932）で読むため、
    # BOM を外すと日本語リテラルが化ける。改行は CRLF に揃える
    $files = @(Get-ChildItem "$here\..\scripts", "$here" -Recurse -Include *.ps1, *.xaml |
        Where-Object { $_.FullName -notlike "*\testdata\*" })

    It "調べる対象のファイルがある" {
        $files.Count -gt 20 | Should Be $true
    }

    It "すべて BOM 付き UTF-8" {
        $bad = @($files | Where-Object {
            $bytes = [System.IO.File]::ReadAllBytes($_.FullName)
            !($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
        } | ForEach-Object { $_.Name })
        ($bad -join ", ") | Should Be ""
    }

    It "すべて改行が CRLF" {
        $bad = @($files | Where-Object {
            $text = [System.IO.File]::ReadAllText($_.FullName)
            [regex]::Matches($text, "(?<!`r)`n").Count -gt 0
        } | ForEach-Object { $_.Name })
        ($bad -join ", ") | Should Be ""
    }

    # Windows で動く Codecov の CLI は codecov.yml を cp1252 として読み、日本語があると止まる（UnicodeDecodeError）。
    # そのため codecov.yml には ASCII の文字だけを書き、説明は .github\workflows\test.yml と docs\00_共通_3_テスト.md に置く
    It "codecov.yml は ASCII の文字だけ" {
        $bytes = [System.IO.File]::ReadAllBytes("$here\..\.github\codecov.yml")
        @($bytes | Where-Object { $_ -gt 0x7F }).Count | Should Be 0
    }
}