# ファイルの文字コードと改行のテスト。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
}

Describe "文字コードと改行" -Tag Meta {
    BeforeAll {
        # PowerShell 5.1 は BOM なしのファイルをシステム既定のコードページ（CP932）で読むため、
        # BOM を外すと日本語リテラルが化ける。改行は CRLF に揃える
        $files = @(Get-ChildItem "$here\..\scripts", "$here" -Recurse -Include *.ps1, *.xaml |
            Where-Object { $_.FullName -notlike "*\testdata\*" })
    }

    It "調べる対象のファイルがある" {
        $files.Count -gt 20 | Should -Be $true
    }

    It "すべて BOM 付き UTF-8" {
        $bad = @($files | Where-Object {
            $bytes = [System.IO.File]::ReadAllBytes($_.FullName)
            !($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
        } | ForEach-Object { $_.Name })
        ($bad -join ", ") | Should -Be ""
    }

    It "すべて改行が CRLF" {
        $bad = @($files | Where-Object {
            $text = [System.IO.File]::ReadAllText($_.FullName)
            [regex]::Matches($text, "(?<!`r)`n").Count -gt 0
        } | ForEach-Object { $_.Name })
        ($bad -join ", ") | Should -Be ""
    }

    # Windows で動く Codecov の CLI は codecov.yml を cp1252 として読み、日本語があると止まる（UnicodeDecodeError）。
    # そのため codecov.yml には ASCII の文字だけを書き、説明は .github\workflows\test.yml と docs\design\testing\ci.md に置く
    It "codecov.yml は ASCII の文字だけ" {
        $bytes = [System.IO.File]::ReadAllBytes("$here\..\.github\codecov.yml")
        @($bytes | Where-Object { $_ -gt 0x7F }).Count | Should -Be 0
    }

    # Actions は run の内容を BOM の無い UTF-8 の一時スクリプトにして渡す。shell: powershell（Windows PowerShell 5.1）
    # はこれを ANSI として読むため、run に ASCII 以外の文字（日本語のメッセージなど）があると文字化けして構文エラーになる
    # （v0.3.0 のタグの release で実際に起きた）。日本語のメッセージなどが要るときは tools\ の BOM 付き UTF-8 のスクリプトに移す。
    # shell: pwsh（PowerShell 7）はこの制限を受けないため調べない
    It "shell: powershell の run は ASCII だけ" {
        $workflows = Get-ChildItem "$here\..\.github\workflows" -Filter *.yml
        $workflows.Count -gt 5 | Should -Be $true

        $bad = @()
        foreach ($workflow in $workflows) {
            $text = [System.IO.File]::ReadAllText($workflow.FullName)
            # ワークフローのステップ（"- name:" か "- uses:" で始まる）ごとに切り分ける
            $steps = [regex]::Split($text, '(?m)^(?=\s*- (?:name|uses):)')
            foreach ($step in $steps) {
                if ($step -notmatch '(?m)^\s*shell:\s*powershell\s*$') {
                    continue
                }
                $runMatch = [regex]::Match($step, '(?m)^\s*run:\s*([\s\S]*)\z')
                if (!$runMatch.Success) {
                    continue
                }
                # 次のステップの手前にあるコメント行（このステップの run とは関係が無い）を除いてから調べる
                $codeLines = @($runMatch.Groups[1].Value -split "`n" | Where-Object { $_.Trim() -notmatch '^#' })
                if (($codeLines -join "`n") -match '[^\x00-\x7F]') {
                    $bad += $workflow.Name
                }
            }
        }
        ($bad -join ", ") | Should -Be ""
    }
}
