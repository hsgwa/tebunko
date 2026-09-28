# ファイルの文字コードと改行のテスト。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"

    # run: の値を、YAML のインデントに基づいて取り出す（ブロックスカラー | 対応）。
    # 次のステップの手前にある YAML のコメント（run より浅いインデント）は含めない。
    # run の中の PowerShell のコメント（run より深いインデント）は含める
    function Get-PowerShellRunText {
        param (
            [string[]]$Lines,
            [int]$RunLineIndex
        )

        if ($Lines[$RunLineIndex] -notmatch '^(?<indent>[ \t]*)run:[ \t]*(?<rest>.*)$') {
            return $null
        }
        $indent = $Matches['indent'].Length
        $rest = $Matches['rest']

        if ($rest -notmatch '^[|>][+-]?[ \t]*$') {
            # run: <コマンド> の 1 行の形
            return $rest
        }

        $collected = @()
        for ($i = $RunLineIndex + 1; $i -lt $Lines.Count; $i++) {
            $line = $Lines[$i]
            if ($line.Trim() -eq "") {
                $collected += $line
                continue
            }
            $lineIndent = ($line -replace '^([ \t]*).*', '$1').Length
            if ($lineIndent -le $indent) {
                break
            }
            $collected += $line
        }
        return ($collected -join "`n")
    }

    # ワークフローの本文から、shell: powershell で run に ASCII 以外の文字があるステップの名を返す
    function Get-BadPowershellRunSteps {
        param (
            [string]$WorkflowText
        )

        $bad = @()
        # CRLF だと $ が \r の手前で止まり行末に掛からないため、先に \n に揃える
        $normalized = $WorkflowText -replace "`r`n", "`n"
        # ワークフローのステップ（"- name:" か "- uses:" で始まる）ごとに切り分ける
        $steps = [regex]::Split($normalized, '(?m)^(?=[ \t]*- (?:name|uses):)')
        foreach ($step in $steps) {
            if ($step -notmatch '(?m)^[ \t]*shell:[ \t]*powershell[ \t]*$') {
                continue
            }
            $lines = $step -split "`r?`n"
            $runIndex = -1
            for ($i = 0; $i -lt $lines.Count; $i++) {
                if ($lines[$i] -match '^[ \t]*run:') {
                    $runIndex = $i
                    break
                }
            }
            if ($runIndex -eq -1) {
                continue
            }
            $run = Get-PowerShellRunText -Lines $lines -RunLineIndex $runIndex
            if ($null -ne $run -and $run -match '[^\x00-\x7F]') {
                $bad += $lines[0].Trim()
            }
        }
        return $bad
    }
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
    # はこれを ANSI として読むため、run に ASCII 以外の文字（日本語のメッセージ・コメントなど）があると文字化けして
    # 構文エラーになる（v0.3.0 のタグの release で実際に起きた）。日本語が要るときは tools\ の BOM 付き UTF-8 の
    # スクリプトに移す。shell: pwsh（PowerShell 7）はこの制限を受けないため調べない
    It "shell: powershell の run は ASCII だけ" {
        $workflows = Get-ChildItem "$here\..\.github\workflows" -Filter *.yml
        $workflows.Count -gt 5 | Should -Be $true

        $bad = @()
        foreach ($workflow in $workflows) {
            $text = [System.IO.File]::ReadAllText($workflow.FullName)
            foreach ($step in (Get-BadPowershellRunSteps $text)) {
                $bad += "$($workflow.Name): $step"
            }
        }
        ($bad -join ", ") | Should -Be ""
    }

    # 上のテストが確かめる Get-BadPowershellRunSteps 自体の判定が正しいことを、YAML の断片で確かめる
    It "<name>" -TestCases @(
        @{
            name     = "run のメッセージの日本語 → 検出する"
            text     = @"
      - name: 例
        shell: powershell
        run: |
          if (`$true) { throw "日本語のメッセージ" }
"@
            expected = $true
        }
        @{
            name     = "run の中の PowerShell のコメントの日本語 → 検出する"
            text     = @"
      - name: 例
        shell: powershell
        run: |
          Write-Host "ok"
          # 日本語のコメント
"@
            expected = $true
        }
        @{
            name     = "次のステップの手前にある YAML のコメントの日本語 → 検出しない"
            text     = @"
      - name: 例
        shell: powershell
        run: |
          Write-Host "ok"

      # 日本語のコメント
      - name: 次
        run: OK
"@
            expected = $false
        }
        @{
            name     = "shell: pwsh の日本語 → 検出しない"
            text     = @"
      - name: 例
        shell: pwsh
        run: |
          throw "日本語のメッセージ"
"@
            expected = $false
        }
        @{
            name     = "run: <コマンド> の 1 行の形の日本語 → 検出する"
            text     = @"
      - name: 例
        shell: powershell
        run: Write-Host "日本語"
"@
            expected = $true
        }
    ) {
        param ($name, $text, $expected)
        (@(Get-BadPowershellRunSteps $text).Count -gt 0) | Should -Be $expected
    }
}
