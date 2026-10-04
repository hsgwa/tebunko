# 起動そのものに失敗したときの知らせ（tebunko\ui\startup_error.ps1）のテスト。
# gui.ps1 の読み込みの一番最初（app_host.ps1 より前）に使うため、ここだけを読み込んで確かめる
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "$scriptsDir\tebunko\ui\startup_error_view.ps1"
    . "$scriptsDir\tebunko\ui\startup_error.ps1"
}

Describe "getStartupErrorDetail" -Tag Unit {
    It "時刻・言語モード・メッセージ・位置を 1 件にまとめ、末尾に空行を入れる" {
        try { throw "壊れた" } catch { $err = $_ }
        $detail = getStartupErrorDetail $err "FullLanguage" "2026-01-02 03:04:05"
        $detail | Should -Match "^==== 2026-01-02 03:04:05 起動・実行中 ====\r\nLanguageMode: FullLanguage\r\n壊れた\r\n"
        $detail | Should -Match "\r\n\r\n$"
    }
}

Describe "getStartupErrorMessage" -Tag Unit {
    It "<Case>" -ForEach @(
        @{ Case = "記録できたら、そのファイル名だけを添える"; File = "C:\共有\営業部\startup_error.txt"; Expected = "予期しないエラーが発生しました。`n壊れた`n`n詳しい内容は startup_error.txt に残しています。" }
        @{ Case = "記録できなかったら、添えない"; File = $null; Expected = "予期しないエラーが発生しました。`n壊れた" }
    ) {
        getStartupErrorMessage "壊れた" $File | Should -Be $Expected
    }
}
Describe "getExistingRecordFile" -Tag Unit {
    It "実際に書けたファイルだけを返す" {
        $file = Join-Path $TestDrive "exists.txt"
        [System.IO.File]::WriteAllText($file, "x")
        getExistingRecordFile $file | Should -Be $file
    }

    It "<case> なら `$null を返す" -ForEach @(
        @{ Case = "ファイルが無い"; Path = { Join-Path $TestDrive "missing.txt" } }
        @{ Case = "パスが空"; Path = { "" } }
    ) {
        getExistingRecordFile (& $Path) | Should -Be $null
    }
}

Describe "getGuiErrorLogFile" -Tag Unit {
    AfterEach {
        Remove-Variable -Name workspace -Scope Script -ErrorAction SilentlyContinue
    }

    It "ワークスペースが決まる前は `$null を返す" {
        Remove-Variable -Name workspace -Scope Script -ErrorAction SilentlyContinue
        getGuiErrorLogFile | Should -Be $null
    }

    It "startGui のような関数の中でワークスペースを決めたあと、その関数を抜けた外側からも記録先を返す（外側の catch から writeErrorLog が使える）" {
        function startGuiLike {
            $script:workspace = [pscustomobject]@{ GuiErrorLogFile = "gui_error_log.txt" }
            throw "画面の組み立ての失敗"
        }
        try { startGuiLike } catch { }
        getGuiErrorLogFile | Should -Be "gui_error_log.txt"
    }
}

Describe "writeStartupErrorFile" -Tag Io {
    BeforeEach {
        $script:savedLocalAppData = $env:LOCALAPPDATA
        $script:savedTemp = $env:TEMP
    }
    AfterEach {
        $env:LOCALAPPDATA = $script:savedLocalAppData
        $env:TEMP = $script:savedTemp
    }

    It "既定の置き場所（LOCALAPPDATA の下）に書ければ、そのパスを返す" {
        $env:LOCALAPPDATA = Join-Path $TestDrive "ok-local"
        $env:TEMP = Join-Path $TestDrive "ok-temp"
        $expected = Join-Path $env:LOCALAPPDATA "tebunko\startup_error.txt"

        writeStartupErrorFile "内容" | Should -Be $expected
        (Get-Content -LiteralPath $expected -Raw) | Should -Match "内容"
    }

    It "LOCALAPPDATA の下に書けなければ、TEMP の下に書く" {
        # LOCALAPPDATA 自体をファイルにして、配下にフォルダを作れなくする
        $blockedLocal = Join-Path $TestDrive "blocked-local.txt"
        [System.IO.File]::WriteAllText($blockedLocal, "x")
        $env:LOCALAPPDATA = $blockedLocal
        $env:TEMP = Join-Path $TestDrive "fallback-temp"
        $expected = Join-Path $env:TEMP "tebunko_startup_error.txt"

        writeStartupErrorFile "内容" | Should -Be $expected
        (Get-Content -LiteralPath $expected -Raw) | Should -Match "内容"
    }

    It "どちらにも書けなければ `$null を返す" {
        $blockedLocal = Join-Path $TestDrive "blocked-local2.txt"
        $blockedTemp = Join-Path $TestDrive "blocked-temp2.txt"
        [System.IO.File]::WriteAllText($blockedLocal, "x")
        [System.IO.File]::WriteAllText($blockedTemp, "x")
        $env:LOCALAPPDATA = $blockedLocal
        $env:TEMP = $blockedTemp

        writeStartupErrorFile "内容" | Should -Be $null
    }
}

Describe "reportStartupFailure（記録の経路の切り替え）" -Tag Gui {
    # writeErrorLog（app_host.ps1）を読み込む前と後で、記録の経路が変わることを確かめる。
    # メッセージボックスは本物を出す（.NET の静的メソッドは差し替えられないため）が、
    # ウィンドウを見つけて閉じる操作はしない（この検証環境では raw な user32 の FindWindow が
    # 確実に失敗し、閉じられないまま止まり続けるため）。別プロセスで動かし、
    # 記録が書けたことを確認したらプロセスごと止める
    BeforeAll {
        function startReportStartupFailureProcess {
            param ([string[]]$setupLines, [string]$message)
            $scriptFile = Join-Path $TestDrive "scenario_$([guid]::NewGuid().ToString('N')).ps1"
            $lines = @(
                "Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase"
                ". '$scriptsDir\tebunko\ui\startup_error_view.ps1'"
                ". '$scriptsDir\tebunko\ui\startup_error.ps1'"
            ) + $setupLines + @(
                "try { throw '$message' } catch { reportStartupFailure `$_ | Out-Null }"
            )
            Set-Content -LiteralPath $scriptFile -Value $lines -Encoding UTF8
            return Start-Process -FilePath "powershell.exe" `
                -ArgumentList @("-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", $scriptFile) `
                -PassThru -WindowStyle Hidden
        }

        function waitForFile {
            param ([string]$path, [int]$timeoutSeconds = 10)
            $watch = [System.Diagnostics.Stopwatch]::StartNew()
            while (-not (Test-Path -LiteralPath $path) -and $watch.Elapsed.TotalSeconds -lt $timeoutSeconds) {
                Start-Sleep -Milliseconds 100
            }
            return Test-Path -LiteralPath $path
        }
    }

    AfterEach {
        if ($script:scenarioProcess -and -not $script:scenarioProcess.HasExited) {
            Stop-Process -Id $script:scenarioProcess.Id -Force -ErrorAction SilentlyContinue
        }
    }

    It "app_host.ps1 を読み込む前は、固定の場所（LOCALAPPDATA）に記録する" {
        $localAppData = Join-Path $TestDrive "early-local"
        $setup = @("`$env:LOCALAPPDATA = '$localAppData'")
        $script:scenarioProcess = startReportStartupFailureProcess $setup "読み込み前の失敗"

        $recordFile = Join-Path $localAppData "tebunko\startup_error.txt"
        (waitForFile $recordFile) | Should -Be $true
        (Get-Content -LiteralPath $recordFile -Raw) | Should -Match "読み込み前の失敗"
    }

    It "writeErrorLog（app_host.ps1）が使えるときは、そちらに記録する" {
        $logFile = Join-Path $TestDrive "guierr\gui_error_log.txt"
        $setup = @(
            ". '$scriptsDir\shared\ui\app_host.ps1'"
            "function getGuiErrorLogFile { '$logFile' }"
        )
        $script:scenarioProcess = startReportStartupFailureProcess $setup "読み込み後の失敗"

        (waitForFile $logFile) | Should -Be $true
        (Get-Content -LiteralPath $logFile -Raw) | Should -Match "読み込み後の失敗"
    }
}
