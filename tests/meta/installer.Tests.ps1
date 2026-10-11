# インストーラー（installer\）の決まりを確かめるテスト（docs/safety/disclosure.md「インストーラー版」）。
# インストーラーそのもののビルド（Inno Setup）は release.yml で行う。ここでは、起動口 tebunko.exe がビルドできることと、
# インストーラーが書き込む先・入れるものを広げていないことを、スクリプトを読んで確かめる。
BeforeDiscovery {
    . "$PSScriptRoot\..\helpers\path_chars.ps1"
    $pathCharCases = getPathCharCases
}

BeforeAll {
    $rootDir = (Resolve-Path "$PSScriptRoot\..\..").Path
    $launcherSource = "$rootDir\installer\tebunko.cs"
    $setupScript = "$rootDir\installer\tebunko.iss"

    function getIssSection {
        # Inno Setup のスクリプトの [名前] の節の行（コメントと空行を除く）
        param (
            [string]$name
        )

        $lines = @()
        $inSection = $false
        foreach ($line in [System.IO.File]::ReadAllLines($setupScript)) {
            $text = $line.Trim()
            if ($text -match '^\[(.+)\]$') {
                $inSection = ($Matches[1] -eq $name)
                continue
            }
            if ($inSection -and $text -ne "" -and !$text.StartsWith(";")) {
                $lines += $text
            }
        }
        return $lines
    }
}

Describe "起動口 tebunko.exe（installer\tebunko.cs）" -Tag Meta {
    BeforeAll {
        $source = [System.IO.File]::ReadAllText($launcherSource)
    }

    It "Windows 標準の csc.exe で、警告なしにビルドできる" {
        $csc = Join-Path ([System.Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()) "csc.exe"
        $out = Join-Path $TestDrive "tebunko.exe"
        $log = & $csc -nologo -codepage:65001 -target:winexe -warnaserror+ "-out:$out" $launcherSource 2>&1
        ($log -join "`n") | Should -Be ""
        (Test-Path -LiteralPath $out) | Should -Be $true
    }

    It "起動するのは Windows の PowerShell 5.1 だけで、実行ポリシーは RemoteSigned（Bypass を使わない）" {
        @([regex]::Matches($source, 'new ProcessStartInfo\(')).Count | Should -Be 1
        $source -match 'WindowsPowerShell\\v1\.0\\powershell\.exe' | Should -Be $true
        $source -match '-ExecutionPolicy RemoteSigned -Command' | Should -Be $true
        $source -match 'Bypass|Unrestricted|EncodedCommand' | Should -Be $false
    }

    It "Windows API を直接呼ばない（P/Invoke）・レジストリに触らない" {
        $source -match 'DllImport|Microsoft\.Win32|Registry' | Should -Be $false
    }

    It "ミューテックスの名前が、インストーラーの AppMutex と同じ" {
        $source -match 'MutexName = "([^"]+)"' | Should -Be $true
        $name = $Matches[1]
        @(getIssSection "Setup") -contains "AppMutex=$name" | Should -Be $true
    }
}

Describe "起動口 tebunko.exe が渡す引数（フォルダ名の文字）" -Tag Meta {
    # tebunko.exe は gui.ps1 のパスを環境変数で渡し、-Command で WildcardPattern.Escape を通して呼ぶ。
    # ' や [ ] などを含むフォルダでも起動できるはず。ビルドした exe の BuildArguments の返す引数で、実際に確かめる
    BeforeAll {
        $csc = Join-Path ([System.Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()) "csc.exe"
        $out = Join-Path $TestDrive "tebunko_args.exe"
        & $csc -nologo -codepage:65001 -target:winexe "-out:$out" $launcherSource | Out-Null
        if (-not [System.IO.File]::Exists($out)) { throw "tebunko.exe をビルドできなかった" }
        # ファイルを掴まないよう、バイト列から読み込む
        $assembly = [System.Reflection.Assembly]::Load([System.IO.File]::ReadAllBytes($out))
        $buildArguments = $assembly.GetType("Program").GetMethod("BuildArguments", [System.Reflection.BindingFlags]"NonPublic,Public,Static")
        $powershell = Join-Path ([Environment]::GetFolderPath("System")) "WindowsPowerShell\v1.0\powershell.exe"
        $script:caseCount = 0
        $script:trialRoots = @()
    }

    AfterAll {
        foreach ($root in $script:trialRoots) {
            if ([System.IO.Directory]::Exists($root)) { [System.IO.Directory]::Delete($root, $true) }
        }
    }

    It "フォルダ名に <label> を含んでも、tebunko.exe の引数で gui.ps1 が起動する" -TestCases $pathCharCases {
        param ($label, $ch)
        $script:caseCount++
        # TestDrive は片付けに -Path を使い、[ を含む名前を消せないので、TEMP に作って .NET で消す
        $script:trialRoots += ,(Join-Path ([System.IO.Path]::GetTempPath()) ("exe_" + [guid]::NewGuid().ToString("N")))
        $guiDir = Join-Path $script:trialRoots[-1] "a${ch}b\scripts\tebunko"
        [System.IO.Directory]::CreateDirectory($guiDir) | Out-Null
        $gui = "$guiDir\gui.ps1"
        [System.IO.File]::WriteAllText($gui, "Set-Content -LiteralPath (Join-Path `$PSScriptRoot 'marker.txt') -Value `$PSScriptRoot -Encoding UTF8`r`nexit 0`r`n", (New-Object System.Text.UTF8Encoding($true)))

        $psi = New-Object System.Diagnostics.ProcessStartInfo($powershell, $buildArguments.Invoke($null, @()))
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.EnvironmentVariables["TEBUNKO_GUI"] = $gui
        $process = [System.Diagnostics.Process]::Start($psi)
        try {
            $process.WaitForExit(20000) | Out-Null
        } finally {
            if (-not $process.HasExited) { Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue }
        }

        $marker = "$guiDir\marker.txt"
        [System.IO.File]::Exists($marker) | Should -Be $true -Because "gui.ps1 が起動しなかった"
        (Get-Content -LiteralPath $marker -Encoding UTF8) | Should -BeExactly $guiDir
    }
}

Describe "起動口 tebunko.exe の終了コード（gui.ps1 の終わり方）" -Tag Meta {
    # 以前の -File と同じ終わり方にする: 読み込めない・実行中に例外なら 0 以外でエラー出力あり、
    # 普通に終われば（途中の外部コマンドが 0 以外で終わっていても）0、exit N なら N
    BeforeDiscovery {
        $endCases = @(
            @{ label = "構文エラーで読み込めない"; body = "if ("; code = 1; hasError = $true }
            @{ label = "実行中に例外"; body = "throw 'boom'"; code = 1; hasError = $true }
            @{ label = "普通に終わる"; body = "Write-Output ok"; code = 0; hasError = $false }
            @{ label = "exit 3 で終わる"; body = "exit 3"; code = 3; hasError = $false }
            @{ label = "外部コマンドが 0 以外で終わったあと、普通に終わる"; body = "cmd.exe /d /c exit 5`r`nWrite-Output ok"; code = 0; hasError = $false }
        )
    }

    BeforeAll {
        $csc = Join-Path ([System.Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()) "csc.exe"
        $out = Join-Path $TestDrive "tebunko_exit.exe"
        & $csc -nologo -codepage:65001 -target:winexe "-out:$out" $launcherSource | Out-Null
        if (-not [System.IO.File]::Exists($out)) { throw "tebunko.exe をビルドできなかった" }
        $assembly = [System.Reflection.Assembly]::Load([System.IO.File]::ReadAllBytes($out))
        $buildArguments = $assembly.GetType("Program").GetMethod("BuildArguments", [System.Reflection.BindingFlags]"NonPublic,Public,Static")
        $powershell = Join-Path ([Environment]::GetFolderPath("System")) "WindowsPowerShell\v1.0\powershell.exe"
        $script:endRoots = @()
    }

    AfterAll {
        foreach ($root in $script:endRoots) {
            if ([System.IO.Directory]::Exists($root)) { [System.IO.Directory]::Delete($root, $true) }
        }
    }

    It "gui.ps1 が <label> とき、終了コードは <code>、エラー出力は <hasError>（有り）" -ForEach $endCases {
        $root = Join-Path ([System.IO.Path]::GetTempPath()) ("exe_end_" + [guid]::NewGuid().ToString("N"))
        $script:endRoots += ,$root
        [System.IO.Directory]::CreateDirectory($root) | Out-Null
        $gui = Join-Path $root "gui.ps1"
        [System.IO.File]::WriteAllText($gui, $body + "`r`n", (New-Object System.Text.UTF8Encoding($true)))

        $psi = New-Object System.Diagnostics.ProcessStartInfo($powershell, $buildArguments.Invoke($null, @()))
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.RedirectStandardError = $true
        $psi.EnvironmentVariables["TEBUNKO_GUI"] = $gui
        $process = [System.Diagnostics.Process]::Start($psi)
        $errorText = ""
        try {
            $errorText = $process.StandardError.ReadToEnd()
            $process.WaitForExit(20000) | Out-Null
        } finally {
            if (-not $process.HasExited) { Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue }
        }

        $process.ExitCode | Should -Be $code
        ($errorText.Trim().Length -gt 0) | Should -Be $hasError
    }
}
Describe "インストーラー（installer\tebunko.iss）" -Tag Meta {
    It "既定は管理者権限なしで入れる" {
        @(getIssSection "Setup") -contains "PrivilegesRequired=lowest" | Should -Be $true
    }

    It "レジストリ・PATH・ファイルの関連付けに書かない（アンインストールの情報は Inno Setup が書く）" {
        @(getIssSection "Registry").Count | Should -Be 0
        @(getIssSection "Setup") | Where-Object { $_ -match '^(ChangesEnvironment|ChangesAssociations)=yes' } | Should -BeNullOrEmpty
    }

    It "入れるのは tebunko.exe・LICENSE・VERSION.txt・scripts\ だけ" {
        $sources = @(getIssSection "Files" | ForEach-Object {
            if ($_ -match 'Source: "\{#StageDir\}\\([^"]+)"') { $Matches[1] }
        })
        ($sources -join ", ") | Should -Be "tebunko.exe, LICENSE, VERSION.txt, scripts\*"
    }

    It "アンインストーラーは、tebunko のものと分かるよう uninstall\ に置く" {
        @(getIssSection "Setup") -contains 'UninstallFilesDir={app}\uninstall' | Should -Be $true
    }

    It "更新のときは、前の版の scripts\ を消してから入れる" {
        @(getIssSection "InstallDelete") -contains 'Type: filesandordirs; Name: "{app}\scripts"' | Should -Be $true
    }

    It "アンインストールのときは、ツールのフォルダに作られる設定と起動失敗の記録も消す" {
        $lines = @(getIssSection "UninstallDelete")
        $lines -contains 'Type: files; Name: "{app}\setting.config"' | Should -Be $true
        $lines -contains 'Type: files; Name: "{app}\startup_error.txt"' | Should -Be $true
    }

    It "起動するのは、入れた tebunko.exe だけ" {
        @(getIssSection "Run" | Where-Object { $_ -notmatch '^Filename: "\{app\}\\tebunko\.exe";' }).Count | Should -Be 0
    }
}

Describe "インストーラーのファイルの文字コードと改行" -Tag Meta {
    # Inno Setup は BOM の無いスクリプトを ANSI として読むため、日本語が化ける
    It "<name> は BOM 付き UTF-8・CRLF" -TestCases @(
        @{ name = "tebunko.cs"; file = "installer\tebunko.cs" }
        @{ name = "tebunko.iss"; file = "installer\tebunko.iss" }
    ) {
        param ($name, $file)
        # 表はテストを探す段階で作られ、BeforeAll の変数がまだ無いため、パスはここで組み立てる
        $path = Join-Path $rootDir $file
        $bytes = [System.IO.File]::ReadAllBytes($path)
        ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) | Should -Be $true
        [regex]::Matches([System.IO.File]::ReadAllText($path), "(?<!`r)`n").Count | Should -Be 0
    }
}
