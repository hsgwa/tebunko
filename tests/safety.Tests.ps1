# 安全性説明書（docs/04_安全性.md）の主張を機械的に検査するテスト。
# 「危険な処理・ライブラリを使っていない」ことを、将来の変更で崩れたら失敗する形で固定する。
# Pester 3.4 以降で実行: Invoke-Pester .\tests\safety.Tests.ps1
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptDir = Join-Path (Split-Path -Parent $here) "scripts"
$launcher = Join-Path (Split-Path -Parent $here) "win_grep.bat"

function getCodeLines {
    # 検査対象のコード行を @{ File; Line; Text } で返す。
    # 説明のコメント（「以前は Add-Type でコンパイルしていた」等）を検査対象に含めないよう、コメントを取り除く
    param (
        [string[]]$paths
    )

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($path in $paths) {
        $number = 0
        foreach ($text in [System.IO.File]::ReadAllLines($path)) {
            $number++
            $code = $text
            if ($code.TrimStart().StartsWith("#")) {
                continue
            }
            # 行の途中のコメントを落とす（引用符の中の # は残す）
            $sharp = $code.IndexOf("#")
            while ($sharp -gt 0) {
                $before = $code.Substring(0, $sharp)
                $quotes = @($before.ToCharArray() | Where-Object { $_ -eq '"' }).Count
                $singles = @($before.ToCharArray() | Where-Object { $_ -eq "'" }).Count
                if (($quotes % 2) -eq 0 -and ($singles % 2) -eq 0) {
                    $code = $before
                    break
                }
                $sharp = $code.IndexOf("#", $sharp + 1)
            }
            if ($code.Trim() -eq "") {
                continue
            }
            $result.Add([pscustomobject]@{
                File = [System.IO.Path]::GetFileName($path)
                Line = $number
                Text = $code
            })
        }
    }
    return $result.ToArray()
}

function findPattern {
    # 指定した正規表現に当たったコード行を "ファイル:行番号" の一覧（文字列）で返す。
    # 1 件も無ければ空文字列。Pester の失敗メッセージに場所が出るよう文字列で返す
    param (
        [object[]]$lines,
        [string]$pattern
    )

    return (@($lines | Where-Object { $_.Text -match $pattern } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ")
}

$scriptFiles = @(Get-ChildItem -LiteralPath $scriptDir -Filter *.ps1 | ForEach-Object { $_.FullName })
$code = getCodeLines $scriptFiles

Describe "危険な処理を使っていないこと（docs/04_安全性.md 2.1）" {
    It "スクリプトが 4 本すべて検査対象になっている" {
        # 検査の取りこぼし（対象 0 件で全項目が通る）を防ぐ
        $scriptFiles.Count | Should Be 4
        ($code.Count -gt 1000) | Should Be $true
    }

    It "文字列を式として実行しない（Invoke-Expression・iex・ScriptBlock の生成）" {
        (findPattern $code 'Invoke-Expression|[^-\w]iex[ (]|ScriptBlock\]::Create') | Should Be ""
    }

    It "難読化したコマンドを実行しない（Base64・EncodedCommand）" {
        (findPattern $code 'FromBase64String|EncodedCommand') | Should Be ""
    }

    It "ネットワーク通信を行わない" {
        (findPattern $code 'Invoke-WebRequest|Invoke-RestMethod|WebClient|HttpClient|Net\.Sockets|Start-BitsTransfer|DownloadFile|DownloadString|System\.Net\.') | Should Be ""
    }

    It "Windows API を直接呼び出さない（P/Invoke）" {
        (findPattern $code 'DllImport|GetDelegateForFunctionPointer') | Should Be ""
    }

    It "実行時にコードをコンパイルしない（Add-Type は標準アセンブリの読み込みだけ）" {
        $addType = @($code | Where-Object { $_.Text -match 'Add-Type' })
        # Add-Type がある行は、すべて -AssemblyName（アセンブリの読み込み）であること
        (@($addType | Where-Object { $_.Text -notmatch '-AssemblyName' } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should Be ""
        ($addType.Count -gt 0) | Should Be $true
    }

    It "レジストリを読み書きしない" {
        (findPattern $code 'HKLM|HKCU|HKEY_|Set-ItemProperty|New-ItemProperty|Registry::') | Should Be ""
    }

    It "権限・サービス・自動起動を変更しない" {
        (findPattern $code 'Set-Acl|icacls|schtasks|New-Service|Start-Service|sc\.exe|-Verb\s+RunAs') | Should Be ""
    }

    It "実行ポリシーを恒久変更せず、Bypass も使わない" {
        (findPattern $code 'Set-ExecutionPolicy') | Should Be ""
        (findPattern $code 'ExecutionPolicy\s+Bypass') | Should Be ""
        # 起動用 .bat も同じ（RemoteSigned で起動する）
        $bat = [System.IO.File]::ReadAllText($launcher)
        ($bat -match 'Bypass') | Should Be $false
        ($bat -match 'ExecutionPolicy RemoteSigned') | Should Be $true
    }

    It "資格情報を入力要求・保存しない" {
        (findPattern $code 'Get-Credential|ConvertTo-SecureString|PSCredential') | Should Be ""
    }

    It "リモート実行を行わない" {
        (findPattern $code 'Invoke-Command|New-PSSession|Enter-PSSession|WinRM') | Should Be ""
    }

    It "外部プロセスの起動は explorer.exe と自分自身（powershell.exe）だけ" {
        $starts = @($code | Where-Object { $_.Text -match 'Start-Process' })
        (@($starts | Where-Object { $_.Text -notmatch 'explorer\.exe|powershell\.exe' } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should Be ""
    }

    It "プロセスの強制終了は common.ps1 の 1 か所だけ（画面の［9 プロセス停止］）" {
        $stops = @($code | Where-Object { $_.Text -match 'Stop-Process' })
        $stops.Count | Should Be 1
        $stops[0].File | Should Be "common.ps1"
    }
}

Describe "Office ファイルを安全に開くこと（docs/04_安全性.md 2.2）" {
    $convert = @($code | Where-Object { $_.File -eq "office_to_tsv.ps1" })

    It "マクロを強制的に無効にしてから開く（AutomationSecurity = 3）" {
        (findPattern $convert 'AutomationSecurity\s*=\s*3') | Should Not Be ""
    }

    It "イベントマクロを発火させない（EnableEvents = false）" {
        (findPattern $convert 'EnableEvents\s*=\s*\$false') | Should Not Be ""
    }

    It "外部リンクを更新しない（AskToUpdateLinks = false・Open の UpdateLinks = 0）" {
        (findPattern $convert 'AskToUpdateLinks\s*=\s*\$false') | Should Not Be ""
        (findPattern $convert 'Workbooks\.Open|\$workbooks\.Open') | Should Not Be ""
        (findPattern $convert '\.Open\(\$openPath,\s*0,\s*\$true') | Should Not Be ""
    }

    It "変換処理の Office は画面に出さない（Visible = false）" {
        (findPattern $convert 'Visible\s*=\s*\$false') | Should Not Be ""
        # 可視にするのは画面から元のファイルを開くときだけ（config_gui.ps1）
        $visible = @($code | Where-Object { $_.Text -match 'Visible\s*=\s*\$true' })
        (@($visible | Where-Object { $_.File -ne "config_gui.ps1" } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should Be ""
    }
}

Describe "変換対象のファイルを書き換えないこと（docs/04_安全性.md 3.2）" {
    It "元のファイルのパスを書き込み・削除の API に渡さない" {
        # 書き込み・削除の呼び出し行に、変換対象（原本）を指す変数が現れないこと。
        # 原本は作業フォルダへコピーしてから開くため、書き込み先は常にコピー側（$tmpPath・$destPath・$copyPath 等）になる。
        # $sourcePath = 変換する原本、$targetFolder・$folder.Path = 変換対象フォルダ、$row.SourcePath = 検索結果の元のファイル
        $writes = @($code | Where-Object { $_.Text -match 'Remove-Item|WriteAllText|WriteAllLines|StreamWriter|\.SaveAs|\[System\.IO\.(File|Directory)\]::Move|Move-Item' })
        ($writes.Count -gt 0) | Should Be $true
        (@($writes | Where-Object { $_.Text -match '\$sourcePath|\$targetFolder|\$row\.SourcePath|\$folder\.Path' } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should Be ""
    }

    It "原本を開くのはコピーを作る処理（copyFileShared）の読み取りだけ" {
        # $sourcePath（原本）を渡す先が、コピーの作成と Office での変換（コピー側へ保存）に限られること
        $uses = @($code | Where-Object { $_.Text -match '\$sourcePath' -and $_.Text -notmatch '^\s*(param|\[string\]\$sourcePath)' })
        ($uses.Count -gt 0) | Should Be $true
        $bad = @($uses | Where-Object {
            $_.Text -match 'Remove-Item|WriteAllText|WriteAllLines|StreamWriter|Move-Item|\[System\.IO\.(File|Directory)\]::(Move|Delete|WriteAll)'
        } | ForEach-Object { "$($_.File):$($_.Line)" })
        ($bad -join ", ") | Should Be ""
        # コピーの作成に使う関数は共有違反を避けて読むだけ（FileShare 指定の読み取り）
        (findPattern $code 'function copyFileShared') | Should Not Be ""
    }

    It "Office の SaveAs の保存先は作業フォルダのパスだけ" {
        $saves = @($code | Where-Object { $_.Text -match '\.SaveAs' })
        $saves.Count | Should Be 3
        (@($saves | Where-Object { $_.Text -notmatch 'SaveAs2?\(\$(tmpPath|destPath)' } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should Be ""
    }

    It "原本は読み取り専用で開く（Excel・Word・PowerPoint）" {
        $convert = @($code | Where-Object { $_.File -eq "office_to_tsv.ps1" })
        # Excel: Open の第 3 引数 ReadOnly = $true / Word: 第 3 引数 ReadOnly = $true / PowerPoint: 第 2 引数 ReadOnly = -1
        (findPattern $convert '\.Open\(\$openPath,\s*0,\s*\$true') | Should Not Be ""
        (findPattern $convert '\$documents\.Open\(\$sourcePath,\s*\$false,\s*\$true') | Should Not Be ""
        (findPattern $convert '\$presentations\.Open\("\$\{sourcePath\}::dummy::",\s*-1') | Should Not Be ""
    }
}

Describe "書き込み先が限られていること（docs/04_安全性.md 3.1）" {
    $common = @($code | Where-Object { $_.File -eq "common.ps1" })

    It "書き込みに使うフォルダの定義は work 配下と TEMP 配下だけ" {
        (findPattern $common '\$\{workDir\}\s*=\s*"\$\{rootDir\}\\work"') | Should Not Be ""
        (findPattern $common '\$\{indexDir\}\s*=\s*"\$\{workDir\}\\index"') | Should Not Be ""
        (findPattern $common '\$\{tmpDir\}\s*=\s*Join-Path\s*\(\[System\.IO\.Path\]::GetTempPath\(\)\)\s*"win_grep\\\$\{PID\}"') | Should Not Be ""
        (findPattern $common '\$\{publishDir\}\s*=\s*"\$\{workDir\}\\') | Should Not Be ""
    }

    It "ドライブ直下・システムフォルダを直接指す書き込み先が無い" {
        (findPattern $code '"[A-Za-z]:\\(Windows|Program Files|Users)') | Should Be ""
        (findPattern $code 'GetFolderPath\("?(System|Windows|ProgramFiles|Startup)') | Should Be ""
    }
}
