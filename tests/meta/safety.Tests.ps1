# 安全性説明書（docs/04_安全性.md）の主張を機械的に検査するテスト。
# 「危険な処理・ライブラリを使っていない」ことを、将来の変更で崩れたら失敗する形で固定する。
# スクリプトを読むだけなので、Office もテストデータも要らない。
$here = (Resolve-Path "$PSScriptRoot\..").Path
$rootDir = (Resolve-Path "$here\..").Path
$scriptsDir = "$rootDir\scripts"
$launcher = "$rootDir\tebunko_grep.bat"

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

$scriptFiles = @(Get-ChildItem -LiteralPath $scriptsDir -Recurse -Filter *.ps1 | ForEach-Object { $_.FullName })
$code = getCodeLines $scriptFiles

Describe "危険な処理を使っていないこと（docs/04_安全性.md 2.1）" -Tag Meta {
    It "scripts 配下のスクリプトがすべて検査対象になっている" {
        # 検査の取りこぼし（対象 0 件で全項目が通る）を防ぐ
        ($scriptFiles.Count -ge 30) | Should Be $true
        ($code.Count -gt 3000) | Should Be $true
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

    It "内部の型（NonPublic）をリフレクションで呼ぶのは、フォルダ選択の 1 か所だけ" {
        # Windows 標準のフォルダ選択を、実行時コンパイルなしで開くため（docs/04_安全性.md 2.1）
        (@($code | Where-Object { $_.Text -match 'NonPublic|Reflection\.BindingFlags' -and $_.File -ne "folder_dialog.ps1" } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should Be ""
        (@($code | Where-Object { $_.File -eq "folder_dialog.ps1" -and $_.Text -match 'NonPublic' }).Count) | Should Be 1
    }

    It "実行時にコードをコンパイルしない（Add-Type は標準アセンブリの読み込みだけ）" {
        $addType = @($code | Where-Object { $_.Text -match 'Add-Type' })
        # Add-Type がある行は、すべて -AssemblyName（アセンブリの読み込み）であること
        (@($addType | Where-Object { $_.Text -notmatch '-AssemblyName' } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should Be ""
        ($addType.Count -gt 0) | Should Be $true
    }

    It "Windows Search への問い合わせは windows_search.ps1 だけで行い、SELECT だけを送る" {
        # 高速検索（docs/02_検索.md）は OLE DB の Search.CollatorDSO で読み取るだけ。ほかのファイルからは DB に触らない
        (@($code | Where-Object { $_.Text -match 'OleDb|CommandText' -and $_.File -ne "windows_search.ps1" } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should Be ""
        (@($code | Where-Object { $_.File -eq "windows_search.ps1" -and $_.Text -match '\^\\s\*SELECT' }).Count -gt 0) | Should Be $true
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

    It "壊れたハッシュ（MD5・SHA-1）と、FIPS 準拠でないハッシュの実装を使わない" {
        # FIPS モードの Windows では、FIPS 準拠でない実装（MD5・*Managed）を作ると例外になり起動できなくなる
        (findPattern $code 'MD5|SHA1|RIPEMD|SHA(256|384|512)Managed|HashAlgorithm\]::Create') | Should Be ""
    }

    It "リモート実行を行わない" {
        (findPattern $code 'Invoke-Command|New-PSSession|Enter-PSSession|WinRM') | Should Be ""
    }

    It "外部プロセスの起動は explorer.exe と自分自身（powershell.exe）だけ" {
        $starts = @($code | Where-Object { $_.Text -match 'Start-Process' })
        (@($starts | Where-Object { $_.Text -notmatch 'explorer\.exe|powershell\.exe' } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should Be ""
    }

    It "プロセスの強制終了は office_process.ps1 の 1 か所だけ（画面の［9 プロセス停止］）" {
        $stops = @($code | Where-Object { $_.Text -match 'Stop-Process' })
        $stops.Count | Should Be 1
        $stops[0].File | Should Be "office_process.ps1"
    }
}

Describe "Office ファイルを安全に開くこと（docs/04_安全性.md 2.2）" -Tag Meta {
    $app = @($code | Where-Object { $_.File -eq "office_app.ps1" })
    $extract = @($code | Where-Object { $_.File -eq "extract_office.ps1" })

    It "マクロを強制的に無効にしてから開く（AutomationSecurity = 3）" {
        (findPattern $app 'AutomationSecurity\s*=\s*3') | Should Not Be ""
    }

    It "イベントマクロを発火させない（EnableEvents = false）" {
        (findPattern $app 'EnableEvents\s*=\s*\$false') | Should Not Be ""
    }

    It "外部リンクを更新しない（AskToUpdateLinks = false・Open の UpdateLinks = 0）" {
        (findPattern $app 'AskToUpdateLinks\s*=\s*\$false') | Should Not Be ""
        (findPattern $extract '\.Open\(\$openPath,\s*0,\s*\$true') | Should Not Be ""
    }

    It "インデクサの Office は画面に出さない（Visible = false）" {
        (findPattern $app 'Visible\s*=\s*\$false') | Should Not Be ""
        # 可視にするのは画面から元のファイルを開くときだけ（ui/open_source.ps1）
        $visible = @($code | Where-Object { $_.Text -match 'Visible\s*=\s*\$true' })
        (@($visible | Where-Object { $_.File -ne "open_source.ps1" } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should Be ""
    }
}

Describe "取り込み対象のファイルを書き換えないこと（docs/04_安全性.md 3.2）" -Tag Meta {
    It "元のファイルのパスを書き込み・削除の API に渡さない" {
        # 書き込み・削除の呼び出し行に、取り込み対象（原本）を指す変数が現れないこと。
        # 原本は作業フォルダへコピーしてから開くため、書き込み先は常にコピー側（$tmpPath・$destPath・$copyPath 等）になる。
        # $targetFolder・$folder.Path = クロール対象フォルダ、$row.SourcePath = 検索結果の元のファイル
        $writes = @($code | Where-Object { $_.Text -match 'Remove-Item|WriteAllText|WriteAllLines|StreamWriter|\.SaveAs|\[System\.IO\.(File|Directory)\]::Move|Move-Item' })
        ($writes.Count -gt 0) | Should Be $true
        (@($writes | Where-Object { $_.Text -match '\$targetFolder|\$row\.SourcePath|\$folder\.Path' } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should Be ""
    }

    It "原本を読むのはコピーを作る処理（copyFileShared）だけ" {
        # copyFileShared は読み取りだけで開き、コピー先（作業フォルダ）へ書く
        (findPattern $code 'function copyFileShared') | Should Not Be ""
        $fs = @($code | Where-Object { $_.File -eq "fs.ps1" -and $_.Text -match '\$sourcePath' })
        (@($fs | Where-Object {
            $_.Text -match 'Remove-Item|WriteAllText|WriteAllLines|StreamWriter|Move-Item|\[System\.IO\.(File|Directory)\]::(Move|Delete|WriteAll)'
        } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should Be ""
        (findPattern $fs '\[System\.IO\.FileAccess\]::Read') | Should Not Be ""
    }

    It "Office の SaveAs の保存先は作業フォルダのパスだけ" {
        $saves = @($code | Where-Object { $_.Text -match '\.SaveAs' })
        $saves.Count | Should Be 3
        (@($saves | Where-Object { $_.Text -notmatch 'SaveAs2?\(\$(tmpPath|destPath)' } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should Be ""
    }

    It "原本は読み取り専用で開く（Excel・Word・PowerPoint）" {
        $extract = @($code | Where-Object { $_.File -eq "extract_office.ps1" })
        # Excel: Open の第 3 引数 ReadOnly = $true / Word: 第 3 引数 ReadOnly = $true / PowerPoint: 第 2 引数 ReadOnly = -1
        (findPattern $extract '\.Open\(\$openPath,\s*0,\s*\$true') | Should Not Be ""
        (findPattern $extract '\$documents\.Open\(\$sourcePath,\s*\$false,\s*\$true') | Should Not Be ""
        (findPattern $extract '\$presentations\.Open\("\$\{sourcePath\}::dummy::",\s*-1') | Should Not Be ""
    }
}

Describe "書き込み先が限られていること（docs/04_安全性.md 3.1）" -Tag Meta {
    It "書き込みに使うフォルダの定義は、データの置き場所（設定ファイル・work）と TEMP 配下だけ" {
        $paths = @($code | Where-Object { $_.File -in @("paths.ps1", "data_dir.ps1", "paths_grep.ps1", "settings_grep.ps1") })
        # データの置き場所は、ツールのフォルダか、書き込めないときの %LOCALAPPDATA%\tebunko\<鍵>
        (findPattern $paths '\$\{dataDir\}\s*=\s*getDataDir') | Should Not Be ""
        (findPattern $paths 'GetFolderPath\("LocalApplicationData"\)') | Should Not Be ""
        (findPattern $paths '\$\{workDir\}\s*=\s*"\$\{dataDir\}\\work"') | Should Not Be ""
        (findPattern $paths '\$\{workDir\}\s*=\s*getWorkDir') | Should Not Be ""
        (findPattern $paths '\$\{indexDir\}\s*=\s*"\$\{workDir\}\\index"') | Should Not Be ""
        (findPattern $paths '\$\{tmpDir\}\s*=\s*Join-Path\s*\(\[System\.IO\.Path\]::GetTempPath\(\)\)\s*"tebunko_grep\\\$\{PID\}"') | Should Not Be ""
        (findPattern $paths '\$\{publishDir\}\s*=\s*"\$\{workDir\}\\') | Should Not Be ""
        (findPattern $paths '\$\{settingsFile\}\s*=\s*"\$\{dataDir\}\\setting\.config"') | Should Not Be ""
    }

    It "ドライブ直下・システムフォルダを直接指す書き込み先が無い" {
        (findPattern $code '"[A-Za-z]:\\(Windows|Program Files|Users)') | Should Be ""
        (findPattern $code 'GetFolderPath\("?(System|Windows|ProgramFiles|Startup)') | Should Be ""
    }

    It "異常終了で残った作業フォルダを次回起動時に回収する" {
        # %TEMP%\tebunko_grep\<PID> に原本のコピーが残り続けないこと（docs/04_安全性.md 4.4）
        (findPattern $code 'function removeStaleTmpDirs') | Should Not Be ""
        (findPattern $code '^removeStaleTmpDirs') | Should Not Be ""
    }
}

Describe "サードパーティの静的解析（docs/04_安全性.md 5.2）" -Tag Meta {
    # PSScriptAnalyzer（Microsoft 提供）で検査する。未導入の環境では飛ばす:
    #   Install-Module PSScriptAnalyzer -Scope CurrentUser
    $hasAnalyzer = (@(Get-Module -ListAvailable PSScriptAnalyzer).Count -gt 0)
    if ($hasAnalyzer) {
        Import-Module PSScriptAnalyzer -ErrorAction SilentlyContinue
    }
    $securitySettings = "$PSScriptRoot\PSScriptAnalyzer.security.psd1"

    # 継承元の型が別ファイルにある場合の TypeNotFound は除く（読み込む順で解決する。
    # tests\meta\structure.Tests.ps1 の「スクリプトの構文」と同じ扱い）
    function formatFindings {
        param ($found)
        return (@($found | Where-Object { $_.RuleName -ne "TypeNotFound" } | ForEach-Object { "$($_.RuleName) $($_.ScriptName):$($_.Line)" }) -join ", ")
    }

    It "安全性にかかわるルールの指摘が 0 件" -Skip:(-not $hasAnalyzer) {
        (formatFindings (Invoke-ScriptAnalyzer -Path $scriptsDir -Recurse -Settings $securitySettings)) | Should Be ""
    }

    It "Error 重大度の指摘が 0 件（全ルール）" -Skip:(-not $hasAnalyzer) {
        (formatFindings (Invoke-ScriptAnalyzer -Path $scriptsDir -Recurse -Severity Error)) | Should Be ""
    }

    It "制限言語モード（Constrained Language Mode）で使えない書き方が無い" -Skip:(-not $hasAnalyzer) {
        # AppLocker・WDAC で制限言語モードを強制している環境向け（docs/04_安全性.md 5.4）
        (formatFindings (Invoke-ScriptAnalyzer -Path $scriptsDir -Recurse -IncludeRule PSUseConstrainedLanguageMode)) | Should Be ""
    }

    It "安全性のルール設定に、検査すべきルールが含まれている" {
        # 設定ファイルからルールを消して指摘 0 件にする、という抜け道を防ぐ
        $settings = Import-LocalizedData -BaseDirectory $PSScriptRoot -FileName "PSScriptAnalyzer.security.psd1"
        @($settings.IncludeRules).Count | Should Be 14
        ($settings.IncludeRules -contains "PSAvoidUsingInvokeExpression") | Should Be $true
        ($settings.IncludeRules -contains "PSAvoidUsingPlainTextForPassword") | Should Be $true
        ($settings.IncludeRules -contains "PSAvoidUsingConvertToSecureStringWithPlainText") | Should Be $true
    }
}

Describe "第三者が検証するための資料がそろっていること（docs/04_安全性.md 5.1・6）" -Tag Meta {
    It "安全性説明書がある" {
        (Test-Path -LiteralPath "$rootDir\docs\04_安全性.md") | Should Be $true
    }

    It "脆弱性の連絡先（.github\SECURITY.md）がある" {
        (Test-Path -LiteralPath "$rootDir\.github\SECURITY.md") | Should Be $true
    }

    It "配布物の完全性を確かめる手順（tools\new_release_files.ps1）がある" {
        (Test-Path -LiteralPath "$rootDir\tools\new_release_files.ps1") | Should Be $true
    }

    It "ライセンス（LICENSE）があり、MIT の条文と著作権表示を含む" {
        $licensePath = "$rootDir\LICENSE"
        (Test-Path -LiteralPath $licensePath) | Should Be $true
        $license = [System.IO.File]::ReadAllText($licensePath)
        # 表題・著作権表示・許諾条項・無保証条項がそろっていること（MIT の全文）
        ($license -like "MIT License*") | Should Be $true
        ($license -match "Copyright \(c\) \d{4} \S+") | Should Be $true
        ($license -like "*Permission is hereby granted, free of charge*") | Should Be $true
        ($license -like "*WITHOUT WARRANTY OF ANY KIND*") | Should Be $true
        # 実名を含めない（AGENTS.md「個人情報を書かない」）。著作権者はアカウント名で表記する
        ($license -like "*hsgwa*") | Should Be $true
    }

    It "部品表（SBOM）があり、第三者の部品を 1 件も含まない" {
        $sbomPath = "$rootDir\sbom.cdx.json"
        (Test-Path -LiteralPath $sbomPath) | Should Be $true
        $sbom = [System.IO.File]::ReadAllText($sbomPath) | ConvertFrom-Json
        $sbom.bomFormat | Should Be "CycloneDX"
        $sbom.specVersion | Should Be "1.6"
        # 本ツール自身のライセンスは SPDX の識別子で記載する（LICENSE と一致させる）
        $sbom.metadata.component.licenses[0].license.id | Should Be "MIT"
        # 構成物は本ツール自身のファイルだけ。パッケージマネージャー由来の部品（purl を持つ）は無い
        (@($sbom.components).Count -gt 0) | Should Be $true
        @($sbom.components | Where-Object { $_.group -ne "tebunko" }).Count | Should Be 0
        @($sbom.components | Where-Object { $_.purl }).Count | Should Be 0
        # 前提ソフトウェア（同梱しないもの）は metadata.properties に記載する
        @($sbom.metadata.properties | Where-Object { $_.name -eq "tebunko:prerequisite" }).Count -gt 0 | Should Be $true
    }
}
