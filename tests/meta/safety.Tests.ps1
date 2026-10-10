# 安全性説明書（docs/safety/index.md）の主張を機械的に検査するテスト。
# 「危険な処理・ライブラリを使っていない」ことを、将来の変更で崩れたら失敗する形で固定する。
# スクリプトを読むだけなので、Office もテストデータも要らない。
BeforeAll {
    $here = (Resolve-Path "$PSScriptRoot\..").Path
    $rootDir = (Resolve-Path "$here\..").Path
    $scriptsDir = "$rootDir\scripts"
    $launcher = "$rootDir\tebunko.bat"
    # 禁止の語の一覧は tools\script_rules.ps1 と共有する（check_release_package.ps1・new_single_script.ps1 も同じ一覧を見る）
    . "$rootDir\tools\script_rules.ps1"

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
                # コメントの落とし方は tools\script_rules.ps1 の stripLineComment と共有する
                $code = stripLineComment $text
                if ($code.Trim() -eq "") {
                    continue
                }
                $result.Add([pscustomobject]@{
                    File = [System.IO.Path]::GetFileName($path)
                    Path = $path
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

    function getWritePlaceViolations {
        # 書き込み先を環境から得る書き方が、許すものの一覧に収まっているかを調べ、外れたものを文字列の一覧で返す（無ければ空）。
        # $codeLines は getCodeLines の形（コメントを除いた行）、$batText は tebunko.bat の全文。
        # 違反を入れた文字列で判定が落ちることを確かめるため、ファイルを読まずに文字列だけで動くようにしてある。
        # 返す文字列の先頭（最初の ":" まで）は決まりの名前。違反のテストはこの名前で、どの決まりで落ちたかを確かめる。
        # 許す一覧（下の $allowed*）に足すときは、足す理由を行のコメントに書き、docs/safety/file-access.md の書き込む場所も合わせて直す
        param (
            [object[]]$codeLines,
            [string]$batText
        )

        $problems = New-Object System.Collections.Generic.List[string]
        # %LOCALAPPDATA%・%TEMP% を指す文字そのものを書かない
        $forbidden = '(?i)APPDATA|%TEMP%|%TMP%|env:TEMP\b|env:TMP\b'
        foreach ($line in @($codeLines | Where-Object { $_.Text -match $forbidden })) {
            $problems.Add("禁止の語: $($line.File):$($line.Line)")
        }
        if ($batText -match $forbidden) {
            $problems.Add("禁止の語: tebunko.bat")
        }
        # 特別なフォルダを得るのは settings.ps1 の UserProfile（既定のワークスペース）の 1 か所だけ
        $allowedFolderPath = 'settings.ps1|[string]$profileDir = [System.Environment]::GetFolderPath("UserProfile")'
        $folderPath = @($codeLines | Where-Object { $_.Text -match '(?i)GetFolderPath\s*\(' } | ForEach-Object { "$($_.File)|$($_.Text.Trim())" })
        foreach ($entry in @($folderPath | Where-Object { $_ -ne $allowedFolderPath })) {
            $problems.Add("GetFolderPath: $entry")
        }
        # 環境変数は SystemRoot（メモ帳を開く）と、画面のテストが差し込む TEBUNKO_GUI_LEFTOVER_FILE だけ
        $allowedEnv = @("systemroot", "tebunko_gui_leftover_file")
        # 加えて、既定のワークスペースを差し替える TEBUNKO_DEFAULT_WORKSPACE は、settings.ps1 が 1 回だけ読む（テストや実機の確かめが既定のワークスペースに書かないための口）
        $workspaceEnvCount = 0
        foreach ($line in $codeLines) {
            foreach ($match in [regex]::Matches($line.Text, '(?i)\$\{?env:(\w+)')) {
                $name = $match.Groups[1].Value.ToLowerInvariant()
                if ($name -eq "tebunko_default_workspace" -and $line.File -eq "settings.ps1") {
                    $workspaceEnvCount++
                } elseif ($allowedEnv -notcontains $name) {
                    $problems.Add("環境変数: $($line.File):$($line.Line) $($match.Value)")
                }
            }
        }
        if ($workspaceEnvCount -gt 1) {
            $problems.Add("環境変数: TEBUNKO_DEFAULT_WORKSPACE を読む行が $workspaceEnvCount 回")
        }
        foreach ($line in @($codeLines | Where-Object { $_.Text -match '(?i)GetEnvironmentVariable' })) {
            $problems.Add("GetEnvironmentVariable: $($line.File):$($line.Line)")
        }
        # %…% を展開するのは、利用者が入れたフォルダの文字列を扱う 2 か所（normalizeFolderPath・getWorkDir）で、各 1 回だけ
        $allowedExpand = @(
            'folder.ps1|$path = [System.Environment]::ExpandEnvironmentVariables($path).Trim()'
            'settings.ps1|$folder = [System.Environment]::ExpandEnvironmentVariables($folder)'
        )
        $expand = @($codeLines | Where-Object { $_.Text -match '(?i)ExpandEnvironmentVariables' } | ForEach-Object { "$($_.File)|$($_.Text.Trim())" })
        foreach ($entry in @($expand | Where-Object { $allowedExpand -notcontains $_ })) {
            $problems.Add("ExpandEnvironmentVariables: $entry")
        }
        # 許す行も、各 1 回だけ（同じ行が増えても通さない）
        foreach ($entry in $allowedExpand) {
            $count = @($expand | Where-Object { $_ -eq $entry }).Count
            if ($count -gt 1) {
                $problems.Add("ExpandEnvironmentVariables: 許した行が $count 回: $entry")
            }
        }
        # tebunko.bat が使う %…% は、起動に使う変数と SystemRoot だけ
        $allowedBat = @("ps1", "pscmd", "systemroot")
        foreach ($match in [regex]::Matches($batText, '%(\w+)%')) {
            if ($allowedBat -notcontains $match.Groups[1].Value.ToLowerInvariant()) {
                $problems.Add("bat の変数: $($match.Value)")
            }
        }
        if ($batText -match '(?i)\$\{?env:') {
            $problems.Add("bat の環境変数")
        }
        # tebunko.bat の中の PowerShell も、特別なフォルダ・環境変数・一時フォルダを得る API を使わない（許す数は 0）
        foreach ($match in [regex]::Matches($batText, '(?i)GetFolderPath|GetEnvironmentVariable|ExpandEnvironmentVariables|GetTempPath|GetTempFileName|New-TemporaryFile')) {
            $problems.Add("bat の API: $($match.Value)")
        }
        return @($problems)
    }

    $scriptFiles = @(Get-ChildItem -LiteralPath $scriptsDir -Recurse -Filter *.ps1 | ForEach-Object { $_.FullName })
    $code = getCodeLines $scriptFiles
}

Describe "危険な処理を使っていないこと（docs/safety/checks.md「検査項目と結果」）" -Tag Meta {
    It "scripts 配下のスクリプトがすべて検査対象になっている" {
        # 検査の取りこぼし（対象 0 件で全項目が通る）を防ぐ
        ($scriptFiles.Count -ge 30) | Should -Be $true
        ($code.Count -gt 3000) | Should -Be $true
    }

    It "文字列を式として実行しない（Invoke-Expression・iex・ScriptBlock の生成）" {
        (findPattern $code $script:bannedCodePatterns.InvokeExpression) | Should -Be ""
    }

    It "難読化したコマンドを実行しない（Base64・EncodedCommand）" {
        (findPattern $code $script:bannedCodePatterns.EncodedCommand) | Should -Be ""
    }

    It "ネットワーク通信を行わない" {
        (findPattern $code $script:bannedCodePatterns.Network) | Should -Be ""
    }

    It "Windows API を直接呼び出さない（P/Invoke）" {
        (findPattern $code $script:bannedCodePatterns.PInvoke) | Should -Be ""
    }

    It "内部の型（NonPublic）をリフレクションで呼ぶのは、フォルダ選択の 1 か所だけ" {
        # Windows 標準のフォルダ選択を、実行時コンパイルなしで開くため（docs/safety/checks.md「検査項目と結果」）
        (@($code | Where-Object { $_.Text -match 'NonPublic|Reflection\.BindingFlags' -and $_.File -ne "folder_dialog.ps1" } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should -Be ""
        (@($code | Where-Object { $_.File -eq "folder_dialog.ps1" -and $_.Text -match 'NonPublic' }).Count) | Should -Be 1
    }

    It "実行時にコードをコンパイルしない（Add-Type は標準アセンブリの読み込みだけ）" {
        $addType = @($code | Where-Object { $_.Text -match 'Add-Type' })
        # Add-Type がある行は、すべて -AssemblyName（アセンブリの読み込み）であること
        (@($addType | Where-Object { $_.Text -notmatch '-AssemblyName' } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should -Be ""
        ($addType.Count -gt 0) | Should -Be $true
    }

    It "Windows Search への問い合わせは windows_search.ps1 だけで行い、SELECT だけを送る" {
        # 高速検索（docs/design/search/index.md）は OLE DB の Search.CollatorDSO で読み取るだけ。ほかのファイルからは DB に触らない
        (@($code | Where-Object { $_.Text -match 'OleDb|CommandText' -and $_.File -ne "windows_search.ps1" } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should -Be ""
        (@($code | Where-Object { $_.File -eq "windows_search.ps1" -and $_.Text -match '\^\\s\*SELECT' }).Count -gt 0) | Should -Be $true
    }

    It "レジストリを読み書きしない" {
        (findPattern $code $script:bannedCodePatterns.Registry) | Should -Be ""
    }

    It "権限・サービス・自動起動を変更しない" {
        (findPattern $code $script:bannedCodePatterns.PrivilegeService) | Should -Be ""
    }

    It "実行ポリシーを恒久変更せず、Bypass も使わない" {
        (findPattern $code $script:bannedCodePatterns.ExecutionPolicySet) | Should -Be ""
        (findPattern $code $script:bannedCodePatterns.ExecutionPolicyBypass) | Should -Be ""
        # 起動用 .bat も同じ（RemoteSigned で起動する）
        $bat = [System.IO.File]::ReadAllText($launcher)
        ($bat -match 'Bypass') | Should -Be $false
        ($bat -match 'ExecutionPolicy RemoteSigned') | Should -Be $true
    }

    It "資格情報を入力要求・保存しない" {
        (findPattern $code $script:bannedCodePatterns.Credential) | Should -Be ""
    }

    It "壊れたハッシュ（MD5・SHA-1）と、FIPS 準拠でないハッシュの実装を使わない" {
        # FIPS モードの Windows では、FIPS 準拠でない実装（MD5・*Managed）を作ると例外になり起動できなくなる
        (findPattern $code $script:bannedCodePatterns.WeakHash) | Should -Be ""
    }

    It "リモート実行を行わない" {
        (findPattern $code $script:bannedCodePatterns.Remote) | Should -Be ""
    }

    It "外部プロセスの起動は explorer.exe・notepad.exe だけ" {
        # notepad.exe は固定のパス（$env:SystemRoot\System32\notepad.exe）だけを許す（既定のアプリの登録に頼らない）
        $starts = @($code | Where-Object { $_.Text -match 'Start-Process' })
        (@($starts | Where-Object { $_.Text -notmatch 'explorer\.exe' -and $_.Text -notmatch 'SystemRoot.*System32\\notepad\.exe' } |
            ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should -Be ""
    }

    It "ProcessStartInfo・[System.Diagnostics.Process]::Start を使うのは open_source.ps1 の openWithShell だけ（1 か所）" {
        $starts = @($code | Where-Object { $_.Text -match 'ProcessStartInfo|\[System\.Diagnostics\.Process\]::Start\(' })
        $starts.Count | Should -Be 2
        (@($starts | Where-Object { $_.File -ne "open_source.ps1" } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should -Be ""
    }

    It "tebunko.bat が起動する外部のプログラムは conhost.exe・powershell.exe・notepad.exe だけ（起動に失敗したときだけ notepad.exe を開く）" {
        $bat = [System.IO.File]::ReadAllText($launcher)
        $names = @([regex]::Matches($bat, '\b[A-Za-z0-9_]+\.exe\b') | ForEach-Object { $_.Value.ToLowerInvariant() } | Sort-Object -Unique)
        (@($names | Where-Object { $_ -notin @("conhost.exe", "powershell.exe", "notepad.exe") }) -join ", ") | Should -Be ""
        ($names -contains "notepad.exe") | Should -Be $true
    }

    It "プロセスの強制終了は office_process.ps1 の 1 か所だけ（起動時の前回残った Office の確認。止める対象は記録のあるものだけ）" {
        $stops = @($code | Where-Object { $_.Text -match 'Stop-Process' })
        $stops.Count | Should -Be 1
        $stops[0].File | Should -Be "office_process.ps1"
    }
}

Describe "Office ファイルを安全に開くこと（docs/safety/checks.md「Office ファイルを開くときの設定」）" -Tag Meta {
    BeforeAll {
        $app = @($code | Where-Object { $_.File -eq "office_app.ps1" })
        $extract = @($code | Where-Object { $_.File -eq "extract_office.ps1" })
    }

    It "マクロを強制的に無効にしてから開く（AutomationSecurity = 3）" {
        (findPattern $app 'AutomationSecurity\s*=\s*3') | Should -Not -Be ""
    }

    It "イベントマクロを発火させない（EnableEvents = false）" {
        (findPattern $app 'EnableEvents\s*=\s*\$false') | Should -Not -Be ""
    }

    It "外部リンクを更新しない（AskToUpdateLinks = false・Open の UpdateLinks = 0）" {
        (findPattern $app 'AskToUpdateLinks\s*=\s*\$false') | Should -Not -Be ""
        # Excel の Workbooks.Open（第 2 引数 UpdateLinks = 0・第 3 引数 ReadOnly = $true）。3.2 の「原本は読み取り専用で開く」も兼ねる
        (findPattern $extract '\.Open\(\$openPath,\s*0,\s*\$true') | Should -Not -Be ""
    }

    It "インデクサの Office は画面に出さない（Visible = false）" {
        (findPattern $app 'Visible\s*=\s*\$false') | Should -Not -Be ""
        # 可視にするのは画面から元のファイルを開くときだけ（ui/open_source.ps1）
        $visible = @($code | Where-Object { $_.Text -match 'Visible\s*=\s*\$true' })
        (@($visible | Where-Object { $_.File -ne "open_source.ps1" } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should -Be ""
    }
}

Describe "細工したOfficeファイル（.docx・.pptx・.xlsx）から身を守ること（docs/safety/checks.md「Office ファイルを開くときの設定」）" -Tag Meta {
    It "XmlDocument への読み込みは newXmlDocument 関数だけで行う（LoadXml を直接呼ばない）" {
        # shared/ui/ は、このツール自身の XAML（利用者の Office ファイルではない）を読むので対象外
        (findPattern @($code | Where-Object { $_.Path -notlike "*\shared\ui\app_host.ps1" }) '\.LoadXml\(') | Should -Be ""
    }

    It "XmlDocument は newXmlDocument 経由でだけ作る（New-Object System.Xml.XmlDocument は office_reader.ps1 の中だけ）" {
        $news = @($code | Where-Object { $_.Text -match 'New-Object\s+System\.Xml\.XmlDocument' })
        (@($news | Where-Object { $_.File -ne "office_reader.ps1" -and $_.Path -notlike "*\shared\ui\app_host.ps1" } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should -Be ""
    }

    It "[xml] への型変換で XML を読み込まない（DTD 無効化を回避できてしまうため）" {
        (findPattern $code '\[xml\]') | Should -Be ""
    }

    It "XmlReaderSettings を作るところは、すべて DTD の処理を禁止する（DtdProcessing = Prohibit）" {
        # 作る数と、Prohibit を入れる数を比べる（ファイルに 1 か所あれば足りる、という見逃しを防ぐ）
        $creations = @($code | Where-Object { $_.Text -match 'New-Object\s+System\.Xml\.XmlReaderSettings' })
        $prohibits = @($code | Where-Object { $_.Text -match 'DtdProcessing\s*=\s*\[System\.Xml\.DtdProcessing\]::Prohibit' })
        ($creations.Count -gt 0) | Should -Be $true
        $prohibits.Count | Should -Be $creations.Count -Because "XmlReaderSettings を作った数だけ Prohibit が要る（作った場所: $(findPattern $creations '.')）"
    }

    It "ZIP の部品・1ファイルの合計のサイズに上限があり、読むのは readZipEntry だけ" {
        (findPattern $code 'zipPartMaxBytes') | Should -Not -Be ""
        (findPattern $code 'zipTotalMaxBytes') | Should -Not -Be ""
        $reads = @($code | Where-Object { $_.Text -match 'function readZipEntry' } | ForEach-Object { $_.File } | Sort-Object -Unique)
        ($reads -join ", ") | Should -Be "office_reader.ps1"
    }

    It "ZIP の部品の中身を実際に読む（entry.Open()・ReadToEnd）のは readZipEntry 関数と、上限を付けたシートの流れ読みだけ" {
        # 上限の判定（zipPartMaxBytes・zipTotalMaxBytes）を迂回して、office_reader.ps1 のほかの関数が
        # 直接 ZIP の中身を読んでしまわないことを、関数の行範囲で確かめる
        $officeReaderLines = @($code | Where-Object { $_.File -eq "office_reader.ps1" })
        $funcStarts = @($officeReaderLines | Where-Object { $_.Text -match '^\s*function\s+\w+' } | Sort-Object Line)
        # readZipEntry のほかに、シートのヘッダー・フッターを流れで読む readXlsxSheetHeaderFooter だけは、
        # 大きさの上限（zipSheetStreamMaxBytes。偽りの申告は MaxCharactersInDocument で打ち切る）を付けて entry.Open() してよい
        $ranges = @()
        foreach ($fn in @("readZipEntry", "readXlsxSheetHeaderFooter")) {
            $start = ($funcStarts | Where-Object { $_.Text -match "function\s+$fn\b" }).Line
            $start | Should -Not -BeNullOrEmpty
            $end = ($funcStarts | Where-Object { $_.Line -gt $start } | Sort-Object Line | Select-Object -First 1).Line
            if (-not $end) { $end = [int]::MaxValue }
            $ranges += , @($start, $end)
        }
        $streamed = @($officeReaderLines | Where-Object { $_.Line -ge $ranges[1][0] -and $_.Line -lt $ranges[1][1] })
        (findPattern $streamed 'MaxCharactersInDocument\s*=\s*\$script:zipSheetStreamMaxBytes') | Should -Not -Be ""

        # 対象は Office の読み取りと取り込みの全ファイル。ExtractToFile・CopyTo も同じく上限を迂回して中身を出せる
        $readerFiles = @($code | Where-Object { $_.Path -like "*\scripts\shared\office\*" -or $_.Path -like "*\scripts\tebunko\indexer\*" })
        # ZIP と関係のない既存の使い方（runspace を開く・状態ファイルを読む）は対象から外す
        $notZip = '^\s*(\$runspace|\$this\.Runspace)\.Open\(\)\s*$|^\s*\$(lines|text)\s*=\s*\$reader\.ReadToEnd\(\)'
        $targets = @($readerFiles | Where-Object { $_.Text -match '\.Open\(\)|ReadToEnd\(\)|ExtractToFile|ExtractToDirectory|Expand-Archive|\.CopyTo\(' -and $_.Text -notmatch $notZip })
        $outside = @($targets | Where-Object {
            $t = $_
            $t.File -ne "office_reader.ps1" -or @($ranges | Where-Object { $t.Line -ge $_[0] -and $t.Line -lt $_[1] }).Count -eq 0
        })
        (@($outside | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should -Be ""
        ($targets.Count -gt 0) | Should -Be $true  # 検査の取りこぼし（対象0件で通る）を防ぐ
    }

    It "readZipEntry の StringBuilder に容量を渡さない（申告の大きさぶんのメモリを先に確保しないため）" {
        $officeReaderLines = @($code | Where-Object { $_.File -eq "office_reader.ps1" })
        $funcStarts = @($officeReaderLines | Where-Object { $_.Text -match '^\s*function\s+\w+' } | Sort-Object Line)
        $start = ($funcStarts | Where-Object { $_.Text -match 'function\s+readZipEntry\b' }).Line
        $start | Should -Not -BeNullOrEmpty
        $end = ($funcStarts | Where-Object { $_.Line -gt $start } | Sort-Object Line | Select-Object -First 1).Line
        if (-not $end) { $end = [int]::MaxValue }
        $inside = @($officeReaderLines | Where-Object { $_.Line -ge $start -and $_.Line -lt $end })
        $builders = @($inside | Where-Object { $_.Text -match 'System\.Text\.StringBuilder' })
        ($builders.Count -gt 0) | Should -Be $true
        # 引数なし（行末が StringBuilder）であること。容量を渡す書き方（括弧・数字）は、どれも当たらない
        (@($builders | Where-Object { $_.Text -notmatch 'System\.Text\.StringBuilder\s*$' } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should -Be ""
    }

    It "XmlDocument を作る（New-Object System.Xml.XmlDocument）のは newXmlDocument 関数の中だけ" {
        # ファイル単位の確かめ（上の「newXmlDocument 経由でだけ作る」）に加え、office_reader.ps1 の
        # ほかの関数が newXmlDocument を経由せずに直接 XmlDocument を作っていないことを、関数の行範囲で確かめる
        $officeReaderLines = @($code | Where-Object { $_.File -eq "office_reader.ps1" })
        $funcStarts = @($officeReaderLines | Where-Object { $_.Text -match '^\s*function\s+\w+' } | Sort-Object Line)
        $start = ($funcStarts | Where-Object { $_.Text -match 'function\s+newXmlDocument\b' }).Line
        $start | Should -Not -BeNullOrEmpty
        $end = ($funcStarts | Where-Object { $_.Line -gt $start } | Sort-Object Line | Select-Object -First 1).Line
        if (-not $end) { $end = [int]::MaxValue }

        $targets = @($officeReaderLines | Where-Object { $_.Text -match 'New-Object\s+System\.Xml\.XmlDocument' })
        $outside = @($targets | Where-Object { $_.Line -lt $start -or $_.Line -ge $end })
        (@($outside | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should -Be ""
        ($targets.Count -gt 0) | Should -Be $true
    }
}

Describe "取り込み対象のファイルを書き換えないこと（docs/safety/file-access.md「取り込み対象のファイルは書き換えない」）" -Tag Meta {
    It "元のファイルのパスを書き込み・削除の API に渡さない" {
        # 書き込み・削除の呼び出し行に、取り込み対象（原本）を指す変数が現れないこと。
        # 原本は作業フォルダへコピーしてから開くため、書き込み先は常にコピー側（$tmpPath・$destPath・$copyPath 等）になる。
        # $targetFolder・$folder.Path = クロール対象フォルダ、$row.SourcePath = 検索結果の元のファイル
        $writes = @($code | Where-Object { $_.Text -match 'Remove-Item|WriteAllText|WriteAllLines|StreamWriter|\.SaveAs|\[System\.IO\.(File|Directory)\]::Move|Move-Item' })
        ($writes.Count -gt 0) | Should -Be $true
        (@($writes | Where-Object { $_.Text -match '\$targetFolder|\$row\.SourcePath|\$folder\.Path' } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should -Be ""
    }

    It "原本を読むのはコピーを作る処理（copyFileShared）だけ" {
        # copyFileShared は読み取りだけで開き、コピー先（作業フォルダ）へ書く
        (findPattern $code 'function copyFileShared') | Should -Not -Be ""
        $fs = @($code | Where-Object { $_.File -eq "fs.ps1" -and $_.Text -match '\$sourcePath' })
        (@($fs | Where-Object {
            $_.Text -match 'Remove-Item|WriteAllText|WriteAllLines|StreamWriter|Move-Item|\[System\.IO\.(File|Directory)\]::(Move|Delete|WriteAll)'
        } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should -Be ""
        (findPattern $fs '\[System\.IO\.FileAccess\]::Read') | Should -Not -Be ""
    }

    It "Office の SaveAs の保存先は作業フォルダのパスだけ" {
        $saves = @($code | Where-Object { $_.Text -match '\.SaveAs' })
        $saves.Count | Should -Be 3
        (@($saves | Where-Object { $_.Text -notmatch 'SaveAs2?\(\$(tmpPath|destPath)' } | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should -Be ""
    }

    It "原本は読み取り専用で開く（Excel・Word・PowerPoint）" {
        $extract = @($code | Where-Object { $_.File -eq "extract_office.ps1" })
        # Word: 第 3 引数 ReadOnly = $true / PowerPoint: 第 2 引数 ReadOnly = -1
        # Excel（Open の第 3 引数 ReadOnly = $true）は、2.2 の「外部リンクを更新しない」で UpdateLinks と一緒に確かめる
        (findPattern $extract '\$documents\.Open\(\$sourcePath,\s*\$false,\s*\$true') | Should -Not -Be ""
        (findPattern $extract '\$presentations\.Open\("\$\{sourcePath\}::dummy::",\s*-1') | Should -Not -Be ""
    }
}

Describe "取り込み対象の拡張子が固定であること（docs/safety/checks.md「検査項目と結果」）" -Tag Meta {
    BeforeAll {
        . "$here\helpers\load.ps1"
    }

    It "テキストの拡張子がちょうど 75（既定のアプリで開く 68＋メモ帳で開く 7）で、Office と合わせて 85 ちょうどである" {
        # 足したら必ずここが落ち、docs/safety/checks.md の見直し（開くと実行される種類を含まないか）を促す
        @(${textExtensions}).Count | Should -Be 75
        @(${textNotepadExtensions}).Count | Should -Be 7
        $expectedNotepad = @(".bat", ".cmd", ".ps1", ".vbs", ".js", ".reg", ".sh")
        (@(${textNotepadExtensions}) | Sort-Object) -join "," | Should -Be (($expectedNotepad | Sort-Object) -join ",")
        # メモ帳で開く拡張子はテキストの拡張子に含まれ、既定のアプリで開く拡張子とは重ならない
        (@(${textNotepadExtensions}) | Where-Object { ${textExtensions} -notcontains $_ }) | Should -BeNullOrEmpty
        (@(${textNotepadExtensions}) | Where-Object { ${textOpenExtensions} -contains $_ }) | Should -BeNullOrEmpty

        @(${officeExtensions} + ${textExtensions}).Count | Should -Be 85

        # 開くと実行・登録になる、または安全性の理由で対象外にした拡張子は、どちらの一覧にも含まれない
        $forbidden = @(
            ".exe", ".com", ".scr", ".hta", ".wsf", ".wsh", ".jse", ".vbe",
            ".lnk", ".url", ".msi", ".cpl", ".inf", ".ts"
        )
        foreach ($ext in $forbidden) {
            (${officeExtensions} -contains $ext) | Should -Be $false
            (${textExtensions} -contains $ext) | Should -Be $false
        }
    }

    It "実行されうるテキストの拡張子は既定のアプリで開かない" {
        foreach ($ext in @(${textNotepadExtensions})) {
            (testTextOpenWithNotepad "a${ext}") | Should -Be $true -Because "拡張子 ${ext} はメモ帳で開く一覧のはず"
        }
        (@(${textNotepadExtensions}) | Where-Object { ${textOpenExtensions} -contains $_ }) | Should -BeNullOrEmpty
    }
}

Describe "書き込み先が限られていること（docs/safety/file-access.md「書き込み・削除する場所」）" -Tag Meta {
    It "書き込みに使うフォルダの定義は、設定ファイルの置き場所（ツールのフォルダか既定のワークスペース）と前の版の片付け先（TEMP）だけ" {
        $paths = @($code | Where-Object { $_.File -in @("paths.ps1", "workspace.ps1", "data_dir.ps1", "settings.ps1") })
        # 設定ファイルの置き場所は、ツールのフォルダか、書き込めないときの既定のワークスペースだけ
        (findPattern $paths 'return Join-Path \(getDataDir \$root \$defaultWorkDir\) "setting\.config"') | Should -Not -Be ""
        (findPattern $paths '\$\{settingsFile\}\s*=\s*getSettingsFilePath\s+\$\{rootDir\}\s+\(getDefaultWorkDir\)') | Should -Not -Be ""
        (findPattern $paths 'return \$fallbackDir') | Should -Not -Be ""
        # ワークスペースは設定から決め、中の場所はワークスペースのフォルダから組み立てる
        (findPattern $paths '\$\{script:workspace\}\s*=\s*\[Workspace\]::new\(\(getWorkDir\)\)') | Should -Not -Be ""
        (findPattern $paths '\$this\.IndexDir\s*=\s*"\$dir\\content_index"') | Should -Not -Be ""
        (findPattern $paths '\$\{legacyTmpParent\}\s*=\s*Join-Path\s*\(\[System\.IO\.Path\]::GetTempPath\(\)\)\s*"tebunko"') | Should -Not -Be ""
        (findPattern $paths '\$this\.TmpRoot\s*=\s*"\$dir\\tmp"') | Should -Not -Be ""
        (findPattern $paths '\$this\.OfficePidRoot\s*=\s*"\$dir\\office_pids"') | Should -Not -Be ""
        (findPattern $paths '\$this\.PublishDir\s*=\s*"\$dir\\') | Should -Not -Be ""
    }

    It "起動に失敗したときの記録の置き場所は、ツールのフォルダだけ（startup_error.ps1・tebunko.bat）" {
        $guiCode = @($code | Where-Object { $_.File -eq "startup_error.ps1" })
        (findPattern $guiCode 'Join-Path\s+\$toolDir\s+"startup_error\.txt"') | Should -Not -Be ""
        $bat = [System.IO.File]::ReadAllText($launcher)
        ($bat -match "Join-Path\s+\`$root\s+'startup_error\.txt'") | Should -Be $true
    }

    It "%LOCALAPPDATA%・%TEMP% に書かない（場所を環境から得る書き方が、許すものの一覧に収まっている。前の版の片付け先の GetTempPath は別の It で確かめる）" {
        $bat = [System.IO.File]::ReadAllText($launcher)
        (getWritePlaceViolations $code $bat) -join ", " | Should -Be ""
    }

    It "書き込み先の検査は、違反を 1 つ入れると期待する決まりで落ちる: <Case>" -TestCases @(
        # Rules = 落ちる決まりの名前（getWritePlaceViolations が返す文字列の先頭）。ちょうどこの集まりで落ちること
        @{ Case = '$env:LOCALAPPDATA'; File = "x.ps1"; Text = '$d = Join-Path $env:LOCALAPPDATA "tebunko"'; Bat = ""; Rules = @("環境変数", "禁止の語") }
        @{ Case = '$Env:LocalAppData（大文字小文字違い）'; File = "x.ps1"; Text = '$d = $Env:LocalAppData'; Bat = ""; Rules = @("環境変数", "禁止の語") }
        @{ Case = '${env:TEMP}'; File = "x.ps1"; Text = '$d = "${env:TEMP}\a"'; Bat = ""; Rules = @("環境変数", "禁止の語") }
        @{ Case = 'env:TEMP'; File = "x.ps1"; Text = '$t = $env:TEMP'; Bat = ""; Rules = @("環境変数", "禁止の語") }
        @{ Case = '環境変数 env:USERPROFILE'; File = "x.ps1"; Text = '$t = $env:USERPROFILE'; Bat = ""; Rules = @("環境変数") }
        @{ Case = 'TEBUNKO_DEFAULT_WORKSPACE を別のファイルで読む'; File = "x.ps1"; Text = '$d = $env:TEBUNKO_DEFAULT_WORKSPACE'; Bat = ""; Rules = @("環境変数") }
        @{ Case = '禁止の語だけ: パスの AppData'; File = "x.ps1"; Text = '$d = "$base\AppData\Local\tebunko"'; Bat = ""; Rules = @("禁止の語") }
        @{ Case = '禁止の語だけ: コードの中の %TEMP%'; File = "x.ps1"; Text = 'cmd /c "echo %TEMP%"'; Bat = ""; Rules = @("禁止の語") }
        @{ Case = 'LocalApplicationData'; File = "x.ps1"; Text = '$d = [System.Environment]::GetFolderPath("LocalApplicationData")'; Bat = ""; Rules = @("GetFolderPath") }
        @{ Case = 'GetFolderPath を別の所で使う'; File = "x.ps1"; Text = '$x = [System.Environment]::GetFolderPath("UserProfile")'; Bat = ""; Rules = @("GetFolderPath") }
        @{ Case = 'GetFolderPath("ApplicationData")'; File = "settings.ps1"; Text = '[string]$profileDir = [System.Environment]::GetFolderPath("ApplicationData")'; Bat = ""; Rules = @("GetFolderPath") }
        @{ Case = 'ExpandEnvironmentVariables("%LOCALAPPDATA%\tebunko")'; File = "settings.ps1"; Text = '$x = [System.Environment]::ExpandEnvironmentVariables("%LOCALAPPDATA%\tebunko")'; Bat = ""; Rules = @("ExpandEnvironmentVariables", "禁止の語") }
        @{ Case = 'ExpandEnvironmentVariables を別の所で使う'; File = "x.ps1"; Text = '$x = [System.Environment]::ExpandEnvironmentVariables($y)'; Bat = ""; Rules = @("ExpandEnvironmentVariables") }
        @{ Case = 'GetEnvironmentVariable'; File = "x.ps1"; Text = '$x = [System.Environment]::GetEnvironmentVariable("TEMP")'; Bat = ""; Rules = @("GetEnvironmentVariable") }
        @{ Case = 'bat の %TEMP%'; File = ""; Text = ""; Bat = 'set "OUT=%TEMP%\a.txt"'; Rules = @("bat の変数", "禁止の語") }
        @{ Case = 'bat の %LOCALAPPDATA%'; File = ""; Text = ""; Bat = 'set "OUT=%LOCALAPPDATA%\a.txt"'; Rules = @("bat の変数", "禁止の語") }
        @{ Case = 'bat の %USERPROFILE%'; File = ""; Text = ""; Bat = 'set "OUT=%USERPROFILE%\a.txt"'; Rules = @("bat の変数") }
        @{ Case = 'bat の $env:TEMP'; File = ""; Text = ""; Bat = 'set "PSCMD=%PSCMD%$p = $env:TEMP"'; Rules = @("bat の環境変数", "禁止の語") }
        @{ Case = 'bat の GetTempPath'; File = ""; Text = ""; Bat = 'set "PSCMD=%PSCMD%$p = [System.IO.Path]::GetTempPath()"'; Rules = @("bat の API") }
        @{ Case = 'bat の GetFolderPath'; File = ""; Text = ""; Bat = 'set "PSCMD=%PSCMD%$p = [Environment]::GetFolderPath(''LocalApplicationData'')"'; Rules = @("bat の API") }
        @{ Case = 'bat の New-TemporaryFile'; File = ""; Text = ""; Bat = 'set "PSCMD=%PSCMD%$p = New-TemporaryFile"'; Rules = @("bat の API") }
        @{ Case = 'bat の GetEnvironmentVariable'; File = ""; Text = ""; Bat = 'set "PSCMD=%PSCMD%$p = [Environment]::GetEnvironmentVariable(''TEMP'')"'; Rules = @("bat の API") }
        @{ Case = 'bat の ExpandEnvironmentVariables'; File = ""; Text = ""; Bat = 'set "PSCMD=%PSCMD%$p = [Environment]::ExpandEnvironmentVariables(''x'')"'; Rules = @("bat の API") }
    ) {
        $lines = @()
        if ($Text -ne "") {
            $lines = @([pscustomobject]@{ File = $File; Line = 1; Text = $Text })
        }
        $found = @(getWritePlaceViolations $lines $Bat | ForEach-Object { ($_ -split ':')[0] } | Sort-Object -Unique)
        ($found -join ",") | Should -Be ((@($Rules) | Sort-Object) -join ",")
    }

    It "書き込み先の検査は、許した ExpandEnvironmentVariables の行が 2 回あると落ちる" {
        $line = '        $path = [System.Environment]::ExpandEnvironmentVariables($path).Trim()'
        $lines = @(
            [pscustomobject]@{ File = "folder.ps1"; Line = 1; Text = $line }
            [pscustomobject]@{ File = "folder.ps1"; Line = 2; Text = $line }
        )
        $found = @(getWritePlaceViolations $lines "" | ForEach-Object { ($_ -split ':')[0] } | Sort-Object -Unique)
        ($found -join ",") | Should -Be "ExpandEnvironmentVariables"
    }

    It "書き込み先の検査は、許すものだけなら通る" {
        $lines = @(
            [pscustomobject]@{ File = "settings.ps1"; Line = 1; Text = '        [string]$profileDir = [System.Environment]::GetFolderPath("UserProfile")' }
            [pscustomobject]@{ File = "folder.ps1"; Line = 2; Text = '        $path = [System.Environment]::ExpandEnvironmentVariables($path).Trim()' }
            [pscustomobject]@{ File = "settings.ps1"; Line = 3; Text = '    $folder = [System.Environment]::ExpandEnvironmentVariables($folder)' }
            [pscustomobject]@{ File = "shell.ps1"; Line = 4; Text = 'Start-Process -FilePath "$env:SystemRoot\System32\notepad.exe"' }
            [pscustomobject]@{ File = "x.ps1"; Line = 5; Text = '$x = $env:TEBUNKO_GUI_LEFTOVER_FILE' }
            [pscustomobject]@{ File = "settings.ps1"; Line = 6; Text = '    $override = $env:TEBUNKO_DEFAULT_WORKSPACE' }
        )
        @(getWritePlaceViolations $lines 'set "PS1=%SystemRoot%\System32\a.exe"
set "PSCMD=%PSCMD%x"').Count | Should -Be 0
    }

    It "書き込み先の検査は、TEBUNKO_DEFAULT_WORKSPACE を settings.ps1 が 2 回読むと落ちる" {
        $lines = @(
            [pscustomobject]@{ File = "settings.ps1"; Line = 1; Text = '$a = $env:TEBUNKO_DEFAULT_WORKSPACE' }
            [pscustomobject]@{ File = "settings.ps1"; Line = 2; Text = '$b = $env:TEBUNKO_DEFAULT_WORKSPACE' }
        )
        $found = @(getWritePlaceViolations $lines "" | ForEach-Object { ($_ -split ':')[0] } | Sort-Object -Unique)
        ($found -join ",") | Should -Be "環境変数"
    }

    It "ドライブ直下・システムフォルダを直接指す書き込み先が無い" {
        (findPattern $code '"[A-Za-z]:\\(Windows|Program Files|Users)') | Should -Be ""
        (findPattern $code 'GetFolderPath\("?(System|Windows|ProgramFiles|Startup)') | Should -Be ""
    }

    It "content_index を Windows Search の対象から外す属性は、fs.ps1 の setNotContentIndexed だけが書き、呼ぶのは決まった 6 ファイルだけ" {
        # NotContentIndexed 属性を直接書くのは setNotContentIndexed（fs.ps1）だけ。
        # 読むだけの testNotContentIndexed（fs.ps1。親フォルダの属性を見て引き継ぐか決める）は除く
        $writes = @($code | Where-Object { $_.Text -match "(?<!set)(?<!test)NotContentIndexed" -and $_.File -ne "fs.ps1" })
        (@($writes | ForEach-Object { "$($_.File):$($_.Line)" }) -join ", ") | Should -Be ""
        (findPattern $code 'function setNotContentIndexed') | Should -Not -Be ""
        (findPattern $code 'function testNotContentIndexed') | Should -Not -Be ""

        # testNotContentIndexed を呼ぶのは newWorkerTmpDir（index_migrate.ps1）だけ
        $testCallers = @($code | Where-Object { $_.Text -match "testNotContentIndexed" -and $_.Text -notmatch "function testNotContentIndexed" })
        (@($testCallers | ForEach-Object { $_.File } | Sort-Object -Unique) -join ", ") | Should -Be "index_migrate.ps1"

        # 呼ぶのは全体に付ける indexer_run.ps1・取り込みの作業フォルダに付ける index_migrate.ps1・
        # 入れた直後に付ける index_store.ps1・pack_store.ps1・source_map.ps1・インポートで入れたフォルダ全体に付ける index_archive.ps1 の 6 ファイルだけ
        $callers = @($code | Where-Object { $_.Text -match "setNotContentIndexed" -and $_.Text -notmatch "function setNotContentIndexed" })
        $callerFiles = @($callers | ForEach-Object { $_.File } | Sort-Object -Unique)
        $callerFiles.Count | Should -Be 6
        ($callerFiles -contains "indexer_run.ps1") | Should -Be $true
        ($callerFiles -contains "index_migrate.ps1") | Should -Be $true
        ($callerFiles -contains "index_store.ps1") | Should -Be $true
        ($callerFiles -contains "pack_store.ps1") | Should -Be $true
        ($callerFiles -contains "source_map.ps1") | Should -Be $true
        ($callerFiles -contains "index_archive.ps1") | Should -Be $true

        # 渡す引数まで確かめる（indexer_run.ps1 は $workspace.IndexDir、index_migrate.ps1 は取り込みの作業フォルダ（$dir）、
        # index_store.ps1 は $bookDir、pack_store.ps1 は $folder、index_archive.ps1 は $targetIndexDir）。
        # -Recurse で下をたどるのは indexer_run.ps1（content_index 全体）と index_archive.ps1（インポートで入れた 1 つのインデックス。Directory.Move で入るため）だけ（ほかは 1 か所ずつなので要らない）
        (findPattern $callers 'setNotContentIndexed\s+\$workspace\.IndexDir\s+-Recurse') | Should -Not -Be ""
        (findPattern $callers 'setNotContentIndexed\s+\$dir\)') | Should -Not -Be ""
        (findPattern $callers 'setNotContentIndexed\s+\$bookDir\)') | Should -Not -Be ""
        (findPattern $callers 'setNotContentIndexed\s+\$folder\)') | Should -Not -Be ""
        (findPattern $callers 'setNotContentIndexed\s+\$targetIndexDir\s+-Recurse') | Should -Not -Be ""
        $recurseCallers = @($callers | Where-Object { $_.Text -match "-Recurse" } | ForEach-Object { $_.File } | Sort-Object -Unique)
        ($recurseCallers -join ", ") | Should -Be "index_archive.ps1, indexer_run.ps1"
    }

    It "%TEMP% を指す書き方は決めた所だけ（取り込みの作業フォルダはワークスペースの tmp\\の下を使う）" {
        # %TEMP% を直接指すのは、前の版の片付けだけに使う場所（${legacyTmpParent}。paths.ps1）だけ
        $tempRefs = @($code | Where-Object { $_.Text -match 'GetTempPath|env:TEMP\b|env:TMP\b|New-TemporaryFile|GetTempFileName' })
        $tempFiles = @($tempRefs | ForEach-Object { $_.File } | Sort-Object -Unique)
        ($tempFiles -join ", ") | Should -Be "paths.ps1"
        (@($tempRefs | Where-Object { $_.File -eq "paths.ps1" }).Count) | Should -Be 1
    }

    It "異常終了で残った作業フォルダを次回起動時に回収する" {
        # %TEMP%\tebunko\<PID>（前の版が残した作業フォルダ）に原本のコピーが残り続けないこと（docs/safety/disclosure.md「原本の一時コピーと、その回収」）
        (findPattern $code 'function removeStaleTmpDirs') | Should -Not -Be ""
        # インデックス作成の始め（invokeIndexer の本体）で呼ぶ
        (findPattern $code '^\s+removeStaleTmpDirs$') | Should -Not -Be ""
    }

    It "エクスポートは利用者が選んだ保存先と、その .tmp だけに書き込む" {
        # ［結果をファイルに出力］（利用者が指定した出力先）と同じ扱い
        $archive = @($code | Where-Object { $_.File -eq "index_archive.ps1" })
        (findPattern $archive '\$tmpPath\s*=\s*"\$\{destPath\}\.tmp"') | Should -Not -Be ""
        (findPattern $archive '\[System\.IO\.Compression\.ZipFile\]::Open\(\$tmpPath') | Should -Not -Be ""
        (findPattern $archive '\[System\.IO\.File\]::Move\(\(toLongPath \$tmpPath\), \(toLongPath \$destPath\)\)') | Should -Not -Be ""
    }
}

Describe "サードパーティの静的解析（docs/safety/scans.md「静的解析: PSScriptAnalyzer（Microsoft）」）" -Tag Meta {
    # PSScriptAnalyzer（Microsoft 提供）で検査する。未導入の環境では飛ばす:
    #   Install-Module PSScriptAnalyzer -Scope CurrentUser
    # -Skip は探索のときに決まるため、導入の有無は Describe の本体で調べる
    $hasAnalyzer = (@(Get-Module -ListAvailable PSScriptAnalyzer).Count -gt 0)

    BeforeAll {
        if (@(Get-Module -ListAvailable PSScriptAnalyzer).Count -gt 0) {
            Import-Module PSScriptAnalyzer -ErrorAction SilentlyContinue
        }
        $securitySettings = "$PSScriptRoot\PSScriptAnalyzer.security.psd1"

        # 継承元の型が別ファイルにある場合の TypeNotFound は除く（読み込む順で解決する。
        # tests\meta\structure.Tests.ps1 の「スクリプトの構文」と同じ扱い）
        function formatFindings {
            param ($found)
            return (@($found | Where-Object { $_.RuleName -ne "TypeNotFound" } | ForEach-Object { "$($_.RuleName) $($_.ScriptName):$($_.Line)" }) -join ", ")
        }
    }

    It "安全性にかかわるルールの指摘が 0 件" -Skip:(-not $hasAnalyzer) {
        (formatFindings (Invoke-ScriptAnalyzer -Path $scriptsDir -Recurse -Settings $securitySettings)) | Should -Be ""
    }

    It "Error 重大度の指摘が 0 件（全ルール）" -Skip:(-not $hasAnalyzer) {
        (formatFindings (Invoke-ScriptAnalyzer -Path $scriptsDir -Recurse -Severity Error)) | Should -Be ""
    }

    It "制限言語モード（Constrained Language Mode）で使えない書き方が無い" -Skip:(-not $hasAnalyzer) {
        # AppLocker・WDAC で制限言語モードを強制している環境向け（docs/safety/scans.md「実行環境の制約との適合」）
        (formatFindings (Invoke-ScriptAnalyzer -Path $scriptsDir -Recurse -IncludeRule PSUseConstrainedLanguageMode)) | Should -Be ""
    }

    It "安全性のルール設定に、検査すべきルールが含まれている" {
        # 設定ファイルからルールを消して指摘 0 件にする、という抜け道を防ぐ
        $settings = Import-LocalizedData -BaseDirectory $PSScriptRoot -FileName "PSScriptAnalyzer.security.psd1"
        @($settings.IncludeRules).Count | Should -Be 14
        ($settings.IncludeRules -contains "PSAvoidUsingInvokeExpression") | Should -Be $true
        ($settings.IncludeRules -contains "PSAvoidUsingPlainTextForPassword") | Should -Be $true
        ($settings.IncludeRules -contains "PSAvoidUsingConvertToSecureStringWithPlainText") | Should -Be $true
    }
}

Describe "第三者が検証するための資料がそろっていること（docs/safety/scans.md「配布物の完全性（カタログ・ハッシュ一覧・来歴の署名）」・docs/safety/supply-chain.md「供給網（サプライチェーン）とライセンス」）" -Tag Meta {
    It "<name>" -TestCases @(
        @{ name = "安全性説明書がある"; file = "docs\safety\index.md" }
        @{ name = "脆弱性の連絡先（.github\SECURITY.md）がある"; file = ".github\SECURITY.md" }
        @{ name = "配布物の完全性を確かめる手順（tools\new_release_files.ps1）がある"; file = "tools\new_release_files.ps1" }
    ) {
        param ($name, $file)
        (Test-Path -LiteralPath "$rootDir\$file") | Should -Be $true
    }

    It "ライセンス（LICENSE）があり、MIT の条文と著作権表示を含む" {
        $licensePath = "$rootDir\LICENSE"
        (Test-Path -LiteralPath $licensePath) | Should -Be $true
        $license = [System.IO.File]::ReadAllText($licensePath)
        # 表題・著作権表示・許諾条項・無保証条項がそろっていること（MIT の全文）
        ($license -like "MIT License*") | Should -Be $true
        ($license -match "Copyright \(c\) \d{4} \S+") | Should -Be $true
        ($license -like "*Permission is hereby granted, free of charge*") | Should -Be $true
        ($license -like "*WITHOUT WARRANTY OF ANY KIND*") | Should -Be $true
        # 実名を含めない（AGENTS.md「個人情報を書かない」）。著作権者はアカウント名で表記する
        ($license -like "*hsgwa*") | Should -Be $true
    }

    It "部品表の雛形（sbom.cdx.json）があり、本体の説明・ライセンス・前提ソフトウェアを持ち、部品は持たない" {
        $sbomPath = "$rootDir\sbom.cdx.json"
        (Test-Path -LiteralPath $sbomPath) | Should -Be $true
        $sbom = [System.IO.File]::ReadAllText($sbomPath) | ConvertFrom-Json
        $sbom.bomFormat | Should -Be "CycloneDX"
        $sbom.specVersion | Should -Be "1.6"
        # 本ツール自身のライセンスは SPDX の識別子で記載する（LICENSE と一致させる）
        $sbom.metadata.component.licenses[0].license.id | Should -Be "MIT"
        # ファイルごとの一覧とハッシュは、配布物を作るときに tools\new_sbom.ps1 が zip の中身から作る。雛形には書かない
        $sbom.PSObject.Properties.Name -contains "components" | Should -Be $false
        $sbom.PSObject.Properties.Name -contains "dependencies" | Should -Be $false
        # 前提ソフトウェア（同梱しないもの）は metadata.properties に記載する
        @($sbom.metadata.properties | Where-Object { $_.name -eq "tebunko:prerequisite" }).Count -gt 0 | Should -Be $true
    }
}
