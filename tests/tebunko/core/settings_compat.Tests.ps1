# 前の版が作った setting.config（見本・golden。tests/testdata/compat/settings/<見本の名前>/）が、
# 今のコードでも壊れたと取り違えずに読め、同じ値になること・書き足してもほかの値を保つことを確かめる（互換テスト）。
# 見本の置き方・作り方は tests/testdata/README.md「前の版のファイル（compat\）」。
BeforeDiscovery {
    $compatSettingsRoot = Join-Path (Resolve-Path "$PSScriptRoot\..\..\testdata").Path "compat\settings"
    $settingsSamples = @(Get-ChildItem -LiteralPath $compatSettingsRoot -Directory | ForEach-Object { @{ Sample = $_.Name } })
}

BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "$PSScriptRoot\..\..\helpers\settings_golden.ps1"

    $script:compatSettingsRoot = Join-Path (Resolve-Path "$PSScriptRoot\..\..\testdata").Path "compat\settings"

    function script:getSettingsSample {
        # 見本の setting.config と、手で書いた expected.json を返す
        param ([string]$name)
        $dir = Join-Path $compatSettingsRoot $name
        $expected = [System.IO.File]::ReadAllText((Join-Path $dir "expected.json"), [System.Text.Encoding]::UTF8) | ConvertFrom-Json
        return @{ Name = $name; SettingsPath = Join-Path $dir "setting.config"; Expected = $expected }
    }

    function script:copySettingsSample {
        # 見本の setting.config を $TestDrive にコピーする（元の見本を書き換えない）
        param ($sample, [string]$label)
        $path = Join-Path $TestDrive "$($sample.Name)_$label\setting.config"
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $path) | Out-Null
        Copy-Item -LiteralPath $sample.SettingsPath -Destination $path -Force
        return $path
    }

    function script:toCanonicalText {
        # 値を、型名を添えた文字列にする（オブジェクトのプロパティの並びに左右されない・配列は並び順のまま・
        # 型の変化（数値の 2 と文字列の "2" など）も見分ける）。Hashtable（IDictionary）はキーで並べ替えて比べる
        param ($value)
        if ($null -eq $value) { return "null" }
        if ($value -is [string]) { return "string:$value" }
        if ($value -is [System.Collections.IDictionary]) {
            $names = @($value.Keys | Sort-Object)
            return "{" + (($names | ForEach-Object { "$_=$(toCanonicalText $value[$_])" }) -join ";") + "}"
        }
        if ($value -is [System.Management.Automation.PSCustomObject]) {
            $names = @($value.PSObject.Properties.Name | Sort-Object)
            return "{" + (($names | ForEach-Object { "$_=$(toCanonicalText $value.$_)" }) -join ";") + "}"
        }
        if ($value -is [System.Collections.IEnumerable]) {
            return "[" + ((@($value) | ForEach-Object { toCanonicalText $_ }) -join ",") + "]"
        }
        return "$($value.GetType().Name):$value"
    }

    function script:walkJsonNode {
        # getJsonPropertyPaths から呼ぶ再帰（配列の添字は付けない。targetFolders[].name のように、要素の形だけを見る）
        param ([System.Collections.Generic.List[string]]$paths, [string]$prefix, $node)
        if ($node -is [System.Management.Automation.PSCustomObject]) {
            foreach ($prop in $node.PSObject.Properties) {
                $path = if ($prefix) { "$prefix.$($prop.Name)" } else { $prop.Name }
                $paths.Add($path)
                walkJsonNode $paths $path $prop.Value
            }
        } elseif ($node -is [array]) {
            foreach ($item in $node) {
                walkJsonNode $paths "$prefix[]" $item
            }
        }
    }

    function script:getJsonPropertyPaths {
        # JSON テキストから、プロパティのパス（配列は要素ごとに分けず [] を付ける）をすべて返す（道具の違いを気にせず比べるため）
        param ([string]$json)
        $paths = New-Object System.Collections.Generic.List[string]
        walkJsonNode $paths "" (ConvertFrom-Json $json)
        return @($paths | Sort-Object -Unique)
    }

    # settings_compat e が名前で呼ぶ関数の表。IsArray は、戻り値が一覧（配列）かどうか
    # （PowerShell は要素が 1 つの配列をそのまま返すと中身だけに崩れるため、@() で包み直す必要があるかの目印）
    $script:settingsFunctionsIsArray = @{
        getTargetFolders   = $true
        readIndexSources   = $true
        readSearchExcludes = $true
        readFileKinds      = $true
        readSearchOption   = $false
        readOpenMode       = $false
        readCloudFiles     = $false
        getWorkDir         = $false
    }
}

Describe "見本をそのまま読む（settings_compat a）" -Tag Io {
    It "<Sample>: repairBrokenSettings が壊れていると判定せず、ファイルをそのまま残す" -TestCases $settingsSamples {
        $sample = getSettingsSample $Sample
        $path = copySettingsSample $sample "a"
        $before = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)

        (repairBrokenSettings $path) | Should -Be ""

        Test-Path -LiteralPath $path | Should -Be $true
        [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8) | Should -Be $before
    }
}

Describe "今のコードでも同じ値として読める（settings_compat b）" -Tag Io {
    It "<Sample>: readSettings が expected.json にあるすべての項目と同じ値を返す" -TestCases $settingsSamples {
        $sample = getSettingsSample $Sample
        $path = copySettingsSample $sample "b"
        $actual = readSettings $path

        foreach ($name in @($sample.Expected.PSObject.Properties.Name)) {
            # functions は settings_compat e で確かめる（readSettings の生の値ではなく、関数の戻り値のため）
            if ($name -eq "functions") { continue }
            # openMode だけ readOpenMode（知らない値なら「通常」に戻す検査も含む）で読む。ほかはそのまま
            if ($name -eq "openMode") {
                (toCanonicalText (readOpenMode $path)) | Should -Be (toCanonicalText $sample.Expected.$name) -Because "キー ${name}"
            } else {
                (toCanonicalText $actual[$name]) | Should -Be (toCanonicalText $sample.Expected.$name) -Because "キー ${name}"
            }
        }
    }
}

Describe "今の版が書き足しても、ほかの値を保つ（settings_compat c）" -Tag Io {
    It "<Sample>: writeOpenMode で開き方を変えても、その値だけが変わり、ほかの値は見本のまま" -TestCases $settingsSamples {
        $sample = getSettingsSample $Sample
        $path = copySettingsSample $sample "c"

        writeOpenMode "new" $path

        readOpenMode $path | Should -Be "new"

        $actual = readSettings $path
        foreach ($name in @($sample.Expected.PSObject.Properties.Name)) {
            if ($name -eq "openMode" -or $name -eq "functions") { continue }
            (toCanonicalText $actual[$name]) | Should -Be (toCanonicalText $sample.Expected.$name) -Because "キー ${name}"
        }
    }
}

Describe "今のコードの関数が返す値も変わっていない（settings_compat e）" -Tag Io {
    # readSettings は一覧の値を生のオブジェクトのまま返すため、settings_compat b・c だけでは、
    # 一覧を整える関数（例: getTargetFolders が enabled=false のフォルダを取り込みの対象に戻す）の変化に気付けない。
    # expected.json の functions（関数名 → 戻り値）を表として使い、名前で関数を呼んで確かめる
    It "<Sample>: expected.json の functions にある関数を名前で呼び、同じ戻り値を返す" -TestCases $settingsSamples {
        $sample = getSettingsSample $Sample
        $path = copySettingsSample $sample "e"
        $sample.Expected.functions | Should -Not -Be $null -Because "見本の expected.json に functions がありません"

        foreach ($functionName in @($sample.Expected.functions.PSObject.Properties.Name)) {
            # 代入の式に if をそのまま使うと、要素が 1 つの配列でも中身だけに崩れて戻るため、文として分ける
            if ($script:settingsFunctionsIsArray[$functionName]) {
                $actual = @(& $functionName $path)
            } else {
                $actual = & $functionName $path
            }
            (toCanonicalText $actual) | Should -Be (toCanonicalText $sample.Expected.functions.$functionName) -Because "関数 ${functionName}"
        }
    }
}

Describe "今の版の書く形が見本にある（settings_compat d）" -Tag Io {
    It "今の版の writeGoldenSettings が書く JSON のプロパティのパスが、どれか 1 つの見本の setting.config に含まれる" {
        $todayPath = Join-Path $TestDrive "today\setting.config"
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $todayPath) | Out-Null
        writeGoldenSettings -path $todayPath
        $todayPaths = @(getJsonPropertyPaths ([System.IO.File]::ReadAllText($todayPath, [System.Text.Encoding]::UTF8)))
        $todayPaths.Count | Should -BeGreaterThan 0

        $sampleDirs = @(Get-ChildItem -LiteralPath $compatSettingsRoot -Directory)
        $sampleDirs.Count | Should -BeGreaterThan 0 -Because "見本が 1 つも無い"

        $matched = @($sampleDirs | Where-Object {
            $samplePaths = New-Object "System.Collections.Generic.HashSet[string]"
            foreach ($path in (getJsonPropertyPaths ([System.IO.File]::ReadAllText((Join-Path $_.FullName "setting.config"), [System.Text.Encoding]::UTF8)))) {
                [void]$samplePaths.Add($path)
            }
            @($todayPaths | Where-Object { -not $samplePaths.Contains($_) }).Count -eq 0
        })
        $matched.Count | Should -BeGreaterThan 0 -Because "今のコードが書くプロパティのパスを、すべて含む見本がありません"
    }
}
