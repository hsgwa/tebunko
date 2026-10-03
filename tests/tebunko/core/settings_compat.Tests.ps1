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
        # 値を、オブジェクトのプロパティの並びに左右されない文字列にする（配列は並び順のまま）
        param ($value)
        if ($null -eq $value) { return "null" }
        if (($value -is [System.Collections.IEnumerable]) -and -not ($value -is [string])) {
            return "[" + ((@($value) | ForEach-Object { toCanonicalText $_ }) -join ",") + "]"
        }
        if ($value -is [System.Management.Automation.PSCustomObject]) {
            $names = @($value.PSObject.Properties.Name | Sort-Object)
            return "{" + (($names | ForEach-Object { "$_=$(toCanonicalText $value.$_)" }) -join ";") + "}"
        }
        return [string]$value
    }

    function script:getJsonPropertyPaths {
        # JSON テキストから、プロパティのパス（配列は [添字] を付ける）をすべて返す（道具の違いを気にせず比べるため）
        param ([string]$json)
        $paths = New-Object System.Collections.Generic.List[string]
        function script:walkJsonNode {
            param ([string]$prefix, $node)
            if ($node -is [System.Management.Automation.PSCustomObject]) {
                foreach ($prop in $node.PSObject.Properties) {
                    $path = if ($prefix) { "$prefix.$($prop.Name)" } else { $prop.Name }
                    $paths.Add($path)
                    walkJsonNode $path $prop.Value
                }
            } elseif ($node -is [array]) {
                for ($i = 0; $i -lt $node.Count; $i++) {
                    walkJsonNode "$prefix[$i]" $node[$i]
                }
            }
        }
        walkJsonNode "" (ConvertFrom-Json $json)
        return @($paths)
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
            if ($name -eq "openMode") { continue }
            (toCanonicalText $actual[$name]) | Should -Be (toCanonicalText $sample.Expected.$name) -Because "キー ${name}"
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
