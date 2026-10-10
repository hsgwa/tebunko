# 前の版のファイル（見本・golden。tests/testdata/compat/index/<見本の名前>/）が、
# 見本として要る部分をそろえていること・今のコードの形式の目印を含んでいることを確かめる。
# 中身が今のコードで読めること自体は tests/tebunko/indexer/index_compat.Tests.ps1 で確かめる（ここでは読まない）。
# zip の中の content_index/ という決め打ちの文字列は、index_compat.Tests.ps1 の check i が確かめる（ここでは見ない）
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\..\helpers\compat.ps1"

    ${compatIndexRoot} = "${testDataDir}\compat\index"
    ${requiredEntries} = @("source", "ws", "export.zip", "file_times.tsv", "expected.json")
    ${compatSettingsRoot} = "${testDataDir}\compat\settings"
}

Describe "前の版のファイル（compat\index）" -Tag Meta {
    It "それぞれの見本に、要る 5 つがそろっている（source・ws・export.zip・file_times.tsv・expected.json）" {
        $problems = New-Object System.Collections.Generic.List[string]
        foreach ($dir in @(Get-ChildItem -LiteralPath ${compatIndexRoot} -Directory)) {
            foreach ($entry in ${requiredEntries}) {
                if (!(Test-Path -LiteralPath "$($dir.FullName)\$entry")) {
                    $problems.Add("$($dir.Name) に $entry がありません")
                }
            }
        }
        ($problems -join "`n") | Should -Be ""
    }

    It "今のコードの形式の目印が、1 つの見本にすべて残っている（見本ごとに調べ、どれか 1 つがそろっていればよい）" {
        $samples = @(Get-ChildItem -LiteralPath ${compatIndexRoot} -Directory)
        $samples.Count | Should -BeGreaterThan 0 -Because "見本が 1 つも無い"
        $problemsBySample = [ordered]@{}

        foreach ($sample in $samples) {
            $problems = New-Object System.Collections.Generic.List[string]
            $wsDir = "$($sample.FullName)\ws"

            # contentIndexVersion: 本文インデックスのファイル（content_index\...\content_index.*.tsv）の先頭行 "版=<contentIndexVersion>"
            $contentIndexFile = @(Get-ChildItem -LiteralPath "$wsDir\content_index" -Recurse -File -Filter "${contentIndexFileNamePrefix}.*.tsv" -ErrorAction SilentlyContinue | Select-Object -First 1)
            if ($contentIndexFile.Count -eq 0) {
                $problems.Add("contentIndexFileNamePrefix（${contentIndexFileNamePrefix}）に当たる本文インデックスのファイルが無い")
            } else {
                $text = [System.IO.File]::ReadAllText($contentIndexFile[0].FullName, [System.Text.Encoding]::Unicode)
                $m = [regex]::Match($text, "版=(\d+)")
                if (!$m.Success -or $m.Groups[1].Value -ne [string]${contentIndexVersion}) {
                    $problems.Add("contentIndexVersion（${contentIndexVersion}）が $($contentIndexFile[0].Name) に見つからない")
                }
            }

            # sourceFolderFileName・systemIndexFileName
            if (@(Get-ChildItem -LiteralPath "$wsDir\content_index" -Recurse -File -Filter ${sourceFolderFileName} -ErrorAction SilentlyContinue).Count -eq 0) {
                $problems.Add("sourceFolderFileName（${sourceFolderFileName}）が無い")
            }
            if (@(Get-ChildItem -LiteralPath "$wsDir\system_index" -Recurse -File -Filter ${systemIndexFileName} -ErrorAction SilentlyContinue).Count -eq 0) {
                $problems.Add("systemIndexFileName（${systemIndexFileName}）が無い")
            }

            # Workspace クラスの葉の名前（content_index・system_index・system_index_state.tsv・ingest_status.tsv）
            $ws = [Workspace]::new($wsDir)
            foreach ($leaf in @($ws.IndexDir, $ws.SystemIndexDir, $ws.SystemIndexStateFile, $ws.StatusFile)) {
                if (!(Test-Path -LiteralPath $leaf)) {
                    $problems.Add("Workspace の $([System.IO.Path]::GetFileName($leaf)) が無い")
                }
            }

            # statusColumns: ingest_status.tsv の見出し行
            if (Test-Path -LiteralPath $ws.StatusFile) {
                $lines = [System.IO.File]::ReadAllLines($ws.StatusFile, [System.Text.Encoding]::UTF8)
                $headerLine = @($lines | Where-Object { $_ -ne "" -and $_ -notmatch "^クロール対象フォルダ`t" } | Select-Object -First 1)
                if ($headerLine.Count -eq 0 -or $headerLine[0] -ne (${statusColumns} -join "`t")) {
                    $problems.Add("ingest_status.tsv の見出しが statusColumns と違う")
                }
            }

            # indexArchiveFormatVersion・indexArchiveManifestFileName・indexArchiveStatusEntryName: export.zip の目録
            $zipPath = "$($sample.FullName)\export.zip"
            if (Test-Path -LiteralPath $zipPath) {
                $manifestText = readZipEntryText $zipPath ${indexArchiveManifestFileName}
                if ($null -eq $manifestText) {
                    $problems.Add("export.zip に indexArchiveManifestFileName（${indexArchiveManifestFileName}）が無い")
                } elseif ([int]($manifestText | ConvertFrom-Json).formatVersion -ne ${indexArchiveFormatVersion}) {
                    $problems.Add("export.zip の formatVersion が indexArchiveFormatVersion（${indexArchiveFormatVersion}）と違う")
                }
                if ((getZipEntryNames $zipPath) -cnotcontains ${indexArchiveStatusEntryName}) {
                    $problems.Add("export.zip に indexArchiveStatusEntryName（${indexArchiveStatusEntryName}）が無い")
                }
            } else {
                $problems.Add("export.zip が無い")
            }

            $problemsBySample[$sample.Name] = $problems
        }

        $complete = @($problemsBySample.Keys | Where-Object { $problemsBySample[$_].Count -eq 0 })
        $report = ($problemsBySample.Keys | ForEach-Object { "${_}: " + ($problemsBySample[$_] -join " / ") }) -join "`n"
        $complete.Count | Should -BeGreaterThan 0 -Because "目印がすべてそろった見本が無い`n$report"
    }
}

Describe "前の版のファイル（compat\settings）" -Tag Meta {
    It "それぞれの見本に、setting.config と expected.json がそろっている" {
        $problems = New-Object System.Collections.Generic.List[string]
        foreach ($dir in @(Get-ChildItem -LiteralPath ${compatSettingsRoot} -Directory)) {
            foreach ($entry in @("setting.config", "expected.json")) {
                if (!(Test-Path -LiteralPath "$($dir.FullName)\$entry")) {
                    $problems.Add("$($dir.Name) に $entry がありません")
                }
            }
        }
        ($problems -join "`n") | Should -Be ""
    }

    It "newSettings の全キーが、どれか 1 つの見本の expected.json にある（期待値の書き忘れを止める）" {
        $samples = @(Get-ChildItem -LiteralPath ${compatSettingsRoot} -Directory)
        $samples.Count | Should -BeGreaterThan 0 -Because "見本が 1 つも無い"

        $covered = New-Object "System.Collections.Generic.HashSet[string]"
        foreach ($sample in $samples) {
            $expectedPath = "$($sample.FullName)\expected.json"
            if (!(Test-Path -LiteralPath $expectedPath)) { continue }
            $expected = [System.IO.File]::ReadAllText($expectedPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
            foreach ($name in @($expected.PSObject.Properties.Name)) {
                [void]$covered.Add($name)
            }
        }

        $missing = @((newSettings).Keys | Where-Object { -not $covered.Contains($_) })
        ($missing -join ", ") | Should -Be "" -Because "見本の expected.json に無い newSettings のキー"
    }
}
