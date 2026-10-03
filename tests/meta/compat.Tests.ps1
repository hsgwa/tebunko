# 前の版のファイル（見本・golden。tests/testdata/compat/index/<見本の名前>/）が、
# 見本として要る部分をそろえていること・今のコードの形式の目印を含んでいることを確かめる。
# 中身が今のコードで読めること自体は tests/tebunko/indexer/index_compat.Tests.ps1 で確かめる（ここでは読まない）。
# zip の中の content_index/ という決め打ちの文字列は、index_compat.Tests.ps1 の check i が確かめる（ここでは見ない）
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"

    ${compatIndexRoot} = "${testDataDir}\compat\index"
    ${requiredEntries} = @("source", "ws", "export.zip", "file_times.tsv", "expected.json")
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

    It "今のコードの形式の目印が、見本のどれかに残っている" {
        $samples = @(Get-ChildItem -LiteralPath ${compatIndexRoot} -Directory)
        $samples.Count | Should -BeGreaterThan 0 -Because "見本が 1 つも無い"
        $problems = New-Object System.Collections.Generic.List[string]

        # packVersion: 集約ファイル（content_index\...\content_index.*.tsv）の先頭行 "版=<packVersion>"
        $packFile = @($samples | ForEach-Object { Get-ChildItem -LiteralPath "$($_.FullName)\ws\content_index" -Recurse -File -Filter "${packFileNamePrefix}.*.tsv" -ErrorAction SilentlyContinue } | Select-Object -First 1)
        if ($packFile.Count -eq 0) {
            $problems.Add("packFileNamePrefix（${packFileNamePrefix}）に当たる集約ファイルが見本に無い")
        } else {
            $text = [System.IO.File]::ReadAllText($packFile[0].FullName, [System.Text.Encoding]::Unicode)
            $m = [regex]::Match($text, "版=(\d+)")
            if (!$m.Success -or $m.Groups[1].Value -ne [string]${packVersion}) {
                $problems.Add("packVersion（${packVersion}）が $($packFile[0].FullName) に見つからない")
            }
        }

        # sourceFolderFileName・systemIndexFileName
        $sourceFolderHit = @($samples | ForEach-Object { Get-ChildItem -LiteralPath "$($_.FullName)\ws\content_index" -Recurse -File -Filter ${sourceFolderFileName} -ErrorAction SilentlyContinue })
        if ($sourceFolderHit.Count -eq 0) {
            $problems.Add("sourceFolderFileName（${sourceFolderFileName}）が見本に無い")
        }
        $systemIndexHit = @($samples | ForEach-Object { Get-ChildItem -LiteralPath "$($_.FullName)\ws\system_index" -Recurse -File -Filter ${systemIndexFileName} -ErrorAction SilentlyContinue })
        if ($systemIndexHit.Count -eq 0) {
            $problems.Add("systemIndexFileName（${systemIndexFileName}）が見本に無い")
        }

        # Workspace クラスの葉の名前（content_index・system_index・system_index_state.tsv・ingest_status.tsv）
        foreach ($sample in $samples) {
            $ws = [Workspace]::new("$($sample.FullName)\ws")
            foreach ($leaf in @($ws.IndexDir, $ws.SystemIndexDir, $ws.SystemIndexStateFile, $ws.StatusFile)) {
                if (!(Test-Path -LiteralPath $leaf)) {
                    $problems.Add("$($sample.Name) に Workspace の $([System.IO.Path]::GetFileName($leaf)) が無い")
                }
            }
        }

        # statusColumns: ingest_status.tsv の見出し行
        foreach ($sample in $samples) {
            $statusPath = "$($sample.FullName)\ws\ingest_status.tsv"
            if (Test-Path -LiteralPath $statusPath) {
                $lines = [System.IO.File]::ReadAllLines($statusPath, [System.Text.Encoding]::UTF8)
                $headerLine = @($lines | Where-Object { $_ -ne "" -and $_ -notmatch "^クロール対象フォルダ`t" } | Select-Object -First 1)
                if ($headerLine.Count -eq 0 -or $headerLine[0] -ne (${statusColumns} -join "`t")) {
                    $problems.Add("$($sample.Name) の ingest_status.tsv の見出しが statusColumns と違う")
                }
            }
        }

        # indexArchiveFormatVersion・indexArchiveManifestFileName・indexArchiveStatusEntryName: export.zip の目録
        foreach ($sample in $samples) {
            $zipPath = "$($sample.FullName)\export.zip"
            if (!(Test-Path -LiteralPath $zipPath)) { continue }
            $archive = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
            try {
                $manifestEntry = $archive.Entries | Where-Object { $_.FullName -ceq ${indexArchiveManifestFileName} } | Select-Object -First 1
                if (!$manifestEntry) {
                    $problems.Add("$($sample.Name) の export.zip に indexArchiveManifestFileName（${indexArchiveManifestFileName}）が無い")
                } else {
                    $stream = $manifestEntry.Open()
                    try {
                        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
                        $manifest = $reader.ReadToEnd() | ConvertFrom-Json
                    } finally {
                        $stream.Dispose()
                    }
                    if ([int]$manifest.formatVersion -ne ${indexArchiveFormatVersion}) {
                        $problems.Add("$($sample.Name) の export.zip の formatVersion が indexArchiveFormatVersion（${indexArchiveFormatVersion}）と違う")
                    }
                }
                $statusEntry = $archive.Entries | Where-Object { $_.FullName -ceq ${indexArchiveStatusEntryName} } | Select-Object -First 1
                if (!$statusEntry) {
                    $problems.Add("$($sample.Name) の export.zip に indexArchiveStatusEntryName（${indexArchiveStatusEntryName}）が無い")
                }
            } finally {
                $archive.Dispose()
            }
        }

        ($problems -join "`n") | Should -Be ""
    }
}
