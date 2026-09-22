# 制限モードが使う関数の呼び出し例（tests/meta/clm.Tests.ps1 が使う）。
#
# 同じ呼び出しを FullLanguage と模擬の制限言語モードで動かし、結果（JSON にしたもの）が同じことを確かめる。
# このファイルも制限言語モードで読み込むため、制限言語モードで使える書き方だけで書く。
# 呼び出す前に、ファイルを書いてよいフォルダ（空）を $caseDir に入れておく。結果にはそのフォルダのパスを入れない
# （2 つのモードでフォルダが違うため。パスは getCaseRelative で $caseDir からの相対にする）。

function getCaseBytes {
    # ファイルの中身（バイト列）を比べられる文字列にする
    param ([string]$path)
    return (@(Get-Content -LiteralPath $path -Encoding Byte -ReadCount 0) -join ",")
}

function getCaseRelative {
    param ([string]$path)
    return $path.Replace($caseDir, "<dir>")
}

$clmCases = [ordered]@{
    # --- shared/core/text.ps1 ---
    "replaceCellNewLine" = { replaceCellNewLine "a`t`"x`r`ny`"`t`"p`nq`rr`"`n`"s`"" }
    "formatTsv"          = {
        @(
            (formatTsv "a`tb`t`r`n`r`nc`r`n`r`n" 1 1),
            (formatTsv "a`t`"x`ny`"" 3 2),
            (formatTsv "`r`n  `r`n" 2 1),
            (formatTsv "" 1 1)
        )
    }
    "countTsvFields"     = { @((countTsvFields "a`tb"), (countTsvFields "`"x`ty`"`tz"), (countTsvFields "")) }
    "toColumnName"       = { @(1, 26, 27, 52, 702, 703, 16384 | ForEach-Object { toColumnName $_ }) }
    "splitTsvCells"      = { @((splitTsvCells "a`t`"b`"`"c`"`t`t`"d`te`"x") -join "|") }
    "cellNewLine"        = { [int][char]${cellNewLine} }

    # --- shared/core/fs.ps1 ---
    "getPathLeaf"        = { @("a\b.tsv", "a\b\", "C:", "C:\", "x", "\\s\share\f.xlsx", "a/b" | ForEach-Object { getPathLeaf $_ }) }
    "getPathParent"      = { @("a\b.tsv", "a\b\c", "b.tsv", "a/b" | ForEach-Object { getPathParent $_ }) }
    "getPathStem"        = { @("a\b.tsv", "c", "d.e.f", ".x" | ForEach-Object { getPathStem $_ }) }
    "toSafeFileName"     = { toSafeFileName 'a<b>c\d*e"f:g?h|i/j' }
    "toLongPath"         = { @("C:\a", "\\s\share\a", "\\?\C:\a", "C:/a", "" | ForEach-Object { toLongPath $_ }) }
    "fromLongPath"       = { @("\\?\C:\a", "\\?\UNC\s\share\a", "C:\a" | ForEach-Object { fromLongPath $_ }) }
    "formatFileTime"     = { formatFileTime (Get-Date -Year 2024 -Month 3 -Day 4 -Hour 5 -Minute 6 -Second 7) }
    "writeListFile"      = {
        $a = Join-Path $caseDir "list\a.txt"
        writeListFile $a @("1行目", "", "𠮷")
        $b = Join-Path $caseDir "list\b.txt"
        writeListFile $b @()
        @((getCaseBytes $a), (getCaseBytes $b), (@(readListFile $a) -join "|"))
    }
    "writeTextLinesAtomic" = {
        $a = Join-Path $caseDir "atomic\a.tsv"
        writeTextLinesAtomic $a @("x`ty", "あ")
        writeTextLinesAtomic $a @("上書き")
        @((getCaseBytes $a), (Test-Path -LiteralPath "$a.tmp"))
    }
    "writeUtf8NoBom"     = {
        $a = Join-Path $caseDir "nobom.json"
        writeUtf8NoBom $a "{`"a`":`"あ𠮷é`"}`r`n"
        $b = Join-Path $caseDir "empty.json"
        writeUtf8NoBom $b ""
        @((getCaseBytes $a), (Get-Item -LiteralPath $b).Length)
    }
    "removeDirectoryRetry" = {
        $d = Join-Path $caseDir "remove\sub"
        New-Item -ItemType Directory -Path $d -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $d "f.txt") -Value "x"
        removeDirectoryRetry (Join-Path $caseDir "remove")
        removeDirectoryRetry (Join-Path $caseDir "none")
        Test-Path -LiteralPath (Join-Path $caseDir "remove")
    }

    # --- shared/core/folder.ps1 ---
    "normalizeFolderPath" = {
        $env:tebunko_grep_CLM_CASE = "C:\data\見積"
        try {
            @(
                '  "C:\data\"  ', "D:\", "D:", "\\server\share\", "C:/data/見積", "//server/share/見積",
                "\\?\C:\data\見積", "\\?\UNC\server\share\見積", "C:\data\\見積", "C:\data\.\見積",
                "C:\data\売上\..\見積", "\\server\share\売上\..\見積", "C:\..\..\a", "\\server\share\..\..\a",
                "%tebunko_grep_CLM_CASE%", "%tebunko_grep_CLM_CASE%\2024", "%tebunko_grep_CLM_NONE%\a",
                "work\index", ".\work\index", "C:\data*", "\\server", "\server\share", "", "C:\a. \b.", "C:\a\...\b", "C:\a\b. ", "C:\a\...", "C:\a\b\..\..\..", "D:\x\..", "\\server\share\x\.. "
            ) | ForEach-Object { (normalizeFolderPath $_).Replace($rootDir, "<root>") }
        } finally {
            Remove-Item -LiteralPath "env:tebunko_grep_CLM_CASE"
        }
    }
    "getPathUnderFolder" = {
        @(
            (getPathUnderFolder "C:\data\見積\2024\a.xlsx" "C:\data\見積"),
            (getPathUnderFolder "c:\DATA\見積\" "C:\data\見積"),
            (getPathUnderFolder "D:\a.xlsx" "D:\"),
            ($null -eq (getPathUnderFolder "C:\data\見積2\a.xlsx" "C:\data\見積"))
        )
    }
    "getFolderPathAliases" = {
        $drives = @{ "Z:" = "\\server\share"; "X:" = "C:\data" }
        @(
            (@(getFolderPathAliases "Z:\見積" $drives) -join "|"),
            (@(getFolderPathAliases "\\server\share\見積" $drives) -join "|"),
            (@(getFolderPathAliases "Z:\" $drives) -join "|"),
            (@(getFolderPathAliases "D:\見積" $drives) -join "|")
        )
    }
    "testSameFolder"     = {
        $drives = @{ "Z:" = "\\server\share" }
        @((testSameFolder "Z:\見積\" "\\SERVER\share\見積" $drives), (testSameFolder "Z:\見積" "Z:\見積2" $drives))
    }
    "getFolderLeafName"  = { @("C:\data\見積", "D:\", "\\server\share", "C:\data\見積\" | ForEach-Object { getFolderLeafName $_ }) }

    # --- tebunko_grep/core/settings_grep.ps1 ---
    "settings"           = {
        $path = Join-Path $caseDir "settings\setting.config"
        writeTargetFolders @(
            (New-Object PSObject -Property ([ordered]@{ Name = "見積"; Path = "C:\data\見積"; Enabled = $true })),
            (New-Object PSObject -Property ([ordered]@{ Name = "見積"; Path = "C:\data\見積\"; Enabled = $false })),
            (New-Object PSObject -Property ([ordered]@{ Name = "SAMPLE"; Path = "D:\sample"; Enabled = $false })),
            (New-Object PSObject -Property ([ordered]@{ Name = "sample"; Path = "E:\sample"; Enabled = $true }))
        ) $path
        writeIndexSources @((New-Object PSObject -Property ([ordered]@{ Name = "外部"; Path = "\\server\share\外部" }))) $path
        setIndexSourceFolder "営業" "F:\営業" $path
        writeSearchExcludes @((New-Object PSObject -Property ([ordered]@{ Path = "C:\data\見積\古い"; Subfolders = $false }))) $path
        writeSearchOption @{ UseRegex = $true; FileFilter = "*.xlsx;!*old*" } $path
        writeOpenMode ${openModeReadOnly} $path
        @(
            (getCaseBytes $path),
            (@(getTargetFolders $path) | ConvertTo-Json -Compress),
            (@(readIndexSources $path) | ConvertTo-Json -Compress),
            (@(readSearchExcludes $path) | ConvertTo-Json -Compress),
            ((readSearchOption $path).GetEnumerator() | Sort-Object Key | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ";",
            (readOpenMode $path)
        )
    }
    "readSettings(文字列・壊れた値)" = {
        $path = Join-Path $caseDir "settings2\setting.config"
        New-Item -ItemType Directory -Path (Join-Path $caseDir "settings2") -Force | Out-Null
        Set-Content -LiteralPath $path -Encoding UTF8 -Value '{"useRegex":" True ","caseSensitive":"no","fileFilter":5,"targetFolders":null,"includeShapes":0}'
        $settings = readSettings $path
        @($settings.useRegex, $settings.caseSensitive, $settings.fileFilter, $settings.targetFolders.Count, $settings.includeShapes)
    }
    "readSettings(空のファイル)" = {
        $path = Join-Path $caseDir "settings3\setting.config"
        New-Item -ItemType Directory -Path (Join-Path $caseDir "settings3") -Force | Out-Null
        Set-Content -LiteralPath $path -Value ([byte[]](0xEF, 0xBB, 0xBF)) -Encoding Byte
        (readSettings $path) | ConvertTo-Json -Compress
    }
    "toSettingBool"      = { @((toSettingBool "FALSE" $true), (toSettingBool " true " $false), (toSettingBool "x" $true), (toSettingBool 1 $false)) }

    # --- tebunko_grep/index/index_name.ps1 ---
    "newIndexName"       = {
        @(
            (newIndexName "C:\data\見積" @("見積", "見積(2)")),
            (newIndexName "D:\" $null),
            (newIndexName "\\server\share" "SHARE"),
            (newIndexName "" @("フォルダ")),
            (newIndexName ("C:\" + ("あ" * 255)) @("あ" * 255)).Length
        )
    }
    "assignIndexNames"   = {
        $targets = @(
            (New-Object PSObject -Property ([ordered]@{ Name = ""; Path = "C:\a\見積"; Enabled = $true })),
            (New-Object PSObject -Property ([ordered]@{ Name = "見積"; Path = "C:\b\見積"; Enabled = $false })),
            (New-Object PSObject -Property ([ordered]@{ Name = ""; Path = "C:\c\売上"; Enabled = $true }))
        )
        $previous = @(@{ Path = "c:\C\売上"; Name = "昔の売上" })
        @(assignIndexNames $targets $previous) | ConvertTo-Json -Compress
    }
    "splitIndexRelPath"  = { @((splitIndexRelPath "見積\2024\a"), (splitIndexRelPath "見積")) | ForEach-Object { "$($_.Name)|$($_.Rest)" } }
    "testIndexName"      = { @("", " a", ("a" * 256), "a:b", "a.", "con.txt", "見積", "ok" | ForEach-Object { testIndexName $_ @("見積") }) }
    "encodeIndexPlace"   = { @("売上", "a_b%c", "x:y*z?", "_", "`t", "" | ForEach-Object { encodeIndexPlace $_ }) }
    "decodeIndexPlace"   = { @("a%5Fb%25c", "%3A%2A%3F", "%5f", "100%", "%09" | ForEach-Object { decodeIndexPlace $_ }) }
    "toIndexFileName"    = { toIndexFileName "シート_1" }
    "splitIndexFileName" = { @("ブック名.xlsx_シート名.tsv", "資料.pptx_スライド003%5Fノート.tsv", "x.tsv") | ForEach-Object { $s = splitIndexFileName $_; "$($s.book)|$($s.sheet)" } }
    "describePlace"      = {
        @(
            @("a.xlsx", "売上"), @("a.xlsx", "売上[図形]"), @("b.docx", "ページ003"), @("b.docx", "脚注"),
            @("c.pptx", "スライド002（非表示）"), @("c.pptx", "スライド002_ノート"), @("c.pptx", "スライド001[コメント]"), @("b.docx", "")
        ) | ForEach-Object { $d = describePlace $_[0] $_[1]; "$($d.Place)|$($d.Kind)" }
    }
    "splitIndexTsvPath"  = {
        @("見積\2024\a.xlsx\売上%5F1.tsv", "見積\a.xlsx_売上.tsv", "a.docx\ページ001.tsv", "見積\x.tsv") | ForEach-Object {
            $s = splitIndexTsvPath $_
            "$($s.Book)|$($s.Place)|$($s.RelDir)"
        }
    }

    # --- tebunko_grep/search/search_query.ps1 ---
    "isValidRegex"       = { @((isValidRegex "a+"), (isValidRegex "a(")) }
    "getRegexScanMode"   = { @("見積", "a.b", "\s", "[^a]", "\A", "(?i)a", "(?=a)", "(?<n>a)", "a\" | ForEach-Object { getRegexScanMode $_ }) }
    "newSearchRegex"     = {
        @(
            @("a.b", $true, $false), @("a.b", $false, $true), @("a(", $false, $false), @("^x$", $false, $false)
        ) | ForEach-Object {
            $r = newSearchRegex $_[0] $_[1] $_[2]
            "$($r.Regex)|$($r.Regex.Options)|$($r.Regex.MatchTimeout)|$($r.SimpleMatch)|$($r.TextRegex.Options)|$($r.ScanMode)"
        }
    }
    "newPlaceExclude"    = { @((newPlaceExclude $false $true), (newPlaceExclude $true $false), (newPlaceExclude $false $false), ($null -eq (newPlaceExclude $true $true))) | ForEach-Object { "$_" } }
    "newFileFilter"      = {
        @("*.xlsx;見積*;!*old*", "見積", "！古い；a?c", "") | ForEach-Object {
            $f = newFileFilter $_
            "$($f.Include)|$($f.Exclude)|$(if ($f.Include) { $f.Include.Options })"
        }
    }

    # --- tebunko_grep/search/source_map.ps1・index/index_store.ps1 ---
    "sourceMap"          = {
        $index = Join-Path $caseDir "index"
        New-Item -ItemType Directory -Path (Join-Path $index "見積\2024") -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $index "D") -Force | Out-Null
        writeSourceFolderFile @(
            (New-Object PSObject -Property ([ordered]@{ Name = "見積"; Path = "C:\data\見積" })),
            (New-Object PSObject -Property ([ordered]@{ Name = "D"; Path = "D:\" })),
            (New-Object PSObject -Property ([ordered]@{ Name = "無い"; Path = "E:\無い" }))
        ) $index
        $status = Join-Path $caseDir "取り込み一覧.tsv"
        Set-Content -LiteralPath $status -Encoding UTF8 -Value @(
            "${statusFolderKey}`tC:\移した\見積`t見積", "${statusFolderKey}`tG:\g`tG", (${statusColumns} -join "`t"), "a.xlsx`t2024`t1`t済`t1`t2024`t`t1"
        )
        $nameMap = getIndexNameMap $status
        $settingsPath = Join-Path $caseDir "none.config"
        # getSourceLocation は既定の取り込み一覧・設定ファイルを読むため、ツール本体のものを読まないよう差し替える
        # （関数の引数の既定値は、呼び出し元のこの変数から決まる）
        $statusFile = $status
        $settingsFile = $settingsPath
        $maps = @{}
        @(
            (getCaseBytes (Join-Path $index "見積\${sourceFolderFileName}")),
            ((readSourceFolderFile (Join-Path $index "d")).GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }),
            (@($nameMap.Keys | Sort-Object) -join "|"), $nameMap["見積"],
            ((getSourceFolderMap (Join-Path $index "見積") $status $settingsPath).GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }),
            (resolveSourcePath @{ Root = $index; RelDir = "見積\2024"; Book = "a.xlsx" } $maps),
            (getSourceLocation @{ Root = (Join-Path $index "見積"); RelDir = "2024"; Book = "a.xlsx" } $maps | ConvertTo-Json -Compress),
            (getSourceLocation @{ Root = (Join-Path $caseDir "よそ"); RelDir = "x"; Book = "a.xlsx" } $maps | ConvertTo-Json -Compress).Replace($caseDir.Replace("\", "\\"), "<dir>"),
            (joinSourcePath "D:\" "" "a.xlsx"), (joinSourcePath "C:\a" "b\c" "")
        )
    }
}
