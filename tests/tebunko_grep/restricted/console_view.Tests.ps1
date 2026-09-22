# 制限モードのコンソールの文言と入力の読み方（tebunko_grep\restricted\console_view.ps1）のテスト。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\tebunko_grep\restricted\console_view.ps1"

function newOption {
    param ([hashtable]$override = @{})
    $option = @{ UseRegex = $false; CaseSensitive = $false; FileFilter = ""; IncludeShapes = $true; IncludeComments = $true }
    foreach ($key in $override.Keys) { $option[$key] = $override[$key] }
    return $option
}

Describe "describeRestrictedOption / getOptionMenuLines" -Tag Unit {
    It "既定の条件" {
        describeRestrictedOption (newOption) | Should Be "文字どおり・大文字と小文字を区別しない・対象ファイル すべて・図形とコメントも検索"
    }

    It "変えた条件" {
        describeRestrictedOption (newOption @{ UseRegex = $true; CaseSensitive = $true; FileFilter = "*.xlsx"; IncludeComments = $false }) |
            Should Be "正規表現・大文字と小文字を区別する・対象ファイル *.xlsx・図形も検索（コメントは除く）"
        describeRestrictedOption (newOption @{ IncludeShapes = $false }) | Should Be "文字どおり・大文字と小文字を区別しない・対象ファイル すべて・コメントも検索（図形は除く）"
        describeRestrictedOption (newOption @{ IncludeShapes = $false; IncludeComments = $false }) | Should Be "文字どおり・大文字と小文字を区別しない・対象ファイル すべて・図形とコメントは除く"
    }

    It "メニューに番号と今の値を出す" {
        $lines = getOptionMenuLines (newOption @{ UseRegex = $true; FileFilter = "見積*" })
        $lines.Count | Should Be 6
        $lines[1] | Should Be "  1. 正規表現を使う: オン"
        $lines[3] | Should Be "  3. 対象ファイル: 見積*"
        (getOptionMenuLines (newOption))[3] | Should Be "  3. 対象ファイル: すべて"
    }
}

Describe "インデックスの選び方" -Tag Unit {
    $indexes = @(
        @{ Name = "見積"; Path = "C:\t\work\index\見積"; SourcePath = "D:\見積" },
        @{ Name = "営業"; Path = "C:\t\work\index\営業"; SourcePath = "" },
        @{ Name = "技術"; Path = "C:\t\work\index\技術"; SourcePath = "" },
        @{ Name = "総務"; Path = "C:\t\work\index\総務"; SourcePath = "" }
    )

    It "インデックスごと外したものだけを「検索しない」とする（中のフォルダだけ外したものは検索する）" {
        $excludes = @(@{ Path = "C:\t\work\index\営業"; Subfolders = $true }, @{ Path = "C:\t\work\index\見積\古い"; Subfolders = $true }, @{ Path = "C:\t\work\index\技術"; Subfolders = $false })
        testIndexChecked $indexes[0] $excludes | Should Be $true
        testIndexChecked $indexes[1] $excludes | Should Be $false
        testIndexChecked $indexes[2] $excludes | Should Be $true
    }

    It "検索するインデックスの表示（3 件まで・全部外した・無い）" {
        describeRestrictedIndexes $indexes @() | Should Be "見積、営業、技術 ほか 1 件（4 / 4 件）"
        describeRestrictedIndexes $indexes @(@{ Path = "C:\t\work\index\営業"; Subfolders = $true }) | Should Be "見積、技術、総務（3 / 4 件）"
        describeRestrictedIndexes $indexes @($indexes | ForEach-Object { @{ Path = $_.Path; Subfolders = $true } }) | Should Match "どのインデックスも選んでいません"
        describeRestrictedIndexes @() @() | Should Match "インデックスがありません"
    }

    It "メニューに番号・選んでいるか・元のフォルダを出す" {
        $lines = getIndexMenuLines $indexes @(@{ Path = "C:\t\work\index\営業"; Subfolders = $true })
        $lines[1] | Should Be "  1. [x] 見積（元のフォルダ: D:\見積）"
        $lines[2] | Should Be "  2. [ ] 営業"
    }

    It "切り替え: 外すときはフォルダ以下すべて、戻すときはそのインデックスの中の記録をすべて消す" {
        $next = switchIndexExclude $indexes[0] @(@{ Path = "C:\t\work\index\営業"; Subfolders = $true })
        @($next | ForEach-Object { "$($_.Path)|$($_.Subfolders)" }) -join "," | Should Be "C:\t\work\index\営業|True,C:\t\work\index\見積|True"
        $back = switchIndexExclude $indexes[0] @(@{ Path = "C:\t\work\index\見積"; Subfolders = $true }, @{ Path = "C:\t\work\index\見積\古い"; Subfolders = $false }, @{ Path = "C:\t\work\index\営業"; Subfolders = $true })
        @($back | ForEach-Object { $_.Path }) -join "," | Should Be "C:\t\work\index\営業"
    }
}

Describe "parseMenuNumber / parseHitChoice" -Tag Unit {
    It "番号（全角も）・範囲の外・番号でないもの" {
        parseMenuNumber " 2 " 5 | Should Be 2
        parseMenuNumber "３" 5 | Should Be 3
        parseMenuNumber "6" 5 | Should Be -1
        parseMenuNumber "0" 5 | Should Be -1
        parseMenuNumber "a" 5 | Should Be 0
        parseMenuNumber "" 5 | Should Be 0
    }

    It "ヒットの選び方: 番号はファイル、f番号はフォルダ" {
        $choice = parseHitChoice "3" 5
        "$($choice.Kind)|$($choice.Number)" | Should Be "file|3"
        $choice = parseHitChoice "f 2" 5
        "$($choice.Kind)|$($choice.Number)" | Should Be "folder|2"
        $choice = parseHitChoice "Ｆ１" 5
        "$($choice.Kind)|$($choice.Number)" | Should Be "folder|1"
        ($null -eq (parseHitChoice "9" 5)) | Should Be $true
        ($null -eq (parseHitChoice "x" 5)) | Should Be $true
    }
}

Describe "getSearchSummaryLines / formatHitListLine / getResultInfoItems" -Tag Unit {
    It "件数・見つからない・打ち切り・文字どおりに切り替えた" {
        (getSearchSummaryLines "りんご" @{ Hits = @(1, 2); Total = 10; Truncated = $false; SimpleMatch = $true } $false 10000) -join "|" |
            Should Be "「りんご」は 2 件見つかりました（TSV 10 件を検索）。"
        (getSearchSummaryLines "a(" @{ Hits = @(); Total = 3; Truncated = $false; SimpleMatch = $true } $true 10000) -join "|" |
            Should Be "「a(」は正規表現として正しくないため、文字どおりに検索しました。|「a(」は見つかりませんでした（TSV 3 件を検索）。"
        (getSearchSummaryLines "a" @{ Hits = @(1); Total = 3; Truncated = $true; SimpleMatch = $false } $true 10000) -join "|" |
            Should Be "「a」は 10000 件を超えたため、10000 件で打ち切りました（条件を絞ってください）。"
    }

    It "ヒットを 1 行で出し、幅に収まらなければ切り詰める" {
        $hit = @{ RelDir = "見積\2024"; Book = "a.xlsx"; Location = "売上"; LineNumber = 12; Line = "りんご`t100`t`"青森$([char]0x2028)ふじ`"" }
        formatHitListLine 3 $hit | Should Be "   3. 見積\2024\a.xlsx  [シート] 売上  12 行目: りんご 100 `"青森 ふじ`""
        $short = formatHitListLine 3 $hit 20
        $short.Length | Should Be 20
        $short.EndsWith("…") | Should Be $true
    }

    It "条件のシートの項目" {
        $items = getResultInfoItems "りんご" @{ UseRegex = $false; CaseSensitive = $false; FileFilter = ""; IncludeShapes = $true; IncludeComments = $true } "見積（1 / 1 件）" @{ Hits = @(1); Total = 5; Truncated = $false } 10000 (Get-Date -Year 2024 -Month 1 -Day 2 -Hour 3 -Minute 4 -Second 5)
        @($items | ForEach-Object { $_[0] }) -join "," | Should Be "検索ワード,条件,検索したインデックス,件数,検索した TSV,検索した日時,見方"
        $items[3][1] | Should Be "1 件"
        $items[5][1] | Should Be "2024/01/02 03:04:05"
        (getResultInfoItems "a" @{} "" @{ Hits = @(); Total = 0; Truncated = $true } 10000 (Get-Date))[3][1] | Should Be "10000 件で打ち切り"
    }
}

Describe "formatHitListLine（相対フォルダが無い）" -Tag Unit {
    It "ファイル名だけを出す" {
        formatHitListLine 1 @{ RelDir = ""; Book = "a.docx"; Location = "ページ001"; LineNumber = 2; Line = "x" } | Should Be "   1. a.docx  [ページ] 1（目安）  2 行目: x"
    }
}
