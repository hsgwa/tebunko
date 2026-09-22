# 制限モードの検索（tebunko_grep\restricted\restricted_search.ps1）のテスト。
# いつもの画面の検索（getIndexTsvFiles・searchIndex）と、同じ TSV から同じヒットを返すことを突き合わせる。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\tebunko_grep\restricted\restricted_search.ps1"

# テスト用のインデックス（work\index にあたるフォルダ）を作る。改行は CRLF・LF・CR を混ぜる
function newSearchFixture {
    $root = Join-Path $TestDrive "index"
    $files = [ordered]@{
        "見積\a.xlsx\売上.tsv"            = "りんご`t100`r`nみかん`tりんごりんご`r`n`r`n`"青森産$([char]0x2028)ふじ`"`tりんご飴`r`n"
        "見積\a.xlsx\売上[図形].tsv"      = "B2`tりんごの図形`r`n"
        "見積\a.xlsx\売上[コメント].tsv"  = "C3`tりんごのコメント`r`n"
        "見積\2024\b.docx\ページ001.tsv"  = "1 行目 りんご`nApple`napple pie`n末尾に改行なし りんご"
        "見積\2024\c.pptx\スライド001.tsv" = "CR だけの改行`rりんご`rりんご 12-34`r"
        "見積\old.xls_Sheet1.tsv"         = "旧形式 りんご 5`r`n"
        "見積\空.xlsx\空.tsv"              = ""
        "営業\深い\階層\d.xlsm\一覧.tsv"   = "a b`ta`tb`r`nX1`tx2`r`n^りんご$`r`n"
    }
    foreach ($rel in $files.Keys) {
        $path = Join-Path $root $rel
        New-Item -ItemType Directory -Path (Split-Path $path -Parent) -Force | Out-Null
        [System.IO.File]::WriteAllText($path, $files[$rel], (New-Object System.Text.UTF8Encoding($true)))
    }
    return (Resolve-Path -LiteralPath $root).ProviderPath
}

# ヒットを比べられる文字列にする
function formatHits {
    param ([object[]]$hits)
    return @($hits | ForEach-Object { "$($_.RelPath)|$($_.RelDir)|$($_.FileName)|$($_.Book)|$($_.Location)|$($_.LineNumber)|$($_.Line)" })
}

Describe "searchRestricted（いつもの画面の検索と同じ結果）" -Tag Io {
    $root = newSearchFixture
    $indexes = @(
        @{ Name = "見積"; Path = "$root\見積" },
        @{ Name = "営業"; Path = "$root\営業" }
    )
    $targets = @($indexes | ForEach-Object { @{ Root = $root; RelPath = $_.Name; Recurse = $true } })
    $guiFiles = (getIndexTsvFiles $targets).Files
    $files = getRestrictedTsvFiles $indexes

    It "列挙した TSV の順と相対パスが同じ" {
        (@($files | ForEach-Object { $_.RelPath }) -join "|") | Should Be (@($guiFiles.Values | ForEach-Object { $_.RelPath }) -join "|")
    }

    $cases = @(
        @{ Word = "りんご"; Simple = $true },
        @{ Word = "りんご"; Simple = $true; Case = $true; Filter = "*.xlsx;!old*" },
        @{ Word = "りんご"; Simple = $true; Shapes = $false; Comments = $false },
        @{ Word = "りんご"; Simple = $true; Filter = "見積" },
        @{ Word = "apple"; Simple = $true },
        @{ Word = "apple"; Simple = $true; Case = $true },
        @{ Word = "^り"; Simple = $false },
        @{ Word = "ご$"; Simple = $false },
        @{ Word = "\d+-\d+"; Simple = $false },
        @{ Word = "a\sb"; Simple = $false },
        @{ Word = "[^x]1"; Simple = $false },
        @{ Word = "(?i)x\d"; Simple = $false },
        @{ Word = "\Aりんご"; Simple = $false },
        @{ Word = "青森産.ふじ"; Simple = $false },
        @{ Word = "^りんご$"; Simple = $true },
        @{ Word = "a("; Simple = $false },
        @{ Word = "存在しない"; Simple = $true }
    )
    foreach ($case in $cases) {
        $name = "「$($case.Word)」（文字どおり: $($case.Simple)、区別: $([bool]$case.Case)、対象: $($case.Filter)、図形: $($case.Shapes -ne $false)）"
        It $name {
            $gui = searchIndex $case.Word $guiFiles $case.Simple 0 100 $null $null ([bool]$case.Case) ([string]$case.Filter) 1 $null ($case.Shapes -ne $false) ($case.Comments -ne $false)
            $restricted = searchRestricted $case.Word $files $case.Simple ([bool]$case.Case) 0 ([string]$case.Filter) ($case.Shapes -ne $false) ($case.Comments -ne $false)
            ((formatHits $restricted.Hits) -join "`n") | Should Be ((formatHits $gui.Hits) -join "`n")
            $restricted.SimpleMatch | Should Be $gui.SimpleMatch
            $restricted.Total | Should Be $gui.Total
            $restricted.Truncated | Should Be $gui.Truncated
        }
    }

    foreach ($limit in @(1, 3, 7, 8)) {
        It "件数の上限 $limit で打ち切る（いつもの画面と同じ件数・打ち切りの判定）" {
            $gui = searchIndex "りんご" $guiFiles $true $limit 100 $null $null $false "" 1
            $restricted = searchRestricted "りんご" $files $true $false $limit
            ((formatHits $restricted.Hits) -join "`n") | Should Be ((formatHits $gui.Hits) -join "`n")
            $restricted.Truncated | Should Be $gui.Truncated
        }
    }

    It "照合が時間切れになったら、いつもの画面と同じメッセージで止める" {
        $big = Join-Path $root "営業\時間切れ.xlsx\重い.tsv"
        New-Item -ItemType Directory -Path (Split-Path $big -Parent) -Force | Out-Null
        [System.IO.File]::WriteAllText($big, ("a" * 30000) + "!`r`n")
        $heavy = getRestrictedTsvFiles @(@{ Name = "営業"; Path = "$root\営業" })
        $regexTimeout = [timespan]::FromMilliseconds(50)
        try {
            { searchRestricted "(a|aa)*$" $heavy $false } | Should Throw "正規表現の照合に時間がかかりすぎるため"
        } finally {
            Remove-Item -LiteralPath (Split-Path $big -Parent) -Recurse -Force
        }
    }

    It "進み具合を知らせる" {
        $many = 1..205 | ForEach-Object { @{ LongPath = (toLongPath "$root\見積\old.xls_Sheet1.tsv"); Root = $root; RelPath = "x$_"; RelDir = ""; FileName = "f"; Book = "f.xlsx"; Location = "s"; Size = 10 } }
        $script:progress = @()
        [void](searchRestricted "りんご" $many $true $false 0 "" $true $true { param($done, $total) $script:progress += "$done/$total" })
        ($script:progress -join ",") | Should Be "100/205,200/205"
    }

    It "読めない TSV は飛ばす" {
        $gone = @(@{ LongPath = (toLongPath "$root\無い.tsv"); Root = $root; RelPath = "無い.tsv"; RelDir = ""; FileName = "無い.tsv"; Book = "無い.tsv"; Location = ""; Size = 0 })
        (searchRestricted "りんご" $gone $true).Hits.Count | Should Be 0
    }

    It "大きな TSV（64MB 超とみなしたもの）も 1 行ずつ読んで同じ結果" {
        $sized = @($files | ForEach-Object { $copy = @{}; foreach ($k in $_.Keys) { $copy[$k] = $_[$k] }; $copy.Size = 100MB; $copy })
        $gui = searchIndex "りんご" $guiFiles $true 0 100 $null $null $false "" 1
        ((formatHits (searchRestricted "りんご" $sized $true).Hits) -join "`n") | Should Be ((formatHits $gui.Hits) -join "`n")
    }
}

Describe "getRestrictedTsvFiles（検索から外したフォルダ）" -Tag Io {
    $root = newSearchFixture
    $indexes = @(@{ Name = "見積"; Path = "$root\見積" }, @{ Name = "営業"; Path = "$root\営業" }, @{ Name = "無い"; Path = "$root\無い" })

    It "フォルダ以下すべてを外す・直下のファイルだけを外す・インデックスごと外す" {
        $excludes = @(
            @{ Path = "$root\見積\2024"; Subfolders = $true },
            @{ Path = "$root\見積"; Subfolders = $false },
            @{ Path = "$root\営業"; Subfolders = $true }
        )
        @(getRestrictedTsvFiles $indexes $excludes).Count | Should Be 0
        $direct = @(getRestrictedTsvFiles $indexes @(@{ Path = "$root\見積"; Subfolders = $false }) | ForEach-Object { $_.RelPath } | Sort-Object)
        $expected = @("見積\2024\b.docx\ページ001.tsv", "見積\2024\c.pptx\スライド001.tsv", "営業\深い\階層\d.xlsm\一覧.tsv") | Sort-Object
        ($direct -join "|") | Should Be ($expected -join "|")
    }

    It "大文字・小文字と末尾の \ の違いは同じフォルダとみなす" {
        $files = @(getRestrictedTsvFiles $indexes @(@{ Path = "$($root.ToUpperInvariant())\営業\深い\"; Subfolders = $true }))
        @($files | Where-Object { $_.RelPath -like "営業*" }).Count | Should Be 0
    }
}

Describe "readRestrictedTsvLines" -Tag Unit {
    It "StreamReader.ReadLine と同じく行に分ける（末尾の改行の後ろは行にしない。空なら 0 行）" {
        foreach ($text in @("", "a", "a`n", "`n", "a`n`nb", "a`nb`n")) {
            $reader = New-Object System.IO.StringReader($text)
            $expected = @()
            while ($null -ne ($line = $reader.ReadLine())) { $expected += $line }
            ((readRestrictedTsvLines $text) -join "|") + "#" + (readRestrictedTsvLines $text).Count | Should Be (($expected -join "|") + "#" + $expected.Count)
        }
    }
}

Describe "searchRestricted（時間切れからの照合し直し・1 行ずつの照合での打ち切り）" -Tag Io {
    $root = Join-Path $TestDrive "timeout"
    New-Item -ItemType Directory -Path "$root\見積\a.xlsx" -Force | Out-Null
    [System.IO.File]::WriteAllText("$root\見積\a.xlsx\売上.tsv", "b`r`n" + ("a" * 5000) + "`r`nb`r`n", (New-Object System.Text.UTF8Encoding($true)))
    $files = getRestrictedTsvFiles @(@{ Name = "見積"; Path = "$root\見積" })

    It "全文への照合が時間切れになったら、1 行ずつ照合し直す（同じ行を二重に数えない）" {
        # 全文では 1 行目の b が見つかった後、5000 文字の行で時間切れになる正規表現（1 行ずつの照合には別の正規表現を使う）
        Mock newSearchRegex {
            $pattern = "b"
            $slow = New-Object regex -ArgumentList @('(?:a|a)+c|b', 'None', [timespan]::FromMilliseconds(100))
            @{ Regex = (New-Object regex -ArgumentList @($pattern, 'None', [timespan]::FromSeconds(5))); SimpleMatch = $false; TextRegex = $slow; ScanMode = "lines" }
        }
        $result = searchRestricted "b" $files $false
        @($result.Hits | ForEach-Object { $_.LineNumber }) -join "," | Should Be "1,3"
    }

    It "1 行ずつ照合するときも、上限を超えたら打ち切る" {
        $result = searchRestricted "(?i)b" $files $false $false 1
        $result.Hits.Count | Should Be 1
        $result.Truncated | Should Be $true
    }
}
