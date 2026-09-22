# 制限モードの検索（状態層）。制限言語モードで動く書き方だけで書く。
#
# いつもの画面の検索（search\search_run.ps1 の getIndexTsvFiles・searchIndex）と同じ TSV を同じ順に照合し、同じヒットを返す。
# 画面の検索は FileStream・StreamReader・List・並列のスレッドを使うが、制限言語モードではどれも使えないため、
# TSV を Get-Content で丸ごと読み、同じ正規表現（newSearchRegex）を全文にかける。照合のしかた（getRegexScanMode の
# lines・filter・scan）と、行・行番号の数え方は searchTsvFiles と同じにする（tests\tebunko_grep\restricted\restricted_search.Tests.ps1 が突き合わせる）。

# 全文を読んで照合する TSV の大きさの上限（search_run.ps1 の searchWholeFileMax と同じ）。これより大きい TSV は 1 行ずつ読む
${restrictedWholeFileMax} = 64MB

# 検索結果の件数の上限（いつもの画面と同じ）
${restrictedSearchLimit} = 10000

function testRestrictedExcluded {
    # TSV（元のファイル）のあるフォルダ（folder。フルパス）が、検索から外したフォルダ（readSearchExcludes）に当たるか。
    # いつもの画面の検索対象ツリーと同じく、Subfolders が $true ならそのフォルダ以下すべて、
    # $false ならそのフォルダ直下のファイル（直下の元のファイル名のフォルダの中の TSV を含む）だけを外す
    param (
        [string]$folder,
        [object[]]$excludes
    )

    foreach ($exclude in $excludes) {
        $rest = getPathUnderFolder $folder $exclude.Path
        if ($null -eq $rest) {
            continue
        }
        if ($exclude.Subfolders -or $rest -eq "") {
            return $true
        }
    }
    return $false
}

function getRestrictedTsvFiles {
    # 検索する TSV を列挙し、元のファイル名（Book）・場所（Location）・元のファイルのあるフォルダ（RelDir）付きで、
    # パス順（Sort-Object と同じ。いつもの画面の getIndexTsvFiles と同じ順）に返す。
    #   indexes : 検索するインデックス（getSearchIndexes の要素。Path はインデックスのフォルダ）
    #   excludes: 検索から外すフォルダ（readSearchExcludes）
    # 各要素は @{ LongPath（\\?\ 付き）; Root（インデックスの置き場所 work\index）; RelPath; RelDir; FileName; Book; Location; Size }
    param (
        [object[]]$indexes,
        [object[]]$excludes = @()
    )

    $found = @{}
    foreach ($index in @($indexes | Where-Object { $_ })) {
        $dir = ([string]$index.Path).TrimEnd("\")
        $longDir = toLongPath $dir
        if (!(Test-Path -LiteralPath $longDir -PathType Container)) {
            continue
        }
        $root = getPathParent $dir
        $rootLength = $root.Length + 1
        foreach ($file in @(Get-ChildItem -LiteralPath $longDir -Recurse -File -Filter "*.tsv" -ErrorAction SilentlyContinue)) {
            $full = fromLongPath $file.FullName
            $relPath = $full.Substring($rootLength)
            $parts = splitIndexTsvPath $relPath
            $folder = if ($parts.RelDir) { "$root\$($parts.RelDir)" } else { $root }
            if ($excludes.Count -gt 0 -and (testRestrictedExcluded $folder $excludes)) {
                continue
            }
            $found[$full] = @{
                LongPath = $file.FullName
                Root     = $root
                RelPath  = $relPath
                RelDir   = [string]$parts.RelDir
                FileName = getPathLeaf $relPath
                Book     = [string]$parts.Book
                Location = [string]$parts.Place
                Size     = $file.Length
            }
        }
    }
    return @($found.Keys | Sort-Object | ForEach-Object { $found[$_] })
}

function newRestrictedHit {
    # ヒット 1 件（いつもの画面の searchTsvFiles のヒットと同じ項目）
    param (
        $file,
        [int]$lineNumber,
        [string]$line
    )

    return @{
        Root = $file.Root; RelPath = $file.RelPath; RelDir = $file.RelDir; FileName = $file.FileName
        Book = $file.Book; Location = $file.Location; LineNumber = $lineNumber; Line = $line
    }
}

function readRestrictedTsvLines {
    # 改行を LF にそろえた全文を、StreamReader.ReadLine と同じ区切り方で行に分ける（末尾の改行の後ろは行にしない。空なら 0 行）
    param (
        [string]$text
    )

    if ($text -eq "") {
        return , [string[]]@()
    }
    $lines = $text.Split([char]10)
    if ($text[$text.Length - 1] -eq [char]10) {
        $lines = @($lines | Select-Object -First ($lines.Count - 1))
    }
    return , [string[]]$lines
}

function searchRestricted {
    # TSV（getRestrictedTsvFiles）をワードで検索する。いつもの画面の searchIndex と同じ条件・同じ結果にする。
    #   simpleMatch    : $true なら文字どおり。$false なら正規表現として探し、正規表現として不正なら文字どおりに探す
    #   limit          : 件数の上限（0 は上限なし）。超えたら打ち切る
    #   fileFilter     : 対象ファイル（newFileFilter）
    #   includeShapes / includeComments: 図形・コメントの場所の TSV も検索する（newPlaceExclude）
    #   onProgress     : TSV を 100 件検索するたびに呼ぶ { param($done, $total) }
    # @{ Hits（ヒットのハッシュテーブルの配列）; SimpleMatch; Total（絞った後の TSV の数）; Truncated } を返す。
    # 1 行の照合に 5 秒を超えたら（newSearchRegex の regexTimeout）、searchIndex と同じメッセージで例外にする
    param (
        [string]$word,
        [object[]]$files,
        [bool]$simpleMatch = $false,
        [bool]$caseSensitive = $false,
        [int]$limit = 0,
        [string]$fileFilter = "",
        [bool]$includeShapes = $true,
        [bool]$includeComments = $true,
        [scriptblock]$onProgress = $null
    )

    $search = newSearchRegex $word $simpleMatch $caseSensitive
    $filter = newFileFilter $fileFilter
    $placeExclude = newPlaceExclude $includeShapes $includeComments
    $targets = @($files | Where-Object {
            (!$filter.Include -or $filter.Include.IsMatch($_.Book)) -and
            (!$filter.Exclude -or !$filter.Exclude.IsMatch($_.Book)) -and
            (!$placeExclude -or !$placeExclude.IsMatch($_.Location))
        })

    $regex = $search.Regex
    $textRegex = $search.TextRegex
    $scanMode = $search.ScanMode
    $rawCheck = $textRegex.ToString().IndexOfAny([char[]]"^$") -lt 0
    $lf = [char]10
    $timeoutMessage = "正規表現の照合に時間がかかりすぎるため、検索を中止しました。正規表現を見直してください。"

    # ヒットは件数分のハッシュテーブルに番号で入れる（配列に += で足すと、件数が多いと遅くなるため）
    $hits = @{}
    $count = 0
    $max = if ($limit -gt 0) { $limit } else { -1 }
    $truncated = $false
    $done = 0

    foreach ($f in $targets) {
        $done++
        if ($onProgress -and ($done % 100) -eq 0) {
            & $onProgress $done $targets.Count
        }
        # 読めない TSV（インデックス作成中に削除された等）は飛ばす
        try {
            if ($f.Size -gt ${restrictedWholeFileMax}) {
                $text = $null
                $lines = @(Get-Content -LiteralPath $f.LongPath -Encoding UTF8 -ErrorAction Stop)
            } else {
                $text = Get-Content -LiteralPath $f.LongPath -Raw -Encoding UTF8 -ErrorAction Stop
                if ($null -eq $text) {
                    $text = ""
                }
                $lines = $null
            }
        } catch {
            continue
        }

        if ($null -ne $text -and $scanMode -ne "scan") {
            $before = $count
            $scanned = $false
            try {
                # ^ $ を含まなければ、改行をそろえる前の全文で一致しない TSV は、そろえても一致しない
                $rawChecked = $false
                if ($rawCheck) {
                    if (!$textRegex.IsMatch($text)) { continue }
                    $rawChecked = $true
                }
                $text = $text.Replace("`r`n", "`n").Replace("`r", "`n")
                if ($scanMode -eq "filter") {
                    if (!$rawChecked -and !$textRegex.IsMatch($text)) { continue }
                } else {
                    $length = $text.Length
                    # 末尾の改行の後ろは行ではない
                    $tailIsLine = $length -gt 0 -and $text[$length - 1] -ne $lf
                    $number = 1
                    $pos = 0  # 行 number の先頭
                    $m = $textRegex.Match($text)
                    while ($m.Success) {
                        $index = $m.Index
                        if ($index -ge $length -and !$tailIsLine) { break }
                        $lineStart = if ($index -eq 0) { 0 } else { $text.LastIndexOf($lf, $index - 1) + 1 }
                        $lineEnd = $text.IndexOf($lf, $index)
                        if ($lineEnd -lt 0) { $lineEnd = $length }
                        if ($lineStart -gt $pos) {
                            $skipped = $text.Substring($pos, $lineStart - $pos)
                            $number += $skipped.Length - $skipped.Replace("`n", "").Length
                            $pos = $lineStart
                        }
                        $hits[$count] = newRestrictedHit $f $number $text.Substring($lineStart, $lineEnd - $lineStart)
                        $count++
                        if ($max -ge 0 -and $count -gt $max) { break }
                        if ($lineEnd -ge $length) { break }
                        # 1 行に複数一致しても 1 件にするため、次の行の先頭から探す
                        $m = $textRegex.Match($text, $lineEnd + 1)
                    }
                    $scanned = $true
                }
            } catch [System.Text.RegularExpressions.RegexMatchTimeoutException] {
                # 全文への照合が時間切れになったら、この TSV の分を捨てて 1 行ずつ照合し直す
                for ($i = $before; $i -lt $count; $i++) { $hits.Remove($i) }
                $count = $before
            }
            if ($max -ge 0 -and $count -gt $max) {
                $truncated = $true
                break
            }
            if ($scanned) {
                continue
            }
        }

        if ($null -eq $lines) {
            $lines = readRestrictedTsvLines $text.Replace("`r`n", "`n").Replace("`r", "`n")
        }
        $number = 0
        try {
            foreach ($line in $lines) {
                $number++
                if (!$regex.IsMatch($line)) { continue }
                $hits[$count] = newRestrictedHit $f $number $line
                $count++
                if ($max -ge 0 -and $count -gt $max) { break }
            }
        } catch [System.Text.RegularExpressions.RegexMatchTimeoutException] {
            throw $timeoutMessage
        }
        if ($max -ge 0 -and $count -gt $max) {
            $truncated = $true
            break
        }
    }

    if ($truncated) {
        $hits.Remove($count - 1)
        $count--
    }
    # if の値にすると、0 件の @() が $null になるため、先に空の配列を入れておく
    $list = @()
    if ($count -gt 0) {
        $list = @(for ($i = 0; $i -lt $count; $i++) { $hits[$i] })
    }
    return @{ Hits = $list; SimpleMatch = $search.SimpleMatch; Total = $targets.Count; Truncated = $truncated }
}
