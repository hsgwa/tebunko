# 高速検索の決まり（tebunko\search\search_gram.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"

function script:getGramTokens {
    # 本文から txt に書く語の一覧を作る（テスト用）
    param ([string]$text)
    $set = New-Object 'System.Collections.Generic.HashSet[uint32]'
    addTextGrams $set $text
    $values = New-Object 'uint32[]' $set.Count
    $set.CopyTo($values)
    [Array]::Sort($values)
    return , @((convertToGramText $values).Trim() -split " ")
}

Describe "convertToGramToken" -Tag Unit {
    It "2 文字を UTF-16LE の 4 バイトの 16 進にする" {
        convertToGramToken "ニタ" | Should Be "xcb30bf30"
        convertToGramToken "ab" | Should Be "x61006200"
    }
}

Describe "getSearchGrams" -Tag Unit {
    It "隣り合う 2 文字を語にし、小文字にそろえる" {
        (getSearchGrams "ニター") -join "," | Should Be "xcb30bf30,xbf30fc30"
        (getSearchGrams "AB") -join "," | Should Be ((getSearchGrams "ab") -join ",")
    }

    It "空白で区切り、1 文字の部分は語にしない" {
        (getSearchGrams "ab c") -join "," | Should Be "x61006200"
        (getSearchGrams "a b c").Count | Should Be 0
        (getSearchGrams "").Count | Should Be 0
    }

    It "同じ語は 1 回だけにする" {
        (getSearchGrams "ababab").Count | Should Be 2
    }

    It "多いときは先頭と末尾を残して均等に間引く" {
        $word = -join (0..40 | ForEach-Object { [char](0x4E00 + $_) })
        $grams = getSearchGrams $word
        $grams.Count | Should Be ${searchGramMax}
        $grams[0] | Should Be (convertToGramToken $word.Substring(0, 2))
        $grams[-1] | Should Be (convertToGramToken $word.Substring($word.Length - 2, 2))
    }
}

Describe "testFastSearchUsable" -Tag Unit {
    It "Windows Search が使え、正規表現がオフで、2 文字以上の部分があれば使える" {
        testFastSearchUsable $true $false "見積" | Should Be $true
    }

    It "正規表現・1 文字・Windows Search が使えないときは使えない" {
        testFastSearchUsable $true $true "見積" | Should Be $false
        testFastSearchUsable $true $false "見" | Should Be $false
        testFastSearchUsable $false $false "見積" | Should Be $false
    }
}

Describe "addTextGrams / convertToGramText" -Tag Unit {
    It "本文に含まれるワードの語は、すべて本文の語に入っている（語の途中・記号・英字の大小・奇数の長さ）" {
        $text = "No`t品名`r`n1`tモニター 27インチ`t`"2,100,000`"`r`nABC-1234型番`t東京都千代田区丸の内"
        $tokens = New-Object System.Collections.Generic.HashSet[string] (, [string[]](getGramTokens $text))
        foreach ($word in @("ニター", "インチ", "100,000", "abc-1234", "1234型", "代田区", "丸の内", "品名", "27イ")) {
            foreach ($gram in (getSearchGrams $word)) {
                $tokens.Contains($gram) | Should Be $true
            }
        }
    }

    It "本文に無い 2 文字は入らない" {
        $tokens = getGramTokens "abcd"
        $tokens -contains (convertToGramToken "ac") | Should Be $false
        $tokens.Count | Should Be 3
    }

    It "1 文字・空の本文" {
        $set = New-Object 'System.Collections.Generic.HashSet[uint32]'
        addTextGrams $set "a"
        addTextGrams $set ""
        $set.Count | Should Be 0
    }

    It "範囲を指定して一部だけを文字列にできる" {
        $values = [uint32[]]@(1, 2, 3)
        convertToGramText $values 1 1 | Should Be "x02000000 "
        convertToGramText $values 3 | Should Be ""
    }
}

Describe "getGramPartCount / getSystemIndexFileNames / testSystemIndexPath" -Tag Unit {
    It "上限を超える分だけ分ける" {
        getGramPartCount 0 100 | Should Be 1
        getGramPartCount 10 100 | Should Be 1
        getGramPartCount 11 100 | Should Be 2
    }

    It "分けないときは 1 つ、分けるときは番号を付ける" {
        (getSystemIndexFileNames 1) -join "," | Should Be ${systemIndexFileName}
        (getSystemIndexFileNames 2) -join "," | Should Be "システムインデックス_1.txt,システムインデックス_2.txt"
    }

    It "パスが 240 文字以上になるなら作らない" {
        testSystemIndexPath "C:\ws\system_index\a" | Should Be $true
        testSystemIndexPath ("C:\" + ("a" * 230)) | Should Be $false
    }
}

Describe "convertFolderRoot / convertItemUrl / convertToScopeUrl" -Tag Unit {
    It "system_index の中のパスを index の中のパスにする" {
        convertFolderRoot "C:\ws\system_index\営業\2024" "C:\ws\system_index" "C:\ws\index" | Should Be "C:\ws\index\営業\2024"
        convertFolderRoot "C:\WS\System_Index" "C:\ws\system_index\" "C:\ws\index" | Should Be "C:\ws\index"
        convertFolderRoot "C:\ws\system_index2\a" "C:\ws\system_index" "C:\ws\index" | Should Be $null
    }

    It "ItemUrl は file: を外して / を \ にするだけ（% は戻さない）" {
        convertItemUrl "file:C:/ws/system_index/タブ%09あり/システムインデックス.txt" | Should Be "C:\ws\system_index\タブ%09あり\システムインデックス.txt"
    }

    It "SCOPE の URL は ' を重ねる" {
        convertToScopeUrl "C:\ws\R&D's\" | Should Be "file:C:/ws/R&D''s"
    }
}

Describe "getRelativePath" -Tag Unit {
    It "root からの相対パス。外・root そのものは null" {
        getRelativePath "C:\ws\system_index\営業\a" "C:\ws\system_index\" | Should Be "営業\a"
        getRelativePath "C:\WS\SYSTEM_INDEX\営業" "C:\ws\system_index" | Should Be "営業"
        getRelativePath "C:\ws\system_index" "C:\ws\system_index" | Should Be $null
        getRelativePath "C:\ws\index\営業" "C:\ws\system_index" | Should Be $null
    }
}

Describe "問い合わせ" -Tag Unit {
    It "候補・分けた txt・反映の判定の問い合わせを組み立てる" {
        $sql = newSystemIndexQuery "C:\ws\system_index\営業" @("x61006200", "x62006300")
        $sql | Should Match "SCOPE='file:C:/ws/system_index/営業'"
        $sql | Should Match "System.FileName = 'システムインデックス.txt'"
        $sql | Should Match ([regex]::Escape("CONTAINS(System.Search.Contents, '`"x61006200`" AND `"x62006300`"')"))
        newSystemIndexSplitQuery "C:\ws\system_index" "x61006200" | Should Match ([regex]::Escape("LIKE 'システムインデックス[_]%'"))
        newSystemIndexStateQuery "C:\ws\system_index" | Should Match "System.Search.GatherTime, System.DateModified"
    }
}

Describe "testSystemIndexReflected" -Tag Unit {
    $ticks = [datetime]::new(2026, 9, 24, 1, 2, 3, 456, [DateTimeKind]::Utc).Ticks
    $truncated = [datetime]::new(2026, 9, 24, 1, 2, 3)

    It "本文を読み終え、更新日時（秒で切り捨て）が同じなら反映済み" {
        testSystemIndexReflected ([datetime]::Now) $truncated $ticks | Should Be $true
    }

    It "本文を読み終えていない・更新日時が違う・値が無いなら未反映" {
        testSystemIndexReflected ([System.DBNull]::Value) $truncated $ticks | Should Be $false
        testSystemIndexReflected ([datetime]::Now) $truncated.AddSeconds(-1) $ticks | Should Be $false
        testSystemIndexReflected ([datetime]::Now) $null $ticks | Should Be $false
        testSystemIndexReflected $null $truncated $ticks | Should Be $false
    }
}

Describe "testFolderInTarget" -Tag Unit {
    It "サブフォルダも含む対象は、中のフォルダも入る" {
        testFolderInTarget "C:\ws\index\営業\2024" "C:\ws\index\営業" $true | Should Be $true
        testFolderInTarget "C:\ws\index\営業" "C:\ws\index\営業\" $true | Should Be $true
        testFolderInTarget "C:\ws\index\営業2" "C:\ws\index\営業" $true | Should Be $false
    }

    It "直下だけの対象は、そのフォルダだけ" {
        testFolderInTarget "C:\ws\index\営業" "C:\ws\index\営業" $false | Should Be $true
        testFolderInTarget "C:\ws\index\営業\2024" "C:\ws\index\営業" $false | Should Be $false
    }
}

Describe "状態ファイルの行" -Tag Unit {
    It "読んで書き戻すと同じになる。形の違う行は無視する" {
        $lines = @(
            "対応済み`t営業`t",
            "対象外`t営業\長い`t",
            "反映待ち`t営業\2024\システムインデックス.txt`t639258025749778837",
            "反映待ち`t壊れた行`tabc",
            "不明`tx`t",
            ""
        )
        $state = convertFromSystemIndexState $lines
        $state.Covered.Contains("営業") | Should Be $true
        $state.Excluded.Contains("営業\長い") | Should Be $true
        $state.Pending["営業\2024\システムインデックス.txt"] | Should Be 639258025749778837
        $state.Pending.Count | Should Be 1
        (convertToSystemIndexState $state) -join "|" | Should Be ($lines[0..2] -join "|")
    }

    It "インデックス名は相対パスの先頭" {
        getIndexNameOfRelPath "営業\2024\a.txt" | Should Be "営業"
        getIndexNameOfRelPath "営業" | Should Be "営業"
    }
}
