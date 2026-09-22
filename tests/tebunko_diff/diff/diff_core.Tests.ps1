# 差分の中心（tebunko_diff\diff\diff_core.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\tebunko_diff\lib.ps1"

function getKeys {
    # 文字の配列を、左右で共通の番号にする
    param ([string[]]$left, [string[]]$right)

    $dict = New-Object 'System.Collections.Generic.Dictionary[string,int]'
    return @{ A = (getLineKeys $left $dict); B = (getLineKeys $right $dict) }
}

function getLcsLength {
    # 総当たり（動的計画法）の最長共通部分列の長さ（テストの答え合わせ用）
    param ([int[]]$a, [int[]]$b)

    $width = $b.Count + 1
    $table = New-Object int[] (($a.Count + 1) * $width)
    for ($i = 1; $i -le $a.Count; $i++) {
        for ($j = 1; $j -le $b.Count; $j++) {
            if ($a[$i - 1] -eq $b[$j - 1]) {
                $table[$i * $width + $j] = $table[($i - 1) * $width + $j - 1] + 1
            } else {
                $up = $table[($i - 1) * $width + $j]
                $left = $table[$i * $width + $j - 1]
                $table[$i * $width + $j] = [Math]::Max($up, $left)
            }
        }
    }
    return $table[$a.Count * $width + $b.Count]
}

function testValidMatch {
    # 対応が正しい形か（同じ番号どうし・両方で順番が増える・同じ相手を 2 回使わない）
    param ([int[]]$a, [int[]]$b, [int[]]$match)

    $last = -1
    for ($i = 0; $i -lt $a.Count; $i++) {
        if ($match[$i] -lt 0) { continue }
        if ($match[$i] -le $last) { return $false }
        if ($a[$i] -ne $b[$match[$i]]) { return $false }
        $last = $match[$i]
    }
    return $true
}

Describe "getLineMatches" -Tag Unit {
    It "同じ並びは全部対応する" {
        $k = getKeys @("a", "b", "c") @("a", "b", "c")
        (getLineMatches $k.A $k.B) -join "," | Should Be "0,1,2"
    }

    It "全部違えば対応しない" {
        $k = getKeys @("a", "b") @("x", "y", "z")
        (getLineMatches $k.A $k.B) -join "," | Should Be "-1,-1"
    }

    It "途中に挿入した行を見つける（前後の行は対応したまま）" {
        $k = getKeys @("a", "b", "c", "d") @("a", "b", "新", "c", "d")
        (getLineMatches $k.A $k.B) -join "," | Should Be "0,1,3,4"
    }

    It "削除した行を見つける" {
        $k = getKeys @("a", "b", "c", "d") @("a", "c", "d")
        (getLineMatches $k.A $k.B) -join "," | Should Be "0,-1,1,2"
    }

    It "空の行が何度も出てくる並びでも、行の挿入をずらさずに見つける" {
        $left = @("見出し", "", "", "明細1", "", "明細2", "", "合計")
        $right = @("見出し", "", "", "明細1", "", "追加", "", "明細2", "", "合計")
        $k = getKeys $left $right
        $match = getLineMatches $k.A $k.B
        testValidMatch $k.A $k.B $match | Should Be $true
        @($match | Where-Object { $_ -ge 0 }).Count | Should Be 8
    }

    It "片方が空" {
        $k = getKeys @() @("a")
        $match = getLineMatches $k.A $k.B
        $match.Count | Should Be 0
        $k = getKeys @("a") @()
        (getLineMatches $k.A $k.B) -join "," | Should Be "-1"
    }

    It "でたらめな並びでも、対応はいつも正しい形になる" {
        $random = New-Object System.Random 7
        for ($round = 0; $round -lt 60; $round++) {
            $a = [int[]]@(1..($random.Next(0, 14)) | ForEach-Object { $random.Next(0, 4) })
            $b = [int[]]@(1..($random.Next(0, 14)) | ForEach-Object { $random.Next(0, 4) })
            $match = getLineMatches $a $b
            testValidMatch $a $b $match | Should Be $true
        }
    }

    It "目印の無い区間（Myers）は最長の対応になる" {
        # 同じ番号が何度も出てくる並び（両方に 1 回ずつだけの番号が無い）は Myers だけで比べる
        $random = New-Object System.Random 11
        for ($round = 0; $round -lt 40; $round++) {
            $a = [int[]]@(1..($random.Next(1, 12)) | ForEach-Object { $random.Next(0, 2) })
            $b = [int[]]@(1..($random.Next(1, 12)) | ForEach-Object { $random.Next(0, 2) })
            $a = [int[]]($a + $a)
            $b = [int[]]($b + $b)
            $match = New-Object int[] $a.Count
            for ($i = 0; $i -lt $a.Count; $i++) { $match[$i] = -1 }
            matchMyers $a $b 0 $a.Count 0 $b.Count $match 1000
            testValidMatch $a $b $match | Should Be $true
            @($match | Where-Object { $_ -ge 0 }).Count | Should Be (getLcsLength $a $b)
        }
    }

    It "違いが上限を超える区間は、対応なし（まとめて置き換え）にする" {
        $a = [int[]]@(1, 2, 1, 2, 1, 2)
        $b = [int[]]@(2, 1, 2, 1, 2, 1)
        $match = New-Object int[] 6
        for ($i = 0; $i -lt 6; $i++) { $match[$i] = -1 }
        matchMyers $a $b 0 6 0 6 $match 1
        @($match | Where-Object { $_ -ge 0 }).Count | Should Be 0
    }

    It "空白の違いを無視・大文字と小文字を区別しない、で同じ番号になる" {
        $dict = New-Object 'System.Collections.Generic.Dictionary[string,int]'
        $a = getLineKeys @("A  b", "x　 y") $dict $true $false
        $b = getLineKeys @("a b", "X y ") $dict $true $false
        ($a -join ",") | Should Be ($b -join ",")
    }
}

Describe "似ている度合い" -Tag Unit {
    It "文字: 同じなら 1、共通する 2 文字の組が無ければ 0" {
        getTextSimilarity "支払は30日" "支払は30日" | Should Be 1
        getTextSimilarity "abc" "xyz" | Should Be 0
        (getTextSimilarity "支払は30日以内" "支払は45日以内") -gt 0.5 | Should Be $true
    }

    It "Excel の行: 値のあるセルのうち、同じ列で同じ値の割合" {
        getCellSimilarity "a`tb`tc" "a`tx`tc" | Should Be (2 / 3)
        getCellSimilarity "`t`t" "`t" | Should Be 1
    }
}

Describe "getAlignedPairs" -Tag Unit {
    $similarity = { param($x, $y) getTextSimilarity $x $y }

    function getPairText {
        param ([string[]]$left, [string[]]$right)

        $k = getKeys $left $right
        $match = getLineMatches $k.A $k.B
        $pairs = getAlignedPairs $match $right.Count $left $right $similarity
        return (@($pairs | ForEach-Object { "$($_[0]):$($_[1]):$(getPairKind $_[2])" }) -join " ")
    }

    It "似ている削除と追加を「変更」に組む" {
        getPairText @("a", "支払は30日以内", "c") @("a", "支払は45日以内", "c") | Should Be "0:0:same 1:1:change 2:2:same"
    }

    It "似ていない削除と追加は組まない" {
        getPairText @("a", "送料", "c") @("a", "保守費用", "c") | Should Be "0:0:same 1:-1:delete -1:1:insert 2:2:same"
    }

    It "数の合わない区間は、似ているものを組み、残りを追加・削除にする" {
        getPairText @("a", "第2条 支払は30日", "z") @("a", "追加の条文", "第2条 支払は45日", "z") | Should Be "0:0:same -1:1:insert 1:2:change 2:3:same"
    }
}

Describe "getCharSegments" -Tag Unit {
    It "違う部分だけを印にする" {
        $segs = getCharSegments "支払は30日以内" "支払は45日以内"
        (@($segs.Left | ForEach-Object { if ($_.Changed) { "[$($_.Text)]" } else { $_.Text } }) -join "") | Should Be "支払は[30]日以内"
        (@($segs.Right | ForEach-Object { if ($_.Changed) { "[$($_.Text)]" } else { $_.Text } }) -join "") | Should Be "支払は[45]日以内"
    }

    It "長すぎる行は、行全体を違う部分にする" {
        $long = "あ" * (${diffMaxCharLength} + 1)
        $segs = getCharSegments $long "い"
        @($segs.Left).Count | Should Be 1
        $segs.Left[0].Changed | Should Be $true
    }
}

Describe "foldSameRows" -Tag Unit {
    function newRows {
        param ([string]$kinds)

        $no = 0
        return @($kinds.ToCharArray() | ForEach-Object {
            $no++
            $row = [DiffRow]::new()
            $row.Kind = if ($_ -eq "c") { "change" } else { "same" }
            $row.LeftNo = [string]$no
            $row.RightNo = [string]$no
            $row
        })
    }

    It "変更の前後 2 行を残して、長く続く同じ行をたたむ" {
        $rows = newRows "ssssssssscsssssssss"
        $folded = foldSameRows $rows
        (@($folded | ForEach-Object { if ($_.Type -eq "Fold") { "F$($_.FoldCount)" } else { $_.Kind[0] } }) -join "") | Should Be "F7sscssF7"
        $folded[0].LeftText | Should Be "同じ行 7 行（行 1〜7）"
    }

    It "たたむほど長くない同じ行は、たたまない" {
        $rows = newRows "csssssc"
        @(foldSameRows $rows | Where-Object { $_.Type -eq "Fold" }).Count | Should Be 0
    }
}

Describe "差分の速さ" -Tag Slow {
    It "10,000 行・違い 100 か所のシートを 5 秒以内に比べる" {
        $left = [string[]]@(1..10000 | ForEach-Object { "行 $_`t値 $($_ % 97)" })
        $list = [System.Collections.Generic.List[string]]::new($left)
        for ($k = 0; $k -lt 100; $k++) { $list[$k * 97 + 5] = "変えた $k" }
        $list.Insert(5000, "挿入")
        $right = [string[]]$list.ToArray()
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        $result = compareOfficeUnits "Excel" ([ordered]@{ S = $left }) ([ordered]@{ S = $right })
        $watch.Stop()
        $result.Inserts | Should Be 101
        $result.Deletes | Should Be 100
        $watch.Elapsed.TotalSeconds -lt 5 | Should Be $true
    }
}


Describe "差分の中心（細かい場合）" -Tag Unit {
    It "Excel の同じ行は、似ている度合い 1" {
        getCellSimilarity "a`tb" "a`tb" | Should Be 1
    }

    It "左に 1 行多く挟まっているときは、その行を削除にして次の行と組む" {
        $left = [string[]]@("a", "無関係の行", "支払は30日以内", "z")
        $right = [string[]]@("a", "支払は45日以内", "z")
        $k = getKeys $left $right
        $match = getLineMatches $k.A $k.B
        $pairs = getAlignedPairs $match $right.Count $left $right { param($x, $y) getTextSimilarity $x $y }
        (@($pairs | ForEach-Object { getPairKind $_[2] }) -join ",") | Should Be "same,delete,change,same"
    }
}
