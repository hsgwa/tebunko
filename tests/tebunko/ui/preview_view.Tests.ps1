# プレビューの判断（tebunko\ui\preview_view.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\preview_view.ps1"
}

Describe "getPreviewTopRow" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "選択行が上から 3 行目に来る"; hit = 50; expected = 48 }
        @{ name = "選択行の前が 2 行ちょうどなら、先頭の行が一番上"; hit = 2; expected = 0 }
        @{ name = "選択行の前が 2 行に満たなければ、先頭から"; hit = 1; expected = 0 }
        @{ name = "選択行が先頭なら、先頭から"; hit = 0; expected = 0 }
        @{ name = "選択行が表に無いときは、先頭から"; hit = -1; expected = 0 }
    ) {
        param ($name, $hit, $expected)
        getPreviewTopRow $hit | Should -Be $expected
    }
}

Describe "getPreviewFillWidth" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "枠が広ければ、行番号の列を除いた幅いっぱいにする"; viewport = 900; base = 300; current = 300; last = 300; expected = 856 }
        @{ name = "合わせたあとの幅が、広がった枠にも追いつく"; viewport = 1000; base = 300; current = 856; last = 856; expected = 956 }
        @{ name = "枠が狭まれば、初めの幅まで戻る"; viewport = 500; base = 300; current = 856; last = 856; expected = 456 }
        @{ name = "初めの幅のほうが広ければ、それより狭くしない"; viewport = 500; base = 640; current = 640; last = 640; expected = 640 }
        @{ name = "枠の幅がまだ分からなければ初めの幅"; viewport = 0; base = 300; current = 300; last = 300; expected = 300 }
        @{ name = "行番号を除くと最小幅に満たなければ初めの幅"; viewport = 50; base = 30; current = 30; last = 30; expected = 30 }
        @{ name = "手で変えた幅（前に合わせた幅と違う）は変えない"; viewport = 900; base = 300; current = 500; last = 856; expected = 500 }
    ) {
        param ($name, $viewport, $base, $current, $last, $expected)
        getPreviewFillWidth $viewport 44 $base $current $last | Should -Be $expected
    }
}

Describe "getPreviewColumnWindow" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "左端: 見えている列と右に 1 列"; widths = @(1..10 | ForEach-Object { 100.0 }); offset = 0; viewport = 250; margin = 1; first = 0; last = 3; left = 0; right = 600 }
        @{ name = "左端: 行番号の列の中までのスクロールは、左端と同じ"; widths = @(1..10 | ForEach-Object { 100.0 }); offset = 20; viewport = 250; margin = 1; first = 0; last = 3; left = 0; right = 600 }
        @{ name = "途中: 見えている列の左右に 1 列ずつ"; widths = @(1..10 | ForEach-Object { 100.0 }); offset = 444; viewport = 250; margin = 1; first = 3; last = 7; left = 300; right = 200 }
        @{ name = "右端: 右の空きは 0"; widths = @(1..10 | ForEach-Object { 100.0 }); offset = 944; viewport = 500; margin = 1; first = 8; last = 9; left = 800; right = 0 }
        @{ name = "見える幅より広い列が 1 つ"; widths = @(100.0, 1000.0, 100.0); offset = 144; viewport = 300; margin = 0; first = 1; last = 1; left = 100; right = 100 }
        @{ name = "列が 0 のときは空の範囲"; widths = @(); offset = 0; viewport = 300; margin = 3; first = 0; last = -1; left = 0; right = 0 }
        @{ name = "全部の列が見える幅に収まるときは、全部の列"; widths = @(100.0, 100.0); offset = 0; viewport = 900; margin = 3; first = 0; last = 1; left = 0; right = 0 }
        @{ name = "スクロールの位置が全部の列の幅より右でも、最後の列"; widths = @(1..10 | ForEach-Object { 100.0 }); offset = 5000; viewport = 300; margin = 1; first = 8; last = 9; left = 800; right = 0 }
        @{ name = "幅が違う列でも、境目の列を選ぶ（幅 10・500・10・500・10 の 2 列目の途中）"; widths = @(10.0, 500.0, 10.0, 500.0, 10.0); offset = 344; viewport = 200; margin = 0; first = 1; last = 1; left = 10; right = 520 }
        @{ name = "見える幅がまだ分からないとき（0）も、先頭の列から"; widths = @(1..10 | ForEach-Object { 100.0 }); offset = 0; viewport = 0; margin = 1; first = 0; last = 1; left = 0; right = 800 }
    ) {
        param ($name, $widths, $offset, $viewport, $margin, $first, $last, $left, $right)
        $prefix = New-Object 'double[]' ($widths.Count + 1)
        for ($i = 0; $i -lt $widths.Count; $i++) { $prefix[$i + 1] = $prefix[$i] + $widths[$i] }
        $window = getPreviewColumnWindow $prefix $offset $viewport 44 $margin
        $window.First | Should -Be $first
        $window.Last | Should -Be $last
        $window.Left | Should -Be $left
        $window.Right | Should -Be $right
    }
}

Describe "toStatusText" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "短い文字列はそのまま"; text = "見積書"; expected = "見積書" }
        @{ name = "改行とタブはスペースにする"; text = "A`r`nB`tC"; expected = "A B C" }
        @{ name = "40 文字を超えたら末尾を省略する"; text = ("あ" * 50); expected = (("あ" * 40) + "…") }
        @{ name = "ちょうど 40 文字なら省略しない"; text = ("あ" * 40); expected = ("あ" * 40) }
    ) {
        param ($name, $text, $expected)
        toStatusText $text | Should -Be $expected
    }
}

Describe "getPreviewMenuItems" -Tag Unit {
    It "元のファイルのこの場所を開くを太字で先頭に出し、区切りのあとにコピーを並べる" {
        $items = @(getPreviewMenuItems)
        ($items | ForEach-Object { $_.Id }) -join "/" | Should -Be "openHere/separator/copyCell/copyRow"
        ($items | ForEach-Object { $_.Header }) -join "/" | Should -Be "元のファイルのこの場所を開く//選んだセルをコピー/この行をコピー"
        @($items | Where-Object { $_.Bold } | ForEach-Object { $_.Id }) | Should -Be @("openHere")
        foreach ($item in $items) {
            $item.Header | Should -Not -Match "Ctrl"
        }
    }
}

Describe "getPreviewReadMore" -Tag Unit {
    It "<Name>" -TestCases @(
        @{ Name = "端から遠ければ頼まない"; Case = @{ First = 100; Last = 120; Count = 300; FirstLine = 50; LastLine = 349 }; Expected = "" }
        @{ Name = "上の端に近ければ、上を 200 行"; Case = @{ First = 20; Last = 40; Count = 300; FirstLine = 500; LastLine = 799 }; Expected = "up:300:200" }
        @{ Name = "上の端から 21 行なら、まだ頼まない"; Case = @{ First = 21; Last = 41; Count = 300; FirstLine = 500; LastLine = 799 }; Expected = "" }
        @{ Name = "下の端に近ければ、下を 200 行（最後の行の次から）"; Case = @{ First = 260; Last = 279; Count = 300; FirstLine = 500; LastLine = 799 }; Expected = "down:800:200" }
        @{ Name = "下の端から 21 行なら、まだ頼まない"; Case = @{ First = 200; Last = 278; Count = 300; FirstLine = 500; LastLine = 799 }; Expected = "" }
        @{ Name = "上も下も近ければ、両方"; Case = @{ First = 5; Last = 20; Count = 30; FirstLine = 500; LastLine = 529 }; Expected = "up:300:200,down:530:200" }
        @{ Name = "1 行目まで 50 行しかなければ、50 行だけ（頼むのは 1 行目から）"; Case = @{ First = 0; Last = 20; Count = 300; FirstLine = 51; LastLine = 350 }; Expected = "up:1:50" }
        @{ Name = "1 行目に届いていれば、上は頼まない"; Case = @{ First = 0; Last = 20; Count = 300; FirstLine = 1; LastLine = 300 }; Expected = "" }
        @{ Name = "上にもう行が無い（done）なら、上は頼まない"; Case = @{ First = 0; Last = 20; Count = 300; FirstLine = 500; LastLine = 799; TopDone = $true }; Expected = "" }
        @{ Name = "下にもう行が無い（終わりまで読んだ）なら、下は頼まない"; Case = @{ First = 280; Last = 299; Count = 300; FirstLine = 500; LastLine = 799; BottomDone = $true }; Expected = "" }
        @{ Name = "上を読んでいる途中なら、上は頼まない（下は頼める）"; Case = @{ First = 5; Last = 295; Count = 300; FirstLine = 500; LastLine = 799; UpLoading = $true }; Expected = "down:800:200" }
        @{ Name = "下を読んでいる途中なら、下は頼まない（上は頼める）"; Case = @{ First = 5; Last = 295; Count = 300; FirstLine = 500; LastLine = 799; DownLoading = $true }; Expected = "up:300:200" }
        @{ Name = "行が無ければ何も頼まない"; Case = @{ First = 0; Last = 0; Count = 0; FirstLine = 0; LastLine = 0 }; Expected = "" }
    ) {
        param ($Name, $Case, $Expected)
        $result = @(getPreviewReadMore $Case.First $Case.Last $Case.Count $Case.FirstLine $Case.LastLine ([bool]$Case.TopDone) ([bool]$Case.BottomDone) ([bool]$Case.UpLoading) ([bool]$Case.DownLoading))
        (@($result | ForEach-Object { "$($_.Side):$($_.FirstLine):$($_.Count)" }) -join ",") | Should -Be $Expected
    }

    It "近づく行数と 1 回の行数は引数で変えられる" {
        $result = @(getPreviewReadMore 3 10 100 1000 1099 $false $false $false $false 5 50)
        (@($result | ForEach-Object { "$($_.Side):$($_.FirstLine):$($_.Count)" }) -join ",") | Should -Be "up:950:50"
    }
}

Describe "getPreviewTrim" -Tag Unit {
    It "<Name>" -TestCases @(
        @{ Name = "上限以下なら捨てない"; First = 1000; Last = 1020; Count = 3000; Side = "none"; Dropped = 0 }
        @{ Name = "上限を超え、上から遠ければ上を 200 行捨てる"; First = 2500; Last = 2520; Count = 3200; Side = "top"; Dropped = 200 }
        @{ Name = "上限を超え、下から遠ければ下を 200 行捨てる"; First = 100; Last = 120; Count = 3200; Side = "bottom"; Dropped = 200 }
        @{ Name = "上と下が同じ遠さなら下を捨てる"; First = 1000; Last = 2000; Count = 3001; Side = "bottom"; Dropped = 200 }
        @{ Name = "遠い側が 200 行より少なければ、見えている行は捨てずにある分だけ"; First = 150; Last = 3199; Count = 3200; Side = "top"; Dropped = 150 }
        @{ Name = "捨てる行が無ければ捨てない"; First = 0; Last = 3199; Count = 3200; Side = "none"; Dropped = 0 }
    ) {
        param ($Name, $First, $Last, $Count, $Side, $Dropped)
        $result = getPreviewTrim $First $Last $Count
        $result.Side | Should -Be $Side
        $result.Count | Should -Be $Dropped
    }
}
