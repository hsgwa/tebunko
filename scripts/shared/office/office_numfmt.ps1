# Excel の表示形式（numFmt）で、セルの値を Excel が画面に出す文字にする（判断層）。制限言語モードで動く書き方だけで書く。
#
# Excel を使わずに .xlsx を読むとき（office_reader_clm.ps1）に使う。xlsx に入っているのは「値」（数値・シリアル値・
# 共有文字列の番号）だけで、画面に見えている文字は styles.xml の表示形式を当てて作る必要があるため。
#
# Excel と同じにできないもの（docs/05_制限モード.md の「取り込める値の違い」に書く）:
#   ・列幅が足りないときの ##### … 列幅は xlsx に入っているが、文字の幅は画面のフォントで決まるため再現しない
#   ・分数（# ?/?）… 近い分数に直す規則が Excel と違うことがある
#   ・和暦以外のカレンダー（[$-411] 以外の暦）
# どちらも検索に使う文字としては元の値が残るため、見つからなくなることはない。

# 組み込みの表示形式（numFmtId → 書式）。Excel が styles.xml に書かない番号の分
# （ECMA-376 Part 1 18.8.30。日本語版 Excel で使う 27〜36・50〜58 の和暦・漢数字は除く）
${builtinNumberFormats} = @{
    0  = "General";  1  = "0";       2  = "0.00";    3  = "#,##0";   4  = "#,##0.00"
    9  = "0%";       10 = "0.00%";   11 = "0.00E+00"; 12 = "# ?/?";  13 = "# ??/??"
    14 = "m/d/yyyy"; 15 = "d-mmm-yy"; 16 = "d-mmm";   17 = "mmm-yy";  18 = "h:mm AM/PM"
    19 = "h:mm:ss AM/PM"; 20 = "h:mm"; 21 = "h:mm:ss"; 22 = "m/d/yyyy h:mm"
    37 = "#,##0;-#,##0"; 38 = "#,##0;[Red]-#,##0"; 39 = "#,##0.00;-#,##0.00"; 40 = "#,##0.00;[Red]-#,##0.00"
    45 = "mm:ss";    46 = "[h]:mm:ss"; 47 = "mm:ss.0"; 48 = "##0.0E+0"; 49 = "@"
    # 日本語版の Excel が使う組み込みの書式（27〜36・50〜58）。ブックには書式が書かれず番号だけが入る
    27 = '[$-411]ge.m.d';  28 = '[$-411]ggge"年"m"月"d"日"'; 29 = '[$-411]ggge"年"m"月"d"日"'
    30 = "m/d/yy";        31 = 'yyyy"年"m"月"d"日"';        32 = 'h"時"mm"分"'
    33 = 'h"時"mm"分"ss"秒"'; 34 = 'yyyy"年"m"月"';          35 = 'm"月"d"日"'
    36 = '[$-411]ge.m.d';  50 = '[$-411]ge.m.d';            51 = '[$-411]ggge"年"m"月"d"日"'
    52 = 'yyyy"年"m"月"';   53 = 'm"月"d"日"';                54 = '[$-411]ggge"年"m"月"d"日"'
    55 = 'yyyy"年"m"月"';   56 = 'm"月"d"日"';                57 = '[$-411]ge.m.d'
    58 = '[$-411]ggge"年"m"月"d"日"'
}

# 曜日（aaa / aaaa）。[datetime]::DayOfWeek は 0 が日曜
${japaneseDayNames} = @("日", "月", "火", "水", "木", "金", "土")

# 和暦（ggge）。開始の日付の新しい順に見る
${japaneseEras} = @(
    @{ Name = "令和"; Short = "R"; Initial = "R"; Start = [datetime]"2019-05-01" },
    @{ Name = "平成"; Short = "H"; Initial = "H"; Start = [datetime]"1989-01-08" },
    @{ Name = "昭和"; Short = "S"; Initial = "S"; Start = [datetime]"1926-12-25" },
    @{ Name = "大正"; Short = "T"; Initial = "T"; Start = [datetime]"1912-07-30" },
    @{ Name = "明治"; Short = "M"; Initial = "M"; Start = [datetime]"1868-01-25" }
)

function getNumberFormatCode {
    # 表示形式の番号（numFmtId）と、ブックに書かれた表示形式（numFmtId → 書式）から、使う書式を返す
    param (
        [string]$id,
        $formats = @{}   # styles.xml の numFmts（番号（文字列）→ 書式）
    )

    if ($formats -and $formats.ContainsKey($id)) {
        return [string]$formats[$id]
    }
    if ($id -match '^\s*[0-9]{1,4}\s*$' -and ${builtinNumberFormats}.ContainsKey([int]$id)) {
        return [string]${builtinNumberFormats}[[int]$id]
    }
    return "General"
}

function splitNumberFormatSections {
    # 表示形式を ; で区切った節（正の数・負の数・ゼロ・文字列）に分ける。
    # 引用符の中とエスケープ（\;）の ; では区切らない
    param (
        [string]$format
    )

    $sections = @()
    $current = ""
    $inQuote = $false
    $chars = $format.ToCharArray()
    for ($i = 0; $i -lt $chars.Length; $i++) {
        $c = [string]$chars[$i]
        if ($c -eq '"') {
            $inQuote = -not $inQuote
            $current += $c
            continue
        }
        if (-not $inQuote -and $c -eq "\") {
            $current += $c
            if ($i + 1 -lt $chars.Length) {
                $i++
                $current += [string]$chars[$i]
            }
            continue
        }
        if (-not $inQuote -and $c -eq ";") {
            $sections += $current
            $current = ""
            continue
        }
        $current += $c
    }
    $sections += $current
    return , $sections
}

function testDateFormatSection {
    # その節が日付・時刻の書式か（引用符の中とエスケープした文字、色・条件の [ ] は見ない）
    param (
        [string]$section
    )

    $bare = removeFormatLiterals $section
    # y m d h s（年月日時分秒）・aaa（曜日）・g（和暦の元号）・[h] [m] [s]（経過時間）のどれかがあれば日付・時刻
    return ($bare -match '[ymdhsg]|\[h\]|\[m\]|\[s\]')
}

function removeFormatLiterals {
    # 書式から、引用符で囲んだ文字・\ でエスケープした文字・[ ] の指定（色・条件・ロケール・経過時間以外）・
    # _ の次の文字（幅合わせ）・* の次の文字（埋め）を取り除く。書式の種類を見分けるために使う
    param (
        [string]$format
    )

    $text = $format -replace '"[^"]*"', ""
    $text = $text -replace '\\.', ""
    $text = $text -replace '_.', ""
    $text = $text -replace '\*.', ""
    # [h] [m] [s]（経過時間）は残し、それ以外の [ ] は取り除く
    $text = $text -replace '\[(?![hms]\])[^\]]*\]', ""
    return $text
}

function newFormatParts {
    # 書式を、そのまま出す文字（リテラル）と、値を当てる部分に分ける。
    # 戻り値は @{ Pattern = リテラルを 1 文字の目印に置き換えた書式; Literals = 目印 → 文字 }
    # （目印は書式に出てこない私用領域の文字にする）
    param (
        [string]$format
    )

    $literals = @{}
    $pattern = ""
    $chars = $format.ToCharArray()
    $next = 0xE000
    for ($i = 0; $i -lt $chars.Length; $i++) {
        $c = [string]$chars[$i]
        $text = $null
        if ($c -eq '"') {
            $text = ""
            for ($i++; $i -lt $chars.Length -and [string]$chars[$i] -ne '"'; $i++) {
                $text += [string]$chars[$i]
            }
        } elseif ($c -eq "\") {
            if ($i + 1 -lt $chars.Length) {
                $i++
                $text = [string]$chars[$i]
            } else {
                $text = ""
            }
        } elseif ($c -eq "_") {
            # 幅合わせ（次の文字の幅の空白）。Excel は空白を出す
            if ($i + 1 -lt $chars.Length) { $i++ }
            $text = " "
        } elseif ($c -eq "*") {
            # 埋め文字（列幅まで繰り返す）。繰り返しは再現しないため、何も出さない
            if ($i + 1 -lt $chars.Length) { $i++ }
            $text = ""
        } elseif ($c -eq "[") {
            $inside = ""
            for ($i++; $i -lt $chars.Length -and [string]$chars[$i] -ne "]"; $i++) {
                $inside += [string]$chars[$i]
            }
            if ($inside -match '^[hms]+$') {
                # 経過時間（[h] [mm] [ss]）は書式として残す
                $pattern += "[" + $inside + "]"
                continue
            }
            if ($inside.StartsWith('$') -and $inside.Contains("-")) {
                # [$¥-411] のような通貨記号つきのロケール指定。記号だけを出す
                $text = $inside.Substring(1, $inside.IndexOf("-") - 1)
            } else {
                # 色（[Red]）・条件（[>100]）・ロケール（[$-411]）は出さない
                $text = ""
            }
        }
        if ($null -eq $text) {
            $pattern += $c
            continue
        }
        if ($text -eq "") {
            continue
        }
        $mark = [string][char]$next
        $next++
        $literals[$mark] = $text
        $pattern += $mark
    }
    return @{ Pattern = $pattern; Literals = $literals }
}

function applyFormatLiterals {
    # newFormatParts で置き換えた目印を、元の文字に戻す
    param (
        [string]$text,
        [hashtable]$literals
    )

    foreach ($mark in $literals.Keys) {
        $text = $text.Replace($mark, [string]$literals[$mark])
    }
    return $text
}

function getExcelDateTime {
    # Excel のシリアル値を日時にする。$null なら日付として出せない値
    #   1900 年方式: 1 = 1900/01/01。Excel にしかない 1900/02/29（60）に合わせて、60 以下は 1 日ずらす
    #   1904 年方式: 0 = 1904/01/01
    param (
        [double]$serial,
        [bool]$date1904 = $false
    )

    if ($serial -lt 0 -or $serial -ge 2958466) {
        return $null
    }
    # 秒より細かい端数は秒に丸める（Excel も秒までで表示する。計算の誤差で 1 秒ずれるのを防ぐ）
    $seconds = $serial * 86400
    $rest = $seconds % 1
    $seconds = $seconds - $rest + $(if ($rest -ge 0.5) { 1 } else { 0 })
    $serial = $seconds / 86400
    if ($date1904) {
        return ([datetime]"1904-01-01").AddSeconds($seconds)
    }
    if ($serial -lt 60) {
        # 1900/01/01 が 1。60 未満は 1 日多く足す（Excel にしかない 1900/02/29 の分）
        return ([datetime]"1899-12-31").AddSeconds($seconds)
    }
    # シリアル値 60 は Excel にしかない 1900/02/29。.NET には無い日付のため 1900/02/28 を返し、
    # 日にちだけ formatExcelDateTime が 29 に差し替える（testExcelFakeLeapDay）
    return ([datetime]"1899-12-30").AddSeconds($seconds)
}

function testExcelFakeLeapDay {
    # Excel にしかない 1900/02/29（1900 年方式のシリアル値 60）か
    param (
        [double]$serial,
        [bool]$date1904 = $false
    )

    return (-not $date1904 -and $serial -ge 60 -and $serial -lt 61)
}

function formatJapaneseEra {
    # 和暦（ggge / gge / ge）の元号と年を返す @{ Name; Short; Initial; Year }。明治より前は $null
    param (
        [datetime]$date
    )

    foreach ($era in ${japaneseEras}) {
        if ($date -ge $era.Start) {
            return @{ Name = $era.Name; Short = $era.Short; Initial = $era.Initial; Year = ($date.Year - $era.Start.Year + 1) }
        }
    }
    return $null
}

function formatExcelDateTime {
    # 日付・時刻の書式を当てる（節はリテラルを目印に置き換えたもの）
    param (
        [string]$pattern,
        [double]$serial,
        [bool]$date1904 = $false
    )

    $date = getExcelDateTime $serial $date1904
    if ($null -eq $date) {
        return $null
    }
    # 経過時間（[h] [m] [s]）は、日付ではなく期間として出す
    $elapsed = $null
    if ($pattern -match '\[(h+|m+|s+)\]') {
        $elapsed = $serial
    }
    $has12Hour = ($pattern -match 'AM/PM|A/P')
    if ($has12Hour) {
        # 午前・午後は先に決めてしまう（AM/PM の A・P・M を、日付の文字として読まないため）
        $marker = $(if ($date.Hour -lt 12) { "AM" } else { "PM" })
        $pattern = $pattern -replace "AM/PM", $marker
        $pattern = $pattern -replace "A/P", $marker.Substring(0, 1)
        $literalMark = [string][char]0xE0FF
        $pattern = $pattern.Replace($marker, $literalMark)
    }
    # Excel にしかない 1900/02/29 は、日にちだけ 29 として出す
    $fakeDay = $(if (testExcelFakeLeapDay $serial $date1904) { 29 } else { 0 })
    $out = ""
    $chars = $pattern.ToCharArray()
    $seenHour = $false
    for ($i = 0; $i -lt $chars.Length; $i++) {
        $c = ([string]$chars[$i])
        $lower = $c.ToLowerInvariant()
        if ($c -eq "[") {
            $inside = ""
            for ($i++; $i -lt $chars.Length -and [string]$chars[$i] -ne "]"; $i++) {
                $inside += [string]$chars[$i]
            }
            $out += formatElapsedPart $inside $serial
            $seenHour = ($inside -match '^h')
            continue
        }
        if ($lower -notmatch '[ymdhsage]') {
            $out += $c
            continue
        }
        # 同じ文字の並び（yyyy・mm など）をまとめて読む
        $run = $c
        while ($i + 1 -lt $chars.Length -and ([string]$chars[$i + 1]).ToLowerInvariant() -eq $lower) {
            $i++
            $run += [string]$chars[$i]
        }
        $count = $run.Length
        switch ($lower) {
            "y" {
                $out += $(if ($count -le 2) { $date.ToString("yy") } else { $date.ToString("yyyy") })
            }
            "d" {
                $day = $(if ($fakeDay -gt 0) { $fakeDay } else { $date.Day })
                if ($count -eq 1) { $out += [string]$day }
                elseif ($count -eq 2) { $out += (padNumberText ([string]$day) 2) }
                elseif ($count -eq 3) { $out += ${japaneseDayNames}[[int]$date.DayOfWeek] }
                else { $out += ${japaneseDayNames}[[int]$date.DayOfWeek] + "曜日" }
            }
            "h" {
                $seenHour = $true
                $hour = $date.Hour
                if ($has12Hour) {
                    $hour = $date.Hour % 12
                    if ($hour -eq 0) { $hour = 12 }
                }
                $out += $(if ($count -eq 1) { [string]$hour } else { padNumberText ([string]$hour) 2 })
            }
            "s" {
                $seenHour = $false
                $out += $(if ($count -eq 1) { [string]$date.Second } else { $date.ToString("ss") })
            }
            "m" {
                # 直前が時（h）なら分、そうでなければ月
                $nextIsSecond = testNextIsSecond $chars ($i + 1)
                if ($seenHour -or $nextIsSecond) {
                    $out += $(if ($count -eq 1) { [string]$date.Minute } else { $date.ToString("mm") })
                    $seenHour = $false
                } elseif ($count -eq 1) { $out += [string]$date.Month }
                elseif ($count -eq 2) { $out += $date.ToString("MM") }
                elseif ($count -eq 3) { $out += $date.ToString("MMM") }
                else { $out += $date.ToString("MMMM") }
            }
            "a" {
                # aaa（曜日）・aaaa（曜日＋曜日）。それ以外の a はそのまま
                if ($count -eq 3) { $out += ${japaneseDayNames}[[int]$date.DayOfWeek] }
                elseif ($count -ge 4) { $out += ${japaneseDayNames}[[int]$date.DayOfWeek] + "曜日" }
                else { $out += $run }
            }
            "g" {
                $era = formatJapaneseEra $date
                if ($null -eq $era) { $out += $run }
                elseif ($count -eq 1) { $out += $era.Initial }
                elseif ($count -eq 2) { $out += $era.Short }
                else { $out += $era.Name }
            }
            "e" {
                # 和暦の年（ggge の e）。元号が無ければ西暦の年
                $era = formatJapaneseEra $date
                $year = $(if ($null -eq $era) { $date.Year } else { $era.Year })
                $out += $(if ($count -ge 2) { padNumberText ([string]$year) 2 } else { [string]$year })
            }
        }
    }
    if ($has12Hour) {
        $out = $out.Replace([string][char]0xE0FF, $marker)
    }
    if ($null -ne $elapsed) {
        # 経過時間の書式では、年月日は出さない（Excel も同じ）
        return $out
    }
    return $out
}

function testNextIsSecond {
    # 書式の i 文字目から見て、次に来る書式の文字が秒（s）か（mm:ss の分の判定に使う）
    param (
        [char[]]$chars,
        [int]$index
    )

    for ($i = $index; $i -lt $chars.Length; $i++) {
        $c = ([string]$chars[$i]).ToLowerInvariant()
        if ($c -eq "s") { return $true }
        if ($c -match '[ymdhga]') { return $false }
    }
    return $false
}

function formatElapsedPart {
    # 経過時間（[h] [mm] [ss]）1 つ分を出す
    param (
        [string]$inside,
        [double]$serial
    )

    $count = $inside.Length
    $kind = $inside.Substring(0, 1).ToLowerInvariant()
    # 秒に直して切り捨てる（[Math] は制限言語モードで使えないため、余りを引く）
    $seconds = $serial * 86400
    $seconds = $seconds - ($seconds % 1)
    $value = switch ($kind) {
        "h" { ($seconds - ($seconds % 3600)) / 3600 }
        "m" { ($seconds - ($seconds % 60)) / 60 }
        default { $seconds }
    }
    return (padNumberText ([string][int64]$value) $count)
}

function padNumberText {
    # 数字の文字列を、けた数まで 0 で埋める
    param (
        [string]$text,
        [int]$width
    )

    $sign = ""
    if ($text.StartsWith("-")) {
        $sign = "-"
        $text = $text.Substring(1)
    }
    while ($text.Length -lt $width) {
        $text = "0" + $text
    }
    return $sign + $text
}

function toNetNumberFormat {
    # Excel の数値の書式（リテラルを目印に置き換えたもの）を .NET の書式にする。
    # Excel と .NET は #・0・,・. ・% ・E+ が同じ意味のため、置き換えるのは ? だけ
    param (
        [string]$pattern
    )

    # ?（けた合わせの空白）は .NET に無いため # と同じに扱う
    return ($pattern -replace '\?', "#")
}

function formatExcelNumber {
    # 数値の書式を当てる（節はリテラルを目印に置き換えたもの）
    param (
        [string]$pattern,
        [double]$value
    )

    if ($pattern -match '^([#0?]*)\s*([#0?]+)/([#0?]+|[0-9]+)$') {
        return (formatExcelFraction $value $Matches[1] $Matches[3])
    }
    if ($pattern -match '/') {
        # 上の形に当てはまらない分数は、小数のまま出す（Excel と違う。docs/05_制限モード.md に書く）
        return (formatGeneralNumber $value)
    }
    $format = toNetNumberFormat $pattern
    if ($format -eq "") {
        return ""
    }
    return $value.ToString($format)
}

function formatExcelFraction {
    # 分数の書式（# ?/? など）を当てる。分母は ? の数のけた数まで（3/8 のように分母が決まっている書式にも対応）
    param (
        [double]$value,
        [string]$wholePart,    # 整数の部分の書式（空なら帯分数にしない）
        [string]$denominator   # 分母の書式（?? なら 2 けたまで、8 なら 8 分の n）
    )

    $sign = $(if ($value -lt 0) { "-" } else { "" })
    if ($value -lt 0) { $value = -$value }
    $whole = 0
    if ($wholePart -ne "") {
        $whole = $value - ($value % 1)
        $value = $value - $whole
    }
    if ($denominator -match '^[0-9]+$') {
        $best = [int]$denominator
        $top = $value * $best
        $top = $top - ($top % 1) + $(if (($top % 1) -ge 0.5) { 1 } else { 0 })
    } else {
        # ? のけた数で表せる分母のうち、いちばん近いものを選ぶ
        $limit = 1
        for ($i = 0; $i -lt $denominator.Length; $i++) { $limit = $limit * 10 }
        $limit = $limit - 1
        $best = 1
        $top = 0
        $error = 1
        for ($den = 1; $den -le $limit; $den++) {
            $num = $value * $den
            $rounded = $num - ($num % 1) + $(if (($num % 1) -ge 0.5) { 1 } else { 0 })
            $diff = $value - ($rounded / $den)
            if ($diff -lt 0) { $diff = -$diff }
            if ($diff -lt $error - 1e-12) {
                $error = $diff
                $best = $den
                $top = $rounded
            }
        }
    }
    if ($top -eq 0) {
        # 分数にならない（ちょうど整数）
        return ($sign + [string][int64]$whole)
    }
    if ($top -eq $best) {
        $whole = $whole + 1
        return ($sign + [string][int64]$whole)
    }
    $fraction = "$([int64]$top)/$([int64]$best)"
    if ($wholePart -eq "") {
        return ($sign + $fraction)
    }
    if ($whole -eq 0) {
        # 整数の部分が 0 のときは、Excel は空白にして分数だけを出す
        return ($sign + " " + $fraction)
    }
    return ($sign + [string][int64]$whole + " " + $fraction)
}

function formatGeneralNumber {
    # 書式が「General」のときの出し方。Excel は 11 けたまで出し、それを超えると指数にする
    param (
        [double]$value
    )

    if ($value -eq 0) {
        return "0"
    }
    $absolute = $(if ($value -lt 0) { -$value } else { $value })
    if ($absolute -ge 1e11 -or $absolute -lt 1e-10) {
        # Excel は 1.23457E+19 のように、仮数を 5 けたまで出す
        $text = $value.ToString("0.#####E+00")
        return $text
    }
    # 11 けた（小数点・符号を除く）に収まるように丸める
    $text = $value.ToString("0.##########")
    $digits = ($text -replace '[^0-9]', "").TrimStart("0")
    if ($digits.Length -gt 11) {
        $text = $value.ToString("G11")
        if ($text.Contains("E")) {
            $text = $value.ToString("0.#####E+00")
        }
    }
    return $text
}

function formatExcelCellText {
    # セルの値（xlsx に入っている文字列）と表示形式から、Excel が画面に出す文字を作る。
    #   $value  : <v> の中身（数値・シリアル値）、または文字列そのもの
    #   $kind   : "number"（数値）・"text"（文字列）・"boolean"・"error"
    #   $format : 表示形式（getNumberFormatCode の結果）
    param (
        [string]$value,
        [string]$kind,
        [string]$format,
        [bool]$date1904 = $false
    )

    if ($kind -eq "error") {
        return $value
    }
    if ($kind -eq "boolean") {
        return $(if ($value -eq "1" -or $value -eq "TRUE") { "TRUE" } else { "FALSE" })
    }
    # splitNumberFormatSections は , を付けて配列で返すため、@() で包まずに受ける（包むと配列 1 個の配列になる）
    $sections = splitNumberFormatSections ([string]$format)
    if ($kind -eq "text") {
        # 4 つ目の節が文字列の書式。無ければそのまま出す
        if ($sections.Count -ge 4) {
            $parts = newFormatParts $sections[3]
            return (applyFormatLiterals ($parts.Pattern -replace '@', $value) $parts.Literals)
        }
        return $value
    }
    if ($value -eq "" -or $value -notmatch '^\s*[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)([eE][+-]?[0-9]+)?\s*$') {
        return $value  # 数値として読めないものは、そのまま出す
    }
    $number = [double]$value

    # 節を選ぶ（正の数・負の数・ゼロ）。負の数の節があれば、符号を落として当てる
    $section = $sections[0]
    if ($number -lt 0 -and $sections.Count -ge 2) {
        $section = $sections[1]
        $number = -$number
    } elseif ($number -eq 0 -and $sections.Count -ge 3) {
        $section = $sections[2]
    }
    if ($section.Trim() -eq "") {
        return ""  # 節が空（"#,##0;-#,##0;" のゼロなど）は何も出さない
    }
    if ($section -match '^\s*(General|標準)\s*$') {
        return (formatGeneralNumber $number)
    }

    $parts = newFormatParts $section
    if (testDateFormatSection $section) {
        $text = formatExcelDateTime $parts.Pattern $number $date1904
        if ($null -ne $text) {
            return (applyFormatLiterals $text $parts.Literals)
        }
        return (formatGeneralNumber $number)
    }
    return (applyFormatLiterals (formatExcelNumber $parts.Pattern $number) $parts.Literals)
}
