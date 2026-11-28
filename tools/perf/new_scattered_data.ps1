# 場所ごとに 1 件ずつ散らばる語の、性能の計測用のデータを作る（perf.yml の入力 scattered が 1 のとき。手元でも使える）。
# 取り込みの一時置き場と同じ形の TSV（tebunko-perfdata の new_index.ps1 と同じ形）と、その語だけの語の表（words.tsv と同じ列）を書く。
#
#   .\tools\perf\new_scattered_data.ps1 -Dest <TSV のフォルダ> -Words <語の表> [-Tool <測る tebunko>] [-Files 19] [-Places 500] [-Lines 1]
#   -Dest   … TSV を作るフォルダ。<Dest>\散在\資料NN.docx\<場所>.tsv を作る（場所は Word のページ「ページNNN」。ファイル名は toIndexFileName）
#   -Words  … 語の表を書くファイル（名前・語・正規表現・件数。1 行目は見出し）。-Dest の外に置く
#   -Tool   … TSV のファイル名（toIndexFileName）を借りる tebunko のフォルダ。既定はこのスクリプトのリポジトリ
#   -Files・-Places … 元のファイルの数と、ファイルごとの場所（ページ）の数
#   -Lines  … 場所ごとの、語を含む行の数（既定 1。ヒットは Files × Places × Lines 件。既定は 9,500 件で、検索の上限 1 万件より少ない）
# 1 つの TSV は、語を含まない行を前後に 1 行ずつ置き、その間に語を含む行を Lines 行置く。乱数は使わず、同じ引数からは同じ中身を作る。
# 書くのは -Dest の下の TSV と -Words のファイルだけで、ワークスペース・設定には触らない。
param (
    [Parameter(Mandatory = $true)][string]$Dest,
    [Parameter(Mandatory = $true)][string]$Words,
    [string]$Tool,
    [int]$Files = 19,
    [int]$Places = 500,
    [int]$Lines = 1
)

$ErrorActionPreference = "Stop"
if ($Files -lt 1 -or $Places -lt 1 -or $Lines -lt 1) { throw "-Files・-Places・-Lines は 1 以上にしてください。" }
# Windows PowerShell 5.1 は -File で起動すると param の既定値の $PSScriptRoot が空になるため、既定のフォルダはここで決める
if (!$Tool) { $Tool = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent }
. ([WildcardPattern]::Escape("$PSScriptRoot\perf_common.ps1".Replace('`', '``')))
# toIndexFileName は index_name.ps1 だけでは動かない（ファイル名の上限 maxFileNameLength を shared の text.ps1 が持つ）ため、lib.ps1 を読む
. ([WildcardPattern]::Escape((resolveTebunkoLib $Tool).Replace('`', '``')))

# 計測の語。perfdata の語（words.tsv）と重ならず、語を含まない行にも出ない文字列
$wordName = "散らばる語"
$word = "散在計測語"

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
# 相対パスは PowerShell の今の場所から決める（.NET の今のフォルダとは違うことがある）
$Dest = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Dest).TrimEnd("\")
$Words = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Words)
[void][System.IO.Directory]::CreateDirectory($Dest)
$root = Join-Path $Dest "散在"
$fileDigits = [Math]::Max(2, "$Files".Length)
$placeDigits = [Math]::Max(3, "$Places".Length)
for ($f = 1; $f -le $Files; $f++) {
    $book = "資料{0}.docx" -f $f.ToString("D$fileDigits")
    $bookDir = Join-Path $root $book
    [void][System.IO.Directory]::CreateDirectory($bookDir)
    for ($p = 1; $p -le $Places; $p++) {
        $page = $p.ToString("D$placeDigits")
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append("$f-$page`t前の行`r`n")
        for ($l = 1; $l -le $Lines; $l++) { [void]$sb.Append("$f-$page-$l`t$word`t本文`r`n") }
        [void]$sb.Append("$f-$page`t後の行`r`n")
        [System.IO.File]::WriteAllText((Join-Path $bookDir (toIndexFileName "ページ$page")), $sb.ToString(), $utf8Bom)
    }
}

$hits = $Files * $Places * $Lines
[void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($Words))
[System.IO.File]::WriteAllText($Words, "名前`t語`t正規表現`t件数`r`n$wordName`t$word`tfalse`t$hits`r`n", $utf8Bom)
Write-Information ("散らばる語のデータを作りました（ファイル {0:N0}・場所 {1:N0}・ヒット {2:N0} 件）。" -f $Files, ($Files * $Places), $hits) -InformationAction Continue
