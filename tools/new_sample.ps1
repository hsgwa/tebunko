# 手での確認・スクリーンショット・性能の確認に使うサンプル（Excel・Word・PowerPoint）を作る。
# 架空の会社 (株)山田商事 の部署・年度のフォルダに、既定 40 件を作る。同じ引数なら同じ中身になる（乱数を使わない）。
# Excel・Word・PowerPoint（COM）が必要。作ったあと tests\testdata\scrub_personal.ps1 で作成者名などを取り除く。
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\new_sample.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\new_sample.ps1 -OutDir sample2 -Count 100
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\new_sample.ps1 -PlanOnly   # 作る一覧だけ返す（Office を使わない）
#
# -Count は数百件まで（上限 500）。7,000 件・16 万件規模の性能用データは、tebunko-perfdata の
# new_index.ps1・new_books.ps1 で作る（同じ目的を 2 か所に持たない）。
# テストデータ（tests\testdata\make_testdata.ps1）は、壊れたファイルや危険な名前まで含む別の目的のもの。
param (
    [string]$OutDir = "sample",
    [int]$Count = 40,
    [switch]$PlanOnly  # 作るファイルの相対パスの一覧だけを返す
)

$ErrorActionPreference = "Stop"

$maxCount = 500
if ($Count -lt 1 -or $Count -gt $maxCount) {
    throw "-Count は 1 から $maxCount までで指定してください（大量データは tebunko-perfdata を使います）: $Count"
}

$companyName = "(株)山田商事"
$departments = @("営業部", "総務部", "経理部", "開発部", "人事部")
$years = @("2023", "2024", "2025")
$topicsByType = @{
    ".xlsx" = @("売上集計", "予算管理", "在庫一覧", "勤務表", "顧客名簿", "経費精算")
    ".docx" = @("議事録", "企画書", "報告書", "規程", "手順書", "見積書")
    ".pptx" = @("会社説明", "月次報告", "新製品", "研修資料", "提案書", "計画概要")
}
$types = @(".xlsx", ".docx", ".pptx")

function getSamplePlan {
    # 作るファイルの相対パスを順に返す。種類・部署・題名・年度を決まった順に巡らせ、一巡したら _2 のように番号を付ける
    param (
        [int]$count
    )

    $plan = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $count; $i++) {
        $type = $types[$i % 3]
        $j = [Math]::Floor($i / 3)
        $dept = $departments[$j % $departments.Count]
        $k = [Math]::Floor($j / $departments.Count)
        $topics = $topicsByType[$type]
        $topic = $topics[$k % $topics.Count]
        $m = [Math]::Floor($k / $topics.Count)
        $year = $years[$m % $years.Count]
        $round = [Math]::Floor($m / $years.Count)
        $name = if ($round -gt 0) { "${topic}_$($round + 1)" } else { $topic }
        $plan.Add("$companyName\$dept\${year}年度\$name$type")
    }
    return , $plan.ToArray()
}

function getSampleText {
    # 相対パスから、ファイルの中に書く固定の文を作る
    param (
        [string]$relPath
    )

    $parts = $relPath.Split('\')
    $title = [System.IO.Path]::GetFileNameWithoutExtension($parts[-1])
    return @{
        Title = $title
        Dept  = $parts[1]
        Year  = $parts[2]
        Lines = @("$($parts[1])の$($parts[2])の$title", "担当: $($parts[1])", "更新: $($parts[2])")
    }
}

$plan = getSamplePlan $Count
if ($PlanOnly) {
    return $plan
}

$outFull = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine((Get-Location).Path, $OutDir))
if (Test-Path -LiteralPath $outFull) {
    if (@(Get-ChildItem -LiteralPath $outFull -Force).Count -gt 0) {
        throw "出力先が空ではありません。別のフォルダを指定するか、先に消してください: $outFull"
    }
}

$excel = $null
$word = $null
$ppt = $null
try {
    foreach ($rel in $plan) {
        $full = Join-Path $outFull $rel
        [void](New-Item -ItemType Directory -Force -Path (Split-Path $full -Parent))
        $text = getSampleText $rel
        $ext = [System.IO.Path]::GetExtension($rel)

        if ($ext -eq ".xlsx") {
            if ($null -eq $excel) {
                $excel = New-Object -ComObject Excel.Application
                $excel.Visible = $false
                $excel.DisplayAlerts = $false
            }
            $wb = $excel.Workbooks.Add()
            $ws = $wb.Worksheets.Item(1)
            $ws.Name = $text.Title
            $ws.Cells.Item(1, 1).Value2 = "項目"
            $ws.Cells.Item(1, 2).Value2 = "内容"
            for ($r = 0; $r -lt $text.Lines.Count; $r++) {
                $ws.Cells.Item($r + 2, 1).Value2 = "行$($r + 1)"
                $ws.Cells.Item($r + 2, 2).Value2 = $text.Lines[$r]
            }
            $wb.SaveAs([string]$full, 51)
            $wb.Close($false)
        }
        elseif ($ext -eq ".docx") {
            if ($null -eq $word) {
                $word = New-Object -ComObject Word.Application
                $word.Visible = $false
                $word.DisplayAlerts = 0
            }
            $doc = $word.Documents.Add()
            $doc.Content.Text = ($text.Lines -join "`r")
            $doc.SaveAs2([string]$full, 12)
            $doc.Close(0)
        }
        else {
            if ($null -eq $ppt) {
                $ppt = New-Object -ComObject PowerPoint.Application
            }
            $pres = $ppt.Presentations.Add(0)
            $slide = $pres.Slides.Add(1, 1)
            $slide.Shapes.Title.TextFrame.TextRange.Text = $text.Lines[0]
            $slide.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = ($text.Lines[1..2] -join "`r")
            $pres.SaveAs([string]$full, 24)
            $pres.Close()
        }
    }
}
finally {
    foreach ($app in @($excel, $word, $ppt)) {
        if ($null -ne $app) {
            try { $app.Quit() } catch { Write-Verbose $_.Exception.Message }
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
        }
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# 作成者名・保存先の絶対パスなど、実行した人の情報を取り除く
& "$PSScriptRoot\..\tests\testdata\scrub_personal.ps1" -Path $outFull

Write-Output "$($plan.Count) 件を作りました: $outFull"
