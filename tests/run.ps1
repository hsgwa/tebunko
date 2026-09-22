# テストの実行（Pester 3.4）
#
#   .\tests\run.ps1              既定（Unit・Io・Meta。Office と Slow は除く）
#   .\tests\run.ps1 -Tag Unit    速い確認だけ
#   .\tests\run.ps1 -All         Office・Slow も含めて全部（Office が必要）
#   .\tests\run.ps1 -Ci          結果の XML とカバレッジ（Cobertura XML）を出し、カバレッジの下限も確かめる
#   .\tests\run.ps1 -Path .\tests\shared\core   指定したフォルダ・ファイルのテストだけ
#
# いずれも失敗したテストの数を終了コードにする（pre-commit フック・CI が見る）。
# 1 件も実行しなかったときも失敗にする（タグの打ち間違いや、タグの渡し方の誤りで何も確かめずに通るのを防ぐ）
param (
    [string[]]$Tag,
    [string[]]$ExcludeTag,
    [string[]]$Path,  # 実行するテストのフォルダ・ファイル（既定は tests 全体）
    [switch]$All,    # Office・Slow も実行する（Excel・Word・PowerPoint が必要）
    [switch]$Ci,     # 結果の XML とカバレッジを出し、失敗数で終了する
    [switch]$Quiet
)

$ErrorActionPreference = "Stop"

# powershell.exe -File で呼ぶと、-Tag Unit,Meta は配列にならず "Unit,Meta" という 1 つの文字列で渡る
# （pre-commit フックがこの呼び方）。どちらの呼び方でも同じになるよう、カンマで分ける
function splitTags([string[]]$tags) {
    @($tags | ForEach-Object { $_ -split "," } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}
$Tag = splitTags $Tag
$ExcludeTag = splitTags $ExcludeTag

# テストは Pester 3.4 の書き方。Pester 5 も入っている環境（GitHub Actions のランナーなど）では 5 が読み込まれて
# 全部失敗するため、3 系を明示して読み込む
Import-Module Pester -MaximumVersion 3.99.99

$testsDir = $PSScriptRoot
$rootDir  = Split-Path $testsDir -Parent
$outDir   = "$rootDir\work\test"

# 既定で外すタグ。Office は COM が要るもの、Slow は時間がかかるもの、Manual は手で確かめるもの
$defaultExclude = @("Office", "Slow", "Manual")
if (!$PSBoundParameters.ContainsKey("ExcludeTag")) {
    $ExcludeTag = if ($All) { @("Manual") } else { $defaultExclude }
}

$arguments = @{
    Script     = if ($Path) { $Path } else { $testsDir }
    PassThru   = $true
    ExcludeTag = $ExcludeTag
}
if ($Tag)   { $arguments.Tag = $Tag }
if ($Quiet) { $arguments.Quiet = $true }

if ($Ci) {
    [System.IO.Directory]::CreateDirectory($outDir) | Out-Null
    $arguments.OutputFile   = "$outDir\results.xml"
    $arguments.OutputFormat = "NUnitXml"
    # カバレッジの対象は判断層・状態層だけにする（画面層は自動テストの対象外）
    $arguments.CodeCoverage = @(Get-ChildItem "$rootDir\scripts" -Recurse -Filter "*.ps1" |
        Where-Object { $_.Name -notmatch "^(gui|shell|app_host)\.ps1$" -and $_.Name -notmatch "_tab\.ps1$" -and $_.Name -notmatch "_dialog\.ps1$" } |
        ForEach-Object { $_.FullName })
}

$result = Invoke-Pester @arguments

if ($result.TotalCount -eq 0) {
    Write-Host "実行したテストが 0 件でした（-Tag $($Tag -join ',') / -ExcludeTag $($ExcludeTag -join ',')）。タグの指定を確かめてください。" -ForegroundColor Red
    exit 1
}

# -Quiet のときは何も表示されないため、失敗したテストだけを出す
if ($Quiet -and $result.FailedCount -gt 0) {
    Write-Host "失敗したテスト（$($result.FailedCount) 件）:" -ForegroundColor Red
    foreach ($test in @($result.TestResult | Where-Object { $_.Result -eq "Failed" })) {
        Write-Host "  $(@($test.Describe, $test.Context, $test.Name | Where-Object { $_ }) -join ' / ')"
        Write-Host "    $($test.FailureMessage)"
    }
}

# Pester 3.4 のカバレッジ（コマンド単位）を、行単位の Cobertura XML にして書き出す（Codecov に送るため）。
# Pester 3.4 はカバレッジをファイルに出せないため、HitCommands / MissedCommands から組み立てる。
# 1 行に実行されたコマンドが 1 つでもあれば、その行は通ったものとする。
# ファイル名はリポジトリからの相対パス（/ 区切り）にする。絶対パスには利用者名が入るため書かない
function writeCobertura($coverage, [string]$path) {
    $invariant = [System.Globalization.CultureInfo]::InvariantCulture
    $rootPrefix = $rootDir.TrimEnd('\') + '\'
    $files = @{}
    foreach ($item in @(@($coverage.HitCommands | ForEach-Object { @{ Command = $_; Hit = 1 } }) +
                        @($coverage.MissedCommands | ForEach-Object { @{ Command = $_; Hit = 0 } }))) {
        $file = [string]$item.Command.File
        if ($file.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase)) { $file = $file.Substring($rootPrefix.Length) }
        $file = $file.Replace('\', '/')
        if (!$files.ContainsKey($file)) { $files[$file] = @{} }
        $line = [int]$item.Command.Line
        $lines = $files[$file]
        $lines[$line] = [Math]::Max([int]$lines[$line], $item.Hit)
    }
    $rate = { param($covered, $valid) if ($valid -gt 0) { ($covered / $valid).ToString("0.####", $invariant) } else { "0" } }

    $allValid = 0
    $allCovered = 0
    foreach ($lines in $files.Values) {
        $allValid += $lines.Count
        $allCovered += @($lines.Values | Where-Object { $_ -gt 0 }).Count
    }

    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Indent = $true
    $settings.Encoding = New-Object System.Text.UTF8Encoding $false
    $writer = [System.Xml.XmlWriter]::Create($path, $settings)
    try {
        $writer.WriteStartDocument()
        $writer.WriteStartElement("coverage")
        $writer.WriteAttributeString("line-rate", (& $rate $allCovered $allValid))
        $writer.WriteAttributeString("branch-rate", "0")
        $writer.WriteAttributeString("lines-covered", [string]$allCovered)
        $writer.WriteAttributeString("lines-valid", [string]$allValid)
        $writer.WriteAttributeString("branches-covered", "0")
        $writer.WriteAttributeString("branches-valid", "0")
        $writer.WriteAttributeString("complexity", "0")
        $writer.WriteAttributeString("version", "0")
        $writer.WriteAttributeString("timestamp", [string][DateTimeOffset]::UtcNow.ToUnixTimeSeconds())
        $writer.WriteStartElement("sources")
        $writer.WriteElementString("source", ".")
        $writer.WriteEndElement()
        $writer.WriteStartElement("packages")
        # フォルダを 1 つのパッケージにする
        foreach ($group in @($files.Keys | Sort-Object | Group-Object { $i = $_.LastIndexOf('/'); if ($i -lt 0) { "." } else { $_.Substring(0, $i) } })) {
            $packageValid = 0
            $packageCovered = 0
            foreach ($file in $group.Group) {
                $packageValid += $files[$file].Count
                $packageCovered += @($files[$file].Values | Where-Object { $_ -gt 0 }).Count
            }
            $writer.WriteStartElement("package")
            $writer.WriteAttributeString("name", $group.Name)
            $writer.WriteAttributeString("line-rate", (& $rate $packageCovered $packageValid))
            $writer.WriteAttributeString("branch-rate", "0")
            $writer.WriteAttributeString("complexity", "0")
            $writer.WriteStartElement("classes")
            foreach ($file in $group.Group) {
                $lines = $files[$file]
                $writer.WriteStartElement("class")
                $writer.WriteAttributeString("name", $file.Substring($file.LastIndexOf('/') + 1))
                $writer.WriteAttributeString("filename", $file)
                $writer.WriteAttributeString("line-rate", (& $rate @($lines.Values | Where-Object { $_ -gt 0 }).Count $lines.Count))
                $writer.WriteAttributeString("branch-rate", "0")
                $writer.WriteAttributeString("complexity", "0")
                $writer.WriteStartElement("methods")
                $writer.WriteEndElement()
                $writer.WriteStartElement("lines")
                foreach ($line in @($lines.Keys | Sort-Object)) {
                    $writer.WriteStartElement("line")
                    $writer.WriteAttributeString("number", [string]$line)
                    $writer.WriteAttributeString("hits", [string]$lines[$line])
                    $writer.WriteAttributeString("branch", "false")
                    $writer.WriteEndElement()
                }
                $writer.WriteEndElement()
                $writer.WriteEndElement()
            }
            $writer.WriteEndElement()
            $writer.WriteEndElement()
        }
        $writer.WriteEndElement()
        $writer.WriteEndElement()
        $writer.WriteEndDocument()
    } finally {
        $writer.Close()
    }
}

if ($Ci) {
    if ($result.CodeCoverage) {
        writeCobertura $result.CodeCoverage "$outDir\coverage.xml"
    }
    $covered = 0
    $total = 0
    if ($result.CodeCoverage) {
        $covered = @($result.CodeCoverage.HitCommands).Count
        $total   = $covered + @($result.CodeCoverage.MissedCommands).Count
    }
    $percent = if ($total -gt 0) { [Math]::Round(100.0 * $covered / $total, 1) } else { 0 }
    "テスト {0} 件中 {1} 件成功 / {2} 件失敗　カバレッジ {3}%（{4}/{5}）" -f `
        $result.TotalCount, $result.PassedCount, $result.FailedCount, $percent, $covered, $total | Write-Host

    # カバレッジの下限（tests\coverage.baseline。下回ったら失敗にする）
    $baselineFile = "$testsDir\coverage.baseline"
    if (Test-Path -LiteralPath $baselineFile) {
        $baseline = [double]((Get-Content -LiteralPath $baselineFile -Raw).Trim())
        if ($percent -lt $baseline) {
            Write-Host "カバレッジが基準 $baseline% を下回りました（$percent%）。" -ForegroundColor Red
            exit 1
        }
    }
    exit $result.FailedCount
}
exit $result.FailedCount