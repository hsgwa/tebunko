# 見本（tests/testdata/compat/）を変える・消す PR に ! が付いているかを確かめる（AGENTS.md「GitHub の運用」の ! の項）。
#
#   .\tools\check_compat_golden.ps1 -Title <PR のタイトル> -Base <base のコミット>
#
# tests/testdata/compat/ の下で A（追加）だけなら通す（見本は足すだけでよい）。M（変更）・D（消す）・R（改名）が
# 1 つでもあれば、タイトルに ! が無いと失敗する（! があれば通す。前の版の見本を消す・直してよいのは ! の PR だけ）。
# 判定は diff の行とタイトルを受け取る純粋な関数（testCompatGoldenProblem）に分け、git を呼ぶ所は薄くする。
# ! の読み取りは tools/commit_message_rules.ps1 と同じ（tools/check_commit_message.ps1 も使う。2 か所に書かない）。
#
# pwsh 7（CI の ubuntu。.github/workflows/title.yml の pr-title）でも Windows PowerShell 5.1 でも動くように書く
# （パスの区切りは git の出力のまま "/" で比べる・core.quotepath=off で日本語のパスを引用符付きにしない。
#   tools/check_markdown_links.ps1 と同じやり方）。
param (
    [Parameter(Mandatory = $true)]
    [string]$Title,
    [Parameter(Mandatory = $true)]
    [string]$Base,
    [string]$Root = (Join-Path $PSScriptRoot ".."),
    # テスト用。指定すると git diff を呼ばず、この行をそのまま使う（git --name-status の行。R は 3 列）
    [string[]]$DiffLines
)

$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "commit_message_rules.ps1")

${compatRoot} = "tests/testdata/compat/"

# git diff --name-status <base>...HEAD -- tests/testdata/compat/ の行を取る（pwsh 7・Linux でも動く書き方）
function getCompatDiffLines([string]$root, [string]$base) {
    $info = New-Object System.Diagnostics.ProcessStartInfo "git", "-c core.quotepath=off diff --name-status ${base}...HEAD -- ${compatRoot}"
    $info.WorkingDirectory = $root
    $info.UseShellExecute = $false
    $info.RedirectStandardOutput = $true
    $info.StandardOutputEncoding = New-Object System.Text.UTF8Encoding($false)
    $process = [System.Diagnostics.Process]::Start($info)
    $text = $process.StandardOutput.ReadToEnd()
    $process.WaitForExit()
    if ($process.ExitCode -ne 0) { throw "git diff が失敗しました" }
    return @($text -split "`r?`n" | Where-Object { $_ })
}

# name-status の 1 行から、見本（compatRoot の下）のパスを取り出す（A 以外は、改名の旧・新の両方を見る）
function getCompatGoldenPaths([string]$line) {
    $cols = $line -split "`t"
    $status = $cols[0]
    $paths = if ($status.StartsWith("R") -or $status.StartsWith("C")) { @($cols[1], $cols[2]) } else { @($cols[1]) }
    return @($paths | Where-Object { $_ -and $_.StartsWith(${compatRoot}) })
}

# diff の行とタイトルから、問題があれば理由を返す（問題が無ければ $null）。git を呼ばない純粋な関数
function testCompatGoldenProblem {
    param (
        [string[]]$DiffLines,
        [bool]$Breaking
    )

    if ($Breaking) { return $null }
    foreach ($line in $DiffLines) {
        if (!$line) { continue }
        $status = ($line -split "`t")[0]
        if ($status.StartsWith("A")) { continue }
        $paths = getCompatGoldenPaths $line
        if ($paths.Count -gt 0) {
            return "見本（${compatRoot}）を変える・消す PR はタイトルに ! が要ります: $status $($paths -join ', ')"
        }
    }
    return $null
}

$breaking = testBreakingTitle $Title
$lines = if ($PSBoundParameters.ContainsKey("DiffLines")) { $DiffLines } else { getCompatDiffLines $Root $Base }
$problem = testCompatGoldenProblem -DiffLines $lines -Breaking $breaking

if ($problem) {
    Write-Host $problem -ForegroundColor Red
    exit 1
}
Write-Host "見本（${compatRoot}）の検査: 問題なし"
exit 0
