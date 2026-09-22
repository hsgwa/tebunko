# GitHub Actions の Windows の実行環境（使い捨て）に、スクリプトの制限を有効にした WDAC のポリシーを当てる（CI の clm.yml から呼ぶ）。
# 開発機では実行しない（ほかのアプリも止まる）。当てた後に起動した PowerShell は、信頼されていないスクリプトを制限言語モード（CLM）で動かす。
#
# ポリシー: Windows に付いている例（DefaultWindows_Enforced.xml。Windows が署名したものだけを許可する）を元にし、
# 実行環境が後の手順で起動する exe（Git・Node など）が止まらないよう、C:\ の決まったフォルダをパスの規則で許可する。
# リポジトリ（D:\a\...）は許可しないので、リポジトリのスクリプトは CLM で動く。
# なお、パスの規則で許可したフォルダのスクリプトは、windows-2022 では CLM のまま、windows-2025 では FullLanguage になった。
$ErrorActionPreference = 'Stop'

$dir = Join-Path $env:RUNNER_TEMP 'wdac'
New-Item -ItemType Directory -Force -Path $dir | Out-Null
$base = Join-Path $env:windir 'schemas\CodeIntegrity\ExamplePolicies\DefaultWindows_Enforced.xml'
if (!(Test-Path -LiteralPath $base)) { throw "例のポリシーが無い: $base" }
Write-Host "例のポリシー: $base"
Write-Host ("OS: " + (Get-CimInstance Win32_OperatingSystem).Caption + ' ' + [Environment]::OSVersion.Version)
Write-Host ("ワークスペース: " + $env:GITHUB_WORKSPACE)

# 実行環境の本体（Runner.Worker.exe）の場所。後の手順の node.exe などはこの下から起動される
$runner = (Get-Process -Name Runner.Worker | Select-Object -First 1).Path
$runnerRoot = Split-Path (Split-Path $runner -Parent) -Parent
Write-Host "実行環境: $runnerRoot"

$allow = @(
    'C:\Program Files\*',
    'C:\Program Files (x86)\*',
    'C:\ProgramData\*',
    'C:\hostedtoolcache\*',
    'C:\tools\*',
    "$runnerRoot\*"
) | Select-Object -Unique
if ($env:GITHUB_WORKSPACE -like 'C:\*') {
    Write-Warning "ワークスペースが C:\ にある。パスの規則でスクリプトまで許可されないかを結果で確かめること"
}

$rules = @()
foreach ($p in $allow) {
    Write-Host "パスで許可: $p"
    $rules += New-CIPolicyRule -FilePathRule $p
}
$merged = Join-Path $dir 'policy.xml'
# Windows Server 2025 では例のポリシーを直接読めない（アクセス拒否）ため、先に写す
$copy = Join-Path $dir 'base.xml'
Copy-Item -LiteralPath $base -Destination $copy
Merge-CIPolicy -PolicyPaths $copy -OutputFilePath $merged -Rules $rules | Out-Null

# 3 = 監査のみ（外して強制にする）、11 = スクリプトの制限を無効（外す）、
# 18 = パスの規則で、利用者が書き込めるフォルダを拒む確認をしない（実行環境のフォルダを許可するため）
Set-RuleOption -FilePath $merged -Option 3 -Delete
Set-RuleOption -FilePath $merged -Option 11 -Delete
Set-RuleOption -FilePath $merged -Option 18
$id = Set-CIPolicyIdInfo -FilePath $merged -PolicyName 'tebunko CLM PoC' -ResetPolicyID
$id = ([xml](Get-Content -LiteralPath $merged -Raw)).SiPolicy.PolicyID
Write-Host "PolicyID: $id"

$cip = Join-Path $dir "$id.cip"
ConvertFrom-CIPolicy -XmlFilePath $merged -BinaryFilePath $cip | Out-Null

$citool = Join-Path $env:windir 'System32\CiTool.exe'
if (Test-Path -LiteralPath $citool) {
    Write-Host 'CiTool で当てる（再起動なし）'
    & $citool --update-policy $cip -json
    if ($LASTEXITCODE -ne 0) { throw "CiTool の終了コード $LASTEXITCODE" }
} else {
    Write-Host 'CiTool が無いので、Active に置いて WMI で当てる'
    $active = Join-Path $env:windir 'System32\CodeIntegrity\CiPolicies\Active'
    New-Item -ItemType Directory -Force -Path $active | Out-Null
    Copy-Item -LiteralPath $cip -Destination $active
    Invoke-CimMethod -Namespace root\Microsoft\Windows\CI -ClassName PS_UpdateAndCompareCIPolicy -MethodName Update -Arguments @{ FilePath = (Join-Path $active "$id.cip") } | Format-List | Out-String | Write-Host
}

$dg = Get-CimInstance -Namespace root\Microsoft\Windows\DeviceGuard -ClassName Win32_DeviceGuard
Write-Host ("UsermodeCodeIntegrityPolicyEnforcementStatus: " + $dg.UsermodeCodeIntegrityPolicyEnforcementStatus + '（2 = 強制、1 = 監査のみ）')
Write-Host ("CodeIntegrityPolicyEnforcementStatus: " + $dg.CodeIntegrityPolicyEnforcementStatus)
