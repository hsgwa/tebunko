# GitHub Actions の Windows の実行環境（使い捨て）に、スクリプトの規則だけを強制した AppLocker のポリシーを当てる（Issue #71 の作業 1 の試作）。
# 開発機では実行しない。WDAC のポリシーを当てられないときの代わりとして試す。
#
# スクリプトは %WINDIR% と %PROGRAMFILES% の下だけを許可する。管理者向けの既定の規則（すべて許可）は入れない（実行環境の利用者は管理者のため）。
# exe の規則は設定しない（止めない）。リポジトリ（D:\a\...）のスクリプトは CLM で動く。
$ErrorActionPreference = 'Stop'

$xml = @'
<AppLockerPolicy Version="1">
  <RuleCollection Type="Script" EnforcementMode="Enabled">
    <FilePathRule Id="1d6c2b3e-0000-4000-8000-000000000001" Name="Windows" Description="" UserOrGroupSid="S-1-1-0" Action="Allow">
      <Conditions><FilePathCondition Path="%WINDIR%\*" /></Conditions>
    </FilePathRule>
    <FilePathRule Id="1d6c2b3e-0000-4000-8000-000000000002" Name="Program Files" Description="" UserOrGroupSid="S-1-1-0" Action="Allow">
      <Conditions><FilePathCondition Path="%PROGRAMFILES%\*" /></Conditions>
    </FilePathRule>
  </RuleCollection>
</AppLockerPolicy>
'@
$path = Join-Path $env:RUNNER_TEMP 'applocker.xml'
Set-Content -LiteralPath $path -Value $xml -Encoding UTF8
Set-AppLockerPolicy -XmlPolicy $path

& sc.exe config AppIDSvc start= auto | Out-Host
Start-Service -Name AppIDSvc
# AppLocker のドライバー（appid.sys）を起こし、ポリシーを読み直させる
& appidtel.exe start | Out-Host
& gpupdate.exe /force | Out-Host
Get-Service -Name AppIDSvc | Format-List Name, Status, StartType | Out-String | Write-Host
Get-AppLockerPolicy -Effective -Xml | Write-Host
# 規則が効いているか（リポジトリのスクリプトは拒否、Windows の下は許可になるはず）
# Test-AppLockerPolicy は .psm1 を扱えない（UnsupportedFileTypeException）ので ps1 だけを調べる
try {
    Test-AppLockerPolicy -XmlPolicy $path -Path (Join-Path $env:GITHUB_WORKSPACE 'tools\clm_poc\probe.ps1') -User Everyone | Format-List | Out-String -Width 300 | Write-Host
} catch {
    Write-Warning ("Test-AppLockerPolicy: " + $_.Exception.Message)
}
