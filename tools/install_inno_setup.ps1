# Inno Setup 7 を持ち運び版でランナーの一時フォルダに入れる（.github\workflows\release.yml が使う）。
#
#   .\tools\install_inno_setup.ps1 -Url <URL> -Sha256 <SHA256> -Installer <保存先の .exe> -Dir <展開先>
#
# 公式のリリースから取り、SHA256 を確かめてから、レジストリに書かない持ち運び版で -Dir に入れる。
# .github\workflows\release.yml の run（shell: powershell）は ASCII だけで書く決まりのため、
# 日本語のメッセージが要るこの手順は tools\ の BOM 付き UTF-8 のスクリプトに置く。
param (
    [Parameter(Mandatory = $true)]
    [string]$Url,
    [Parameter(Mandatory = $true)]
    [string]$Sha256,
    [Parameter(Mandatory = $true)]
    [string]$Installer,
    [Parameter(Mandatory = $true)]
    [string]$Dir
)

$ErrorActionPreference = "Stop"

Invoke-WebRequest -Uri $Url -OutFile $Installer -UseBasicParsing
$hash = (Get-FileHash -LiteralPath $Installer -Algorithm SHA256).Hash
if ($hash -ne $Sha256) {
    throw "Inno Setup の SHA256 が違います: $hash"
}
$p = Start-Process -FilePath $Installer -ArgumentList "/CURRENTUSER /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /NOICONS /PORTABLE=1 /TASKS=`"`" /DIR=`"$Dir`"" -Wait -PassThru
if ($p.ExitCode -ne 0) {
    throw "Inno Setup を入れられませんでした（終了コード $($p.ExitCode)）"
}
