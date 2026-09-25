# テストの共通の準備。各テストの先頭で dot-source する。
#   . "$PSScriptRoot\..\..\helpers\load.ps1"
#
# $here は tests フォルダを指す（テストの置き場所の深さに関わらず同じ）。

$here          = (Resolve-Path "$PSScriptRoot\..").Path
${scriptsDir}  = (Resolve-Path "$here\..\scripts").Path
${testDataDir} = "$here\testdata"

. "${scriptsDir}\tebunko_grep\lib.ps1"
. "$PSScriptRoot\tsv.ps1"
. "$PSScriptRoot\office.ps1"
