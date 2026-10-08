# テストの共通の準備。各テストの先頭で dot-source する。
#   . "$PSScriptRoot\..\..\helpers\load.ps1"
#
# $here は tests フォルダを指す（テストの置き場所の深さに関わらず同じ）。

$here          = (Resolve-Path "$PSScriptRoot\..").Path
${scriptsDir}  = (Resolve-Path "$here\..\scripts").Path
${testDataDir} = "$here\testdata"

. "${scriptsDir}\tebunko\lib.ps1"
# ワークスペースは起動口（startGui・invokeIndexerMain）が initWorkspace を呼んで決める。
# テストも、既定のワークスペース（setting.config の workspaceFolder）に揃えるため、ここで呼ぶ
# （テストごとに変えるときは newTestWorkspace で $workspace を差し替える）
initWorkspace
. "$PSScriptRoot\tsv.ps1"
. "$PSScriptRoot\workspace.ps1"
. "$PSScriptRoot\cfb.ps1"
