# テストの共通の準備。各テストの先頭で dot-source する。
#   . "$PSScriptRoot\..\..\helpers\load.ps1"
#
# $here は tests フォルダを指す（テストの置き場所の深さに関わらず同じ）。

$here          = (Resolve-Path "$PSScriptRoot\..").Path
${scriptsDir}  = (Resolve-Path "$here\..\scripts").Path
${testDataDir} = "$here\testdata"

# 既定のワークスペースを、利用者の本物（%USERPROFILE%\Documents\tebunko_ws）にしない。lib.ps1 を読んだ時点で ${settingsFile} が
# 既定のワークスペースから決まるため、その前に環境変数 TEBUNKO_DEFAULT_WORKSPACE を入れる。tests\run.ps1 を通したときは入っている。
# 通さずに Invoke-Pester を直接流したときは、このプロセスだけの使い捨てを入れる
. "$here\..\tools\isolation\isolation_common.ps1"
if ([string]::IsNullOrWhiteSpace($env:TEBUNKO_DEFAULT_WORKSPACE)) {
    $env:TEBUNKO_DEFAULT_WORKSPACE = newIsolatedFolder "tebunko-test-ws"
    # 使い捨てのフォルダは、このプロセスが終わるときに消す（run.ps1 を通したときは run.ps1 が消す）
    $global:tebunkoTestIsolatedFolder = $env:TEBUNKO_DEFAULT_WORKSPACE
    [void](Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
        Remove-Item -LiteralPath $global:tebunkoTestIsolatedFolder -Recurse -Force -ErrorAction SilentlyContinue
    })
}

. "${scriptsDir}\tebunko\lib.ps1"
# ワークスペースは起動口（startGui・invokeIndexerMain）が initWorkspace を呼んで決める。
# テストも、既定のワークスペース（setting.config の workspaceFolder）に揃えるため、ここで呼ぶ
# （テストごとに変えるときは newTestWorkspace で $workspace を差し替える）
initWorkspace
# 決まったワークスペースが本物を指していたら、テストを始めずに止める（設定ファイルの workspaceFolder が本物を指す場合など）
assertNotRealWorkspace $workspace.Dir
. "$PSScriptRoot\tsv.ps1"
. "$PSScriptRoot\workspace.ps1"
. "$PSScriptRoot\cfb.ps1"
