# カバレッジの計測の対象にするファイルの一覧（tests/run.ps1 と tests/meta/structure.Tests.ps1 が使う）。
# 対象は判断層・状態層だけ。画面層は自動テストの対象外にする。
#   ・ui\ の下は、判断層（*_view.ps1）だけを対象にする（画面層のファイルを足しても、名前を足さずに分母から外れる）
#   ・起動口の gui.ps1 も外す

function getCoverageTargets {
    param (
        [string]$scriptsRoot
    )

    $uiPattern = [regex]::Escape([System.IO.Path]::DirectorySeparatorChar + "ui" + [System.IO.Path]::DirectorySeparatorChar)
    return @(Get-ChildItem -LiteralPath $scriptsRoot -Recurse -Filter "*.ps1" |
        Where-Object {
            $_.Name -ne "gui.ps1" -and
            ($_.FullName -notmatch $uiPattern -or $_.Name -match "_view\.ps1$")
        } |
        ForEach-Object { $_.FullName })
}
