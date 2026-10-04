# フォルダ構成の基本（どのツールからも使う）。

# このファイルは scripts\shared\core\ に置く。3 つ上がリポジトリ直下になる（zip 版）。
# 単一 .ps1 版（展開せずに動く試験版。設計は docs/design/structure/single-script.md）は、
# scripts\ の下の階層が無いため、${bundledScriptPath}（結合の道具が入れる、自分自身のパス）があれば
# その .ps1 自身の置き場所を根フォルダとする
${pathsBundled} = Get-Variable -Name bundledScriptPath -ErrorAction SilentlyContinue
${rootDir} = if ($null -ne ${pathsBundled} -and ${pathsBundled}.Value) {
    Split-Path -Parent ${pathsBundled}.Value
} else {
    (Resolve-Path "$PSScriptRoot\..\..\..").Path
}
Remove-Variable -Name pathsBundled -ErrorAction SilentlyContinue
# データ（設定ファイル・work）の置き場所は data_dir.ps1 で決める（ツールのフォルダに書き込めないことがあるため）

${utf8Bom} = New-Object System.Text.UTF8Encoding($true)