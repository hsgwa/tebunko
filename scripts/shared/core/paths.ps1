# フォルダ構成の基本（どのツールからも使う）。

# このファイルは scripts\shared\core\ に置く。3 つ上がリポジトリ直下になる
${rootDir} = (Resolve-Path "$PSScriptRoot\..\..\..").Path
# データ（設定ファイル・work）の置き場所は data_dir.ps1 で決める（ツールのフォルダに書き込めないことがあるため）

${utf8Bom} = New-Object System.Text.UTF8Encoding($true)