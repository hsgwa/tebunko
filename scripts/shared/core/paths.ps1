# フォルダ構成の基本（どのツールからも使う）。

# このファイルは scripts\shared\core\ に置く。3 つ上がリポジトリ直下になる
${rootDir} = (Resolve-Path "$PSScriptRoot\..\..\..").Path
${workDir} = "${rootDir}\work"

${utf8Bom} = New-Object System.Text.UTF8Encoding($true)