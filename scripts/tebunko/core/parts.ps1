# 部品（lib・indexerLib・indexer）の読み込み先をまとめる。
# zip 版は scripts 配下の実際のファイルのパスを返し、単一 .ps1 版（展開せずに動く試験版）は
# 自分自身のパス（$bundledScriptPath）と、どの部分を読み込むかを示す引数を返す。
# 設計は docs/design/structure/single-script.md。
#
# 単一 .ps1 版かどうかは、先頭で定義する変数 ${bundledScriptPath} の有無だけで決める
# （スクリプトの先頭で $PSCommandPath を入れておく。結合の道具 tools/new_single_script.ps1 が作る）。
# 別スレッド・別プロセス・別ランスペースで部品を読み込むところは、パスの文字列を直接組み立てず、必ずここを通す。

function getPartLoad {
    # 戻り値は @{ Path = <読み込む .ps1 のパス>; Args = <渡す引数（ハッシュテーブル）> }。
    # zip 版は Args が空（引数を渡さず、そのパスをそのまま読み込める）。
    # 単一 .ps1 版は Args に Part を入れる（自分自身を読み込み直すときに、どの部分かを伝える）
    param (
        [ValidateSet("lib", "indexerLib", "indexer")]
        [string]$name
    )

    $bundled = Get-Variable -Name bundledScriptPath -ErrorAction SilentlyContinue
    if ($null -ne $bundled -and $bundled.Value) {
        return @{ Path = $bundled.Value; Args = @{ Part = $name } }
    }

    $path = switch ($name) {
        "lib" { "$PSScriptRoot\..\lib.ps1" }
        "indexerLib" { "$PSScriptRoot\..\indexer\indexer_lib.ps1" }
        "indexer" { "$PSScriptRoot\..\indexer.ps1" }
    }
    return @{ Path = (Resolve-Path $path).Path; Args = @{} }
}

function getPartInitScript {
    # 文字列として組み立てて別スレッドに渡す dot-source（BackgroundQueue の初期化文字列）。
    # getPartLoad の戻り値を受け取り、". '<パス>' -Part '<名前>'" のような 1 行の文字列にする
    param (
        [hashtable]$load
    )

    $path = $load.Path.Replace("'", "''")
    $argsText = (@($load.Args.GetEnumerator() | ForEach-Object { "-$($_.Key) '$([string]$_.Value -replace "'", "''")'" })) -join " "
    if ($argsText) {
        return ". '$path' $argsText"
    }
    return ". '$path'"
}
