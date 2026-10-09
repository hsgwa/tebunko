# データ（設定ファイル）の置き場所の既定（どのツールからも使う）。
# ツールのフォルダに書き込めればそこに置く。書き込めないとき（Program Files・読み取り専用の共有フォルダに置いたとき）は、
# ツールが渡した逃げ先に置く（逃げ先はツールが決める。tebunko は既定のワークスペース）。
# work の置き場所は、ツールの側で決める（tebunko は設定の workspaceFolder）。

function testWritableFolder {
    # フォルダにファイルを作れるか。試しに作ったファイルは閉じると消える（DeleteOnClose）。フォルダが無ければ $false
    param (
        [string]$dir
    )

    if ([string]::IsNullOrWhiteSpace($dir) -or -not [System.IO.Directory]::Exists($dir)) {
        return $false
    }
    $probe = Join-Path $dir (".tebunko_write_test_" + [guid]::NewGuid().ToString("N") + ".tmp")
    try {
        $stream = New-Object System.IO.FileStream($probe, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write,
            [System.IO.FileShare]::None, 1, [System.IO.FileOptions]::DeleteOnClose)
        $stream.Dispose()
        return $true
    } catch {
        return $false
    }
}

function getDataDir {
    # 設定ファイルを置くフォルダを返す。ツールのフォルダに書き込めればそれ、書き込めなければ渡された逃げ先。
    # 逃げ先の既定値は持たない（shared はツールの既定の置き場所を知らない）
    param (
        [string]$root = ${rootDir},
        [Parameter(Mandatory = $true)]
        [string]$fallbackDir
    )

    if (testWritableFolder $root) {
        return $root
    }
    return $fallbackDir
}
