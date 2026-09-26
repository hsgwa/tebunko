# データ（設定ファイル・work）の置き場所の既定（どのツールからも使う）。
# ツールのフォルダに書き込めればそこに置く（以前の版と同じ）。書き込めないとき（Program Files・読み取り専用の共有フォルダに
# 置いたとき）だけ、利用者ごとの場所（%LOCALAPPDATA%\tebunko\<ツールのフォルダの鍵>）に置く。
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
    # 設定ファイルと既定の work を置くフォルダを返す。
    # 利用者ごとの場所の鍵は getFolderKey の先頭 16 文字（フォルダの名前にするため短くする。インデックスの TSV のパスが長くなりすぎないように）
    param (
        [string]$root = ${rootDir},
        [string]$fallbackBase = [System.Environment]::GetFolderPath("LocalApplicationData")
    )

    if (testWritableFolder $root) {
        return $root
    }
    return Join-Path $fallbackBase ("tebunko\" + (getFolderKey $root).Substring(0, 16))
}

${dataDir} = getDataDir
