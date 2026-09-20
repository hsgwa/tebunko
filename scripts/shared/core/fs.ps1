# ファイルの読み書き（行ファイル・原子的な書き込み・長いパス・排他）。

function readListFile {
    # 1行1件のファイルを読み込む（空行を除く）。ファイルが無ければ空配列
    param (
        [string]$path
    )

    if (!(Test-Path -LiteralPath $path)) {
        return @()
    }
    return @(Get-Content -LiteralPath $path -Encoding UTF8 | Where-Object { $_.Trim() -ne "" })
}

function writeListFile {
    param (
        [string]$path,
        [string[]]$lines
    )

    if ($null -eq $lines) {
        $lines = [string[]]@()
    }
    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
    [System.IO.File]::WriteAllLines($path, $lines, ${utf8Bom})
}

function formatFileTime {
    # 変換一覧に記録する日時の書式（秒まで）。更新の有無はこの文字列で比べる
    param (
        [datetime]$time
    )

    return $time.ToString("yyyy/MM/dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
}

function writeTextLinesAtomic {
    # 途中で中断してもファイルが壊れないよう、一時ファイルに書いてから置き換える
    param (
        [string]$path,
        [object[]]$lines
    )

    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
    $tmpPath = "${path}.tmp"
    [System.IO.File]::WriteAllLines($tmpPath, [string[]]@($lines), ${utf8Bom})
    if (Test-Path -LiteralPath $path) {
        # $null は空文字列として渡されて例外になるため、[NullString]::Value（バックアップを作らない）を渡す
        [System.IO.File]::Replace($tmpPath, $path, [NullString]::Value)
    } else {
        [System.IO.File]::Move($tmpPath, $path)
    }
}

function toSafeFileName {
    # ファイル名に使えない文字を全角に変換
    param (
        [string]$name
    )

    $name = $name.Replace(">", "＞")
    $name = $name.Replace("<", "＜")
    $name = $name.Replace("\", "￥")
    $name = $name.Replace("*", "＊")
    $name = $name.Replace('"', '”')
    $name = $name.Replace(":", "：")
    $name = $name.Replace("?", "？")
    $name = $name.Replace("|", "｜")
    $name = $name.Replace("/", "／")

    return $name
}


${maxFileNameLength} = 255  # Windows のファイル名（パスの区切りの間の1つ）の上限


function toLongPath {
    # パスの先頭に \\?\ を付け、260文字を超えるパスもファイル操作（System.IO・-LiteralPath）で扱えるようにする。
    # ネットワークのパス（\\server\share\…）は \\?\UNC\server\share\… にする。付いていればそのまま返す。
    # Join-Path は \\?\ 付きのパスを扱えないため、パスを組み立てた後、ファイル操作に渡す直前に使う
    param (
        [string]$path
    )

    if (!$path -or $path.StartsWith("\\?\")) {
        return $path
    }
    $path = $path.Replace("/", "\")
    if ($path.StartsWith("\\")) {
        return "\\?\UNC\" + $path.Substring(2)
    }
    return "\\?\" + $path
}

function copyFileShared {
    # 元のファイルを占有せずにコピーする（変換は、このコピーを開いて行う）。
    # File.Copy は元のファイルをほかのアプリの書き込みを拒否して開くため、コピーの間は利用者が上書き保存できず、
    # 利用者がファイルを開いて編集中だとコピーできないことがある。
    # ここでは読み取りだけで開き、ほかのアプリの読み書き・削除・名前変更を妨げない。
    # コピー先は通常の属性（読み取り専用などを付けない）で作り、既にあれば上書きする
    param (
        [string]$sourcePath,
        [string]$destPath
    )

    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $source = New-Object System.IO.FileStream((toLongPath $sourcePath), [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
    try {
        $dest = New-Object System.IO.FileStream((toLongPath $destPath), [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        try {
            $source.CopyTo($dest, 1MB)
        } finally {
            $dest.Dispose()
        }
    } finally {
        $source.Dispose()
    }
}

function fromLongPath {
    # toLongPath で付けた \\?\ を外し、通常のパスに戻す（Get-ChildItem の FullName 等から相対パスを求めるため）
    param (
        [string]$path
    )

    if ($path.StartsWith("\\?\UNC\")) {
        return "\\" + $path.Substring(8)
    }
    if ($path.StartsWith("\\?\")) {
        return $path.Substring(4)
    }
    return $path
}

function removeDirectoryRetry {
    # フォルダを中身ごと削除する。ウイルス対策ソフト・エクスプローラーが一時的に掴んでいることがあるため、少し待って数回試す
    param (
        [string]$path,
        [int]$tries = 3,
        [int]$waitMilliseconds = 200
    )

    # 中に長いパス（260文字超）のファイルがあっても削除できるよう \\?\ 付きで削除する
    $longPath = toLongPath $path
    for ($i = 1; $true; $i++) {
        if (![System.IO.Directory]::Exists($longPath)) {
            return
        }
        try {
            Remove-Item -LiteralPath $longPath -Recurse -Force
            return
        } catch {
            if ($i -ge $tries) {
                throw
            }
            Start-Sleep -Milliseconds $waitMilliseconds
        }
    }
}

function newAppMutex {
    # 同じツール（配置フォルダ）の処理を二重に動かさないための名前付きミューテックスを作り、@{ Mutex; Acquired } を返す。
    # Acquired が $false なら、ほかで実行中。プロセスが終われば解放されるため、強制終了されても残らない
    #   name: 処理の種類（"gui" = 画面、"convert" = 変換）
    param (
        [string]$name,
        [string]$dir = ${rootDir}
    )

    $md5 = New-Object System.Security.Cryptography.MD5CryptoServiceProvider
    try {
        $key = [BitConverter]::ToString($md5.ComputeHash([System.Text.Encoding]::UTF8.GetBytes(([string]$dir).ToLowerInvariant()))).Replace("-", "")
    } finally {
        $md5.Dispose()
    }
    $createdNew = $false
    $mutex = New-Object System.Threading.Mutex($true, "Local\win_grep_${name}_${key}", [ref]$createdNew)
    return @{ Mutex = $mutex; Acquired = $createdNew }
}
