# ファイルの読み書き（行ファイル・原子的な書き込み・長いパス・排他）。

function readListFile {
    # 1行1件のファイルを読み込む（空行を除く）。ファイルが無ければ空配列。
    # 読めない（ほかが書き込み中など）ときは例外にする。呼び出し側の $ErrorActionPreference が既定の Continue でも
    # 空の一覧と取り違えないよう -ErrorAction Stop を付ける
    param (
        [string]$path
    )

    if (!(Test-Path -LiteralPath $path)) {
        return @()
    }
    return @(Get-Content -LiteralPath $path -Encoding UTF8 -ErrorAction Stop | Where-Object { $_.Trim() -ne "" })
}

function getPathLeaf {
    # パスの最後の部分（ファイル名・フォルダ名）を返す。[System.IO.Path]::GetFileName と同じく、
    # 最後の \ / : より後ろを返す（末尾が区切りなら空）。制限言語モードでは System.IO.Path を呼べないため文字列で求める
    param (
        [string]$path
    )

    return $path.Substring($path.LastIndexOfAny([char[]]"\/:") + 1)
}

function getPathParent {
    # パスの親（最後の \ / より前）を返す。区切りが無ければ空。
    # [System.IO.Path]::GetDirectoryName の代わり（制限言語モードで使う）。ドライブ直下・UNC の共有直下の扱いは
    # GetDirectoryName と違うため、ファイルやフォルダのパス（相対パスを含む）の親を求めるときだけ使う
    param (
        [string]$path
    )

    $i = $path.LastIndexOfAny([char[]]"\/")
    if ($i -lt 0) {
        return ""
    }
    return $path.Substring(0, $i)
}

function getPathStem {
    # ファイル名から拡張子を除いたもの（[System.IO.Path]::GetFileNameWithoutExtension と同じ。最後の . より前）
    param (
        [string]$path
    )

    $leaf = getPathLeaf $path
    $i = $leaf.LastIndexOf(".")
    if ($i -lt 0) {
        return $leaf
    }
    return $leaf.Substring(0, $i)
}

function writeUtf8NoBom {
    # 文字列を BOM なしの UTF-8 で書く。制限言語モードでは System.Text.Encoding を使えず、
    # Set-Content -Encoding UTF8 は BOM を付けるため、バイト列を組み立てて書く（設定ファイルのような小さなファイル用）
    param (
        [string]$path,
        [string]$text
    )

    if (${fullLanguage}) {
        [System.IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding($false)))
        return
    }
    $bytes = @(for ($i = 0; $i -lt $text.Length; $i++) {
            $code = [int]$text[$i]
            if ($code -ge 0xD800 -and $code -le 0xDBFF -and $i + 1 -lt $text.Length) {
                # サロゲートペア（𠮷 など）は 1 つの文字（4 バイト）にする
                $low = [int]$text[$i + 1]
                if ($low -ge 0xDC00 -and $low -le 0xDFFF) {
                    $code = 0x10000 + (($code - 0xD800) -shl 10) + ($low - 0xDC00)
                    $i++
                }
            }
            if ($code -lt 0x80) {
                $code
            } elseif ($code -lt 0x800) {
                0xC0 -bor ($code -shr 6); 0x80 -bor ($code -band 0x3F)
            } elseif ($code -lt 0x10000) {
                0xE0 -bor ($code -shr 12); 0x80 -bor (($code -shr 6) -band 0x3F); 0x80 -bor ($code -band 0x3F)
            } else {
                0xF0 -bor ($code -shr 18); 0x80 -bor (($code -shr 12) -band 0x3F); 0x80 -bor (($code -shr 6) -band 0x3F); 0x80 -bor ($code -band 0x3F)
            }
        })
    if ($bytes.Count -eq 0) {
        New-Item -ItemType File -Path $path -Force | Out-Null
        return
    }
    Set-Content -LiteralPath $path -Value ([byte[]]$bytes) -Encoding Byte
}

function writeListFile {
    param (
        [string]$path,
        [string[]]$lines
    )

    if ($null -eq $lines) {
        $lines = [string[]]@()
    }
    if (${fullLanguage}) {
        [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
        [System.IO.File]::WriteAllLines($path, $lines, ${utf8Bom})
        return
    }
    # 制限言語モード: Set-Content -Encoding UTF8 も BOM 付き・行ごとに CRLF で書く（WriteAllLines と同じ中身）
    New-Item -ItemType Directory -Path (getPathParent $path) -Force | Out-Null
    writeUtf8BomLines $path $lines
}

function writeUtf8BomLines {
    # 行の配列を BOM 付き UTF-8・行ごとに CRLF で書く（制限言語モード用。[System.IO.File]::WriteAllLines と同じ中身）。
    # 行が無いときも、WriteAllLines と同じく BOM だけのファイルにする
    param (
        [string]$path,
        [string[]]$lines
    )

    if ($lines.Count -eq 0) {
        Set-Content -LiteralPath $path -Value ([byte[]](0xEF, 0xBB, 0xBF)) -Encoding Byte
        return
    }
    Set-Content -LiteralPath $path -Value $lines -Encoding UTF8
}

function formatFileTime {
    # 取り込み一覧に記録する日時の書式（秒まで）。更新の有無はこの文字列で比べる
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

    $tmpPath = "${path}.tmp"
    if (${fullLanguage}) {
        [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
        [System.IO.File]::WriteAllLines($tmpPath, [string[]]@($lines), ${utf8Bom})
    } else {
        New-Item -ItemType Directory -Path (getPathParent $path) -Force | Out-Null
        writeUtf8BomLines $tmpPath ([string[]]@($lines))
    }
    # 書いた直後のファイルは、ウイルス対策ソフト等が一時的に掴んでいて置き換えられないことがあるため、少し待って数回試す
    for ($i = 1; $true; $i++) {
        try {
            if (!${fullLanguage}) {
                # 制限言語モードでは File.Replace を呼べないため、上書きの移動で置き換える
                Move-Item -LiteralPath $tmpPath -Destination $path -Force -ErrorAction Stop
            } elseif (Test-Path -LiteralPath $path) {
                # $null は空文字列として渡されて例外になるため、[NullString]::Value（バックアップを作らない）を渡す
                [System.IO.File]::Replace($tmpPath, $path, [NullString]::Value)
            } else {
                [System.IO.File]::Move($tmpPath, $path)
            }
            return
        } catch {
            if ($i -ge 5) {
                throw
            }
            Start-Sleep -Milliseconds 200
        }
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
    # 元のファイルを占有せずにコピーする（インデックス作成は、このコピーを開いて行う）。
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
        if (!(Test-Path -LiteralPath $longPath -PathType Container)) {
            return
        }
        try {
            # 呼び出し側の $ErrorActionPreference によらず、消せなければ catch に来るよう -ErrorAction Stop を付ける
            # （付けないと既定の Continue では例外にならず、試し直しも失敗の報告もせずに戻ってしまう）
            Remove-Item -LiteralPath $longPath -Recurse -Force -ErrorAction Stop
            return
        } catch {
            if ($i -ge $tries) {
                throw
            }
            Start-Sleep -Milliseconds $waitMilliseconds
        }
    }
}

function getFolderKey {
    # フォルダのパスから、名前付きミューテックス・イベントの名前に使う鍵（16 進 64 文字）を作る。大文字と小文字は区別しない。
    # 安全性のためではなく、パスを名前に使える長さと文字にするためのハッシュ。
    # FIPS 準拠の実装（SHA256CryptoServiceProvider）を使う。MD5 や SHA256Managed は、
    # FIPS モードを有効にした Windows では作るときに例外になり、起動できなくなるため
    param (
        [string]$dir
    )

    $sha = New-Object System.Security.Cryptography.SHA256CryptoServiceProvider
    try {
        return [BitConverter]::ToString($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($dir.ToLowerInvariant()))).Replace("-", "")
    } finally {
        $sha.Dispose()
    }
}

function newAppMutex {
    # 同じツール（配置フォルダ）の処理を二重に動かさないための名前付きミューテックスを作り、@{ Mutex; Acquired } を返す。
    # Acquired が $false なら、ほかで実行中。プロセスが終われば解放されるため、強制終了されても残らない
    #   name: 処理の種類（"gui" = 画面、"indexer" = インデックス作成）
    param (
        [string]$name,
        [string]$dir = ${rootDir}
    )

    $key = getFolderKey $dir
    $createdNew = $false
    $mutex = New-Object System.Threading.Mutex($true, "Local\${appId}_${name}_${key}", [ref]$createdNew)
    return @{ Mutex = $mutex; Acquired = $createdNew }
}
