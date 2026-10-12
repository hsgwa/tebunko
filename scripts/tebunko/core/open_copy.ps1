# ［新規で開く］のためのコピー（テキスト・PDF など、「新規」の開き方が無い種類のファイル）の作成と片付け（状態層）。
# コピーは $workspace.TmpRoot\<PC の鍵>\open_copy\<12 桁の文字>\<元のファイル名> に置く（getOpenCopyDir）。
# 片付けは GUI の起動時にだけ行い、1 日より古いものを消す（開いているコピーは消せずに残る。閉じるときには消さない）。

# open_copy の置き場所の名前（片付けが、消してよい場所かを見るときにも使う）
${openCopyDirName} = "open_copy"
# 開くためのコピーを残す日数（これより古いものを起動時に消す）
${openCopyKeepDays} = 1

function getOpenCopyDir {
    # 開くためのコピーを置くフォルダ（$workspace.TmpRoot\<PC の鍵>\open_copy）を返す。文字列を組み立てるだけで、ファイルには触らない。
    # 取り込みの作業フォルダ（<PC の鍵>\<PID>）と同じ親の下に置くが、名前が数字でないため、取り込みの片付け（removeStaleProcessDirs）は触らない
    param (
        [Workspace]$workspace
    )

    return Join-Path $workspace.TmpRoot (Join-Path (getMachineKey) ${openCopyDirName})
}

function testOpenCopyDir {
    # copyDir が、open_copy の置き場所として使ってよい形か（空・ドライブの直下・名前が open_copy でないものは $false）。
    # コピーの作成・片付けは、これが真のときだけ行う
    param (
        [string]$copyDir
    )

    if (!$copyDir -or !$copyDir.Trim()) {
        return $false
    }
    $trimmed = $copyDir.TrimEnd([char]92, [char]47)
    if ($trimmed -eq "" -or $trimmed.EndsWith(":")) {
        return $false
    }
    return ([System.IO.Path]::GetFileName($trimmed) -ceq ${openCopyDirName})
}

function getOpenCopyName {
    # 開くためのコピーの名前（元のファイルの名前のまま。パスの区切りが紛れていても、ファイル名だけを使う）。
    # 同じ名前のコピーが重ならないよう、置く場所は copySourceForOpen が 1 回ごとに別のフォルダにする
    param (
        [string]$book
    )

    return [System.IO.Path]::GetFileName(([string]$book).TrimEnd('.', ' ').Replace("/", "\"))
}

function copySourceForOpen {
    # 元のファイルのコピーを copyDir の下に作り、そのパスを返す。作れなければ空文字を返す（元のファイルは変えない）。
    # 同じ名前のコピーが重ならないよう、1 回ごとに別のフォルダ（copyDir\<12 桁の文字>）に置く。名前は元のファイルのまま（getOpenCopyName）
    param (
        [string]$sourcePath,
        [string]$book,
        [string]$copyDir
    )

    if (!(testOpenCopyDir $copyDir)) {
        return ""
    }
    $name = getOpenCopyName $book
    if (!$name) {
        return ""
    }
    $dir = Join-Path $copyDir ([guid]::NewGuid().ToString("N").Substring(0, 12))
    $destPath = Join-Path $dir $name
    if ($destPath.Length -ge 259) {
        return ""
    }
    try {
        [System.IO.Directory]::CreateDirectory((toLongPath $dir)) | Out-Null
        copyFileShared $sourcePath $destPath
    } catch {
        return ""
    }
    return $destPath
}

function clearOldOpenCopies {
    # copyDir の下の、作ってから keepDays 日より古いコピー（1 つ下のフォルダごと）を消し、消したフォルダの数を返す。
    # 消すのは open_copy の 1 つ下のフォルダの中のファイルだけ（testOpenCopyDir が偽なら何もしない）。
    # リンク（ジャンクション・シンボリックリンク）は辿らず、使用中で消せないコピーは残す。再帰の削除は使わない
    param (
        [string]$copyDir,
        [datetime]$now = (Get-Date),
        [int]$keepDays = ${openCopyKeepDays}
    )

    if (!(testOpenCopyDir $copyDir) -or !(Test-Path -LiteralPath $copyDir -PathType Container)) {
        return 0
    }
    $limit = $now.AddDays(-$keepDays)
    $prefix = $copyDir.TrimEnd([char]92) + "\"
    $cleared = 0
    foreach ($old in @(Get-ChildItem -LiteralPath $copyDir -Directory -Force)) {
        if ($old.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
            continue
        }
        if (!$old.FullName.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase) -or $old.CreationTime -ge $limit) {
            continue
        }
        foreach ($copy in @(Get-ChildItem -LiteralPath $old.FullName -File -Force)) {
            if ($copy.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
                continue
            }
            try {
                [System.IO.File]::SetAttributes((toLongPath $copy.FullName), [System.IO.FileAttributes]::Normal)
                [System.IO.File]::Delete((toLongPath $copy.FullName))
            } catch {
                # 開いているコピーは残す（次の起動で消す）
            }
        }
        try {
            [System.IO.Directory]::Delete((toLongPath $old.FullName), $false)
            $cleared++
        } catch {
            # 中にコピーが残っている（開いている）・既に無い
        }
    }
    return $cleared
}
