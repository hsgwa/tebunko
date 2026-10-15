# ［新規で開く］のためのコピー（テキスト・PDF など、「新規」の開き方が無い種類のファイル）の作成と片付け（状態層）。
# コピーは $workspace.TmpRoot\<PC の鍵>\open_copy\<12 桁の文字>\<元の名前>（コピー yyyyMMdd-HHmm）.<拡張子> に置く（getOpenCopyDir・getOpenCopyName）。
# 片付けは GUI の起動時にだけ行い、1 日より古いものを消す（開いているコピーは消せずに残る。閉じるときには消さない）。
# 作る・消すときは、置き場所がワークスペースの TmpRoot\<PC の鍵>\open_copy そのものであり、リンク（ジャンクション・シンボリックリンク）で
# 外へつながっていないことを確かめてから行う（testOpenCopyDir）。

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
    # copyDir が、open_copy の置き場所として使ってよいものか。コピーの作成・片付けは、これが真のときだけ行う。
    #   絶対パス（ドライブ付き・UNC）で、途中に .. や余計な区切りを含まず、tmpRoot\<PC の鍵>\open_copy と一致すること
    #   copyDir・その親（<PC の鍵>）・その親（tmpRoot）が、リンク（ジャンクション・シンボリックリンク。testPathIsLink）でないこと（まだ無ければ問題にしない）
    # 空・ドライブの直下・相対パス・別の場所の open_copy・リンクの先は $false
    param (
        [string]$copyDir,
        [string]$tmpRoot
    )

    if (!$copyDir -or !$copyDir.Trim() -or !$tmpRoot -or !$tmpRoot.Trim()) {
        return $false
    }
    $trimmed = $copyDir.TrimEnd([char]92, [char]47)
    if ($trimmed -notmatch '^([A-Za-z]:\\|\\\\[^\\])' -or $tmpRoot -notmatch '^([A-Za-z]:\\|\\\\[^\\])') {
        return $false
    }
    try {
        $full = [System.IO.Path]::GetFullPath($trimmed)
        $expected = [System.IO.Path]::GetFullPath((Join-Path (Join-Path $tmpRoot (getMachineKey)) ${openCopyDirName}))
    } catch {
        return $false
    }
    if (![string]::Equals($trimmed, $full, [System.StringComparison]::OrdinalIgnoreCase) -or
        ![string]::Equals($full, $expected.TrimEnd([char]92), [System.StringComparison]::OrdinalIgnoreCase)) {
        return $false
    }
    $keyDir = [System.IO.Path]::GetDirectoryName($full)
    foreach ($checked in @($full, $keyDir, [System.IO.Path]::GetDirectoryName($keyDir))) {
        if (testPathIsLink $checked) {
            return $false
        }
    }
    return $true
}

function getOpenCopyName {
    # 開くためのコピーの名前 <元の名前>（コピー yyyyMMdd-HHmm）.<拡張子>（コピーだと分かる名前にする）。
    # パスの区切りが紛れていても、ファイル名だけを使う。名前が空なら空文字。
    # 同じ名前のコピーが重ならないよう、置く場所は copySourceForOpen が 1 回ごとに別のフォルダにする
    param (
        [string]$book,
        [datetime]$now = (Get-Date)
    )

    $leaf = [System.IO.Path]::GetFileName(([string]$book).TrimEnd('.', ' ').Replace("/", "\"))
    if (!$leaf) {
        return ""
    }
    $extension = [System.IO.Path]::GetExtension($leaf)
    $baseName = $leaf.Substring(0, $leaf.Length - $extension.Length)
    $stamp = $now.ToString("yyyyMMdd-HHmm", [System.Globalization.CultureInfo]::InvariantCulture)
    return "${baseName}（コピー ${stamp}）${extension}"
}

function getOpenCopyBlocker {
    # コピーを作らずに断る理由（ファイルには触らない）。作れるなら空文字。
    #   NoName  … 名前が空
    #   TooLong … コピーのパスが長すぎる（260 文字の制限）
    # 名前に [ ] を含んでも断らない（コピーも開くのも、パスをそのまま（リテラルで）扱う）
    param (
        [string]$book,
        [string]$copyDir,
        [datetime]$now = (Get-Date)
    )

    $name = getOpenCopyName $book $now
    if (!$name) {
        return "NoName"
    }
    # <copyDir>\<12 桁>\<名前>
    if ($copyDir.TrimEnd([char]92).Length + 1 + 12 + 1 + $name.Length -ge 259) {
        return "TooLong"
    }
    return ""
}

function copySourceForOpen {
    # 元のファイルのコピーを copyDir の下に作る（元のファイルは変えない）。@{ Path; Reason } を返す。
    #   作れたとき: Path = コピーのパス、Reason = ""
    #   作れなかったとき: Path = ""、Reason = Place（置き場所の確かめで断った）・NoName／TooLong（getOpenCopyBlocker の結果）・Failed（作る・書くのに失敗した）
    # 同じ名前のコピーが重ならないよう、1 回ごとに別のフォルダ（copyDir\<12 桁の文字>）に置く。名前は getOpenCopyName。
    param (
        [string]$sourcePath,
        [string]$book,
        [string]$copyDir,
        [string]$tmpRoot,
        [datetime]$now = (Get-Date)
    )

    if (!(testOpenCopyDir $copyDir $tmpRoot)) {
        return @{ Path = ""; Reason = "Place" }
    }
    $blocker = getOpenCopyBlocker $book $copyDir $now
    if ($blocker) {
        return @{ Path = ""; Reason = $blocker }
    }
    $name = getOpenCopyName $book $now
    $dir = Join-Path $copyDir ([guid]::NewGuid().ToString("N").Substring(0, 12))
    $destPath = Join-Path $dir $name
    try {
        if (!(newDirectoryUnder $dir $copyDir)) {
            return @{ Path = ""; Reason = "Failed" }
        }
        copyFileShared $sourcePath $destPath
    } catch {
        return @{ Path = ""; Reason = "Failed" }
    }
    return @{ Path = $destPath; Reason = "" }
}

function getOpenCopyLatestTime {
    # 1 つ下のフォルダの「最後に使った日時」。フォルダの作成・更新と、中のファイルの更新のうち新しいもの
    # （編集して保存したコピーを、作ってから 1 日たったからといってすぐ消さないため）
    param (
        [System.IO.DirectoryInfo]$folder
    )

    $latest = $folder.CreationTime
    if ($folder.LastWriteTime -gt $latest) {
        $latest = $folder.LastWriteTime
    }
    foreach ($copy in @(Get-ChildItem -LiteralPath $folder.FullName -File -Force)) {
        if ($copy.LastWriteTime -gt $latest) {
            $latest = $copy.LastWriteTime
        }
    }
    return $latest
}

function clearOldOpenCopies {
    # copyDir の下の、最後に使ってから keepDays 日より古いコピー（1 つ下のフォルダごと）を消し、消したフォルダの数を返す。
    # 最後に使った日時は getOpenCopyLatestTime。現在より先の日時（時計のずれ）は新しい扱いになり、消さずに残る（消さない側に倒す）。
    # 消すのは open_copy の 1 つ下のフォルダの中のファイルだけ（testOpenCopyDir が偽なら何もしない。フォルダごとに消す直前にも確かめ、偽になっていればそこで止める）。
    # リンク（ジャンクション・シンボリックリンク）は辿らず、使用中で消せないコピーは残す。再帰の削除は使わない
    param (
        [string]$copyDir,
        [string]$tmpRoot,
        [datetime]$now = (Get-Date),
        [int]$keepDays = ${openCopyKeepDays}
    )

    if (!(testOpenCopyDir $copyDir $tmpRoot) -or !(Test-Path -LiteralPath $copyDir -PathType Container)) {
        return 0
    }
    $limit = $now.AddDays(-$keepDays)
    $prefix = $copyDir.TrimEnd([char]92) + "\"
    $cleared = 0
    foreach ($old in @(Get-ChildItem -LiteralPath $copyDir -Directory -Force)) {
        if (testLinkItem $old) {
            continue
        }
        if (!$old.FullName.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase) -or (getOpenCopyLatestTime $old) -ge $limit) {
            continue
        }
        # 確かめてから消すまでの間に置き場所がリンクに替わっていないか、フォルダごとに消す直前にもう一度確かめる
        if (!(testOpenCopyDir $copyDir $tmpRoot)) {
            break
        }
        # 消すフォルダ自身も、列挙してから消すまでの間にリンクに替わっていないかを確かめる（リンクなら辿らない）
        if (testPathIsLink $old.FullName) {
            continue
        }
        foreach ($copy in @(Get-ChildItem -LiteralPath $old.FullName -File -Force)) {
            if (testLinkItem $copy) {
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
