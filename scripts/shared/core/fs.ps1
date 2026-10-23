# ファイルの読み書き（行ファイル・原子的な書き込み・長いパス・排他）。


# ---- パスの状態（見つかった・見つからない・接続できない・その他） ----
#
# Test-Path・File.Exists・Directory.Exists は、届かない共有でも多くの場合に例外を出さず $false を返す
# （「見つからない」と区別できない）。getPathState は、例外を投げる呼び出し（File.GetAttributes）で調べる。

${pathStateFound}       = "Found"
${pathStateMissing}     = "Missing"
${pathStateUnreachable} = "Unreachable"
${pathStateOther}       = "Other"

# 接続できない（届かない共有・ネットワークの問題）を表す Win32 のエラー番号
# ERROR_BAD_NETPATH(53)・ERROR_BAD_NET_NAME(67)・ERROR_NETNAME_DELETED(64)・ERROR_SEM_TIMEOUT(121)・
# ERROR_NETWORK_UNREACHABLE(1231)・ERROR_HOST_UNREACHABLE(1232)・ERROR_NO_NET_OR_BAD_PATH(1222)
${unreachableWin32Codes} = @(53, 67, 64, 121, 1231, 1232, 1222)

# ---- クラウドにだけあるファイル（OneDrive・SharePoint などのオンデマンド） ----
#
# 一覧（Get-ChildItem）が返す属性だけで見分ける。ファイルを開かないので、ダウンロードは起きない。
#   RECALL_ON_DATA_ACCESS(0x400000)・RECALL_ON_OPEN(0x40000)・OFFLINE(0x1000) のどれかが付いていれば、
#   中身は手元に無い（開くとダウンロードされる）。
#   PINNED(0x80000)・UNPINNED(0x100000)・ReparsePoint だけでは、クラウドにだけあるとは言えない。

${attributeRecallOnDataAccess} = 0x400000
${attributeRecallOnOpen}       = 0x40000
${attributeOffline}            = 0x1000

function testCloudOnlyAttributes {
    # ファイルの属性（数値）が、中身が手元に無い（クラウドにだけある）ことを表すとき $true
    param (
        [int]$attributes
    )

    $mask = ${attributeRecallOnDataAccess} -bor ${attributeRecallOnOpen} -bor ${attributeOffline}
    return (($attributes -band $mask) -ne 0)
}

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
        [object[]]$lines,
        [System.Text.Encoding]$encoding = ${utf8Bom}  # 文字コード（既定は BOM 付き UTF-8）
    )

    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
    $tmpPath = "${path}.tmp"
    [System.IO.File]::WriteAllLines($tmpPath, [string[]]@($lines), $encoding)
    # 書いた直後のファイルは、ウイルス対策ソフト等が一時的に掴んでいて置き換えられないことがあるため、少し待って数回試す
    for ($i = 1; $true; $i++) {
        try {
            if (Test-Path -LiteralPath $path) {
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
    # 元のファイルに影響を出さずにコピーする（インデックス作成は、元のファイルでなくこのコピーを読んで行う）。
    # 書き込み中のファイルも、読み取りだけで開いてコピーできる（書きかけの中身が入ることは、直接読む場合と同じ）。
    # 排他で開かれたファイルはコピーできず、取り込みの失敗になる。
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

function invokeWithSourceCopy {
    # 元のファイルを作業領域へコピーし（copyFileShared）、コピーを読む処理（readBlock）を呼んでその戻り値を返し、最後にコピーを消す。
    # 取り込みの入口（テキスト・PDF・Excel・Word・PowerPoint）が、元のファイルを直接開かないために使う。
    # extraFiles は、読む処理が途中で作った作業ファイルの一覧（List[string]。読む処理の中で足してよい）。コピーと一緒に、あれば消す。
    # 消せなかったとき: 読む処理が例外なら、その例外をそのまま通す（消せなかったことは握る）。
    #   読む処理が成功したのに消せなかったときだけ「作業領域のコピーを消せませんでした: <元のメッセージ>」を投げる。
    # timing（hashtable。既定 $null）を渡すと、copyFileShared にかかったミリ秒を timing['CopyMs'] に足す（無ければ作る）。性能の計測用。$null のときは何もしない。
    # readBlock は子のスコープで動くので、中で代入した変数は呼んだ側に残らない。要る値は出力（戻り値）で返す。
    # 読む処理の中の COM の呼び出しなどの出力が混ざらないよう、出力は [void]・$null = で捨て、返す値は最後に 1 つだけ出す
    param (
        [string]$sourcePath,
        [string]$copyPath,
        [scriptblock]$readBlock,
        [System.Collections.Generic.List[string]]$extraFiles = $null,
        [hashtable]$timing = $null
    )

    $removeError = $null
    try {
        if ($null -ne $timing) {
            $watch = [System.Diagnostics.Stopwatch]::StartNew()
            try {
                copyFileShared $sourcePath $copyPath
            } finally {
                $timing['CopyMs'] = [double]$timing['CopyMs'] + $watch.Elapsed.TotalMilliseconds
            }
        } else {
            copyFileShared $sourcePath $copyPath
        }
        # 読む処理の出力は、まとめずにそのまま呼んだ側へ流す（配列・hashtable・何も出さない、のどれも、直接呼んだときと同じ形で届く）
        & $readBlock
    } finally {
        $targets = New-Object System.Collections.Generic.List[string]
        $targets.Add($copyPath)
        if ($null -ne $extraFiles) {
            $targets.AddRange($extraFiles)
        }
        foreach ($target in $targets) {
            try {
                if (Test-Path -LiteralPath (toLongPath $target)) {
                    Remove-Item -LiteralPath (toLongPath $target) -Force -ErrorAction Stop
                }
            } catch {
                if (!$removeError) {
                    $removeError = $_.Exception.Message
                }
            }
        }
    }
    if ($removeError) {
        throw "作業領域のコピーを消せませんでした: $removeError"
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

function testPathUnder {
    # path が root の「真下より下」（root そのものは含まない）かを返す。消す・書く直前に、
    # 状態ファイルや設定から読んだ名前でつないだパスが、決めた置き場の外へ出ていないかを確かめるのに使う。
    # \\?\ を外してから GetFullPath で正規化する（.. や . を解き、/ を \ にする）ため、\\?\ の有無・大文字小文字・末尾の \ に依らない。
    # リンクは見ない（辿るかどうかは削除の部品が決める）。解けない（不正な）パスは $false
    param (
        [string]$path,
        [string]$root
    )

    if (!$path -or !$root) {
        return $false
    }
    try {
        $fullPath = [System.IO.Path]::GetFullPath((fromLongPath $path))
        $fullRoot = [System.IO.Path]::GetFullPath((fromLongPath $root)).TrimEnd("\")
    } catch {
        return $false
    }
    return $fullPath.StartsWith("$fullRoot\", [System.StringComparison]::OrdinalIgnoreCase) -and
        $fullPath.TrimEnd("\").Length -gt $fullRoot.Length
}

function trimTrailingSeparator {
    # パスの末尾の \ と / を落とす（リンクの判定・削除の入口で使う。念のため、末尾の区切りでリンクの先を見ることが無いようにする）。
    # 途中の区切りは変えない（\\?\ で始まらないパスの / は、あとで toLongPath が \ にそろえる。\\?\ で始まるパスの / はそのまま）。
    # 落とした結果が空か : で終わるパス（C:\ ・ \\?\C:\ ・ \ ・ / ）と、\\?\ だけのパスは、そのまま返す。
    # それ以外は、UNC の根（\\server\share\）や \\?\Volume{...}\ も、末尾を落とす
    param (
        [string]$path
    )

    if (!$path) {
        return $path
    }
    $trimmed = $path.TrimEnd([char[]]@("\", "/"))
    if (!$trimmed -or $trimmed.EndsWith(":") -or $trimmed -eq "\\?") {
        return $path
    }
    return $trimmed
}

function removeDirectoryRetry {
    # フォルダを中身ごと削除する。ウイルス対策ソフト・エクスプローラーが一時的に掴んでいることがあるため、少し待って数回試す
    param (
        [string]$path,
        [int]$tries = 3,
        [int]$waitMilliseconds = 200
    )

    if (!$path) {
        return  # 空のパスは消すものが無い
    }
    # 中に長いパス（260文字超）のファイルがあっても削除できるよう \\?\ 付きで削除する
    $longPath = toLongPath $path
    for ($i = 1; $true; $i++) {
        if (![System.IO.Directory]::Exists($longPath)) {
            # Exists は、無いときだけでなく、届かない・読めないときも $false になる。「無い」と確かめられたときだけ成功にする
            # （フォルダでないもの（ファイル）が同じ名前であるときも、消すフォルダは無いので成功にする）
            $state = getPathState $path
            if ($state.State -eq ${pathStateMissing} -or ($state.State -eq ${pathStateFound} -and !$state.IsDirectory)) {
                return
            }
            if ($state.State -ne ${pathStateFound}) {
                # 届かない・読めない（権限が無いなど）は、待っても変わらないので、試し直さずに例外にする。理由は OS の文面をそのまま添える
                throw "フォルダがあるかどうか確かめられません: $path（$($state.Message)）"
            }
            # フォルダとして在るのに Exists が偽（開けない・消している途中など）。少し待って試し直す
            if ($i -ge $tries) {
                throw "フォルダがあるのに開けないため、消せません: $path"
            }
            Start-Sleep -Milliseconds $waitMilliseconds
            continue
        }
        try {
            # リンク（ジャンクション・シンボリックリンク）は辿らず、リンクそのものだけを外す（removeTreeNoFollow）。
            # 消せなければ例外になり、catch で試し直す
            removeTreeNoFollow $longPath
            # 例外が出なくても消えていないことがある（別のプロセスが同じ名前で作り直したとき・消す側の取りこぼしなど）。
            # 消えたことを .NET で確かめてから成功にする（呼ぶ側は、成功を見て取り込み一覧・状態・ログに記録する）
            if ([System.IO.Directory]::Exists($longPath)) {
                throw "フォルダが消えていません: $(fromLongPath $longPath)"
            }
            return
        } catch {
            if ($i -ge $tries) {
                throw
            }
            Start-Sleep -Milliseconds $waitMilliseconds
        }
    }
}

function getLinkKind {
    # item（Get-Item・Get-ChildItem・DirectoryInfo などが返す FileSystemInfo）の種類を返す。
    #   Link    … ジャンクション・シンボリックリンク（属性に ReparsePoint が付き、LinkType がそのどちらか）
    #   Unknown … 属性が読めない、またはフォルダなのに ReparsePoint が付き LinkType がジャンクション・シンボリックリンクと確かめられない
    #             （OneDrive のオンデマンドの置き場のフォルダなど）。リンクかどうか分からないので、辿らず、消しもしない
    #   Plain   … それ以外（ReparsePoint の付いたファイルは、辿るものが無いので Plain）
    # 先に属性の ReparsePoint を見て、付いているときだけ LinkType を読む（多くのファイルで余計な問い合わせをしない）。
    # 5.1 ではプロパティを読んだときの例外は $null になる（catch に届かない）ため、$null を読めなかったものとして扱う
    param (
        $item
    )

    $attributes = $item.Attributes
    if ($null -eq $attributes) {
        return "Unknown"
    }
    if (!($attributes -band [System.IO.FileAttributes]::ReparsePoint)) {
        return "Plain"
    }
    $linkType = [string]$item.LinkType
    if ($linkType -eq "Junction" -or $linkType -eq "SymbolicLink") {
        return "Link"
    }
    if ($attributes -band [System.IO.FileAttributes]::Directory) {
        return "Unknown"
    }
    return "Plain"
}

function testLinkItem {
    # item が、辿ってはいけないもの（getLinkKind が Link か Unknown）なら $true
    param (
        $item
    )

    return ((getLinkKind $item) -ne "Plain")
}

function testPathIsLink {
    # path そのもの（途中の階層は見ない）がジャンクション・シンボリックリンクなら $true。
    # 無い・空のパスは $false（まだ作っていない）。属性が読めない、リンクかどうか確かめられないフォルダ（getLinkKind が Unknown）のときは $true（辿らない側に倒す）
    param (
        [string]$path
    )

    if (!$path) {
        return $false
    }
    $long = toLongPath (trimTrailingSeparator $path)
    try {
        $attributes = [System.IO.File]::GetAttributes($long)
    } catch [System.IO.FileNotFoundException] {
        return $false
    } catch [System.IO.DirectoryNotFoundException] {
        return $false
    } catch {
        return $true
    }
    if (!($attributes -band [System.IO.FileAttributes]::ReparsePoint)) {
        return $false
    }
    if ($attributes -band [System.IO.FileAttributes]::Directory) {
        return (testLinkItem (New-Object System.IO.DirectoryInfo $long))
    }
    return (testLinkItem (New-Object System.IO.FileInfo $long))
}

function testPathHasLink {
    # base の下の path まで、途中の各階層（path を含み、base とそれより上は含まない）に、リンクがあれば $true。
    # 利用者が選んだ置き場（base）そのものが別の場所へのリンクでも、構成を変えずに使えるよう、base は見ない。
    # path が base と同じなら $false。base の下になければ、確かめられないので $true
    param (
        [string]$base,
        [string]$path
    )

    try {
        $baseFull = [System.IO.Path]::GetFullPath((fromLongPath $base)).TrimEnd("\")
        $pathFull = [System.IO.Path]::GetFullPath((fromLongPath $path)).TrimEnd("\")
    } catch {
        return $true
    }
    if ($pathFull.Equals($baseFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $false
    }
    if (!$pathFull.StartsWith("$baseFull\", [System.StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }
    $current = $baseFull
    foreach ($name in $pathFull.Substring($baseFull.Length + 1).Split("\")) {
        $current = "$current\$name"
        if (testPathIsLink $current) {
            return $true
        }
        if (!(Test-Path -LiteralPath (toLongPath $current))) {
            return $false  # ここから先はまだ無い
        }
    }
    return $false
}

function newDirectoryUnder {
    # root の下に path のフォルダ（途中の階層を含む）を作り、作れたら（すでにあるときも）$true を返す。
    # root から path までの途中の階層（path を含む）にリンクがあるとき、path が root の下でないときは、何も作らずに $false を返す。
    # 作ったあとにもう一度確かめ、リンクに替わっていれば $false を返す（作ってしまった分は消さない。中には何も書かない）。
    # root そのもの（とそれより上）がリンクのときは、構成を変えずに使えるよう、リンクとは見ない（testPathHasLink）。
    # path が root と同じときは、root を（無ければ）作って $true を返す。
    # 作れなかった理由（権限など）は例外で返る。ログは呼ぶ側が出す（shared/ はツールのログを知らない）
    param (
        [string]$path,
        [string]$root
    )

    if (!$path -or !$root -or (testPathHasLink $root $path)) {
        return $false
    }
    [System.IO.Directory]::CreateDirectory((toLongPath $path)) | Out-Null
    return !(testPathHasLink $root $path)
}

function removeEntryOnly {
    # ファイル 1 つ・リンク（ジャンクション・シンボリックリンク）1 つ・空のフォルダ 1 つ（item は FileSystemInfo）を、そのものだけ消す。
    # リンクの先は消さない。読み取り専用の属性は外してから消す
    param (
        [System.IO.FileSystemInfo]$item
    )

    $path = $item.FullName
    $attributes = $item.Attributes
    if ($attributes -band [System.IO.FileAttributes]::ReadOnly) {
        # 読み取り専用だけを外す（ほかの属性は変えない）
        [System.IO.File]::SetAttributes($path, ($attributes -band (-bnot [System.IO.FileAttributes]::ReadOnly)))
    }
    if ($item -is [System.IO.DirectoryInfo]) {
        [System.IO.Directory]::Delete($path, $false)
    } else {
        [System.IO.File]::Delete($path)
    }
}

function removeTreeNoFollow {
    # フォルダを中身ごと削除する。ジャンクション・シンボリックリンクは辿らず、リンクそのものだけを外す（先は消さない）。
    # path そのものがリンクなら、そのリンクだけを外す。フォルダが無ければ何もしない。
    # 先に全体を数える（リンクは辿らない）。リンクかどうか確かめられないフォルダ（getLinkKind が Unknown）が 1 つでもあれば、
    # 降りず、何も消さずに例外にする。何も消さないと言えるのは、この数える段で見つけた Unknown のときだけ。
    # そのあとで、ファイル・リンク・フォルダ（下から順）を消す。
    # 降りる直前と、フォルダのファイルを消す直前に、そのフォルダがリンクでないかを確かめ直す。
    # 数えたあとにフォルダがリンクに替わっていたら例外にする。それまでに消した分は戻らない。
    # 消す直前に確かめ直すのは、そのフォルダ自身だけ。上の階層が途中で替わったときは見つけられない。
    # 読み取り専用の属性は外して消す（Remove-Item -Force と同じ働き）。消せなければ例外にする
    param (
        [string]$path
    )

    $long = toLongPath (trimTrailingSeparator $path)
    if (![System.IO.Directory]::Exists($long)) {
        return
    }
    $filesByDirectory = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[string]]'
    $links = New-Object System.Collections.Generic.List[string]
    $directories = New-Object System.Collections.Generic.List[string]
    $pending = New-Object System.Collections.Generic.Stack[string]
    $pending.Push($long)
    while ($pending.Count -gt 0) {
        $current = New-Object System.IO.DirectoryInfo ($pending.Pop())
        $currentKind = getLinkKind $current
        if ($currentKind -eq "Unknown") {
            throw "リンクかどうか確かめられないフォルダのため、消しませんでした: $(fromLongPath $current.FullName)"
        }
        if ($currentKind -eq "Link") {
            $links.Add($current.FullName)
            continue
        }
        $directories.Add($current.FullName)
        $filesByDirectory[$current.FullName] = New-Object System.Collections.Generic.List[string]
        foreach ($entry in @($current.EnumerateFileSystemInfos("*", [System.IO.SearchOption]::TopDirectoryOnly))) {
            if ($entry -isnot [System.IO.DirectoryInfo]) {
                $filesByDirectory[$current.FullName].Add($entry.FullName)
            } elseif ((getLinkKind $entry) -eq "Link") {
                $links.Add($entry.FullName)
            } else {
                # 確かめられないもの（Unknown）も積む。降りる直前の確かめで例外にする（列挙の途中では投げない）
                $pending.Push($entry.FullName)
            }
        }
    }
    foreach ($directory in $directories) {
        if ((getLinkKind (New-Object System.IO.DirectoryInfo $directory)) -ne "Plain") {
            throw "数えたあとでリンクに替わったフォルダのため、中のファイルを消しませんでした: $(fromLongPath $directory)"
        }
        foreach ($file in $filesByDirectory[$directory]) {
            removeEntryOnly (New-Object System.IO.FileInfo $file)
        }
    }
    foreach ($link in $links) {
        removeEntryOnly (New-Object System.IO.DirectoryInfo $link)
    }
    for ($i = $directories.Count - 1; $i -ge 0; $i--) {
        removeEntryOnly (New-Object System.IO.DirectoryInfo $directories[$i])
    }
}

function getFilesNoFollow {
    # root の下のファイルとフォルダを、ジャンクション・シンボリックリンクは辿らずに列挙する。
    # 再帰の列挙（Get-ChildItem -Recurse・AllDirectories）の代わりに使う（tests/meta/link_safety.Tests.ps1 が、それらを許した関数の外で書かせない）。
    # root そのものがリンクかは見ない（呼ぶ側が決める）。root が作業フォルダの置き場のようにリンクでよい場所のときは、その先の中を列挙する。
    # 戻り値: @{ Files; Directories; Links; HasError; Reason }
    #   Files       … ファイル（FileInfo。FullName は \\?\ 付き）。リンクのファイルは入れない
    #   Directories … フォルダ（DirectoryInfo。root は含まない）。リンクのフォルダは入れない
    #   Links       … 飛ばしたリンク（ファイル・フォルダ）の、\\?\ の付かない通常のパス
    #   HasError    … 読めなかったフォルダがあったか。読めたところまでは Files・Directories に入る（途中で例外にしない）
    #   Reason      … 最初に読めなかった理由
    #   Stopped     … shouldStop が $true を返して、途中で止めたか。フォルダ 1 つを読むごとに見て、止めたら読んだところまでを返す
    # リンクかどうかの見分けは getLinkKind。リンクかどうか確かめられないフォルダ（ReparsePoint が付き、LinkType が取れない。
    # OneDrive のオンデマンドのフォルダなど）は、SkipUnknown が無ければ降りる（読むだけのため）。属性を書くときは SkipUnknown を付け、降りずに Links に入れる。
    # Hidden の付いたものは、IncludeHidden が無ければ入れない（Get-ChildItem の既定と同じ。System だけのものは入れる）。
    # 並びは Get-ChildItem -Recurse と同じ（フォルダごとに、その中身を並べてから、サブフォルダの下へ進む）
    # TopOnly を付けると root の直下だけを読み、サブフォルダの下へは進まない（直下のフォルダは Directories に入る。直下のリンクは Links に入る）
    param (
        [string]$root,
        [switch]$IncludeHidden,
        [switch]$SkipUnknown,
        [switch]$TopOnly,
        [scriptblock]$shouldStop = $null
    )

    $files = New-Object System.Collections.Generic.List[object]
    $directories = New-Object System.Collections.Generic.List[object]
    $links = New-Object System.Collections.Generic.List[string]
    $hasError = $false
    $reason = ""
    $stopped = $false
    $reparseFlag = [System.IO.FileAttributes]::ReparsePoint
    $hiddenFlag = [System.IO.FileAttributes]::Hidden
    $pending = New-Object System.Collections.Generic.Stack[string]
    $pending.Push((toLongPath (trimTrailingSeparator $root)))
    while ($pending.Count -gt 0) {
        if ($null -ne $shouldStop -and (& $shouldStop)) {
            $stopped = $true
            break
        }
        $directory = New-Object System.IO.DirectoryInfo ($pending.Pop())
        try {
            # 列挙の途中で読めなくなることがあるので、配列にして受ける（例外はこの try で受ける）
            $entries = @($directory.EnumerateFileSystemInfos("*", [System.IO.SearchOption]::TopDirectoryOnly))
        } catch {
            $hasError = $true
            if (!$reason) {
                $reason = $_.Exception.Message
            }
            continue
        }
        $subDirectories = New-Object System.Collections.Generic.List[string]
        foreach ($entry in $entries) {
            $attributes = $entry.Attributes
            if (!$IncludeHidden -and ($attributes -band $hiddenFlag)) {
                continue
            }
            if ($attributes -band $reparseFlag) {
                # 属性に ReparsePoint が付いているものだけ、リンクかを調べる（多くのファイルで余計な問い合わせをしない）
                $kind = getLinkKind $entry
                if ($kind -eq "Link" -or ($kind -eq "Unknown" -and $SkipUnknown)) {
                    $links.Add((fromLongPath $entry.FullName))
                    continue
                }
            }
            if ($entry -is [System.IO.DirectoryInfo]) {
                $directories.Add($entry)
                if (!$TopOnly) { $subDirectories.Add($entry.FullName) }
            } else {
                $files.Add($entry)
            }
        }
        # Get-ChildItem -Recurse と同じ順（そのフォルダの中身、次に先頭のサブフォルダの下すべて、…）にするため、逆から積む
        for ($i = $subDirectories.Count - 1; $i -ge 0; $i--) {
            $pending.Push($subDirectories[$i])
        }
    }
    return @{ Files = $files.ToArray(); Directories = $directories.ToArray(); Links = $links.ToArray(); HasError = $hasError; Reason = $reason; Stopped = $stopped }
}

function moveDirectoryRetry {
    # フォルダを移動する。removeDirectoryRetry と同じ理由（ウイルス対策ソフト・エクスプローラーが、
    # 作ったばかり・書いたばかりのフォルダを一時的に掴んでいることがある）で、待つ時間を倍にしながら数回試す
    # （削除より、掴まれている時間が長くかかることがあるため、既定の回数・最初の待ち時間を大きくしている）
    param (
        [string]$source,
        [string]$destination,
        [int]$tries = 5,
        [int]$waitMilliseconds = 200
    )

    $longSource = toLongPath $source
    $longDestination = toLongPath $destination
    $wait = $waitMilliseconds
    for ($i = 1; $true; $i++) {
        try {
            [System.IO.Directory]::Move($longSource, $longDestination)
            return
        } catch {
            if ($i -ge $tries) {
                throw
            }
            Start-Sleep -Milliseconds $wait
            $wait *= 2
        }
    }
}

function testSameDriveRoot {
    # 2 つのパスが同じドライブ（同じ根）にあるかを返す。同じなら Directory.Move で移せ、違えば写してから消す。
    # 呼ぶ側のテストが、別のドライブを用意せずに、この判定を差し替えられるよう、関数にしてある
    param (
        [string]$first,
        [string]$second
    )

    return [System.IO.Path]::GetPathRoot($first) -eq [System.IO.Path]::GetPathRoot($second)
}

function testRemovablePath {
    # path を、root の下で消す・移す・作ってよいかを返す。次のすべてが成り立つときだけ $true。
    # ・path と root が絶対パス（ドライブか UNC。\\?\ 付きも可）である（相対パスは現在の場所に依って意味が変わるため偽）
    # ・path が root の真下より下である（root そのもの・root より上・似た名前の別のフォルダは偽。testPathUnder）
    # ・root から path までの途中の階層にリンクが無い（testPathHasLink）。root そのものがリンクの構成は今までどおり動く。
    #   まだ無い階層は見ない（移動の先・作る先に使うため、根の下の無いパスは真）
    # 確かめられない（読めない・例外）ときは $false
    # leafLinkOk を付けると、path そのものがリンクでもよい（途中の階層だけ見る）。消す・移すフォルダそのものがリンクのときは、
    # リンクだけを外す・動かす（先は辿らない）ため、これまでどおり扱えるように、removeDirectoryUnder・moveDirectoryUnder が付ける
    param (
        [string]$path,
        [string]$root,
        [switch]$leafLinkOk
    )

    if (!$path -or !$root) {
        return $false
    }
    $absolutePattern = '^(\\\\|[A-Za-z]:[\\/])'
    if (($path -notmatch $absolutePattern) -or ($root -notmatch $absolutePattern)) {
        return $false
    }
    try {
        if (!(testPathUnder $path $root)) {
            return $false
        }
        $linkFrom = $path
        if ($leafLinkOk) {
            $linkFrom = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath((fromLongPath $path)).TrimEnd("\"))
        }
        return !(testPathHasLink $root $linkFrom)
    } catch {
        return $false
    }
}

function removeDirectoryUnder {
    # path のフォルダを、root の下で途中にリンクが無いことを確かめてから、中身ごと消す（removeDirectoryRetry）。
    # 確かめ（testRemovablePath）が偽なら何も消さずに $false、消したら（無くて消すものが無いときも）$true を返す。
    # 消せなかった理由は例外のまま返る。ログは呼ぶ側が出す
    param (
        [string]$path,
        [string]$root,
        [int]$tries = 3
    )

    if (!(testRemovablePath $path $root -leafLinkOk)) {
        return $false
    }
    removeDirectoryRetry $path $tries
    return $true
}

function removeFileUnder {
    # path のファイル 1 つを、root の下で途中にリンクが無いことを確かめてから消す（removeEntryOnly。path 自身がリンクなら、リンクだけを外す）。
    # 確かめ（testRemovablePath）が偽なら何も消さずに $false、消したら（無くて消すものが無いときも）$true を返す。
    # 消せなかった理由は例外のまま返る。ログは呼ぶ側が出す
    param (
        [string]$path,
        [string]$root
    )

    if (!(testRemovablePath $path $root -leafLinkOk)) {
        return $false
    }
    $long = toLongPath $path
    if ([System.IO.File]::Exists($long)) {
        removeEntryOnly (New-Object System.IO.FileInfo $long)
    }
    return $true
}

function moveDirectoryUnder {
    # source のフォルダを destination へ、どちらも root の下で途中にリンクが無いことを確かめてから移す（moveDirectoryRetry）。
    # 移すと元の場所から無くなるため、消すのと同じ確かめを元と先の両方に行う。偽なら何も移さずに $false、移したら $true を返す。
    # 元と先で根が違うときは、先の根を root、元の根を sourceRoot に渡す（省略したら元も root の下で確かめる）
    param (
        [string]$source,
        [string]$destination,
        [string]$root,
        [string]$sourceRoot = ""
    )

    if ($sourceRoot -eq "") {
        $sourceRoot = $root
    }
    if (!(testRemovablePath $source $sourceRoot -leafLinkOk) -or !(testRemovablePath $destination $root -leafLinkOk)) {
        return $false
    }
    moveDirectoryRetry $source $destination
    return $true
}

function testAnyEntry {
    # 列挙（[System.IO.Directory]::EnumerateFiles などの結果）に、predicate が $true を返すもの（省略したら何でも）が
    # 1 つでもあるか。見つけたところで列挙をやめ、列挙子を必ず Dispose する。
    # foreach を break・return・throw で途中で抜けると列挙子が Dispose されず、調べていた下のフォルダを開いたまま残る。
    # GC されるまで、そのフォルダと上のフォルダを移動・削除できなくなる（Directory.Move がアクセス拒否になる）ため、
    # 途中でやめる列挙はこの関数を通す
    param (
        [System.Collections.IEnumerable]$entries,
        [scriptblock]$predicate
    )

    $enumerator = $entries.GetEnumerator()
    try {
        while ($enumerator.MoveNext()) {
            if ($null -eq $predicate -or (& $predicate $enumerator.Current)) {
                return $true
            }
        }
        return $false
    } finally {
        if ($enumerator -is [System.IDisposable]) {
            $enumerator.Dispose()
        }
    }
}

function selectFirstEntries {
    # 列挙（[System.IO.Directory]::EnumerateFileSystemEntries などの結果）の先頭から count 件までを配列にして返す。
    # count 件で列挙をやめ、列挙子を必ず Dispose する（testAnyEntry と同じ理由。Select-Object -First で途中でやめる形にしない）
    param (
        [System.Collections.IEnumerable]$entries,
        [int]$count
    )

    $list = New-Object System.Collections.Generic.List[object]
    $enumerator = $entries.GetEnumerator()
    try {
        while ($list.Count -lt $count -and $enumerator.MoveNext()) {
            $list.Add($enumerator.Current)
        }
    } finally {
        if ($enumerator -is [System.IDisposable]) {
            $enumerator.Dispose()
        }
    }
    return , $list.ToArray()
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
    #   app: ツールの名前（ミューテックスの名前の先頭。呼ぶ側が渡す。空なら例外にする）
    param (
        [string]$name,
        [string]$dir = ${rootDir},
        [string]$app = ""
    )

    if ([string]::IsNullOrEmpty($app)) {
        throw "newAppMutex: ツールの名前（app）が空です"
    }
    $key = getFolderKey $dir
    $createdNew = $false
    $mutex = New-Object System.Threading.Mutex($true, "Local\${app}_${name}_${key}", [ref]$createdNew)
    return @{ Mutex = $mutex; Acquired = $createdNew }
}

function openLockFile {
    # 鍵のファイルを、ほかから開けない形（FileShare.None）で開いて返す。閉じると消える（DeleteOnClose）。
    # ほかが持っていれば $null。持っているかどうかは「ファイルがあるか」でなく「開けるか」で決める
    # （強制終了で残ったファイルは誰も開いていないので開ける。プロセスが終われば OS が閉じる）。
    # ミューテックス（newAppMutex）と違い、別のセッション・別の PC・別のパスの書き方でも効く（共有フォルダでも、サーバーが開き方を守る場合）
    # 共有違反・ロック違反（Win32 の 32・33）は「ほかが持っている」。アクセス拒否（Win32 の 5）は、前の持ち主が閉じた直後で
    # ファイルが削除待ちのときにも起きるため、32・33 と同じく attempts の回数までやり直す。最後まで拒否なら例外にする。
    # それ以外の失敗（親フォルダが無い・接続できないなど）は、そのまま投げる
    #   attempts: 試す回数（1 なら待たずに 1 回だけ）。画面のスレッドでは 1 にする
    param (
        [string]$path,
        [int]$attempts = 3,
        [int]$waitMilliseconds = 200
    )

    $longPath = toLongPath $path
    $denied = $false
    for ($i = 1; $i -le $attempts; $i++) {
        try {
            return [System.IO.FileStream]::new($longPath, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite,
                [System.IO.FileShare]::None, 4096, [System.IO.FileOptions]::DeleteOnClose)
        } catch {
            $inner = $_.Exception
            while ($inner -is [System.Management.Automation.MethodInvocationException] -and $null -ne $inner.InnerException) {
                $inner = $inner.InnerException
            }
            if ($inner -is [System.UnauthorizedAccessException]) {
                $denied = $true
            } elseif ($inner -is [System.IO.IOException] -and (($inner.HResult -band 0xFFFF) -eq 32 -or ($inner.HResult -band 0xFFFF) -eq 33)) {
                $denied = $false
            } else {
                throw
            }
        }
        if ($i -lt $attempts) {
            Start-Sleep -Milliseconds $waitMilliseconds
        }
    }
    if ($denied) {
        throw "インデックス作成の鍵のファイル（${path}）を開けませんでした。ほかのインデックス作成が終わった直後か、ワークスペースに書き込む権限がありません。しばらく待ってからやり直し、続くときはワークスペースのフォルダに書き込めるかを確かめてください。"
    }
    return $null
}

function getPathErrorKind {
    # File.GetAttributes などが投げた例外（PowerShell から呼ぶと MethodInvocationException に包まれる）を、
    # 見つからない（FileNotFoundException・DirectoryNotFoundException）・接続できない（Win32 のエラー番号）・
    # その他（アクセス拒否・ログオンの失敗など上のどれにも当たらないもの）に分ける。純粋な関数（例外を受け取って値を返すだけ）
    param (
        $exception  # System.Exception（MethodInvocationException でもそのままでもよい）
    )

    $inner = $exception
    while ($inner -is [System.Management.Automation.MethodInvocationException] -and $null -ne $inner.InnerException) {
        $inner = $inner.InnerException
    }
    if ($inner -is [System.IO.FileNotFoundException] -or $inner -is [System.IO.DirectoryNotFoundException]) {
        return ${pathStateMissing}
    }
    $code = $inner.HResult -band 0xFFFF
    if (${unreachableWin32Codes} -contains $code) {
        return ${pathStateUnreachable}
    }
    return ${pathStateOther}
}

function getPathState {
    # パスの状態を、例外を投げる呼び出し（File.GetAttributes）で調べる。260 文字を超えるパスも扱えるよう \\?\ を付けて渡す。
    # 返すもの: @{ State（上の定数）; IsDirectory; Message（見つかったとき以外の例外の文面） }
    param (
        [string]$path
    )

    try {
        $attributes = [System.IO.File]::GetAttributes((toLongPath $path))
        return @{ State = ${pathStateFound}; IsDirectory = [bool]($attributes -band [System.IO.FileAttributes]::Directory); Message = "" }
    } catch {
        return @{ State = (getPathErrorKind $_.Exception); IsDirectory = $false; Message = $_.Exception.Message }
    }
}

function invokeWithNamedMutex {
    # 名前付きミューテックスを取って action を実行し、action の出力を返す（終わったら必ず手放す）。
    # 同じスレッドの入れ子は通す（ミューテックスは同じスレッドなら再び取れる）。
    # timeoutMilliseconds を過ぎても取れなければ、「排他の待ちが時間切れになりました」の例外にする。
    # 前の持ち主が手放さずに終わっていた（abandoned）ときは、取れたものとして続ける
    #   mutexName: ミューテックスの名前（呼ぶ側が決める。同じ名前どうしで排他する）。
    #   action から呼び出し元の変数を参照できるよう、この関数の変数の名前は目立つものにしてある（name・key・path などは使わない）
    param (
        [string]$mutexName,
        [int]$timeoutMilliseconds,
        [scriptblock]$action
    )

    $mutex = New-Object System.Threading.Mutex($false, $mutexName)
    try {
        $acquired = $false
        try {
            $acquired = $mutex.WaitOne($timeoutMilliseconds)
        } catch {
            # PowerShell では MethodInvocationException に包まれて届く。abandoned でも所有は移っている
            $inner = $_.Exception
            while ($null -ne $inner -and $inner -isnot [System.Threading.AbandonedMutexException]) {
                $inner = $inner.InnerException
            }
            if ($null -eq $inner) {
                throw
            }
            $acquired = $true
        }
        if (!$acquired) {
            throw "排他の待ちが時間切れになりました（${timeoutMilliseconds} ミリ秒。${mutexName}）。ほかの処理が終わってからやり直してください。"
        }
        try {
            & $action
        } finally {
            $mutex.ReleaseMutex()
        }
    } finally {
        $mutex.Dispose()
    }
}
