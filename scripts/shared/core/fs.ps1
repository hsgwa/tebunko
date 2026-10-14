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

function setNotContentIndexed {
    # フォルダ（またはファイル）に「内容のインデックスを作成しない」属性（NotContentIndexed）を付ける。
    # 既に付いているものには付け直さない（Changed に数えない）ので、ほかの属性（読み取り専用など）は変えない。
    #   Recurse 無し: path 自身と、フォルダなら直下のファイルに付ける（下のフォルダには付けない）
    #   Recurse あり: path 自身と、下のすべてのフォルダ・ファイルに付ける
    # Recurse でフォルダを渡したときは、根フォルダの属性を下より先に確かめる。既に付いていれば
    # （前回すべて付け終えて根に付けたということなので）下はたどらず、GetAttributes 1 回だけで終える。
    # 根の属性は、下をすべて付け終えて成功したときにだけ最後に付ける（列挙・付与の途中で失敗した回は
    # 根に付けない。次回もう一度下からたどり直せるようにするため）。
    # 1件ずつの失敗（アクセス拒否など）も、列挙そのものの失敗（無いパス・列挙の途中で読めなくなったフォルダなど）も
    # 例外にせず、@{ Ok; Changed; Reason }（最初に見つかった理由を Reason に）を返す。
    # 呼び出し元（インデックス作成）は、Ok が $false でも処理を止めず、ログに 1 行残すだけにする
    param (
        [string]$path,
        [switch]$Recurse
    )

    $flag = [System.IO.FileAttributes]::NotContentIndexed
    $state = getPathState $path
    if ($state.State -ne ${pathStateFound}) {
        $reason = if ($state.Message) { $state.Message } else { "パスが見つかりません（${path}）" }
        return @{ Ok = $false; Changed = 0; Reason = $reason }
    }

    $long = toLongPath $path
    $deferRoot = [bool]($Recurse -and $state.IsDirectory)
    if ($deferRoot) {
        try {
            $rootAttrs = [System.IO.File]::GetAttributes($long)
        } catch {
            return @{ Ok = $false; Changed = 0; Reason = $_.Exception.Message }
        }
        if ($rootAttrs -band $flag) {
            return @{ Ok = $true; Changed = 0; Reason = "" }
        }
    }

    # 先に対象（path 自身と、必要なら中身）をすべて列挙してから属性を付ける。deferRoot のときは
    # 根自身をここでは含めず、下だけを集める（根は下がすべて終わったあとに付ける）。
    # 列挙そのものが例外になったら（途中で消えた・読めないフォルダがあるなど）、途中まで集めた分は捨てて Ok=$false にする
    $items = New-Object System.Collections.Generic.List[object]
    if (!$deferRoot) {
        $items.Add(@{ Path = $path; IsDirectory = $state.IsDirectory })
    }
    try {
        if ($state.IsDirectory) {
            $dirInfo = New-Object System.IO.DirectoryInfo $long
            if ($Recurse) {
                foreach ($entry in $dirInfo.EnumerateFileSystemInfos("*", [System.IO.SearchOption]::AllDirectories)) {
                    $items.Add(@{ Path = (fromLongPath $entry.FullName); IsDirectory = ($entry -is [System.IO.DirectoryInfo]) })
                }
            } else {
                foreach ($file in $dirInfo.EnumerateFiles("*", [System.IO.SearchOption]::TopDirectoryOnly)) {
                    $items.Add(@{ Path = (fromLongPath $file.FullName); IsDirectory = $false })
                }
            }
        }
    } catch {
        return @{ Ok = $false; Changed = 0; Reason = $_.Exception.Message }
    }

    $changed = 0
    $ok = $true
    $reason = ""
    foreach ($item in $items) {
        try {
            # フォルダ・ファイルどちらも [System.IO.File]::Get/SetAttributes で扱う（Directory クラスには同名のメソッドが無い）
            $itemLong = toLongPath $item.Path
            $attrs = [System.IO.File]::GetAttributes($itemLong)
            if ($attrs -band $flag) {
                continue
            }
            [System.IO.File]::SetAttributes($itemLong, ($attrs -bor $flag))
            $changed++
        } catch {
            $ok = $false
            if (!$reason) {
                $reason = $_.Exception.Message
            }
        }
    }

    if ($deferRoot -and $ok) {
        try {
            [System.IO.File]::SetAttributes($long, ($rootAttrs -bor $flag))
            $changed++
        } catch {
            $ok = $false
            if (!$reason) {
                $reason = $_.Exception.Message
            }
        }
    }

    return @{ Ok = $ok; Changed = $changed; Reason = $reason }
}

function testNotContentIndexed {
    # path に「内容のインデックスを作成しない」属性（NotContentIndexed）が付いているかを返す。
    # 属性を確かめられなければ（パスが無い・アクセスできないなど）例外にせず $false を返す
    param (
        [string]$path
    )

    try {
        $attrs = [System.IO.File]::GetAttributes((toLongPath $path))
        return [bool]($attrs -band [System.IO.FileAttributes]::NotContentIndexed)
    } catch {
        return $false
    }
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
            # Exists は、無いときだけでなく、届かない・読めないときも $false になる。「無い」と確かめられたときだけ成功にする
            # （フォルダでないもの（ファイル）が同じ名前であるときも、消すフォルダは無いので成功にする）
            $state = getPathState $path
            if ($state.State -eq ${pathStateMissing} -or ($state.State -eq ${pathStateFound} -and !$state.IsDirectory)) {
                return
            }
            if ($i -ge $tries) {
                throw "フォルダの有無を確かめられません: $path（$($state.State)）"
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
    param (
        [string]$name,
        [string]$dir = ${rootDir}
    )

    $key = getFolderKey $dir
    $createdNew = $false
    $mutex = New-Object System.Threading.Mutex($true, "Local\${appId}_${name}_${key}", [ref]$createdNew)
    return @{ Mutex = $mutex; Acquired = $createdNew }
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
