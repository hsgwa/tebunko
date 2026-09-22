# 検索結果から元のファイルの場所を求める（元のフォルダ.txt の読み書きを含む）。

function writeSourceFolderFile {
    # 各インデックスのフォルダに、インデックス名とクロール対象フォルダの対応（元のフォルダ.txt）を書き出す。
    # 1行目は説明、2行目は "インデックス名<TAB>クロール対象フォルダ"。
    # インデックス1件につき1ファイルのため、<インデックス名> のフォルダだけを別の PC・場所へコピーしても
    # 元のファイルの場所が分かる（work\index ごとコピーした場合は readSourceFolderFile が各フォルダを読む）
    param (
        [object[]]$folders,  # assignIndexNames の結果（@{ Path; Name }）
        [string]$dir = ${indexDir}
    )

    $header = "# 検索結果から元のファイルを開くときに使う、インデックス名とクロール対象フォルダの対応です（インデックス作成のたびに作り直します）"
    $items = @($folders | Where-Object { $_ -and $_.Name })

    # 以前の版は、インデックスのフォルダ直下にも全インデックス分の 元のフォルダ.txt を書いていた。
    # そこにしか記録の無いインデックス（クロール対象フォルダから外したものなど）の分を各フォルダへ移してから、直下のファイルを消す
    $rootPath = Join-Path $dir ${sourceFolderFileName}
    $names = @($items | ForEach-Object { $_.Name })
    foreach ($line in @(readListFile $rootPath)) {
        $fields = $line.Split("`t")
        if ($fields.Count -eq 2 -and $fields[0] -ne "" -and $fields[1] -ne "" -and $names -notcontains $fields[0]) {
            $items += New-Object PSObject -Property ([ordered]@{ Name = $fields[0]; Path = $fields[1] })
            $names += $fields[0]
        }
    }

    foreach ($folder in $items) {
        # インデックスのフォルダがまだ無い（1件も取り込んでいない）場合は作らない
        $indexPath = Join-Path $dir $folder.Name
        if (Test-Path -LiteralPath (toLongPath $indexPath) -PathType Container) {
            writeListFile (Join-Path $indexPath ${sourceFolderFileName}) @($header, "$($folder.Name)`t$($folder.Path)")
        }
    }

    if (Test-Path -LiteralPath $rootPath) {
        Remove-Item -LiteralPath $rootPath -Force
    }
}

function readSourceFolderFile {
    # インデックスのフォルダ（dir）直下の 元のフォルダ.txt を読み、インデックス名 → クロール対象フォルダ を返す
    # （大文字・小文字を区別しない）。ファイルが無ければ空。
    # ファイルは各インデックスのフォルダ（work\index\<インデックス名>）に置くため、dir にはそのフォルダを渡す。
    # dir の下を探し回らないのは、インデックスのフォルダの下が元のファイル1つにつき1フォルダ（数万個）になるため
    param (
        [string]$dir
    )

    # ハッシュテーブルは大文字・小文字を区別しない（制限モード（制限言語モード）からも使うため Dictionary にしない）
    $map = @{}
    foreach ($line in @(readListFile (Join-Path $dir ${sourceFolderFileName}))) {
        $fields = $line.Split("`t")
        if ($fields.Count -eq 2 -and $fields[0] -ne "" -and $fields[1] -ne "") {
            $map[$fields[0]] = $fields[1]
        }
    }
    return , $map
}

function getSourceFolderMap {
    # インデックスのフォルダ（dir）の インデックス名 → 元のフォルダ（今そのフォルダが置かれている場所）を返す。
    # 次の順に読み、後のもので上書きする（後のものが優先）:
    #   1. dir 直下の 元のフォルダ.txt … インデックスを作ったときの場所。インデックスをコピーしても付いてくる
    #   2. 既定のインデックス（work\index）なら取り込み一覧の記録
    #   3. 設定のインデックス名に対する場所（targetFolders・indexSources）… 利用者が指定した「今の場所」のため最も優先する
    param (
        [string]$dir,
        [string]$statusPath = ${statusFile},
        [string]$settingsPath = ${settingsFile}
    )

    $map = readSourceFolderFile $dir
    if ((Test-Path -LiteralPath ${indexDir} -PathType Container) -and
        (testSameFolder $dir (Resolve-Path -LiteralPath ${indexDir}).ProviderPath)) {
        foreach ($entry in (getIndexNameMap $statusPath).GetEnumerator()) {
            $map[$entry.Key] = $entry.Value
        }
    }
    foreach ($folder in @(getTargetFolders $settingsPath | Where-Object { $_.Name })) {
        $map[$folder.Name] = $folder.Path
    }
    foreach ($source in @(readIndexSources $settingsPath)) {
        $map[$source.Name] = $source.Path
    }
    return , $map
}

function joinSourcePath {
    # フォルダ・相対フォルダ（空でも可）・ファイル名をつなぐ（ドライブ直下 "D:\" でも \ が重ならないようにする）
    param (
        [string]$folder,
        [string]$rest,
        [string]$name = ""
    )

    $path = $folder.TrimEnd("\")
    foreach ($part in @($rest, $name)) {
        if ($part) {
            $path += "\$part"
        }
    }
    return $path
}

function getSourceLocation {
    # 検索結果の元のファイルの場所を @{ Name（インデックス名）; Folder（元のフォルダ）; Rest（その下の相対フォルダ）; Known } で返す。
    # インデックスのフォルダ（Root）の下は "インデックス名\相対フォルダ" のため、インデックス名から元のフォルダを引く（getSourceFolderMap）。
    # 検索対象にインデックス名のフォルダ（…\index\<インデックス名>）を直接指定した場合は、親フォルダの記録を使う。
    # 元のフォルダが分からなければ Known = $false・Folder = "" とする（Name は返すため、フォルダを選んでもらえば設定に記録できる）
    #   maps: フォルダ → getSourceFolderMap の結果 のキャッシュ（読んだ結果を追加する）
    param (
        $hit,
        [hashtable]$maps = @{}
    )

    $root = ([string]$hit.Root).TrimEnd("\")
    $relDir = [string]$hit.RelDir
    $candidates = @()
    if ($relDir) {
        $parts = splitIndexRelPath $relDir
        $candidates += @{ Dir = $root; Name = $parts.Name; Rest = $parts.Rest }
        # インデックスのフォルダの中の記録（<インデックス名> のフォルダだけを別の場所へコピーした場合）
        $candidates += @{ Dir = (Join-Path $root $parts.Name); Name = $parts.Name; Rest = $parts.Rest }
    }
    # 検索対象にインデックス名のフォルダ（…\index\<インデックス名>）を直接指定した場合。
    # そのフォルダ自身の記録を先に見て、無ければ親フォルダの記録を見る
    $leaf = Split-Path $root -Leaf
    $candidates += @{ Dir = $root; Name = $leaf; Rest = $relDir }
    $parent = Split-Path $root -Parent
    if ($parent) {
        $candidates += @{ Dir = $parent; Name = $leaf; Rest = $relDir }
    }

    foreach ($candidate in $candidates) {
        if (!$maps.ContainsKey($candidate.Dir)) {
            $maps[$candidate.Dir] = getSourceFolderMap $candidate.Dir
        }
        $map = $maps[$candidate.Dir]
        if ($map.ContainsKey($candidate.Name)) {
            return @{ Name = $candidate.Name; Folder = $map[$candidate.Name]; Rest = $candidate.Rest; Known = $true }
        }
    }

    # 分からない場合も、インデックス名と、その下の相対フォルダは分かる（検索対象にインデックス名のフォルダを直接指定した場合は Rest がすべて）
    if ($candidates.Count -gt 0) {
        return @{ Name = $candidates[0].Name; Folder = ""; Rest = $candidates[0].Rest; Known = $false }
    }
    return @{ Name = (Split-Path $root -Leaf); Folder = ""; Rest = ""; Known = $false }
}

function findMovedSource {
    # 元のファイルが見つからないとき、選んでもらったフォルダ（picked）の中から探す。
    # picked は元のフォルダ（インデックスのルート）に当たるフォルダでも、その下のどのフォルダ（ファイルのあるフォルダなど）に当たるフォルダでもよい。
    # 上の階層に当たるとみなす方から順に試し、見つかれば @{ Path（見つかったファイル）; Root（元のフォルダに当たるフォルダ。遡れなければ ""） }、
    # 見つからなければ $null を返す。Root を設定に記録すれば（setIndexSourceFolder）、同じインデックスのほかのファイルも開ける
    param (
        [string]$picked,
        [string]$rest,
        [string]$book
    )

    $picked = normalizeFolderPath $picked
    $segments = [string[]]@($rest.Split([char[]]@("\"), [System.StringSplitOptions]::RemoveEmptyEntries))
    for ($skip = 0; $skip -le $segments.Count; $skip++) {
        # 先頭の skip 個のフォルダは picked より上、残りは picked の下にあるとみなす
        $below = [string]::Join("\", $segments, $skip, $segments.Count - $skip)
        $candidate = joinSourcePath $picked $below $book
        if (Test-Path -LiteralPath (toLongPath $candidate) -PathType Leaf) {
            # 選んだフォルダは「元のフォルダ＋先頭 skip 個のフォルダ」に当たるため、skip 個上が元のフォルダに当たる
            $root = $picked
            for ($i = 0; $i -lt $skip -and $root; $i++) {
                $root = Split-Path $root -Parent
            }
            return @{ Path = $candidate; Root = [string]$root }
        }
    }
    return $null
}

function resolveSourcePath {
    # 検索結果の元のファイルのパスを返す（ファイルがあるかは確かめない）。元のフォルダが分からなければ $null
    #   maps: getSourceLocation のキャッシュ
    param (
        $hit,
        [hashtable]$maps = @{}
    )

    $location = getSourceLocation $hit $maps
    if (!$location.Known) {
        return $null
    }
    return (joinSourcePath $location.Folder $location.Rest $hit.Book)
}
