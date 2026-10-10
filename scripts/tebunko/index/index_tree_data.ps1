# ［検索］の検索対象ツリーのデータ集め（状態層。画面に触らない）。
# ワークスペースのフォルダを読むので、届かないネットワークのワークスペースでは、画面のスレッドでなく裏の列（network）から呼ぶ。
# 画面（ui\index_tree.ps1・ui\types.ps1 の IndexNode）は、この結果を受け取って組み立てるだけにする。

function testIndexBookName {
    # 元のファイルごとのフォルダと分かる拡張子か。Office は拡張子の形（.xls・.doc・.ppt で始まり 4〜5 文字）で見分け、
    # テキストは対象の拡張子の一覧（shared\core\text_file.ps1 の testTextExtension。一覧はそこにだけある）で見分ける
    param (
        [string]$name
    )

    $ext = [System.IO.Path]::GetExtension($name).ToLowerInvariant()
    if ($ext.Length -ge 4 -and $ext.Length -le 5 -and
        ($ext.StartsWith(".xls") -or $ext.StartsWith(".doc") -or $ext.StartsWith(".ppt"))) {
        return $true
    }
    return (testTextExtension $name)
}

function testIndexBookDirPath {
    # 元のファイルごとのフォルダ（集約する前の TSV・中身が空のファイルのフォルダ）か。名前が .xlsx などで終わる本物のフォルダと
    # 区別するため、サブフォルダも集約ファイル（content_index.*.tsv）も無いことも見る（pack_store.ps1 の testIndexBookDir と同じ判定）
    param (
        [string]$dir
    )

    if (-not (testIndexBookName ([System.IO.Path]::GetFileName($dir.TrimEnd('\'))))) {
        return $false
    }
    try {
        # 列挙子（Enumerate*）を途中で抜けると、GC まで調べたフォルダを掴んだまま残り、インポートの上書きで
        # content_index\<名前> を移動できなくなる。直下だけなので配列（Get*）で受ける（ここの関数はどれも同じ）
        $long = toLongPath $dir
        if ([System.IO.Directory]::GetDirectories($long).Length -gt 0) { return $false }
        if ([System.IO.Directory]::GetFiles($long, ${packFilePattern}).Length -gt 0) { return $false }
        return $true
    } catch {
        return $false
    }
}

function testIndexFolderHasSubfolders {
    # インデックスのフォルダの直下に、本のフォルダでないサブフォルダがあるか
    param (
        [string]$dir
    )

    try {
        foreach ($sub in [System.IO.Directory]::GetDirectories((toLongPath $dir))) {
            if (-not (testIndexBookDirPath $sub)) { return $true }
        }
        return $false
    } catch {
        return $false
    }
}

function testIndexFolderHasFiles {
    # インデックスのフォルダの直下に、ファイル（集約ファイル・集約する前の TSV）があるか
    param (
        [string]$dir
    )

    try {
        # 集約ファイル（content_index.<拡張子>.tsv。pack_format.ps1 の packFilePattern）か、集約する前の TSV があれば、フォルダ直下にファイルがある
        $long = toLongPath $dir
        if ([System.IO.Directory]::GetFiles($long, ${packFilePattern}).Length -gt 0) { return $true }
        foreach ($sub in [System.IO.Directory]::GetDirectories($long)) {
            if (-not (testIndexBookDirPath $sub)) { continue }
            if ([System.IO.Directory]::GetFiles($sub, "*.tsv").Length -gt 0) { return $true }
        }
        return $false
    } catch {
        return $false
    }
}

function getIndexFolderChildren {
    # インデックスのフォルダ（dir）の子を @{ HasFiles; Folders = @(@{ Name; HasSubfolders }); Error } で返す。
    # Folders は本のフォルダを除いた名前順。読めなかったときは Error に文面を入れ、Folders は空にする
    param (
        [string]$dir
    )

    $names = New-Object System.Collections.Generic.List[string]
    try {
        foreach ($sub in [System.IO.Directory]::GetDirectories((toLongPath $dir))) {
            if (testIndexBookDirPath $sub) { continue }
            $names.Add([System.IO.Path]::GetFileName($sub))
        }
    } catch {
        return @{ HasFiles = $false; Folders = @(); Error = $_.Exception.Message }
    }
    $names.Sort([System.StringComparer]::CurrentCultureIgnoreCase)
    $folders = @($names | ForEach-Object {
        @{ Name = $_; HasSubfolders = (testIndexFolderHasSubfolders (Join-Path $dir $_)) }
    })
    return @{ HasFiles = (testIndexFolderHasFiles $dir); Folders = $folders; Error = "" }
}

function getIndexTreeData {
    # 検索対象ツリーの材料を集める。
    #   dir: インデックスのフォルダ（work\index）  paths: 展開していたフォルダ・チェックを外したフォルダ（インデックスの根から先祖すべての子も返す）
    # 返すもの: @{ State（getPathState の State）; Message; Root; Sources（getSourceFolderMap の結果）;
    #   Indexes = @(@{ Name; SourcePath; Exists; HasSubfolders }); Children = @{ <フォルダ> = getIndexFolderChildren の結果 } }
    # インデックスのフォルダに接続できない・その他のときは、列挙せずに State と文面だけを返す
    # （Test-Path が $false を返して「インデックスが無い」と取り違えないため）
    param (
        [string]$dir,
        [string]$statusPath,
        [string]$settingsPath,
        [string[]]$paths = @()
    )

    $result = @{ State = ${pathStateMissing}; Message = ""; Root = $dir; Sources = @{}; Indexes = @(); Children = @{} }
    $state = getPathState $dir
    if ($state.State -ne ${pathStateFound}) {
        $result.State = $state.State
        $result.Message = $state.Message
        return $result
    }
    $result.State = ${pathStateFound}

    $data = getSearchIndexData $dir $statusPath $settingsPath
    $root = [string]$data.Root
    $result.Root = $root
    $result.Sources = $data.Sources
    $indexes = New-Object System.Collections.Generic.List[object]
    foreach ($index in @($data.Indexes)) {
        $path = "$root\$($index.Name)"
        $exists = [System.IO.Directory]::Exists((toLongPath $path))
        $indexes.Add(@{
            Name          = $index.Name
            SourcePath    = $(if ($index.SourcePath) { [string]$index.SourcePath } else { "" })
            Exists        = $exists
            HasSubfolders = ($exists -and (testIndexFolderHasSubfolders $path))
        })
    }
    $result.Indexes = $indexes.ToArray()

    # 先祖すべての子を入れる（深い階層の展開・除外を戻すため）
    foreach ($index in $indexes) {
        if (-not $index.Exists) { continue }
        $indexRoot = "$root\$($index.Name)"
        foreach ($target in @($paths)) {
            $target = ([string]$target).TrimEnd("\")
            $isRoot = $target.Equals($indexRoot, [System.StringComparison]::OrdinalIgnoreCase)
            if (-not $isRoot -and -not $target.StartsWith("$indexRoot\", [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            $segments = @($target.Substring($indexRoot.Length).Split([char[]]@("\"), [System.StringSplitOptions]::RemoveEmptyEntries))
            $current = $indexRoot
            for ($i = 0; $i -le $segments.Count; $i++) {
                if (-not $result.Children.ContainsKey($current)) {
                    $result.Children[$current] = getIndexFolderChildren $current
                }
                if ($i -ge $segments.Count) { break }
                $next = $segments[$i]
                $known = $result.Children[$current]
                if (-not (@($known.Folders) | Where-Object { $_.Name -ieq $next })) { break }
                $current = "$current\$next"
            }
        }
    }
    return $result
}
