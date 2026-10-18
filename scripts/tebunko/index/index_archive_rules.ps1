# インデックスのエクスポート・インポートの判断（目録の確かめ・zip のエントリー名の安全性・名前の決め方）。
# 画面に触らず、ファイルにも触らない（判断層）。

# 目録（tebunko-index.json）の形式
${indexArchiveFormat}            = "tebunko-index"
${indexArchiveFormatVersion}     = 2  # この版が読める形式の版（インデックスの格納形式の版 ${indexFormatVersion} と同じ番号にそろえる。本文インデックスのファイルの形・zip の中の配置・目録・取り込み一覧の列を変えたら上げる）
${indexArchiveManifestFileName}  = "tebunko-index.json"
${indexArchiveStatusEntryName}   = "ingest_status.tsv"
${indexArchiveMaxManifestBytes}  = 64MB   # 目録（tebunko-index.json）自体の大きさの上限（zip bomb 対策）
${indexArchiveMaxEntries}        = 100000 # zip のエントリーの数の上限（目録自身を含む）
${indexArchiveMaxTotalBytes}     = 10GB   # 目録の files の size の合計の上限（空き容量を調べられないときも効く）
${indexArchiveMaxStatusBytes}    = 256MB  # 取り込み一覧（ingest_status.tsv）の大きさの上限（展開したあと全量を読むため）
${indexArchiveFreeSpaceMargin}   = 1GB    # インポートに求める空き容量の余裕（目録の合計 + この量）

# 同じ名前のインデックスがあったときの扱い（importIndex・getImportIndexName に渡す）
${importCollisionRename}    = "Rename"
${importCollisionOverwrite} = "Overwrite"
${importCollisionCancel}    = "Cancel"


# testPathSegments（区切りごとの確かめ）は index_name.ps1 にある（取り込み一覧の相対パスの確かめ testIndexRelPath と共通）


function testIndexArchiveEntryPath {
    # zip のエントリー名（/ 区切り）が、展開先のパスとして安全かを調べ、使えない理由を返す（使えれば空文字列）。
    # \ 区切り・絶対パス（/ 始まり・ドライブ文字・UNC）は、どれも展開先のパスの外に出るおそれがあるため受け付けない（zip slip 対策）
    param (
        [string]$path
    )

    if ($path -eq "") {
        return "エントリーの名前が空です。"
    }
    if ($path.IndexOf("\") -ge 0) {
        return "エントリーの名前に \ が含まれています（/ 区切りにしてください）: ${path}"
    }
    if ($path.StartsWith("/")) {
        return "エントリーの名前が / で始まっています: ${path}"
    }
    if ($path -match "^[A-Za-z]:") {
        return "エントリーの名前がドライブ文字で始まっています: ${path}"
    }
    $reason = testPathSegments ($path.Split("/"))
    if ($reason) {
        return "エントリー「${path}」の名前が使えません（${reason}）"
    }
    return ""
}


function testIndexArchiveEntryLocation {
    # zip のエントリー名（testIndexArchiveEntryPath を通ったもの）が、持ち出すファイルの置き場所として
    # 許されているか（ingest_status.tsv、または content_index/ の下の本文インデックスのファイルの名前の型）を返す
    param (
        [string]$path
    )

    if ($path -ceq ${indexArchiveStatusEntryName}) {
        return $true
    }
    if (!$path.StartsWith("content_index/", [System.StringComparison]::Ordinal)) {
        return $false
    }
    $name = $path.Substring("content_index/".Length)
    if ($name -eq "") {
        return $false
    }
    $lastSlash = $name.LastIndexOf("/")
    $fileName = if ($lastSlash -ge 0) { $name.Substring($lastSlash + 1) } else { $name }
    return ($null -ne (readContentIndexFileName $fileName))
}


function getManifestSourceFolder {
    # 目録の sourceFolder を返す。絶対パス・UNC でない、または制御文字を含むときは空（呼ぶ側に指定させる）
    param (
        $manifest
    )

    $folder = [string]$manifest.sourceFolder
    if ($folder -eq "") {
        return ""
    }
    if ($folder.IndexOfAny([char[]]@(0..31)) -ge 0) {
        return ""
    }
    if ($folder -notmatch "^[A-Za-z]:\\" -and !$folder.StartsWith("\\")) {
        return ""
    }
    return $folder
}


function testIndexArchiveManifest {
    # 目録（ConvertFrom-Json の結果。読めなければ $null）と、zip の中の全エントリー名（目録自身を含む）・
    # 目録エントリーの展開後の大きさから、この zip をインポートしてよいかを調べ、使えない理由を返す（使えれば空文字列）。
    # 1 つでも受け付けない点があればインポート全体を止めるため、最初に見つかった理由だけを返す
    param (
        $manifest,
        [string[]]$entryNames,
        [long]$manifestBytes
    )

    if ($manifestBytes -gt ${indexArchiveMaxManifestBytes}) {
        return "目録（${indexArchiveManifestFileName}）が大きすぎます。"
    }
    if (@($entryNames).Count -gt ${indexArchiveMaxEntries}) {
        return "zip のエントリーの数が多すぎます。"
    }
    if ($null -eq $manifest) {
        return "tebunko のインデックスの zip ではない、または壊れています（目録を読めません）。"
    }
    if ([string]$manifest.format -ne ${indexArchiveFormat}) {
        return "tebunko のインデックスの zip ではない、または壊れています（形式が違います）。"
    }
    $version = 0
    if (-not [int]::TryParse([string]$manifest.formatVersion, [ref]$version) -or $version -lt 1) {
        return "tebunko のインデックスの zip ではない、または壊れています（形式の版が読めません）。"
    }
    if ($version -lt ${indexArchiveFormatVersion}) {
        return "古い形式のインデックスの zip のため、読み込めません。${indexRebuildMessage}"
    }
    if ($version -gt ${indexArchiveFormatVersion}) {
        return "新しい版の tebunko で作られたインデックスのため、この版では読み込めません。"
    }

    $nameReason = testIndexName ([string]$manifest.indexName) @()
    if ($nameReason) {
        return "目録のインデックス名が使えません（${nameReason}）"
    }

    # zip の中に同じ名前のエントリーが 2 つあれば止める（ZipArchive は許すため、ここで数える）
    $seenEntries = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($entryName in $entryNames) {
        if (!$seenEntries.Add($entryName)) {
            return "zip の中に同じ名前のエントリーが 2 つあります: ${entryName}"
        }
    }

    $fileMap = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    $totalBytes = [long]0
    foreach ($file in @($manifest.files | Where-Object { $_ })) {
        $path = [string]$file.path
        $reason = testIndexArchiveEntryPath $path
        if ($reason) {
            return $reason
        }
        if (!(testIndexArchiveEntryLocation $path)) {
            return "目録に、許されない場所のエントリーがあります: ${path}"
        }
        if ($fileMap.ContainsKey($path)) {
            return "目録に同じパスが 2 つあります: ${path}"
        }
        $size = [long]0
        if (-not [long]::TryParse([string]$file.size, [ref]$size) -or $size -lt 0) {
            return "目録のファイルの大きさが読めません: ${path}"
        }
        if ([string]$file.sha256 -eq "") {
            return "目録のファイルの SHA-256 がありません: ${path}"
        }
        if ($path -ceq ${indexArchiveStatusEntryName} -and $size -gt ${indexArchiveMaxStatusBytes}) {
            return "${indexArchiveStatusEntryName} が大きすぎます。"
        }
        $totalBytes += $size
        if ($totalBytes -gt ${indexArchiveMaxTotalBytes}) {
            return "インデックスの合計の大きさが大きすぎます。"
        }
        $fileMap[$path] = $file
    }
    if (!$fileMap.ContainsKey(${indexArchiveStatusEntryName})) {
        return "目録に ${indexArchiveStatusEntryName} がありません。"
    }

    # zip の中のエントリー（目録自身を除く）が、すべて目録にあるか。目録にあるのに zip に無いものが無いか
    $zipEntries = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($entryName in $entryNames) {
        if ($entryName -ieq ${indexArchiveManifestFileName}) {
            continue
        }
        [void]$zipEntries.Add($entryName)
        if (!$fileMap.ContainsKey($entryName)) {
            return "目録に無いエントリーが zip にあります: ${entryName}"
        }
    }
    foreach ($path in $fileMap.Keys) {
        if (!$zipEntries.Contains($path)) {
            return "目録にあるのに zip に無いエントリーがあります: ${path}"
        }
    }

    return ""
}


function testImportedStatusRelPath {
    # ingest_status.tsv の1行の相対パス（\ 区切り。インデックス名は付いていない）が安全に使えるかを返す（使えれば空文字列）
    param (
        [string]$relPath
    )

    if ($relPath -eq "") {
        return "相対パスが空です。"
    }
    return (testPathSegments ($relPath.Split("\")))
}


function testImportedStatusRow {
    # ingest_status.tsv の1行（列は statusColumns と同じ）が受け付けられるかを返す（使えれば空文字列）
    param (
        [string]$line
    )

    $fields = $line.Split("`t")
    if ($fields.Count -ne ${statusColumns}.Count) {
        return "列の数が違います。"
    }
    if (@(${stateNew}, ${stateDone}, ${stateFailed}) -notcontains $fields[3]) {
        return "状態の値（$($fields[3])）が違います。"
    }
    return (testImportedStatusRelPath $fields[0])
}


function testImportedStatusLines {
    # ingest_status.tsv 全体（見出し行 + データ行）が受け付けられるかを返す（使えれば空文字列）。
    # 見出しが違う・列や状態や相対パスがおかしい行がある・同じ相対パスが2つある、のいずれかで止める
    param (
        [string[]]$lines
    )

    if ($lines.Count -eq 0 -or $lines[0] -ne (${statusColumns} -join "`t")) {
        return "ingest_status.tsv の見出しが違います。"
    }
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    for ($i = 1; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ($line -eq "") {
            continue
        }
        $reason = testImportedStatusRow $line
        if ($reason) {
            return "ingest_status.tsv の $($i + 1) 行目が読めません（${reason}）"
        }
        $relPath = $line.Split("`t")[0]
        if (!$seen.Add($relPath)) {
            return "ingest_status.tsv に同じ相対パスが 2 つあります: ${relPath}"
        }
    }
    return ""
}


function testImportedContentIndexFile {
    # 取り込む本文インデックスのファイル 1 つを、検索と更新が読めるかで確かめ、断る理由を返す（読めれば空文字列）。
    # 断るのは、大きさが contentIndexSearchMaxBytes を超えるもの（このときは中身を渡さなくてよい）と、
    # 検索（readContentIndexPlaces）か更新（splitContentIndexTextByBook）が例外にするもの。
    # 形の決まりはここに写さず、2 つの関数をそのまま呼ぶ（読み取りの側の決まりが変わっても、取り込みの判定が付いてくる）
    param (
        [long]$size,
        [AllowNull()]$text   # 文字列。 大きさだけを先に確かめるときは $null（読み取りの判定は飛ばす）
    )

    if ($size -gt ${contentIndexSearchMaxBytes}) {
        return "大きさが上限（$([long](${contentIndexSearchMaxBytes} / 1MB)) MB）を超えています"
    }
    if ($null -eq $text) {
        return ""
    }
    try {
        [void](readContentIndexPlaces $text)
        [void](splitContentIndexTextByBook $text)
    } catch {
        return [string]$_.Exception.Message
    }
    return ""
}


function getIndexFolderConflict {
    # 元のフォルダが、ほかのクロール対象のインデックスと重なるかを調べ、直してほしい内容を返す（問題なければ空文字列）。
    # 画面の追加と詳細欄の編集（testIndexEditInput）とインポート（importIndex）の両方が使う、1 つの決まり。
    # 同じフォルダは、getTargetFolders が 2 つ目以降を読まないため設定に残らない。入れ子のフォルダは、
    # 同じファイルが 2 つのインデックスに入り、取り込みも検索結果も二重になる
    param (
        [string]$folder,   # 正規化済みの元のフォルダ
        $others            # 比べる相手（Name・Path を持つ行。置き換える・編集中の行は呼び出し側で外しておく）
    )

    foreach ($other in @($others)) {
        if (testSameFolder $other.Path $folder) {
            return "「${folder}」のインデックス [$($other.Name)] が既にあります。"
        }
        if (testFolderUnder $folder $other.Path) {
            return "「${folder}」は、インデックス [$($other.Name)]（$($other.Path)）の中のフォルダです。" +
                "同じファイルが二重にインデックスに入るため、登録できません。検索する範囲を絞るときは［検索］の検索対象で外してください。"
        }
        if (testFolderUnder $other.Path $folder) {
            return "「${folder}」の中には、インデックス [$($other.Name)]（$($other.Path)）があります。" +
                "同じファイルが二重にインデックスに入るため、登録できません。まとめるときは、先に [$($other.Name)] を削除してください。"
        }
    }
    return ""
}


function getImportIndexName {
    # 同じ名前のインデックスがあったときの扱い（collisionMode）から、インポートで使う名前を決める。
    #   Rename    : usedNames と重ならない名前にする（今の newIndexName の決まり「名前(2)」「名前(3)」…）
    #   Overwrite : suggestedName のまま（そのインデックスを置き換える）
    #   Cancel    : 何もしない（$null）
    param (
        [string]$suggestedName,
        $usedNames = @(),
        [string]$collisionMode
    )

    if ($collisionMode -eq ${importCollisionCancel}) {
        return $null
    }
    if ($collisionMode -eq ${importCollisionOverwrite}) {
        return $suggestedName
    }
    return (newIndexName $suggestedName $usedNames)
}


function getExportFileName {
    # エクスポートの既定のファイル名（"<インデックス名>_インデックス_<yyyyMMdd>.zip"）を返す。
    # usedNames（保存先フォルダに既にあるファイル名）と重なれば "(2)" "(3)" … を付ける
    param (
        [string]$indexName,
        $usedNames = @(),
        [datetime]$now = (Get-Date)
    )

    $base = "$(toSafeFileName $indexName)_インデックス_$($now.ToString('yyyyMMdd'))"
    $used = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($usedName in @($usedNames)) {
        if ($usedName) {
            [void]$used.Add([string]$usedName)
        }
    }
    $name = "${base}.zip"
    for ($i = 2; $used.Contains($name); $i++) {
        $name = "${base}(${i}).zip"
    }
    return $name
}
