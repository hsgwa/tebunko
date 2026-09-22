# setting.config の読み書き（tebunko_grep の設定）。

# 設定ファイル（画面が読み書きする。インデクサはクロール対象フォルダを読む）。内容は JSON
${settingsFile} = "${rootDir}\setting.config"
# 以前の設定ファイル（設定ファイルと同じフォルダの config\*.txt）。設定ファイルが無いときだけ読み込んで移す
${legacyConfigDirName}        = "config"
${legacyTargetFolderFileName} = "変換対象フォルダパス.txt"
${legacySourceReplaceFileName} = "元のフォルダの置き換え.txt"
${legacySearchOptionFileName} = "検索オプション.txt"

function newSettings {
    # 設定の既定値。設定ファイル（JSON）のキーと同じ
    return [ordered]@{
        targetFolders      = @()      # クロール対象フォルダ: @{ name（インデックス名）; path（今フォルダが置かれている場所）; enabled }（記載順。enabled が false は登録のみで取り込まない）
        indexSources       = @()      # 取り込まないインデックスの元のフォルダ: @{ name; path }（別のPC・場所で作ったインデックスを検索するとき）
        searchExcludes     = @()      # 画面の検索対象ツリーでチェックを外したフォルダ: @{ path（フルパス）; subfolders（false はフォルダ直下のファイルだけ） }
        useRegex           = $false   # 検索ワードを正規表現として扱う
        caseSensitive      = $false   # 英字の大文字と小文字を区別する
        fileFilter         = ""       # 対象ファイル（元のファイル名のワイルドカード。; 区切り、! で始まるものは除外。空ならすべて）
        includeShapes      = $true    # 図形（テキストボックス等）の文字も検索する（場所 "<元の場所>[図形]"。index_name.ps1 の objectPlacePattern）
        includeComments    = $true    # コメントも検索する（場所 "<元の場所>[コメント]"）
        openMode           = ${openModeNormal}  # 検索結果の元のファイルの開き方: 通常（編集する）/ 読み取り専用 / 新規（元のファイルを基にした無題の文書。占有しない）
    }
}

function readSettings {
    # 設定を newSettings と同じ形で返す。記載の無い項目は既定値。
    # 設定ファイルが無ければ、同じフォルダの config\ にある以前の設定ファイル（*.txt）から移して保存する（それも無ければ既定値）
    param (
        [string]$path = ${settingsFile}
    )

    $settings = newSettings
    if (!(Test-Path -LiteralPath $path)) {
        if (readLegacySettings $settings (Join-Path (getPathParent $path) ${legacyConfigDirName})) {
            writeSettings $settings $path
        }
        return $settings
    }

    # 制限モード（制限言語モード）からも読むため、System.IO.File ではなく Get-Content で読む（BOM の有無によらず UTF-8）
    # 空のファイル（BOM だけを含む）では何も返らない
    $json = Get-Content -LiteralPath $path -Raw -Encoding UTF8 -ErrorAction Stop
    if ($null -eq $json -or $json.Trim() -eq "") {
        return $settings
    }
    try {
        $data = ConvertFrom-Json $json
    } catch {
        throw "$(getPathLeaf $path) を読み込めません。（$($_.Exception.Message)）"
    }
    # 中身が null だけなら、空のファイルと同じく既定値（配列・数値だけのときと同じ扱い）
    if ($null -eq $data) {
        return $settings
    }
    foreach ($key in @($settings.Keys)) {
        $property = $data.PSObject.Properties[$key]
        if ($null -eq $property -or $null -eq $property.Value) {
            continue
        }
        if ($settings[$key] -is [array]) {
            $settings[$key] = @($property.Value | Where-Object { $null -ne $_ })
        } elseif ($settings[$key] -is [string]) {
            $settings[$key] = [string]$property.Value
        } else {
            $settings[$key] = toSettingBool $property.Value $settings[$key]
        }
    }
    return $settings
}

function toSettingBool {
    # 設定ファイルの真偽値を読む。手で書き換えた "false" などの文字列は [bool] にすると $true になるため、
    # 文字列は true / false として読み、読めなければ既定値（default）にする
    param (
        $value,
        [bool]$default
    )

    if ($value -is [string]) {
        # [bool]::TryParse と同じく、前後の空白を除いた true / false を大文字・小文字を区別せずに読む
        # （TryParse は [ref] を渡すため、制限言語モードでは使えない）
        switch ($value.Trim()) {
            "true" { return $true }
            "false" { return $false }
        }
        return $default
    }
    return [bool]$value
}

function writeSettings {
    param (
        $settings,
        [string]$path = ${settingsFile}
    )

    New-Item -ItemType Directory -Path (getPathParent $path) -Force | Out-Null
    writeUtf8NoBom $path (ConvertTo-Json -InputObject $settings -Depth 5)
}

function updateSettings {
    # 設定ファイルを読み直して key の値だけを変えて保存する（ほかの項目は、ほかの画面・処理が保存した内容を保つ）
    param (
        [string]$key,
        $value,
        [string]$path = ${settingsFile}
    )

    $settings = readSettings $path
    $settings[$key] = $value
    writeSettings $settings $path
}

function readLegacySettings {
    # 以前の設定ファイル（1行1件のテキスト）を settings に読み込む。1つも無ければ $false
    param (
        $settings,
        [string]$dir
    )

    $found = $false
    $file = Join-Path $dir ${legacyTargetFolderFileName}
    if (Test-Path -LiteralPath $file) {
        $found = $true
        # 行頭が # の行はチェックなし
        # インデックス名は以前の設定ファイルには無いため空にする（取り込み時に割り当てる。assignIndexNames）
        $settings.targetFolders = @(readListFile $file | ForEach-Object { $_.Trim() } | ForEach-Object {
            New-Object PSObject -Property ([ordered]@{ name = ""; path = (normalizeFolderPath $_.TrimStart("#")); enabled = -not $_.StartsWith("#") })
        } | Where-Object { $_.path -ne "" })
    }
    $file = Join-Path $dir ${legacySearchOptionFileName}
    if (Test-Path -LiteralPath $file) {
        $found = $true
        # "正規表現=オン" の行
        $settings.useRegex = @(readListFile $file | Where-Object { $_ -match "^\s*正規表現\s*=\s*オン\s*$" }).Count -gt 0
    }
    return $found
}

function getTargetFolders {
    # クロール対象フォルダの一覧（記載順）を返す: @{ Name; Path; Enabled }。
    #   Name   : インデックス名（work\index 直下のフォルダ名）。インデックスの「名前」で、フォルダの置き場所（Path）とは分けて持つ。
    #            Path を書き換えても Name が同じなら同じインデックスとして扱う（取り込み直さない）。空なら取り込み時に割り当てる（assignIndexNames）
    #   Path   : そのフォルダが今置かれている場所
    #   Enabled: false はチェックなし（登録のみで取り込まない）
    # 同じフォルダ・同じ名前は最初のものだけ使う（名前の重複は、2 つ目以降を空にして割り当て直す）。
    # 書き方が違うだけで同じフォルダを指す場合（ネットワークドライブと UNC パスなど）も同じフォルダとみなす
    param (
        [string]$path = ${settingsFile}
    )

    # 制限モード（制限言語モード）からも読むため、List・HashSet ではなく配列と、
    # ToUpperInvariant したものをキーにしたハッシュテーブル（大文字・小文字を区別しない集合）で持つ
    $folders = @()
    $seenPath = @{}
    $seenName = @{}
    foreach ($item in @((readSettings $path).targetFolders)) {
        $folder = normalizeFolderPath ([string]$item.path)
        if ($folder -eq "" -or $seenPath.ContainsKey($folder.ToUpperInvariant())) {
            continue
        }
        $seenPath[$folder.ToUpperInvariant()] = $true
        if (@($folders | Where-Object { testSameFolder $_.Path $folder }).Count -gt 0) {
            continue
        }
        $name = toSafeFileName ([string]$item.name).Trim()
        if ($name -ne "") {
            if ($seenName.ContainsKey($name.ToUpperInvariant())) {
                $name = ""
            } else {
                $seenName[$name.ToUpperInvariant()] = $true
            }
        }
        $folders += New-Object PSObject -Property ([ordered]@{ Name = $name; Path = $folder; Enabled = ($item.enabled -ne $false) })
    }
    return $folders
}

function writeTargetFolders {
    # クロール対象フォルダの一覧（@{ Name; Path; Enabled } の配列）を保存する
    param (
        [object[]]$folders,
        [string]$path = ${settingsFile}
    )

    updateSettings "targetFolders" ([object[]]@($folders | Where-Object { $_ } | ForEach-Object {
        New-Object PSObject -Property ([ordered]@{ name = [string]$_.Name; path = $_.Path; enabled = [bool]$_.Enabled })
    })) $path
}

function readIndexSources {
    # 取り込まないインデックスの元のフォルダ（indexSources）を @{ Name; Path } の配列で返す。
    # 別の PC・場所で作ったインデックスを検索するとき、そのインデックス名の元のフォルダを覚えておくために使う
    param (
        [string]$path = ${settingsFile}
    )

    $items = @()
    $seen = @{}
    foreach ($item in @((readSettings $path).indexSources)) {
        $name = ([string]$item.name).Trim()
        $folder = normalizeFolderPath ([string]$item.path)
        if ($name -eq "" -or $folder -eq "" -or $seen.ContainsKey($name.ToUpperInvariant())) {
            continue
        }
        $seen[$name.ToUpperInvariant()] = $true
        $items += New-Object PSObject -Property ([ordered]@{ Name = $name; Path = $folder })
    }
    return $items
}

function writeIndexSources {
    param (
        [object[]]$sources,
        [string]$path = ${settingsFile}
    )

    updateSettings "indexSources" ([object[]]@($sources | Where-Object { $_ } | ForEach-Object {
        New-Object PSObject -Property ([ordered]@{ name = $_.Name; path = $_.Path })
    })) $path
}

function setIndexSourceFolder {
    # インデックス名に対する元のフォルダ（今そのフォルダが置かれている場所）を設定に記録する。
    # クロール対象フォルダにある名前ならそのパスを書き換え、無ければ indexSources に記録する。
    # 「名前」と「置き場所」を分けて持つため、フォルダを移した場合もこの 1 か所を書き換えるだけで済む
    param (
        [string]$name,
        [string]$folder,
        [string]$path = ${settingsFile}
    )

    $name = ([string]$name).Trim()
    $folder = normalizeFolderPath $folder
    if ($name -eq "" -or $folder -eq "") {
        return
    }

    $targets = @(getTargetFolders $path)
    if (@($targets | Where-Object { $_.Name -eq $name }).Count -gt 0) {
        writeTargetFolders @($targets | ForEach-Object {
            if ($_.Name -eq $name) { New-Object PSObject -Property ([ordered]@{ Name = $_.Name; Path = $folder; Enabled = $_.Enabled }) } else { $_ }
        }) $path
        return
    }

    $sources = @(@(readIndexSources $path | Where-Object { $_.Name -ne $name }) + @(New-Object PSObject -Property ([ordered]@{ Name = $name; Path = $folder })))
    writeIndexSources $sources $path
}

function readSearchExcludes {
    # 画面の検索対象ツリーでチェックを外したフォルダを @{ Path; Subfolders } の配列で返す（設定が無ければ空 = すべて検索する）。
    #   Subfolders: $true はフォルダ以下すべて、$false はフォルダ直下のファイルだけを外す
    param (
        [string]$path = ${settingsFile}
    )

    $result = @()
    foreach ($item in @((readSettings $path).searchExcludes)) {
        $folder = ([string]$item.path).Trim().TrimEnd("\")
        if ($folder -eq "") {
            continue
        }
        $subfolders = if ($null -eq $item.subfolders) { $true } else { toSettingBool $item.subfolders $true }
        $result += New-Object PSObject -Property ([ordered]@{ Path = $folder; Subfolders = $subfolders })
    }
    return $result
}

function writeSearchExcludes {
    # 検索対象ツリーでチェックを外したフォルダ（@{ Path; Subfolders } の配列）を保存する
    param (
        [object[]]$excludes,
        [string]$path = ${settingsFile}
    )

    updateSettings "searchExcludes" ([object[]]@($excludes | Where-Object { $_ -and $_.Path } | ForEach-Object {
        [ordered]@{ path = [string]$_.Path; subfolders = [bool]$_.Subfolders }
    })) $path
}


${searchOptionKeys} = [ordered]@{
    UseRegex = "useRegex"; CaseSensitive = "caseSensitive"; FileFilter = "fileFilter"
    IncludeShapes = "includeShapes"; IncludeComments = "includeComments"
}


function readSearchOption {
    # 画面の検索オプションを @{ UseRegex; CaseSensitive; FileFilter; IncludeShapes; IncludeComments } で返す。
    # 設定が無ければ、文字どおり・大文字と小文字を区別しない・対象ファイルはすべて・図形とコメントも検索する
    param (
        [string]$path = ${settingsFile}
    )

    $settings = readSettings $path
    $option = @{}
    foreach ($name in ${searchOptionKeys}.Keys) {
        $option[$name] = $settings[${searchOptionKeys}[$name]]
    }
    return $option
}

function writeSearchOption {
    # 画面の検索オプションを保存する。option（readSearchOption と同じ形）にある項目だけを変える
    param (
        [hashtable]$option,
        [string]$path = ${settingsFile}
    )

    $settings = readSettings $path
    foreach ($name in ${searchOptionKeys}.Keys) {
        if ($option.ContainsKey($name)) {
            $settings[${searchOptionKeys}[$name]] = $option[$name]
        }
    }
    writeSettings $settings $path
}

function readOpenMode {
    # 元のファイルの開き方（${openModes} のいずれか）を返す。設定が無い・知らない値なら「通常」
    param (
        [string]$path = ${settingsFile}
    )

    $mode = (readSettings $path).openMode
    if (${openModes} -contains $mode) {
        return $mode
    }
    return ${openModeNormal}
}

function writeOpenMode {
    param (
        [string]$mode,
        [string]$path = ${settingsFile}
    )

    updateSettings "openMode" $mode $path
}
