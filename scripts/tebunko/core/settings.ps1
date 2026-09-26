# setting.config の読み書き（tebunko の設定）。

# 設定ファイル（画面が読み書きする。インデクサはクロール対象フォルダと work の置き場所を読む）。内容は JSON。
# ツールのフォルダに書き込めないときは、利用者ごとの場所に置く（shared\core\data_dir.ps1 の getDataDir）
${settingsFile} = "${dataDir}\setting.config"

# 検索結果から元のファイルを開くときの開き方（設定 openMode の値）
${openModeNormal}   = "normal"    # そのまま開く（編集する）
${openModeReadOnly} = "readOnly"  # 読み取り専用で開く（誤って上書きしない）
${openModeNew}      = "new"       # 新規（元のファイルを基にした無題の文書）で開く。元のファイルを占有しない
${openModes}        = @(${openModeNormal}, ${openModeReadOnly}, ${openModeNew})

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
        workspaceFolder    = ""       # ワークスペース（インデックス・取り込み一覧・ログを置くフォルダ）。空なら既定（%USERPROFILE%\Documents\tebunko_ws。getDefaultWorkDir）
        ingestThreads      = 0        # Office を使わずに読むファイル（.docx・.pptx など）の読み取りのスレッドの数（1〜4。0 はコア数から決める。getIngestWorkerCount）。Excel・Word・PowerPoint は種類ごとに 1 つ
    }
}

function readSettings {
    # 設定を newSettings と同じ形で返す。記載の無い項目は既定値。設定ファイルが無ければ既定値
    param (
        [string]$path = ${settingsFile}
    )

    $settings = newSettings
    if (!(Test-Path -LiteralPath $path)) {
        return $settings
    }

    $json = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
    if ($json.Trim() -eq "") {
        return $settings
    }
    try {
        $data = ConvertFrom-Json $json
    } catch {
        throw "$([System.IO.Path]::GetFileName($path)) を読み込めません。（$($_.Exception.Message)）"
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
        } elseif ($settings[$key] -is [int]) {
            # 数値。手で書き換えて数値にできないときは既定値
            $number = 0
            if ([int]::TryParse([string]$property.Value, [ref]$number)) {
                $settings[$key] = $number
            }
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
        $parsed = $false
        if ([bool]::TryParse($value.Trim(), [ref]$parsed)) {
            return $parsed
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

    # 書き込みの途中で止まっても設定ファイルが壊れないよう、一時ファイルに書いてから置き換える（BOM なし UTF-8）
    writeTextLinesAtomic $path @((ConvertTo-Json -InputObject $settings -Depth 5)) (New-Object System.Text.UTF8Encoding($false))
}

function invokeSettingsLocked {
    # 設定ファイルの「読む → 変える → 書く」の一続きを、ほかの画面・インデクサ（別のプロセスも）と同時に行わないよう、
    # 設定ファイルごとの名前付きミューテックスの中で action を実行する（同じスレッドの入れ子は通す）。
    # action の出力を返す。5 秒待っても取れなければ例外（invokeWithNamedMutex）
    param (
        [string]$path,
        [scriptblock]$action,
        [int]$timeoutMilliseconds = 5000
    )

    $lockKey = getFolderKey ([System.IO.Path]::GetFullPath($path))
    return invokeWithNamedMutex "Local\${appId}_settings_${lockKey}" $timeoutMilliseconds $action
}

function updateSettings {
    # 設定ファイルを読み直して key の値だけを変えて保存する（ほかの項目は、ほかの画面・処理が保存した内容を保つ）
    param (
        [string]$key,
        $value,
        [string]$path = ${settingsFile}
    )

    invokeSettingsLocked $path {
        $settings = readSettings $path
        $settings[$key] = $value
        writeSettings $settings $path
    }
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

    $folders = New-Object System.Collections.Generic.List[object]
    $seenPath = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $seenName = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($item in @((readSettings $path).targetFolders)) {
        $folder = normalizeFolderPath ([string]$item.path)
        if ($folder -eq "" -or -not $seenPath.Add($folder)) {
            continue
        }
        if (@($folders | Where-Object { testSameFolder $_.Path $folder }).Count -gt 0) {
            continue
        }
        $name = toSafeFileName ([string]$item.name).Trim()
        if ($name -ne "" -and -not $seenName.Add($name)) {
            $name = ""
        }
        $folders.Add([pscustomobject]@{ Name = $name; Path = $folder; Enabled = ($item.enabled -ne $false) })
    }
    return $folders.ToArray()
}

function writeTargetFolders {
    # クロール対象フォルダの一覧（@{ Name; Path; Enabled } の配列）を保存する
    param (
        [object[]]$folders,
        [string]$path = ${settingsFile}
    )

    updateSettings "targetFolders" ([object[]]@($folders | Where-Object { $_ } | ForEach-Object {
        [pscustomobject]@{ name = [string]$_.Name; path = $_.Path; enabled = [bool]$_.Enabled }
    })) $path
}

function mergeAssignedIndexNames {
    # 今の設定のクロール対象フォルダ（current。getTargetFolders の形）に、インデクサ・画面が割り当てたインデックス名（assigned。
    # assignIndexNames の結果）を反映した一覧を返す。インデクサが始めに読んだ一覧を丸ごと書き戻すと、その間の画面での変更を消すため、
    # 書き戻す直前に読み直した一覧へ、割り当てた名前だけを足す。
    # 突き合わせる鍵はパス（大文字と小文字は区別しない）。名前は、今の一覧で名前が空の項目にだけ付ける。
    # 同じ名前をほかの項目が使っていれば付けない（空のまま。次の機会に割り当て直す）
    param (
        [object[]]$current,
        [object[]]$assigned
    )

    $assignedNames = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($item in @($assigned | Where-Object { $_ -and $_.Name })) {
        $assignedNames[[string]$item.Path] = [string]$item.Name
    }
    $used = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($item in @($current | Where-Object { $_ -and $_.Name })) {
        [void]$used.Add([string]$item.Name)
    }

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($item in @($current | Where-Object { $_ })) {
        $name = [string]$item.Name
        if ($name -eq "" -and $assignedNames.ContainsKey([string]$item.Path) -and $used.Add($assignedNames[[string]$item.Path])) {
            $name = $assignedNames[[string]$item.Path]
        }
        $result.Add([pscustomobject]@{ Name = $name; Path = $item.Path; Enabled = $item.Enabled })
    }
    return $result.ToArray()
}

function saveAssignedIndexNames {
    # 割り当てたインデックス名（assigned。assignIndexNames の結果）を設定に保存する。
    # 設定を読み直し（ほかの画面・インデクサの変更を保つ）、mergeAssignedIndexNames で名前だけを足して書く（読む → 書くを排他の中で行う）
    param (
        [object[]]$assigned,
        [string]$path = ${settingsFile}
    )

    invokeSettingsLocked $path {
        writeTargetFolders (mergeAssignedIndexNames @(getTargetFolders $path) $assigned) $path
    }
}

function readIndexSources {
    # 取り込まないインデックスの元のフォルダ（indexSources）を @{ Name; Path } の配列で返す。
    # 別の PC・場所で作ったインデックスを検索するとき、そのインデックス名の元のフォルダを覚えておくために使う
    param (
        [string]$path = ${settingsFile}
    )

    $items = New-Object System.Collections.Generic.List[object]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($item in @((readSettings $path).indexSources)) {
        $name = ([string]$item.name).Trim()
        $folder = normalizeFolderPath ([string]$item.path)
        if ($name -eq "" -or $folder -eq "" -or -not $seen.Add($name)) {
            continue
        }
        $items.Add([pscustomobject]@{ Name = $name; Path = $folder })
    }
    return $items.ToArray()
}

function writeIndexSources {
    param (
        [object[]]$sources,
        [string]$path = ${settingsFile}
    )

    updateSettings "indexSources" ([object[]]@($sources | Where-Object { $_ } | ForEach-Object {
        [pscustomobject]@{ name = $_.Name; path = $_.Path }
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

    invokeSettingsLocked $path {
        $targets = @(getTargetFolders $path)
        if (@($targets | Where-Object { $_.Name -eq $name }).Count -gt 0) {
            writeTargetFolders @($targets | ForEach-Object {
                if ($_.Name -eq $name) { [pscustomobject]@{ Name = $_.Name; Path = $folder; Enabled = $_.Enabled } } else { $_ }
            }) $path
            return
        }

        $sources = @(@(readIndexSources $path | Where-Object { $_.Name -ne $name }) + @([pscustomobject]@{ Name = $name; Path = $folder }))
        writeIndexSources $sources $path
    }
}

function readSearchExcludes {
    # 画面の検索対象ツリーでチェックを外したフォルダを @{ Path; Subfolders } の配列で返す（設定が無ければ空 = すべて検索する）。
    #   Subfolders: $true はフォルダ以下すべて、$false はフォルダ直下のファイルだけを外す
    param (
        [string]$path = ${settingsFile}
    )

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($item in @((readSettings $path).searchExcludes)) {
        $folder = ([string]$item.path).Trim().TrimEnd("\")
        if ($folder -eq "") {
            continue
        }
        $subfolders = if ($null -eq $item.subfolders) { $true } else { toSettingBool $item.subfolders $true }
        $result.Add([pscustomobject]@{ Path = $folder; Subfolders = $subfolders })
    }
    return $result.ToArray()
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

    invokeSettingsLocked $path {
        $settings = readSettings $path
        foreach ($name in ${searchOptionKeys}.Keys) {
            if ($option.ContainsKey($name)) {
                $settings[${searchOptionKeys}[$name]] = $option[$name]
            }
        }
        writeSettings $settings $path
    }
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

function getDefaultWorkDir {
    # 既定のワークスペース（%USERPROFILE%\Documents\tebunko_ws）。高速検索のため、Windows Search の索引の対象になる場所に置く。
    # OneDrive にリダイレクトされた「ドキュメント」ではなく、プロファイルの直下の Documents を使う（インデックスが同期でクラウドに上がらないように）
    param (
        [string]$profileDir = [System.Environment]::GetFolderPath("UserProfile")
    )

    return Join-Path $profileDir "Documents\tebunko_ws"
}

# 既定のワークスペースが空でないときの文言（画面の［既定に戻す］・起動時、インデクサで共通）
function getDefaultWorkspaceError {
    param (
        [string]$folder
    )

    return "「${folder}」は空のフォルダではありません。ワークスペースには別の空のフォルダを選んでください（［8 設定］の［変更…］）。"
}

function testDefaultWorkspace {
    # 既定のワークスペースを使えるか: @{ Usable; Folder; Message }。
    # 使える: 無い（使うときに作る）・空・前から使っているワークスペース（index か 取り込み一覧.tsv がある）。
    # それ以外（ほかのファイルが置いてある）は、インデックスのファイルと混ざるため使わせない
    param (
        [string]$folder = (getDefaultWorkDir)
    )

    $result = @{ Usable = $true; Folder = $folder; Message = "" }
    if (![System.IO.Directory]::Exists($folder)) {
        return $result
    }
    $names = @([System.IO.Directory]::EnumerateFileSystemEntries($folder) | Select-Object -First 1000 | ForEach-Object { [System.IO.Path]::GetFileName($_) })
    if ($names.Count -eq 0 -or ($names -contains "index") -or ($names -contains [System.IO.Path]::GetFileName($workspace.StatusFile))) {
        return $result
    }
    $result.Usable = $false
    $result.Message = getDefaultWorkspaceError $folder
    return $result
}

function getWorkspaceBlockMessage {
    # 今のワークスペースが既定の場所で、そこにほかのファイルが置いてあるなら、その文言（使えるなら空）。
    # インデックスのファイルと混ざるため、インデックス作成を始めず、［8 設定］で別のフォルダを選んでもらう
    param (
        [string]$current = $workspace.Dir,
        [string]$defaultDir = (getDefaultWorkDir)
    )

    if (!(testSameFolder $current $defaultDir)) {
        return ""
    }
    return (testDefaultWorkspace $defaultDir).Message
}

function getWorkDir {
    # work の置き場所を返す。設定 workspaceFolder が空なら既定（getDefaultWorkDir）。
    # 手で書いた相対パスは、設定ファイルのフォルダからとみなす
    param (
        [string]$path = ${settingsFile}
    )

    $folder = ([string](readSettings $path).workspaceFolder).Trim()
    if ($folder -eq "") {
        return getDefaultWorkDir
    }
    $folder = [System.Environment]::ExpandEnvironmentVariables($folder)
    return [System.IO.Path]::GetFullPath([System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($path), $folder)).TrimEnd("\")
}

function writeWorkspaceFolder {
    # work の置き場所を保存する。既定の場所なら空にする（ツールのフォルダを移しても既定のまま付いてくるように）
    param (
        [string]$folder,
        [string]$path = ${settingsFile}
    )

    $folder = ([string]$folder).Trim().TrimEnd("\")
    if ($folder -ne "" -and (testSameFolder $folder (getDefaultWorkDir))) {
        $folder = ""
    }
    updateSettings "workspaceFolder" $folder $path
}
