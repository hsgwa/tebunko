# 検索対象ツリーでチェックを外したフォルダ（設定の searchExcludes）の読み書きと、「フォルダとその下」の比較。
# インデックスの改名・削除（IndexCatalog）、ワークスペースの切り替え（WorkspaceMover）、検索対象ツリーの保存が、同じ部品を通す。
# 設定の読み書きは settings.ps1（invokeSettingsLocked・readSettings・updateSettings）を使う。

function testFolderOrUnder {
    # path が folder そのもの、または folder の下にあるかを返す（フォルダとその下の比較は、ここだけで行う）。
    # 大文字・小文字は区別しない（序数の比較。NTFS と同じ扱いで、文化による同一視はしない）。
    # folder の末尾の \ は無いものとして比べ、"<folder>\" で始まるものだけを「下」にする（Sales の下に Sales2 は入らない）。
    # folder が空なら、どれも含まない
    param (
        [string]$path,
        [string]$folder
    )

    $folder = ([string]$folder).TrimEnd("\")
    if ($folder -eq "" -or $null -eq $path) {
        return $false
    }
    return $path.Equals($folder, [System.StringComparison]::OrdinalIgnoreCase) -or
        $path.StartsWith($folder + "\", [System.StringComparison]::OrdinalIgnoreCase)
}

function findFolderOrAncestor {
    # path 自身か、その先祖のうち、folders（HashSet[string]。OrdinalIgnoreCase・末尾の \ なし）に入っている最初のものを返す（無ければ $null）。
    # 結果は testFolderOrUnder と同じ（返すのは folders の表記ではなく、path 側の表記の先祖。大文字・小文字は区別しない・\ の区切りでだけ上る）。多くのフォルダを 1 件ずつ比べる代わりに、先祖を辞書引きする
    param (
        [string]$path,
        $folders
    )

    if ($null -eq $path) {
        return $null
    }
    $current = $path.TrimEnd("\")
    while ($current -ne "") {
        if ($folders.Contains($current)) {
            return $current
        }
        $index = $current.LastIndexOf("\")
        if ($index -le 0) {
            return $null
        }
        $current = $current.Substring(0, $index)
    }
    return $null
}

function groupByFolderOrAncestor {
    # items（文字列のパス。property を渡せば、その名前の値をパスとする）を、folders（findFolderOrAncestor と同じ集合）の
    # どのフォルダの下かで 1 回に振り分ける。戻り値は 見つかったフォルダ → items の List の辞書（OrdinalIgnoreCase）。
    # 辞書のキーは findFolderOrAncestor の戻り値のため、folders の表記ではなく path 側の表記になる（大文字・小文字が違うことがある）。
    # どのフォルダの下にも無いものは入らない
    param (
        $items,
        $folders,
        [string]$property = ""
    )

    $groups = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($item in @($items)) {
        $path = if ($property) { [string]$item.$property } else { [string]$item }
        $key = findFolderOrAncestor $path $folders
        if ($null -eq $key) { continue }
        if (-not $groups.ContainsKey($key)) { $groups[$key] = New-Object System.Collections.Generic.List[object] }
        $groups[$key].Add($item)
    }
    return $groups
}

function testSameSearchExcludes {
    # 2 つの除外の一覧（@{ Path; Subfolders } の配列）が、件数・順・Path（大文字・小文字も含めて）・Subfolders まで同じかを返す
    # （順だけ違っても同じとは見なさない。保存を省いてよいかの判定に使う）
    param (
        [object[]]$left,
        [object[]]$right
    )

    $left = @($left)
    $right = @($right)
    if ($left.Count -ne $right.Count) {
        return $false
    }
    for ($i = 0; $i -lt $left.Count; $i++) {
        if (-not ([string]$left[$i].Path).Equals([string]$right[$i].Path, [System.StringComparison]::Ordinal) -or
            [bool]$left[$i].Subfolders -ne [bool]$right[$i].Subfolders) {
            return $false
        }
    }
    return $true
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

function removeSearchExcludesUnder {
    # 検索対象ツリーでチェックを外した記録のうち、folder（インデックスのフォルダ）とその下のものを消す（比較は testFolderOrUnder）。
    # 記録は利便性のための一時的な情報のため、消せなくても例外にしない（インデックスの改名・削除は止めない）。
    # 消す記録が無ければ設定ファイルは書き換えない
    param (
        [string]$folder,
        [string]$path = ${settingsFile}
    )

    $folder = ([string]$folder).Trim().TrimEnd("\")
    if ($folder -eq "") {
        return
    }

    try {
        invokeSettingsLocked -path $path -action {
            $excludes = @(readSearchExcludes $path)
            $kept = @($excludes | Where-Object { -not (testFolderOrUnder $_.Path $folder) })
            if ($kept.Count -lt $excludes.Count) {
                writeSearchExcludes $kept $path
            }
        } | Out-Null
    } catch {
    }
}

function moveSearchExcludes {
    # 検索対象ツリーでチェックを外したフォルダ（searchExcludes。インデックスの下のフルパス）のうち、
    # ワークスペース from の index の下のものを、to の index の下に付け替えて保存する。付け替えた数を返す
    param (
        [string]$from,
        [string]$to,
        [string]$path = ${settingsFile}
    )

    $fromIndex = ([Workspace]::new($from.TrimEnd("\"))).IndexDir
    $toIndex = ([Workspace]::new($to.TrimEnd("\"))).IndexDir
    return invokeSettingsLocked -path $path -action {
        $count = 0
        $excludes = @(readSearchExcludes $path | ForEach-Object {
            $folder = $_.Path
            if (testFolderOrUnder $folder $fromIndex) {
                $folder = $toIndex + $folder.Substring($fromIndex.Length)
                $count++
            }
            [pscustomobject]@{ Path = $folder; Subfolders = $_.Subfolders }
        })
        if ($count -gt 0) {
            writeSearchExcludes $excludes $path
        }
        return $count
    }
}
