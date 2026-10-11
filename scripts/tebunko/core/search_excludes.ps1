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
