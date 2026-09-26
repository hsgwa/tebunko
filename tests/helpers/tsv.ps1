# テストで使う TSV の作成。

function newTsv {
    param (
        [string]$path,
        [string[]]$lines
    )

    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
    [System.IO.File]::WriteAllText($path, (($lines -join "`r`n") + "`r`n"), ${utf8Bom})
}

function newPackIndex {
    # TSV のインデックス（tsvRoot）から、フォルダごとの集約ファイル（packRoot。同じ相対パス）を作り、getPackFiles の結果を返す
    param ([string]$tsvRoot, [string]$packRoot)
    foreach ($folder in (findIndexFoldersWithBooks $tsvRoot)) {
        [void](convertIndexFolderToPack $folder ($packRoot + $folder.Substring($tsvRoot.Length)))
    }
    return , (getPackFiles $packRoot)
}
