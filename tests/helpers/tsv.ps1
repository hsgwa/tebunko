# テストで使う TSV の作成。

function newTsv {
    param (
        [string]$path,
        [string[]]$lines
    )

    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
    [System.IO.File]::WriteAllText($path, (($lines -join "`r`n") + "`r`n"), ${utf8Bom})
}

function newContentIndexFiles {
    # TSV のインデックス（tsvRoot）から、フォルダごとの本文インデックスのファイル（contentIndexRoot。同じ相対パス）を作り、findContentIndexFiles の結果を返す
    param ([string]$tsvRoot, [string]$contentIndexRoot)
    foreach ($folder in (findIndexFoldersWithBooks $tsvRoot)) {
        [void](convertFolderToContentIndex $folder ($contentIndexRoot + $folder.Substring($tsvRoot.Length)))
    }
    return , (findContentIndexFiles $contentIndexRoot)
}
