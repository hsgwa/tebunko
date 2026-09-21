# テストで使う TSV の作成。

function newTsv {
    param (
        [string]$path,
        [string[]]$lines
    )

    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
    [System.IO.File]::WriteAllText($path, (($lines -join "`r`n") + "`r`n"), ${utf8Bom})
}