# 前の版のファイル（見本。tests/testdata/compat/index/）を読むテストの共通の部品。
#   tests/tebunko/indexer/index_compat.Tests.ps1・tests/meta/compat.Tests.ps1 が dot-source する。

function getZipEntryNames {
    # zip の中のファイル名（FullName）を並べ替えて返す
    param ([string]$zipPath)
    $archive = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
    try {
        return @($archive.Entries | ForEach-Object { $_.FullName } | Sort-Object)
    } finally {
        $archive.Dispose()
    }
}

function readZipEntryText {
    # zip の中のファイル（UTF-8）の中身を返す。無ければ $null
    param ([string]$zipPath, [string]$entryName)
    $archive = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
    try {
        $entry = $archive.Entries | Where-Object { $_.FullName -ceq $entryName } | Select-Object -First 1
        if (!$entry) { return $null }
        $stream = $entry.Open()
        try {
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
            return $reader.ReadToEnd()
        } finally {
            $stream.Dispose()
        }
    } finally {
        $archive.Dispose()
    }
}