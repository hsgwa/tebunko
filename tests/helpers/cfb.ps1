# テストで使う、複合ドキュメント形式（CFB。[MS-CFB]）の最小限のファイルの作成。
# バージョン3・セクター512バイトで、FATセクター1つ・ディレクトリセクター1つ以上を持つ形にする
# （office_protection.ps1・extract_office.ps1 のテストが、暗号化の判定を確かめるのに使う）。

${cfbFatSectMarker}    = [uint32]0xFFFFFFFDL
${cfbEndOfChainMarker} = [uint32]0xFFFFFFFEL
${cfbFreeSectMarker}   = [uint32]0xFFFFFFFFL

function writeDirectoryEntry {
    # 128バイトのディレクトリエントリ1つ分を $buffer の $offset に書く（名前・種類だけ。ほかは0のまま）
    param ([byte[]]$buffer, [int]$offset, [string]$name, [byte]$objectType)

    if ($name -ne "") {
        $nameBytes = [System.Text.Encoding]::Unicode.GetBytes($name + "`0")
        [System.Array]::Copy($nameBytes, 0, $buffer, $offset, $nameBytes.Length)
        [System.Array]::Copy([System.BitConverter]::GetBytes([uint16]$nameBytes.Length), 0, $buffer, $offset + 64, 2)
        $buffer[$offset + 66] = $objectType
    }
}

function buildCfbDirectorySectors {
    # ディレクトリセクター（1セクターにつき Root Entry を含めて4エントリまで）の配列を組み立てる
    param (
        [int]$sectorSize,
        [string[][]]$entrySectors
    )

    $dirSectors = New-Object 'System.Collections.Generic.List[byte[]]'
    for ($d = 0; $d -lt $entrySectors.Count; $d++) {
        $buffer = New-Object byte[] $sectorSize
        if ($d -eq 0) {
            writeDirectoryEntry $buffer 0 "Root Entry" 5
            $names = @($entrySectors[$d])
            for ($i = 0; $i -lt [Math]::Min(3, $names.Count); $i++) {
                writeDirectoryEntry $buffer (($i + 1) * 128) $names[$i] 2
            }
        } else {
            $names = @($entrySectors[$d])
            for ($i = 0; $i -lt [Math]::Min(4, $names.Count); $i++) {
                writeDirectoryEntry $buffer ($i * 128) $names[$i] 2
            }
        }
        $dirSectors.Add($buffer)
    }
    return $dirSectors
}

function newCompoundFile {
    # 最小限のCFB（バージョン3・セクター512バイト）を組み立てる。
    # セクター0 = FAT、セクター1以降 = ディレクトリ（1セクターにつき Root Entry を含めて4エントリまで）。
    #   entrySectors: ディレクトリセクターごとの、ルート以外のエントリ名の配列（@() の配列）
    #   fatOverrides: セクター番号 → 上書きする次のセクター番号（壊れたCFBを作るため）
    param (
        [string]$path,
        [string[][]]$entrySectors = @(@()),
        [hashtable]$fatOverrides = @{}
    )

    $sectorSize = 512
    $numDirSectors = $entrySectors.Count

    $header = New-Object byte[] 512
    $signature = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1)
    [System.Array]::Copy($signature, 0, $header, 0, 8)
    [System.Array]::Copy([System.BitConverter]::GetBytes([uint16]3), 0, $header, 26, 2)   # メジャーバージョン
    [System.Array]::Copy([System.BitConverter]::GetBytes([uint16]9), 0, $header, 30, 2)   # セクターシフト（512バイト）
    [System.Array]::Copy([System.BitConverter]::GetBytes([uint16]6), 0, $header, 32, 2)   # ミニセクターシフト（未使用）
    [System.Array]::Copy([System.BitConverter]::GetBytes([uint32]1), 0, $header, 48, 4)    # 先頭のディレクトリセクター
    [System.Array]::Copy([System.BitConverter]::GetBytes(${cfbEndOfChainMarker}), 0, $header, 68, 4)  # DIFATセクターは使わない
    [System.Array]::Copy([System.BitConverter]::GetBytes([uint32]0), 0, $header, 76, 4)    # DIFAT[0] = FATはセクター0
    for ($i = 1; $i -lt 109; $i++) {
        [System.Array]::Copy([System.BitConverter]::GetBytes(${cfbFreeSectMarker}), 0, $header, 76 + $i * 4, 4)
    }

    $fat = @{}
    $fat[0] = ${cfbFatSectMarker}
    for ($d = 0; $d -lt $numDirSectors; $d++) {
        $secIndex = 1 + $d
        $fat[$secIndex] = $(if ($d -lt $numDirSectors - 1) { [uint32]($secIndex + 1) } else { ${cfbEndOfChainMarker} })
    }
    foreach ($key in $fatOverrides.Keys) {
        $fat[[int]$key] = [uint32]$fatOverrides[$key]
    }

    $fatSector = New-Object byte[] $sectorSize
    for ($i = 0; $i -lt 128; $i++) {
        $value = $(if ($fat.ContainsKey($i)) { $fat[$i] } else { ${cfbFreeSectMarker} })
        [System.Array]::Copy([System.BitConverter]::GetBytes($value), 0, $fatSector, $i * 4, 4)
    }

    $dirSectors = buildCfbDirectorySectors $sectorSize $entrySectors

    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
    $stream = [System.IO.File]::Create($path)
    try {
        $stream.Write($header, 0, $header.Length)
        $stream.Write($fatSector, 0, $fatSector.Length)
        foreach ($sector in $dirSectors) {
            $stream.Write($sector, 0, $sector.Length)
        }
    } finally {
        $stream.Dispose()
    }
}

function newCompoundFileWithDifat {
    # DIFATセクター経由でFATセクターの並びを求める最小限のCFBを作る
    # （ヘッダーの109件は使わず空にし、先頭のDIFATセクター1つだけでFATセクターの場所を示す）。
    #   loopDifat: $true なら、DIFATセクターの「次」を自分自身にして輪にする（例外を出さず戻ることを確かめる用）
    param (
        [string]$path,
        [string[][]]$entrySectors = @(@()),
        [switch]$loopDifat
    )

    $sectorSize = 512
    $numDirSectors = $entrySectors.Count
    $difatSectorIndex = 1 + $numDirSectors   # セクター0=FAT、1..numDirSectors=ディレクトリ、その次がDIFAT

    $header = New-Object byte[] 512
    $signature = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1)
    [System.Array]::Copy($signature, 0, $header, 0, 8)
    [System.Array]::Copy([System.BitConverter]::GetBytes([uint16]3), 0, $header, 26, 2)
    [System.Array]::Copy([System.BitConverter]::GetBytes([uint16]9), 0, $header, 30, 2)
    [System.Array]::Copy([System.BitConverter]::GetBytes([uint16]6), 0, $header, 32, 2)
    [System.Array]::Copy([System.BitConverter]::GetBytes([uint32]1), 0, $header, 48, 4)
    [System.Array]::Copy([System.BitConverter]::GetBytes([uint32]$difatSectorIndex), 0, $header, 68, 4)
    [System.Array]::Copy([System.BitConverter]::GetBytes([uint32]1), 0, $header, 72, 4)
    for ($i = 0; $i -lt 109; $i++) {
        [System.Array]::Copy([System.BitConverter]::GetBytes(${cfbFreeSectMarker}), 0, $header, 76 + $i * 4, 4)
    }

    $fat = @{}
    $fat[0] = ${cfbFatSectMarker}
    for ($d = 0; $d -lt $numDirSectors; $d++) {
        $secIndex = 1 + $d
        $fat[$secIndex] = $(if ($d -lt $numDirSectors - 1) { [uint32]($secIndex + 1) } else { ${cfbEndOfChainMarker} })
    }

    $fatSector = New-Object byte[] $sectorSize
    for ($i = 0; $i -lt 128; $i++) {
        $value = $(if ($fat.ContainsKey($i)) { $fat[$i] } else { ${cfbFreeSectMarker} })
        [System.Array]::Copy([System.BitConverter]::GetBytes($value), 0, $fatSector, $i * 4, 4)
    }

    $dirSectors = buildCfbDirectorySectors $sectorSize $entrySectors

    # DIFATセクター: 最初のエントリ（FATセクターの並びの0番目）= セクター0（FAT）、残りは空、末尾4バイトは次のDIFATセクター
    $difatSector = New-Object byte[] $sectorSize
    [System.Array]::Copy([System.BitConverter]::GetBytes([uint32]0), 0, $difatSector, 0, 4)
    for ($i = 1; $i -lt 127; $i++) {
        [System.Array]::Copy([System.BitConverter]::GetBytes(${cfbFreeSectMarker}), 0, $difatSector, $i * 4, 4)
    }
    $nextDifat = $(if ($loopDifat) { [uint32]$difatSectorIndex } else { ${cfbEndOfChainMarker} })
    [System.Array]::Copy([System.BitConverter]::GetBytes($nextDifat), 0, $difatSector, 127 * 4, 4)

    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
    $stream = [System.IO.File]::Create($path)
    try {
        $stream.Write($header, 0, $header.Length)
        $stream.Write($fatSector, 0, $fatSector.Length)
        foreach ($sector in $dirSectors) {
            $stream.Write($sector, 0, $sector.Length)
        }
        $stream.Write($difatSector, 0, $difatSector.Length)
    } finally {
        $stream.Dispose()
    }
}
