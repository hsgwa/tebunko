# 暗号化されたファイルの見分け（状態層）: ファイルの先頭バイト列と、複合ドキュメント形式（CFB。[MS-CFB]）の
# ディレクトリのエントリ名を読み、判断層（office_protection_view.ps1）に渡す。
# ファイルを読むだけで、Office・COMには触らない（getOfficeFileProtection を extract_office.ps1 が呼ぶ）。
#
# 壊れたCFB（FATの鎖が輪になる・範囲外のセクター・ディレクトリが途中で切れる）でも例外を出さず、
# 決まった時間で戻る（読める範囲まで返す。ヘッダーさえ読めなければ $null）。
# ディレクトリのエントリ名だけが目的のため、ミニFAT・ストリームの中身は読まない。

${cfbHeadLength}      = 4096          # 先頭バイト列の読み取り量（判断層の判定に使う分）
${cfbFreeSect}        = [uint32]0xFFFFFFFFL
${cfbEndOfChain}      = [uint32]0xFFFFFFFEL
${cfbFatSect}         = [uint32]0xFFFFFFFDL
${cfbDifSect}         = [uint32]0xFFFFFFFCL
${cfbMaxChainSectors} = 200000         # FAT・ディレクトリの鎖をたどる上限（壊れたファイルで止まらないため）

function readFileHead {
    # ファイルの先頭（既定は4096バイト）を読む。開けない・読めなければ空の配列
    param (
        [string]$path,
        [int]$length = ${cfbHeadLength}
    )

    try {
        $stream = [System.IO.File]::OpenRead((toLongPath $path))
    } catch {
        return [byte[]]@()
    }
    try {
        $buffer = New-Object byte[] $length
        $read = $stream.Read($buffer, 0, $length)
        if ($read -eq $length) {
            return $buffer
        }
        $head = New-Object byte[] ([Math]::Max($read, 0))
        if ($read -gt 0) {
            [System.Array]::Copy($buffer, $head, $read)
        }
        return $head
    } catch {
        return [byte[]]@()
    } finally {
        $stream.Dispose()
    }
}

function readUInt32LE {
    # 範囲外なら FREESECT（鎖の終わりと同じ扱いにする安全な既定値）を返す
    param (
        [byte[]]$bytes,
        [int]$offset
    )

    if ($null -eq $bytes -or $offset -lt 0 -or ($offset + 4) -gt $bytes.Length) {
        return ${cfbFreeSect}
    }
    return [System.BitConverter]::ToUInt32($bytes, $offset)
}

function readCfbSector {
    # セクター番号（0起点）1つ分のバイト列。範囲外・読めなければ $null
    param (
        $stream,
        [double]$sectorSize,
        [long]$fileLength,
        [uint32]$secId
    )

    if ($secId -eq ${cfbFreeSect} -or $secId -eq ${cfbEndOfChain} -or $secId -eq ${cfbFatSect} -or $secId -eq ${cfbDifSect}) {
        return $null
    }
    $offset = [int64]$sectorSize * ([int64]$secId + 1)
    if ($offset -lt 0 -or ($offset + $sectorSize) -gt $fileLength) {
        return $null
    }
    try {
        [void]$stream.Seek($offset, [System.IO.SeekOrigin]::Begin)
        $buffer = New-Object byte[] ([int]$sectorSize)
        if ($stream.Read($buffer, 0, [int]$sectorSize) -ne [int]$sectorSize) {
            return $null
        }
        return $buffer
    } catch {
        return $null
    }
}

function buildCfbFatMap {
    # セクター番号 → 次のセクター番号（FATの鎖） の対応表を作る。
    # DIFAT（ヘッダーの109件 + DIFATセクターの鎖）をたどってFATセクターの並びを求め、その並びの順に読む
    param (
        $stream,
        [byte[]]$header,
        [double]$sectorSize,
        [long]$fileLength
    )

    $entriesPerSector = [int]($sectorSize / 4)
    $fatSectorIds = New-Object 'System.Collections.Generic.List[uint32]'
    for ($i = 0; $i -lt 109; $i++) {
        $value = readUInt32LE $header (76 + $i * 4)
        if ($value -ne ${cfbFreeSect}) {
            $fatSectorIds.Add($value)
        }
    }

    $difatSec = readUInt32LE $header 68
    $seenDifat = New-Object 'System.Collections.Generic.HashSet[uint32]'
    $steps = 0
    while ($difatSec -ne ${cfbEndOfChain} -and $difatSec -ne ${cfbFreeSect} -and $steps -lt ${cfbMaxChainSectors} -and $seenDifat.Add($difatSec)) {
        $buffer = readCfbSector $stream $sectorSize $fileLength $difatSec
        if ($null -eq $buffer) {
            break
        }
        $entriesInSector = $entriesPerSector - 1
        for ($i = 0; $i -lt $entriesInSector; $i++) {
            $value = readUInt32LE $buffer ($i * 4)
            if ($value -ne ${cfbFreeSect}) {
                $fatSectorIds.Add($value)
            }
        }
        $difatSec = readUInt32LE $buffer ($entriesInSector * 4)
        $steps++
    }

    $fatMap = New-Object 'System.Collections.Generic.Dictionary[uint32,uint32]'
    $seenFat = New-Object 'System.Collections.Generic.HashSet[uint32]'
    $sectorIndex = 0
    foreach ($fatSecId in $fatSectorIds) {
        if ($sectorIndex -ge ${cfbMaxChainSectors}) {
            break
        }
        if ($seenFat.Add($fatSecId)) {
            $buffer = readCfbSector $stream $sectorSize $fileLength $fatSecId
            if ($null -ne $buffer) {
                for ($i = 0; $i -lt $entriesPerSector; $i++) {
                    $fatMap[[uint32]($sectorIndex + $i)] = (readUInt32LE $buffer ($i * 4))
                }
            }
        }
        $sectorIndex += $entriesPerSector
    }
    return $fatMap
}

function readCfbDirectoryEntryNames {
    # ディレクトリ（ルート・ストレージ・ストリーム）のエントリ名を、鎖をたどって集める
    param (
        $stream,
        [double]$sectorSize,
        [long]$fileLength,
        [System.Collections.Generic.Dictionary[uint32, uint32]]$fatMap,
        [uint32]$firstDirSector
    )

    $names = New-Object 'System.Collections.Generic.List[string]'
    $entriesPerSector = [int]($sectorSize / 128)
    $seen = New-Object 'System.Collections.Generic.HashSet[uint32]'
    $sec = $firstDirSector
    $steps = 0
    while ($sec -ne ${cfbEndOfChain} -and $sec -ne ${cfbFreeSect} -and $steps -lt ${cfbMaxChainSectors} -and $seen.Add($sec)) {
        $buffer = readCfbSector $stream $sectorSize $fileLength $sec
        if ($null -eq $buffer) {
            break
        }
        for ($i = 0; $i -lt $entriesPerSector; $i++) {
            $entryOffset = $i * 128
            $nameLength = [System.BitConverter]::ToUInt16($buffer, $entryOffset + 64)
            if ($nameLength -ge 2 -and $nameLength -le 64) {
                $name = [System.Text.Encoding]::Unicode.GetString($buffer, $entryOffset, $nameLength - 2)
                if ($name -ne "") {
                    $names.Add($name)
                }
            }
        }
        if (!$fatMap.ContainsKey($sec)) {
            break
        }
        $sec = $fatMap[$sec]
        $steps++
    }
    return $names
}

function readCompoundEntryNames {
    # CFBのディレクトリのエントリ名の配列を返す。CFBでない・読めない・壊れている場合は $null
    # （判断層は $null を、権限保護と分からない「旧形式」として扱う）
    param (
        [string]$path
    )

    try {
        $stream = [System.IO.File]::OpenRead((toLongPath $path))
    } catch {
        return $null
    }
    try {
        $fileLength = $stream.Length
        if ($fileLength -lt 512) {
            return $null
        }
        $header = New-Object byte[] 512
        if ($stream.Read($header, 0, 512) -ne 512) {
            return $null
        }
        $signature = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1)
        for ($i = 0; $i -lt 8; $i++) {
            if ($header[$i] -ne $signature[$i]) {
                return $null
            }
        }

        $sectorShift = [System.BitConverter]::ToUInt16($header, 30)
        if ($sectorShift -ne 9 -and $sectorShift -ne 12) {
            return $null
        }
        $sectorSize = [double][Math]::Pow(2, $sectorShift)
        $firstDirSector = readUInt32LE $header 48

        $fatMap = buildCfbFatMap $stream $header $sectorSize $fileLength
        return @(readCfbDirectoryEntryNames $stream $sectorSize $fileLength $fatMap $firstDirSector)
    } catch {
        return $null
    } finally {
        $stream.Dispose()
    }
}

function getOfficeFileProtection {
    # ファイルを読み、暗号化の種類（getOfficeProtectionKind の戻り値）を返す
    param (
        [string]$path
    )

    $head = readFileHead $path
    $cfbNames = readCompoundEntryNames $path
    return getOfficeProtectionKind $head $cfbNames
}
