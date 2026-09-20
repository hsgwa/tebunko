# フォルダのパスと一覧（正規化・同一判定・ドライブ・エクスプローラー風の一覧）。

function normalizeFolderPath {
    # フォルダパスを1つの書き方にそろえる（書き方の違いで同じフォルダを別のフォルダとみなさないため）。
    #   ・前後の空白・" を取り除く
    #   ・環境変数（%USERPROFILE% など）を展開する
    #   ・/ を \ にそろえ、長いパス用の \\?\ ・ \\?\UNC\ を外す（ファイル操作に渡す直前に toLongPath で付け直す）
    #   ・重なった \ ・ . ・ .. を解決する
    #   ・相対パスは win_grep のフォルダ（$rootDir）からとみなして絶対パスにする
    #   ・末尾の \ を取り除く（ドライブ直下は "D:\" のまま。UNC の共有直下は "\\server\share"）
    # パスとして解釈できない場合（* ? を含む・共有名の無い \\server など）は、書かれたとおりに扱う
    param (
        [string]$path
    )

    $path = $path.Trim().Trim('"').Trim()
    if ($path -eq "") {
        return ""
    }
    if ($path.IndexOf("%") -ge 0) {
        $path = [System.Environment]::ExpandEnvironmentVariables($path).Trim()
    }
    $path = (fromLongPath ($path.Replace("/", "\"))).TrimEnd("\")
    if ($path -eq "") {
        return ""
    }
    # ドライブ名だけ（"D:"）はドライブ直下とする。
    # GetFullPath はそのドライブの「現在のフォルダ」を返すことがあるため、先に決める
    if ($path -match "^[A-Za-z]:$") {
        return "${path}\"
    }
    try {
        if (![System.IO.Path]::IsPathRooted($path)) {
            # 相対パスは win_grep のフォルダからとみなす
            $path = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine(${rootDir}, $path)).TrimEnd("\")
        } elseif ($path -match "^[A-Za-z]:\\" -or $path.StartsWith("\\")) {
            $path = [System.IO.Path]::GetFullPath($path).TrimEnd("\")
        }
        # \ ひとつで始まるパス（"\server\share"）は、書き間違えた UNC パスのことが多いため、
        # 今のドライブのパスに直さず、書かれたとおりに扱う
    } catch {
        # パスとして解釈できない場合は、書かれたとおりに扱う（存在しないフォルダとして扱われる）
    }
    if ($path -match "^[A-Za-z]:$") {
        return "${path}\"
    }
    return $path
}

function getPathUnderFolder {
    # path が folder 自身か folder の下なら、folder からの相対パス（folder 自身は ""）を返す。
    # folder の下でなければ $null（大文字・小文字は区別しない）。
    # パスの文字数で切り出すと、書き方が少し違うだけ（末尾の \ ・\\?\ 付きなど）で取り違えるため、この関数を通す
    param (
        [string]$path,
        [string]$folder
    )

    $folder = $folder.TrimEnd("\")
    $path = $path.TrimEnd("\")
    if ($folder -eq "" -or $path -eq "") {
        return $null
    }
    if ($path.Equals($folder, [System.StringComparison]::OrdinalIgnoreCase)) {
        return ""
    }
    if ($path.StartsWith("${folder}\", [System.StringComparison]::OrdinalIgnoreCase)) {
        return $path.Substring($folder.Length + 1)
    }
    return $null
}


# ドライブ文字の割り当て先（ネットワークドライブ）は CIM で調べる（getDriveTargets 参照）。
# ※以前は Win32 API（QueryDosDevice）で subst も解決していたが、実行時コンパイル（csc.exe）を無くすため CIM に変更した。

${driveTargets} = $null


function getDriveTargets {
    # ドライブ文字（"Z:"）→ 割り当て先（ネットワークドライブは "\\server\share"）を返す。
    # 同じプロセスでは1回だけ調べる（変換・検索の途中で割り当てが変わることは想定しない）。
    # subst で割り当てたドライブは解決しない（CIM で取れないため。実運用ではネットワークドライブが主）。
    if ($null -ne ${script:driveTargets}) {
        return ${script:driveTargets}
    }
    $map = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
    try {
        # DriveType=4 はネットワークドライブ。DeviceID="Z:"、ProviderName="\\server\share"
        foreach ($d in @(Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=4" -ErrorAction SilentlyContinue)) {
            if ($d.ProviderName) {
                $map[$d.DeviceID] = $d.ProviderName.TrimEnd("\")
            }
        }
    } catch {
        # 調べられない環境では別名なしとする（パスを書かれたとおりに使う）
    }
    ${script:driveTargets} = $map
    return $map
}

function getFolderPathAliases {
    # 同じ場所を指す別の書き方を、path 自身を先頭にして返す（ドライブの割り当てをたどる）。
    # ファイルサーバーのフォルダは、ネットワークドライブ（Z:\…）と UNC パス（\\server\share\…）のどちらでも書けるため、
    # 設定に書いた書き方と、インデックスに記録した書き方が違っても同じフォルダと分かるようにする。
    #   Z: が \\server\share のネットワークドライブのとき
    #     "Z:\見積"             → @("Z:\見積", "\\server\share\見積")
    #     "\\server\share\見積" → @("\\server\share\見積", "Z:\見積")
    param (
        [string]$path,
        $drives = (getDriveTargets)  # ドライブ文字 → 割り当て先（テストで差し替える）
    )

    $result = New-Object System.Collections.Generic.List[string]
    $result.Add($path)
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    [void]$seen.Add($path.TrimEnd("\"))

    foreach ($entry in $drives.GetEnumerator()) {
        $alias = $null
        # ドライブ文字 → 割り当て先
        $rest = getPathUnderFolder $path $entry.Key
        if ($null -ne $rest) {
            $alias = if ($rest -eq "") { $entry.Value } else { joinSourcePath $entry.Value $rest }
        } else {
            # 割り当て先 → ドライブ文字
            $rest = getPathUnderFolder $path $entry.Value
            if ($null -ne $rest) {
                $alias = if ($rest -eq "") { "$($entry.Key)\" } else { joinSourcePath "$($entry.Key)\" $rest }
            }
        }
        if ($alias -and $seen.Add($alias.TrimEnd("\"))) {
            $result.Add($alias)
        }
    }
    return $result.ToArray()
}

function testSameFolder {
    # 2つのパスが同じフォルダを指すか（末尾の \ ・大文字と小文字の違いと、ネットワークドライブ・subst の割り当てをたどる）
    param (
        [string]$a,
        [string]$b,
        $drives = (getDriveTargets)  # ドライブ文字 → 割り当て先（テストで差し替える）
    )

    $target = $b.TrimEnd("\")
    foreach ($alias in @(getFolderPathAliases $a $drives)) {
        if ($alias.TrimEnd("\").Equals($target, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }
    return $false
}

function testFolderUnder {
    # path が folder 自身か folder の下のフォルダかを返す（testSameFolder と同じく、書き方の違い・ドライブの割り当てをたどる）。
    # 変換対象フォルダが入れ子になると、同じファイルが2つのインデックスに入り、検索結果にも二重に出るため、その確認に使う
    param (
        [string]$path,
        [string]$folder,
        $drives = (getDriveTargets)  # ドライブ文字 → 割り当て先（テストで差し替える）
    )

    if ($path -eq "" -or $folder -eq "") {
        return $false
    }
    foreach ($alias in @(getFolderPathAliases $path $drives)) {
        if ($null -ne (getPathUnderFolder $alias $folder)) {
            return $true
        }
    }
    return $false
}


# ---- フォルダ選択（エクスプローラー風のフォルダ選択ダイアログが使う） ----


function joinFolderPath {
    # フォルダのパスとその中の名前をつなぐ（ドライブ直下 "C:\" ・共有フォルダ直下 "\\server\share" で \ が重ならないようにする）
    param (
        [string]$folder,
        [string]$name
    )

    if ($folder -eq "") {
        return $name
    }
    return ($folder.TrimEnd("\") + "\" + $name)
}

function getParentFolderPath {
    # 1つ上のフォルダを返す。これより上へはたどれない場合（ドライブ直下 "C:\"・共有フォルダ直下 "\\server\share"）は ""。
    # フォルダ選択ダイアログの［↑］（1つ上へ）で使う
    param (
        [string]$path
    )

    $path = ([string]$path).Trim().TrimEnd("\")
    if ($path -eq "" -or $path -match "^[A-Za-z]:$") {
        return ""
    }
    if ($path.StartsWith("\\")) {
        # "\\server\share" までで1つのフォルダ（共有フォルダ）。サーバー名だけ・共有名までなら、これより上は無い
        if (@($path.Substring(2) -split "\\" | Where-Object { $_ -ne "" }).Count -le 2) {
            return ""
        }
    }
    $parent = ([string][System.IO.Path]::GetDirectoryName($path)).TrimEnd("\")
    if ($parent -eq "") {
        return ""
    }
    if ($parent -match "^[A-Za-z]:$") {
        # ドライブ直下は "C:\"（normalizeFolderPath と同じ書き方）
        return "${parent}\"
    }
    return $parent
}

function testVisibleEntry {
    # 隠し・システムの属性が付いていない（エクスプローラーの既定で見える）ものかを返す
    param (
        [System.IO.FileSystemInfo]$entry
    )

    return (($entry.Attributes -band ([System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System)) -eq 0)
}

function testHasSubFolders {
    # フォルダの中にサブフォルダ（隠し・システムを除く）があるかを返す。
    # フォルダ選択ダイアログのツリーで ▷（展開できる印）を出すかの判定に使う。
    # 1つ見つかった時点で打ち切るため、中身の多いフォルダでも待たない
    param (
        [string]$path
    )

    try {
        $dir = New-Object System.IO.DirectoryInfo ((toLongPath $path))
        foreach ($sub in $dir.EnumerateDirectories()) {
            if (testVisibleEntry $sub) {
                return $true
            }
        }
    } catch {
        # 開けないフォルダ（権限が無い・切れているネットワークドライブなど）は、サブフォルダ無しとして扱う
    }
    return $false
}

function getFolderEntries {
    # フォルダの中身を、フォルダ選択ダイアログの一覧に出す順（フォルダが先、それぞれ名前順）で返す。
    # 隠し・システムのフォルダとファイルは出さない（エクスプローラーの既定と同じ）。
    #   Entries     : @{ Name; Path; IsFolder; IsOffice; Updated（DateTime。取れなければ $null） } の配列
    #   FolderCount : 一覧に出したフォルダの数
    #   OfficeCount : 一覧に出した Office ファイルの数（そのフォルダが目的のフォルダかの目安になる）
    #   Truncated   : 中身が limit 件を超えて打ち切ったか（中身の多いフォルダで画面が固まらないようにする）
    #   Error       : 開けなかった理由（開けたときは ""）
    # foldersOnly を付けるとフォルダだけを返す（ツリーの読み込み用。ファイルを数えない分だけ速い）
    param (
        [string]$path,
        [int]$limit = 2000,
        [switch]$foldersOnly
    )

    $folders = New-Object System.Collections.Generic.List[object]
    $files = New-Object System.Collections.Generic.List[object]
    $truncated = $false
    $message = ""
    if (([string]$path).Trim() -eq "") {
        return @{ Entries = @(); FolderCount = 0; OfficeCount = 0; Truncated = $false; Error = "フォルダを指定してください。" }
    }
    try {
        $dir = New-Object System.IO.DirectoryInfo ((toLongPath $path))
        $scanned = 0
        foreach ($entry in $dir.EnumerateFileSystemInfos()) {
            # 中身が非常に多いフォルダでも待たせないよう、見た件数でも打ち切る（隠しファイルばかりのフォルダ対策）
            $scanned++
            if ($scanned -gt ($limit * 10)) {
                $truncated = $true
                break
            }
            if (-not (testVisibleEntry $entry)) {
                continue
            }
            if (($folders.Count + $files.Count) -ge $limit) {
                $truncated = $true
                break
            }
            $isFolder = (($entry.Attributes -band [System.IO.FileAttributes]::Directory) -ne 0)
            if ($foldersOnly -and -not $isFolder) {
                continue
            }
            $updated = $null
            try {
                $updated = $entry.LastWriteTime
            } catch {
                # 更新日時が取れなくても一覧には出す
            }
            $item = [pscustomobject]@{
                Name     = $entry.Name
                Path     = (joinFolderPath $path $entry.Name)
                IsFolder = $isFolder
                IsOffice = ((-not $isFolder) -and (testOfficeFile $entry.Name))
                Updated  = $updated
            }
            if ($isFolder) {
                $folders.Add($item)
            } else {
                $files.Add($item)
            }
        }
    } catch [System.UnauthorizedAccessException] {
        $message = "このフォルダを開く権限がありません。"
    } catch [System.IO.DirectoryNotFoundException] {
        $message = "フォルダが見つかりません。"
    } catch {
        $message = "フォルダを開けません（$($_.Exception.Message)）。"
    }
    $sortedFolders = @($folders | Sort-Object -Property Name)
    $sortedFiles = @($files | Sort-Object -Property Name)
    return @{
        Entries     = @($sortedFolders + $sortedFiles)
        FolderCount = $sortedFolders.Count
        OfficeCount = @($sortedFiles | Where-Object { $_.IsOffice }).Count
        Truncated   = $truncated
        Error       = $message
    }
}

function getQuickFolders {
    # フォルダ選択ダイアログの左側に出す「よく使う場所」（実際にあるフォルダだけ）。
    #   @{ Name; Path } の配列
    $items = New-Object System.Collections.Generic.List[object]
    $places = @(
        @{ Name = "デスクトップ"; Path = [System.Environment]::GetFolderPath("DesktopDirectory") },
        @{ Name = "ドキュメント"; Path = [System.Environment]::GetFolderPath("MyDocuments") },
        @{ Name = "ダウンロード"; Path = (joinFolderPath ([System.Environment]::GetFolderPath("UserProfile")) "Downloads") }
    )
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($place in $places) {
        $path = ([string]$place.Path).TrimEnd("\")
        if ($path -eq "" -or -not $seen.Add($path)) {
            continue
        }
        try {
            if (-not [System.IO.Directory]::Exists((toLongPath $path))) {
                continue
            }
        } catch {
            continue
        }
        $items.Add([pscustomobject]@{ Name = $place.Name; Path = $path })
    }
    return $items.ToArray()
}

function getComputerFolders {
    # フォルダ選択ダイアログの「PC」の下に出すドライブの一覧（使えるドライブだけ）。
    #   @{ Name = "Windows (C:)"; Path = "C:\" } の配列
    # ネットワークドライブは割り当て先（\\server\share）を名前に出す
    param (
        $drives = (getDriveTargets)  # ドライブ文字 → 割り当て先（テストで差し替える）
    )

    $items = New-Object System.Collections.Generic.List[object]
    $found = @()
    try {
        $found = [System.IO.DriveInfo]::GetDrives()
    } catch {
        # ドライブを調べられない環境では、ツリーにドライブを出さない（アドレスバーからは開ける）
        return $items.ToArray()
    }
    foreach ($drive in $found) {
        try {
            if (-not $drive.IsReady) {
                continue
            }
            $letter = $drive.Name.TrimEnd("\")
            $label = ""
            try {
                $label = ([string]$drive.VolumeLabel).Trim()
            } catch {
                # ラベルが取れないドライブは種類の名前で出す
            }
            if ($drive.DriveType -eq [System.IO.DriveType]::Network -and $drives.ContainsKey($letter)) {
                $label = $drives[$letter]
            }
            if ($label -eq "") {
                $label = switch ($drive.DriveType) {
                    "Network"   { "ネットワークドライブ" }
                    "Removable" { "リムーバブルディスク" }
                    "CDRom"     { "DVD ドライブ" }
                    default     { "ローカルディスク" }
                }
            }
            $items.Add([pscustomobject]@{ Name = "$label ($letter)"; Path = "${letter}\" })
        } catch {
            # 準備できていないドライブ（切れているネットワークドライブなど）は飛ばす
            continue
        }
    }
    return $items.ToArray()
}

function getFolderLeafName {
    # フォルダ名（ドライブ直下はドライブ名、UNC の共有直下は共有名）を返す
    param (
        [string]$folderPath
    )

    $path = $folderPath.TrimEnd("\")
    $leaf = [System.IO.Path]::GetFileName($path)
    if ($leaf -eq "") {
        $leaf = $path.TrimEnd(":")
    }
    return $leaf
}
