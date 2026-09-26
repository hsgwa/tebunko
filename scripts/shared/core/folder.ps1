# フォルダのパス（正規化・同一判定・ドライブ・フォルダ選択の開始フォルダ）。

function normalizeFolderPath {
    # フォルダパスを1つの書き方にそろえる（書き方の違いで同じフォルダを別のフォルダとみなさないため）。
    #   ・前後の空白・" を取り除く
    #   ・環境変数（%USERPROFILE% など）を展開する
    #   ・/ を \ にそろえ、長いパス用の \\?\ ・ \\?\UNC\ を外す（ファイル操作に渡す直前に toLongPath で付け直す）
    #   ・重なった \ ・ . ・ .. を解決する
    #   ・相対パスはツールのフォルダ（$rootDir）からとみなして絶対パスにする
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
            # 相対パスはツールのフォルダからとみなす
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
    # 同じプロセスでは1回だけ調べる（インデックス作成・検索の途中で割り当てが変わることは想定しない）。
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
        $drives = $null  # ドライブ文字 → 割り当て先（テストで差し替える。$null は getDriveTargets）
    )

    $target = $b.TrimEnd("\")
    # 書き方どおりに同じなら、ドライブの割り当て（CIM。初回は 0.2 秒以上、切断されたネットワークドライブがあるとさらにかかる）を調べない
    if ($a.TrimEnd("\").Equals($target, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }
    if ($null -eq $drives) {
        $drives = getDriveTargets
    }
    foreach ($alias in @(getFolderPathAliases $a $drives)) {
        if ($alias.TrimEnd("\").Equals($target, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }
    return $false
}

function testFolderUnder {
    # path が folder 自身か folder の下のフォルダかを返す（testSameFolder と同じく、書き方の違い・ドライブの割り当てをたどる）。
    # クロール対象フォルダが入れ子になると、同じファイルが2つのインデックスに入り、検索結果にも二重に出るため、その確認に使う
    param (
        [string]$path,
        [string]$folder,
        $drives = $null  # ドライブ文字 → 割り当て先（テストで差し替える。$null は getDriveTargets）
    )

    if ($path -eq "" -or $folder -eq "") {
        return $false
    }
    # 書き方どおりに下にあれば、ドライブの割り当て（CIM）を調べない
    if ($null -ne (getPathUnderFolder $path $folder)) {
        return $true
    }
    if ($null -eq $drives) {
        $drives = getDriveTargets
    }
    foreach ($alias in @(getFolderPathAliases $path $drives)) {
        if ($null -ne (getPathUnderFolder $alias $folder)) {
            return $true
        }
    }
    return $false
}


# ---- フォルダ選択の開始フォルダ ----


function getExistingAncestorFolder {
    # folder が今もあればそのまま、無ければその上の今もあるフォルダを返す（フォルダ選択を開く場所）。どこにも無ければ空
    param (
        [string]$folder
    )

    $dir = $folder
    while ($dir) {
        if (Test-Path -LiteralPath (toLongPath $dir) -PathType Container) {
            return $dir
        }
        $dir = Split-Path $dir -Parent
    }
    return ""
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
