# setting.config（JSON）の読み書き（どのツールからも使う）。
# 1 つのファイルを複数のツールで使うため、書き戻すときは、書く側が知らないキー（ほかのツールの設定）も残す。
# キーと既定値は、ツールごとに持つ（検索は settings_grep.ps1、比較は settings_diff.ps1）。

# 設定ファイル（ツールを置いたフォルダの直下）
${settingsFile} = "${rootDir}\setting.config"

function readSettingsData {
    # 設定ファイルを読み、中身（PSCustomObject）を返す。ファイルが無い・空・中身が null なら $null を返す。
    # JSON として読めなければ、ファイル名の入ったメッセージで例外にする
    param (
        [string]$path = ${settingsFile}
    )

    if (!(Test-Path -LiteralPath $path)) {
        return $null
    }
    $json = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
    if ($json.Trim() -eq "") {
        return $null
    }
    try {
        return (ConvertFrom-Json $json)
    } catch {
        throw "$([System.IO.Path]::GetFileName($path)) を読み込めません。（$($_.Exception.Message)）"
    }
}

function mergeSettingValues {
    # 既定値（settings。キーの一覧でもある）に、読んだ中身（data）の値を型をそろえて入れる。
    # 記載の無い・null のキーは既定値のまま。settings を書き換えて返す
    param (
        $settings,
        $data
    )

    if ($null -eq $data) {
        return $settings
    }
    foreach ($key in @($settings.Keys)) {
        $property = $data.PSObject.Properties[$key]
        if ($null -eq $property -or $null -eq $property.Value) {
            continue
        }
        $settings[$key] = toSettingValue $property.Value $settings[$key]
    }
    return $settings
}

function toSettingValue {
    # 読んだ値（value）を、既定値（default）と同じ型にする。読めない値は既定値にする
    param (
        $value,
        $default
    )

    if ($default -is [array]) {
        return @($value | Where-Object { $null -ne $_ })
    }
    if ($default -is [string]) {
        return [string]$value
    }
    if ($default -is [bool]) {
        return (toSettingBool $value $default)
    }
    if ($default -is [int]) {
        $parsed = 0
        if ([int]::TryParse([string]$value, [ref]$parsed)) {
            return $parsed
        }
        return $default
    }
    if ($default -is [System.Collections.IDictionary]) {
        # 入れ子の設定（キーの決まったもの）。既定値の写しに、読んだ値を入れる
        $copy = [ordered]@{}
        foreach ($key in @($default.Keys)) {
            $copy[$key] = $default[$key]
        }
        return (mergeSettingValues $copy $value)
    }
    return $value
}

function toSettingBool {
    # 設定ファイルの真偽値を読む。手で書き換えた "false" などの文字列は [bool] にすると $true になるため、
    # 文字列は true / false として読み、読めなければ既定値（default）にする
    param (
        $value,
        [bool]$default
    )

    if ($value -is [string]) {
        $parsed = $false
        if ([bool]::TryParse($value.Trim(), [ref]$parsed)) {
            return $parsed
        }
        return $default
    }
    return [bool]$value
}

function writeSettingsFile {
    # 設定（settings）を書き込む。ファイルにあって settings に無いキー（ほかのツールの設定）は、そのまま残す。
    # removeKeys に挙げたキー（使わなくなった設定）はファイルから消す。
    # 今のファイルが JSON として読めないときは、settings だけで書き直す（読めない内容は残せないため）
    param (
        $settings,
        [string]$path = ${settingsFile},
        [string[]]$removeKeys = @()
    )

    $merged = [ordered]@{}
    foreach ($key in @($settings.Keys)) {
        $merged[$key] = $settings[$key]
    }
    $current = $null
    try {
        $current = readSettingsData $path
    } catch {
        $current = $null
    }
    if ($current -is [System.Management.Automation.PSCustomObject]) {
        foreach ($property in $current.PSObject.Properties) {
            if (!$merged.Contains($property.Name) -and $removeKeys -notcontains $property.Name) {
                $merged[$property.Name] = $property.Value
            }
        }
    }

    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
    [System.IO.File]::WriteAllText($path, (ConvertTo-Json -InputObject $merged -Depth 6), (New-Object System.Text.UTF8Encoding($false)))
}
