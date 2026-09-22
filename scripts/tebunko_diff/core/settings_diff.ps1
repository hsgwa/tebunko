# setting.config のうち、比較（tebunko_diff）の設定。キーは diff で始める。
# 読み書きは shared\core\settings.ps1（検索の設定のキーは残す）。
# 比較元・比較先のパスは保存しない（画面を開いている間だけ覚え、閉じたら忘れる）。

# 使わなくなったキー（以前の版が保存していた比較元・比較先のパス）。設定を保存するときにファイルから消す
${diffObsoleteSettingKeys} = @("diffFileLeft", "diffFileRight", "diffFolderLeft", "diffFolderRight")

function newDiffSettings {
    # 比較の設定の既定値。設定ファイル（JSON）のキーと同じ
    return [ordered]@{
        diffMode        = "file"   # トグル（file = ファイル、folder = フォルダ）。最後に使った側
        diffSubfolders  = $true    # ［サブフォルダも比較］（［フォルダ］のときだけ使う）
        diffHideSame    = $true    # ［同じファイルを隠す］（［フォルダ］のときだけ使う）
        diffOptions     = [ordered]@{  # 比べ方（両方の側で共通）
            includeShapes    = $true
            includeComments  = $true
            ignoreWhitespace = $false
            caseSensitive    = $true
        }
        diffView        = "side"   # ファイルの差分の表示（side = 左右に並べる、list = 一覧）
        diffFoldSame    = $true    # ［同じ行をたたむ］
        diffTreeHeight  = 260      # フォルダのツリーの高さ（px）
    }
}

function readDiffSettings {
    # 比較の設定を newDiffSettings と同じ形で返す。記載の無い項目は既定値
    param (
        [string]$path = ${settingsFile}
    )

    $settings = newDiffSettings
    $settings = mergeSettingValues $settings (readSettingsData $path)
    if (@("file", "folder") -notcontains $settings.diffMode) { $settings.diffMode = "file" }
    if (@("side", "list") -notcontains $settings.diffView) { $settings.diffView = "side" }
    $settings.diffTreeHeight = [Math]::Max(80, [Math]::Min(2000, [int]$settings.diffTreeHeight))
    return $settings
}

function updateDiffSettings {
    # 比較の設定のうち values のキーだけを変えて保存する（ほかのキー・検索の設定は残す）
    param (
        [System.Collections.IDictionary]$values,
        [string]$path = ${settingsFile}
    )

    $settings = readDiffSettings $path
    foreach ($key in @($values.Keys)) {
        if (!$settings.Contains($key)) {
            throw "比較の設定に無いキーです: $key"
        }
        $settings[$key] = $values[$key]
    }
    writeSettingsFile $settings $path ${diffObsoleteSettingKeys}
}
