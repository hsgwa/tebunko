# テキストファイルの拡張子ごとの扱い（取り込むか・開き方）を選ぶ画面の、文言と判断（判断層）。
# 画面に触らないため、そのままテストできる（tests/tebunko/ui/settings/text_ext_view.Tests.ps1）。
# 選べるのは shared\core\text_file.ps1 の ${textExtensions} の拡張子だけ（足せない）。実行・登録になる ${textNotepadExtensions} は
# 「既定のアプリで開く」を選べない。設定の読み書きは core\settings.ps1（readTextExtensionSettings・writeTextExtensionSettings）。
# ダイアログは ui\settings\text_ext_dialog.ps1。

function getTextExtensionWayText {
    # 扱いの値（default・notepad・exclude）の、画面に出す文言
    param (
        [string]$way
    )

    switch ($way) {
        "default" { return "既定のアプリで開く" }
        "notepad" { return "メモ帳で開く" }
        "exclude" { return "取り込まない" }
        default   { return "" }
    }
}

function getTextExtensionOptions {
    # その拡張子で選べる扱い（@{ Value; Text } の並び）。実行・登録になる拡張子は「既定のアプリで開く」を選べない
    param (
        [string]$extension
    )

    $ways = @("default", "notepad", "exclude")
    if (${textNotepadExtensions} -contains $extension) {
        $ways = @("notepad", "exclude")
    }
    return @($ways | ForEach-Object { @{ Value = $_; Text = (getTextExtensionWayText $_) } })
}

function getTextExtensionRows {
    # ダイアログの行（${textExtensions} の順）。@{ Extension; Options（選べる扱い）; Way（今の扱い） } の並び。
    # excluded（取り込まない）・notepad（メモ帳で開く）は設定の値。取り込まないが優先。実行・登録になる拡張子は、取り込む限りメモ帳
    param (
        [object[]]$excluded = @(),
        [object[]]$notepad = @()
    )

    $excludedSet = @(@($excluded) | ForEach-Object { ([string]$_).ToLowerInvariant() })
    $notepadSet = @(@($notepad) | ForEach-Object { ([string]$_).ToLowerInvariant() })
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($extension in ${textExtensions}) {
        $way = "default"
        if ($excludedSet -contains $extension) {
            $way = "exclude"
        } elseif ((${textNotepadExtensions} -contains $extension) -or ($notepadSet -contains $extension)) {
            $way = "notepad"
        }
        $rows.Add(@{ Extension = $extension; Options = @(getTextExtensionOptions $extension); Way = $way })
    }
    return $rows.ToArray()
}

function convertTextExtensionRows {
    # ダイアログで選んだ行（@{ Extension; Way } の並び）から、設定に書く 2 つの配列を作る: @{ Excluded; Notepad }。
    # 設定は既定からの差だけを持つため、実行・登録になる拡張子の「メモ帳」は Notepad に書かない（常にメモ帳）。
    # ${textExtensions} に無い拡張子・知らない扱いは捨てる
    param (
        [object[]]$rows
    )

    $excluded = New-Object System.Collections.Generic.List[string]
    $notepad = New-Object System.Collections.Generic.List[string]
    foreach ($row in @($rows)) {
        $extension = ([string]$row.Extension).ToLowerInvariant()
        if (${textExtensions} -notcontains $extension) {
            continue
        }
        if ($row.Way -eq "exclude") {
            $excluded.Add($extension)
        } elseif ($row.Way -eq "notepad" -and ${textNotepadExtensions} -notcontains $extension) {
            $notepad.Add($extension)
        }
    }
    return @{ Excluded = $excluded.ToArray(); Notepad = $notepad.ToArray() }
}

function getTextExtensionSummary {
    # ［設定］の枠に出す、今の状態の要約。既定のまま（2 つとも空）なら「既定のままです」、
    # そうでなければ「取り込む 75 個のうち 74 個 ・ メモ帳で開く 8 個」（メモ帳の数は、取り込むもののうち）
    param (
        [object[]]$excluded = @(),
        [object[]]$notepad = @()
    )

    $rows = @(getTextExtensionRows $excluded $notepad)
    $included = @($rows | Where-Object { $_.Way -ne "exclude" })
    $notepadCount = @($included | Where-Object { $_.Way -eq "notepad" }).Count
    $defaultNotepadCount = @(${textNotepadExtensions}).Count
    if ($included.Count -eq $rows.Count -and $notepadCount -eq $defaultNotepadCount) {
        return "既定のままです"
    }
    return "取り込む $($rows.Count) 個のうち $($included.Count) 個 ・ メモ帳で開く $notepadCount 個"
}

function getTextExtensionDialogTexts {
    # ダイアログの文言: @{ Title; Note; Header（列の見出し）; Reset; Ok; Cancel }
    return @{
        Title     = "テキストファイルの拡張子"
        Note      = "「取り込まない」にした拡張子のファイルは、次にインデックス作成で件数を数えたときにインデックスから外します。取り込むに戻すと、次のインデックス作成で取り込みます。"
        Extension = "拡張子"
        Way       = "扱い"
        Reset     = "既定に戻す"
        Ok        = "OK"
        Cancel    = "キャンセル"
    }
}
