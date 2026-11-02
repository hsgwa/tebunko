# テキストファイルの開き方（検索結果から開くとき、拡張子ごとに「既定のアプリで開く」か「エディタで開く」か）を選ぶ画面の、文言と判断（判断層）。
# 画面に触らないため、そのままテストできる（tests/tebunko/ui/settings/text_open_view.Tests.ps1）。
# 選べるのは shared\core\text_file.ps1 の ${textExtensions} の拡張子だけ（足せない）。実行・登録になる ${textNotepadExtensions} は
# 「既定のアプリで開く」を選べない（エディタ固定）。設定の読み書きは core\settings.ps1（readTextExtensionSettings・writeTextOpenSettings）。
# 取り込むかどうか（インデックスごと）は ui\index\text_ext_view.ps1。ダイアログは ui\settings\text_open_dialog.ps1。
# 「エディタ」は、既定はメモ帳（固定のパス）。設定 textEditorPath があり、testTextEditorPath を通り、ファイルがあるときだけそのプログラムで開く。

function getTextOpenWayText {
    # 開き方の値（default・editor）の、画面に出す文言
    param (
        [string]$way
    )

    switch ($way) {
        "default" { return "既定のアプリで開く" }
        "editor"  { return "エディタで開く" }
        default   { return "" }
    }
}

function getTextOpenDefaultWay {
    # その拡張子の既定の開き方。実行・登録になる拡張子（${textNotepadExtensions}）は editor、そのほかは default
    param (
        [string]$extension
    )

    if (${textNotepadExtensions} -contains $extension) {
        return "editor"
    }
    return "default"
}

function getTextOpenOptions {
    # その拡張子で選べる開き方（@{ Value; Text } の並び）。実行・登録になる拡張子は「既定のアプリで開く」を選べない
    param (
        [string]$extension
    )

    $ways = @("default", "editor")
    if (${textNotepadExtensions} -contains $extension) {
        $ways = @("editor")
    }
    return @($ways | ForEach-Object { @{ Value = $_; Text = (getTextOpenWayText $_) } })
}

function getTextOpenRows {
    # ダイアログの行（${textExtensions} の順）。@{ Extension; Options（選べる開き方）; Way（今の開き方） } の並び。
    # notepad は設定のエディタで開く拡張子（60 個のうち）。実行・登録になる拡張子は、書いてあっても常にエディタ
    param (
        [object[]]$notepad = @()
    )

    $notepadSet = @(@($notepad) | ForEach-Object { ([string]$_).ToLowerInvariant() })
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($extension in ${textExtensions}) {
        $way = getTextOpenDefaultWay $extension
        if ($notepadSet -contains $extension) {
            $way = "editor"
        }
        $rows.Add(@{ Extension = $extension; Options = @(getTextOpenOptions $extension); Way = $way })
    }
    return $rows.ToArray()
}

function convertTextOpenRows {
    # ダイアログで選んだ行（@{ Extension; Way } の並び）から、設定に書く notepadTextExtensions の配列を作る。
    # 設定は既定からの差だけを持つため、実行・登録になる拡張子（常にエディタ）は書かない。${textExtensions} に無い拡張子・知らない開き方は捨てる
    param (
        [object[]]$rows
    )

    $notepad = New-Object System.Collections.Generic.List[string]
    foreach ($row in @($rows)) {
        $extension = ([string]$row.Extension).ToLowerInvariant()
        if (${textExtensions} -notcontains $extension) {
            continue
        }
        if ($row.Way -eq "editor" -and ${textNotepadExtensions} -notcontains $extension) {
            $notepad.Add($extension)
        }
    }
    return $notepad.ToArray()
}

function getTextOpenChangeSummary {
    # 既定から変えたものの要約。rows は @{ Extension; Way } の並び。
    # 変えたものが無ければ「すべて既定のままです（エディタで開くのは、実行・登録になる 15 個）」、あれば「変更したもの 2 個（エディタで開く）」
    param (
        [object[]]$rows
    )

    $changed = @(@($rows) | Where-Object { $_.Way -ne (getTextOpenDefaultWay $_.Extension) })
    if ($changed.Count -eq 0) {
        return "すべて既定のままです（エディタで開くのは、実行・登録になる $(@(${textNotepadExtensions}).Count) 個）"
    }
    return "変更したもの $($changed.Count) 個（エディタで開く）"
}

function getTextOpenSummary {
    # ［設定］の枠に出す、今の状態の要約（ダイアログの要約と同じ言い方。getTextOpenChangeSummary）
    param (
        [object[]]$notepad = @()
    )

    return (getTextOpenChangeSummary @(getTextOpenRows $notepad))
}

function getTextOpenListView {
    # ダイアログの一覧の見せ方。rows は @{ Extension; Way } の並び、filter は絞り込み欄の文字、changedOnly は［変更したものだけ］。
    # 返す値: @{ Visible（見せる拡張子の並び）; Summary（要約）; Empty（一致が無いか）; ShowRestrictedHeader（実行・登録になる形式の見出しを出すか） }
    # 絞り込みは、拡張子に含まれるか（大文字・小文字を区別せず、先頭の . は有っても無くてもよい）。要約は
    # 何も絞らないとき変更したものの要約、［変更したものだけ］のとき「変更したもの 3 個を表示」、文字で絞るとき「75 個中 2 個を表示」
    param (
        [object[]]$rows,
        [string]$filter = "",
        [bool]$changedOnly = $false
    )

    $needle = ([string]$filter).Trim().ToLowerInvariant().TrimStart(".")
    $visible = New-Object System.Collections.Generic.List[string]
    foreach ($row in @($rows)) {
        if ($changedOnly -and $row.Way -eq (getTextOpenDefaultWay $row.Extension)) {
            continue
        }
        if ($needle -ne "" -and -not $row.Extension.TrimStart(".").Contains($needle)) {
            continue
        }
        $visible.Add($row.Extension)
    }
    if ($needle -ne "") {
        $summary = "$(@($rows).Count) 個中 $($visible.Count) 個を表示"
    } elseif ($changedOnly) {
        $summary = "変更したもの $($visible.Count) 個を表示"
    } else {
        $summary = getTextOpenChangeSummary $rows
    }
    return @{
        Visible              = $visible.ToArray()
        Summary              = $summary
        Empty                = ($visible.Count -eq 0)
        ShowRestrictedHeader = (@($visible | Where-Object { ${textNotepadExtensions} -contains $_ }).Count -gt 0)
    }
}

# エディタに指定できない、スクリプトや命令を動かすプログラムのファイル名（拡張子なし。小文字）。実行・登録になる拡張子のファイルをエディタに渡すと、
# そのスクリプトが動いてしまうため。この表は漏れうる（ここに無い、別のプログラムでスクリプトを動かせるものは防げない）。docs/safety/checks.md「エディタの差し替え」に限界を書く
${textEditorDenyNames} = @(
    "powershell", "powershell_ise", "pwsh", "cmd", "wscript", "cscript", "mshta", "bash", "wsl", "sh",
    "python", "pythonw", "py", "pyw", "node", "ruby", "perl", "php", "lua", "java", "javaw",
    "conhost", "wt", "regedit", "reg", "msbuild", "rundll32", "regsvr32", "msiexec", "cmstp", "installutil", "regasm", "regsvcs", "forfiles", "pcalua", "explorer"
)

function testTextEditorDenied {
    # エディタに指定されたパスが、スクリプトや命令を動かすプログラム（${textEditorDenyNames}）か。ファイル名だけを見る
    # （大文字・小文字を区別せず、フォルダの付いた形・拡張子の有無も同じに扱う。区切りは \ と /）
    param (
        $path
    )

    if ($path -isnot [string] -or $path -eq "") {
        return $false
    }
    $leaf = $path.Substring($path.LastIndexOfAny([char[]]@([char]92, [char]47)) + 1).Trim().TrimEnd(".", " ").ToLowerInvariant()
    $name = $(if ($leaf.EndsWith(".exe")) { $leaf.Substring(0, $leaf.Length - 4) } else { $leaf })
    return (${textEditorDenyNames} -contains $name)
}

function testTextEditorDeniedByVersion {
    # 実行ファイルの版の情報の元の名前（OriginalFilename）か内部名（InternalName）が、スクリプトや命令を動かすプログラム（${textEditorDenyNames}）か。
    # 別の名前にコピーしたものを止めるため。末尾は .mui を 1 回、そのあと .exe か .dll を 1 回外す（繰り返さない）。大文字・小文字は区別しない
    param (
        $originalFilename,
        $internalName
    )

    foreach ($value in @($originalFilename, $internalName)) {
        if ($value -isnot [string] -or $value -eq "") {
            continue
        }
        $name = $value.Trim().TrimEnd(".", " ").ToLowerInvariant()
        if ($name.EndsWith(".mui")) {
            $name = $name.Substring(0, $name.Length - 4)
        }
        if ($name.EndsWith(".exe") -or $name.EndsWith(".dll")) {
            $name = $name.Substring(0, $name.Length - 4)
        }
        if (${textEditorDenyNames} -contains $name) {
            return $true
        }
    }
    return $false
}

function testTextEditorPath {
    # 設定 textEditorPath の値が、使ってよい形か（文字列だけを見る。ファイルがあるかは openWithEditor の側で確かめる）。
    # ドライブ文字から始まる絶対パスで、拡張子が .exe。UNC・相対・環境変数（%）・引用符・引数付きの書き方、
    # Start-Process がワイルドカードとして扱う [ ] * ?、ファイル名に使えない文字は不可
    param (
        $path
    )

    if ($path -isnot [string] -or $path -eq "") {
        return $false
    }
    if ($path -notmatch '^[A-Za-z]:\\[^\\]') {
        return $false
    }
    # ~ は 8.3 の短い名前（POWERS~1.EXE のように、ファイル名の表に当たらずに別のプログラムを指せる）、ドライブ文字の後ろの : は別のデータストリーム（a.exe:powershell.exe）を指せるため、どちらも拒む
    if ($path -match '["%~\[\]\*\?<>\|]' -or $path.IndexOf(':', 2) -ge 0 -or $path.IndexOfAny([System.IO.Path]::GetInvalidPathChars()) -ge 0) {
        return $false
    }
    if ($path -match '(^|\\)\.\.?(\\|$)') {
        return $false
    }
    if (-not $path.EndsWith(".exe", [System.StringComparison]::OrdinalIgnoreCase)) {
        return $false
    }
    # スクリプトや命令を動かすプログラムはエディタにできない（実行・登録になる拡張子のファイルを渡すと動いてしまうため）
    return (-not (testTextEditorDenied $path))
}

function getTextEditorChoice {
    # 「エディタで開く」に使うプログラムを決める。configured は設定 textEditorPath、exists は configured のファイルがあるか（呼ぶ側が調べる）、
    # notepad は固定のメモ帳のパス。返す値: @{ Path（起動するプログラム）; Notice（使えなかったときの知らせ。無ければ空）; Label（画面に出す名前） }
    # 空ならメモ帳。値があっても、形が正しくない・ファイルが無いときは、既定のアプリには戻さずメモ帳で開き、知らせる。
    # versionNames: 版の情報（getTextEditorVersionNames の値 @{ Ok; OriginalFilename; InternalName }）。省略は「読んでいない」で、名前だけで決める。
    # 読めなかった（Ok が $false）とき、元の名前・内部名が表に当たるときは、スクリプトを動かすプログラムとして拒む（どちらも空なら拒まない）
    param (
        $configured,
        [bool]$exists,
        [string]$notepad,
        $versionNames = $null
    )

    if ($configured -is [string] -and $configured -ne "") {
        $deniedByVersion = $false
        $unreadable = $false
        if ($versionNames -is [hashtable]) {
            if (-not $versionNames.Ok) {
                $unreadable = $true
            } else {
                $deniedByVersion = testTextEditorDeniedByVersion $versionNames.OriginalFilename $versionNames.InternalName
            }
        }
        if ((testTextEditorPath $configured) -and $exists -and -not $deniedByVersion -and -not $unreadable) {
            return @{ Path = $configured; Notice = ""; Label = [System.IO.Path]::GetFileNameWithoutExtension($configured) }
        }
        if ((testTextEditorDenied $configured) -or $deniedByVersion) {
            return @{ Path = $notepad; Notice = "設定のエディタ（textEditorPath）は、スクリプトを動かすプログラムのため使えません。メモ帳で開きます。"; Label = "メモ帳" }
        }
        return @{ Path = $notepad; Notice = "設定のエディタ（textEditorPath）が使えないため、メモ帳で開きます。"; Label = "メモ帳" }
    }
    return @{ Path = $notepad; Notice = ""; Label = "メモ帳（既定）" }
}

function addTextEditorNotice {
    # ステータスの文に、設定のエディタを使えなかった知らせ（無ければ空）を後ろに足す
    param (
        [string]$status,
        [string]$notice
    )

    if ($notice -eq "") {
        return $status
    }
    return "${status} ${notice}"
}

function getTextEditorLabel {
    # ［設定］の「エディタ」の行・ダイアログの下の行に出す、使うエディタの名前（設定が空なら「メモ帳（既定）」）。
    # 使えない値のときは、実際に開くメモ帳の名前を出す（ファイルの有無は見ない。開くときに知らせる）
    param (
        $configured
    )

    if ($configured -is [string] -and (testTextEditorPath $configured)) {
        return [System.IO.Path]::GetFileNameWithoutExtension($configured)
    }
    return "メモ帳（既定）"
}

function getTextOpenDialogTexts {
    # ダイアログと［設定］の枠の文言
    return @{
        Title             = "テキストの開き方"
        Note              = "検索結果・プレビューからテキストファイルを開くときの開き方を、拡張子ごとに選びます（全インデックス共通）。「エディタで開く」はメモ帳で開きます。取り込むかどうかは、インデックスの更新の確認で、インデックスごとに選べます。"
        Extension         = "拡張子"
        Way               = "検索結果から開くとき"
        Reset             = "既定に戻す"
        Ok                = "OK"
        Cancel            = "キャンセル"
        ChangedOnly       = "変更したものだけ"
        FilterPlaceholder = "拡張子で絞り込み"
        Empty             = "一致する拡張子がありません"
        RestrictedHeader  = "実行・登録になる形式（「エディタで開く」だけ選べます）"
        RestrictedNote    = "実行できる形式のため、エディタ固定"
        EditorTitle       = "エディタ"
        EditorHint        = "「エディタで開く」に使うアプリ。いまはメモ帳だけ"
        EditorSettingsHint = "「エディタで開く」に使うアプリです。今は{0}だけです。"
        FutureMark        = "将来の対応"
        Change            = "変更…"
        SettingsHint      = "拡張子ごとに、検索結果・プレビューから開くときの開き方（既定のアプリ／エディタ）を選びます（全インデックス共通）。"
    }
}
