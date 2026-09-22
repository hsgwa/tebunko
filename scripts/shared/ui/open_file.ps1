# 元のファイルを開く（どのツールからも使う。検索結果・比較の結果から開く）。
# 開き方（${openModes}）は shared\office\office_files.ps1。Excel はシートとセルを選んで開く

function openWithShell {
    # ファイルを既定のアプリで開く。開き方（mode）は、エクスプローラーの右クリックメニューと同じ動詞で行う。
    #   読み取り専用 → OpenAsReadOnly、新規 → New（元のファイルを基にした無題の文書。元のファイルを占有しない）
    # その動詞が登録されていない種類のファイルは、そのまま開いて $false を返す
    param (
        [string]$path,
        [string]$mode
    )

    $verb = switch ($mode) {
        ${openModeReadOnly} { "OpenAsReadOnly" }
        ${openModeNew}      { "New" }
        default             { $null }
    }
    if ($verb) {
        # Start-Process は [ ] をワイルドカードとして扱うため使わない。
        # ProcessStartInfo.Verbs は New を一覧に含めないため、実行してみて、無ければ例外で分かる
        $info = New-Object System.Diagnostics.ProcessStartInfo($path)
        $info.UseShellExecute = $true
        $info.Verb = $verb
        try {
            [void][System.Diagnostics.Process]::Start($info)
            return $true
        } catch [System.ComponentModel.Win32Exception] {
            # その動詞が登録されていない
        }
    }
    Invoke-Item -LiteralPath $path
    return (-not $verb)
}

function openInExcel {
    # 表示中の Excel（無ければ新しく起動）でブックを開き、該当シートの該当セルを選択する。
    # 開き方（mode）: 通常 = そのまま開く / 読み取り専用 = ReadOnly で開く /
    #                 新規 = 元のファイルを基にした新しいブック（無題）として開く（元のファイルを占有しない）
    param (
        [string]$path,
        [string]$location,
        [string]$cell,
        [string]$mode = ${openModeNormal}
    )

    $excel = $null
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        # インデクサがバックグラウンドで使っている Excel は使わない
        if (!$excel.Visible) {
            $excel = $null
        }
    } catch {
        $excel = $null
    }
    if ($null -eq $excel) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.UserControl = $true
    }

    $book = $null
    if ($mode -eq ${openModeNew}) {
        # 元のファイルをテンプレートとして新しいブックを作る（読み込んだ後は元のファイルを開いたままにしない）
        $book = $excel.Workbooks.Add($path)
    } else {
        # 既に開いているブックがあれば、そのまま使う（同じブックを二重に開けないため。開き方も既に開いたときのまま）
        foreach ($openBook in $excel.Workbooks) {
            if ($openBook.FullName -eq $path) {
                $book = $openBook
                break
            }
        }
        if ($null -eq $book) {
            #   引数: Filename, UpdateLinks, ReadOnly
            $book = $excel.Workbooks.Open($path, [Type]::Missing, ($mode -eq ${openModeReadOnly}))
        }
    }
    $book.Activate()

    # 場所はシート名。名前が同じシートを選ぶ。
    # 以前の版のインデックスは、ファイル名に使えない文字を全角に置き換えてあるため、同じ名前のシートが無ければ
    # 全角に置き換えて一致するシートを選ぶ（`衝突"` と `衝突”` のように、置き換えると重なるシートがあるため、同じ名前を優先する）
    $target = $null
    $sameSafeName = $null
    foreach ($sheet in $book.Worksheets) {
        if ($sheet.Name -eq $location) {
            $target = $sheet
            break
        }
        if ($null -eq $sameSafeName -and (toSafeFileName $sheet.Name) -eq $location) {
            $sameSafeName = $sheet
        }
    }
    if ($null -eq $target) {
        $target = $sameSafeName
    }
    if ($null -ne $target) {
        $target.Activate()
        if ($cell) {
            $target.Range($cell).Select()
        }
    }
    if ($excel.WindowState -eq -4140) {
        # 最小化されていれば元に戻す（xlMinimized → xlNormal）
        $excel.WindowState = -4143
    }
    # Excel は Visible にして前面に出す（P/Invoke の SetForegroundWindow は実行時コンパイル（csc.exe）を無くすため廃止）
    $excel.Visible = $true
    try { $excel.ActiveWindow.Activate() } catch { }
}

