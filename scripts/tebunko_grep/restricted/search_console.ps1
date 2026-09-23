# 制限モードのコンソール（画面層。手で確かめる）。制限言語モードで動く書き方だけで書く。
#
# 検索ワードを入れるだけで検索し、結果を Excel のブック（result_book.ps1）で開く。Excel が無い PC では、
# 結果をこのコンソールに番号付きの一覧で出し、番号で元のファイル・フォルダを開く。
# 条件（正規表現・大文字と小文字・対象ファイル・図形・コメント）と、検索するインデックスは、
# いつもの画面と同じく setting.config に保存する（どちらのモードで変えても、もう一方に引き継がれる）。
# 文言と入力の読み方は console_view.ps1 に置く。

function readRestrictedSetting {
    # 設定を読む。読めない（壊れた JSON など）ときは既定値で続ける
    param (
        [scriptblock]$read,
        $default
    )

    try {
        return (& $read)
    } catch {
        Write-Host "  設定を読めませんでした（$($_.Exception.Message)）。既定の値で続けます。" -ForegroundColor Yellow
        return $default
    }
}

function saveRestrictedSetting {
    # 設定を保存する。書けない（work・ツールのフォルダが読み取り専用など）ときは、そのことを出して続ける
    param (
        [scriptblock]$write
    )

    try {
        & $write
    } catch {
        Write-Host "  設定を保存できませんでした（$($_.Exception.Message)）。今回だけ変えて続けます。" -ForegroundColor Yellow
    }
}

function editRestrictedOption {
    # 条件を変えるメニュー
    param (
        [hashtable]$option
    )

    while ($true) {
        Write-Host ""
        foreach ($line in (getOptionMenuLines $option)) { Write-Host $line }
        $answer = Read-Host "番号"
        if ($answer.Trim() -eq "") {
            return
        }
        $number = parseMenuNumber $answer ${restrictedOptionItems}.Count
        if ($number -le 0) {
            Write-Host "  番号で選んでください。" -ForegroundColor Yellow
            continue
        }
        $key = ${restrictedOptionItems}[$number - 1].Key
        if ($key -eq "FileFilter") {
            Write-Host "  対象ファイル（例: *.xlsx;見積*;!*old*。; で区切り、! で始まるものは除く。空ですべて）"
            $option[$key] = (Read-Host "  対象ファイル").Trim()
        } else {
            $option[$key] = !$option[$key]
        }
        $copy = $option
        saveRestrictedSetting { writeSearchOption $copy }
    }
}

function editRestrictedIndexes {
    # 検索するインデックスを選ぶメニュー
    param (
        [object[]]$indexes
    )

    while ($true) {
        $excludes = @(readRestrictedSetting { readSearchExcludes } @())
        Write-Host ""
        if ($indexes.Count -eq 0) {
            Write-Host (describeRestrictedIndexes $indexes $excludes)
            return
        }
        foreach ($line in (getIndexMenuLines $indexes $excludes)) { Write-Host $line }
        $answer = (Read-Host "番号").Trim()
        if ($answer -eq "") {
            return
        }
        if ($answer -eq "a" -or $answer -eq "ａ") {
            $kept = @($excludes | Where-Object { $e = $_; @($indexes | Where-Object { $null -ne (getPathUnderFolder $e.Path $_.Path) }).Count -eq 0 })
            saveRestrictedSetting { writeSearchExcludes $kept }
            continue
        }
        $number = parseMenuNumber $answer $indexes.Count
        if ($number -le 0) {
            Write-Host "  番号で選んでください。" -ForegroundColor Yellow
            continue
        }
        $next = switchIndexExclude $indexes[$number - 1] $excludes
        saveRestrictedSetting { writeSearchExcludes $next }
    }
}

function showRestrictedHitList {
    # Excel が無い PC: ヒットを一覧で出し、番号で元のファイル（f番号でフォルダ）を開く
    param (
        [object[]]$hits,
        [hashtable]$maps
    )

    $shown = @($hits | Select-Object -First ${restrictedConsoleListMax})
    $width = 120
    try { $width = $Host.UI.RawUI.WindowSize.Width - 1 } catch { }
    for ($i = 0; $i -lt $shown.Count; $i++) {
        Write-Host (formatHitListLine ($i + 1) $shown[$i] $width)
    }
    if ($hits.Count -gt $shown.Count) {
        Write-Host "  （先頭の $($shown.Count) 件だけを出しました）"
    }
    while ($true) {
        $answer = Read-Host "番号で元のファイルを開きます（f番号: フォルダ、空で Enter: 戻る）"
        if ($answer.Trim() -eq "") {
            return
        }
        $choice = parseHitChoice $answer $shown.Count
        if ($null -eq $choice) {
            Write-Host "  番号で選んでください。" -ForegroundColor Yellow
            continue
        }
        $path = resolveSourcePath $shown[$choice.Number - 1] $maps
        if (!$path) {
            Write-Host "  元のフォルダが分かりません（いつもの画面で、インデックスの元のフォルダを選んでください）。" -ForegroundColor Yellow
            continue
        }
        if ($choice.Kind -eq "folder") {
            $path = getPathParent $path
        }
        if (!(Test-Path -LiteralPath $path)) {
            Write-Host "  見つかりません: $path" -ForegroundColor Yellow
            continue
        }
        Invoke-Item -LiteralPath $path
    }
}

function invokeRestrictedSearch {
    # 1 回の検索（検索して、結果を Excel で開くか、一覧で出す）
    param (
        [string]$word,
        [hashtable]$option,
        [object[]]$indexes,
        [bool]$hasExcel
    )

    $excludes = @(readRestrictedSetting { readSearchExcludes } @())
    $targets = @($indexes | Where-Object { testIndexChecked $_ $excludes })
    if ($targets.Count -eq 0) {
        Write-Host (describeRestrictedIndexes $indexes $excludes) -ForegroundColor Yellow
        return
    }
    Write-Host "  検索しています…"
    $files = getRestrictedTsvFiles $targets $excludes
    $progress = { param($done, $total) Write-Progress -Activity "検索しています" -Status "$done / $total" -PercentComplete ([int](100 * $done / $total)) }
    try {
        $result = searchRestricted $word $files (!$option.UseRegex) ([bool]$option.CaseSensitive) ${restrictedSearchLimit} ([string]$option.FileFilter) ([bool]$option.IncludeShapes) ([bool]$option.IncludeComments) $progress
    } catch {
        Write-Progress -Activity "検索しています" -Completed
        Write-Host "  $($_.Exception.Message)" -ForegroundColor Yellow
        return
    }
    Write-Progress -Activity "検索しています" -Completed
    foreach ($line in (getSearchSummaryLines $word $result ([bool]$option.UseRegex) ${restrictedSearchLimit})) {
        Write-Host "  $line"
    }
    if ($result.Hits.Count -eq 0) {
        return
    }
    $maps = @{}
    if (!$hasExcel) {
        showRestrictedHitList $result.Hits $maps
        return
    }
    $regex = (newSearchRegex $word (!$option.UseRegex) ([bool]$option.CaseSensitive)).Regex
    $info = getResultInfoItems $word $option (describeRestrictedIndexes $indexes $excludes) $result ${restrictedSearchLimit} (Get-Date)
    try {
        Write-Host "  結果のブックを作っています…"
        $book = saveResultBook $result $regex $info ${resultBookDir} $maps
        Invoke-Item -LiteralPath $book
        Write-Host "  Excel で開きました: $book"
    } catch {
        Write-Host "  結果のブックを作れませんでした（$($_.Exception.Message)）。一覧で出します。" -ForegroundColor Yellow
        showRestrictedHitList $result.Hits $maps
    }
}

function runRestrictedConsole {
    # 制限モードの入口（start.ps1 から呼ぶ）。終わるまで戻らない
    param (
        [bool]$hasExcel
    )

    removeOldResultBooks
    $option = readRestrictedSetting { readSearchOption } @{ UseRegex = $false; CaseSensitive = $false; FileFilter = ""; IncludeShapes = $true; IncludeComments = $true }
    while ($true) {
        $indexes = @(readRestrictedSetting { getSearchIndexes } @())
        $excludes = @(readRestrictedSetting { readSearchExcludes } @())
        Write-Host ""
        Write-Host "検索するインデックス: $(describeRestrictedIndexes $indexes $excludes)"
        Write-Host "条件: $(describeRestrictedOption $option)"
        $word = Read-Host "検索ワード（空で Enter: メニュー）"
        if ($word -ne "") {
            invokeRestrictedSearch $word $option $indexes $hasExcel
            continue
        }
        Write-Host ""
        Write-Host "  1. 検索の条件を変える"
        Write-Host "  2. 検索するインデックスを選ぶ"
        Write-Host "  3. インデックスを作成する（Word・PowerPoint・Excel）"
        Write-Host "  4. インデックスを管理する（追加・削除・取り込むものの選択）"
        Write-Host "  0. 終わる"
        switch ((Read-Host "番号（空で Enter: 戻る）").Trim()) {
            "1" { editRestrictedOption $option }
            "2" { editRestrictedIndexes $indexes }
            "3" { runRestrictedIndexing }
            "4" { editRestrictedCrawlTargets }
            "0" { return }
        }
    }
}
