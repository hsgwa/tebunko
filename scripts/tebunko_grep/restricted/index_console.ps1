# 制限モードのインデックス管理とインデックス作成のコンソール（画面層。手で確かめる）。制限言語モードで動く書き方だけで書く。
#
# クロール対象フォルダ（インデックス）の追加・削除・取り込む／取り込まないの切り替えと、インデックス作成を行う。
# 設定（setting.config）と取り込み一覧は、いつもの画面と同じものを使う（どちらのモードで変えても、もう一方に引き継がれる）。
# 名前・場所の変更は、いつもの画面で行う（インデックスのフォルダの改名に、制限言語モードで使えない型が要るため）。
# 文言と入力の読み方は console_view.ps1 に置く。

function editRestrictedCrawlTargets {
    # インデックスを管理するメニュー
    while ($true) {
        $folders = @(readRestrictedSetting { getTargetFolders } @())
        Write-Host ""
        foreach ($line in (getCrawlMenuLines $folders)) { Write-Host $line }
        $answer = Read-Host "番号"
        if ($answer.Trim() -eq "") {
            return
        }
        $choice = parseCrawlChoice $answer $folders.Count
        if ($null -eq $choice) {
            Write-Host "  番号で選んでください。" -ForegroundColor Yellow
            continue
        }
        if ($choice.Kind -eq "add") {
            addRestrictedCrawlTarget $folders
            continue
        }
        $folder = $folders[$choice.Number - 1]
        if ($choice.Kind -eq "switch") {
            $folder.Enabled = !$folder.Enabled
            saveRestrictedSetting { writeTargetFolders $folders }
            continue
        }
        removeRestrictedCrawlTarget $folders $folder
    }
}

function addRestrictedCrawlTarget {
    # インデックスを追加する（一覧に加えるだけ。中のファイルを読むのはインデックス作成のとき）
    param (
        [object[]]$folders
    )

    Write-Host "  Office ファイル（Word・PowerPoint・Excel）の入っているフォルダのパスを入れてください（エクスプローラーのアドレス欄からコピーできます）。"
    $path = normalizeFolderPath (Read-Host "  フォルダ").Trim().Trim('"')
    if ($path -eq "") {
        return
    }
    if (!(Test-Path -LiteralPath $path -PathType Container)) {
        Write-Host "  フォルダが見つかりません: $path" -ForegroundColor Yellow
        return
    }
    $suggested = newIndexName $path @($folders | ForEach-Object { $_.Name })
    $name = (Read-Host "  インデックス名（空で Enter: $suggested）").Trim()
    if ($name -eq "") {
        $name = $suggested
    }
    $message = testRestrictedFolderInput $path $name $folders
    if ($message) {
        Write-Host "  $message" -ForegroundColor Yellow
        return
    }
    $next = @($folders) + @(New-Object PSObject -Property ([ordered]@{ Name = $name; Path = $path; Enabled = $true }))
    saveRestrictedSetting { writeTargetFolders $next }
    Write-Host "  インデックス [$name] を追加しました。メニューの［インデックスを作成する］で中身を取り込みます。"
}

function removeRestrictedCrawlTarget {
    # インデックスを削除する（一覧から外し、インデックスのフォルダと取り込み一覧の記録も削除する）
    param (
        [object[]]$folders,
        $folder
    )

    Write-Host "  インデックス [$($folder.Name)]（$($folder.Path)）を削除します。元のフォルダのファイルは削除しません。"
    if ((Read-Host "  削除してよければ y").Trim() -notmatch '^[yｙYＹ]$') {
        return
    }
    $reason = enterIndexingLock
    if ($reason) {
        Write-Host "  $reason" -ForegroundColor Yellow
        return
    }
    try {
        removeIndex $folder.Name
        $next = @($folders | Where-Object { $_ -ne $folder })
        saveRestrictedSetting { writeTargetFolders $next }
        Write-Host "  インデックス [$($folder.Name)] を削除しました。"
    } catch {
        Write-Host "  削除できませんでした（$($_.Exception.Message)）。" -ForegroundColor Yellow
    } finally {
        exitIndexingLock
    }
}

function runRestrictedIndexing {
    # インデックスを作成する（Word・PowerPoint だけ）
    Write-Host ""
    Write-Host "インデックスを作成します。制限モードで取り込めるのは Word・PowerPoint（.docx .docm .pptx .pptm）だけです。"
    Write-Host "Excel と旧形式（.xls .doc .ppt など）は、いつもの画面が使える PC でインデックスを作成すると取り込みます。"
    Write-Host "途中で止めるときは Ctrl+C を押します（取り込んだ分は残り、次回は続きから取り込みます）。"
    try {
        $result = invokeRestrictedIndexing
    } catch {
        Write-Host "  $($_.Exception.Message)" -ForegroundColor Yellow
        return
    }
    foreach ($line in (getIndexingSummaryLines $result)) {
        Write-Host "  $line"
    }
}
