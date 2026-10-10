# 現行の画面を、状態ごとに全部（40 枚）自動で撮る道具。GUI を直す前の記録として docs\images\screens\ に置き、
# docs\design\gui\screens\ の写真として載せる。詳しい経緯・撮る状態の一覧・撮らないものは docs\design\gui\screens\index.md。
#
#   .\tools\capture_screens.ps1                 すべての状態を撮り直す
#   .\tools\capture_screens.ps1 -Only <ID...>   指定した状態だけを撮り直す（例: -Only index-tab/empty search-tab/results）
#   .\tools\capture_screens.ps1 -OutDir <フォルダ>   写真の置き場所（既定 docs\images\screens）
#
# -Only は、powershell.exe -File で呼ぶとカンマ区切りの 1 つの文字列で渡ることがある（tests\run.ps1 の -Tag と同じ）。
# どちらの呼び方でも同じになるよう、カンマでも分ける
#
# 流れ:
#   1. そろえる条件を確かめる（倍率 100%・ライトモード・画面 1 つ・Office が動いていない）。合わなければ理由を出して止まる
#   2. 一時フォルダを、空いているドライブの文字（Z: から下へ）に subst で割り当てる（利用者名を含むパスを画面に出さないため）
#   3. 場面ごとに、tebunko を写して起動し、画面遷移の一覧（docs\design\testing\gui-smoke.md）の操作を行い、状態ごとに撮る
#   4. 撮った窓の外（デスクトップ・後ろの窓）は無地で塗り、利用者名・コンピューター名・利用者のフォルダのパスを含む部品は
#      塗りつぶして C:\Users\test\... と描き直す
#   5. subst を外し、一時フォルダ・偽のプロセスを片づける
#
# 本体（scripts\）は変えない。判断だけの関数（ID の一覧・そろえる条件・塗りつぶし）は tools\capture\capture_common.ps1 に分けてあり、
# 単体テストは tests\tools\capture_screens.Tests.ps1。画面を動かす部分（このファイル）は実機で確かめる。
param (
    [string[]]$Only = @(),
    [string]$OutDir = ""
)

$ErrorActionPreference = "Stop"

$Only = @($Only | ForEach-Object { $_ -split "," } | ForEach-Object { $_.Trim() } | Where-Object { $_ })

$repoRoot = (Resolve-Path "$PSScriptRoot\..").Path
if (!$OutDir) { $OutDir = Join-Path $repoRoot "docs\images\screens" }

. "$PSScriptRoot\capture\capture_common.ps1"
. "$repoRoot\tests\helpers\load.ps1"
. "$repoRoot\tests\gui\gui_helpers.ps1"

# 状態の一覧（${captureIds}）は tools\capture\capture_common.ps1 にある（tests\meta\screens.Tests.ps1 も読むため）

# ---- そろえる条件 ----

function getCaptureAppliedDpi {
    $value = (Get-ItemProperty -LiteralPath "HKCU:\Control Panel\Desktop\WindowMetrics" -Name "AppliedDPI" -ErrorAction SilentlyContinue).AppliedDPI
    if (!$value) { return 96 }
    return [int]$value
}

function getCaptureLightTheme {
    $value = (Get-ItemProperty -LiteralPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -ErrorAction SilentlyContinue).AppsUseLightTheme
    if ($null -eq $value) { return $true }
    return ([int]$value -ne 0)
}

function assertCaptureEnvironment {
    $problems = getCaptureEnvironmentProblems `
        -AppliedDpi (getCaptureAppliedDpi) `
        -LightTheme (getCaptureLightTheme) `
        -ScreenCount ([System.Windows.Forms.Screen]::AllScreens.Count) `
        -OfficeProcessCount (@(Get-Process -Name EXCEL, WINWORD, POWERPNT -ErrorAction SilentlyContinue).Count)
    if ($problems.Count -gt 0) {
        throw "撮る前のそろえる条件に合いません。`n- " + ($problems -join "`n- ")
    }
}

# ---- 一時フォルダを subst のドライブに割り当てる（利用者名を含むパスを画面に出さないため） ----

function useCaptureDrive {
    param ([string]$Target)
    [void][IO.Directory]::CreateDirectory($Target)
    for ($code = [int][char]'Z'; $code -ge [int][char]'D'; $code--) {
        $letter = [char]$code
        if (Test-Path -LiteralPath "${letter}:\") { continue }
        & subst "${letter}:" $Target | Out-Null
        if (Test-Path -LiteralPath "${letter}:\") { return "${letter}:" }
    }
    throw "空いているドライブの文字（Z: から D:）が見つかりません"
}

function removeCaptureDrive {
    param ([string]$Drive)
    if ($Drive) { & subst $Drive /D 2>$null | Out-Null }
}

# ---- 撮る・塗る ----

$script:captureFont = New-Object Drawing.Font("Segoe UI", 9)
$script:captureBg = [Drawing.Color]::FromArgb(235, 235, 235)

function saveCaptureBitmap {
    # 撮った Bitmap を PNG で保存し、大きさを確かめて Sizes に足し、結果を 1 行ログに出す
    # （captureGuiState・captureStartupSplash の両方が使う）
    param ($Bitmap, [string]$Id, [string]$OutDir, [System.Collections.Generic.List[long]]$Sizes, [string[]]$Painted = @())

    $path = getCaptureImagePath -Id $Id -OutDir $OutDir
    [void][IO.Directory]::CreateDirectory((Split-Path $path -Parent))
    $Bitmap.Save($path, [Drawing.Imaging.ImageFormat]::Png)
    $bytes = (Get-Item -LiteralPath $path).Length
    if (!(testCaptureImageSize -Bytes $bytes)) {
        throw "写真が 200 KB を超えました（$Id・$([Math]::Round($bytes / 1KB)) KB）"
    }
    [void]$Sizes.Add($bytes)
    if ($Painted.Count -gt 0) {
        Write-Host "  [$Id] 塗った: $($Painted -join '・')"
    } else {
        Write-Host "  [$Id]"
    }
}

function captureGuiState {
    # 状態 1 つ分を撮る。Ids に Id が含まれないときは何もしない（-Only で絞ったとき、遷移の途中の操作はそのまま行う）。
    # Primary を渡さなければ本体の窓（$S.Window）。Extra はダイアログ・メニューなど、本体と重ねて撮るほかの窓
    param (
        $S,
        [string]$Id,
        [string[]]$Ids,
        [string]$OutDir,
        [string]$UserName,
        [string]$ComputerName,
        [string]$UserProfile,
        [System.Collections.Generic.List[long]]$Sizes,
        $Primary = $null,
        $Extra = @()
    )

    if ($Ids -notcontains $Id) { return }
    if (!$Primary) { $Primary = $S.Window }
    activateGuiWindow $S $Primary

    $windows = @($Primary) + @($Extra)
    # 窓が出た直後は、まだ大きさが決まっていない（GetWindowRect が 0 x 0）ことがあるため、決まるまで待つ
    foreach ($w in $windows) {
        [void](waitGui $S "窓の大きさが決まる（$Id）" ${guiDefaultTimeout} { $r = getGuiWindowRect $w; if ($r.Width -gt 0 -and $r.Height -gt 0) { $r } })
    }
    # 大きさが決まっても、前の窓の絵がまだ残って見えることがある（確認が続けて出るときなど）。強制的に描き直させて少し待つ
    foreach ($w in $windows) { redrawGuiWindow $w }
    Start-Sleep -Milliseconds 300
    # 窓そのものの範囲はネイティブの GetWindowRect（UI オートメーションの BoundingRectangle は、描画前の窓で 0 x 0 になることがある）
    $rects = @($windows | ForEach-Object { getGuiWindowRect $_ })
    # Measure-Object の Maximum・Minimum は [double] で返るため、大きさは必ず [int] に直す
    # （[double] のまま New-Object Drawing.Bitmap に渡すと、コンストラクターの解決を誤り「パラメーターが有効ではない」で失敗する）
    $left = [int](($rects | ForEach-Object { $_.Left } | Measure-Object -Minimum).Minimum)
    $top = [int](($rects | ForEach-Object { $_.Top } | Measure-Object -Minimum).Minimum)
    $right = [int](($rects | ForEach-Object { $_.Right } | Measure-Object -Maximum).Maximum)
    $bottom = [int](($rects | ForEach-Object { $_.Bottom } | Measure-Object -Maximum).Maximum)
    $width = $right - $left
    $height = $bottom - $top
    if ($width -le 0 -or $height -le 0) {
        throw "撮る範囲の大きさがおかしい（$Id）: $width x $height（rects: $(($rects | ForEach-Object { "$($_.Left),$($_.Top),$($_.Width),$($_.Height)" }) -join ' | ')）"
    }

    $bitmap = New-Object Drawing.Bitmap($width, $height)
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.CopyFromScreen($left, $top, 0, 0, (New-Object Drawing.Size($width, $height)))

        # 窓の外（デスクトップ・後ろの窓・通知）を無地で塗る
        $region = New-Object Drawing.Region((New-Object Drawing.Rectangle(0, 0, $width, $height)))
        $bgBrush = New-Object Drawing.SolidBrush($script:captureBg)
        foreach ($r in $rects) {
            $region.Exclude((New-Object Drawing.Rectangle(($r.Left - $left), ($r.Top - $top), $r.Width, $r.Height)))
        }
        $graphics.FillRegion($bgBrush, $region)

        # 利用者名・コンピューター名・利用者のフォルダのパスを含む部品を塗りつぶして描き直す
        $painted = New-Object System.Collections.ArrayList
        foreach ($w in $windows) {
            $candidates = @(findAllGui $w -Type Text) + @(findAllGui $w -Type Edit) + @(findAllGui $w -Type Pane)
            foreach ($el in $candidates) {
                $text = getGuiValue $el
                if (!(testCaptureSensitiveText -Text $text -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile)) { continue }
                $er = getGuiRect $el
                if ($er.Width -le 0 -or $er.Height -le 0) { continue }
                $local = New-Object Drawing.Rectangle(($er.Left - $left), ($er.Top - $top), $er.Width, $er.Height)
                $graphics.FillRectangle([Drawing.Brushes]::White, $local)
                $redacted = getCaptureRedactedText -Text $text -UserProfile $UserProfile -UserName $UserName -ComputerName $ComputerName
                $graphics.DrawString($redacted, $script:captureFont, [Drawing.Brushes]::Black, [float]$local.X, [float]($local.Y + 1))
                [void]$painted.Add($redacted)
            }
        }

        saveCaptureBitmap -Bitmap $bitmap -Id $Id -OutDir $OutDir -Sizes $Sizes -Painted @($painted)
    } finally {
        if ($region) { $region.Dispose() }
        if ($bgBrush) { $bgBrush.Dispose() }
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

# ---- 場面: window（起動中・メニュー・about・壊れた設定） ----

function captureBrokenConfigScene {
    param ($Ids, $Root, $OutDir, $UserName, $ComputerName, $UserProfile, $Sizes)
    if ($Ids -notcontains "window/settings-broken") { return }

    $dir = Join-Path $Root "broken"
    $tool = newGuiTool $dir
    [IO.File]::WriteAllText($tool.Config, "{ 壊れた json", (New-Object Text.UTF8Encoding($false)))

    $S = startGui $tool "window-settings-broken"
    invokeGuiScene $S {
        setGuiStep $S "壊れた設定ファイルの知らせ"
        $window = waitGuiWindow $S "壊れた設定ファイルの知らせ" -Text "設定ファイルが壊れていた"
        # 設定が既定に戻るため、本体の窓は既定のワークスペース（使い捨てに差し替えたもの）の中身で変わる
        # （利用者名を含まないため塗りつぶしでも消えない。計画で 5・22・35 を撮らないものにした理由と同じ）。
        # メッセージボックスの窓だけを撮り、本体は重ねない
        captureGuiState -S $S -Id "window/settings-broken" -Ids $Ids -OutDir $OutDir `
            -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes -Primary $window
        # 起動のごく早い段階で出るこのメッセージボックスは、OK への BM_CLICK だけ・Enter キーだけでは閉じないことがある
        # （手元では、両方を送ってようやく閉じた）。閉じるまで両方を送り直す
        waitGui $S "壊れた設定ファイルの知らせが閉じる" ${guiDefaultTimeout} {
            try {
                pressGuiMessageOk $window
            } catch { }
            Start-Sleep -Milliseconds 300
            !(@(getGuiOtherWindows $S) | Where-Object { (@(getGuiTexts $_) -join " ") -like "*設定ファイルが壊れていた*" })
        } | Out-Null

        # 既定に戻った設定で、既定のワークスペース（使い捨てに差し替えたもの）が空でなければ、続けて
        # 「空のフォルダではありません」の警告も出る。それが出ていれば閉じるだけにする
        $extraWarning = @(getGuiOtherWindows $S) | Select-Object -First 1
        if ($extraWarning) {
            closeGuiNativeMessage $S $extraWarning "既定のワークスペースの警告"
        }

        closeGui $S
    }
}

function captureStarterScene {
    # 起動中の表示・index-tab の空・追加・編集・チェック・削除の確認・作成の確認・完了・normal、search-tab の no-index、window のメニュー・about
    param ($Ids, $Root, $OutDir, $UserName, $ComputerName, $UserProfile, $Sizes)

    $dir = Join-Path $Root "starter"
    $tool = newGuiTool $dir
    $source = Join-Path (Join-Path $Root "starter-元") "営業"
    newGuiSourceFolder $source

    $S = startGui $tool "starter"
    try {
        invokeGuiScene $S {
            setGuiStep $S "起動時のタブ（インデックスが無い）"
            waitGui $S "［インデックス管理］が選ばれる" ${guiDefaultTimeout} { (getGuiSelectedTab $S) -eq "IndexTab" } | Out-Null
            captureGuiState -S $S -Id "index-tab/empty" -Ids $Ids -OutDir $OutDir `
                -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes

            setGuiStep $S "［検索］（インデックスが無い）"
            selectGuiTab $S "SearchTab" "GoIndexTabButton"
            captureGuiState -S $S -Id "search-tab/no-index" -Ids $Ids -OutDir $OutDir `
                -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes
            selectGuiTab $S "IndexTab" "NewIndexButton"

            setGuiStep $S "「バージョン情報」"
            clickGui $S $S.Window "AboutLink" "バージョン情報"
            $about = waitGuiWindow $S "「バージョン情報」のダイアログ" -Id "VersionText"
            captureGuiState -S $S -Id "window/about" -Ids $Ids -OutDir $OutDir `
                -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes -Extra @($about)
            clickGui $S $about "CloseButton" "［閉じる］"
            waitGuiWindowClosed $S $about "「バージョン情報」"

            setGuiStep $S "［＋ フォルダを追加］"
            clickGui $S $S.Window "NewIndexButton" "［＋ フォルダを追加］"
            $dialog = waitGuiWindow $S "追加のダイアログ" -Id "FolderBox"
            captureGuiState -S $S -Id "index-tab/add" -Ids $Ids -OutDir $OutDir `
                -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes -Extra @($dialog)

            setGuiStep $S "［＋ フォルダを追加］入力が足りないまま［OK］"
            clickGui $S $dialog "OkButton" "［OK］"
            waitGui $S "注意（ErrorText）" ${guiDefaultTimeout} { (getGuiText (findGui $dialog -Id "ErrorText")) -ne "" } | Out-Null
            captureGuiState -S $S -Id "index-tab/add-error" -Ids $Ids -OutDir $OutDir `
                -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes -Extra @($dialog)

            setGuiStep $S "フォルダを入れて追加"
            setGuiText $S (findGui $dialog -Id "FolderBox") $source
            waitGui $S "名前が自動で入る" ${guiDefaultTimeout} { (getGuiValue (findGui $dialog -Id "NameBox")) -eq "営業" } | Out-Null
            clickGui $S $dialog "OkButton" "［OK］"
            waitGuiWindowClosed $S $dialog "追加のダイアログ"
            $row = waitGui $S "一覧に加わる" ${guiDefaultTimeout} { @(getGuiGridRows (findGui $S.Window -Id "IndexGrid")) | Select-Object -First 1 }

            # 編集のダイアログ（index-tab/edit）は、［アクション ▾］の［編集…］で開く（行の右クリックとダブルクリックは
            # UI オートメーションからは開けない）
            setGuiStep $S "［アクション ▾］→［編集…］"
            selectGui $row
            clickGuiAction $S "ActionEdit" "［編集…］"
            $dialog = waitGuiWindow $S "インデックスの編集のダイアログ" -Id "FolderBox"
            captureGuiState -S $S -Id "index-tab/edit" -Ids $Ids -OutDir $OutDir `
                -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes -Extra @($dialog)
            clickGui $S $dialog "CancelButton" "［キャンセル］"
            waitGuiWindowClosed $S $dialog "編集のダイアログ"

            setGuiStep $S "行のチェックを付ける"
            $row = @(getGuiGridRows (findGui $S.Window -Id "IndexGrid"))[0]
            checkGuiRow $S $row
            captureGuiState -S $S -Id "index-tab/checked" -Ids $Ids -OutDir $OutDir `
                -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes

            setGuiStep $S "［削除］"
            clickGuiAction $S "ActionDelete" "［削除…］"
            $deleteConfirm = waitGuiWindow $S "削除の確認" -Id "HeadingText" -Text "インデックスを削除しますか"
            captureGuiState -S $S -Id "index-tab/delete-confirm" -Ids $Ids -OutDir $OutDir `
                -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes -Extra @($deleteConfirm)
            clickGuiByName $S $deleteConfirm "キャンセル"
            waitGuiWindowClosed $S $deleteConfirm "削除の確認"

            setGuiStep $S "［すべて更新］"
            clickGui $S $S.Window "IndexingButton" "［すべて更新］"
            $startConfirm = waitGuiWindow $S "取り込みの確認" -Id "StartButton" -Timeout ${guiIndexTimeout}
            captureGuiState -S $S -Id "index-tab/start-confirm" -Ids $Ids -OutDir $OutDir `
                -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes -Extra @($startConfirm)
            clickGui $S $startConfirm "StartButton" "確認の［更新を開始］"
            waitGuiWindowClosed $S $startConfirm "取り込みの確認"

            setGuiStep $S "取り込みの完了"
            waitGui $S "取り込みの完了" ${guiIndexTimeout} {
                (getGuiText (findGui $S.Window -Id "IndexSummaryText")) -like "*集約ファイル*" -and
                    (findGui $S.Window -Id "IndexingButton").Current.IsEnabled
            } | Out-Null
            captureGuiState -S $S -Id "index-tab/done" -Ids $Ids -OutDir $OutDir `
                -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes
            captureGuiState -S $S -Id "index-tab/normal" -Ids $Ids -OutDir $OutDir `
                -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes

            closeGui $S
        }
    } finally {
        stopGui $S
    }
}

function captureFailedScene {
    # index-tab/failed（失敗したファイルの一覧があり、タブに ⚠）
    param ($Ids, $Root, $OutDir, $UserName, $ComputerName, $UserProfile, $Sizes)
    if ($Ids -notcontains "index-tab/failed") { return }

    $dir = Join-Path $Root "failed"
    $tool = newGuiTool $dir
    $source = Join-Path (Join-Path $Root "failed-元") "営業"
    newGuiSourceFolder $source -Broken
    $config = readGuiConfig $tool
    $config | Add-Member -NotePropertyName targetFolders -NotePropertyValue @(@{ name = "営業"; path = $source; enabled = $true }) -Force
    writeGuiConfig $tool.Dir $config

    $S = startGui $tool "failed"
    invokeGuiScene $S {
        selectGuiTab $S "IndexTab" "IndexingButton"
        setGuiStep $S "取り込みの完了（失敗あり）"
        startGuiIndexing $S
        waitGui $S "取り込みの完了・失敗 1 件" ${guiIndexTimeout} {
            (getGuiText (findGui $S.Window -Id "IndexSummaryText")) -like "*集約ファイル*" -and
                (getGuiText (findGui $S.Window -Id "FailedHeading")) -like "*失敗したファイル 1 件*" -and
                (findGui $S.Window -Id "IndexingButton").Current.IsEnabled
        } | Out-Null
        captureGuiState -S $S -Id "index-tab/failed" -Ids $Ids -OutDir $OutDir `
            -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes
        closeGui $S
    }
}

function captureHeavyScene {
    # index-tab の running・stop-confirm・interrupted、settings-tab/running-warning、window/close-confirm
    param ($Ids, $Root, $OutDir, $UserName, $ComputerName, $UserProfile, $Sizes)

    $needed = @("index-tab/running", "index-tab/stop-confirm", "index-tab/interrupted", "settings-tab/running-warning", "window/close-confirm")
    if (@($needed | Where-Object { $Ids -contains $_ }).Count -eq 0) { return }

    $tooFast = "取り込みが終わってしまい、間に合わなかった。tools\capture_screens.ps1 の copies（ファイルの数）を増やす"
    $dir = Join-Path $Root "heavy"
    $tool = newGuiTool $dir @{ ingestThreads = 1 }
    $source = Join-Path (Join-Path $Root "heavy-元") "大量"
    newGuiSourceFolder $source -Copies 100
    $config = readGuiConfig $tool
    $config | Add-Member -NotePropertyName targetFolders -NotePropertyValue @(@{ name = "大量"; path = $source; enabled = $true }) -Force
    writeGuiConfig $tool.Dir $config

    $S = startGui $tool "heavy-1"
    invokeGuiScene $S {
        setGuiStep $S "取り込みを始める"
        startGuiIndexing $S
        waitGui $S "取り込み中" ${guiDefaultTimeout} { testGuiIndexing $S } | Out-Null
        if (!(testGuiIndexing $S)) { throw $tooFast }
        # 行の［中止］・全体の進み・ステータスバーの 1 行がそろうまで待ってから撮る
        $runningRow = @(getGuiGridRows (findGui $S.Window -Id "IndexGrid"))[0]
        waitGui $S "行に［中止］が出る" ${guiDefaultTimeout} { findGui $runningRow -Id "IndexRowStopButton" } | Out-Null
        waitGui $S "進み具合の件数が出る" ${guiIndexTimeout} { (getGuiIndexingBannerText $S) -like "*/*" } | Out-Null
        if (!(testGuiIndexing $S)) { throw $tooFast }
        captureGuiState -S $S -Id "index-tab/running" -Ids $Ids -OutDir $OutDir `
            -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes

        setGuiStep $S "取り込み中の［設定］の［変更…］"
        selectGuiTab $S "SettingsTab" "ChangeWorkspaceButton"
        clickGui $S $S.Window "ChangeWorkspaceButton" "［変更…］"
        $warning = waitGuiWindow $S "作成中の警告" -Text "更新中はワークスペースを変えられません"
        captureGuiState -S $S -Id "settings-tab/running-warning" -Ids $Ids -OutDir $OutDir `
            -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes -Extra @($warning)
        closeGuiNativeMessage $S $warning "作成中の警告"
        selectGuiTab $S "IndexTab" "IndexingStopButton"

        if (!(testGuiIndexing $S)) { throw $tooFast }
        setGuiStep $S "［中止］"
        clickGui $S $S.Window "IndexingStopButton" "［中止］"
        $stopConfirm = waitGuiWindow $S "中止の確認" -Id "HeadingText" -Text "中止しますか"
        captureGuiState -S $S -Id "index-tab/stop-confirm" -Ids $Ids -OutDir $OutDir `
            -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes -Extra @($stopConfirm)
        clickGuiByName $S $stopConfirm "中止する"
        waitGui $S "更新が止まる" ${guiIndexTimeout} {
            $b = findGui $S.Window -Id "IndexingButton"
            $b.Current.IsEnabled -and $b.Current.Name -eq "すべて更新" -and (findGui $S.Window -Id "IndexingResumeButton")
        } | Out-Null
        closeGui $S
    }

    if (@("index-tab/interrupted", "window/close-confirm") | Where-Object { $Ids -contains $_ }) {
        $S = startGui $tool "heavy-2"
        invokeGuiScene $S {
            $tooFastGuard = { if (!(testGuiIndexing $S)) { throw $tooFast } }
            setGuiStep $S "中断した取り込み"
            selectGuiTab $S "IndexTab" "IndexingButton"
            waitGui $S "まだ取り込んでいないファイルがある" ${guiDefaultTimeout} {
                (getGuiIndexingBannerText $S) -like "*更新を中断しました*"
            } | Out-Null
            captureGuiState -S $S -Id "index-tab/interrupted" -Ids $Ids -OutDir $OutDir `
                -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes

            setGuiStep $S "続きから再開して閉じる"
            startGuiIndexing $S
            waitGui $S "取り込み中" ${guiDefaultTimeout} { testGuiIndexing $S } | Out-Null
            if (!(testGuiIndexing $S)) { throw $tooFast }
            closeGuiWindowAsync $S $S.Window
            $closeConfirm = waitGuiWindow $S "閉じる確認" -Id "HeadingText" -Text "中止して閉じますか" -Guard $tooFastGuard
            captureGuiState -S $S -Id "window/close-confirm" -Ids $Ids -OutDir $OutDir `
                -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes -Extra @($closeConfirm)
            clickGuiByName $S $closeConfirm "閉じない"
            waitGuiWindowClosed $S $closeConfirm "閉じる確認"

            # 片づけ: もう一度閉じて、今度は「中止して閉じる」で終える（closeGui は、この確認を扱えない）
            setGuiStep $S "取り込みを止めて閉じる"
            closeGuiWindowAsync $S $S.Window
            $confirm2 = waitGuiWindow $S "閉じる確認" -Id "HeadingText" -Text "中止して閉じますか" -Guard $tooFastGuard
            clickGuiByName $S $confirm2 "中止して閉じる"
            waitGui $S "取り込みを止めて画面が終了する" ${guiIndexTimeout} -AllowExited { $S.Process.HasExited } | Out-Null
        }
    }
}

function captureSearchScene {
    # search-tab の状態（no-index・starter を除く）
    param ($Ids, $Root, $OutDir, $UserName, $ComputerName, $UserProfile, $Sizes)

    $needed = @(${captureIds} | Where-Object { $_ -like "search-tab/*" -and $_ -ne "search-tab/no-index" })
    if (@($needed | Where-Object { $Ids -contains $_ }).Count -eq 0) { return }

    $dir = Join-Path $Root "search"
    $tool = newGuiTool $dir
    newGuiSampleIndex $tool $Root "営業"
    # 件数の上限（10,000 件）を超えるための、語を繰り返すだけの本文インデックスのファイル
    $limitRoot = Join-Path $Root "search-limit-tsv"
    $limitLines = 1..10005 | ForEach-Object { "上限テスト" }
    newTsv "$limitRoot\大量\上限.xlsx\$(toIndexFileName "上限")" $limitLines
    [void](newContentIndexFiles $limitRoot "$($tool.Work)\content_index")

    $S = startGui $tool "search"
    invokeGuiScene $S {
        setGuiStep $S "起動時の［検索］"
        waitGui $S "［検索］が選ばれる" ${guiDefaultTimeout} { (getGuiSelectedTab $S) -eq "SearchTab" } | Out-Null
        captureGuiState -S $S -Id "search-tab/initial" -Ids $Ids -OutDir $OutDir `
            -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes

        setGuiStep $S "正規表現で不正な式"
        toggleGui (findGui $S.Window -Id "RegexCheck")
        setGuiText $S (findGui $S.Window -Id "WordBox") "("
        waitGui $S "吹き出し（RegexBalloon）が出る" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "RegexBalloonText")) -eq "正規表現が正しくありません" } | Out-Null
        captureGuiState -S $S -Id "search-tab/regex-error" -Ids $Ids -OutDir $OutDir `
            -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes
        setGuiText $S (findGui $S.Window -Id "WordBox") ""
        toggleGui (findGui $S.Window -Id "RegexCheck")

        setGuiStep $S "検索対象の［すべて解除］"
        clickGui $S $S.Window "UncheckAllIndexButton" "［解除］"
        waitGui $S "検索対象が「なし」になる" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "TargetCountText")) -like "検索対象 0 / *" } | Out-Null
        captureGuiState -S $S -Id "search-tab/tree-none" -Ids $Ids -OutDir $OutDir `
            -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes
        clickGui $S $S.Window "CheckAllIndexButton" "［すべて］"
        waitGui $S "検索対象が戻る" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "TargetCountText")) -match "^検索対象 (\d+) / \1$" } | Out-Null

        $search = {
            param ($word)
            setGuiText $S (findGui $S.Window -Id "WordBox") $word
            waitGuiEnabled $S (findGui $S.Window -Id "SearchButton") "［検索］"
            clickGui $S $S.Window "SearchButton" "［検索］"
        }
        $summary = { getGuiText (findGui $S.Window -Id "SummaryText") }
        $hitRows = { getGuiHitRows (findGui $S.Window -Id "ResultGrid") }

        setGuiStep $S "見つからない検索"
        & $search "存在しないはずのことば"
        waitGui $S "見つからない（ステータスバーに 0 件）" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "StatusText")) -like "検索しました：* 0 件" } | Out-Null
        captureGuiState -S $S -Id "search-tab/no-results" -Ids $Ids -OutDir $OutDir `
            -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes

        setGuiStep $S "件数の上限で打ち切り"
        & $search "上限テスト"
        waitGui $S "上限で打ち切り" ${guiIndexTimeout} { (getGuiText (findGui $S.Window -Id "StatusText")) -like "*件を超えたため、ここで打ち切りました*" } | Out-Null
        captureGuiState -S $S -Id "search-tab/limit" -Ids $Ids -OutDir $OutDir `
            -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes

        setGuiStep $S "検索して結果を選ぶ"
        & $search "単価"
        waitGui $S "該当 2 件" ${guiDefaultTimeout} { (& $summary) -like "一致 2 件（*" } | Out-Null
        clickGui $S $S.Window "ExpandAllButton" "［すべて開く］"
        $row = waitGui $S "結果の行" ${guiDefaultTimeout} { @(& $hitRows) | Select-Object -Last 1 }
        selectGui $row
        waitGui $S "プレビューが出る" ${guiDefaultTimeout} {
            @(findAllGui (findGui $S.Window -Id "PreviewScroll") -Type Text | Where-Object { $_.Current.Name -eq "りんご" }).Count -gt 0
        } | Out-Null
        # 行を選ぶと、部品を選ぶだけで出るツールヒント（元のファイルのパス）が重なって写ることがある。
        # フォーカスを移す・ツールヒントの窓が消えるまで待つ、の両方を試したが消えなかった（架空の subst ドライブの
        # パスで個人情報ではないため、写り込んだまま撮る）
        captureGuiState -S $S -Id "search-tab/results" -Ids $Ids -OutDir $OutDir `
            -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes

        # 結果の行・プレビューの右クリックのメニューは、UI オートメーションの Invoke でも、ネイティブの
        # WM_CONTEXTMENU でも開けなかった（フォーカスで出るツールヒントを拾うだけだった）ため、撮らない
        # （docs\design\gui\screens\index.md「撮らないもの」）

        setGuiStep $S "すべて折りたたむ・絞り込み"
        clickGui $S $S.Window "CollapseAllButton" "［すべて折りたたむ］"
        waitGui $S "結果の行が隠れる" ${guiDefaultTimeout} { @(& $hitRows).Count -eq 0 } | Out-Null
        captureGuiState -S $S -Id "search-tab/collapsed" -Ids $Ids -OutDir $OutDir `
            -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes
        clickGui $S $S.Window "ExpandAllButton" "［すべて開く］"
        waitGui $S "結果の行が戻る" ${guiDefaultTimeout} { @(& $hitRows).Count -eq 2 } | Out-Null
        setGuiText $S (findGui $S.Window -Id "FilterBox") "議事録"
        waitGui $S "絞り込んだ件数" ${guiDefaultTimeout} { @(& $hitRows).Count -eq 1 } | Out-Null
        captureGuiState -S $S -Id "search-tab/filtered" -Ids $Ids -OutDir $OutDir `
            -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes
        setGuiText $S (findGui $S.Window -Id "FilterBox") ""

        setGuiStep $S "元のファイルが見つからない確認"
        waitGui $S "結果の行が戻る" ${guiDefaultTimeout} { @(& $hitRows).Count -eq 2 } | Out-Null
        selectGui (@(& $hitRows) | Select-Object -Last 1)
        waitGuiEnabled $S (findGui $S.Window -Id "OpenButton") "［開く］"
        clickGui $S $S.Window "OpenButton" "［開く］"
        $missing = waitGuiWindow $S "見つからない確認" -Id "HeadingText" -Text "が見つかりません"
        captureGuiState -S $S -Id "search-tab/missing-source" -Ids $Ids -OutDir $OutDir `
            -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes -Extra @($missing)
        clickGuiByName $S $missing "キャンセル"
        waitGuiWindowClosed $S $missing "見つからない確認"

        setGuiStep $S "最小の大きさ"
        resizeGuiWindow $S 1024 640
        Start-Sleep -Milliseconds 300
        captureGuiState -S $S -Id "search-tab/min-width" -Ids $Ids -OutDir $OutDir `
            -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes

        closeGui $S
    }
}

function captureSettingsScene {
    # settings-tab の normal・empty-confirm・nonempty-confirm・index-confirm・invalid-warning
    param ($Ids, $Root, $OutDir, $UserName, $ComputerName, $UserProfile, $Sizes)

    $needed = @("settings-tab/normal", "settings-tab/empty-confirm", "settings-tab/nonempty-confirm", "settings-tab/index-confirm", "settings-tab/invalid-warning")
    if (@($needed | Where-Object { $Ids -contains $_ }).Count -eq 0) { return }

    $dir = Join-Path $Root "settings"
    $tool = newGuiTool $dir
    newGuiSampleIndex $tool $Root "営業"
    $emptyDir = Join-Path $Root "settings-空"
    $nonEmptyDir = Join-Path $Root "settings-空でない"
    $sharedDir = Join-Path $Root "settings-共有"
    [void][IO.Directory]::CreateDirectory($emptyDir)
    [void][IO.Directory]::CreateDirectory($nonEmptyDir)
    Set-Content -LiteralPath "$nonEmptyDir\メモ.txt" -Value "メモ" -Encoding UTF8
    newGuiSampleIndex @{ Work = $sharedDir } (Join-Path $Root "settings-共有-元") "共有"

    $S = startGui $tool "settings"
    invokeGuiScene $S {
        $changeWorkspace = {
            param ($path)
            clickGui $S $S.Window "ChangeWorkspaceButton" "［変更…］"
            useGuiFolderPicker $S $path
        }

        setGuiStep $S "［設定］（既定でないワークスペース）"
        selectGuiTab $S "SettingsTab" "ChangeWorkspaceButton"
        captureGuiState -S $S -Id "settings-tab/normal" -Ids $Ids -OutDir $OutDir `
            -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes

        setGuiStep $S "使えないフォルダ（今のインデックスの中）"
        & $changeWorkspace "$($tool.Work)\content_index\営業"
        $invalid = waitGuiWindow $S "使えないフォルダの警告" -Text "インデックスのフォルダの中です"
        captureGuiState -S $S -Id "settings-tab/invalid-warning" -Ids $Ids -OutDir $OutDir `
            -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes -Extra @($invalid)
        closeGuiNativeMessage $S $invalid "使えないフォルダの警告"

        setGuiStep $S "空のフォルダの確認"
        & $changeWorkspace $emptyDir
        $emptyConfirm = waitGuiWindow $S "ワークスペースを変える確認" -Id "HeadingText" -Text "へ移動します"
        captureGuiState -S $S -Id "settings-tab/empty-confirm" -Ids $Ids -OutDir $OutDir `
            -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes -Extra @($emptyConfirm)
        clickGuiByName $S $emptyConfirm "キャンセル"
        waitGuiWindowClosed $S $emptyConfirm "ワークスペースを変える確認"

        setGuiStep $S "空でないフォルダの確認"
        & $changeWorkspace $nonEmptyDir
        $nonEmptyConfirm = waitGuiWindow $S "空でないフォルダの確認" -Id "HeadingText" -Text "空ではありません"
        captureGuiState -S $S -Id "settings-tab/nonempty-confirm" -Ids $Ids -OutDir $OutDir `
            -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes -Extra @($nonEmptyConfirm)
        clickGuiByName $S $nonEmptyConfirm "キャンセル"
        waitGuiWindowClosed $S $nonEmptyConfirm "空でないフォルダの確認"

        setGuiStep $S "インデックスのあるフォルダの確認"
        & $changeWorkspace $sharedDir
        $indexConfirm = waitGuiWindow $S "インデックスのあるフォルダの確認" -Id "HeadingText" -Text "すでにインデックスがあります"
        captureGuiState -S $S -Id "settings-tab/index-confirm" -Ids $Ids -OutDir $OutDir `
            -UserName $UserName -ComputerName $ComputerName -UserProfile $UserProfile -Sizes $Sizes -Extra @($indexConfirm)
        clickGuiByName $S $indexConfirm "キャンセル"
        waitGuiWindowClosed $S $indexConfirm "インデックスのあるフォルダの確認"

        closeGui $S
    }
}

function newCaptureLeftoverFile {
    # 偽の行（TEBUNKO_GUI_LEFTOVER_FILE が指す JSON）を作る。本物の Office は使わない。Kind: basic（3 件）・many（12 件）・partial（3 件。Word の 1 件が確認の後に変わる）
    param ([string]$Path, [string]$Kind)

    $rows = New-Object System.Collections.ArrayList
    if ($Kind -eq "many") {
        $apps = @(1..7 | ForEach-Object { "EXCEL" }) + @(1..3 | ForEach-Object { "WINWORD" }) + @(1..2 | ForEach-Object { "POWERPNT" })
        for ($i = 0; $i -lt $apps.Count; $i++) {
            [void]$rows.Add(@{ Id = 12840 + $i * 311; ProcessName = $apps[$i]; StartTime = "2030-10-04T18:{0:00}:00" -f (20 + $i) })
        }
    } else {
        [void]$rows.Add(@{ Id = 12840; ProcessName = "EXCEL"; StartTime = "2030-10-04T18:32:00" })
        [void]$rows.Add(@{ Id = 15012; ProcessName = "EXCEL"; StartTime = "2030-10-04T18:32:00" })
        $word = @{ Id = 9316; ProcessName = "WINWORD"; StartTime = "2030-10-04T18:41:00" }
        if ($Kind -eq "partial") { $word.StopStatus = "Changed" }
        [void]$rows.Add($word)
    }
    [IO.File]::WriteAllText($Path, (ConvertTo-Json -InputObject @($rows.ToArray())), (New-Object Text.UTF8Encoding($false)))
}

function captureLeftoverScene {
    # window/leftover 系（起動時の「前回残った Office」の確認と、終了したあとのステータス）。
    # 偽の行を環境変数で差し込む（本物の Office には触れない。終了も偽の結果を作るだけ）
    param ($Ids, $Root, $OutDir, $UserName, $ComputerName, $UserProfile, $Sizes)

    $needed = @("window/leftover", "window/leftover-open", "window/leftover-many", "window/leftover-search", "window/leftover-killed", "window/leftover-partial")
    if (@($needed | Where-Object { $Ids -contains $_ }).Count -eq 0) { return }

    $run = {
        param ([string]$Name, [string]$Kind, [scriptblock]$Scene)
        $tool = newGuiTool (Join-Path $Root "leftover-$Name")
        newGuiSampleIndex $tool $Root "営業"
        $file = Join-Path $Root "leftover-$Name.json"
        newCaptureLeftoverFile $file $Kind
        $env:TEBUNKO_GUI_LEFTOVER_FILE = $file
        try {
            $S = startGui $tool "leftover-$Name"
        } finally {
            Remove-Item Env:\TEBUNKO_GUI_LEFTOVER_FILE -ErrorAction SilentlyContinue
        }
        invokeGuiScene $S { & $Scene $S }
    }
    $common = @{ Ids = $Ids; OutDir = $OutDir; UserName = $UserName; ComputerName = $ComputerName; UserProfile = $UserProfile; Sizes = $Sizes }

    & $run "basic" "basic" {
        param ($S)
        $dialog = waitGuiWindow $S "Office の終了" -Id "HeadingText" -Text "残ったまま動いています"
        captureGuiState -S $S -Id "window/leftover-search" @common -Extra @($dialog)
        captureGuiState -S $S -Id "window/leftover" @common -Primary $dialog
        toggleGui (findGui $dialog -Id "LeftoverDetailToggle")
        waitGui $S "詳細が開く" ${guiDefaultTimeout} { (@(getGuiTexts $dialog) -join " ") -like "*12840*" } | Out-Null
        captureGuiState -S $S -Id "window/leftover-open" @common -Primary $dialog
        clickGui $S $dialog "LeftoverCancelButton" "［今回は終了しない］"
        waitGuiWindowClosed $S $dialog "Office の終了"
        closeGui $S
    }
    & $run "many" "many" {
        param ($S)
        $dialog = waitGuiWindow $S "Office の終了" -Id "HeadingText" -Text "残ったまま動いています"
        toggleGui (findGui $dialog -Id "LeftoverDetailToggle")
        waitGui $S "詳細が開く" ${guiDefaultTimeout} { (@(getGuiTexts $dialog) -join " ") -like "*12840*" } | Out-Null
        captureGuiState -S $S -Id "window/leftover-many" @common -Primary $dialog
        clickGui $S $dialog "LeftoverCancelButton" "［今回は終了しない］"
        waitGuiWindowClosed $S $dialog "Office の終了"
        closeGui $S
    }
    & $run "killed" "basic" {
        param ($S)
        answerGuiConfirm $S "Office の終了" "残ったまま動いています" "終了する"
        waitGui $S "結果がステータスに出る" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "StatusText")) -like "Office を*" } | Out-Null
        captureGuiState -S $S -Id "window/leftover-killed" @common
        closeGui $S
    }
    & $run "partial" "partial" {
        param ($S)
        answerGuiConfirm $S "Office の終了" "残ったまま動いています" "終了する"
        waitGui $S "結果がステータスに出る" ${guiDefaultTimeout} { (getGuiText (findGui $S.Window -Id "StatusText")) -like "Office を*" } | Out-Null
        captureGuiState -S $S -Id "window/leftover-partial" @common
        closeGui $S
    }
}

# ---- メイン ----

$ids = resolveCaptureIds -Only $Only -Ids ${captureIds}
assertCaptureEnvironment

$userName = $env:USERNAME
$computerName = $env:COMPUTERNAME
$userProfile = [System.Environment]::GetFolderPath("UserProfile")

$tempBase = Join-Path ([IO.Path]::GetTempPath()) ("tebunko-capture-" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
$drive = $null
$sizes = New-Object System.Collections.Generic.List[long]
# 撮る間は、使い捨ての既定のワークスペースに差し替える（環境変数 TEBUNKO_DEFAULT_WORKSPACE。起動する画面にも引き継がれる）。
# 設定が既定に戻る場面（captureBrokenConfigScene）も、利用者の既定のワークスペースには触れない。終わったら元に戻し、既定のワークスペースの前後を比べる
$workspaceGuard = startWorkspaceGuard -Prefix "tebunko-capture-ws"
try {
    $drive = useCaptureDrive $tempBase
    $root = "$drive\"
    Write-Host "撮る状態: $($ids.Count) 件（$($drive) を写す先にする）"

    captureBrokenConfigScene $ids $root $OutDir $userName $computerName $userProfile $sizes
    captureStarterScene      $ids $root $OutDir $userName $computerName $userProfile $sizes
    captureFailedScene       $ids $root $OutDir $userName $computerName $userProfile $sizes
    captureHeavyScene        $ids $root $OutDir $userName $computerName $userProfile $sizes
    captureSearchScene       $ids $root $OutDir $userName $computerName $userProfile $sizes
    captureSettingsScene     $ids $root $OutDir $userName $computerName $userProfile $sizes
    captureLeftoverScene     $ids $root $OutDir $userName $computerName $userProfile $sizes

    $total = ($sizes | Measure-Object -Sum).Sum
    Write-Host "撮った写真: $($sizes.Count) 枚・合計 $([Math]::Round($total / 1KB)) KB"
    if ($total -gt 5MB) {
        throw "写真の合計が 5 MB を超えました（$([Math]::Round($total / 1MB, 1)) MB）。範囲・形式を見直してください。"
    }
} finally {
    if ($drive) { removeCaptureDrive $drive }
    Remove-Item -LiteralPath $tempBase -Recurse -Force -ErrorAction SilentlyContinue
    $guardReport = finishWorkspaceGuard $workspaceGuard
}
if ($guardReport.ExitCode -ne 0) {
    throw ($guardReport.Lines -join "`n")
}
