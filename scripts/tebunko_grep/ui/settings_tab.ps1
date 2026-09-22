# ［8 設定］タブ（インデックスの置き場所・設定ファイルの場所）。文言と可否の判定は settings_view.ps1。

$script:restartRequested = $false  # インデックスの置き場所を変えたため、閉じたあと開き直す（gui.ps1）

function updateSettingsView {
    # インデックスの置き場所（work）と設定ファイルの場所を表示する
    $view = getWorkDirView ${workDir} (getDefaultWorkDir)
    $ui.WorkDirText.Text = $view.Path
    $ui.WorkDirNote.Text = $view.Note
    $ui.ResetWorkDirButton.Visibility = if ($view.CanReset) { "Visible" } else { "Collapsed" }

    $file = getSettingsFileView ${settingsFile} ${rootDir}
    $ui.SettingsFileText.Text = $file.Path
    $ui.SettingsFileNote.Text = $file.Note
}

function testWorkDirChangeable {
    # インデックス作成中は置き場所を変えない（インデクサが今の work に書いている。画面を使わずに起動したものも含む）
    if ((isIndexing) -or (findRunningIndexer)) {
        showMessage "インデックス作成中はインデックスの置き場所を変えられません。インデックス作成が終わるまでお待ちください（［中止］で止められます）。" "OK" "Warning" | Out-Null
        return $false
    }
    if ($script:indexBusy) {
        showMessage "前のインデックスの削除が終わるまでお待ちください。" "OK" "Warning" | Out-Null
        return $false
    }
    return $true
}

function chooseWorkDir {
    # ［変更…］。フォルダを選んで、インデックスの置き場所にする
    if (!(testWorkDirChangeable)) {
        return
    }
    $folder = selectFolder "インデックス・取り込み一覧・ログを置くフォルダを選んでください。" ${workDir}
    if ($null -eq $folder) {
        return
    }
    applyWorkDir $folder
}

function resetWorkDir {
    # ［既定に戻す］。設定ファイルと同じフォルダの work に戻す（無ければ作る）
    if (!(testWorkDirChangeable)) {
        return
    }
    $folder = getDefaultWorkDir
    try {
        [System.IO.Directory]::CreateDirectory($folder) | Out-Null
    } catch {
        # 作れなければ、書き込めないフォルダとして下の確認で伝える
    }
    applyWorkDir $folder
}

function applyWorkDir {
    # 確かめてから置き場所を保存し、画面を開き直す（work の中のファイルの場所は、読み込み時に決まるため）
    param (
        [string]$folder
    )

    $check = testWorkFolderChoice $folder ${workDir} (testWritableFolder $folder)
    if ($check.Kind -eq "same") {
        setStatus "今と同じ置き場所です"
        return
    }
    if ($check.Kind -eq "error") {
        showMessage $check.Message "OK" "Warning" | Out-Null
        return
    }

    $folder = normalizeFolderPath $folder
    $confirm = newWorkFolderConfirm $folder ${workDir} (Test-Path -LiteralPath (Join-Path $folder "index") -PathType Container)
    $facts = @($confirm.Facts | ForEach-Object {
        if ($_.Kind -eq "kept") { factKept $_.Title $_.Detail } else { factNext $_.Title $_.Detail }
    })
    $answer = showConfirm -heading $confirm.Heading -facts $facts -hint $confirm.Hint `
        -choices @(@{ Text = $confirm.ChoiceText; Value = "change" })
    if ($answer -ne "change") {
        return
    }

    writeWorkFolder $folder
    $script:restartRequested = $true
    $window.Close()
}

# ---- イベント ----

$ui.ChangeWorkDirButton.Add_Click({ safe { chooseWorkDir } })
$ui.ResetWorkDirButton.Add_Click({ safe { resetWorkDir } })
