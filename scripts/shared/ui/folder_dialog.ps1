# フォルダ選択（Windows 標準のダイアログ）と、フォルダのドラッグ＆ドロップ。

${folderPickOption} = [uint32]0x20                  # FOS_PICKFOLDERS（ファイルではなくフォルダを選ぶ）
${folderPickCanceled} = [int]0x800704C7             # HRESULT_FROM_WIN32(ERROR_CANCELLED)（［キャンセル］で閉じた）
${folderPickFileName} = "フォルダーの選択"          # 予備のダイアログのファイル名の欄に入れておく名前

function selectFolder {
    # Windows 標準のエクスプローラー形式のダイアログでフォルダを選んでもらい、選んだフォルダを返す（キャンセルなら $null）。
    #   description はタイトルバーに出す。initialPath が無ければ、その上の今もあるフォルダから開く。
    # ※エクスプローラー形式のダイアログ（COM の IFileOpenDialog）をフォルダ選択で開くには、
    #   インターフェースの定義が要る。自分で定義すると実行時コンパイル（csc.exe）が要るため（12.2）、
    #   WinForms（読み込み済み）が内部に持つ定義 FileDialogNative+IFileDialog をリフレクションで呼ぶ。
    #   内部の型（NonPublic）を呼ぶのはここだけ（tests\meta\safety.Tests.ps1 が確かめる）。
    #   内部の型が使えないとき（.NET の変更など）は、同じ見た目のファイルを開くダイアログで、フォルダの中に入って選んでもらう。
    param (
        [string]$description,
        [string]$initialPath,
        [System.Windows.Window]$owner = $window
    )

    $start = getExistingAncestorFolder (normalizeFolderPath $initialPath)
    $hwnd = [IntPtr]::Zero
    if ($owner) {
        $hwnd = (New-Object System.Windows.Interop.WindowInteropHelper $owner).Handle
    }
    try {
        return showFolderPicker $description $start $hwnd
    } catch {
        writeErrorLog "フォルダ選択（標準のフォルダ選択を開けないため、ファイルを開くダイアログで選んでもらう）" $_
    }
    return showFolderByFileDialog $description $start $hwnd
}

function showFolderPicker {
    # エクスプローラー形式のダイアログ（IFileOpenDialog をフォルダ選択にしたもの）。開けなければ例外を投げる
    param (
        [string]$description,
        [string]$start,
        [IntPtr]$hwnd
    )

    $flags = [System.Reflection.BindingFlags]"Instance,Public,NonPublic"
    $fileDialogType = [System.Windows.Forms.FileDialog]
    $assembly = $fileDialogType.Assembly
    $nativeType = $assembly.GetType("System.Windows.Forms.FileDialogNative+IFileDialog", $true)
    $itemType = $assembly.GetType("System.Windows.Forms.FileDialogNative+IShellItem", $true)
    $nameType = $assembly.GetType("System.Windows.Forms.FileDialogNative+SIGDN", $true)

    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = $description
    $dialog.InitialDirectory = $start

    $path = ""
    $native = $fileDialogType.GetMethod("CreateVistaDialog", $flags).Invoke($dialog, $null)
    try {
        # タイトル・開始フォルダなどを WinForms に設定させ、フォルダ選択の指定を足す
        $fileDialogType.GetMethod("OnBeforeVistaDialog", $flags).Invoke($dialog, @($native)) | Out-Null
        $options = [uint32]$fileDialogType.GetMethod("get_Options", $flags).Invoke($dialog, $null)
        $nativeType.GetMethod("SetOptions").Invoke($native, @($options -bor ${folderPickOption})) | Out-Null
        $hr = [int]$nativeType.GetMethod("Show").Invoke($native, @($hwnd))
        if ($hr -eq ${folderPickCanceled}) {
            return $null
        }
        if ($hr -ne 0) {
            throw "フォルダ選択のダイアログがエラーを返しました（0x$($hr.ToString('X8'))）"
        }
        # 選んだフォルダ（IShellItem）のファイルシステムのパス（SIGDN_FILESYSPATH）を読む。
        # out 引数は、渡した配列の同じ位置に入って戻る
        $resultArgs = [object[]]@($null)
        $nativeType.GetMethod("GetResult").Invoke($native, $resultArgs) | Out-Null
        $item = $resultArgs[0]
        try {
            $nameArgs = [object[]]@([Enum]::Parse($nameType, "SIGDN_FILESYSPATH"), $null)
            $itemType.GetMethod("GetDisplayName").Invoke($item, $nameArgs) | Out-Null
            $path = normalizeFolderPath ([string]$nameArgs[1])
        } finally {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($item) | Out-Null
        }
    } finally {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($native) | Out-Null
    }
    if ($path -eq "") {
        return $null
    }
    return $path
}

function showFolderByFileDialog {
    # 予備のフォルダ選択（標準のフォルダ選択を開けないときに使う）。見た目は同じエクスプローラー形式で、
    # 公開の API だけを使う。ファイルは出さず、選びたいフォルダの中に入って［開く］を押してもらう
    param (
        [string]$description,
        [string]$start,
        [IntPtr]$hwnd
    )

    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "$description（フォルダの中に入って［開く］を押してください）"
    $dialog.InitialDirectory = $start
    $dialog.FileName = ${folderPickFileName}
    $dialog.Filter = "フォルダー|*.tebunko-folder"   # どのファイルにも当たらない条件にして、フォルダだけを出す
    $dialog.CheckFileExists = $false
    $dialog.ValidateNames = $false
    $dialog.AddExtension = $false
    $ownerWindow = New-Object System.Windows.Forms.NativeWindow
    if ($hwnd -ne [IntPtr]::Zero) {
        $ownerWindow.AssignHandle($hwnd)
    }
    try {
        $result = $dialog.ShowDialog($ownerWindow)
    } finally {
        $ownerWindow.ReleaseHandle()
        $dialog.Dispose()
    }
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
    }
    return normalizeFolderPath ([System.IO.Path]::GetDirectoryName($dialog.FileName))
}

function getDroppedFolders {
    param (
        [System.Windows.DragEventArgs]$e
    )

    if (!$e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
        return @()
    }
    return @($e.Data.GetData([System.Windows.DataFormats]::FileDrop) | Where-Object { Test-Path -LiteralPath $_ -PathType Container })
}

function onFolderDragOver {
    param ($sender, [System.Windows.DragEventArgs]$e)

    $e.Effects = if ((getDroppedFolders $e).Count -gt 0) { "Copy" } else { "None" }
    $e.Handled = $true
}
