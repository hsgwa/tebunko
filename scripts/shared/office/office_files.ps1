# 取り込み対象にする Office ファイルの拡張子。

# 取り込み対象にする Office ファイルの拡張子（小文字）と、その種類（Excel・Word・PowerPoint）の表。拡張子と種類の対応はここにだけある。
# どのツールからも使うため shared/ に置く。種類の判定（getOfficeFileKind）は、tebunko/ と shared/ のどちらからもここを呼ぶ
${officeFileKinds} = [ordered]@{
    ".xlsx" = "Excel"; ".xlsm" = "Excel"; ".xls" = "Excel"; ".xlsb" = "Excel"
    ".docx" = "Word"; ".docm" = "Word"; ".doc" = "Word"
    ".pptx" = "PowerPoint"; ".pptm" = "PowerPoint"; ".ppt" = "PowerPoint"
}

# 取り込み対象にする Office ファイルの拡張子（小文字。上の表の拡張子）。
# インデクサ（indexer.ps1）の対象の判定で使う
${officeExtensions} = @(${officeFileKinds}.Keys)

function getOfficeFileKind {
    # ファイル名・パスの拡張子から、Office の種類（Excel / Word / PowerPoint）を返す。表（officeFileKinds）に無い拡張子は空文字。
    # 大文字小文字は区別しない。名前に使えない文字があっても例外にしない（拡張子は最後の「.」から後ろを文字列で切り出す）
    param (
        [string]$name
    )

    if ($name -notmatch '(\.[^.\\/]*)$') { return "" }
    $extension = $Matches[1].ToLowerInvariant()
    if (${officeFileKinds}.Contains($extension)) { return ${officeFileKinds}[$extension] }
    return ""
}
