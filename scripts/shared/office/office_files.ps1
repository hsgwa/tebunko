# 変換対象にする Office ファイルの判定。

# 変換対象にする Office ファイルの拡張子（小文字）。
# 変換処理（indexer.ps1）の対象の判定と、フォルダ選択の一覧（Office ファイルかどうかの色分け）で使う
${officeExtensions} = @(
    ".xlsx", ".xlsm", ".xls", ".xlsb",
    ".docx", ".docm", ".doc",
    ".pptx", ".pptm", ".ppt"
)

function testOfficeFile {
    # ファイル名が変換対象の Office ファイル（Excel・Word・PowerPoint）かを返す
    param (
        [string]$name
    )

    if ($name -eq "") {
        return $false
    }
    return (${officeExtensions} -contains [System.IO.Path]::GetExtension($name).ToLowerInvariant())
}
