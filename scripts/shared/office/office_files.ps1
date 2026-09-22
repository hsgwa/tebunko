# 取り込み対象にする Office ファイルの判定。

# 取り込み対象にする Office ファイルの拡張子（小文字）。
# インデクサ（indexer.ps1）の対象の判定と、フォルダ選択の一覧（Office ファイルかどうかの色分け）で使う
${officeExtensions} = @(
    ".xlsx", ".xlsm", ".xls", ".xlsb",
    ".docx", ".docm", ".doc",
    ".pptx", ".pptm", ".ppt"
)

function testOfficeFile {
    # ファイル名が取り込み対象の Office ファイル（Excel・Word・PowerPoint）かを返す
    param (
        [string]$name
    )

    if ($name -eq "") {
        return $false
    }
    return (${officeExtensions} -contains [System.IO.Path]::GetExtension($name).ToLowerInvariant())
}
