# 取り込み対象にする Office ファイルの拡張子。

# 取り込み対象にする Office ファイルの拡張子（小文字）。
# インデクサ（indexer.ps1）の対象の判定で使う
${officeExtensions} = @(
    ".xlsx", ".xlsm", ".xls", ".xlsb",
    ".docx", ".docm", ".doc",
    ".pptx", ".pptm", ".ppt"
)
