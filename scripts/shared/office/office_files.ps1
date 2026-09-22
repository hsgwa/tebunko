# 取り込み対象にする Office ファイルの判定。

# 取り込み対象にする Office ファイルの拡張子（小文字）。
# インデクサの対象の判定、比較のフォルダの列挙、フォルダ選択の一覧（Office ファイルかどうかの色分け）で使う
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

function findOfficeFiles {
    # フォルダ配下のOfficeファイルを検索し（インデクサのクロール・比較のフォルダの列挙）、@{ Root; Files; HasError（アクセスできないフォルダがあった） } を返す。
    # Root は実際に列挙したフォルダ（\\?\ の付かない通常のパス）。相対パスはこの Root から求める
    # （設定に書かれたパスは、末尾の \ ・ドライブ文字と UNC パスなど書き方が違うことがあるため、文字数で切り出さない）
    param (
        [string]$targetFolder
    )

    # \\?\ を付けないと、パスが約248文字を超えるフォルダの中を検索できない（アクセスできないフォルダ扱いになる）。
    # 見つかったファイルの FullName は \\?\ 付きになる（fromLongPath で戻す）
    $scanErrors = $null
    $root = (Resolve-Path -LiteralPath $targetFolder).ProviderPath
    $files = @(Get-ChildItem -LiteralPath (toLongPath $root) -Recurse -File -ErrorAction SilentlyContinue -ErrorVariable scanErrors |
        Where-Object { (${officeExtensions} -contains $_.Extension.ToLower()) -and -not $_.Name.StartsWith('~$') })

    return @{ Root = $root; Files = $files; HasError = (@($scanErrors).Count -gt 0) }
}
