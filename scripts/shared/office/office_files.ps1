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

# 元のファイルを開くときの開き方（検索の設定 openMode の値。比較から開くときも同じ値を使う）
${openModeNormal}   = "normal"    # そのまま開く（編集する）
${openModeReadOnly} = "readOnly"  # 読み取り専用で開く（誤って上書きしない）
${openModeNew}      = "new"       # 新規（元のファイルを基にした無題の文書）で開く。元のファイルを占有しない
${openModes}        = @(${openModeNormal}, ${openModeReadOnly}, ${openModeNew})

function describeIngestError {
    # 取り込み・比較の抽出で発生した例外から、画面に表示する失敗の理由（インデクサは取り込み一覧のエラー列にも書く）を返す。
    # Officeアプリのメッセージは分かりにくい（パスワード付きでも「入力したパスワードが間違っています」等）ため、
    # よくある原因は「原因（詳細: 元のメッセージ）」の形に言い換える
    param (
        [System.Exception]$exception
    )

    # メソッド呼び出しの例外（「"7" 個の引数を指定して "Open" を呼び出し中に例外が発生しました」等）は、中の例外が本来の理由
    $base = $exception.GetBaseException()
    $message = ($base.Message -replace "\s+", " ").Trim()
    $code = "{0:X8}" -f $base.HResult
    if ($message -eq "") {
        $message = "エラーコード 0x${code}"
    }

    # スクリプト自身が throw したメッセージは、利用者向けに書いてあるためそのまま使う
    if ($base.GetType() -eq [System.Management.Automation.RuntimeException]) {
        return $message
    }

    if ($message -match "パスワード|password") {
        # 元のメッセージ（パスワードが間違っています等）は、パスワードを入力していない利用者には誤解を招くため付けない
        return "読み取りパスワードが設定されているため開けません（パスワード付きのファイルは取り込めません）"
    }

    $cause = $null
    if ($base -is [System.IO.FileNotFoundException] -or $base -is [System.IO.DirectoryNotFoundException]) {
        $cause = "ファイルが見つかりません（取り込み中に移動・削除・名前変更された可能性があります）"
    } elseif ($base -is [System.UnauthorizedAccessException] -or $code -eq "80070005") {
        $cause = "ファイルを読むアクセス権がありません"
    } elseif ($base -is [System.IO.PathTooLongException]) {
        $cause = "パスが長すぎるため読めません"
    } elseif ($base -is [System.OutOfMemoryException]) {
        # 巨大なシート（テキストにして約 1GB 超）は、整形（prettyTsv）で一度に読み込めずメモリ不足になる
        $cause = "シート・文書が大きすぎて取り込めません（メモリが不足しました）"
    } elseif (@("80070020", "80070021") -contains $code) {
        # 共有違反・ロック違反
        $cause = "ほかのアプリ・利用者がファイルを使用中のため読めません（ファイルを閉じてから再取り込みしてください）"
    } elseif (@("80040154", "80080005", "800401F3") -contains $code) {
        # クラス未登録・サーバーの起動失敗・ProgID 不正
        $cause = "Officeアプリ（Excel・Word・PowerPoint）を起動できませんでした（インストール・ライセンス認証の状態を確認してください）"
    } elseif (@("800706BA", "800706BE", "80010105", "80010108") -contains $code) {
        # RPC サーバーを利用できない・呼び出し失敗・サーバーで例外・切断
        $cause = "Officeアプリが異常終了したか、内部でエラーが発生しました（ファイルが壊れている、または大きすぎる可能性があります）"
    } elseif (@("80010001", "8001010A") -contains $code) {
        # 呼び出しの拒否・処理中のため後で再試行
        $cause = "Officeアプリが応答しませんでした（Officeアプリでダイアログを表示中などの可能性があります）"
    } elseif ($base -is [System.IO.InvalidDataException] -or $base -is [System.Xml.XmlException] -or
              $message -match "ファイル形式|破損|壊れ|corrupt|file format") {
        $cause = "ファイルが壊れているか、拡張子と中身の形式が一致していません"
    }

    if ($cause) {
        return "${cause}（詳細: ${message}）"
    }
    return $message
}
