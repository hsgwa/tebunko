# 元のファイルを開くとき（ネットワークにあるときの確認）の文言（判断層）。
# 画面に触らないため、そのままテストできる（tests\tebunko\ui\search\open_source_view.Tests.ps1）。

function getSourceCheckingStatus {
    # 元のファイルの場所を確かめている間のステータス（ネットワークにあるときだけ出す。ローカルはその場で開くため出さない）
    param (
        [string]$path
    )

    return "元のファイルを確かめています…：${path}（共有フォルダに接続できないときは、しばらくかかります）"
}

function getSourceLookingStatus {
    # 元のファイルの場所の記録を読んでいる間のステータス（ネットワークのワークスペースにあるときだけ出す）
    return "パスを調べています…（共有フォルダに接続できないときは、しばらくかかります）"
}

function getSourceLookupFailedStatus {
    # 元のファイルの場所の記録を読めなかったときのステータス
    param (
        [string]$message
    )

    if ($message) {
        return "元のファイルの場所を調べられませんでした：${message}"
    }
    return "元のファイルの場所を調べられませんでした"
}

function getSourceBookRejectedStatus {
    # インデックスの「ファイル名」が、開いてよい名前（testSourceBookName）でなかったときのステータス。開かない
    param (
        [string]$book
    )

    return "取り込み対象のファイル名ではないため、開きません：$(formatDisplayText $book)"
}

function getSourceNotFoundStatus {
    # 見つからない・確認で選ばなかったときのステータス
    param (
        [string]$path
    )

    return "元のファイルが見つかりません：$(formatDisplayText $path)"
}

function getSourceUnreachableStatus {
    # 接続できないと分かったときのステータス
    param (
        [string]$folder
    )

    return "元のフォルダに接続できません：${folder}"
}

function getSourceConnectFailureStatus {
    # 接続できない・その他のときの、ダイアログを出した直後のステータス
    param (
        [string]$state,
        [string]$folder,
        [string]$message
    )

    if ($state -eq "Unreachable") {
        return getSourceUnreachableStatus $folder
    }
    return "元のファイルを確かめられませんでした：${message}"
}

function getSourceConnectFailureDialog {
    # 接続できない・その他のときの確認ダイアログの中身。
    #   state: "Unreachable"（接続できない）・"Other"（その他。アクセス拒否・ログオンの失敗など）
    # 返すもの: @{ Heading; Title; Detail; Hint }（Heading はダイアログの見出し、Title・Detail は知らせの1行、Hint は選択肢の説明）
    param (
        [string]$book,
        [string]$state,
        [string]$folder,
        [string]$message
    )

    $heading = "${book} を開けません"
    if ($state -eq "Unreachable") {
        return @{
            Heading = $heading
            Title   = "元のフォルダに接続できません"
            Detail  = $folder
            Hint    = "ネットワーク・VPN の接続を確かめてから、もう一度開いてください"
        }
    }
    return @{
        Heading = $heading
        Title   = "元のファイルを確かめられませんでした"
        Detail  = $message
        Hint    = "アクセスの権限・サインインを確かめてください"
    }
}

# ---- もらったインデックスの元のフォルダを確かめる ----

# マクロ（ファイルの中に仕込まれたプログラム）を持てる形式。もらったインデックスのものは、読み取り専用で開く。
# 取り込みの対象の拡張子（${officeExtensions}）のうち、マクロを持てるものを全部入れる（.xltm・.dotm などは取り込みの対象外のため入れない）
${macroCapableExtensions} = @(".xlsm", ".xlsb", ".xls", ".docm", ".doc", ".pptm", ".ppt")

function testNameInList {
    # name が list のどれかと等しいか。大文字・小文字を区別せず、文字の並びそのもので比べる（OrdinalIgnoreCase。
    # PowerShell の -ieq・-contains はカルチャに従い、ソフトハイフンなどの見えない文字を無視して一致してしまう。
    # 設定の名前の辞書（source_map.ps1）と同じ比べ方にそろえる）
    param (
        [string]$name,
        [string[]]$list = @()
    )

    foreach ($item in @($list)) {
        if ([string]::Equals($item, $name, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }
    return $false
}

function testSourceReceived {
    # もらったインデックス（このワークスペースで自分が作ったのではないもの）か。名前が分からないときも、もらったものとして扱う。
    #   crawledNames: 自分で作ったインデックスの名前（設定の targetFolders）
    param (
        [string]$name,
        [string[]]$crawledNames = @()
    )

    if ($name -eq "") {
        return $true
    }
    return !(testNameInList $name $crawledNames)
}

function testSourceNeedsConfirm {
    # 元のフォルダに触れる前に、利用者に確かめるかを返す。
    # 元のフォルダが分かっていて（known）、そのインデックス名が確かめ済みの名前（設定の targetFolders・indexSources）に無いとき真。
    # 元のフォルダが分からないときは偽（フォルダを選んでもらう流れに進む）。名前が空のときは、確かめ済みと言えないので毎回確かめる
    param (
        [string]$name,
        [bool]$known,
        [string[]]$confirmedNames = @()
    )

    if (!$known) {
        return $false
    }
    if ($name -eq "") {
        return $true
    }
    return !(testNameInList $name $confirmedNames)
}

function getSourceConfirmDialog {
    # もらったインデックスの元のフォルダを確かめるダイアログの中身。
    #   mayConnect: 元のフォルダがローカルのドライブと分からないか（真なら、接続とサインイン情報について注意を足す）
    #   recordable: 名前を記録できるか（testSourceNameRecordable。偽なら「次からは聞きません」と言わず、毎回確かめると伝える）
    # 返すもの: @{ Heading; Title; Detail; Hint; UseText; PickText }
    param (
        [string]$book,
        [string]$name,
        [string]$folder,
        [bool]$mayConnect,
        [Parameter(Mandatory)][bool]$recordable
    )

    $hint = "［このフォルダを使う］を選ぶと、次からはこのインデックスについて聞きません"
    if (!$recordable) {
        $hint = "このインデックスの名前は記録できないため、［このフォルダを使う］を選んでも、開くたびに確かめます"
    }
    if ($mayConnect) {
        $hint = "開くと、このフォルダに接続し、Windows のサインイン情報が送られることがあります。心当たりのない場所なら、［フォルダを選ぶ］で別のフォルダを選んでください。" + $hint
    }
    return @{
        Heading  = "$(formatDisplayText $book) の元のフォルダを確かめてください"
        Title    = "インデックス [${name}] に書かれた元のフォルダ"
        Detail   = (formatDisplayText $folder)
        Hint     = $hint
        UseText  = "このフォルダを使う"
        PickText = "フォルダを選ぶ"
    }
}

function getSourceConfirmCanceledStatus {
    # 元のフォルダの確認でキャンセルしたときのステータス
    return "開くのをやめました"
}

function getSourceOpenMode {
    # 元のファイルの開き方を決める。もらったインデックス（received）のマクロを持てる形式（${macroCapableExtensions}）は、
    # 開き方が通常・新規でも読み取り専用にする。
    # 返すもの: @{ Mode; Notice; Strict }
    #   Notice: 開けたときに添えて出す知らせ（読み取り専用に替えなかったとき、もう読み取り専用だったときは空）
    #   Strict: 読み取り専用で開けないとき、通常に落とさず開かないこと
    param (
        [string]$book,
        [string]$mode,
        [bool]$received
    )

    # Windows は名前の終わりの . と空白を落として開く（evil.xlsm. は evil.xlsm として開く）ので、落としてから拡張子を見る
    $extension = [System.IO.Path]::GetExtension(([string]$book).TrimEnd('.', ' ')).ToLowerInvariant()
    if (!$received -or !(testNameInList $extension ${macroCapableExtensions})) {
        return @{ Mode = $mode; Notice = ""; Strict = $false }
    }
    $notice = ""
    if ($mode -ne ${openModeReadOnly}) {
        $notice = "もらったインデックスのため、マクロを持てる形式は読み取り専用で開きます"
    }
    return @{ Mode = ${openModeReadOnly}; Notice = $notice; Strict = $true }
}

function getOfficeOpenKind {
    # ファイルの開き方の種類（Excel / Office）。Office の形式でなければ空文字。種類は getOfficeFileKind（取り込みの対象の拡張子の表）で決め、
    # Word・PowerPoint はまとめて Office とする。
    #   Excel  … COM で開く（新規は Workbooks.Add）
    #   Office … Word・PowerPoint。エクスプローラーの動詞（読み取り専用は OpenAsReadOnly、新規は New）で開く
    # 読み取り専用で開けるか（testReadOnlySupported）・新規で開くやり方（getOpenNewMethod）は、ここで決める
    param (
        [string]$book
    )

    # Windows は名前の終わりの「.」と空白を落として開くため、同じように落としてから見る
    $kind = getOfficeFileKind (([string]$book).TrimEnd('.', ' '))
    if ($kind -eq "Excel") { return "Excel" }
    if ($kind) { return "Office" }
    return ""
}

function testReadOnlySupported {
    # 読み取り専用で開けるファイルか。Office の形式だけ。
    # テキスト・PDF などは既定のアプリに読み取り専用の開き方が無く、通常と同じ開き方になるため、メニューに出さない
    param (
        [string]$book
    )

    return ((getOfficeOpenKind $book) -ne "")
}

function getOpenNewMethod {
    # ［新規で開く］のやり方を決める。
    #   Excel      … 元のファイルを基にした新しいブック（Workbooks.Add。今までどおり）
    #   Verb       … エクスプローラーの New の動詞（Word・PowerPoint）
    #   Copy       … ほかの種類（テキスト・PDF・そのほか）は、元のファイルのコピーを開く（元のファイルを占有しない）
    param (
        [string]$book
    )

    switch (getOfficeOpenKind $book) {
        "Excel" { return "Excel" }
        "Office" { return "Verb" }
    }
    return "Copy"
}

function getOpenCopyFailedStatus {
    # 開くためのコピーを作れなかったときのステータス。blocker は copySourceForOpen の Reason（空なら、理由を添えない）
    param (
        [string]$path,
        [string]$blocker = ""
    )

    $reason = switch ($blocker) {
        "TooLong" { "ファイル名かフォルダが長すぎるためです。" }
        "Place"   { "コピーの置き場所を確かめられなかったためです。" }
        "Failed"  { "コピーを書き込めなかったためです。" }
        default   { "" }
    }
    return "新規で開くためのコピーを作れませんでした。${reason}元のファイルを開くには［開く］を使ってください：${path}"
}

function getOpenCopyOpenedStatus {
    # コピーを開いたときのステータス。コピーだと分かるよう、コピーのパスと、元のファイルは変わらないことを出す
    #   app: 開いたアプリ（"メモ帳" か空）
    param (
        [string]$path,
        [string]$copyPath,
        [string]$app = ""
    )

    $by = if ($app) { "${app}で" } else { "" }
    return "コピーを${by}新規で開きました（コピー：${copyPath}。保存してもこのコピーに書かれ、元のファイルは変わりません）：${path}"
}

function getSourceReadOnlyFailedStatus {
    # 読み取り専用で開けなかったとき（通常で開くとマクロが動くことがあるため、開かない）のステータス
    param (
        [string]$path
    )

    return "もらったインデックスのファイルは読み取り専用で開けなかったため、開きませんでした。［ファイルのパスをコピー］で場所を調べ、確かめてから開いてください：${path}"
}
function getResultSaveFailedText {
    # 検索結果を search_results.txt に書けなかったときの文（ステータスと検索画面の上の帯に同じ文を出す）
    #   kind: Locked（ほかのアプリが開いている）/ Denied（読み取り専用・書き込み権限が無い。アプリを閉じても直らない）
    param (
        [string]$kind,
        [string]$resultFile
    )

    if ($kind -eq "Denied") {
        return "search_results.txt に書き込む権限がありません（読み取り専用など）。${resultFile} を確かめてから、もう一度出力してください。"
    }
    return "search_results.txt に書き込めません。開いているアプリを閉じてから、もう一度出力してください。"
}
