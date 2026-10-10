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

function getSourceNotFoundStatus {
    # 見つからない・確認で選ばなかったときのステータス
    param (
        [string]$path
    )

    return "元のファイルが見つかりません：${path}"
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
    #   isNetwork: 元のフォルダがネットワークの場所か（真なら、接続とサインイン情報について注意を足す）
    # 返すもの: @{ Heading; Title; Detail; Hint; UseText; PickText }
    param (
        [string]$book,
        [string]$name,
        [string]$folder,
        [bool]$isNetwork
    )

    $hint = "［このフォルダを使う］を選ぶと、次からはこのインデックスについて聞きません"
    if ($isNetwork) {
        $hint = "開くと、このフォルダに接続し、Windows のサインイン情報が送られることがあります。心当たりのない場所なら、［フォルダを選ぶ］で別のフォルダを選んでください。" + $hint
    }
    return @{
        Heading  = "${book} の元のフォルダを確かめてください"
        Title    = "インデックス [${name}] に書かれた元のフォルダ"
        Detail   = $folder
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

function getSourceReadOnlyFailedStatus {
    # 読み取り専用で開けなかったとき（通常で開くとマクロが動くことがあるため、開かない）のステータス
    param (
        [string]$path
    )

    return "もらったインデックスのファイルは読み取り専用で開けなかったため、開きませんでした。［ファイルのパスをコピー］で場所を調べ、確かめてから開いてください：${path}"
}
