# 元のファイルを開くとき（ネットワークにあるときの確認）の文言（判断層）。
# 画面に触らないため、そのままテストできる（tests\tebunko\ui\search\open_source_view.Tests.ps1）。

function getSourceCheckingStatus {
    # 元のファイルの場所を確かめている間のステータス（ネットワークにあるときだけ出す。ローカルはその場で開くため出さない）
    param (
        [string]$path
    )

    return "元のファイルを確かめています…：${path}（共有フォルダに接続できないときは、しばらくかかります）"
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
