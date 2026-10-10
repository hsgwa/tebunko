# ワークスペースの中を調べる処理のうち、起動と切り替えで使うもの。
# ワークスペースが届かないネットワークの場所にあっても画面が止まらないよう、ネットワークの場所なら裏の列（network）で調べる。

$script:legacyIndexMessage = ""               # 前の版のインデックスが見つかったときの知らせ（無ければ空）
$script:legacyIndexRequest = @{ Value = 0 }   # 調べの依頼番号（新しい依頼が出たら、前の依頼の結果は捨てる）

function testStartupIndexExists {
    # 起動時の画面を選ぶために、インデックスがあるかを調べる。
    # ネットワークの場所は、画面のスレッドで調べない（届かないと止まる）。ある前提で［検索］を開き、
    # 中身の有無は［検索］の画面が裏で調べる
    param (
        [string]$dir
    )

    if (testNetworkPath $dir) {
        return $true
    }
    return (testIndexExists)
}

function applyLegacyIndexMessage {
    # 前の版のインデックスの知らせを覚えて、あればステータスに出す（起動時に別の知らせを出しているときは出さない）
    param (
        [string]$message
    )

    $script:legacyIndexMessage = $message
    if ($message -and !$script:workspaceBlock) {
        setStatus $message
    }
}

function refreshLegacyIndexMessage {
    # 今のワークスペースに前の版のインデックスがあるかを調べ、知らせを $script:legacyIndexMessage に入れる。
    # ローカルの場所はこの場で調べる。ネットワークの場所は裏の列で調べ、分かったらステータスに出す
    $dir = [string]$workspace.Dir
    $script:legacyIndexMessage = ""
    $requestBox = $script:legacyIndexRequest
    $requestBox.Value++
    $requestId = $requestBox.Value

    if (!(testNetworkPath $dir)) {
        $script:legacyIndexMessage = getLegacyIndexMessage $dir (getLegacyIndexState $dir).HasLegacyIndex
        return
    }

    $apply = ${function:applyLegacyIndexMessage}   # 終わったときの処理は、関数を変数に取って呼ぶ（クロージャからは関数の名前を引けないため）
    startJob {
        param ($dir)
        getLegacyIndexMessage $dir (getLegacyIndexState $dir).HasLegacyIndex
    } @($dir) {
        param ($output, $errorText)
        if ($requestId -ne $requestBox.Value -or $errorText -or !$output -or $output.Count -eq 0) {
            # 待っている間にワークスペースを切り替えた・調べられなかった。知らせは出さない
            return
        }
        & $apply ([string]$output[0])
    }.GetNewClosure() (getWorkspaceJobQueue $dir)
}
