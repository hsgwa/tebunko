# コミットメッセージの 1 行目・PR や Issue のタイトルが Conventional Commits の形かどうかの判定（AGENTS.md「GitHub の運用」）。
# 副作用（exit・Write-Host など）を持たない部品だけを置く。tools/check_commit_message.ps1 と tools/check_compat_golden.ps1 の
# 両方から dot-source して使う（型(範囲)!: 説明 を読み取る正規表現を 2 か所に書かない）。

# 使える型と使い分け。PR に付けるラベル（.github/release.yml がリリースノートの分類に使う）の目安も並べる
$types = [ordered]@{
    feat     = "機能の追加・変更（ラベル enhancement）"
    fix      = "不具合の修正（ラベル bug）"
    docs     = "文書だけの変更（ラベル documentation）"
    refactor = "動きを変えない書き直し（ラベル internal）"
    perf     = "速さの改善（ラベル enhancement）"
    test     = "テストだけの追加・修正（ラベル internal）"
    style    = "書式だけの変更（空白・改行など）（ラベル internal）"
    build    = "配布物の作り方・依存の更新（依存の更新はラベル dependencies、ほかは internal）"
    ci       = "CI・git のフック・開発用の道具（ラベル internal）"
    chore    = "上のどれにも当たらないもの（ラベル internal）"
    revert   = "前の変更の取り消し（取り消した PR と同じラベル）"
}

# git が自動で作るメッセージ
$generated = '^(Merge |Revert "|fixup! |squash! |amend! )'

# 「<型>(<範囲>)!: <説明>」の形。! は前の版と互換が無くなる変更（breaking）の印
$titlePattern = '^(?<type>[^(!:\s]+)(\((?<scope>[^()]*)\))?(?<breaking>!)?: (?<subject>.*)$'

# 1 行目を調べ、問題があれば理由を返す（問題が無ければ $null）。git が自動で作るメッセージは調べない
function findProblem([string]$line) {
    if ($line -match $generated) { return $null }
    if ($line -cnotmatch $titlePattern) {
        return "「<型>: <説明>」の形になっていません（型の後ろに半角のコロンと空白が 1 つ要ります）"
    }
    # 後の -match で $Matches が上書きされるため、先に取り出す
    $type = $Matches.type
    $scope = $Matches.scope
    $subject = $Matches.subject
    # [ordered] のキーは大文字・小文字を区別しないため、小文字であることは別に確かめる
    if (!$types.Contains($type) -or $type -cne $type.ToLowerInvariant()) {
        return "型「$type」は使えません"
    }
    if ($null -ne $scope -and $scope -notmatch '^\S+$') {
        return "範囲「($scope)」は空白を含まない 1 語にしてください"
    }
    if ($subject -notmatch '^\S') {
        return "説明がありません（コロンの後ろの空白は 1 つにしてください）"
    }
    return $null
}

# タイトルの型に ! が付いているか（前の版と互換が無くなる変更かどうか）。
# 形が違う・git が自動で作るメッセージのときは $false（check_commit_message.ps1 が別に形を止めるため、ここでは問わない）
function testBreakingTitle([string]$line) {
    if ($line -match $generated) { return $false }
    if ($line -cmatch $titlePattern) {
        return [bool]$Matches.breaking
    }
    return $false
}
