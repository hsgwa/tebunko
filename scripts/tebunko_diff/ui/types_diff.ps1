# 比較の画面で使う型。NotifyBase（shared\ui\types.ps1）を継承するため、その後に読み込む。
# 左右に並べる行（DiffRow）・場所（PlaceDiff）・ツリーの行（TreeRow）は判断層（diff\*.ps1）にある。
# それらは裏の runspace で作られるため、画面は型の名前で見分けず（-is を使わず）、プロパティだけを読む

# 左右の差分の一覧の状態（一覧の Tag に入れ、Excel のセルの横の位置を左右そろって動かす）
class DiffViewState : NotifyBase {
    [double]$CellOffset
    [void] SetOffset([double]$value) {
        $this.CellOffset = $value
        $this.Raise("CellOffset")
    }
}

# 場所の見出し（シート・本文・スライドとノート など）
class PlaceTab {
    [string]$Text
    [string]$Status
    [int]$Index
}

# Excel の列の見出し
class ColumnHead {
    [string]$Text
    [double]$Width
    [string]$Kind = ""          # "" / insert（追加した列）/ delete（削除した列）/ empty（相手側にだけある列の空き）
}
