# tebunko の画面で使う型: ［インデックス管理］の一覧・詳細・失敗した行。
# shared\ui\types.ps1 を先に読み込んでおくこと（NotifyBase を継承する）。

# ［インデックス管理］の取り込みに失敗したファイル1件
class FailRow {
    [string]$RelPath
    [string]$Reason
    [string]$IngestedText
    [string]$SourcePath
}

# インデックスの詳細のパネルの値の 1 行（getIndexDetailView）
class DetailRow {
    [string]$Label
    [string]$Value
}

# インデックス更新の確認ダイアログに出すインデックス 1 件（取り込み予定.tsv の 1 行）
class PlanRow {
    [string]$Name
    [string]$Path
    [string]$TotalText    # 対象ファイル数（"1,243"。数えない行は "－"）
    [string]$StatusText   # ステータスのバッジの文言（"要更新" / "最新" / "対象外" / "フォルダなし"）
    [string]$Level        # バッジの色（Wait / Ok / None / Ng）
    [string]$DetailText   # バッジの ToolTip（更新するファイルの内訳）
}

# ［インデックス管理］のインデックス一覧 1 件。プログラムから変えたときに画面へ反映するため通知する。
# 行のチェックの TwoWay バインドは値の往復に使い、保存はチェックボックスの Click で行う（PS class はセッターにロジックを書けないため）
class FolderItem : NotifyBase {
    [string]$Name          # インデックス名（work\index 直下のフォルダ名）
    [string]$Path
    [bool]$Enabled         # 設定の enabled（［すべて更新］で更新する対象か。画面にチェックは出さない）
    [bool]$Checked         # 一覧のチェック（一時の選択。保存しない。［アクション ▼］・行の操作の対象）
    [string]$StatusText
    [object]$StatusBrush
    [string]$FileCountText
    [string]$FileCountToolTip
    [string]$LastIngestedText
    [bool]$StatusChecked   # フォルダの有無を調べ終えたか（別スレッドで調べる。refreshFolderStatus）
    [bool]$FolderExists    # 調べた結果、フォルダがあったか
    # 一覧の「ステータス」列（本文の取り込みの状態。getIndexRowView）。色は XAML で Level から選ぶ（Ok/Wait/Ng/Run/None）
    [string]$IndexText = "－"
    [string]$IndexToolTip
    [string]$IndexLevel = "None"
    [double]$IndexPercent = 0   # 更新中のバッジの下の棒の長さ（0～100）
    [string]$IndexSub = ""   # ステータスのバッジの下の補足（途中で止まったときの「残り N 件」。無ければ空）
    # 一覧の行の右端のボタン（getIndexRowActions）。RowAction は Update / Stop / None
    [string]$RowAction = "Update"
    [bool]$RowUpdateEnabled = $false

    FolderItem() {}   # 既定のコンストラクタを明示する（理由は shared\ui\types.ps1 の NotifyBase）

    [void] SetEnabled([bool]$value) { if ($this.Enabled -ne $value) { $this.Enabled = $value; $this.Raise("Enabled") } }
    [void] SetRowChecked([bool]$value) { if ($this.Checked -ne $value) { $this.Checked = $value; $this.Raise("Checked") } }
    [void] SetName([string]$value) { if ($this.Name -ne $value) { $this.Name = $value; $this.Raise("Name") } }
    [void] SetPath([string]$value) { if ($this.Path -ne $value) { $this.Path = $value; $this.Raise("Path") } }
    [void] SetStatus([string]$text, [object]$brush) { $this.StatusText = $text; $this.StatusBrush = $brush; $this.Raise("StatusText"); $this.Raise("StatusBrush") }
    [void] SetStats([string]$countText, [string]$toolTip, [string]$lastIngested) {
        $this.FileCountText = $countText; $this.FileCountToolTip = $toolTip; $this.LastIngestedText = $lastIngested
        $this.Raise("FileCountText"); $this.Raise("FileCountToolTip"); $this.Raise("LastIngestedText")
    }
    [void] SetIndexState([string]$text, [string]$toolTip, [string]$level, [string]$sub, [double]$percent) {
        $this.IndexText = $text; $this.IndexToolTip = $toolTip; $this.IndexLevel = $level; $this.IndexSub = $sub; $this.IndexPercent = $percent
        $this.Raise("IndexText"); $this.Raise("IndexToolTip"); $this.Raise("IndexLevel"); $this.Raise("IndexSub"); $this.Raise("IndexPercent")
    }
    [void] SetRowActions([string]$action, [bool]$updateEnabled) {
        $this.RowAction = $action; $this.RowUpdateEnabled = $updateEnabled
        $this.Raise("RowAction"); $this.Raise("RowUpdateEnabled")
    }
}
