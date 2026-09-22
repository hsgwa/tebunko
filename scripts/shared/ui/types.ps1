# 画面で使う型のうち、どのツールからも使うもの。
# 継承元になるため、ほかの型より先に読み込む。

# ---- 画面で使う型（PowerShell class。実行時コンパイル（csc.exe）を出さないため C# から移した） ----
# INotifyPropertyChanged は NotifyBase を継承して実装する（PS class はプロパティのセッターに
# ロジックを書けないため、値を変える側が Set*/Raise を呼ぶ）。HitRow は件数が多いので生成時は生データだけ持ち、
# 表示用（強調セグメント・セル解析）は画面に見えた行だけ Prepare() で作る（LoadingRow から呼ぶ）。

class NotifyBase : System.ComponentModel.INotifyPropertyChanged {
    hidden [System.ComponentModel.PropertyChangedEventHandler] $handler
    [void] add_PropertyChanged([System.ComponentModel.PropertyChangedEventHandler]$h) { $this.handler = [Delegate]::Combine($this.handler, $h) }
    [void] remove_PropertyChanged([System.ComponentModel.PropertyChangedEventHandler]$h) { $this.handler = [Delegate]::Remove($this.handler, $h) }
    [void] Raise([string]$name) { if ($this.handler) { $this.handler.Invoke($this, (New-Object System.ComponentModel.PropertyChangedEventArgs $name)) } }
}

# 確認ダイアログ（showConfirm）に並べる「実行するとこうなります」の 1 行
class ConfirmFact {
    [string]$Mark          # ✓ 残る・✗ 消える・→ 続けて起きること
    [object]$MarkBrush
    [string]$Title         # 何がどうなるか（利用者から見える言葉で書く）
    [string]$Detail = ""   # 具体的な場所・件数など（空なら行ごと出さない）
}

# 左のツリーの1項目。データだけを持ち、中身の読み込みは loadFolderNode（shared\core\folder.ps1 の getFolderEntries を呼ぶ）で行う。
# 展開（IsExpanded）・選択（IsSelected）は TreeViewItem と TwoWay バインドし、
# 展開したときの読み込みは TreeView の Expanded イベントで駆動する（PS class はセッターにロジックを書けないため）
class FolderNode : NotifyBase {
    [string]$Name
    [string]$Path                 # 開くフォルダ。見出し（「よく使う場所」「PC」）は ""
    [string]$ToolTip
    [bool]$IsHeader               # 見出し（フォルダではないので、選んでも移動しない）
    [bool]$IsPlaceholder          # 「読み込み中…」（▷ を出すためだけの子）
    [bool]$Loaded                 # 子を読み込み済みか
    [FolderNode]$Parent
    [System.Collections.ObjectModel.ObservableCollection[FolderNode]]$Children

    [bool]$IsExpanded
    [bool]$IsSelected

    FolderNode([FolderNode]$parent, [string]$name, [string]$path) {
        $this.Parent = $parent
        $this.Name = $name
        $this.Path = $path
        $this.ToolTip = $path
        $this.Children = New-Object System.Collections.ObjectModel.ObservableCollection[FolderNode]
    }

    [void] SetExpanded([bool]$value) {
        if ($this.IsExpanded -eq $value) { return }
        $this.IsExpanded = $value
        $this.Raise("IsExpanded")
    }

    [void] SetSelected([bool]$value) {
        if ($this.IsSelected -eq $value) { return }
        $this.IsSelected = $value
        $this.Raise("IsSelected")
    }

    [void] AddPlaceholder() {
        $node = [FolderNode]::new($this, "読み込み中…", "")
        $node.IsPlaceholder = $true
        $this.Children.Add($node)
    }

    # スクリーンリーダー・自動化ツールにはフォルダ名で見えるようにする（既定では型名になる）
    [string] ToString() { return $this.Name }
}

# 右の一覧の1行。フォルダは選べ、ファイルは「目的のフォルダかどうか」を確かめるために出すだけ（選べない）
class FolderEntry {
    [string]$Name
    [string]$Path
    [bool]$IsFolder
    [bool]$IsOffice
    [string]$Kind
    [string]$UpdatedText

    # スクリーンリーダー・自動化ツールには名前で見えるようにする（既定では型名になる）
    [string] ToString() { return $this.Name }
}

# ［9 プロセス停止］の1行
class ProcRow {
    [int]$Id
    [string]$AppName
    [bool]$Background
    [string]$StartText
    [string]$MemoryText
    [string]$TitleText
}
