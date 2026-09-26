# 画面で使う型のうち、どのツールからも使うもの。
# 継承元になるため、ほかの型より先に読み込む。

# ---- 画面で使う型（PowerShell class。実行時コンパイル（csc.exe）を出さないため C# から移した） ----
# INotifyPropertyChanged は NotifyBase を継承して実装する（PS class はプロパティのセッターに
# ロジックを書けないため、値を変える側が Set*/Raise を呼ぶ）。HitRow は件数が多いので生成時は生データだけ持ち、
# 表示用（強調セグメント・セル解析）は画面に見えた行だけ Prepare() で作る（LoadingRow から呼ぶ）。

class NotifyBase : System.ComponentModel.INotifyPropertyChanged {
    hidden [System.ComponentModel.PropertyChangedEventHandler] $handler
    # 既定のコンストラクタを明示する。Windows PowerShell 5.1 では、メソッドを持ち、コンストラクタも初期値付きのフィールドも無いクラスを
    # Set-PSDebug -Trace の最中に作ると ArgumentOutOfRangeException になる（テストのカバレッジの計測がトレースを使う）。
    # 継承するクラスも同じ形なら、それぞれ明示する（tests\meta\classes.Tests.ps1 が確かめる）
    NotifyBase() {}
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
