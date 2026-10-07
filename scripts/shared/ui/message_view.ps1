# メッセージの画面（showMessage・確認・エラーの知らせ）の、種類・ボタン・文の分け方（判断層。画面に触らない）。
# 画面に出す部分は shell.ps1。ここは「どの種類にどのアイコンと色を使うか」「どのボタンをどう並べるか」を決める。

# 種類ごとの印。Kind は共通の呼び名、Label は読み上げ・画面の確認に使う日本語の名前、Icon は theme.xaml の Geometry の名前、Brush は色の名前（Icon が空のときは印を出さない）。
# showMessage の icon 引数（MessageBoxImage の名前）と、確認・エラーの形から決める
function getMessageLook {
    param (
        [string]$icon
    )

    switch -Regex ($icon) {
        "^(Error|Stop|Hand)$"           { return @{ Kind = "Error";    Label = "エラー";   Icon = "Icon.CircleAlert";         Brush = "Danger.Text" } }
        "^(Warning|Exclamation)$"       { return @{ Kind = "Warning";  Label = "警告";     Icon = "Icon.TriangleAlert";       Brush = "Warn" } }
        "^Question$"                    { return @{ Kind = "Question"; Label = "確認";     Icon = "Icon.CircleQuestionMark";  Brush = "Accent" } }
        "^(Information|Asterisk)$"      { return @{ Kind = "Info";     Label = "お知らせ"; Icon = "Icon.Info";                Brush = "Accent" } }
        default                         { return @{ Kind = "None";     Label = "";         Icon = "";                         Brush = "" } }
    }
}

# ボタンの並び。左から右へ。Result は showMessage の戻り値（MessageBoxResult の名前）。
# 既定のボタン（Enter）は default で決める（None なら最初のボタン）。Cancel は Esc で押されるボタン。Primary は主なボタン（塗りつぶし）
function getMessageButtons {
    param (
        [string]$buttons = "OK",
        [string]$default = "None"
    )

    $specs = @(switch ($buttons) {
        "OKCancel"    { @(@{ Text = "OK"; Result = "OK" }, @{ Text = "キャンセル"; Result = "Cancel" }) }
        "YesNo"       { @(@{ Text = "はい"; Result = "Yes" }, @{ Text = "いいえ"; Result = "No" }) }
        "YesNoCancel" { @(@{ Text = "はい"; Result = "Yes" }, @{ Text = "いいえ"; Result = "No" }, @{ Text = "キャンセル"; Result = "Cancel" }) }
        default       { @(@{ Text = "OK"; Result = "OK" }) }
    })
    $defaultResult = if (@($specs | Where-Object { $_.Result -eq $default }).Count -gt 0) { $default } else { $specs[0].Result }
    $result = @()
    foreach ($spec in $specs) {
        $isDefault = ($spec.Result -eq $defaultResult)
        $result += , @{
            Text = $spec.Text
            Result = $spec.Result
            IsDefault = $isDefault
            Primary = $isDefault
            # OK だけなら Esc でも閉じる。取りやめがあるものは、その取りやめで閉じる。はい・いいえだけのものは Esc では閉じない（OS 標準と同じ）
            IsCancel = ($spec.Result -eq "Cancel") -or ($buttons -eq "OK")
        }
    }
    return , $result
}

# ×や Alt+F4 で閉じたときの戻り値（OS 標準のメッセージボックスと同じ）
function getMessageCloseResult {
    param (
        [string]$buttons = "OK"
    )

    switch ($buttons) {
        "OK"       { return "OK" }
        "YesNo"    { return "No" }
        default    { return "Cancel" }
    }
}

# 文を、見出し（最初の段落）と補足・詳しい内容に分ける。段落の区切りは空行。文言そのものは変えない。
# 補足が長い（4 行を超える・240 字を超える）ときは、選べる・スクロールできる詳細の枠に回す
function getMessageParts {
    param (
        [string]$message
    )

    $text = ([string]$message) -replace "`r`n", "`n"
    $index = $text.IndexOf("`n`n")
    if ($index -lt 0) {
        return @{ Heading = $text; Hint = ""; Detail = "" }
    }
    $heading = $text.Substring(0, $index)
    $rest = $text.Substring($index + 2).Trim("`n")
    if ($rest.Length -gt 240 -or @($rest -split "`n").Count -gt 4) {
        return @{ Heading = $heading; Hint = ""; Detail = $rest }
    }
    return @{ Heading = $heading; Hint = $rest; Detail = "" }
}
