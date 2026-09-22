# 起動時の確認の判断（いつもの画面で起動するか、制限モードで起動するか）と、利用者に見せる文言。
#
# 判断層。制限言語モード（CLM）でも読み込めて動く書き方だけで書く（start.ps1 が、画面を開く前に読み込むため）。
# 入力の $facts は getStartupFacts（startup_check.ps1）が集めるハッシュテーブル:
#   LanguageMode … $ExecutionContext.SessionState.LanguageMode の文字列（FullLanguage / ConstrainedLanguage など）
#   WpfError     … 画面の部品（WPF）を読み込めなかったときのメッセージ。読み込めたとき・調べなかったときは空
#   HasExcel     … Excel が入っているか
#   WorkDir      … work フォルダの場所
#   CanWriteWork … work フォルダに書き込めるか

${startupModeNormal}     = "normal"      # いつもの画面（gui.ps1）で起動する
${startupModeRestricted} = "restricted"  # 制限モード（コンソール）で起動する

# 起動のしかたを決める。画面は WPF と PowerShell class を使うため、FullLanguage で WPF を読み込めるときだけ開ける
function getStartupMode {
    param ($facts)

    if ($facts.LanguageMode -eq "FullLanguage" -and !$facts.WpfError) {
        return ${startupModeNormal}
    }
    return ${startupModeRestricted}
}

# 制限モードで起動するときに、コンソールへ出す説明（理由・使えない機能・いつもの画面を使うための頼み方）
function getRestrictedStartupLines {
    param ($facts)

    $clm = $facts.LanguageMode -ne "FullLanguage"
    $lines = @("制限モードで起動します。")
    if ($clm) {
        $lines += "  理由: この PC では PowerShell が制限言語モード（$($facts.LanguageMode)）で動いています。"
    } else {
        $lines += "  理由: 画面の部品（WPF）を読み込めませんでした（$($facts.WpfError)）。"
    }
    $lines += "  使えない機能: 画面（ウィンドウ）と、旧形式（.xls .doc .ppt）・パスワード付きのファイルの取り込み"
    $lines += "  インデックスの作成: 新形式（.docx .docm .pptx .pptm .xlsx .xlsm）だけ取り込めます。ほかはいつもの画面が使える PC で取り込んでください。"
    if (!$facts.HasExcel) {
        $lines += "  検索結果: Excel が見つからないため、この画面に一覧で出します。"
    }
    if (!$facts.CanWriteWork) {
        $lines += "  注意: work フォルダ（$($facts.WorkDir)）に書き込めないため、インデックスの作成と設定の保存ができません。"
    }
    if ($clm) {
        $lines += "  いつもの画面を使うには: 管理者にツールのフォルダの実行許可を依頼してください。"
    } else {
        $lines += "  いつもの画面を使うには: この PC で .NET Framework の画面の部品（WPF）が動くかを、管理者に確かめてもらってください。"
    }
    return ,$lines
}
