# 起動時の確認の判断（tebunko_grep\restricted\startup_view.ps1）のテスト。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\tebunko_grep\restricted\startup_view.ps1"

# 起動時の確認の事実（getStartupFacts が返す形）。既定はいつもの画面を開ける PC
function newFacts {
    param ([hashtable]$override = @{})

    $facts = @{
        LanguageMode = "FullLanguage"
        WpfError     = ""
        HasExcel     = $true
        WorkDir      = "C:\tools\tebunko_grep\work"
        CanWriteWork = $true
    }
    foreach ($key in $override.Keys) { $facts[$key] = $override[$key] }
    return $facts
}

Describe "getStartupMode" -Tag Unit {
    It "FullLanguage で WPF を読み込めれば、いつもの画面" {
        getStartupMode (newFacts) | Should Be ${startupModeNormal}
    }

    It "制限言語モードなら制限モード" {
        getStartupMode (newFacts @{ LanguageMode = "ConstrainedLanguage" }) | Should Be ${startupModeRestricted}
    }

    It "FullLanguage でも WPF を読み込めなければ制限モード" {
        getStartupMode (newFacts @{ WpfError = "ファイルが見つかりません" }) | Should Be ${startupModeRestricted}
    }

    It "Excel が無い・work に書き込めないことは、起動のしかたに関係しない" {
        getStartupMode (newFacts @{ HasExcel = $false; CanWriteWork = $false }) | Should Be ${startupModeNormal}
    }
}

Describe "getRestrictedStartupLines" -Tag Unit {
    It "制限言語モードの理由・使えない機能・頼み方を出す" {
        $lines = getRestrictedStartupLines (newFacts @{ LanguageMode = "ConstrainedLanguage" })
        ($lines -join "`n") | Should Be (@(
                "制限モードで起動します。"
                "  理由: この PC では PowerShell が制限言語モード（ConstrainedLanguage）で動いています。"
                "  使えない機能: 画面（ウィンドウ）と、旧形式（.xls .doc .ppt）・パスワード付きのファイルの取り込み"
                "  インデックスの作成: 新形式（.docx .docm .pptx .pptm .xlsx .xlsm）だけ取り込めます。ほかはいつもの画面が使える PC で取り込んでください。"
                "  いつもの画面を使うには: 管理者にツールのフォルダの実行許可を依頼してください。"
            ) -join "`n")
    }

    It "WPF を読み込めなかったときは、その理由と確かめ方を出す" {
        $lines = getRestrictedStartupLines (newFacts @{ WpfError = "ファイルが見つかりません" })
        $lines[1] | Should Be "  理由: 画面の部品（WPF）を読み込めませんでした（ファイルが見つかりません）。"
        $lines[-1] | Should Be "  いつもの画面を使うには: この PC で .NET Framework の画面の部品（WPF）が動くかを、管理者に確かめてもらってください。"
    }

    It "Excel が無ければ、結果をコンソールに出すことを伝える" {
        $lines = getRestrictedStartupLines (newFacts @{ LanguageMode = "ConstrainedLanguage"; HasExcel = $false })
        $lines -contains "  検索結果: Excel が見つからないため、この画面に一覧で出します。" | Should Be $true
    }

    It "work に書き込めなければ、その場所と、できないことを伝える" {
        $lines = getRestrictedStartupLines (newFacts @{ LanguageMode = "ConstrainedLanguage"; CanWriteWork = $false })
        $lines -contains "  注意: work フォルダ（C:\tools\tebunko_grep\work）に書き込めないため、インデックスの作成と設定の保存ができません。" | Should Be $true
    }

    It "1 行だけでも配列で返す" {
        $lines = getRestrictedStartupLines (newFacts @{ LanguageMode = "ConstrainedLanguage" })
        ($lines -is [array]) | Should Be $true
    }
}
