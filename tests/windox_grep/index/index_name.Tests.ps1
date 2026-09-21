# インデックス名と TSV のファイル名（windox_grep\index\index_name.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"

Describe "encodeIndexPlace / decodeIndexPlace" -Tag Unit {
    It "ファイル名に使えない文字・_・% を %XX にし、元に戻せる" {
        $place = 'a<b>c\d*e:f?g|h/i"j_k%l'
        $encoded = encodeIndexPlace $place
        $encoded | Should Be 'a%3Cb%3Ec%5Cd%2Ae%3Af%3Fg%7Ch%2Fi%22j%5Fk%25l'
        decodeIndexPlace $encoded | Should BeExactly $place
    }

    It "制御文字も %XX にする" {
        encodeIndexPlace "a`tb" | Should Be 'a%09b'
        decodeIndexPlace 'a%09b' | Should Be "a`tb"
    }

    It "半角の記号と全角の記号は別の名前になる（衝突しない）" {
        encodeIndexPlace '衝突"' | Should Not Be (encodeIndexPlace '衝突”')
    }

    It "使える文字（全角記号・空白・&'#() 等）はそのまま返す" {
        # PowerShell は ” を " と同じに扱うため、' で囲む
        encodeIndexPlace 'シート1 (2)&''#＜” 全角　' | Should Be 'シート1 (2)&''#＜” 全角　'
    }

    It "符号化で作らない %XX はそのまま返す（以前の版のシート名 100% 等）" {
        decodeIndexPlace '100%' | Should Be '100%'
        decodeIndexPlace '%41%2a' | Should Be '%41%2a'
    }
}

Describe "toIndexFileName" -Tag Unit {
    It "<場所>.tsv にし、場所は符号化する（元のファイル名はフォルダ名にするため入れない）" {
        toIndexFileName "記号<>_1" | Should Be "記号%3C%3E%5F1.tsv"
    }

    It "場所が長くても255文字を超えなければ返す" {
        (toIndexFileName ("あ" * 251)).Length | Should Be 255
    }

    It "255文字を超えると、文字数の分かるメッセージで例外にする" {
        { toIndexFileName ("あ" * 252) } | Should Throw "256 文字。上限 255 文字"
    }
}

Describe "splitIndexTsvPath" -Tag Unit {
    It "今の形式（<ファイル名.xlsx>\<場所>.tsv）をファイル名・場所・フォルダに分ける" {
        foreach ($case in @(
            @("A.xlsx", "old.xlsx_1"),
            @("A.xlsx_old.xlsx", "1"),
            @("コピー.xls_old.xlsx", "Sheet1"),
            @("終わりが_.xlsx", "_"),
            @("100%.xlsx", "100%"),
            @("ア_イ_ウ.xlsx", "シ_ー_ト"),
            @("[確定]報告書.xlsx", "衝突`""),
            @("資料.pptx", "スライド003_ノート"),
            @("報告書.docx", "ヘッダー・フッター")
        )) {
            $parts = splitIndexTsvPath "営業\2024\$($case[0])\$(toIndexFileName $case[1])"
            $parts.Book | Should BeExactly $case[0]
            $parts.Place | Should BeExactly $case[1]
            $parts.RelDir | Should Be "営業\2024"
        }
    }

    It "インデックスフォルダの直下のファイルは、フォルダが空になる" {
        $parts = splitIndexTsvPath "A.xlsx\Sheet1.tsv"
        $parts.Book | Should Be "A.xlsx"
        $parts.Place | Should Be "Sheet1"
        $parts.RelDir | Should Be ""
    }

    It "以前の形式（<ファイル名.xlsx>_<場所>.tsv）も分けられる" {
        $parts = splitIndexTsvPath "営業\コピー.xls_old.xlsx_Sheet1.tsv"
        $parts.Book | Should Be "コピー.xls_old.xlsx"
        $parts.Place | Should Be "Sheet1"
        $parts.RelDir | Should Be "営業"

        # さらに以前の形式（場所に _ をそのまま入れていた版）
        $parts = splitIndexTsvPath "売上.xls_2024_上期.tsv"
        $parts.Book | Should Be "売上.xls"
        $parts.Place | Should Be "2024_上期"
        $parts.RelDir | Should Be ""
    }

    It "Officeファイル以外の名前のフォルダにあるTSVは、ファイル名をそのまま返す" {
        $parts = splitIndexTsvPath "メモ\other.tsv"
        $parts.Book | Should Be "other.tsv"
        $parts.Place | Should Be ""
        $parts.RelDir | Should Be "メモ"
    }
}

Describe "splitIndexFileName" -Tag Unit {
    It "ブック名とシート名に分解する" {
        $name = splitIndexFileName "book.xlsx_Sheet1.tsv"
        $name.book | Should Be "book.xlsx"
        $name.sheet | Should Be "Sheet1"
    }

    It "場所の %XX を元に戻す" {
        $name = splitIndexFileName "売上.xls_2024%5F上期%22.tsv"
        $name.book | Should Be "売上.xls"
        $name.sheet | Should Be '2024_上期"'
    }

    It "ファイル名に .xls_ 等を含んでも、最後の _ で分解する" {
        $name = splitIndexFileName "コピー.xls_old.xlsx_Sheet1.tsv"
        $name.book | Should Be "コピー.xls_old.xlsx"
        $name.sheet | Should Be "Sheet1"
    }

    It "以前の版のTSV（シート名の _ を符号化していない）も分解できる" {
        $name = splitIndexFileName "売上.xls_2024_上期.tsv"
        $name.book | Should Be "売上.xls"
        $name.sheet | Should Be "2024_上期"

        $name = splitIndexFileName "ア_イ_ウ.xlsx_シ_ト＜＞.tsv"
        $name.book | Should Be "ア_イ_ウ.xlsx"
        $name.sheet | Should Be "シ_ト＜＞"
    }

    It "拡張子が大文字でも分解できる" {
        $name = splitIndexFileName "大文字.XLSX_Sheet1.tsv"
        $name.book | Should Be "大文字.XLSX"
        $name.sheet | Should Be "Sheet1"
    }

    It "xlsm も扱える" {
        (splitIndexFileName "macro.xlsm_A.tsv").book | Should Be "macro.xlsm"
    }

    It "Word・PowerPointのファイル名と場所に分解する" {
        $name = splitIndexFileName "報告書.docx_ページ001.tsv"
        $name.book | Should Be "報告書.docx"
        $name.sheet | Should Be "ページ001"

        $name = splitIndexFileName "旧.doc_ヘッダー・フッター.tsv"
        $name.book | Should Be "旧.doc"
        $name.sheet | Should Be "ヘッダー・フッター"

        $name = splitIndexFileName "提案.pptx_スライド003%5Fノート.tsv"
        $name.book | Should Be "提案.pptx"
        $name.sheet | Should Be "スライド003_ノート"
    }

    It "形式外のファイル名はそのままブック名として返す" {
        $name = splitIndexFileName "other.tsv"
        $name.book | Should Be "other.tsv"
        $name.sheet | Should Be ""
    }
}

Describe "newIndexName / assignIndexNames" -Tag Unit {
    It "フォルダ名（ドライブ直下はドライブ名、UNC は共有名）を使い、重複すれば (2) を付ける" {
        $used = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        newIndexName "C:\data\見積" $used | Should Be "見積"
        newIndexName "D:\" $used | Should Be "D"
        newIndexName "\\server\share" $used | Should Be "share"
        [void]$used.Add("見積")
        [void]$used.Add("見積(2)")
        newIndexName "E:\見積" $used | Should Be "見積(3)"
    }

    It '使用中の名前が $null・配列・空でも落ちずに名前を作る' {
        # インデックスが 1 つも無いとき、呼び出し側から $null が渡ることがある（画面の新規作成ダイアログ）
        newIndexName "C:\data\sample" $null | Should Be "sample"
        newIndexName "C:\data\見積" | Should Be "見積"
        newIndexName "C:\data\見積" @() | Should Be "見積"
        newIndexName "C:\data\見積" @("見積") | Should Be "見積(2)"
        # インデックスが 1 件のときは、集合が展開されて文字列 1 個で渡ることがある
        newIndexName "C:\data\見積" "見積" | Should Be "見積(2)"
        newIndexName "C:\data\見" "見積" | Should Be "見"
        # 前方一致・大文字小文字違いで取り違えない
        newIndexName "C:\data\見積" @("見積書") | Should Be "見積"
        newIndexName "C:\data\sample" @("SAMPLE") | Should Be "sample(2)"
    }

    It "設定に名前があればそれを使い、フォルダの場所が変わっても同じ名前のままにする" {
        $targets = @(
            [pscustomobject]@{ Name = "見積"; Path = "\server\新しい場所\見積書"; Enabled = $true },
            [pscustomobject]@{ Name = ""; Path = "F:\売上"; Enabled = $true }
        )
        # 前回は別の場所だったが、名前が同じなので同じインデックスとして扱う
        $previous = @([pscustomobject]@{ Path = "C:\data\見積"; Name = "見積" })
        $folders = @(assignIndexNames $targets $previous)
        $folders.Count | Should Be 2
        $folders[0].Name | Should Be "見積"
        $folders[0].Path | Should Be "\server\新しい場所\見積書"
        $folders[1].Name | Should Be "売上"
    }

    It "名前が無ければ、前回の変換一覧の同じフォルダの名前を使い、無ければフォルダ名から作る" {
        $targets = @(
            [pscustomobject]@{ Name = ""; Path = "C:\data\見積"; Enabled = $false },
            [pscustomobject]@{ Name = ""; Path = "E:\new\見積"; Enabled = $true },
            [pscustomobject]@{ Name = ""; Path = "F:\売上"; Enabled = $true }
        )
        $previous = @(
            [pscustomobject]@{ Path = "c:\data\見積"; Name = "見積" },
            [pscustomobject]@{ Path = "G:\削除した\報告"; Name = "報告" }
        )
        $folders = @(assignIndexNames $targets $previous)
        $folders[0].Name | Should Be "見積"
        $folders[0].Enabled | Should Be $false
        # 前回の名前（削除したフォルダの名前を含む）と重複しない名前を付ける
        $folders[1].Name | Should Be "見積(2)"
        $folders[2].Name | Should Be "売上"
    }

    It "設定にある名前は、ほかのフォルダの名前には使わない" {
        $targets = @(
            [pscustomobject]@{ Name = ""; Path = "E:\新\見積"; Enabled = $true },
            [pscustomobject]@{ Name = "見積"; Path = "C:\data\見積"; Enabled = $true }
        )
        $folders = @(assignIndexNames $targets @())
        $folders[0].Name | Should Be "見積(2)"
        $folders[1].Name | Should Be "見積"
    }
}

Describe "splitIndexRelPath" -Tag Unit {
    It "先頭のインデックス名と残りに分ける（残りの \ や 2 はそのまま）" {
        $parts = splitIndexRelPath "excel\2024\見積\A社2.xlsx"
        $parts.Name | Should Be "excel"
        $parts.Rest | Should Be "2024\見積\A社2.xlsx"
        $parts = splitIndexRelPath "excel"
        $parts.Name | Should Be "excel"
        $parts.Rest | Should Be ""
    }
}

Describe "testIndexName" -Tag Unit {
    It "使える名前なら空文字列を返す" {
        testIndexName "営業部 2025" | Should Be ""
    }

    It "空なら入力を促す" {
        testIndexName "" | Should Match "入力してください"
    }

    It "前後に空白があれば使えない" {
        testIndexName " 営業" | Should Match "前後に空白"
        testIndexName "営業 " | Should Match "前後に空白"
    }

    It "ファイル名に使えない文字があれば使えない" {
        testIndexName "営業\部" | Should Match "使えない文字"
        testIndexName "営業:部" | Should Match "使えない文字"
    }

    It "末尾が . なら使えない" {
        testIndexName "営業." | Should Match "最後に \."
    }

    It "Windows の予約語は使えない（大文字・小文字を区別しない）" {
        testIndexName "con" | Should Match "使えない名前"
        testIndexName "LPT1" | Should Match "使えない名前"
    }

    It "ほかのインデックスと同じ名前は使えない（大文字・小文字を区別しない）" {
        testIndexName "Sales" @("sales", "tech") | Should Match "ほかのインデックスが使っています"
    }

    It "長すぎる名前は使えない" {
        testIndexName ("あ" * 256) | Should Match "長すぎます"
    }
}
