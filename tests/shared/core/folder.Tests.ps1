# フォルダのパスと一覧（shared\core\folder.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"

Describe "normalizeFolderPath" -Tag Io {
    It "前後の空白・引用符と末尾の \ を取り除き、ドライブ直下は \ を残す" {
        normalizeFolderPath '  "C:\data\"  ' | Should Be "C:\data"
        normalizeFolderPath "D:\" | Should Be "D:\"
        normalizeFolderPath "D:" | Should Be "D:\"
        normalizeFolderPath "\\server\share\" | Should Be "\\server\share"
    }

    It "/ を \ にそろえ、長いパス用の \\?\ ・ \\?\UNC\ を外す" {
        normalizeFolderPath "C:/data/見積" | Should Be "C:\data\見積"
        normalizeFolderPath "//server/share/見積" | Should Be "\\server\share\見積"
        normalizeFolderPath "\\?\C:\data\見積" | Should Be "C:\data\見積"
        normalizeFolderPath "\\?\UNC\server\share\見積" | Should Be "\\server\share\見積"
    }

    It "重なった \ ・ . ・ .. を解決する" {
        normalizeFolderPath "C:\data\\見積" | Should Be "C:\data\見積"
        normalizeFolderPath "C:\data\.\見積" | Should Be "C:\data\見積"
        normalizeFolderPath "C:\data\売上\..\見積" | Should Be "C:\data\見積"
        normalizeFolderPath "\\server\share\売上\..\見積" | Should Be "\\server\share\見積"
    }

    It "環境変数を展開する" {
        $env:tebunko_TEST_FOLDER = "C:\data\見積"
        try {
            normalizeFolderPath "%tebunko_TEST_FOLDER%" | Should Be "C:\data\見積"
            normalizeFolderPath "%tebunko_TEST_FOLDER%\2024" | Should Be "C:\data\見積\2024"
        } finally {
            Remove-Item Env:\tebunko_TEST_FOLDER
        }
    }

    It "相対パスは tebunko のフォルダからとみなす" {
        normalizeFolderPath "work\index" | Should Be "${rootDir}\work\index"
        normalizeFolderPath ".\work\index" | Should Be "${rootDir}\work\index"
    }

    It "パスとして解釈できない場合は書かれたとおりに扱う" {
        normalizeFolderPath "C:\data*" | Should Be "C:\data*"
        normalizeFolderPath "\\server" | Should Be "\\server"
        # \ ひとつで始まるパスは、書き間違えた UNC パスのことが多いため、今のドライブのパスに直さない
        normalizeFolderPath "\server\share" | Should Be "\server\share"
    }
}

Describe "getPathUnderFolder" -Tag Io {
    It "フォルダからの相対パスを返す（フォルダ自身は空。末尾の \ ・大文字と小文字は区別しない）" {
        getPathUnderFolder "C:\data\見積\2024\a.xlsx" "C:\data\見積" | Should Be "2024\a.xlsx"
        getPathUnderFolder "c:\DATA\見積\" "C:\data\見積" | Should Be ""
        getPathUnderFolder "D:\a.xlsx" "D:\" | Should Be "a.xlsx"
        getPathUnderFolder "\\server\share\a.xlsx" "\\server\share" | Should Be "a.xlsx"
    }

    It "フォルダの下でなければ `$null（フォルダ名の途中では一致しない）" {
        ($null -eq (getPathUnderFolder "C:\data\見積2\a.xlsx" "C:\data\見積")) | Should Be $true
        ($null -eq (getPathUnderFolder "D:\a.xlsx" "C:\data")) | Should Be $true
        ($null -eq (getPathUnderFolder "C:\data" "")) | Should Be $true
    }
}

Describe "getFolderPathAliases / testSameFolder" -Tag Io {
    # ドライブの割り当ては PC によって違うため、テストでは割り当てを差し替える
    $drives = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $drives["Z:"] = "\\server\share"
    $drives["X:"] = "C:\data"

    It "ネットワークドライブと UNC パスを行き来する" {
        @(getFolderPathAliases "Z:\見積" $drives) | Should Be @("Z:\見積", "\\server\share\見積")
        @(getFolderPathAliases "\\server\share\見積" $drives) | Should Be @("\\server\share\見積", "Z:\見積")
    }

    It "subst のドライブと割り当て元のフォルダを行き来する" {
        @(getFolderPathAliases "X:\見積" $drives) | Should Be @("X:\見積", "C:\data\見積")
        @(getFolderPathAliases "C:\data\見積" $drives) | Should Be @("C:\data\見積", "X:\見積")
    }

    It "ドライブ直下・共有直下" {
        @(getFolderPathAliases "Z:\" $drives) | Should Be @("Z:\", "\\server\share")
        @(getFolderPathAliases "\\server\share" $drives) | Should Be @("\\server\share", "Z:\")
    }

    It "割り当ての無いパスは自身だけ" {
        @(getFolderPathAliases "D:\見積" $drives) | Should Be @("D:\見積")
    }

    It "書き方が違っても同じフォルダと分かる" {
        testSameFolder "Z:\見積" "\\server\share\見積" $drives | Should Be $true
        testSameFolder "\\server\share\見積\" "Z:\見積" $drives | Should Be $true
        testSameFolder "X:\見積" "C:\data\見積" $drives | Should Be $true
        testSameFolder "C:\data\見積" "X:\見積" $drives | Should Be $true
        testSameFolder "Z:\見積" "Z:\売上" $drives | Should Be $false
        testSameFolder "C:\見積" "\\server\share\見積" $drives | Should Be $false
    }

    It "この PC のドライブの割り当てを使っても例外にならない" {
        @(getFolderPathAliases "${rootDir}\work")[0] | Should Be "${rootDir}\work"
        testSameFolder "${rootDir}\work" "${rootDir}\work\" | Should Be $true
    }

    It "入れ子のフォルダが分かる（testFolderUnder）" {
        testFolderUnder "C:\data\見積\2024" "C:\data\見積" $drives | Should Be $true
        testFolderUnder "C:\data\見積" "C:\data\見積" $drives | Should Be $true   # 同じフォルダも「中」とする
        testFolderUnder "C:\data\見積" "C:\data\見積\2024" $drives | Should Be $false
        testFolderUnder "C:\data\見積2" "C:\data\見積" $drives | Should Be $false  # 名前の先頭が同じだけ
        testFolderUnder "Z:\見積\2024" "\\server\share\見積" $drives | Should Be $true
        testFolderUnder "C:\data\見積\2024" "X:\見積" $drives | Should Be $true
        testFolderUnder "" "C:\data" $drives | Should Be $false
        testFolderUnder "C:\data" "" $drives | Should Be $false
    }

    It "書き方どおりに同じ・下にあるなら、ドライブの割り当て（CIM）を調べない" {
        # 画面の起動時に毎回呼ばれるため、切断されたネットワークドライブがあっても待たされないようにする
        Mock getDriveTargets { throw "ドライブの割り当てを調べた" }
        testSameFolder "C:\data\見積" "c:\DATA\見積\" | Should Be $true
        testFolderUnder "C:\data\見積\2024" "C:\data\見積" | Should Be $true
        Assert-MockCalled getDriveTargets -Times 0
    }
}

Describe "getParentFolderPath" -Tag Io {
    It "1つ上のフォルダを返す" {
        getParentFolderPath "C:\data\見積\2025" | Should Be "C:\data\見積"
        getParentFolderPath "C:\data" | Should Be "C:\"
        getParentFolderPath "\\server\share\見積\2025" | Should Be "\\server\share\見積"
        getParentFolderPath "\\server\share\見積" | Should Be "\\server\share"
    }

    It "これ以上たどれないフォルダ（ドライブ直下・共有フォルダ直下）は空文字列" {
        getParentFolderPath "C:\" | Should Be ""
        getParentFolderPath "C:" | Should Be ""
        getParentFolderPath "\\server\share" | Should Be ""
        getParentFolderPath "\\server" | Should Be ""
        getParentFolderPath "" | Should Be ""
    }
}

Describe "joinFolderPath" -Tag Io {
    It "ドライブ直下・共有フォルダ直下でも \ が重ならない" {
        joinFolderPath "C:\" "data" | Should Be "C:\data"
        joinFolderPath "C:\data" "見積" | Should Be "C:\data\見積"
        joinFolderPath "\\server\share" "見積" | Should Be "\\server\share\見積"
    }
}

Describe "getFolderEntries / testHasSubFolders" -Tag Io {
    $root = "$TestDrive\folderEntries"
    New-Item -ItemType Directory -Path "$root\B社" -Force | Out-Null
    New-Item -ItemType Directory -Path "$root\A社" -Force | Out-Null
    New-Item -ItemType Directory -Path "$root\A社\2025" -Force | Out-Null
    New-Item -ItemType Directory -Path "$root\_隠しフォルダ" -Force | Out-Null
    (Get-Item "$root\_隠しフォルダ").Attributes = "Directory, Hidden"
    Set-Content "$root\見積.xlsx" "x" -Encoding UTF8
    Set-Content "$root\メモ.txt" "x" -Encoding UTF8
    Set-Content "$root\隠し.txt" "x" -Encoding UTF8
    (Get-Item "$root\隠し.txt").Attributes = "Hidden"

    It "フォルダを先に、それぞれ名前順で返す" {
        $result = getFolderEntries $root
        # 名前順は Windows の並び（カタカナが漢字より先）
        @($result.Entries | ForEach-Object { $_.Name }) -join "," | Should Be "A社,B社,メモ.txt,見積.xlsx"
    }

    It "フォルダ数と Office ファイル数を返す" {
        $result = getFolderEntries $root
        $result.FolderCount | Should Be 2
        $result.OfficeCount | Should Be 1
        $result.Truncated | Should Be $false
        $result.Error | Should Be ""
    }

    It "隠し・システムのフォルダとファイルは返さない（エクスプローラーの既定と同じ）" {
        $names = @((getFolderEntries $root).Entries | ForEach-Object { $_.Name })
        $names -contains "_隠しフォルダ" | Should Be $false
        $names -contains "隠し.txt" | Should Be $false
    }

    It "パスと種別を返す" {
        $entry = @((getFolderEntries $root).Entries | Where-Object { $_.Name -eq "見積.xlsx" })[0]
        $entry.Path | Should Be "$root\見積.xlsx"
        $entry.IsFolder | Should Be $false
        $entry.IsOffice | Should Be $true
        $entry.Updated | Should Not BeNullOrEmpty
    }

    It "foldersOnly はフォルダだけを返す（ツリーの読み込み用）" {
        $result = getFolderEntries $root -foldersOnly
        @($result.Entries | ForEach-Object { $_.Name }) -join "," | Should Be "A社,B社"
        $result.OfficeCount | Should Be 0
    }

    It "件数が多いときは打ち切り、Truncated を立てる" {
        $result = getFolderEntries $root 2
        @($result.Entries).Count | Should Be 2
        $result.Truncated | Should Be $true
    }

    It "開けないフォルダは Error に理由を入れる（一覧は空）" {
        $result = getFolderEntries "$root\ありません"
        @($result.Entries).Count | Should Be 0
        $result.Error | Should Match "見つかりません"
    }

    It "フォルダを指定しなければ Error を返す" {
        (getFolderEntries "").Error | Should Match "指定してください"
    }

    It "サブフォルダがあるかを返す（ツリーの ▷ の判定）" {
        testHasSubFolders $root | Should Be $true
        testHasSubFolders "$root\B社" | Should Be $false
        testHasSubFolders "$root\ありません" | Should Be $false
    }
}

Describe "getFolderLeafName" -Tag Io {
    It "フォルダ名を返す（ドライブ直下はドライブ名、UNC は共有名、末尾の \ は無視）" {
        getFolderLeafName "C:\data\見積\" | Should Be "見積"
        getFolderLeafName "D:\" | Should Be "D"
        getFolderLeafName "\server\share" | Should Be "share"
    }
}

Describe "normalizeFolderPath（区切りだけ・ドライブ直下に戻るパス）" -Tag Io {
    It "\ だけのパスは空にする" {
        normalizeFolderPath "\" | Should Be ""
        normalizeFolderPath "\\?\" | Should Be ""
    }

    It ".. でドライブ直下まで戻ったら \ を付ける" {
        normalizeFolderPath "C:\data\.." | Should Be "C:\"
    }
}

Describe "joinFolderPath / getParentFolderPath（フォルダが無い・相対パス）" -Tag Io {
    It "フォルダが空なら名前だけを返す" {
        joinFolderPath "" "見積" | Should Be "見積"
    }

    It "1 つ上の無い相対パスは空文字列" {
        getParentFolderPath "見積" | Should Be ""
    }
}

Describe "getFolderEntries（隠しファイルばかりのフォルダ・開けないパス）" -Tag Io {
    It "隠しファイルばかりでも、見た件数が上限の 10 倍を超えたら打ち切る" {
        $root = "$TestDrive\隠しばかり"
        [System.IO.Directory]::CreateDirectory($root) | Out-Null
        for ($i = 0; $i -lt 12; $i++) {
            $path = "$root\隠し$i.txt"
            [System.IO.File]::WriteAllText($path, "x")
            (Get-Item -LiteralPath $path).Attributes = "Hidden"
        }
        $result = getFolderEntries $root 1
        @($result.Entries).Count | Should Be 0
        $result.Truncated | Should Be $true
    }

    It "パスとして開けない書き方は、理由を付けて Error に入れる" {
        $result = getFolderEntries "C:\a<b>"
        @($result.Entries).Count | Should Be 0
        $result.Error | Should Match "^フォルダを開けません（"
    }
}

Describe "getQuickFolders" -Tag Io {
    It "実際にあるフォルダだけを、重複させずに返す" {
        $items = @(getQuickFolders)
        foreach ($item in $items) {
            [System.IO.Directory]::Exists($item.Path) | Should Be $true
        }
        @($items | ForEach-Object { $_.Path } | Sort-Object -Unique).Count | Should Be $items.Count
        @($items | Where-Object { @("デスクトップ", "ドキュメント", "ダウンロード") -notcontains $_.Name }).Count | Should Be 0
    }
}

Describe "getComputerFolders" -Tag Io {
    It "使えるドライブを「名前 (C:)」とドライブ直下のパスで返す" {
        $empty = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
        $items = @(getComputerFolders $empty)
        $system = $env:SystemDrive
        $item = @($items | Where-Object { $_.Path -eq "${system}\" })
        $item.Count | Should Be 1
        $item[0].Name | Should Match " \($([regex]::Escape($system))\)$"
        foreach ($other in $items) {
            $other.Path | Should Match "^[A-Z]:\\$"
        }
    }
}
