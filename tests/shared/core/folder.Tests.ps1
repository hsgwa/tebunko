# フォルダのパスと一覧（shared\core\folder.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "normalizeFolderPath" -Tag Io {
    BeforeDiscovery {
        # -TestCases は探索のときに作るため、ここでもリポジトリ直下を求めておく
        $rootDir = (Resolve-Path "$PSScriptRoot\..\..\..").Path
    }

    # cases = 渡すパス（path）と期待する結果（expected）の組
    It "<name>" -TestCases @(
        @{ name = "前後の空白・引用符と末尾の \ を取り除き、ドライブ直下は \ を残す"; cases = @(
                @{ path = '  "C:\data\"  '; expected = "C:\data" }
                @{ path = "D:\"; expected = "D:\" }
                @{ path = "D:"; expected = "D:\" }
                @{ path = "\\server\share\"; expected = "\\server\share" }
            )
        }
        @{ name = "/ を \ にそろえ、長いパス用の \\?\ ・ \\?\UNC\ を外す"; cases = @(
                @{ path = "C:/data/見積"; expected = "C:\data\見積" }
                @{ path = "//server/share/見積"; expected = "\\server\share\見積" }
                @{ path = "\\?\C:\data\見積"; expected = "C:\data\見積" }
                @{ path = "\\?\UNC\server\share\見積"; expected = "\\server\share\見積" }
            )
        }
        @{ name = "重なった \ ・ . ・ .. を解決する"; cases = @(
                @{ path = "C:\data\\見積"; expected = "C:\data\見積" }
                @{ path = "C:\data\.\見積"; expected = "C:\data\見積" }
                @{ path = "C:\data\売上\..\見積"; expected = "C:\data\見積" }
                @{ path = "\\server\share\売上\..\見積"; expected = "\\server\share\見積" }
            )
        }
        @{ name = ".. でドライブ直下まで戻ったら \ を付ける"; cases = @(
                @{ path = "C:\data\.."; expected = "C:\" }
            )
        }
        @{ name = "\ だけのパスは空にする"; cases = @(
                @{ path = "\"; expected = "" }
                @{ path = "\\?\"; expected = "" }
            )
        }
        @{ name = "相対パスは tebunko のフォルダからとみなす"; cases = @(
                @{ path = "work\index"; expected = "${rootDir}\work\index" }
                @{ path = ".\work\index"; expected = "${rootDir}\work\index" }
            )
        }
        @{ name = "パスとして解釈できない場合は書かれたとおりに扱う"; cases = @(
                @{ path = "C:\data*"; expected = "C:\data*" }
                @{ path = "\\server"; expected = "\\server" }
                # \ ひとつで始まるパスは、書き間違えた UNC パスのことが多いため、今のドライブのパスに直さない
                @{ path = "\server\share"; expected = "\server\share" }
            )
        }
    ) {
        param ($name, $cases)
        foreach ($case in $cases) {
            normalizeFolderPath $case.path | Should -Be $case.expected
        }
    }

    It "環境変数を展開する" {
        $env:tebunko_TEST_FOLDER = "C:\data\見積"
        try {
            normalizeFolderPath "%tebunko_TEST_FOLDER%" | Should -Be "C:\data\見積"
            normalizeFolderPath "%tebunko_TEST_FOLDER%\2024" | Should -Be "C:\data\見積\2024"
        } finally {
            Remove-Item Env:\tebunko_TEST_FOLDER
        }
    }
}

Describe "getPathUnderFolder" -Tag Io {
    It "フォルダからの相対パスを返す（フォルダ自身は空。末尾の \ ・大文字と小文字は区別しない）" {
        getPathUnderFolder "C:\data\見積\2024\a.xlsx" "C:\data\見積" | Should -Be "2024\a.xlsx"
        getPathUnderFolder "c:\DATA\見積\" "C:\data\見積" | Should -Be ""
        getPathUnderFolder "D:\a.xlsx" "D:\" | Should -Be "a.xlsx"
        getPathUnderFolder "\\server\share\a.xlsx" "\\server\share" | Should -Be "a.xlsx"
    }

    It "フォルダの下でなければ `$null（フォルダ名の途中では一致しない）" {
        ($null -eq (getPathUnderFolder "C:\data\見積2\a.xlsx" "C:\data\見積")) | Should -Be $true
        ($null -eq (getPathUnderFolder "D:\a.xlsx" "C:\data")) | Should -Be $true
        ($null -eq (getPathUnderFolder "C:\data" "")) | Should -Be $true
    }
}

Describe "getFolderPathAliases / testSameFolder" -Tag Io {
    # ドライブの割り当ては PC によって違うため、テストでは割り当てを差し替える
    BeforeAll {
        $drives = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
        $drives["Z:"] = "\\server\share"
        $drives["X:"] = "C:\data"
    }

    It "ネットワークドライブと UNC パスを行き来する" {
        @(getFolderPathAliases "Z:\見積" $drives) | Should -Be @("Z:\見積", "\\server\share\見積")
        @(getFolderPathAliases "\\server\share\見積" $drives) | Should -Be @("\\server\share\見積", "Z:\見積")
    }

    It "subst のドライブと割り当て元のフォルダを行き来する" {
        @(getFolderPathAliases "X:\見積" $drives) | Should -Be @("X:\見積", "C:\data\見積")
        @(getFolderPathAliases "C:\data\見積" $drives) | Should -Be @("C:\data\見積", "X:\見積")
    }

    It "ドライブ直下・共有直下" {
        @(getFolderPathAliases "Z:\" $drives) | Should -Be @("Z:\", "\\server\share")
        @(getFolderPathAliases "\\server\share" $drives) | Should -Be @("\\server\share", "Z:\")
    }

    It "割り当ての無いパスは自身だけ" {
        @(getFolderPathAliases "D:\見積" $drives) | Should -Be @("D:\見積")
    }

    It "書き方が違っても同じフォルダと分かる" {
        testSameFolder "Z:\見積" "\\server\share\見積" $drives | Should -Be $true
        testSameFolder "\\server\share\見積\" "Z:\見積" $drives | Should -Be $true
        testSameFolder "X:\見積" "C:\data\見積" $drives | Should -Be $true
        testSameFolder "C:\data\見積" "X:\見積" $drives | Should -Be $true
        testSameFolder "Z:\見積" "Z:\売上" $drives | Should -Be $false
        testSameFolder "C:\見積" "\\server\share\見積" $drives | Should -Be $false
    }

    It "この PC のドライブの割り当てを使っても例外にならない" {
        @(getFolderPathAliases "${rootDir}\work")[0] | Should -Be "${rootDir}\work"
        testSameFolder "${rootDir}\work" "${rootDir}\work\" | Should -Be $true
    }

    It "入れ子のフォルダが分かる（testFolderUnder）" {
        testFolderUnder "C:\data\見積\2024" "C:\data\見積" $drives | Should -Be $true
        testFolderUnder "C:\data\見積" "C:\data\見積" $drives | Should -Be $true   # 同じフォルダも「中」とする
        testFolderUnder "C:\data\見積" "C:\data\見積\2024" $drives | Should -Be $false
        testFolderUnder "C:\data\見積2" "C:\data\見積" $drives | Should -Be $false  # 名前の先頭が同じだけ
        testFolderUnder "Z:\見積\2024" "\\server\share\見積" $drives | Should -Be $true
        testFolderUnder "C:\data\見積\2024" "X:\見積" $drives | Should -Be $true
        testFolderUnder "" "C:\data" $drives | Should -Be $false
        testFolderUnder "C:\data" "" $drives | Should -Be $false
    }

    It "書き方どおりに同じ・下にあるなら、ドライブの割り当て（CIM）を調べない" {
        # 画面の起動時に毎回呼ばれるため、切断されたネットワークドライブがあっても待たされないようにする
        Mock getDriveTargets { throw "ドライブの割り当てを調べた" }
        testSameFolder "C:\data\見積" "c:\DATA\見積\" | Should -Be $true
        testFolderUnder "C:\data\見積\2024" "C:\data\見積" | Should -Be $true
        Should -Invoke getDriveTargets -Times 0 -Exactly
    }
}

Describe "getFolderLeafName" -Tag Io {
    It "フォルダ名を返す（ドライブ直下はドライブ名、UNC は共有名、末尾の \ は無視）" {
        getFolderLeafName "C:\data\見積\" | Should -Be "見積"
        getFolderLeafName "D:\" | Should -Be "D"
        getFolderLeafName "\server\share" | Should -Be "share"
    }
}

Describe "getExistingAncestorFolder" -Tag Io {
    BeforeAll {
        $root = "$TestDrive\共有"
        New-Item -ItemType Directory -Force -Path "$root\営業部" | Out-Null
    }

    It "フォルダがあればそのまま返す" {
        getExistingAncestorFolder "$root\営業部" | Should -Be "$root\営業部"
    }

    It "フォルダが無ければ、その上の今もあるフォルダを返す" {
        getExistingAncestorFolder "$root\営業部\2024\見積" | Should -Be "$root\営業部"
    }

    It "上のどこにも無ければ空を返す" {
        # 使われていないドライブ名を選ぶ（無ければこのケースは確かめない）
        $used = @([System.IO.DriveInfo]::GetDrives() | ForEach-Object { $_.Name.Substring(0, 1) })
        $free = @([char[]]"QRSTUVWXYZ" | Where-Object { $used -notcontains "$_" })
        if ($free.Count -gt 0) {
            getExistingAncestorFolder "$($free[0]):\営業部\見積" | Should -Be ""
        }
    }

    It "空のパスは空を返す" {
        getExistingAncestorFolder "" | Should -Be ""
    }
}
