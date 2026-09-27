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

    It "どちらかがローカルなら、別名で一致しえないため CIM を照会しない（`$drives を渡さないとき）" {
        # getDriveTargets（CIM）が返す別名はネットワークドライブ⇔UNCの組しか無いため、
        # ローカルとネットワークの組は、別名をたどっても一致しない
        Mock getDriveTargets { throw "ドライブの割り当てを調べた" }
        testSameFolder "C:\data\見積" "\\server\share\見積" | Should -Be $false
        testSameFolder "\\server\share\見積" "C:\data\見積" | Should -Be $false
        testFolderUnder "C:\data\見積\2024" "\\server\share\見積" | Should -Be $false
        testFolderUnder "\\server\share\見積\2024" "C:\data\見積" | Should -Be $false
        Should -Invoke getDriveTargets -Times 0 -Exactly
    }

    It "`$drives を渡したとき（テスト用の割り当て）は、ローカルどうしの別名（subst）も今のとおり調べる" {
        testSameFolder "X:\見積" "C:\data\見積" $drives | Should -Be $true
        testFolderUnder "C:\data\見積\2024" "X:\見積" $drives | Should -Be $true
    }
}

Describe "testNetworkPath" -Tag Unit {
    BeforeAll {
        $driveTypes = @{
            "C:\" = [System.IO.DriveType]::Fixed
            "D:\" = [System.IO.DriveType]::Removable
            "E:\" = [System.IO.DriveType]::CDRom
            "R:\" = [System.IO.DriveType]::Ram
            "Z:\" = [System.IO.DriveType]::Network
            "Y:\" = [System.IO.DriveType]::NoRootDirectory  # 切断されたドライブがこう見えることがある
            "U:\" = [System.IO.DriveType]::Unknown
        }
        $driveType = { param ($drive) $driveTypes[$drive] }
    }

    It "<name>" -TestCases @(
        @{ name = "空はローカル扱い"; path = ""; expected = $false }
        @{ name = "UNC はネットワーク"; path = "\\server\share\見積"; expected = $true }
        @{ name = "\\?\UNC\ もネットワーク"; path = "\\?\UNC\server\share\見積"; expected = $true }
        @{ name = "\\?\C:\ はドライブ文字として扱う（Fixed はローカル）"; path = "\\?\C:\data\見積"; expected = $false }
        @{ name = "\\.\ もドライブ文字として扱う"; path = "\\.\C:\data\見積"; expected = $false }
        @{ name = "Fixed はローカル"; path = "C:\data\見積"; expected = $false }
        @{ name = "Removable はローカル"; path = "D:\見積"; expected = $false }
        @{ name = "CDRom はローカル"; path = "E:\見積"; expected = $false }
        @{ name = "Ram はローカル"; path = "R:\見積"; expected = $false }
        @{ name = "Network はネットワーク"; path = "Z:\見積"; expected = $true }
        @{ name = "NoRootDirectory（切断したドライブ）は安全な側でネットワーク"; path = "Y:\見積"; expected = $true }
        @{ name = "Unknown も安全な側でネットワーク"; path = "U:\見積"; expected = $true }
        @{ name = "ドライブ文字の無い相対パスはローカル扱い"; path = "見積\2024"; expected = $false }
        @{ name = "/ を \ にそろえて調べる"; path = "//server/share/見積"; expected = $true }
    ) {
        param ($name, $path, $expected)
        testNetworkPath $path $driveType | Should -Be $expected
    }

    It "ドライブの種類が調べられなければ安全な側（ネットワーク）にする" {
        testNetworkPath "Q:\見積" { param ($drive) throw "調べられない" } | Should -Be $true
    }

    It "既定はこの PC の実際のドライブの種類を見る（C: は通常 Fixed）" {
        testNetworkPath "${rootDir}\work" | Should -Be $false
    }
}

Describe "testAnyNetworkPath" -Tag Unit {
    It "1 つでもネットワークのパスがあれば `$true" {
        testAnyNetworkPath @("C:\data", "\\server\share") | Should -Be $true
    }

    It "すべてローカルなら `$false" {
        testAnyNetworkPath @("C:\data", "D:\見積") | Should -Be $false
    }

    It "空の一覧は `$false" {
        testAnyNetworkPath @() | Should -Be $false
    }
}

Describe "setDriveTargets" -Tag Unit {
    AfterEach {
        setDriveTargets $null
    }

    It "getDriveTargets のキャッシュを外から設定できる（CIM を照会しない）" {
        Mock Get-CimInstance { throw "CIM を照会した" }
        $custom = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
        $custom["Q:"] = "\\server\share"
        setDriveTargets $custom
        (getDriveTargets)["Q:"] | Should -Be "\\server\share"
        Should -Invoke Get-CimInstance -Times 0 -Exactly
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

    It "skipNetwork が `$true でも、ローカルのパスは今どおり調べる" {
        getExistingAncestorFolder "$root\営業部\2024\見積" $true | Should -Be "$root\営業部"
    }

    It "skipNetwork が `$true なら、ネットワークのパスは調べずに空を返す" {
        Mock Test-Path { throw "調べた" }
        getExistingAncestorFolder "\\server\share\営業部" $true | Should -Be ""
        Should -Invoke Test-Path -Times 0 -Exactly
    }

}
