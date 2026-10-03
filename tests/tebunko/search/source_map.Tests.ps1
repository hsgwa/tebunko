# 元のファイルの場所（tebunko\search\source_map.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "getIndexNameMap / resolveSourcePath" -Tag Io {
    It "インデックス名からクロール対象フォルダを引き、元のファイルのパスを返す" {
        $path = "$TestDrive\status_map.tsv"
        writeStatusFile @(
            [pscustomobject]@{ Path = "C:\data\見積"; Name = "見積" },
            [pscustomobject]@{ Path = "D:\"; Name = "D" }
        ) @() $path
        $map = getIndexNameMap $path
        $map["見積"] | Should -Be "C:\data\見積"

        $root = "C:\tebunko\work\index"
        $maps = @{ $root = $map; "C:\tebunko\work" = $map }
        $hit = [pscustomobject]@{ Root = $root; RelDir = "見積\2024"; Book = "A社.xlsx" }
        resolveSourcePath $hit $maps | Should -Be "C:\data\見積\2024\A社.xlsx"
        $hit = [pscustomobject]@{ Root = $root; RelDir = "見積\2024\見積\2"; Book = "A社.xlsx" }
        resolveSourcePath $hit $maps | Should -Be "C:\data\見積\2024\見積\2\A社.xlsx"
        $hit = [pscustomobject]@{ Root = $root; RelDir = "d"; Book = "直下.xlsx" }
        resolveSourcePath $hit $maps | Should -Be "D:\直下.xlsx"
        $hit = [pscustomobject]@{ Root = "$TestDrive\none_index"; RelDir = "不明\x"; Book = "a.xlsx" }
        resolveSourcePath $hit @{} | Should -Be $null
    }

    It "設定のインデックス名の場所を、source_folder.txt・取り込み一覧より優先する" {
        # 「名前」と「今の置き場所」を設定で分けて持つため、フォルダを移したら設定の場所だけを見る
        $dir = "$TestDrive\優先\index"
        $settings = "$TestDrive\優先\setting.config"
        $status = "$TestDrive\優先\status.tsv"
        [void](New-Item -ItemType Directory -Path "$dir\見積" -Force)
        writeSourceFolderFile @([pscustomobject]@{ Path = "C:\作った時の場所\見積"; Name = "見積" }) $dir
        writeStatusFile @([pscustomobject]@{ Path = "C:\取り込み一覧の場所\見積"; Name = "見積" }) @() $status
        writeTargetFolders @([pscustomobject]@{ Name = "見積"; Path = "\server\今の場所\見積"; Enabled = $true }) $settings

        $map = getSourceFolderMap $dir $status $settings
        $map["見積"] | Should -Be "\server\今の場所\見積"

        # 取り込まないインデックス（indexSources）も同じように優先する
        setIndexSourceFolder "営業" "E:\今の営業" $settings
        (getSourceFolderMap $dir $status $settings)["営業"] | Should -Be "E:\今の営業"
    }

    It "既定のインデックスは取り込み一覧の記録を使う" {
        # リポジトリの work\content_index を作らないよう、既定のインデックスの場所をテスト用のフォルダに向ける
        $indexDir = "$TestDrive\default_work\content_index"
        $workspace = newTestWorkspace @{} "$TestDrive\default_work"
        [System.IO.Directory]::CreateDirectory($indexDir) | Out-Null
        $status = "$TestDrive\status_default.tsv"
        writeStatusFile @([pscustomobject]@{ Path = "C:\新\見積"; Name = "見積" }) @() $status
        $map = getSourceFolderMap (Resolve-Path -LiteralPath $indexDir).ProviderPath $status "$TestDrive\設定なし.config"
        $map["見積"] | Should -Be "C:\新\見積"
    }
}

Describe "writeSourceFolderFile / readSourceFolderFile / getSourceLocation" -Tag Io {
    It "インデックス名とクロール対象フォルダの対応を、各インデックスのフォルダに書き出して読み込む（説明の行は無視する）" {
        $dir = "$TestDrive\copied[1]\index"
        [void][System.IO.Directory]::CreateDirectory("$dir\見積")
        [void][System.IO.Directory]::CreateDirectory("$dir\D")
        writeSourceFolderFile @(
            [pscustomobject]@{ Path = "C:\data\見積"; Name = "見積" },
            [pscustomobject]@{ Path = "D:\"; Name = "D" }
        ) $dir
        # インデックスのフォルダ直下（全インデックス分）には書かない
        Test-Path -LiteralPath (Join-Path $dir ${sourceFolderFileName}) | Should -Be $false
        (readSourceFolderFile "$dir\見積").Count | Should -Be 1
        (readSourceFolderFile "$dir\見積")["見積"] | Should -Be "C:\data\見積"
        (readSourceFolderFile "$dir\D")["d"] | Should -Be "D:\"
        (readSourceFolderFile "$TestDrive\none_dir").Count | Should -Be 0
    }

    It "別の場所にコピーしたインデックスでも、source_folder.txt から元の場所が分かる" {
        $dir = "$TestDrive\別PC\index"
        [void][System.IO.Directory]::CreateDirectory("$dir\見積")
        writeSourceFolderFile @([pscustomobject]@{ Path = "C:\data\見積"; Name = "見積" }) $dir
        $hit = [pscustomobject]@{ Root = $dir; RelDir = "見積\2024"; Book = "A社.xlsx" }
        $location = getSourceLocation $hit
        $location.Known | Should -Be $true
        $location.Folder | Should -Be "C:\data\見積"
        $location.Rest | Should -Be "2024"
        resolveSourcePath $hit @{} | Should -Be "C:\data\見積\2024\A社.xlsx"
    }

    It "インデックスのフォルダの中に source_folder.txt を書き、そのフォルダだけをコピーしても元の場所が分かる" {
        $dir = "$TestDrive\作った PC\index"
        [void][System.IO.Directory]::CreateDirectory("$dir\見積")
        [void][System.IO.Directory]::CreateDirectory("$dir\営業")
        writeSourceFolderFile @(
            [pscustomobject]@{ Path = "C:\data\見積"; Name = "見積" },
            [pscustomobject]@{ Path = "C:\data\営業"; Name = "営業" }
        ) $dir

        # <インデックス名> のフォルダだけを別の場所（ほかの PC の work\index 直下など）へコピーした場合
        $other = "$TestDrive\別 PC\index"
        [void][System.IO.Directory]::CreateDirectory($other)
        Copy-Item -LiteralPath "$dir\見積" -Destination "$other\見積" -Recurse
        $hit = [pscustomobject]@{ Root = $other; RelDir = "見積\2024"; Book = "A社.xlsx" }
        resolveSourcePath $hit @{} | Should -Be "C:\data\見積\2024\A社.xlsx"
    }

    It "インデックスのフォルダがまだ無ければ、その中には書かない" {
        $dir = "$TestDrive\未取り込み\index"
        writeSourceFolderFile @([pscustomobject]@{ Path = "C:\data\見積"; Name = "見積" }) $dir
        Test-Path -LiteralPath "$dir\見積" | Should -Be $false
        (readSourceFolderFile $dir).Count | Should -Be 0
    }

    It "インデックス名のフォルダを検索対象にした場合は、そのフォルダの source_folder.txt を使う" {
        $dir = "$TestDrive\別PC2\index"
        [void][System.IO.Directory]::CreateDirectory("$dir\見積")
        writeSourceFolderFile @([pscustomobject]@{ Path = "C:\data\見積"; Name = "見積" }) $dir
        $hit = [pscustomobject]@{ Root = "$dir\見積"; RelDir = "2024"; Book = "A社.xlsx" }
        resolveSourcePath $hit @{} | Should -Be "C:\data\見積\2024\A社.xlsx"
        $hit = [pscustomobject]@{ Root = "$dir\見積"; RelDir = ""; Book = "直下.xlsx" }
        resolveSourcePath $hit @{} | Should -Be "C:\data\見積\直下.xlsx"
    }

    It "元の場所が分からなければ Known = false・Folder = 空とし、インデックス名と相対フォルダを返す" {
        $hit = [pscustomobject]@{ Root = "$TestDrive\記録なし\index\"; RelDir = "営業\2024"; Book = "a.xlsx" }
        $location = getSourceLocation $hit
        $location.Known | Should -Be $false
        $location.Name | Should -Be "営業"
        $location.Folder | Should -Be ""
        $location.Rest | Should -Be "2024"
        resolveSourcePath $hit @{} | Should -Be $null
    }

    It "読んだ対応はキャッシュに入れ、次からはファイルを読まない" {
        $dir = "$TestDrive\cache\index"
        [void][System.IO.Directory]::CreateDirectory("$dir\見積")
        writeSourceFolderFile @([pscustomobject]@{ Path = "C:\data\見積"; Name = "見積" }) $dir
        $maps = @{}
        $hit = [pscustomobject]@{ Root = $dir; RelDir = "見積"; Book = "a.xlsx" }
        resolveSourcePath $hit $maps | Should -Be "C:\data\見積\a.xlsx"
        Remove-Item -LiteralPath (Join-Path "$dir\見積" ${sourceFolderFileName})
        resolveSourcePath $hit $maps | Should -Be "C:\data\見積\a.xlsx"
    }
}

Describe "findSourceFileState" -Tag Io {
    BeforeAll {
        $dir = "$TestDrive\fsState\見積"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        [System.IO.File]::WriteAllText("$dir\a.xlsx", "x")
    }

    It "見つかれば State=Found・Path" {
        $location = @{ Folder = $dir; Rest = "" }
        $state = findSourceFileState $location "a.xlsx"
        $state.State | Should -Be "Found"
        $state.Path | Should -Be "$dir\a.xlsx"
    }

    It "記録した場所に無くても、別名（ネットワークドライブ⇔UNC）で見つかれば State=Found・Alias" {
        $drives = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
        $drives["Z:"] = "\\server\share"
        $moved = "$TestDrive\fsStateAlias\見積"
        [System.IO.Directory]::CreateDirectory($moved) | Out-Null
        [System.IO.File]::WriteAllText("$moved\a.xlsx", "x")
        Mock getFolderPathAliases { @("Z:\見積", $moved) }
        Mock testNetworkPath { $true }

        $location = @{ Folder = "Z:\見積"; Rest = "" }
        $state = findSourceFileState $location "a.xlsx" $drives
        $state.State | Should -Be "Found"
        $state.Path | Should -Be "$moved\a.xlsx"
        $state.Alias | Should -Be $moved
    }

    It "見つからなければ State=Missing・Initial に上の今もあるフォルダ" {
        $location = @{ Folder = "$dir\無い"; Rest = "" }
        $state = findSourceFileState $location "a.xlsx"
        $state.State | Should -Be "Missing"
        $state.Initial | Should -Be $dir
    }

    It "接続できない・その他のときは、上のフォルダも別名も調べない（getPathState は 1 回だけ呼ぶ）" -TestCases @(
        @{ name = "接続できない"; state = "Unreachable" }
        @{ name = "その他"; state = "Other" }
    ) {
        param ($name, $state)
        Mock getPathState { @{ State = $state; Message = "テストの失敗" } }
        Mock getFolderPathAliases { throw "調べた" }
        Mock getExistingAncestorFolder { throw "調べた" }

        $location = @{ Folder = "\\server\share\見積"; Rest = "" }
        $result = findSourceFileState $location "a.xlsx"
        $result.State | Should -Be $state
        $result.Message | Should -Be "テストの失敗"
        Should -Invoke getPathState -Times 1 -Exactly
    }

    It "ローカルのパスでは別名を求めない（getDriveTargets を呼ばない）" {
        Mock getDriveTargets { throw "CIM を照会した" }
        $location = @{ Folder = "$dir\無い"; Rest = "" }

        findSourceFileState $location "a.xlsx" | Out-Null
        Should -Invoke getDriveTargets -Times 0 -Exactly
    }
}

Describe "joinSourcePath" -Tag Io {
    It "フォルダ・相対フォルダ・ファイル名をつなぐ" {
        joinSourcePath "C:\data\" "2024\見積" "a.xlsx" | Should -Be "C:\data\2024\見積\a.xlsx"
        joinSourcePath "D:\" "" "a.xlsx" | Should -Be "D:\a.xlsx"
        joinSourcePath "C:\data" "2024" | Should -Be "C:\data\2024"
    }

    It "共有フォルダ直下・名前なしでもつなげる" {
        joinSourcePath "\\server\share" "見積" "a.xlsx" | Should -Be "\\server\share\見積\a.xlsx"
        joinSourcePath "\\server\share\" "" "" | Should -Be "\\server\share"
    }
}

Describe "findMovedSource" -Tag Io {
    BeforeAll {
        $moved = "$TestDrive\移動先\見積 [新]"
        [System.IO.Directory]::CreateDirectory("$moved\2024\A社") | Out-Null
        [System.IO.File]::WriteAllText("$moved\2024\A社\見積.xlsx", "")
    }

    It "インデックスの元のフォルダに当たるフォルダを選んだ場合は、そのフォルダを Root にする" {
        $found = findMovedSource $moved "2024\A社" "見積.xlsx"
        $found.Path | Should -Be "$moved\2024\A社\見積.xlsx"
        $found.Root | Should -Be $moved
    }

    It "ファイルのあるフォルダを選んだ場合は、相対フォルダの分だけ上を Root にする" {
        $found = findMovedSource "$moved\2024\A社\" "2024\A社" "見積.xlsx"
        $found.Path | Should -Be "$moved\2024\A社\見積.xlsx"
        $found.Root | Should -Be $moved
    }

    It "途中のフォルダを選んだ場合も Root は同じになる" {
        (findMovedSource "$moved\2024" "2024\A社" "見積.xlsx").Root | Should -Be $moved
    }

    It "相対フォルダが無ければ、選んだフォルダが Root になる" {
        [System.IO.File]::WriteAllText("$moved\直下.xlsx", "")
        $found = findMovedSource $moved "" "直下.xlsx"
        $found.Path | Should -Be "$moved\直下.xlsx"
        $found.Root | Should -Be $moved
    }

    It "見つからなければ `$null" {
        findMovedSource "$TestDrive\移動先" "2024\B社" "見積.xlsx" | Should -Be $null
    }

    It "相対フォルダの前後・途中に余分な \ があっても同じように探す" {
        $found = findMovedSource $moved "\2024\\A社\" "見積.xlsx"
        $found.Path | Should -Be "$moved\2024\A社\見積.xlsx"
        $found.Root | Should -Be $moved
    }

    It "ファイルと同じ名前のフォルダは、見つかったことにしない" {
        [System.IO.Directory]::CreateDirectory("$moved\2024\A社\フォルダ.xlsx") | Out-Null
        findMovedSource $moved "2024\A社" "フォルダ.xlsx" | Should -Be $null
    }
}
