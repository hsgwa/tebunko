# インデックスの集計・一覧・取り込み（tebunko\index\index_store.ps1。改名・削除は index_catalog.Tests.ps1）のテスト
BeforeDiscovery {
    # -TestCases の表が使う一覧の状態（$stateDone など）
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "getIndexStats" -Tag Io {
    It "インデックス名ごとに件数と最終取り込み日時を集計する" {
        $path = "$TestDrive\stats.tsv"
        writeStatusFile @(
            [pscustomobject]@{ Path = "C:\data"; Name = "営業" },
            [pscustomobject]@{ Path = "D:\tech"; Name = "技術" }
        ) @(
            (newStatusRow "営業\a.xlsx" "2025/01/10 12:34:56" "1" $stateDone "1" "2026/09/18 10:00:00"),
            (newStatusRow "営業\b.xlsx" "2025/01/10 12:34:56" "1" $stateFailed "" "2026/09/19 11:00:00" "原因"),
            (newStatusRow "営業\c.docx" "2025/01/10 12:34:56" "1" $stateNew),
            (newStatusRow "技術\d.pptx" "2025/01/10 12:34:56" "1" $stateDone "2" "2026/09/17 09:00:00")
        ) $path

        $stats = getIndexStats (readStatusFile $path).Rows
        $stats["営業"].Total | Should -Be 3
        $stats["営業"].Done | Should -Be 1
        $stats["営業"].Failed | Should -Be 1
        $stats["営業"].Pending | Should -Be 1
        $stats["営業"].LastIngested | Should -Be "2026/09/19 11:00:00"
        $stats["技術"].Total | Should -Be 1
        $stats["技術"].LastIngested | Should -Be "2026/09/17 09:00:00"
    }

}

Describe "testStatusRowFailed" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "済は失敗に数えない"; state = $stateDone; errorText = ""; expected = $false }
        @{ name = "失敗は失敗に数える"; state = $stateFailed; errorText = "原因"; expected = $true }
        @{ name = "エラーが空の失敗も失敗に数える"; state = $stateFailed; errorText = ""; expected = $true }
        @{ name = "エラーが空の未取り込みは待ち（失敗に数えない）"; state = $stateNew; errorText = ""; expected = $false }
        @{ name = "エラー（理由）が書いてある未取り込みは失敗に数える"; state = $stateNew; errorText = ${postponedReasonRunning}; expected = $true }
    ) {
        param ($name, $state, $errorText, $expected)
        testStatusRowFailed (newStatusRow "営業\a.ppt" "" "" $state "" "" $errorText) | Should -Be $expected
    }
}

Describe "getIndexStats（エラーが書いてある未取り込み）" -Tag Unit {
    It "「未取り込み」でエラーが書いてある行は Failed に数え、Pending には数えない" {
        $rows = [ordered]@{
            "営業\a.ppt" = (newStatusRow "営業\a.ppt" "" "" $stateNew "" "" ${postponedReasonInUse})
            "営業\b.ppt" = (newStatusRow "営業\b.ppt" "" "" $stateNew)
            "営業\c.ppt" = (newStatusRow "営業\c.ppt" "" "" $stateDone "1" "2026/09/19 10:00:00")
        }
        $stats = getIndexStats $rows
        $stats["営業"].Total | Should -Be 3
        $stats["営業"].Failed | Should -Be 1
        $stats["営業"].Pending | Should -Be 1
        $stats["営業"].Done | Should -Be 1
    }
}

Describe "getSearchIndexes" -Tag Io {
    It "インデックスのフォルダが無ければ空" {
        @(getSearchIndexes "$TestDrive\無いフォルダ\index").Count | Should -Be 0
    }

    It "work\index 直下のフォルダをインデックス 1 件として返し、元のフォルダも返す" {
        $dir = "$TestDrive\一覧\index"
        $settings = "$TestDrive\一覧\setting.config"
        foreach ($name in @("見積", "営業")) {
            [void](New-Item -ItemType Directory -Path "$dir\$name" -Force)
        }
        writeTargetFolders @(
            [pscustomobject]@{ Name = "見積"; Path = "C:\data\見積"; Enabled = $true },
            [pscustomobject]@{ Name = "営業"; Path = "C:\data\営業"; Enabled = $true }) $settings

        $indexes = @(getSearchIndexes $dir "$TestDrive\取り込み一覧なし.tsv" $settings)
        $indexes.Count | Should -Be 2
        # ［1 インデックス作成］の一覧と同じ並び
        $indexes[0].Name | Should -Be "見積"
        $indexes[0].SourcePath | Should -Be "C:\data\見積"
        $indexes[0].Path | Should -Be (Resolve-Path -LiteralPath "$dir\見積").ProviderPath
        $indexes[1].Name | Should -Be "営業"
        $indexes[1].SourcePath | Should -Be "C:\data\営業"
    }

    It "一覧に無いインデックス（コピーしたものなど）も名前順で後ろに並べる" {
        $dir = "$TestDrive\コピー\index"
        $settings = "$TestDrive\コピー\setting.config"
        foreach ($name in @("報告書", "あとから", "見積")) {
            [void](New-Item -ItemType Directory -Path "$dir\$name" -Force)
        }
        writeTargetFolders @([pscustomobject]@{ Name = "見積"; Path = "C:\data\見積"; Enabled = $true }) $settings

        $indexes = @(getSearchIndexes $dir "$TestDrive\取り込み一覧なし.tsv" $settings)
        @($indexes | ForEach-Object { $_.Name }) -join "," | Should -Be "見積,あとから,報告書"
        # 元のフォルダが分からないものは空
        $indexes[1].SourcePath | Should -Be ""
    }

    It "一覧にも取り込み一覧にも無いインデックスは、そのフォルダの source_folder.txt から元のフォルダを読む" {
        $dir = "$TestDrive\コピー2\index"
        $settings = "$TestDrive\コピー2\setting.config"
        [void](New-Item -ItemType Directory -Path "$dir\営業" -Force)
        # ほかの PC で作ったインデックスをフォルダごとコピーした状態
        writeSourceFolderFile @([pscustomobject]@{ Path = "\\server\営業"; Name = "営業" }) $dir

        $indexes = @(getSearchIndexes $dir "$TestDrive\取り込み一覧なし.tsv" $settings)
        $indexes.Count | Should -Be 1
        $indexes[0].Name | Should -Be "営業"
        $indexes[0].SourcePath | Should -Be "\\server\営業"
    }
}

Describe "getIndexNameMap（見出し行まで読む）" -Tag Io {
    It "クロール対象フォルダの行だけを読み、見出し行の後は読まない" {
        $path = "$TestDrive\name_map.tsv"
        writeListFile $path @(
            "${statusFolderKey}`tC:\data\見積`t見積",
            "${statusFolderKey}`tC:\old",
            ($statusColumns -join "`t"),
            "${statusFolderKey}`tC:\x`tx")
        $map = getIndexNameMap $path
        $map.Count | Should -Be 1
        $map["見積"] | Should -Be "C:\data\見積"
        (getIndexNameMap "$TestDrive\none_status.tsv").Count | Should -Be 0
    }
}

Describe "getIndexTsvCounts / testIndexComplete" -Tag Io {
    BeforeAll {
        function newTestIndex {
            # テスト用のインデックス（work\index 相当）を作る
            param ([string]$dir)

            [System.IO.Directory]::CreateDirectory("$dir\営業\2024\A社.xlsx") | Out-Null
            writeListFile "$dir\営業\2024\A社.xlsx\明細.tsv" @("a")
            writeListFile "$dir\営業\2024\A社.xlsx\表紙.tsv" @("b")
            [System.IO.Directory]::CreateDirectory("$dir\営業\空.xlsx") | Out-Null
            [System.IO.Directory]::CreateDirectory("$dir\営業\資料.docx") | Out-Null
            writeListFile "$dir\営業\資料.docx\ページ001.tsv" @("c")
        }
    }

    It "元のファイル1つ分のフォルダごとにTSVの数を数える" {
        $dir = "$TestDrive\index1"
        newTestIndex $dir
        $counts = getIndexTsvCounts $dir
        $counts["営業\2024\A社.xlsx"] | Should -Be 2
        $counts["営業\資料.docx"] | Should -Be 1
    }

    It "TSVの無いフォルダ（内容が空のファイル）は0件として数える" {
        $dir = "$TestDrive\index2"
        newTestIndex $dir
        $counts = getIndexTsvCounts $dir
        $counts.ContainsKey("営業\空.xlsx") | Should -Be $true
        $counts["営業\空.xlsx"] | Should -Be 0
    }

    It "大文字・小文字を区別しない" {
        $dir = "$TestDrive\index3"
        newTestIndex $dir
        (getIndexTsvCounts $dir)["営業\2024\a社.XLSX"] | Should -Be 2
    }

    It "インデックスのフォルダの直下のTSVは数えない" {
        $dir = "$TestDrive\index4"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        writeListFile "$dir\ブック.xlsx_シート.tsv" @("a")
        (getIndexTsvCounts $dir).Count | Should -Be 0
    }

    # newTestIndex のインデックスに対して、取り込み一覧の行（相対パス・TSV数）がそろっているか
    It "<name>" -TestCases @(
        @{ name = "TSVがそろっていれば「済」のままにする"; relPath = "営業\2024\A社.xlsx"; tsvCount = "2"; complete = $true }
        @{ name = "インデックスのフォルダを直接削除した場合は、そろっていないとする"; relPath = "営業\2024\消えたブック.xlsx"; tsvCount = "3"; complete = $false }
        @{ name = "TSVが足りない場合も、そろっていないとする"; relPath = "営業\2024\A社.xlsx"; tsvCount = "5"; complete = $false }
        @{ name = "内容が空のファイル（TSV 0 件）は、フォルダがあればそろっているとする"; relPath = "営業\空.xlsx"; tsvCount = "0"; complete = $true }
    ) {
        param ($name, $relPath, $tsvCount, $complete)
        $dir = "$TestDrive\index_complete"
        newTestIndex $dir
        $row = newStatusRow $relPath "2025/01/10 12:34:56" "100" ${stateDone} $tsvCount
        testIndexComplete $row $row.相対パス (getIndexTsvCounts $dir) | Should -Be $complete
    }

    It "本文インデックスのファイルは、フォルダと拡張子ごとに数える（0 バイトは壊れているとする）" {
        $dir = "$TestDrive\index_content_index1"
        newTestIndex $dir
        writeListFile "$dir\営業\2024\content_index.xlsx.001.tsv" @("a")
        [System.IO.File]::WriteAllText("$dir\営業\2024\content_index.docx.001.tsv", "")
        $counts = getIndexTsvCounts $dir
        # 番号を除いた「フォルダ\content_index.<拡張子>」で数え、どれか 1 つでも 0 バイトなら壊れているとする
        writeListFile "$dir\営業\2024\content_index.xlsx.002.tsv" @("b")
        writeListFile "$dir\営業\2024\content_index.docx.002.tsv" @("c")
        $counts = getIndexTsvCounts $dir
        $counts["営業\2024\content_index.xlsx"] | Should -Be 1
        $counts["営業\2024\content_index.docx"] | Should -Be ${indexBrokenCount}
        # 本文インデックスのファイルはフォルダの TSV の数には入れない
        $counts["営業\2024"] | Should -Be 0
    }

    It "本文インデックスのファイルがあれば、元のファイルごとのフォルダが無くても「済」のままにする" {
        $dir = "$TestDrive\index_content_index2"
        writeListFile "$dir\営業\2024\content_index.xlsx.001.tsv" @("a")
        [System.IO.File]::WriteAllText("$dir\営業\2024\content_index.docx.001.tsv", "")
        $counts = getIndexTsvCounts $dir
        $row = newStatusRow "営業\2024\B社.xlsx" "2025/01/10 12:34:56" "100" ${stateDone} "3"
        testIndexComplete $row $row.相対パス $counts | Should -Be $true
        # 拡張子の本文インデックスのファイルが無い・壊れているなら、そろっていないとする
        $row = newStatusRow "営業\2024\C社.pptx" "2025/01/10 12:34:56" "100" ${stateDone} "1"
        testIndexComplete $row $row.相対パス $counts | Should -Be $false
        $row = newStatusRow "営業\2024\D社.docx" "2025/01/10 12:34:56" "100" ${stateDone} "1"
        testIndexComplete $row $row.相対パス $counts | Should -Be $false
    }

    It "0 バイトのTSVがあるフォルダは、壊れているとして作り直す" {
        $dir = "$TestDrive\index9"
        newTestIndex $dir
        # 書き込みの途中で電源が落ちた場合など（空のシート・ページは保存しないため、0 バイトのTSVは異常）
        [System.IO.File]::WriteAllBytes("$dir\営業\2024\A社.xlsx\途中.tsv", (New-Object byte[] 0))
        $counts = getIndexTsvCounts $dir
        $counts["営業\2024\A社.xlsx"] | Should -Be ${indexBrokenCount}
        $counts["営業\資料.docx"] | Should -Be 1   # ほかのファイルは巻き込まない

        $row = newStatusRow "営業\2024\A社.xlsx" "2025/01/10 12:34:56" "100" ${stateDone} "2"
        testIndexComplete $row $row.相対パス $counts | Should -Be $false
        $other = newStatusRow "営業\資料.docx" "2025/01/10 12:34:56" "100" ${stateDone} "1"
        testIndexComplete $other $other.相対パス $counts | Should -Be $true
    }

    It "0 バイトのTSVが先に見つかっても、後のTSVで数え直さない" {
        $dir = "$TestDrive\index10"
        [System.IO.Directory]::CreateDirectory("$dir\営業\B社.xlsx") | Out-Null
        [System.IO.File]::WriteAllBytes("$dir\営業\B社.xlsx\001_途中.tsv", (New-Object byte[] 0))
        writeListFile "$dir\営業\B社.xlsx\002_あと.tsv" @("a")
        (getIndexTsvCounts $dir)["営業\B社.xlsx"] | Should -Be ${indexBrokenCount}
    }

    It "TSVの数を記録していない行・数えられなかった場合は確認しない" {
        $row = newStatusRow "営業\2024\A社.xlsx" "2025/01/10 12:34:56" "100" ${stateDone} ""
        testIndexComplete $row $row.相対パス (getIndexTsvCounts "$TestDrive\none_index2") | Should -Be $true
        $done = newStatusRow "営業\2024\A社.xlsx" "2025/01/10 12:34:56" "100" ${stateDone} "2"
        testIndexComplete $done $done.相対パス $null | Should -Be $true
    }
}

Describe "publishIndexFiles" -Tag Io {
    It "以前のインデックスは残さず入れ替える（シートの削除・名前変更に追従する）" {
        $from = "$TestDrive\pub2\tmp"
        [System.IO.Directory]::CreateDirectory($from) | Out-Null
        $bookDir = "$TestDrive\pub2\index\営業\A社.xlsx"
        [System.IO.Directory]::CreateDirectory($bookDir) | Out-Null
        writeListFile "$bookDir\前のシート.tsv" @("old")
        writeListFile "$from\新しいシート.tsv" @("new")

        publishIndexFiles $from $bookDir "$TestDrive\pub2\出力\A社.xlsx"

        Test-Path -LiteralPath "$bookDir\前のシート.tsv" | Should -Be $false
        Test-Path -LiteralPath "$bookDir\新しいシート.tsv" | Should -Be $true
        # 入れた bookDir とその TSV は、親から属性を継がないため、ここで NotContentIndexed が付く
        ([System.IO.File]::GetAttributes($bookDir) -band [System.IO.FileAttributes]::NotContentIndexed) | Should -Not -Be 0
        ([System.IO.File]::GetAttributes("$bookDir\新しいシート.tsv") -band [System.IO.FileAttributes]::NotContentIndexed) | Should -Not -Be 0
    }

    It "TSVが1件も無ければ、空のフォルダにする（内容が空のファイル）" {
        $from = "$TestDrive\pub3\tmp"
        [System.IO.Directory]::CreateDirectory($from) | Out-Null
        $bookDir = "$TestDrive\pub3\index\営業\空.xlsx"

        publishIndexFiles $from $bookDir "$TestDrive\pub3\出力\空.xlsx"

        Test-Path -LiteralPath $bookDir -PathType Container | Should -Be $true
        @(Get-ChildItem -LiteralPath $bookDir -Filter "*.tsv").Count | Should -Be 0
    }

    It "前回の出力用フォルダが残っていても入れ替えられる" {
        $from = "$TestDrive\pub4\tmp"
        [System.IO.Directory]::CreateDirectory($from) | Out-Null
        writeListFile "$from\明細.tsv" @("a")
        $staging = "$TestDrive\pub4\出力\A社.xlsx"
        [System.IO.Directory]::CreateDirectory($staging) | Out-Null
        writeListFile "$staging\前回の残り.tsv" @("old")
        $bookDir = "$TestDrive\pub4\index\営業\A社.xlsx"

        publishIndexFiles $from $bookDir $staging

        @(Get-ChildItem -LiteralPath $bookDir -Filter "*.tsv" | ForEach-Object { $_.Name }) | Should -Be "明細.tsv"
    }
}

Describe "publishIndexFiles / getIndexTsvCounts（まれな状況）" -Tag Io {
    BeforeAll {
        $storePath = "${scriptsDir}\tebunko\index\index_store.ps1"
        function script:findStoreLine {
            param ([string]$pattern)
            return @(Select-String -LiteralPath $storePath -Pattern $pattern)[0].LineNumber
        }
    }

    It "フォルダごと移せないとき（別のドライブへのリンクなど）は、1 件ずつ移して集めたフォルダを消す" {
        $from = "$TestDrive\pub\tmp"
        [System.IO.Directory]::CreateDirectory($from) | Out-Null
        writeListFile "$from\明細.tsv" @("a")
        writeListFile "$from\表紙.tsv" @("b")
        $bookDir = "$TestDrive\pub\index\営業\A社.xlsx"
        $staging = "$TestDrive\pub\出力\A社.xlsx"
        # フォルダごとの移動の直前に移し先を作っておくと、移動が IOException になる
        $point = Set-PSBreakpoint -Script $storePath -Line (findStoreLine '^\s+\[System\.IO\.Directory\]::Move\(\(toLongPath \$stagingDir\)') -Action {
            [System.IO.Directory]::CreateDirectory((toLongPath $bookDir)) | Out-Null
        }
        try {
            publishIndexFiles $from $bookDir $staging
        } finally {
            Remove-PSBreakpoint -Breakpoint $point
        }

        $names = @(Get-ChildItem -LiteralPath $bookDir -Filter "*.tsv" | ForEach-Object { $_.Name })
        $names.Count | Should -Be 2
        ($names -contains "明細.tsv") | Should -Be $true
        ($names -contains "表紙.tsv") | Should -Be $true
        [System.IO.Directory]::Exists($staging) | Should -Be $false
    }

    It "列挙の途中で失敗したら（アクセス権が無い等）null を返す（呼び出し元は確認を省く）" {
        $dir = "$TestDrive\cnt\index"
        [System.IO.Directory]::CreateDirectory("$dir\営業\A社.xlsx") | Out-Null
        writeListFile "$dir\営業\A社.xlsx\明細.tsv" @("a")
        # 今の利用者から一覧の表示を拒否したフォルダを置くと、列挙が UnauthorizedAccessException になる（テストの一時フォルダだけを変え、最後に戻す）
        $locked = New-Object System.IO.DirectoryInfo "$dir\営業\読めない"
        $locked.Create()
        $user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $deny = New-Object System.Security.AccessControl.FileSystemAccessRule($user, "ListDirectory", "Deny")
        $acl = $locked.GetAccessControl("Access")
        $acl.AddAccessRule($deny)
        $locked.SetAccessControl($acl)
        try {
            getIndexTsvCounts $dir | Should -Be $null
        } finally {
            $acl = $locked.GetAccessControl("Access")
            [void]$acl.RemoveAccessRule($deny)
            $locked.SetAccessControl($acl)
        }
    }
}
