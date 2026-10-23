# 取り込み対象の元のファイルが、インデックスの作成・更新・エクスポート・削除の前後で変わらないことの通しのテスト。
# 元のファイルを書き換えない・消さない決まり（docs/safety/file-access.md「取り込み対象のファイルは書き換えない」）を、
# 動かして確かめる。書く API の許可リスト（tests/meta/source_write.Tests.ps1）は書き方を見るのに対し、こちらは結果を見る。
# 使い捨てのクロール対象フォルダに、テキスト・PDF・Word・PowerPoint・読み取り専用属性・パスワード付き・
# 外のフォルダを指すジャンクションを置き、操作の前後で次を比べる（元のフォルダと、ジャンクションの先の外のフォルダの両方）。
#   ファイルの数・種類・ハッシュ・属性・作成日時（UTC）・更新日時（UTC）（フォルダは更新日時を除く。getTreeSnapshot）
# 最終アクセス日時は OS が読むたびに記録しうるので比べない（tebunko の書き込みではない）。
# Io は Office が無くても流れる（Word・PowerPoint の新形式はファイルを直接読む）。Excel と旧形式（.ppt）は COM が要るので Office。
# Office は CI のランナーに無く既定で外れるため、リリース前に `.\tests\run.ps1 -All` で流す（docs/design/testing/ci.md）。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "$PSScriptRoot\..\..\helpers\indexer.ps1"

    function getTreeSnapshot {
        # root の下の全項目（ファイル・フォルダ・リンク）を、リンクの中へ降りずに、1 項目 1 行の文字列の配列にする。
        # 行: 相対パス | 種類 | 属性 | 作成日時 | 更新日時 | ハッシュ（ファイルだけ）
        param ([string]$root)

        $rows = New-Object System.Collections.Generic.List[string]
        $stack = New-Object System.Collections.Generic.Stack[string]
        $stack.Push($root)
        while ($stack.Count -gt 0) {
            $dir = $stack.Pop()
            foreach ($item in ([System.IO.DirectoryInfo]::new($dir)).GetFileSystemInfos()) {
                $relative = $item.FullName.Substring($root.Length).TrimStart("\")
                $isLink = ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0
                $isDirectory = ($item.Attributes -band [System.IO.FileAttributes]::Directory) -ne 0
                $kind = if ($isLink) { "リンク" } elseif ($isDirectory) { "フォルダ" } else { "ファイル" }
                $hash = ""
                if ($kind -eq "ファイル") {
                    $hash = (Get-FileHash -LiteralPath $item.FullName -Algorithm SHA256).Hash
                } elseif ($kind -eq "フォルダ") {
                    $stack.Push($item.FullName)
                }
                # フォルダの更新日時は比べない（NTFS は、フォルダ自身の更新日時を親の一覧へ遅れて反映するため、何も書かなくても揺れる）。
                # フォルダの中でファイルが増えた・消えたことは、中の項目の行で分かる
                $written = if ($kind -eq "フォルダ") { "-" } else { "$($item.LastWriteTimeUtc.Ticks)" }
                $rows.Add("$relative | $kind | $([int]$item.Attributes) | $($item.CreationTimeUtc.Ticks) | $written | $hash")
            }
        }
        return @($rows | Sort-Object)
    }

    function assertUnchanged {
        # 元のフォルダと外のフォルダが、before と同じであること。違う行を名前で示す
        param ($before, [string]$source, [string]$outside, [string]$step)

        $afterSource = @(getTreeSnapshot $source)
        $afterOutside = @(getTreeSnapshot $outside)
        $diffSource = @(Compare-Object $before.Source $afterSource | ForEach-Object { "$($_.SideIndicator) $($_.InputObject)" })
        $diffOutside = @(Compare-Object $before.Outside $afterOutside | ForEach-Object { "$($_.SideIndicator) $($_.InputObject)" })
        ($diffSource -join "`n") | Should -BeNullOrEmpty -Because "$step のあと、元のフォルダが変わった"
        ($diffOutside -join "`n") | Should -BeNullOrEmpty -Because "$step のあと、ジャンクションの先のフォルダが変わった"
    }

    function newSourceTree {
        # 取り込み対象のフォルダ（source）と、そこからジャンクションで指す外のフォルダ（outside）を作る。
        # 返すもの: @{ Source; Outside; Junction（作れたか） }。files は testdata の中の追加するファイル（相対パス → 置き場所）
        param ([string]$name, [hashtable]$files)

        $source = Join-Path $TestDrive "$name\営業部"
        $outside = Join-Path $TestDrive "$name\外"
        [System.IO.Directory]::CreateDirectory("$source\資料") | Out-Null
        [System.IO.Directory]::CreateDirectory($outside) | Out-Null
        [System.IO.File]::WriteAllText("$source\メモ.txt", "山田さんの連絡先は後で確認する", (New-Object System.Text.UTF8Encoding($true)))
        [System.IO.File]::WriteAllText("$source\読み取り専用.txt", "守る本文の読み取り専用", (New-Object System.Text.UTF8Encoding($true)))
        [System.IO.File]::SetAttributes("$source\読み取り専用.txt", [System.IO.FileAttributes]::ReadOnly)
        [System.IO.File]::WriteAllText("$outside\外のひみつ.txt", "外のフォルダのひみつ", (New-Object System.Text.UTF8Encoding($true)))
        foreach ($relative in $files.Keys) {
            Copy-Item -LiteralPath $files[$relative] -Destination (Join-Path $source $relative)
        }
        $junction = newTestJunction "$source\資料\リンク" $outside
        return @{ Source = $source; Outside = $outside; Junction = $junction }
    }

    function invokeSourceScenario {
        # 作成・更新・エクスポート・削除のそれぞれの前後で、元のフォルダと外のフォルダが変わらないことを確かめる
        param ($tree, [string]$name)

        $root = newRoot
        writeTestSettings $root @(@{ name = "営業"; path = $tree.Source; enabled = $true })
        $settingsPath = "$root\setting.config"
        $workspace = [Workspace]::new("$root\work")
        $before = @{ Source = @(getTreeSnapshot $tree.Source); Outside = @(getTreeSnapshot $tree.Outside) }
        $before.Source.Count | Should -BeGreaterThan 3

        # 作成
        runIndexer $root | Should -Be 0
        $rows = @((readTestStatus $root).Rows.Keys)
        $rows.Count | Should -BeGreaterThan 2 -Because "取り込みが動いていないと、比べる意味がない"
        @($rows | Where-Object { $_ -match "外のひみつ" }).Count | Should -Be 0 -Because "ジャンクションの先は取り込まない"
        assertUnchanged $before $tree.Source $tree.Outside "インデックスの作成"

        # 更新（元のファイルを足す・書き換えるのはテストの側。そのあとの状態を比べる基準にする）
        [System.IO.File]::WriteAllText("$($tree.Source)\追加.txt", "あとから足した本文", (New-Object System.Text.UTF8Encoding($true)))
        [System.IO.File]::WriteAllText("$($tree.Source)\メモ.txt", "山田さんの連絡先を確認した", (New-Object System.Text.UTF8Encoding($true)))
        $before = @{ Source = @(getTreeSnapshot $tree.Source); Outside = @(getTreeSnapshot $tree.Outside) }
        runIndexer $root | Should -Be 0
        @((readTestStatus $root).Rows.Keys | Where-Object { $_ -match "追加\.txt" }).Count | Should -Be 1
        assertUnchanged $before $tree.Source $tree.Outside "インデックスの更新"

        # エクスポート
        $zip = Join-Path $TestDrive "$name.zip"
        $null = exportIndex "営業" $zip $workspace $settingsPath
        [System.IO.File]::Exists($zip) | Should -BeTrue
        assertUnchanged $before $tree.Source $tree.Outside "エクスポート"

        # インデックスの削除
        (newIndexCatalog $workspace.Dir $settingsPath).RemoveIndex("営業")
        [System.IO.Directory]::Exists((Join-Path $workspace.IndexDir "営業")) | Should -BeFalse
        assertUnchanged $before $tree.Source $tree.Outside "インデックスの削除"
    }
}

Describe "元のファイルが変わらない（通し）" -Tag Io {
    It "テキスト・PDF・Word・PowerPoint・読み取り専用・パスワード付き・ジャンクションの置かれたフォルダで、作成・更新・エクスポート・削除の前後が同じ" {
        $tree = newSourceTree "io" @{
            "議事録.docx"       = "${testDataDir}\office\Word\形式\大文字拡張子.DOCX"
            "資料\提案.pptx"     = "${testDataDir}\office\PowerPoint\基本.pptx"
            "仕様書.pdf"         = "${testDataDir}\compat\index\v0.4.0+pdf\source\仕様書.pdf"
            "パスワード付き.docx" = "${testDataDir}\office\Word\異常系\書き込みパスワード付き.docx"
        }
        if (!$tree.Junction) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        invokeSourceScenario $tree "io"
    }

    It "ジャンクションが無いフォルダでも同じ（ジャンクションを作れない環境でも、元が変わらないことは確かめる）" {
        $tree = newSourceTree "io_nolink" @{
            "議事録.docx" = "${testDataDir}\office\Word\形式\大文字拡張子.DOCX"
            "仕様書.pdf"   = "${testDataDir}\compat\index\v0.4.0+pdf\source\仕様書.pdf"
        }
        if ($tree.Junction) {
            [System.IO.Directory]::Delete("$($tree.Source)\資料\リンク")
        }
        invokeSourceScenario $tree "io_nolink"
    }
}

Describe "元のファイルが変わらない（通し・Excel と旧形式）" -Tag Office {
    It "Excel（パスワード付きを含む）・旧形式の PowerPoint を含むフォルダで、作成・更新・エクスポート・削除の前後が同じ" {
        $tree = newSourceTree "office" @{
            "議事録.docx"             = "${testDataDir}\office\Word\形式\大文字拡張子.DOCX"
            "売上.xlsx"               = "${testDataDir}\office\Excel\シート名記号.xlsx"
            "書き込みパスワード付き.xlsx" = "${testDataDir}\office\Excel\保護\書き込みパスワード付き.xlsx"
            "読み取りパスワード付き.xlsx" = "${testDataDir}\office\Excel\異常系\読み取りパスワード付き.xlsx"
            "資料\旧形式.ppt"           = "${testDataDir}\office\PowerPoint\異常系\読み取りパスワード付き旧形式.ppt"
        }
        if (!$tree.Junction) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        invokeSourceScenario $tree "office"
    }
}

Describe "元のファイルの前後の比べ方（自己テスト）" -Tag Io {
    It "<name> を見逃さない" -TestCases @(
        @{ name = "中身の書き換え（日時も戻した場合）"; change = { param ($dir) $f = "$dir\メモ.txt"; $t = [System.IO.File]::GetLastWriteTimeUtc($f); [System.IO.File]::WriteAllText($f, "別の本文"); [System.IO.File]::SetLastWriteTimeUtc($f, $t) } }
        @{ name = "更新日時だけの変更"; change = { param ($dir) [System.IO.File]::SetLastWriteTimeUtc("$dir\メモ.txt", [datetime]"2001-01-01") } }
        @{ name = "作成日時だけの変更"; change = { param ($dir) [System.IO.File]::SetCreationTimeUtc("$dir\メモ.txt", [datetime]"2001-01-01") } }
        @{ name = "属性だけの変更（読み取り専用を外す）"; change = { param ($dir) [System.IO.File]::SetAttributes("$dir\読み取り専用.txt", [System.IO.FileAttributes]::Normal) } }
        @{ name = "ファイルの削除"; change = { param ($dir) [System.IO.File]::SetAttributes("$dir\読み取り専用.txt", [System.IO.FileAttributes]::Normal); [System.IO.File]::Delete("$dir\読み取り専用.txt") } }
        @{ name = "ファイルの追加"; change = { param ($dir) [System.IO.File]::WriteAllText("$dir\資料\増えた.txt", "x") } }
    ) {
        param ($name, $change)
        $tree = newSourceTree "selftest_$([guid]::NewGuid().ToString('N'))" @{}
        $before = @(getTreeSnapshot $tree.Source)
        @(getTreeSnapshot $tree.Source) | Should -Be $before
        & $change $tree.Source
        @(Compare-Object $before @(getTreeSnapshot $tree.Source)).Count | Should -BeGreaterThan 0
    }
}
