# クロール対象フォルダ・ワークスペースの名前に記号が入っていても、取り込みと検索が通ることのテスト（Office は使わない）。
# 名前に [ ] や ' が入っていると、Pester の TestDrive の後片付けが失敗することがあるため、%TEMP% の下に専用のフォルダを作って .NET で消す。
# ワークスペースの名前の [ ] は、中間ファイルを置けないため取り込まない決まり（indexer.Tests.ps1）なので、ここでは入れない。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "$PSScriptRoot\..\..\helpers\indexer.ps1"

    $base = Join-Path ([System.IO.Path]::GetTempPath()) "tebunko-pathchars-$([guid]::NewGuid().ToString('N'))"
    [System.IO.Directory]::CreateDirectory($base) | Out-Null
}

AfterAll {
    if ([System.IO.Directory]::Exists($base)) { [System.IO.Directory]::Delete($base, $true) }
}

Describe "indexer.ps1（名前に記号があるフォルダ）" -Tag Io {
    It "<label>: 取り込んだ語で検索すると、正しい名前・行番号でヒットする" -TestCases @(
        @{ label = "一重引用符"; source = "O'Neil"; work = "ws's"; file = "メモ's.txt" }
        @{ label = "角括弧"; source = "[1]資料"; work = "ws&1"; file = "[確定]メモ.txt" }
        @{ label = "引用符・&・角括弧の組み合わせ"; source = "O'aa&[]fds"; work = "O'aa&fds"; file = "O'aa&[]メモ.txt" }
        @{ label = "バッククォートと角括弧"; source = '`[a]`'; work = 'ws`'; file = 'メモ`[1].txt' }
    ) {
        param ($label, $source, $work, $file)
        $dir = Join-Path $base $source
        [System.IO.Directory]::CreateDirectory("$dir\営業") | Out-Null
        [System.IO.File]::WriteAllText("$dir\営業\$file", "1行目`r`n検索語がある2行目`r`n", ${utf8Bom})
        $root = Join-Path $base "tool-$work"
        $workspace = Join-Path $root $work
        [System.IO.Directory]::CreateDirectory($workspace) | Out-Null
        $settings = newSettings
        $settings.targetFolders = @(@{ name = ""; path = $dir; enabled = $true })
        $settings.workspaceFolder = $workspace
        writeSettings $settings "$root\setting.config"

        runIndexer $root | Should -Be 0

        $status = readStatusFile "$workspace\store\ingest_status.tsv"
        $key = "$source\営業\$file"
        $status.Rows[$key].状態 | Should -Be ${stateDone}
        $contentIndexFiles = (getContentIndexFiles @("$workspace\store\content")).ContentIndexFiles
        $hits = @((expandHitRuns (searchContentIndex "検索語" $contentIndexFiles $true).Runs))
        $hits.Count | Should -Be 1
        $hits[0].Book | Should -Be $file
        $hits[0].LineNumber | Should -Be 2
        $hits[0].Line | Should -Be "検索語がある2行目"
    }
}