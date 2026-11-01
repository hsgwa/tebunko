# 公開してよい内容かを確かめるスクリプト（tools\check_commit.ps1）のテスト。
# 一時の git リポジトリに check_commit.ps1 を写し、ステージした Office ファイルを調べさせる。
# check_commit.ps1 が頼るのは、自分の 1 つ上のフォルダ（調べるリポジトリ）と git だけ
BeforeAll {
    . "$PSScriptRoot\..\helpers\git_repo.ps1"
    Add-Type -AssemblyName System.IO.Compression
    $script:savedGitEnv = clearGitEnv
    $utf8 = New-Object System.Text.UTF8Encoding($false)

    # ZIP のバイト列を作る。entries は名前と中身（文字列かバイト列）の組
    function newZipBytes([object[]]$entries) {
        $stream = New-Object System.IO.MemoryStream
        $archive = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create, $true)
        try {
            foreach ($item in $entries) {
                $bytes = if ($item.Bytes) { $item.Bytes } else { $utf8.GetBytes($item.Text) }
                $writer = $archive.CreateEntry($item.Name).Open()
                $writer.Write($bytes, 0, $bytes.Length)
                $writer.Dispose()
            }
        } finally {
            $archive.Dispose()
        }
        return , $stream.ToArray()
    }

    # 一時のリポジトリに、埋め込みブックを持つ docx を 1 つだけステージして check_commit.ps1 -Staged を流す。終了コードと出力を返す
    function checkEmbedded([byte[]]$inner) {
        $repo = Join-Path $TestDrive "repo"
        if (Test-Path -LiteralPath $repo) { Remove-Item -LiteralPath $repo -Recurse -Force }
        initTestRepo $repo $TestDrive
        New-Item -ItemType Directory (Join-Path $repo "tools") | Out-Null
        Copy-Item -LiteralPath "$PSScriptRoot\..\..\tools\check_commit.ps1" -Destination (Join-Path $repo "tools\check_commit.ps1")
        git -C $repo config user.email "1+test@users.noreply.github.com"
        git -C $repo config user.name "test"
        $docx = newZipBytes @(
            @{ Name = "docProps/core.xml"; Text = "<cp:coreProperties><dc:creator>test</dc:creator></cp:coreProperties>" }
            @{ Name = "word/embeddings/Book.xlsx"; Bytes = $inner }
        )
        [System.IO.File]::WriteAllBytes((Join-Path $repo "グラフ.docx"), $docx)
        git -C $repo add "グラフ.docx"
        $out = & (Join-Path $repo "tools\check_commit.ps1") -Staged *>&1 | Out-String
        return @{ Code = $LASTEXITCODE; Out = $out }
    }

    # 埋め込みブック（xlsx）のバイト列。作成者名を引数で変える。名前は実行時に組み立てる（ソースに書くと自分自身が止まる）
    function newBook([string]$creator) {
        return , (newZipBytes @(
                @{ Name = "docProps/core.xml"; Text = "<cp:coreProperties><dc:creator>$creator</dc:creator></cp:coreProperties>" }
                @{ Name = "xl/workbook.xml"; Text = "<workbook/>" }
            ))
    }
}

AfterAll {
    restoreGitEnv $script:savedGitEnv
}

Describe "check_commit.ps1（ZIP の中の ZIP）" -Tag Io {
    It "ZIP の中の ZIP の作成者名を見つける" {
        # 埋め込みブックの作成者だけに名前がある（外側の core.xml は test）。中の ZIP は圧縮されたバイト列なので、開かないと見えない
        $result = checkEmbedded (newBook ("Yamada" + " Taro"))
        $result.Code | Should -Be 1
        $result.Out | Should -Match ([regex]::Escape("グラフ.docx: word/embeddings/Book.xlsx の docProps/core.xml の 作成者名:"))

        # 作成者が test なら通る
        (checkEmbedded (newBook "test")).Code | Should -Be 0
    }

    It "開けない中の ZIP（壊れたデータ）は、バイト列として調べて止まらない" {
        $broken = [byte[]]@(0x50, 0x4B, 0x03, 0x04, 1, 2, 3, 4, 5, 6, 7, 8)
        (checkEmbedded $broken).Code | Should -Be 0
    }
}
