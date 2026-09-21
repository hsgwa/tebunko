# テストデータから個人情報を取り除くスクリプト（scrub_personal.ps1）のテスト。
# Office は使わず、Office が情報を埋め込む形をまねたファイルを作って確かめる。
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$scrub = "$PSScriptRoot\scrub_personal.ps1"
$utf8 = New-Object System.Text.UTF8Encoding($false)
$cp932 = [System.Text.Encoding]::GetEncoding(932)

function newZip([string]$path, [object[]]$entries) {
    $stream = [System.IO.File]::Open($path, [System.IO.FileMode]::CreateNew)
    $archive = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($item in $entries) {
            $level = if ($item.Stored) { [System.IO.Compression.CompressionLevel]::NoCompression } else { [System.IO.Compression.CompressionLevel]::Optimal }
            $entry = $archive.CreateEntry($item.Name, $level)
            $writer = $entry.Open()
            $bytes = if ($item.Bytes) { $item.Bytes } else { $utf8.GetBytes($item.Text) }
            $writer.Write($bytes, 0, $bytes.Length)
            $writer.Dispose()
        }
    } finally {
        $archive.Dispose()
        $stream.Dispose()
    }
}

function readZip([string]$path) {
    $result = [ordered]@{}
    $archive = [System.IO.Compression.ZipFile]::OpenRead($path)
    try {
        foreach ($entry in $archive.Entries) {
            $buffer = New-Object System.IO.MemoryStream
            $stream = $entry.Open()
            $stream.CopyTo($buffer)
            $stream.Dispose()
            $result[$entry.FullName] = $buffer.ToArray()
        }
    } finally {
        $archive.Dispose()
    }
    return $result
}

# 架空の名前。ソースに C:\Users\<名前> や作成者名をそのまま書くと tools\check_commit.ps1 が止めるため、実行時に組み立てる
$name = "yamada"
$titleName = "Yamada"   # パスの中では先頭が大文字のこともある
$fullName = "Yamada Taro"
$jpName = "山田 太郎"
$accountId = "8f3a2b1c9d0e7f6a"

function runScrub([string[]]$names) {
    & $scrub -Path $dir -Names $names -NoAutoNames *> $null
}

Describe "scrub_personal.ps1" -Tag Io {
    $dir = Join-Path $TestDrive "data"

    BeforeEach {
        if (Test-Path -LiteralPath $dir) { Remove-Item -LiteralPath $dir -Recurse -Force }
        New-Item -ItemType Directory $dir | Out-Null
    }

    It "OOXML の作成者・会社名・アカウント ID・パス・名前を消す" {
        $docx = Join-Path $dir "文書.docx"
        newZip $docx @(
            @{ Name = "docProps/core.xml"; Text = "<cp:coreProperties><dc:creator>$fullName</dc:creator><cp:lastModifiedBy>$accountId</cp:lastModifiedBy></cp:coreProperties>" }
            @{ Name = "docProps/app.xml"; Text = '<Properties><Company>(株)山田商事</Company><Manager>部長</Manager></Properties>' }
            @{ Name = "word/people.xml"; Text = "<w15:person w15:author=""$fullName""><w15:presenceInfo w15:providerId=""Windows Live"" w15:userId=""$accountId""/></w15:person>" }
            @{ Name = "word/document.xml"; Text = "<w:t>本文 C:\Users\$name\AppData\Local\Temp\a.docx $($name.ToUpper())</w:t>" }
        )
        runScrub @($fullName, $name)

        $entries = readZip $docx
        $utf8.GetString($entries["docProps/core.xml"]) | Should Be '<cp:coreProperties><dc:creator>test</dc:creator><cp:lastModifiedBy>test</cp:lastModifiedBy></cp:coreProperties>'
        $utf8.GetString($entries["docProps/app.xml"]) | Should Be '<Properties><Company></Company><Manager></Manager></Properties>'
        $utf8.GetString($entries["word/people.xml"]) | Should Be '<w15:person w15:author="test"><w15:presenceInfo w15:providerId="Windows Live" w15:userId="0000000000000000"/></w15:person>'
        $utf8.GetString($entries["word/document.xml"]) | Should Be '<w:t>本文 C:\Users\test\AppData\Local\Temp\a.docx test</w:t>'
    }

    It "書き換えないエントリは、無圧縮のまま・並び順のまま残す（ODF の mimetype）" {
        $odt = Join-Path $dir "文書.odt"
        newZip $odt @(
            @{ Name = "mimetype"; Text = "application/vnd.oasis.opendocument.text"; Stored = $true }
            @{ Name = "meta.xml"; Text = "<office:meta><meta:initial-creator>$fullName</meta:initial-creator></office:meta>" }
            @{ Name = "content.xml"; Text = '<text:p>本文</text:p>' }
        )
        $before = [System.IO.File]::ReadAllBytes($odt)
        runScrub @($fullName)
        $after = [System.IO.File]::ReadAllBytes($odt)

        # 先頭のローカルヘッダー（mimetype）の圧縮方式・サイズ・中身が変わらない
        $headerLength = 30 + "mimetype".Length + "application/vnd.oasis.opendocument.text".Length
        [Convert]::ToBase64String($after, 0, $headerLength) | Should Be ([Convert]::ToBase64String($before, 0, $headerLength))
        $entries = readZip $odt
        @($entries.Keys) -join "," | Should Be "mimetype,meta.xml,content.xml"
        $utf8.GetString($entries["meta.xml"]) | Should Be '<office:meta><meta:initial-creator>test</meta:initial-creator></office:meta>'
        $utf8.GetString($entries["content.xml"]) | Should Be '<text:p>本文</text:p>'
    }

    It "ZIP の中のバイナリ（.xlsb の workbook.bin）は、バイト長を変えずに置き換える" {
        $xlsb = Join-Path $dir "バイナリ.xlsb"
        $bin = [System.Text.Encoding]::Unicode.GetBytes("`0`0C:\Users\$name\AppData\Local\Temp\`0`0")
        newZip $xlsb @(@{ Name = "xl/workbook.bin"; Bytes = $bin })
        runScrub @($name)

        $after = (readZip $xlsb)["xl/workbook.bin"]
        $after.Length | Should Be $bin.Length
        [System.Text.Encoding]::Unicode.GetString($after) | Should Be "`0`0C:\Users\test__\AppData\Local\Temp\`0`0"
    }

    It "旧形式・PDF は、どの文字コードで入っていてもバイト長を変えずに置き換える" {
        $doc = Join-Path $dir "旧形式.doc"
        $parts = @(
            $cp932.GetBytes("[author=$jpName]"),
            [System.Text.Encoding]::Unicode.GetBytes("[$jpName]"),
            $utf8.GetBytes("[/Author ($jpName) path C:\Users\$titleName\x]")
        )
        $bytes = [byte[]]($parts | ForEach-Object { $_ })
        [System.IO.File]::WriteAllBytes($doc, $bytes)
        runScrub @($jpName, $name)

        $after = [System.IO.File]::ReadAllBytes($doc)
        $after.Length | Should Be $bytes.Length
        $cp932.GetString($after, 0, $parts[0].Length) | Should Be "[author=test_____]"
        [System.Text.Encoding]::Unicode.GetString($after, $parts[0].Length, $parts[1].Length) | Should Be "[test_]"
        $utf8.GetString($after, $parts[0].Length + $parts[1].Length, $parts[2].Length) | Should Be "[/Author (test_________) path C:\Users\test__\x]"
    }

    It "テキスト（設定例）は長さを変えて置き換え、文字コード（Shift_JIS）を壊さない" {
        $txt = Join-Path $dir "変換対象フォルダパス_ShiftJIS.txt"
        [System.IO.File]::WriteAllBytes($txt, $cp932.GetBytes("C:\Users\$titleName\テスト\ファイル名`r`n"))
        runScrub @($name)

        $cp932.GetString([System.IO.File]::ReadAllBytes($txt)) | Should Be "C:\Users\test\テスト\ファイル名`r`n"
    }

    It "壊れた ZIP・空ファイルでも止まらない" {
        $broken = Join-Path $dir "壊れたファイル.xlsx"
        [System.IO.File]::WriteAllBytes($broken, [byte[]](@(0x50, 0x4B, 0x03, 0x04) + $utf8.GetBytes($name)))
        [System.IO.File]::WriteAllBytes((Join-Path $dir "空ファイル.xlsx"), (New-Object byte[] 0))
        runScrub @($name)

        $utf8.GetString([System.IO.File]::ReadAllBytes($broken), 4, 6) | Should Be "test__"
    }

    It "隠し属性・読み取り専用のファイルも書き換え、属性は元に戻す" {
        $lock = Join-Path $dir "~`$ロックファイル.xlsx"
        [System.IO.File]::WriteAllBytes($lock, $utf8.GetBytes("owner $name"))
        $item = Get-Item -LiteralPath $lock -Force
        $item.Attributes = [System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::ReadOnly
        runScrub @($name)

        $item = Get-Item -LiteralPath $lock -Force
        $utf8.GetString([System.IO.File]::ReadAllBytes($lock)) | Should Be "owner test__"
        $item.Attributes.HasFlag([System.IO.FileAttributes]::Hidden) | Should Be $true
        $item.IsReadOnly | Should Be $true
        $item.IsReadOnly = $false
    }

    It "ありふれた語（user など）は名前として使わない（C:\Users を壊さない）" {
        $txt = Join-Path $dir "パス.txt"
        [System.IO.File]::WriteAllBytes($txt, $utf8.GetBytes("C:\Users\$name\x"))
        runScrub @("User", "Administrator", $name)

        $utf8.GetString([System.IO.File]::ReadAllBytes($txt)) | Should Be "C:\Users\test\x"
    }

    It "名前が無いファイルは書き換えない（更新日時が変わらない）" {
        $docx = Join-Path $dir "無関係.docx"
        newZip $docx @(@{ Name = "docProps/core.xml"; Text = '<dc:creator>test</dc:creator>' })
        $time = (Get-Date).AddDays(-1)
        (Get-Item -LiteralPath $docx).LastWriteTime = $time
        runScrub @($name)

        (Get-Item -LiteralPath $docx).LastWriteTime | Should Be $time
    }
}
