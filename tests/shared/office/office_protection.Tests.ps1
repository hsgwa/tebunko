# 暗号化されたファイルの見分け・状態層（shared\office\office_protection.ps1）のテスト。
# CFB（複合ドキュメント形式）は、テストの中で最小限のものを組み立てる（本物はテストデータの既存のファイルを使う。
# 組み立てる関数 newCompoundFile は tests\helpers\cfb.ps1 にあり、load.ps1 が読み込む）。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "readFileHead" -Tag Io {
    BeforeAll {
        $dir = Join-Path $TestDrive "head"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
    }

    It "ファイルの先頭を読む（既定は4096バイト。ファイルが短ければ全部）" {
        $path = Join-Path $dir "short.bin"
        [System.IO.File]::WriteAllBytes($path, [byte[]](1, 2, 3))
        (@(readFileHead $path) -join ",") | Should -Be "1,2,3"
    }

    It "0バイトのファイルは空の配列" {
        $path = Join-Path $dir "empty.bin"
        [System.IO.File]::WriteAllBytes($path, [byte[]]@())
        @(readFileHead $path).Count | Should -Be 0
    }

    It "無いファイルは空の配列（例外にしない）" {
        @(readFileHead (Join-Path $dir "no-such-file.bin")).Count | Should -Be 0
    }
}

Describe "readCompoundEntryNames" -Tag Io {
    BeforeAll {
        $dir = Join-Path $TestDrive "cfb"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
    }

    It "CFBでないファイルは `$null" {
        $path = Join-Path $dir "notcfb.bin"
        [System.IO.File]::WriteAllBytes($path, [byte[]](0x50, 0x4B, 0x03, 0x04))
        readCompoundEntryNames $path | Should -Be $null
    }

    It "512バイト未満のファイルは `$null" {
        $path = Join-Path $dir "toosmall.bin"
        [System.IO.File]::WriteAllBytes($path, [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1))
        readCompoundEntryNames $path | Should -Be $null
    }

    It "1セクターのディレクトリのエントリ名を読む" {
        $path = Join-Path $dir "one-sector.cfb"
        newCompoundFile $path @(@("WordDocument", "1Table"))
        @(readCompoundEntryNames $path) | Should -Be @("Root Entry", "WordDocument", "1Table")
    }

    It "複数セクターに分かれたディレクトリの鎖をたどる" {
        $path = Join-Path $dir "two-sectors.cfb"
        newCompoundFile $path @(@("a", "b", "c"), @("d", "e"))
        @(readCompoundEntryNames $path) | Should -Be @("Root Entry", "a", "b", "c", "d", "e")
    }

    It "FATの鎖が輪になっていても、例外を出さず速く戻り、読めた分だけ返す" {
        $path = Join-Path $dir "loop.cfb"
        newCompoundFile $path @(@("a"), @("b")) @{ 2 = 1 }   # セクター2 → 1 → 2 → ... の輪
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $names = @(readCompoundEntryNames $path)
        $sw.Stop()
        $sw.Elapsed.TotalSeconds | Should -BeLessThan 10
        $names | Should -Be @("Root Entry", "a", "b")
    }

    It "FATが範囲外のセクターを指していても、例外を出さず、そこまでに読めた分だけ返す" {
        $path = Join-Path $dir "outofrange.cfb"
        newCompoundFile $path @(@("a")) @{ 1 = 999999 }
        $names = @(readCompoundEntryNames $path)
        $names | Should -Be @("Root Entry", "a")
    }

    It "ディレクトリの鎖が途中で切れていても、読めた分だけ返す" {
        # セクター1（1つ目のディレクトリセクター）の次を、どこも指していない値にする
        $path = Join-Path $dir "truncated.cfb"
        newCompoundFile $path @(@("a", "b"), @("c", "d")) @{ 1 = 5 }
        @(readCompoundEntryNames $path) | Should -Be @("Root Entry", "a", "b")
    }

    It "DIFATセクターの先にあるFATから、次のセクターの名前を読める" {
        $path = Join-Path $dir "difat.cfb"
        newCompoundFileWithDifat $path @(@("a", "b", "c"), @("d", "e"))
        @(readCompoundEntryNames $path) | Should -Be @("Root Entry", "a", "b", "c", "d", "e")
    }

    It "DIFATの鎖が輪になっていても、例外を出さず速く戻る" {
        $path = Join-Path $dir "difat-loop.cfb"
        newCompoundFileWithDifat $path @(@("a", "b", "c"), @("d", "e")) -loopDifat
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $names = @(readCompoundEntryNames $path)
        $sw.Stop()
        $sw.Elapsed.TotalSeconds | Should -BeLessThan 10
        # DIFATが輪でも、FATセクターの並び自体（セクター0）は読めているため、ディレクトリの鎖はふつうにたどれる
        $names | Should -Be @("Root Entry", "a", "b", "c", "d", "e")
    }
}

Describe "getOfficeFileProtection" -Tag Io {
    BeforeAll {
        $dir = Join-Path $TestDrive "protection"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
    }

    It "新形式の権限保護（IRM・秘密度ラベル）は Rights" {
        $path = Join-Path $dir "irm-new.docx"
        newCompoundFile $path @(@(([char]6 + "DataSpaces"), "DRMEncryptedDataSpace", "DRMEncryptedTransform"))
        getOfficeFileProtection $path | Should -Be "Rights"
    }

    It "旧形式の権限保護（IRM）は Rights" {
        $path = Join-Path $dir "irm-legacy.doc"
        newCompoundFile $path @(@(([char]9 + "DRMContent"), ([char]9 + "DRMDataSpace")))
        getOfficeFileProtection $path | Should -Be "Rights"
    }

    It "パスワード付き（新形式）は Password" {
        $path = Join-Path $dir "password-new.xlsx"
        newCompoundFile $path @(@(([char]6 + "DataSpaces"), "StrongEncryptionDataSpace", "EncryptionInfo", "EncryptedPackage"))
        getOfficeFileProtection $path | Should -Be "Password"
    }

    It "権限保護と分からないCFB（旧形式）は Legacy" {
        $path = Join-Path $dir "legacy.doc"
        newCompoundFile $path @(@("WordDocument", "1Table"))
        getOfficeFileProtection $path | Should -Be "Legacy"
    }

    It "壊れたCFB（ディレクトリの鎖が途中で切れる）は、読めた名前だけで判定し Legacy になる" {
        $path = Join-Path $dir "broken.doc"
        newCompoundFile $path @(@("a", "b"), @("c", "d")) @{ 1 = 5 }
        getOfficeFileProtection $path | Should -Be "Legacy"
    }

    It "ランダムなバイト列の .docx は Unknown（形式の分からないバイナリ）" {
        $path = Join-Path $dir "random.docx"
        $random = New-Object byte[] 2048
        (New-Object System.Random(1)).NextBytes($random)
        $random[0] = 0x12  # ZIP・CFBの先頭バイトと重ならないようにする
        [System.IO.File]::WriteAllBytes($path, $random)
        getOfficeFileProtection $path | Should -Be "Unknown"
    }

    # 既存のテストデータで、今読めているものを壊さない（実際のファイルで確かめる）
    It "<path>" -TestCases @(
        @{ path = "Word\異常系\読み取りパスワード付き.docx"; kind = "Password" }
        @{ path = "PowerPoint\異常系\読み取りパスワード付き.pptx"; kind = "Password" }
        @{ path = "Excel\異常系\読み取りパスワード付き.xlsx"; kind = "Password" }
        @{ path = "Word\異常系\読み取りパスワード付き旧形式.doc"; kind = "Legacy" }
        @{ path = "PowerPoint\異常系\読み取りパスワード付き旧形式.ppt"; kind = "Legacy" }
        @{ path = "Word\形式\旧形式.doc"; kind = "Legacy" }
        @{ path = "Excel\異常系\中身はHTML.xls"; kind = "Text" }
        @{ path = "Word\異常系\中身はHTML.doc"; kind = "Text" }
        @{ path = "Word\異常系\中身はRTF.doc"; kind = "Text" }
        @{ path = "Word\形式\リッチテキスト.rtf"; kind = "Text" }
    ) {
        param ($path, $kind)
        getOfficeFileProtection (Join-Path ${testDataDir} "office\$path") | Should -Be $kind
    }
}
