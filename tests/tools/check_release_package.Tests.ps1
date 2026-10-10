# 配布 zip の検査（tools\check_release_package.ps1）のテスト
BeforeAll {
    $rootDir = (Resolve-Path "$PSScriptRoot\..\..").Path
    $check = "$rootDir\tools\check_release_package.ps1"
    $newReleasePackage = "$rootDir\tools\new_release_package.ps1"
    $version = "v9.9.9"
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $outDir = Join-Path $TestDrive "out"
    & $newReleasePackage -Version $version -OutDir $outDir 3>$null | Out-Null
    $goodZip = Join-Path $outDir "tebunko-$version.zip"

    # 検査を実行して、終了コードと出力（通らなかった項目）を返す
    function invokeCheck([string]$zip, [string]$dir = $outDir, [string]$ver = $version) {
        $output = & $check -ZipPath $zip -OutDir $dir -Version $ver 6>&1 | ForEach-Object { "$_" }
        [pscustomobject]@{ Code = $LASTEXITCODE; Text = ($output -join "`n") }
    }

    # 本物の作業ツリーを汚さずに確かめるための clone を作る。
    # コミット前のフックから動くときは、GIT_DIR などがフックを動かしたリポジトリ（このコミット）を指している。
    # 外さずに git clone すると、そちらを向いてしまい、コミット中の索引を壊す（tests\tools\check_release_tag.Tests.ps1 と同じ理由）。
    # clone する間だけ外し、終わったら戻す
    function newIsolatedClone([string]$destName) {
        $saved = @{}
        foreach ($name in @("GIT_DIR", "GIT_WORK_TREE", "GIT_INDEX_FILE", "GIT_PREFIX", "GIT_OBJECT_DIRECTORY", "GIT_COMMON_DIR")) {
            $saved[$name] = [Environment]::GetEnvironmentVariable($name)
            [Environment]::SetEnvironmentVariable($name, $null)
        }
        try {
            $dest = Join-Path $TestDrive $destName
            # --no-hardlinks: $TestDrive が別のドライブのとき、既定のハードリンクは失敗する（CI のランナーで実際に起きた）
            & git clone -q --no-hardlinks --local $rootDir $dest
            return $dest
        } finally {
            foreach ($name in $saved.Keys) { [Environment]::SetEnvironmentVariable($name, $saved[$name]) }
        }
    }

    # 正しい zip をコピーして、指定したエントリーを書き換えた zip を返す。Mutate は byte[]（削除は $null）を返す
    function newBrokenZip([string]$name, [string]$entry, [scriptblock]$mutate, [switch]$add) {
        $path = Join-Path $TestDrive "$name.zip"
        Copy-Item -LiteralPath $goodZip -Destination $path
        $zip = [System.IO.Compression.ZipFile]::Open($path, [System.IO.Compression.ZipArchiveMode]::Update)
        try {
            $existing = $zip.GetEntry($entry)
            $bytes = $null
            if ($existing) {
                $stream = $existing.Open()
                try {
                    $ms = New-Object System.IO.MemoryStream
                    $stream.CopyTo($ms)
                    $bytes = $ms.ToArray()
                } finally { $stream.Dispose() }
                $existing.Delete()
            }
            $new = & $mutate $bytes
            if ($null -ne $new) {
                $e = $zip.CreateEntry($entry)
                $w = $e.Open()
                try { $w.Write([byte[]]$new, 0, $new.Length) } finally { $w.Dispose() }
            }
        } finally {
            $zip.Dispose()
        }
        $path
    }
}

Describe "check_release_package.ps1" -Tag Io {
    It "正しい zip では通り、終了コードが 0 になる" {
        $result = invokeCheck $goodZip
        $result.Text | Should -Match "通りました"
        $result.Code | Should -Be 0
    }

    It "<Name> のとき、通らず、<Expected> を列挙する" -TestCases @(
        @{ Name = "ファイルが 1 つ無い"; Entry = "tebunko/scripts/tebunko/lib.ps1"; Mutate = { $null }; Expected = "zip に無いファイル: tebunko/scripts/tebunko/lib.ps1" }
        @{ Name = "同梱のフォントが無い"; Entry = "tebunko/scripts/shared/fonts/RethinkSans-wght.ttf"; Mutate = { $null }; Expected = "同梱のフォント・ライセンスの文面が zip に無い: scripts/shared/fonts/RethinkSans-wght.ttf" }
        @{ Name = "同梱のフォント（斜体）が無い"; Entry = "tebunko/scripts/shared/fonts/RethinkSans-Italic-wght.ttf"; Mutate = { $null }; Expected = "同梱のフォント・ライセンスの文面が zip に無い: scripts/shared/fonts/RethinkSans-Italic-wght.ttf" }
        @{ Name = "フォントのライセンスの文面（OFL）が無い"; Entry = "tebunko/scripts/shared/fonts/OFL.txt"; Mutate = { $null }; Expected = "同梱のフォント・ライセンスの文面が zip に無い: scripts/shared/fonts/OFL.txt" }
        @{ Name = "アイコンの形のライセンスの文面（Lucide）が無い"; Entry = "tebunko/scripts/shared/fonts/LICENSE-Lucide.txt"; Mutate = { $null }; Expected = "同梱のフォント・ライセンスの文面が zip に無い: scripts/shared/fonts/LICENSE-Lucide.txt" }
        @{ Name = "余分なファイルがある"; Entry = "tebunko/scripts/extra.ps1"; Mutate = { [byte[]][char[]]"# extra" }; Expected = "zip に余分なファイル: tebunko/scripts/extra.ps1" }
        @{ Name = "1 バイト書き換わっている"; Entry = "tebunko/tebunko.bat"; Mutate = { param($b) $c = [byte[]]$b.Clone(); $c[$c.Length - 1] = $c[$c.Length - 1] -bxor 1; $c }; Expected = "SHA256SUMS.txt のハッシュと一致しません: tebunko.bat" }
        @{ Name = "1 バイト書き換わっていて、カタログとも合わない"; Entry = "tebunko/scripts/tebunko/lib.ps1"; Mutate = { param($b) $c = [byte[]]$b.Clone(); $c[$c.Length - 1] = $c[$c.Length - 1] -bxor 1; $c }; Expected = "カタログの検証に失敗しました" }
        @{ Name = "構文エラーがある"; Entry = "tebunko/scripts/tebunko/core/paths.ps1"; Mutate = { param($b) [byte[]]($b + [System.Text.Encoding]::UTF8.GetBytes("`r`nif (`r`n")) }; Expected = "構文エラー: scripts\tebunko\core\paths.ps1" }
        @{ Name = "xaml が XML として読めない"; Entry = "tebunko/scripts/shared/xaml/theme.xaml"; Mutate = { param($b) [byte[]]($b + [System.Text.Encoding]::UTF8.GetBytes("<")) }; Expected = "XML として読めません: scripts\shared\xaml\theme.xaml" }
        @{ Name = "dot-source の先が無い"; Entry = "tebunko/scripts/tebunko/core/settings.ps1"; Mutate = { $null }; Expected = "dot-source の先がありません: scripts\tebunko\lib.ps1 -> scripts\tebunko\core\settings.ps1" }
        @{ Name = "VERSION.txt の SHA が違う"; Entry = "tebunko/VERSION.txt"; Mutate = { [byte[]]((New-Object System.Text.UTF8Encoding($true)).GetPreamble() + [System.Text.Encoding]::UTF8.GetBytes("v9.9.9`r`n" + ("0" * 40) + "`r`n")) }; Expected = "VERSION.txt がタグ名とコミットの SHA になっていません" }
    ) {
        $zip = newBrokenZip "broken" $Entry $Mutate
        $result = invokeCheck $zip
        $result.Code | Should -Be 1
        $result.Text | Should -BeLike "*$Expected*"
    }

    It "タグ名が VERSION.txt と違うと通らない" {
        $result = invokeCheck $goodZip $outDir "v9.9.8"
        $result.Code | Should -Be 1
        $result.Text | Should -BeLike "*VERSION.txt がタグ名とコミットの SHA になっていません*"
    }

    It "通らなかった項目を、最初の 1 つで止めずにすべて列挙し、終了コードが 1 になる" {
        $zip = newBrokenZip "multi" "tebunko/tebunko.bat" { param($b) [byte[]]($b + 65) }
        $result = invokeCheck $zip
        $result.Code | Should -Be 1
        $result.Text | Should -BeLike "*SHA256SUMS.txt のハッシュと一致しません: tebunko.bat*"
        $result.Text | Should -BeLike "*部品表のハッシュと一致しません: tebunko.bat*"
        $result.Text | Should -BeLike "*カタログの検証に失敗しました*"
    }

    It "ハッシュ一覧に単一 .ps1 版の行があっても、無くても通る（行の追記が検査の前でも後でもよい）" {
        $withOut = Join-Path $TestDrive "withSingle"
        Copy-Item -LiteralPath $outDir -Destination $withOut -Recurse
        $single = Join-Path $withOut "tebunko-$version.ps1"
        [System.IO.File]::WriteAllText($single, "# single", (New-Object System.Text.UTF8Encoding($true)))
        $hash = (Get-FileHash -LiteralPath $single -Algorithm SHA256).Hash
        Add-Content -LiteralPath (Join-Path $withOut "SHA256SUMS.txt") -Value "$hash  tebunko-$version.ps1" -Encoding UTF8
        (invokeCheck $goodZip $withOut).Code | Should -Be 0
    }

    It "ハッシュ一覧の単一 .ps1 版の行が、ファイルと違うと通らない" {
        $badOut = Join-Path $TestDrive "badSingle"
        Copy-Item -LiteralPath $outDir -Destination $badOut -Recurse
        [System.IO.File]::WriteAllText((Join-Path $badOut "tebunko-$version.ps1"), "# single", (New-Object System.Text.UTF8Encoding($true)))
        Add-Content -LiteralPath (Join-Path $badOut "SHA256SUMS.txt") -Value ("0" * 64 + "  tebunko-$version.ps1") -Encoding UTF8
        $result = invokeCheck $goodZip $badOut
        $result.Code | Should -Be 1
        $result.Text | Should -BeLike "*SHA256SUMS.txt のハッシュと一致しません: tebunko-$version.ps1*"
    }

    It "ハッシュ一覧が無いと通らない" {
        $emptyOut = Join-Path $TestDrive "noSums"
        Copy-Item -LiteralPath $outDir -Destination $emptyOut -Recurse
        Remove-Item -LiteralPath (Join-Path $emptyOut "SHA256SUMS.txt")
        $result = invokeCheck $goodZip $emptyOut
        $result.Code | Should -Be 1
        $result.Text | Should -BeLike "*SHA256SUMS.txt がありません*"
    }
}

Describe "new_release_files.ps1 のカタログ" -Tag Io {
    It "追跡していないファイルが scripts\ にあっても、カタログは zip の中身と一致する（検査が通る）" {
        # 本物の作業ツリーを汚さないよう、clone した先に追跡していないファイルを置く
        $cloneDir = newIsolatedClone "clone-untracked"
        Set-Content -LiteralPath (Join-Path $cloneDir "scripts\untracked_for_test.ps1") -Value "# untracked"
        $dir = Join-Path $TestDrive "untracked"
        & "$cloneDir\tools\new_release_package.ps1" -Version $version -OutDir $dir 3>$null | Out-Null
        $result = invokeCheck (Join-Path $dir "tebunko-$version.zip") $dir
        $result.Code | Should -Be 0
    }
}

Describe "new_release_package.ps1 の作業ツリーの検査" -Tag Io {
    It "追跡しているファイルに変更があると警告し、変更が無ければ警告しない" {
        $dir = Join-Path $TestDrive "warn"
        $warnings = @()
        & $newReleasePackage -Version $version -OutDir $dir -WarningVariable warnings 3>$null | Out-Null
        $tracked = & git -C $rootDir status --porcelain --untracked-files=no -- scripts tebunko.bat README.md LICENSE
        @($warnings).Count | Should -Be $(if ($tracked) { 1 } else { 0 })
    }

    It "追跡しているファイルに変更があると、変えたファイルの名前を含む警告を 1 件出す" {
        # 本物の作業ツリーを汚さないよう、clone した先で書き換える
        $cloneDir = newIsolatedClone "clone-warn"
        Add-Content -LiteralPath (Join-Path $cloneDir "tebunko.bat") -Value "rem test"
        $dir = Join-Path $TestDrive "warn-changed"
        $warnings = @()
        & "$cloneDir\tools\new_release_package.ps1" -Version $version -OutDir $dir -WarningVariable warnings 3>$null | Out-Null
        @($warnings).Count | Should -Be 1
        "$($warnings[0])" | Should -BeLike "*tebunko.bat*"
    }
}
