# ワークスペースの中の置き場の名前（content_index・ingest_status.tsv・store）を、コードに直接書かないことを確かめる。
# 置き場の名前は Workspace クラス（scripts/tebunko/core/workspace.ps1）が 1 か所で持ち、ほかは Workspace のプロパティ
# （IndexDir・StatusFile・StoreDir・FormatFile・OldIndexDir・OldStatusFile）から取る。名前を直書きすると、置き場を替えたときに
# 直し忘れる所が出る（前の版の置き場と今の置き場を取り違えないため）。
# コメントは対象にしない。名前の一部になっているだけのもの（content_index_format.ps1 のファイル名、content_index.xlsx.001.tsv の
# ファイル名の型）は当たらない。
# 直書きしてよい所は $allowed に、ファイルと理由を対にして書く。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"

    # 名前ごとの照合。store は区切りの記号（\ / 引用符）に挟まれているときだけ当てる（restore・storeName などの単語に当てない）
    ${namePatterns} = [ordered]@{
        "content_index"     = 'content_index(?![_.A-Za-z0-9])'
        "ingest_status.tsv" = 'ingest_status\.tsv'
        "store"             = '(^|[\\/"''`])store($|[\\/"''`])'
    }

    # 直書きしてよい所: ファイル（リポジトリの相対パス）→ 理由と、当てる名前（空なら全部）
    ${allowed} = @{
        "scripts/tebunko/core/workspace.ps1"            = "Workspace クラスが置き場の名前を持つ唯一の所"
        "scripts/tebunko/index/index_archive_rules.ps1" = "zip の中の名前（書き出した zip の形式。置き場の名前とは別）"
        "scripts/tebunko/index/index_archive.ps1"       = "zip の中の名前（書き出した zip の形式。置き場の名前とは別）"
    }

    function findDirectWorkspaceNames {
        # ファイル path のコード（コメント以外）に直書きされた置き場の名前を「<相対パス>:<行> <名前>」の並びで返す。
        # exempt: 行の文字列がこの正規表現に当たる行は数えない（定数の定義 1 行だけを許すとき）
        param (
            [string]$path,
            [string]$relative,
            [string]$exempt = ""
        )

        $tokens = $null
        $errors = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
        $lines = [System.IO.File]::ReadAllLines($path, [System.Text.Encoding]::UTF8)
        $found = New-Object System.Collections.Generic.List[string]
        foreach ($token in $tokens) {
            if ($token.Kind -eq [System.Management.Automation.Language.TokenKind]::Comment) { continue }
            $text = $token.Text
            foreach ($name in ${namePatterns}.Keys) {
                if ($text -notmatch ${namePatterns}[$name]) { continue }
                $lineNumber = $token.Extent.StartLineNumber
                if ($exempt -and $lines[$lineNumber - 1] -match $exempt) { continue }
                $found.Add("${relative}:${lineNumber} ${name}")
            }
        }
        return $found.ToArray()
    }

    function getScanFiles {
        # scripts/・tools/ の .ps1（リポジトリの相対パスで返す）
        $root = (Resolve-Path "$PSScriptRoot\..\..").ProviderPath
        $files = foreach ($scanRoot in @("scripts", "tools")) {
            Get-ChildItem -LiteralPath (Join-Path $root $scanRoot) -Recurse -File -Filter *.ps1
        }
        return @($files | ForEach-Object { @{ Path = $_.FullName; Relative = $_.FullName.Substring($root.Length + 1).Replace("\", "/") } })
    }
}

Describe "ワークスペースの置き場の名前の直書き" -Tag Meta {
    It "scripts/・tools/ のコードに、置き場の名前を直書きしていない" {
        $found = New-Object System.Collections.Generic.List[string]
        foreach ($file in @(getScanFiles)) {
            if (${allowed}.ContainsKey($file.Relative)) { continue }
            # 本文インデックスのファイル名の先頭の語を定める定数の 1 行だけは直書きしてよい
            $exempt = $(if ($file.Relative -eq "scripts/tebunko/index/content_index_format.ps1") { '^\$\{contentIndexFileNamePrefix\}\s*=' } else { "" })
            foreach ($item in @(findDirectWorkspaceNames $file.Path $file.Relative $exempt)) { $found.Add($item) }
        }
        ($found -join "`n") | Should -Be ""
    }

    It "直書きを見つけられる（検査が効いている）" -TestCases @(
        @{ Code = '$dir = "$work\content_index"'; Name = "content_index" }
        @{ Code = '$file = Join-Path $work "ingest_status.tsv"'; Name = "ingest_status.tsv" }
        @{ Code = '$dir = "$work\store\content"'; Name = "store" }
        @{ Code = '$dir = Join-Path $work "store"'; Name = "store" }
    ) {
        $path = Join-Path $TestDrive "sample.ps1"
        [System.IO.File]::WriteAllText($path, $Code, (New-Object System.Text.UTF8Encoding($true)))
        $found = @(findDirectWorkspaceNames $path "sample.ps1")
        $found | Should -Contain "sample.ps1:1 ${Name}"
    }

    It "コメント・名前の一部は直書きに数えない" -TestCases @(
        @{ Code = '# content_index\ に置く' }
        @{ Code = '$file = "content_index.xlsx.001.tsv"' }
        @{ Code = '. "$PSScriptRoot\index\content_index_format.ps1"' }
        @{ Code = '$word = "restore"' }
    ) {
        $path = Join-Path $TestDrive "sample.ps1"
        [System.IO.File]::WriteAllText($path, $Code, (New-Object System.Text.UTF8Encoding($true)))
        @(findDirectWorkspaceNames $path "sample.ps1").Count | Should -Be 0
    }

    It "直書きを許しているファイルが、実際にある" {
        $relatives = @(getScanFiles | ForEach-Object { $_.Relative })
        foreach ($relative in ${allowed}.Keys) {
            $relatives | Should -Contain $relative
        }
    }
}
