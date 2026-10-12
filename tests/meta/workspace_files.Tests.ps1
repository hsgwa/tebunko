# ワークスペースに tebunko が作るファイル・フォルダ（Workspace.Entries()）が、設計書 docs/design/structure/io-files.md の
# 「増え方と片付け」のワークスペースの表にすべて載っていることを確かめる。
# 新しく足した人が、増え方と上限・片付けを決めずに入れられないようにするため。
# 名前は完全一致で比べる（index が content_index に当たらないように）。ワークスペースの外の置き場の表は対象外。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"

    function getDocumentedWorkspaceNames {
        # 「## 増え方と片付け」の「### ワークスペース」の表の 1 列目（バッククォートで囲んだ名前）を返す
        param (
            [string[]]$lines
        )

        $names = New-Object System.Collections.Generic.List[string]
        $inSection = $false
        $inTable = $false
        foreach ($line in $lines) {
            if ($line -match '^## ') {
                $inSection = ($line -match '^## 増え方と片付け\s*$')
                $inTable = $false
                continue
            }
            if (-not $inSection) { continue }
            if ($line -match '^### ') {
                $inTable = ($line -match '^### ワークスペース（')
                continue
            }
            if ($inTable -and $line -match '^\| `([^`]+)` \|') {
                $names.Add($Matches[1])
            }
        }
        return $names.ToArray()
    }

    $script:docLines = [System.IO.File]::ReadAllLines("$here\..\docs\design\structure\io-files.md", [System.Text.Encoding]::UTF8)
}

Describe "ワークスペースのファイルの一覧" -Tag Meta {
    It "Workspace.Entries() の名前がすべて、設計書の「増え方と片付け」の表にある" {
        $documented = @(getDocumentedWorkspaceNames $script:docLines)
        $documented.Count | Should -BeGreaterThan 0
        $missing = @((newWorkspace "C:\ws").Entries() | ForEach-Object { [System.IO.Path]::GetFileName($_) } | Where-Object { $documented -cnotcontains $_ })
        ($missing -join ", ") | Should -Be ""
    }

    It "表から 1 行消すと見つかる（検査が効いている）" {
        $lines = @($script:docLines | Where-Object { $_ -notmatch '^\| `tmp` \|' })
        $documented = @(getDocumentedWorkspaceNames $lines)
        $missing = @((newWorkspace "C:\ws").Entries() | ForEach-Object { [System.IO.Path]::GetFileName($_) } | Where-Object { $documented -cnotcontains $_ })
        $missing | Should -Contain "tmp"
    }
}
