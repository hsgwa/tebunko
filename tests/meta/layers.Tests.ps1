# フォルダ構成の決まりごとのテスト（文脈と層の分け方を保つ）。
. "$PSScriptRoot\..\helpers\load.ps1"

# ファイルが dot-source している相手を返す（. "$PSScriptRoot\..." の形だけを見る）
function getSourcedFiles {
    param ([string]$path)

    $dir = Split-Path $path -Parent
    $text = [System.IO.File]::ReadAllText($path)
    $result = @()
    foreach ($match in [regex]::Matches($text, '(?m)^\s*\.\s+"\$PSScriptRoot\\([^"]+)"')) {
        $full = Join-Path $dir $match.Groups[1].Value
        if (Test-Path -LiteralPath $full) {
            $result += (Resolve-Path -LiteralPath $full).Path
        }
    }
    return $result
}

Describe "依存の向き" -Tag Meta {
    # shared はどのツールからも使う部品。ツール（scripts\tebunko など）のフォルダを知っていてはいけない。
    # 製品の名前 tebunko は shared でも使う（%LOCALAPPDATA%\tebunko など）ため、ツールのフォルダを指す書き方だけを探す
    It "shared 配下にツールのフォルダが出てこない" {
        $toolNames = @(Get-ChildItem "${scriptsDir}" -Directory | Where-Object { $_.Name -ne "shared" } | ForEach-Object { [regex]::Escape($_.Name) })
        $pattern = "\.\.\\(" + ($toolNames -join "|") + ")\\|scripts[\\/](" + ($toolNames -join "|") + ")\b"
        $found = @(Get-ChildItem "${scriptsDir}\shared" -Recurse -Include *.ps1, *.xaml |
            Select-String -Pattern $pattern |
            ForEach-Object { "$($_.Filename):$($_.LineNumber)" })
        ($found -join ", ") | Should Be ""
    }

    It "ツール同士は互いを読み込まない" {
        $tools = @(Get-ChildItem "${scriptsDir}" -Directory | Where-Object { $_.Name -ne "shared" })
        foreach ($tool in $tools) {
            $others = @($tools | Where-Object { $_.Name -ne $tool.Name } | ForEach-Object { $_.Name })
            if ($others.Count -eq 0) { continue }
            foreach ($file in (Get-ChildItem $tool.FullName -Recurse -Filter "*.ps1")) {
                foreach ($sourced in (getSourcedFiles $file.FullName)) {
                    foreach ($other in $others) {
                        ($sourced -like "*\scripts\$other\*") | Should Be $false
                    }
                }
            }
        }
    }
}

Describe "読み込み漏れ" -Tag Meta {
    # 起動口からたどれないファイルは、足したのに読み込み忘れている
    It "すべての .ps1 が起動口からたどれる" {
        $entries = @("${scriptsDir}\tebunko\gui.ps1", "${scriptsDir}\tebunko\indexer.ps1")
        $seen = New-Object 'System.Collections.Generic.HashSet[string]'
        $queue = New-Object System.Collections.Queue
        foreach ($entry in $entries) {
            [void]$seen.Add((Resolve-Path -LiteralPath $entry).Path)
            $queue.Enqueue((Resolve-Path -LiteralPath $entry).Path)
        }
        while ($queue.Count -gt 0) {
            foreach ($next in (getSourcedFiles $queue.Dequeue())) {
                if ($seen.Add($next)) { $queue.Enqueue($next) }
            }
        }
        $all = @(Get-ChildItem "${scriptsDir}" -Recurse -Filter "*.ps1" | ForEach-Object { $_.FullName })
        $missing = @($all | Where-Object { !$seen.Contains($_) } | ForEach-Object { Split-Path $_ -Leaf })
        ($missing -join ", ") | Should Be ""
    }
}

Describe "判断層" -Tag Meta {
    # 判断層（入力は素の値、出力は素の値）は画面に触らない。触らないからテストが書ける
    It "判断層のファイルに画面への依存が無い" {
        $files = @(
            "${scriptsDir}\shared\core\text.ps1"
            "${scriptsDir}\tebunko\index\index_name.ps1"
            "${scriptsDir}\tebunko\index\pack_format.ps1"
            "${scriptsDir}\tebunko\search\search_query.ps1"
            "${scriptsDir}\tebunko\search\search_gram.ps1"
            "${scriptsDir}\tebunko\indexer\indexer_decide.ps1"
        ) + @(Get-ChildItem "${scriptsDir}" -Recurse -Filter "*_view.ps1" | ForEach-Object { $_.FullName })
        foreach ($file in $files) {
            $text = [System.IO.File]::ReadAllText($file)
            $hit = [regex]::Matches($text, '\$ui\.|\$window|System\.Windows\.Media')
            "$(Split-Path $file -Leaf): $($hit.Count)" | Should Be "$(Split-Path $file -Leaf): 0"
        }
    }
}