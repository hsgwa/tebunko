# インデックス作成の鍵の取り方の決まりを確かめるテスト。
# 鍵はミューテックス（同じセッション）とワークスペースの鍵のファイル（別のセッション・PC・パスの書き方）の 2 つで、
# 両方を取るのは enterIndexerLock（indexer_state.ps1）だけ。ミューテックスだけを直接取ると、別のセッション・PC のインデックス作成と重なる。
BeforeAll {
    $here = (Resolve-Path "$PSScriptRoot\..").Path
    $scriptsDir = (Resolve-Path "$here\..\scripts").Path
    $script:statePath = "$scriptsDir\tebunko\indexer\indexer_state.ps1"

    function getIndexerMutexUsers {
        # newAppMutex "indexer" を書いているスクリプトの一覧（コメントの行は数えない）
        param ([string]$root)
        $found = New-Object System.Collections.Generic.List[string]
        foreach ($file in Get-ChildItem -LiteralPath $root -Recurse -Filter *.ps1) {
            foreach ($line in [System.IO.File]::ReadAllLines($file.FullName)) {
                if ($line.TrimStart().StartsWith("#")) { continue }
                if ($line -match 'newAppMutex\s+"indexer"') {
                    $found.Add($file.FullName.Substring($root.Length + 1))
                    break
                }
            }
        }
        return $found.ToArray()
    }
}

Describe "インデックス作成の鍵（ミューテックスと鍵のファイル）" -Tag Meta {
    It "newAppMutex ""indexer"" を書くのは indexer_state.ps1 だけ（鍵は enterIndexerLock 経由で取る）" {
        $users = @(getIndexerMutexUsers $scriptsDir)
        $users | Should -Be @("tebunko\indexer\indexer_state.ps1")
    }

    It "検査が効いている（ミューテックスだけを取る例を見つけられる）" {
        $dir = Join-Path $TestDrive "scripts_sample"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        [System.IO.File]::WriteAllText("$dir\bad.ps1", "`$m = newAppMutex ""indexer"" `$dir `"x`"`r`n")
        [System.IO.File]::WriteAllText("$dir\comment.ps1", "# newAppMutex ""indexer"" は使わない`r`n")
        @(getIndexerMutexUsers $dir) | Should -Be @("bad.ps1")
    }

    It "enterIndexerLock は、ミューテックスを取ってから鍵のファイルを開く" {
        $text = [System.IO.File]::ReadAllText($statePath)
        $start = $text.IndexOf("function enterIndexerLock")
        $start | Should -BeGreaterThan -1
        $mutexAt = $text.IndexOf('= newAppMutex "indexer"', $start)
        $fileAt = $text.IndexOf("= openLockFile", $start)
        $mutexAt | Should -BeGreaterThan -1
        $fileAt | Should -BeGreaterThan $mutexAt
    }
}
