# 変換対象の決定（tebunko_grep\convert\convert_plan.ps1 の createTargetList）のテスト。
. "$PSScriptRoot\..\..\helpers\load.ps1"
. "${scriptsDir}\tebunko_grep\convert\convert_plan.ps1"

function newPrevious {
    param ($rows = @())
    $map = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $rows) { $map[$row.相対パス] = $row }
    return , $map
}

function newCounts {
    param ([hashtable]$pairs = @{})
    $counts = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($key in $pairs.Keys) { $counts[$key] = $pairs[$key] }
    return , $counts
}

Describe "createTargetList" -Tag Io {
    $source = Join-Path $TestDrive "src"
    [System.IO.Directory]::CreateDirectory($source) | Out-Null
    $file = Join-Path $source "a.xlsx"
    [System.IO.File]::WriteAllText($file, "dummy")
    $info = Get-Item -LiteralPath $file
    $updated = formatFileTime $info.LastWriteTime
    $size = [string]$info.Length
    $folder = @{ Path = $source; Name = "売上" }
    # getIndexFiles / removeBookDir が実際のインデックスを見ないよう、テスト用のフォルダに向ける
    ${indexDir} = Join-Path $TestDrive "index"

    It "一覧に無いファイルは変換対象になる（新規）" {
        $result = createTargetList $folder (newPrevious) $null
        $result.Targets.Count | Should Be 1
        $result.Targets[0].相対パス | Should Be "売上\a.xlsx"
        $result.Plan.新規 | Should Be 1
        $result.Plan.ファイル数 | Should Be 1
        $result.Plan.変換対象 | Should Be 1
    }

    It "変換済みで更新が無ければ変換しない" {
        $previous = newPrevious @((newStatusRow "売上\a.xlsx" $updated $size ${stateDone} 1 $updated "" "2"))
        $result = createTargetList $folder $previous (newCounts @{ "売上\a.xlsx" = 1 })
        $result.Targets.Count | Should Be 0
        $result.Plan.変換対象 | Should Be 0
        $result.Rows.Count | Should Be 1
    }

    It "前の抽出版で変換したファイルは、更新が無くても変換し直す（更新ありに数える）" {
        $previous = newPrevious @((newStatusRow "売上\a.xlsx" $updated $size ${stateDone} 1 $updated))
        $result = createTargetList $folder $previous (newCounts @{ "売上\a.xlsx" = 1 })
        $result.Targets.Count | Should Be 1
        $result.Plan.更新あり | Should Be 1
    }

    It "変換済みでも TSV が無ければ変換し直す（変換結果なし）" {
        $previous = newPrevious @((newStatusRow "売上\a.xlsx" $updated $size ${stateDone} 1 $updated))
        $result = createTargetList $folder $previous (newCounts)
        $result.Targets.Count | Should Be 1
        $result.Plan.変換結果なし | Should Be 1
    }

    It "更新されていれば変換対象になる（更新あり）" {
        $previous = newPrevious @((newStatusRow "売上\a.xlsx" "2000/01/01 00:00:00" $size ${stateDone} 1 $updated))
        $result = createTargetList $folder $previous (newCounts @{ "売上\a.xlsx" = 1 })
        $result.Targets.Count | Should Be 1
        $result.Plan.更新あり | Should Be 1
    }

    It "前回失敗して更新が無ければ、変換対象ではなく失敗として返す" {
        $previous = newPrevious @((newStatusRow "売上\a.xlsx" $updated $size ${stateFailed} 0 $updated "開けませんでした"))
        $result = createTargetList $folder $previous (newCounts @{ "売上\a.xlsx" = 1 })
        $result.Targets.Count | Should Be 0
        $result.Failed.Count | Should Be 1
        $result.Plan.前回失敗 | Should Be 1
    }

    It "前回「未変換」で終わっていれば変換対象になる（前回未完了）" {
        $previous = newPrevious @((newStatusRow "売上\a.xlsx" $updated $size ${stateNew}))
        $result = createTargetList $folder $previous (newCounts @{ "売上\a.xlsx" = 1 })
        $result.Targets.Count | Should Be 1
        $result.Plan.前回未完了 | Should Be 1
    }
}