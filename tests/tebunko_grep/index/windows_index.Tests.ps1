# Windows インデックス（tebunko_grep\index\windows_index.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"

function script:newIndexTree {
    # index\営業\2024\A社.xlsx\明細.tsv などの小さなインデックスを作る
    param ([string]$root)
    $index = "$root\index"
    foreach ($item in @(
            @{ Path = "営業\2024\A社.xlsx\明細.tsv"; Text = "No`t品名`r`n1`tモニター 27インチ`r`n" },
            @{ Path = "営業\2024\A社.xlsx\表紙.tsv"; Text = "見積書`r`n" },
            @{ Path = "営業\2024\B社.docx\ページ001.tsv"; Text = "保守サービス`r`n" },
            @{ Path = "営業\2024\旧.xls_Sheet1.tsv"; Text = "以前の形式`r`n" },
            @{ Path = "営業\2024\2月\C社.xlsx\一覧.tsv"; Text = "千代田区`r`n" })) {
        $path = "$index\$($item.Path)"
        [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
        [System.IO.File]::WriteAllText($path, $item.Text, ${utf8Bom})
    }
    return $index
}

function script:readTokens {
    param ([string]$path)
    return New-Object System.Collections.Generic.HashSet[string] (, [string[]](([System.IO.File]::ReadAllText($path)).Trim() -split " "))
}

Describe "writeWindowsIndexFolder" -Tag Io {
    It "フォルダ直下の TSV と、直下のブックの TSV から txt を作る（サブフォルダのブックは入れない）" {
        $index = newIndexTree "$TestDrive\w1"
        $system = "$TestDrive\w1\system_index"
        $result = writeWindowsIndexFolder "$index\営業\2024" $index $system
        $result.Rel | Should Be "営業\2024"
        $result.Excluded | Should Be $false
        $result.Files.Count | Should Be 1
        $result.Files[0].Rel | Should Be "営業\2024\Windowsインデックス.txt"
        $txt = "$system\営業\2024\Windowsインデックス.txt"
        $result.Files[0].Ticks | Should Be ([System.IO.File]::GetLastWriteTimeUtc($txt).Ticks)
        $tokens = readTokens $txt
        foreach ($word in @("ニター", "見積", "サービス", "以前の形式")) {
            foreach ($gram in (getSearchGrams $word)) {
                $tokens.Contains($gram) | Should Be $true
            }
        }
        foreach ($gram in (getSearchGrams "千代田区")) {
            $tokens.Contains($gram) | Should Be $false
        }
    }

    It "大きいときは語の範囲で分け、前の txt は消す" {
        $index = newIndexTree "$TestDrive\w2"
        $system = "$TestDrive\w2\system_index"
        [void](writeWindowsIndexFolder "$index\営業\2024" $index $system)
        $windowsIndexPartBytes = 100   # 10 語ごとに分ける
        $result = writeWindowsIndexFolder "$index\営業\2024" $index $system
        $result.Files.Count | Should BeGreaterThan 1
        [System.IO.File]::Exists("$system\営業\2024\Windowsインデックス.txt") | Should Be $false
        $all = New-Object System.Collections.Generic.HashSet[string]
        foreach ($file in $result.Files) {
            $all.UnionWith((readTokens "$system\$($file.Rel)"))
        }
        foreach ($gram in (getSearchGrams "モニター")) {
            $all.Contains($gram) | Should Be $true
        }
    }

    It "パスが長すぎるなら作らずに対象外にする" {
        $index = newIndexTree "$TestDrive\w3"
        $deep = "$index\営業\" + ("長いフォルダ名" * 35)
        [System.IO.Directory]::CreateDirectory((toLongPath "$deep\D社.xlsx")) | Out-Null
        [System.IO.File]::WriteAllText((toLongPath "$deep\D社.xlsx\一覧.tsv"), "長い`r`n", ${utf8Bom})
        try {
            $result = writeWindowsIndexFolder $deep $index "$TestDrive\w3\system_index"
            $result.Excluded | Should Be $true
            $result.Files.Count | Should Be 0
        } finally {
            # Pester の後片付けは長いパスを消せないため、ここで消す
            removeDirectoryRetry "$TestDrive\w3"
        }
    }

    It "TSV が無くなったフォルダは txt を消す" {
        $index = newIndexTree "$TestDrive\w4"
        $system = "$TestDrive\w4\system_index"
        [void](writeWindowsIndexFolder "$index\営業\2024\2月" $index $system)
        Remove-Item -LiteralPath "$index\営業\2024\2月\C社.xlsx" -Recurse
        $result = writeWindowsIndexFolder "$index\営業\2024\2月" $index $system
        $result.Files.Count | Should Be 0
        [System.IO.File]::Exists("$system\営業\2024\2月\Windowsインデックス.txt") | Should Be $false
    }
}

Describe "getWindowsIndexFolderTsvPaths" -Tag Io {
    It "フォルダが無ければ空" {
        (getWindowsIndexFolderTsvPaths "$TestDrive\無いフォルダ").Count | Should Be 0
    }
}

Describe "writeWindowsIndexFolders" -Tag Io {
    It "フォルダが無ければ何もしない" {
        (writeWindowsIndexFolders @() "$TestDrive\i" "$TestDrive\s").Count | Should Be 0
    }

    It "並列でも 1 つずつでも同じ結果になり、始めた順に返す" {
        $index = newIndexTree "$TestDrive\p"
        $folders = @("$index\営業\2024", "$index\営業\2024\2月")
        $one = writeWindowsIndexFolders $folders $index "$TestDrive\p\s1" 1
        $two = writeWindowsIndexFolders $folders $index "$TestDrive\p\s2" 2
        ($two | ForEach-Object { $_.Rel }) -join "|" | Should Be "営業\2024|営業\2024\2月"
        for ($i = 0; $i -lt 2; $i++) {
            [System.IO.File]::ReadAllText("$TestDrive\p\s2\$($two[$i].Files[0].Rel)") | Should Be ([System.IO.File]::ReadAllText("$TestDrive\p\s1\$($one[$i].Files[0].Rel)"))
        }
    }

    It "止めるよう求められたら、始めていない分は作らない" {
        $index = newIndexTree "$TestDrive\stop"
        $results = writeWindowsIndexFolders @("$index\営業\2024", "$index\営業\2024\2月") $index "$TestDrive\stop\s" 1 { $true }
        $results.Count | Should Be 0
    }
}

Describe "readWindowsIndexState / updateWindowsIndexState" -Tag Io {
    It "無ければ空。書き換えて読み直せる" {
        $path = "$TestDrive\state\Windowsインデックスの状態.tsv"
        (readWindowsIndexState $path).Pending.Count | Should Be 0
        updateWindowsIndexState { param ($s) [void]$s.Covered.Add("営業"); $s.Pending["営業\a\Windowsインデックス.txt"] = 5 } $path | Should Be $true
        $state = readWindowsIndexState $path
        $state.Covered.Contains("営業") | Should Be $true
        $state.Pending["営業\a\Windowsインデックス.txt"] | Should Be 5
        [System.IO.File]::ReadAllBytes($path)[0] | Should Be 0xEF
    }

    It "ほかが開いている間は、書き換えも読み込みもあきらめる" {
        $path = "$TestDrive\locked.tsv"
        [void](updateWindowsIndexState { param ($s) [void]$s.Covered.Add("a") } $path)
        $stream = [System.IO.FileStream]::new($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        try {
            updateWindowsIndexState { param ($s) } $path | Should Be $false
            readWindowsIndexState $path | Should Be $null
        } finally {
            $stream.Dispose()
        }
    }
}

Describe "setWindowsIndexResults / markWindowsIndexChanged / removeWindowsIndexEntries" -Tag Io {
    It "作り直した結果で、そのフォルダの行だけを置き換える" {
        $state = convertFromWindowsIndexState @("反映待ち`t営業\a\Windowsインデックス_1.txt`t1", "反映待ち`t営業\a\b\Windowsインデックス.txt`t2", "対象外`t営業\a`t")
        setWindowsIndexResults $state @(@{ Rel = "営業\a"; Files = @(@{ Rel = "営業\a\Windowsインデックス.txt"; Ticks = 3 }); Excluded = $false })
        $state.Pending.ContainsKey("営業\a\Windowsインデックス_1.txt") | Should Be $false
        $state.Pending["営業\a\Windowsインデックス.txt"] | Should Be 3
        $state.Pending["営業\a\b\Windowsインデックス.txt"] | Should Be 2
        $state.Excluded.Count | Should Be 0
        setWindowsIndexResults $state @(@{ Rel = "営業\c"; Files = @(); Excluded = $true })
        $state.Excluded.Contains("営業\c") | Should Be $true
    }

    It "TSV を入れ替えたフォルダを、日時 0 で反映待ちにする" {
        $path = "$TestDrive\mark.tsv"
        markWindowsIndexChanged @("営業\a") $path | Should Be $true
        (readWindowsIndexState $path).Pending["営業\a\Windowsインデックス.txt"] | Should Be 0
    }

    It "フォルダとその中の行、インデックスの対応済みを消す" {
        $state = convertFromWindowsIndexState @("対応済み`t営業`t", "反映待ち`t営業\a\Windowsインデックス.txt`t1", "反映待ち`t営業2\a\Windowsインデックス.txt`t1", "対象外`t営業\b`t")
        removeWindowsIndexEntries $state "営業" "営業"
        $state.Covered.Count | Should Be 0
        $state.Excluded.Count | Should Be 0
        @($state.Pending.Keys) -join "|" | Should Be "営業2\a\Windowsインデックス.txt"
    }
}

Describe "getWindowsIndexStaleFolders" -Tag Io {
    It "txt が無い・TSV より古い・取り込みの途中で止まった・TSV が無くなったフォルダを返す" {
        $index = newIndexTree "$TestDrive\stale"
        $system = "$TestDrive\stale\system_index"
        $state = newWindowsIndexState
        # txt が無い
        (getWindowsIndexStaleFolders $index $system $state) -join "|" | Should Be "$index\営業\2024|$index\営業\2024\2月"
        setWindowsIndexResults $state (writeWindowsIndexFolders @("$index\営業\2024", "$index\営業\2024\2月") $index $system 1)
        (getWindowsIndexStaleFolders $index $system $state).Count | Should Be 0
        # TSV の方が新しい
        [System.IO.File]::SetLastWriteTimeUtc("$index\営業\2024\A社.xlsx\表紙.tsv", [datetime]::UtcNow.AddMinutes(5))
        (getWindowsIndexStaleFolders $index $system $state) -join "|" | Should Be "$index\営業\2024"
        setWindowsIndexResults $state (writeWindowsIndexFolders @("$index\営業\2024") $index $system 1)
        [System.IO.File]::SetLastWriteTimeUtc("$index\営業\2024\A社.xlsx\表紙.tsv", [datetime]::UtcNow.AddMinutes(-5))
        (getWindowsIndexStaleFolders $index $system $state).Count | Should Be 0
        # 取り込みの途中で止まった（反映待ちの日時が txt と合わない）
        $state.Pending["営業\2024\2月\Windowsインデックス.txt"] = 1
        (getWindowsIndexStaleFolders $index $system $state) -join "|" | Should Be "$index\営業\2024\2月"
        # TSV が無くなった
        setWindowsIndexResults $state (writeWindowsIndexFolders @("$index\営業\2024\2月") $index $system 1)
        Remove-Item -LiteralPath "$index\営業\2024\2月\C社.xlsx" -Recurse
        (getWindowsIndexStaleFolders $index $system $state) -join "|" | Should Be "$index\営業\2024\2月"
    }

    It "対象外のフォルダ・index が無いときは返さない" {
        $index = newIndexTree "$TestDrive\stale2"
        $state = convertFromWindowsIndexState @("対象外`t営業\2024`t", "対象外`t営業\2024\2月`t")
        (getWindowsIndexStaleFolders $index "$TestDrive\stale2\system_index" $state).Count | Should Be 0
        (getWindowsIndexStaleFolders "$TestDrive\none" "$TestDrive\none2" $state).Count | Should Be 0
    }
}

Describe "updateWindowsIndexes" -Tag Io {
    It "作り直しが要るフォルダを作り、インデックスを対応済みにする。2 回目は何もしない" {
        $index = newIndexTree "$TestDrive\u1"
        $system = "$TestDrive\u1\system_index"
        $path = "$TestDrive\u1\state.tsv"
        $result = updateWindowsIndexes $index $system $path "$TestDrive\u1\stop"
        $result.Built | Should Be 2
        $result.Unfinished | Should Be 0
        $state = readWindowsIndexState $path
        $state.Covered.Contains("営業") | Should Be $true
        $state.Pending.Count | Should Be 2
        (updateWindowsIndexes $index $system $path "$TestDrive\u1\stop").Built | Should Be 0
    }

    It "無くなったインデックスの txt と状態の行を消す" {
        $index = newIndexTree "$TestDrive\u2"
        $system = "$TestDrive\u2\system_index"
        $path = "$TestDrive\u2\state.tsv"
        [void](updateWindowsIndexes $index $system $path "$TestDrive\u2\stop")
        Remove-Item -LiteralPath "$index\営業" -Recurse
        [void](updateWindowsIndexes $index $system $path "$TestDrive\u2\stop")
        [System.IO.Directory]::Exists("$system\営業") | Should Be $false
        $state = readWindowsIndexState $path
        $state.Covered.Count | Should Be 0
        $state.Pending.Count | Should Be 0
    }

    It "中止要求があれば作らず、そのインデックスは対応済みにしない" {
        $index = newIndexTree "$TestDrive\u3"
        [System.IO.File]::WriteAllText("$TestDrive\u3\stop", "")
        $result = updateWindowsIndexes $index "$TestDrive\u3\system_index" "$TestDrive\u3\state.tsv" "$TestDrive\u3\stop"
        $result.Built | Should Be 0
        $result.Unfinished | Should Be 1
        (readWindowsIndexState "$TestDrive\u3\state.tsv").Covered.Count | Should Be 0
    }

    It "状態ファイルを読めなければ何もしない" {
        $index = newIndexTree "$TestDrive\u4"
        $path = "$TestDrive\u4\state.tsv"
        [void](updateWindowsIndexState { param ($s) } $path)
        $stream = [System.IO.FileStream]::new($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        try {
            (updateWindowsIndexes $index "$TestDrive\u4\system_index" $path "$TestDrive\u4\stop").Unfinished | Should Be -1
        } finally {
            $stream.Dispose()
        }
    }
}

Describe "removeIndex / renameIndex の Windows インデックス" -Tag Io {
    It "インデックスを削除・名前変更すると、同じワークスペースの system_index の分を消す" {
        $index = newIndexTree "$TestDrive\ix"
        [void](updateWindowsIndexes $index "$TestDrive\ix\system_index" "$TestDrive\ix\Windowsインデックスの状態.tsv" "$TestDrive\ix\stop")
        renameIndex "営業" "営業2" $index "$TestDrive\ix\取り込み一覧.tsv"
        [System.IO.Directory]::Exists("$TestDrive\ix\system_index\営業") | Should Be $false
        (readWindowsIndexState "$TestDrive\ix\Windowsインデックスの状態.tsv").Covered.Count | Should Be 0
        [void](updateWindowsIndexes $index "$TestDrive\ix\system_index" "$TestDrive\ix\Windowsインデックスの状態.tsv" "$TestDrive\ix\stop")
        [System.IO.Directory]::Exists("$TestDrive\ix\system_index\営業2") | Should Be $true
        removeIndex "営業2" $index "$TestDrive\ix\取り込み一覧.tsv"
        [System.IO.Directory]::Exists("$TestDrive\ix\system_index\営業2") | Should Be $false
    }
}

Describe "removeWindowsIndexOf" -Tag Io {
    It "system_index の同じフォルダと状態の行を消す" {
        $index = newIndexTree "$TestDrive\rm"
        $system = "$TestDrive\rm\system_index"
        $path = "$TestDrive\rm\state.tsv"
        $results = writeWindowsIndexFolders @("$index\営業\2024") $index $system 1
        [void](updateWindowsIndexState { param ($s) setWindowsIndexResults $s $results; [void]$s.Covered.Add("営業") } $path)
        removeWindowsIndexOf "営業" $system $path | Should Be $true
        [System.IO.Directory]::Exists("$system\営業") | Should Be $false
        $state = readWindowsIndexState $path
        $state.Covered.Count | Should Be 0
        $state.Pending.Count | Should Be 0
    }

    It "インデックスの中のフォルダだけを消すときは、インデックスの対応済みを残す" {
        $index = newIndexTree "$TestDrive\rm2"
        $system = "$TestDrive\rm2\system_index"
        $path = "$TestDrive\rm2\state.tsv"
        $results = writeWindowsIndexFolders @("$index\営業\2024", "$index\営業\2024\2月") $index $system 1
        [void](updateWindowsIndexState { param ($s) setWindowsIndexResults $s $results; [void]$s.Covered.Add("営業") } $path)
        removeWindowsIndexOf "営業\2024\2月" $system $path | Should Be $true
        $state = readWindowsIndexState $path
        $state.Covered.Contains("営業") | Should Be $true
        @($state.Pending.Keys) -join "|" | Should Be "営業\2024\Windowsインデックス.txt"
    }
}
