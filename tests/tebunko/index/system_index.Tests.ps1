# システムインデックス（tebunko\index\system_index.ps1）のテスト
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

Describe "writeSystemIndexFolder" -Tag Io {
    It "フォルダ直下の TSV と、直下のブックの TSV から txt を作る（サブフォルダのブックは入れない）" {
        $index = newIndexTree "$TestDrive\w1"
        $system = "$TestDrive\w1\system_index"
        $result = writeSystemIndexFolder "$index\営業\2024" $index $system
        $result.Rel | Should Be "営業\2024"
        $result.Excluded | Should Be $false
        $result.Files.Count | Should Be 1
        $result.Files[0].Rel | Should Be "営業\2024\システムインデックス.txt"
        $txt = "$system\営業\2024\システムインデックス.txt"
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
        [void](writeSystemIndexFolder "$index\営業\2024" $index $system)
        $systemIndexPartBytes = 100   # 10 語ごとに分ける
        $result = writeSystemIndexFolder "$index\営業\2024" $index $system
        $result.Files.Count | Should BeGreaterThan 1
        [System.IO.File]::Exists("$system\営業\2024\システムインデックス.txt") | Should Be $false
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
            $result = writeSystemIndexFolder $deep $index "$TestDrive\w3\system_index"
            $result.Excluded | Should Be $true
            $result.Files.Count | Should Be 0
        } finally {
            # Pester の後片付けは長いパスを消せないため、ここで消す
            removeDirectoryRetry "$TestDrive\w3"
        }
    }

    It "集約ファイルからは、メタ情報の行（ファイル名・シート名）を除いて txt を作る" {
        $index = "$TestDrive\w5\index"
        [System.IO.Directory]::CreateDirectory("$index\営業") | Out-Null
        writePackFile "$index\営業\content.xlsx.001.tsv" (convertToPackText @(@{ Name = "山田商事.xlsx"; Places = @(@{ Place = "見積"; Text = "保守サービス`r`n" }) }))
        $system = "$TestDrive\w5\system_index"
        $result = writeSystemIndexFolder "$index\営業" $index $system
        $result.Files.Count | Should Be 1
        $tokens = readTokens "$system\営業\システムインデックス.txt"
        foreach ($gram in (getSearchGrams "サービス")) {
            $tokens.Contains($gram) | Should Be $true
        }
        foreach ($word in @("山田商事", "ファイル名", "シート")) {
            @((getSearchGrams $word) | Where-Object { $tokens.Contains($_) }).Count | Should Be 0
        }
    }

    It "中身（texts）を渡せば、ファイルを読まずにその中身から作る。空なら txt を消す" {
        $index = newIndexTree "$TestDrive\w6"
        $system = "$TestDrive\w6\system_index"
        $text = convertToPackText @(@{ Name = "E社.xlsx"; Places = @(@{ Place = "S"; Text = "渡した中身`r`n" }) })
        [void](writeSystemIndexFolder "$index\営業\2024" $index $system @($text))
        $tokens = readTokens "$system\営業\2024\システムインデックス.txt"
        foreach ($gram in (getSearchGrams "渡した中身")) {
            $tokens.Contains($gram) | Should Be $true
        }
        # フォルダの TSV は読まない
        @((getSearchGrams "見積") | Where-Object { $tokens.Contains($_) }).Count | Should Be 0
        $result = writeSystemIndexFolder "$index\営業\2024" $index $system @()
        $result.Files.Count | Should Be 0
        [System.IO.File]::Exists("$system\営業\2024\システムインデックス.txt") | Should Be $false
    }

    It "TSV が無くなったフォルダは txt を消す" {
        $index = newIndexTree "$TestDrive\w4"
        $system = "$TestDrive\w4\system_index"
        [void](writeSystemIndexFolder "$index\営業\2024\2月" $index $system)
        Remove-Item -LiteralPath "$index\営業\2024\2月\C社.xlsx" -Recurse
        $result = writeSystemIndexFolder "$index\営業\2024\2月" $index $system
        $result.Files.Count | Should Be 0
        [System.IO.File]::Exists("$system\営業\2024\2月\システムインデックス.txt") | Should Be $false
    }
}

Describe "writeSystemIndexFolders" -Tag Io {
    It "並列でも 1 つずつでも同じ結果になり、始めた順に返す" {
        $index = newIndexTree "$TestDrive\p"
        $folders = @("$index\営業\2024", "$index\営業\2024\2月")
        $one = writeSystemIndexFolders $folders $index "$TestDrive\p\s1" 1
        $two = writeSystemIndexFolders $folders $index "$TestDrive\p\s2" 2
        ($two | ForEach-Object { $_.Rel }) -join "|" | Should Be "営業\2024|営業\2024\2月"
        for ($i = 0; $i -lt 2; $i++) {
            [System.IO.File]::ReadAllText("$TestDrive\p\s2\$($two[$i].Files[0].Rel)") | Should Be ([System.IO.File]::ReadAllText("$TestDrive\p\s1\$($one[$i].Files[0].Rel)"))
        }
    }

    It "並列で作っていて書けなかったら、作ったことにせず例外にする" {
        $index = newIndexTree "$TestDrive\pe"
        # system_index という名前のファイルがあると、フォルダを作れない
        [System.IO.File]::WriteAllText("$TestDrive\pe\system_index", "")
        { writeSystemIndexFolders @("$index\営業\2024", "$index\営業\2024\2月") $index "$TestDrive\pe\system_index" 2 } | Should Throw
    }

    It "止めるよう求められたら、始めていない分は作らない" {
        $index = newIndexTree "$TestDrive\stop"
        $results = writeSystemIndexFolders @("$index\営業\2024", "$index\営業\2024\2月") $index "$TestDrive\stop\s" 1 { $true }
        $results.Count | Should Be 0
    }
}

Describe "readSystemIndexState / updateSystemIndexState" -Tag Io {
    It "無ければ空。書き換えて読み直せる" {
        $path = "$TestDrive\state\システムインデックスの状態.tsv"
        (readSystemIndexState $path).Pending.Count | Should Be 0
        updateSystemIndexState { param ($s) [void]$s.Covered.Add("営業"); $s.Pending["営業\a\システムインデックス.txt"] = 5 } $path | Should Be $true
        $state = readSystemIndexState $path
        $state.Covered.Contains("営業") | Should Be $true
        $state.Pending["営業\a\システムインデックス.txt"] | Should Be 5
        [System.IO.File]::ReadAllBytes($path)[0] | Should Be 0xEF
    }

    It "ほかが開いている間は、書き換えも読み込みもあきらめる" {
        $path = "$TestDrive\locked.tsv"
        [void](updateSystemIndexState { param ($s) [void]$s.Covered.Add("a") } $path)
        $stream = [System.IO.FileStream]::new($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        try {
            updateSystemIndexState { param ($s) } $path | Should Be $false
            readSystemIndexState $path | Should Be $null
        } finally {
            $stream.Dispose()
        }
    }
}

Describe "setSystemIndexResults / markSystemIndexChanged / removeSystemIndexEntries" -Tag Io {
    It "作り直した結果で、そのフォルダの行だけを置き換える" {
        $state = convertFromSystemIndexState @("反映待ち`t営業\a\システムインデックス_1.txt`t1", "反映待ち`t営業\a\b\システムインデックス.txt`t2", "対象外`t営業\a`t")
        setSystemIndexResults $state @(@{ Rel = "営業\a"; Files = @(@{ Rel = "営業\a\システムインデックス.txt"; Ticks = 3 }); Excluded = $false })
        $state.Pending.ContainsKey("営業\a\システムインデックス_1.txt") | Should Be $false
        $state.Pending["営業\a\システムインデックス.txt"] | Should Be 3
        $state.Pending["営業\a\b\システムインデックス.txt"] | Should Be 2
        $state.Excluded.Count | Should Be 0
        setSystemIndexResults $state @(@{ Rel = "営業\c"; Files = @(); Excluded = $true })
        $state.Excluded.Contains("営業\c") | Should Be $true
    }

    It "TSV を入れ替えたフォルダを、日時 0 で反映待ちにする" {
        $path = "$TestDrive\mark.tsv"
        markSystemIndexChanged @("営業\a") $path | Should Be $true
        (readSystemIndexState $path).Pending["営業\a\システムインデックス.txt"] | Should Be 0
    }

    It "フォルダとその中の行、インデックスの対応済みを消す" {
        $state = convertFromSystemIndexState @("対応済み`t営業`t", "反映待ち`t営業\a\システムインデックス.txt`t1", "反映待ち`t営業2\a\システムインデックス.txt`t1", "対象外`t営業\b`t")
        removeSystemIndexEntries $state "営業" "営業"
        $state.Covered.Count | Should Be 0
        $state.Excluded.Count | Should Be 0
        @($state.Pending.Keys) -join "|" | Should Be "営業2\a\システムインデックス.txt"
    }
}

Describe "getSystemIndexStaleFolders" -Tag Io {
    It "txt が無い・TSV より古い・取り込みの途中で止まった・TSV が無くなったフォルダを返す" {
        $index = newIndexTree "$TestDrive\stale"
        $system = "$TestDrive\stale\system_index"
        $state = newSystemIndexState
        # txt が無い
        (getSystemIndexStaleFolders $index $system $state) -join "|" | Should Be "$index\営業\2024|$index\営業\2024\2月"
        setSystemIndexResults $state (writeSystemIndexFolders @("$index\営業\2024", "$index\営業\2024\2月") $index $system 1)
        (getSystemIndexStaleFolders $index $system $state).Count | Should Be 0
        # TSV の方が新しい
        [System.IO.File]::SetLastWriteTimeUtc("$index\営業\2024\A社.xlsx\表紙.tsv", [datetime]::UtcNow.AddMinutes(5))
        (getSystemIndexStaleFolders $index $system $state) -join "|" | Should Be "$index\営業\2024"
        setSystemIndexResults $state (writeSystemIndexFolders @("$index\営業\2024") $index $system 1)
        [System.IO.File]::SetLastWriteTimeUtc("$index\営業\2024\A社.xlsx\表紙.tsv", [datetime]::UtcNow.AddMinutes(-5))
        (getSystemIndexStaleFolders $index $system $state).Count | Should Be 0
        # 取り込みの途中で止まった（反映待ちの日時が txt と合わない）
        $state.Pending["営業\2024\2月\システムインデックス.txt"] = 1
        (getSystemIndexStaleFolders $index $system $state) -join "|" | Should Be "$index\営業\2024\2月"
        # TSV が無くなった
        setSystemIndexResults $state (writeSystemIndexFolders @("$index\営業\2024\2月") $index $system 1)
        Remove-Item -LiteralPath "$index\営業\2024\2月\C社.xlsx" -Recurse
        (getSystemIndexStaleFolders $index $system $state) -join "|" | Should Be "$index\営業\2024\2月"
    }

    It "対象外のフォルダ・index が無いときは返さない" {
        $index = newIndexTree "$TestDrive\stale2"
        $state = convertFromSystemIndexState @("対象外`t営業\2024`t", "対象外`t営業\2024\2月`t")
        (getSystemIndexStaleFolders $index "$TestDrive\stale2\system_index" $state).Count | Should Be 0
        (getSystemIndexStaleFolders "$TestDrive\none" "$TestDrive\none2" $state).Count | Should Be 0
    }
}

Describe "updateSystemIndexes" -Tag Io {
    It "作り直しが要るフォルダを作り、インデックスを対応済みにする。2 回目は何もしない" {
        $index = newIndexTree "$TestDrive\u1"
        $system = "$TestDrive\u1\system_index"
        $path = "$TestDrive\u1\state.tsv"
        $result = updateSystemIndexes $index $system $path { $false }
        $result.Built | Should Be 2
        $result.Unfinished | Should Be 0
        $state = readSystemIndexState $path
        $state.Covered.Contains("営業") | Should Be $true
        $state.Pending.Count | Should Be 2
        (updateSystemIndexes $index $system $path { $false }).Built | Should Be 0
    }

    It "無くなったインデックスの txt と状態の行を消す" {
        $index = newIndexTree "$TestDrive\u2"
        $system = "$TestDrive\u2\system_index"
        $path = "$TestDrive\u2\state.tsv"
        [void](updateSystemIndexes $index $system $path { $false })
        Remove-Item -LiteralPath "$index\営業" -Recurse
        [void](updateSystemIndexes $index $system $path { $false })
        [System.IO.Directory]::Exists("$system\営業") | Should Be $false
        $state = readSystemIndexState $path
        $state.Covered.Count | Should Be 0
        $state.Pending.Count | Should Be 0
    }

    It "中止を求められていれば（shouldStop）作らず、そのインデックスは対応済みにしない" {
        $index = newIndexTree "$TestDrive\u3"
        $result = updateSystemIndexes $index "$TestDrive\u3\system_index" "$TestDrive\u3\state.tsv" { $true }
        $result.Built | Should Be 0
        $result.Unfinished | Should Be 1
        (readSystemIndexState "$TestDrive\u3\state.tsv").Covered.Count | Should Be 0
    }

    # Pester 3 の Mock は Describe・Context の中の後のテストにも効くため、Context で囲む
    Context "状態ファイルに書けないとき" {
        It "状態ファイルに書けなければ、次のインデックス作成に回すと知らせる（対応済みにしない）" {
            $index = newIndexTree "$TestDrive\u5"
            Mock updateSystemIndexState { $false }
            $result = updateSystemIndexes $index "$TestDrive\u5\system_index" "$TestDrive\u5\state.tsv" { $false }
            $result.Built | Should Be 2
            (readSystemIndexState "$TestDrive\u5\state.tsv").Covered.Count | Should Be 0
        }
    }

    It "状態ファイルを読めなければ何もしない" {
        $index = newIndexTree "$TestDrive\u4"
        $path = "$TestDrive\u4\state.tsv"
        [void](updateSystemIndexState { param ($s) } $path)
        $stream = [System.IO.FileStream]::new($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        try {
            (updateSystemIndexes $index "$TestDrive\u4\system_index" $path { $false }).Unfinished | Should Be -1
        } finally {
            $stream.Dispose()
        }
    }
}

Describe "removeIndex / renameIndex の システムインデックス" -Tag Io {
    It "インデックスを削除・名前変更すると、同じワークスペースの system_index の分を消す" {
        $index = newIndexTree "$TestDrive\ix"
        [void](updateSystemIndexes $index "$TestDrive\ix\system_index" "$TestDrive\ix\システムインデックスの状態.tsv" { $false })
        renameIndex "営業" "営業2" $index "$TestDrive\ix\取り込み一覧.tsv"
        [System.IO.Directory]::Exists("$TestDrive\ix\system_index\営業") | Should Be $false
        (readSystemIndexState "$TestDrive\ix\システムインデックスの状態.tsv").Covered.Count | Should Be 0
        [void](updateSystemIndexes $index "$TestDrive\ix\system_index" "$TestDrive\ix\システムインデックスの状態.tsv" { $false })
        [System.IO.Directory]::Exists("$TestDrive\ix\system_index\営業2") | Should Be $true
        removeIndex "営業2" $index "$TestDrive\ix\取り込み一覧.tsv"
        [System.IO.Directory]::Exists("$TestDrive\ix\system_index\営業2") | Should Be $false
    }
}

Describe "removeSystemIndexOf" -Tag Io {
    It "system_index の同じフォルダと状態の行を消す" {
        $index = newIndexTree "$TestDrive\rm"
        $system = "$TestDrive\rm\system_index"
        $path = "$TestDrive\rm\state.tsv"
        $results = writeSystemIndexFolders @("$index\営業\2024") $index $system 1
        [void](updateSystemIndexState { param ($s) setSystemIndexResults $s $results; [void]$s.Covered.Add("営業") } $path)
        removeSystemIndexOf "営業" $system $path | Should Be $true
        [System.IO.Directory]::Exists("$system\営業") | Should Be $false
        $state = readSystemIndexState $path
        $state.Covered.Count | Should Be 0
        $state.Pending.Count | Should Be 0
    }

    It "インデックスの中のフォルダだけを消すときは、インデックスの対応済みを残す" {
        $index = newIndexTree "$TestDrive\rm2"
        $system = "$TestDrive\rm2\system_index"
        $path = "$TestDrive\rm2\state.tsv"
        $results = writeSystemIndexFolders @("$index\営業\2024", "$index\営業\2024\2月") $index $system 1
        [void](updateSystemIndexState { param ($s) setSystemIndexResults $s $results; [void]$s.Covered.Add("営業") } $path)
        removeSystemIndexOf "営業\2024\2月" $system $path | Should Be $true
        $state = readSystemIndexState $path
        $state.Covered.Contains("営業") | Should Be $true
        @($state.Pending.Keys) -join "|" | Should Be "営業\2024\システムインデックス.txt"
    }
}
