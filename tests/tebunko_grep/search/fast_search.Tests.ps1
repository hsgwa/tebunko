# 高速検索（tebunko_grep\search\fast_search.ps1）のテスト。
# Windows Search の代わりに、system_index の txt を実際に読んで同じ問い合わせに答える偽物を使い、
# 高速検索で集めた集約ファイルを照合した結果が、すべての集約ファイルを照合した結果と同じになることを確かめる。
. "$PSScriptRoot\..\..\helpers\load.ps1"

function script:newFastWorkspace {
    # ワークスペース（index・system_index・状態ファイル）を作り、すべてのフォルダの txt を作って対応済みにする
    param ([string]$root)
    $index = "$root\index"
    foreach ($item in @(
            @{ Path = "営業\2024\A社.xlsx\明細.tsv"; Text = "No`t品名`r`n1`tモニター 27インチ`r`n2`t保守サービス`r`n" },
            @{ Path = "営業\2024\B社.docx\ページ001.tsv"; Text = "東京都千代田区丸の内`r`nモニター台`r`n" },
            @{ Path = "営業\2024\2月\C社.xlsx\一覧.tsv"; Text = "千代田区の見積`r`nABC-1234型番`r`n" },
            @{ Path = "営業\2025\D社.xlsx\表紙.tsv"; Text = "見積書（確定）`r`n" },
            @{ Path = "総務\E社.xlsx\一覧.tsv"; Text = "モニター`r`n備品`r`n" })) {
        $path = "$index\$($item.Path)"
        [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
        [System.IO.File]::WriteAllText($path, $item.Text, ${utf8Bom})
    }
    # インデックス作成と同じく、フォルダごとの集約ファイルにして TSV を消す
    foreach ($folder in (findIndexFoldersWithBooks $index)) {
        [void](updateIndexFolderPack $folder)
    }
    $system = "$root\system_index"
    $statePath = "$root\システムインデックスの状態.tsv"
    $state = newSystemIndexState
    $results = writeSystemIndexFolders (getSystemIndexStaleFolders $index $system $state) $index $system 1
    [void](updateSystemIndexState { param ($s) setSystemIndexResults $s $results; [void]$s.Covered.Add("営業"); [void]$s.Covered.Add("総務") } $statePath)
    return @{ Index = $index; System = $system; State = $statePath }
}

function script:newFakeWindowsSearch {
    # system_index の txt を読んで問い合わせに答える。unreflected の txt（相対パス）は、本文を読み終えていないことにする
    param ([string]$system, [string[]]$unreflected = @())
    # GetNewClosure で作ると別のモジュールの中で動き、テストで読み込んだ関数が見えないため、値は script: に置く
    $script:fakeSystem = $system
    $script:fakeUnreflected = $unreflected
    return {
        param ($sql)
        $system = $script:fakeSystem
        $unreflected = $script:fakeUnreflected
        $rows = New-Object 'System.Collections.Generic.List[object[]]'
        $scope = convertItemUrl ([regex]::Match($sql, "SCOPE='([^']*)'").Groups[1].Value.Replace("''", "'"))
        $files = @([System.IO.Directory]::GetFiles($scope, "システムインデックス*.txt", [System.IO.SearchOption]::AllDirectories))
        if ($sql -match "GatherTime, System.DateModified") {
            foreach ($file in $files) {
                $rel = getRelativePath $file $system
                $ticks = [System.IO.File]::GetLastWriteTimeUtc($file).Ticks
                $gather = if ($unreflected -contains $rel) { [System.DBNull]::Value } else { [datetime]::Now }
                $rows.Add(@("file:$($file.Replace('\', '/'))", $gather, [datetime]::new($ticks - ($ticks % [timespan]::TicksPerSecond))))
            }
            return , $rows
        }
        $grams = @([regex]::Matches($sql, '"(x[0-9a-f]{8})"') | ForEach-Object { $_.Groups[1].Value })
        $split = $sql -match "\[_\]"
        foreach ($file in $files) {
            $name = [System.IO.Path]::GetFileName($file)
            if ($split -ne ($name -ne ${systemIndexFileName})) { continue }
            if ($unreflected -contains (getRelativePath $file $system)) { continue }
            $tokens = ([System.IO.File]::ReadAllText($file)).Trim() -split " "
            if (@($grams | Where-Object { $tokens -notcontains $_ }).Count -eq 0) {
                $rows.Add(@("file:$($file.Replace('\', '/'))"))
            }
        }
        return , $rows
    }
}

function script:hitKeys {
    param ($result)
    return (@($result.Hits | ForEach-Object { "$($_.RelPath)`t$($_.Book)`t$($_.Location)`t$($_.LineNumber)" }) -join "`n")
}

function script:compareSearch {
    # 高速検索とすべての照合で結果が同じか。高速検索の結果を返す
    param ($ws, [string]$word, [object[]]$folders, [scriptblock]$query)
    $fast = getFastSearchPackFiles $word $folders $query $ws.Index $ws.System $ws.State
    $fast | Should Not Be $null
    $all = getIndexPackFiles $folders
    hitKeys (searchPackIndex $word $fast.Packs $true) | Should Be (hitKeys (searchPackIndex $word $all.Packs $true))
    return $fast
}

Describe "getFastSearchPackFiles" -Tag Io {
    It "反映済みなら候補のフォルダだけを照合し、結果はすべての照合と同じ" {
        $ws = newFastWorkspace "$TestDrive\f1"
        $query = newFakeWindowsSearch $ws.System
        $folders = @(@{ Root = $ws.Index; RelPath = "営業"; Recurse = $true }, @{ Root = $ws.Index; RelPath = "総務"; Recurse = $true })
        foreach ($word in @("モニター", "千代田区", "ニター", "abc-1234", "見積", "存在しない語", "27 インチ")) {
            [void](compareSearch $ws $word $folders $query)
        }
        $fast = compareSearch $ws "千代田区" $folders $query
        $fast.Fast.Candidates | Should Be 2
        $fast.Packs.Count | Should Be 3   # 2024（A社の xlsx・B社の docx）と 2月（C社の xlsx）
        (compareSearch $ws "存在しない語" $folders $query).Packs.Count | Should Be 0
    }

    It "反映済みの行は状態から消す" {
        $ws = newFastWorkspace "$TestDrive\f2"
        (readSystemIndexState $ws.State).Pending.Count | Should BeGreaterThan 0
        [void](getFastSearchPackFiles "モニター" @(@{ Root = $ws.Index; RelPath = "営業"; Recurse = $true }) (newFakeWindowsSearch $ws.System) $ws.Index $ws.System $ws.State)
        (readSystemIndexState $ws.State).Pending.Count | Should Be 0
    }

    It "反映されていないフォルダは、候補でなくても照合する" {
        $ws = newFastWorkspace "$TestDrive\f3"
        $query = newFakeWindowsSearch $ws.System @("営業\2024\2月\システムインデックス.txt")
        $fast = compareSearch $ws "千代田区" @(@{ Root = $ws.Index; RelPath = "営業"; Recurse = $true }) $query
        $fast.Fast.Unreflected | Should Be 1
        @($fast.Packs | Where-Object { $_.RelPath -like "*\2月\*" }).Count | Should Be 1
    }

    It "集約ファイルを書き直したフォルダ（反映待ちの日時 0）は、txt を書いたのと同じ秒の中でも照合する" {
        $ws = newFastWorkspace "$TestDrive\f4"
        newTsv "$($ws.Index)\営業\2025\D社.xlsx\表紙.tsv" @("見積書（確定）", "追加した行")
        [void](updateIndexFolderPack "$($ws.Index)\営業\2025")
        [void](markSystemIndexChanged @("営業\2025") $ws.State)
        $fast = compareSearch $ws "追加した行" @(@{ Root = $ws.Index; RelPath = "営業"; Recurse = $true }) (newFakeWindowsSearch $ws.System)
        @($fast.Packs | Where-Object { $_.RelPath -like "*\2025\*" }).Count | Should Be 1
    }

    It "対象外のフォルダ・対応済みでないインデックスは照合する" {
        $ws = newFastWorkspace "$TestDrive\f5"
        [void](updateSystemIndexState { param ($s) [void]$s.Excluded.Add("営業\2025"); [void]$s.Covered.Remove("総務") } $ws.State)
        Remove-Item -LiteralPath "$($ws.System)\営業\2025" -Recurse
        Remove-Item -LiteralPath "$($ws.System)\総務" -Recurse
        $folders = @(@{ Root = $ws.Index; RelPath = "営業"; Recurse = $true }, @{ Root = $ws.Index; RelPath = "総務"; Recurse = $true })
        $fast = compareSearch $ws "見積" $folders (newFakeWindowsSearch $ws.System)
        @($fast.Packs | Where-Object { $_.RelPath -like "*\2025\*" }).Count | Should Be 1
        [void](compareSearch $ws "モニター" $folders (newFakeWindowsSearch $ws.System))
    }

    It "分けた txt でも、すべての語が見つかったフォルダを候補にする" {
        $ws = newFastWorkspace "$TestDrive\f6"
        $systemIndexPartBytes = 100
        $results = writeSystemIndexFolders @("$($ws.Index)\営業\2024") $ws.Index $ws.System 1
        $results[0].Files.Count | Should BeGreaterThan 1
        [void](updateSystemIndexState { param ($s) setSystemIndexResults $s $results } $ws.State)
        $folders = @(@{ Root = $ws.Index; RelPath = "営業"; Recurse = $true })
        $fast = compareSearch $ws "保守サービス" $folders (newFakeWindowsSearch $ws.System)
        @($fast.Packs | Where-Object { $_.RelPath -like "*\2024\content.xlsx.001.tsv" }).Count | Should Be 1
        [void](compareSearch $ws "丸の内" $folders (newFakeWindowsSearch $ws.System))
    }

    It "直下だけの検索対象は、そのフォルダの中だけを照合する" {
        $ws = newFastWorkspace "$TestDrive\f7"
        $fast = compareSearch $ws "千代田区" @(@{ Root = $ws.Index; RelPath = "営業\2024"; Recurse = $false }) (newFakeWindowsSearch $ws.System)
        @($fast.Packs | Where-Object { $_.RelPath -like "*\2月\*" }).Count | Should Be 0
    }

    It "インデックス全体・別の場所・無いフォルダの検索対象は、すべてを列挙する" {
        $ws = newFastWorkspace "$TestDrive\f9"
        $query = newFakeWindowsSearch $ws.System
        $fast = compareSearch $ws "モニター" @($ws.Index) $query
        $fast.Packs.Count | Should Be 5
        $fast = getFastSearchPackFiles "モニター" @(@{ Root = $ws.Index; RelPath = "無い"; Recurse = $true }) $query $ws.Index $ws.System $ws.State
        $fast.Folders[0].Exists | Should Be $false
    }

    It "使えないとき（語が無い・状態ファイルを読めない・問い合わせの失敗・Windows Search を開けない）は null" {
        $ws = newFastWorkspace "$TestDrive\f10"
        $folders = @(@{ Root = $ws.Index; RelPath = "営業"; Recurse = $true })
        getFastSearchPackFiles "見" $folders (newFakeWindowsSearch $ws.System) $ws.Index $ws.System $ws.State | Should Be $null
        getFastSearchPackFiles "見積" $folders { throw "失敗" } $ws.Index $ws.System $ws.State | Should Be $null
        $stream = [System.IO.FileStream]::new($ws.State, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        try {
            getFastSearchPackFiles "見積" $folders (newFakeWindowsSearch $ws.System) $ws.Index $ws.System $ws.State | Should Be $null
        } finally {
            $stream.Dispose()
        }
        Mock openWindowsSearch { $null }
        getFastSearchPackFiles "見積" $folders $null $ws.Index $ws.System $ws.State | Should Be $null
    }
}
