# 前の結果の使い回し（tebunko\search\search_narrow.ps1・search_run.ps1 の invokeSearchRequest）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    function hitKeys($request) {
        $keys = New-Object System.Collections.Generic.List[string]
        $hit = $null
        while ($request.Queue.TryDequeue([ref]$hit)) {
            $keys.Add("$($hit.Root)|$($hit.RelPath)|$($hit.Book)|$($hit.Location)|$($hit.LineNumber)|$($hit.Line)")
        }
        return , $keys
    }

    # 前の語の検索（記録を残す）→ 次の語を、記録を使う要求と使わない要求で流し、結果を返す
    function invokeBoth($first, $second, $folders, [hashtable]$option = @{}) {
        $state = @{ Record = $null }
        $request1 = newSearchRequest $first $true $folders 10000 $option
        invokeSearchRequest $request1 $null $null $state
        $narrow = newSearchRequest $second $true $folders 10000 $option
        invokeSearchRequest $narrow $null $null $state
        $full = newSearchRequest $second $true $folders 10000 $option
        invokeSearchRequest $full
        return @{ First = $request1; Narrow = $narrow; Full = $full; State = $state }
    }
}

Describe "前の結果の使い回し" -Tag Io {
    BeforeAll {
        $tsvRoot = Join-Path $TestDrive "narrow_tsv"
        newTsv "$tsvRoot\見積A社.xlsx\Sheet1.tsv" @("見積書`t100", "見積`t200", "見積書 見積書`t見積書類", "ABC見積", "Abcd")
        newTsv "$tsvRoot\見積A社.xlsx\$(toIndexFileName "Sheet1[図形]")" @("C2`t納期見積書", "見積書`t納期")
        newTsv "$tsvRoot\見積A社.xlsx\$(toIndexFileName "Sheet1[コメント]")" @("B10`t見積書は2個", "2`t見積")
        newTsv "$tsvRoot\sub\請求.docx\ページ001.tsv" @("見積書を送る", "見積について", "2個`t見積 書類")
        newTsv "$tsvRoot\sub\deep\メモ.docx\ページ001.tsv" @("見積書類 abcd")
        $contentIndexRoot = Join-Path $TestDrive "narrow_content_index"
        [void](newContentIndexFiles $tsvRoot $contentIndexRoot)
    }

    It "絞った結果が、全体を探し直した結果と同じ（<first> → <second>）" -TestCases @(
        @{ first = "見積"; second = "見積書"; option = @{} }
        @{ first = "積書"; second = "見積書"; option = @{} }
        @{ first = "見積"; second = "見積 書"; option = @{} }
        @{ first = "abc"; second = "ABCD"; option = @{} }
        @{ first = "abc"; second = "abcd"; option = @{ CaseSensitive = $true } }
        @{ first = "見積"; second = "見積書"; option = @{ FileKinds = @("excel") } }
        @{ first = "見積"; second = "見積書"; option = @{ IncludeShapes = $false; IncludeComments = $false } }
        @{ first = "2"; second = "2個"; option = @{} }
        @{ first = "見積"; second = "見積ぜんぜん無い"; option = @{} }
        @{ first = "納期"; second = "納期見"; option = @{} }
    ) {
        $r = invokeBoth $first $second @($contentIndexRoot) $option
        $r.Narrow.Narrowed | Should -Be $true
        $r.Narrow.NarrowReason | Should -Be "Narrow"
        $r.Full.Narrowed | Should -Be $false
        $narrowKeys = hitKeys $r.Narrow
        $fullKeys = hitKeys $r.Full
        ($narrowKeys -join "`n") | Should -Be ($fullKeys -join "`n")
        $r.Narrow.HitTotal | Should -Be $r.Full.HitTotal
        $r.Narrow.Total | Should -Be $r.Full.Total
        $r.Narrow.IndexTotal | Should -Be $r.Full.IndexTotal
        $r.Narrow.Truncated | Should -Be $r.Full.Truncated
        $r.Narrow.Cancelled | Should -Be $false
        @($r.Narrow.Folders).Count | Should -Be @($r.Full.Folders).Count
        $r.Narrow.Folders[0].Exists | Should -Be $r.Full.Folders[0].Exists
    }

    It "絞った結果から、さらに絞れる（見積 → 見積書 → 見積書類）" {
        $state = @{ Record = $null }
        foreach ($word in @("見積", "見積書", "見積書類")) {
            $request = newSearchRequest $word $true @($contentIndexRoot) 10000
            invokeSearchRequest $request $null $null $state
            $full = newSearchRequest $word $true @($contentIndexRoot) 10000
            invokeSearchRequest $full
            ((hitKeys $request) -join "`n") | Should -Be ((hitKeys $full) -join "`n")
            if ($word -ne "見積") { $request.Narrowed | Should -Be $true }
        }
    }

    It "検索対象を文字列とハッシュで渡しても絞れる" {
        $state = @{ Record = $null }
        $request1 = newSearchRequest "見積" $true @($contentIndexRoot) 10000
        invokeSearchRequest $request1 $null $null $state
        $request2 = newSearchRequest "見積書" $true @(@{ Root = $contentIndexRoot; RelPath = ""; Recurse = $true }) 10000
        invokeSearchRequest $request2 $null $null $state
        $request2.Narrowed | Should -Be $true
    }

    It "RelPath 付き・Recurse なしの検索対象でも同じ結果になる" {
        $folders = @(@{ Root = $contentIndexRoot; RelPath = "sub"; Recurse = $false })
        $r = invokeBoth "見積" "見積書" $folders
        $r.Narrow.Narrowed | Should -Be $true
        ((hitKeys $r.Narrow) -join "`n") | Should -Be ((hitKeys $r.Full) -join "`n")
    }

    Context "絞らない" {
        BeforeEach {
            $tsv = Join-Path $TestDrive "narrow_change_tsv_$([guid]::NewGuid().ToString('N'))"
            newTsv "$tsv\A社.xlsx\Sheet1.tsv" @("見積書`t100", "見積")
            newTsv "$tsv\B社.xlsx\Sheet1.tsv" @("見積書類")
            $script:root = Join-Path $TestDrive "narrow_change_idx_$([guid]::NewGuid().ToString('N'))"
            [void](newContentIndexFiles $tsv $script:root)
            $script:state = @{ Record = $null }
            $first = newSearchRequest "見積" $true @($script:root) 10000
            invokeSearchRequest $first $null $null $script:state
            $script:state.Record | Should -Not -BeNullOrEmpty
        }

        It "<name>" -TestCases @(
            @{ name = "ファイルを書き直した" }
            @{ name = "ファイルを足した" }
            @{ name = "ファイルを消した" }
        ) {
            $target = @(findContentIndexFiles $script:root)[0].Path
            if ($name -eq "ファイルを書き直した") {
                $text = [System.IO.File]::ReadAllText($target, [System.Text.Encoding]::Unicode)
                [System.IO.File]::WriteAllText($target, $text.Replace("見積書", "見積書み"), [System.Text.Encoding]::Unicode)
            } elseif ($name -eq "ファイルを足した") {
                $more = Join-Path $TestDrive "narrow_more_$([guid]::NewGuid().ToString('N'))"
                newTsv "$more\C社.xlsx\Sheet1.tsv" @("見積書追加")
                [void](newContentIndexFiles $more $script:root)
            } else {
                Remove-Item -LiteralPath $target -Force
            }
            $request = newSearchRequest "見積書" $true @($script:root) 10000
            invokeSearchRequest $request $null $null $script:state
            $request.Narrowed | Should -Be $false
            $request.NarrowReason | Should -Be "IndexChanged"
            $full = newSearchRequest "見積書" $true @($script:root) 10000
            invokeSearchRequest $full
            ((hitKeys $request) -join "`n") | Should -Be ((hitKeys $full) -join "`n")
        }

        It "取り消した要求・打ち切った要求では記録を置き換えない" {
            $record = $script:state.Record
            $cancelled = newSearchRequest "見積書" $true @($script:root) 10000
            $cancelled.Stop = $true
            invokeSearchRequest $cancelled $null $null $script:state
            [object]::ReferenceEquals($script:state.Record, $record) | Should -Be $true
            $truncated = newSearchRequest "見積" $true @($script:root) 1
            invokeSearchRequest $truncated $null $null $script:state
            $truncated.Truncated | Should -Be $true
            [object]::ReferenceEquals($script:state.Record, $record) | Should -Be $true
        }

        It "上限なしの要求は記録しない" {
            $state = @{ Record = $null }
            $request = newSearchRequest "見積" $true @($script:root) 0
            invokeSearchRequest $request $null $null $state
            $state.Record | Should -BeNullOrEmpty
        }

        It "列挙の途中で取り消されたら、Cancelled にして記録を変えない" {
            $record = $script:state.Record
            Mock getContentIndexFiles { param($folders, $onProgress, $stopSignal) $stopSignal.Stop = $true; return @{ Folders = @(); ContentIndexFiles = @() } }
            $request = newSearchRequest "見積書" $true @($script:root) 10000
            invokeSearchRequest $request $null $null $script:state
            $request.Cancelled | Should -Be $true
            $request.Narrowed | Should -Be $false
            $request.Queue.IsEmpty | Should -Be $true
            [object]::ReferenceEquals($script:state.Record, $record) | Should -Be $true
        }

        It "記録を渡さない要求は、NarrowReason が空のまま全体を探す" {
            $request = newSearchRequest "見積書" $true @($script:root) 10000
            invokeSearchRequest $request
            $request.NarrowReason | Should -BeNullOrEmpty
            $request.Narrowed | Should -Be $false
        }
    }

    It "テキストの長い行を切ったヒットがあるときは、記録しない" {
        $tsv = Join-Path $TestDrive "narrow_long_tsv"
        newTsv "$tsv\長い.log\$(toIndexFileName "本文")" @(("あ" * 3000) + "対象語" + ("あ" * 3000))
        $idx = Join-Path $TestDrive "narrow_long_idx"
        [void](newContentIndexFiles $tsv $idx)
        $state = @{ Record = $null }
        $request = newSearchRequest "対象語" $true @($idx) 10000
        invokeSearchRequest $request $null $null $state
        $request.Queue.Count | Should -Be 1
        $state.Record | Should -BeNullOrEmpty
    }
}

Describe "前の結果の使い回しの部品" -Tag Unit {
    It "testLiteralSearchWord: <word>（正規表現 <regex>）は <expected>" -TestCases @(
        @{ word = "見積"; regex = $false; expected = $true }
        @{ word = "見積"; regex = $true; expected = $true }
        @{ word = "見積.書"; regex = $true; expected = $false }
        @{ word = "見積 書"; regex = $true; expected = $false }
        @{ word = "見積("; regex = $true; expected = $true }
    ) {
        testLiteralSearchWord $word (!$regex) | Should -Be $expected
    }

    It "toSearchFolderKeys: 文字列とハッシュは同じ形になる" {
        $a = @(toSearchFolderKeys @("C:\IDX"))
        $b = @(toSearchFolderKeys @(@{ Root = "c:\idx"; RelPath = "\"; Recurse = $true }))
        $a | Should -Be $b
    }

    It "testRecordableHitLines: テキストの先頭か末尾が … の行だけ偽" {
        testRecordableHitLines @([pscustomobject]@{ Book = "a.xlsx"; Line = "…あ" }) | Should -Be $true
        testRecordableHitLines @([pscustomobject]@{ Book = "a.log"; Line = "あ" }) | Should -Be $true
        testRecordableHitLines @([pscustomobject]@{ Book = "a.log"; Line = "あ…" }) | Should -Be $false
    }

    Context "記録" {
        BeforeEach {
            $script:req = newSearchRequest "見積" $true @("C:\idx") 100
            $script:files = @(@{ Path = "C:\idx.tsv"; Ticks = 5; Size = 10 }, @{ Path = "C:\idx.tsv"; Ticks = 6; Size = 11 })
            $script:hits = @(
                [pscustomobject]@{ Book = "a.xlsx"; Location = "S"; LineNumber = 1; Line = "見積書" }
                [pscustomobject]@{ Book = "a.xlsx"; Location = "S[図形]"; LineNumber = 2; Line = "見積`t納期" }
                [pscustomobject]@{ Book = "a.xlsx"; Location = "S"; LineNumber = 3; Line = "見積" }
            )
        }

        It "newSearchRecord: 終えた要求なら記録を作り、記録できない要求は null" {
            $record = newSearchRecord $script:req @("C:\idx") $script:files $script:hits
            $record.Paths.Count | Should -Be 2
            $record.Word | Should -Be "見積"
            $script:req.Truncated = $true
            newSearchRecord $script:req @("C:\idx") $script:files $script:hits | Should -BeNullOrEmpty
            $script:req.Truncated = $false
            $script:req.Skipped = @("x")
            newSearchRecord $script:req @("C:\idx") $script:files $script:hits | Should -BeNullOrEmpty
        }

        It "testSearchRecordFiles: 同じ一覧だけ真" {
            $record = newSearchRecord $script:req @("C:\idx") $script:files $script:hits
            testSearchRecordFiles $record $script:files | Should -Be $true
            testSearchRecordFiles $record @($script:files[0]) | Should -Be $false
            testSearchRecordFiles $record @($script:files[0], @{ Path = "C:\idx.tsv"; Ticks = 7; Size = 11 }) | Should -Be $false
        }

        It "selectNarrowHits: 一致する行だけ選び、図形・コメントはセル番地だけの一致を除く" {
            $record = newSearchRecord $script:req @("C:\idx") $script:files $script:hits
            $selected = selectNarrowHits $record "見積書" $false
            $selected.Count | Should -Be 1
            $selected[0].LineNumber | Should -Be 1
        }

        It "invokeNarrowSearch: 一覧が違えば false を返し、IndexChanged にする" {
            $record = newSearchRecord $script:req @("C:\idx") $script:files $script:hits
            Mock getContentIndexFiles { return @{ Folders = @(); ContentIndexFiles = @(@{ Path = "C:\idx\z.tsv"; Ticks = 1; Size = 1 }) } }
            $next = newSearchRequest "見積書" $true @("C:\idx") 100
            invokeNarrowSearch $next @{ Record = $record } | Should -Be $false
            $next.NarrowReason | Should -Be "IndexChanged"
        }
    }
}
