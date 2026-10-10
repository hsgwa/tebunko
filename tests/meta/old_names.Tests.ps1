# 本文インデックスのコードの名前を pack から contentIndex に改めた後、古い名前が戻らないことを確かめるテスト。
# 表は、名前を改めたときに消えた名前。新しい名前を足すときは、この表を直さない（古い名前を使い直さない限り当たらない）。
# 残してよい所は $skip に、ファイルと理由を対にして書く。
BeforeDiscovery {
    $oldNames = @(
    @{ Kind = '関数'; Pattern = 'getPackFileKind'; New = 'getContentIndexFileKind'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'getPackExtension'; New = 'getContentIndexExtension'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'getPackFileName'; New = 'getContentIndexFileName'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'readPackFileName'; New = 'readContentIndexFileName'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'splitPackBooksByExtension'; New = 'splitContentIndexBooksByExtension'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'encodePackValue'; New = 'encodeContentIndexValue'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'decodePackValue'; New = 'decodeContentIndexValue'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'convertPlaceToPackMeta'; New = 'convertPlaceToContentIndexMeta'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'convertPackMetaToPlace'; New = 'convertContentIndexMetaToPlace'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'convertToPackBody'; New = 'convertToContentIndexBody'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'convertBookToPackBlock'; New = 'convertBookToContentIndexBlock'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'convertToPackText'; New = 'convertToContentIndexText'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'planPackParts'; New = 'planContentIndexParts'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'measurePackPartBytes'; New = 'measureContentIndexPartBytes'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'readPackPlaces'; New = 'readContentIndexPlaces'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'splitPackTextByBook'; New = 'splitContentIndexTextByBook'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'getPackContentText'; New = 'getContentIndexBodyText'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'writePackFile'; New = 'writeContentIndexFile'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'readPackText'; New = 'readContentIndexText'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'readPackContext'; New = 'readContentIndexContext'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'readPackLines'; New = 'readContentIndexLines'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'convertIndexFolderToPack'; New = 'convertFolderToContentIndex'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'updateIndexFolderPack'; New = 'updateFolderContentIndex'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'getIndexPackFiles'; New = 'getContentIndexFiles'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'getPackFiles'; New = 'findContentIndexFiles'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'getFastSearchPackFiles'; New = 'getFastSearchContentIndexFiles'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'getPackKeyOf'; New = 'getContentIndexKeyOf'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'searchPackFiles'; New = 'searchContentIndexFiles'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'searchPackIndex'; New = 'searchContentIndex'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'splitPackTasks'; New = 'splitContentIndexTasks'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'newPackWorkerPool'; New = 'newContentIndexWorkerPool'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'newPackIndex'; New = 'newContentIndexFiles'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'writeHitPack'; New = 'writeHitContentIndex'; Docs = $true }
    @{ Kind = '関数'; Pattern = 'testPackCellPrefixed'; New = 'testContentIndexCellPrefixed'; Docs = $true }
    @{ Kind = 'ファイル名'; Pattern = 'pack_format'; New = 'content_index_format'; Docs = $true }
    @{ Kind = 'ファイル名'; Pattern = 'pack_store'; New = 'content_index_store'; Docs = $true }
    @{ Kind = 'ファイル名'; Pattern = 'pack_search'; New = 'content_index_search'; Docs = $true }
    @{ Kind = 'ファイル名'; Pattern = 'convert_to_pack'; New = 'convert_to_content_index'; Docs = $true }
    @{ Kind = '変数（$pack・$packs）'; Pattern = '\$packs?(?![A-Za-z0-9_])'; New = '$contentIndexFile・$contentIndexFiles'; Docs = $false }
    @{ Kind = 'pack で始まる名前（packX）'; Pattern = '(?<![A-Za-z0-9_])pack[A-Z][A-Za-z0-9_]*'; New = 'contentIndexX'; Docs = $false }
    @{ Kind = '語の末尾の Pack（xxxPack・xxxPacks）'; Pattern = '(?<=[a-z])Packs?(?![A-Za-z0-9_])'; New = 'xxxContentIndex・xxxContentIndexFiles'; Docs = $false }
    @{ Kind = '名前の末尾の _pack'; Pattern = '[A-Za-z0-9]_pack(?![A-Za-z0-9_])'; New = 'xxx_content_index'; Docs = $false }
    @{ Kind = '項目 PackPath'; Pattern = 'PackPath(?![A-Za-z0-9_])'; New = 'ContentIndexPath'; Docs = $false }
    )
}

BeforeAll {
    $repo = (Resolve-Path "$PSScriptRoot\..\..").Path
    # 残してよい所（パスは $repo からの相対。前方一致）と理由
    $script:skip = @(
        @{ Path = "tests\testdata\"; Reason = "前の版が作ったデータの見本（互換の確かめ）。名前の確かめの対象ではない" }
        @{ Path = "tools\measure_perf.ps1"; Reason = "計測結果の形（Packs など）は前の版と比べるため変えない" }
        @{ Path = "tools\perf\"; Reason = "計測の道具は、名前を改める前の版も測るため、古い名前も受け付ける" }
        @{ Path = "tests\tools\perf_"; Reason = "計測の道具のテスト。結果の形は変えない" }
        @{ Path = "tests\tools\measure_perf.Tests.ps1"; Reason = "計測の道具が新旧の名前を選ぶことの確かめ。古い名前を使う" }
        @{ Path = "tests\meta\old_names.Tests.ps1"; Reason = "このテスト自身が古い名前を表に持つ" }
        @{ Path = "tests\meta\structure.Tests.ps1"; Reason = "名前に pack を使わない検査が、例として pack を書く" }
        @{ Path = "docs\design\testing\perf.md"; Reason = "計測の道具が、名前を改める前の版の関数名も受け付けることの説明" }
        @{ Path = "docs\design\testing\perf-check.md"; Reason = "性能テストの上限の名前（packLimitSeconds）。性能テストは名前を改めていない" }
    )
    $script:targets = New-Object System.Collections.Generic.List[object]
    foreach ($dir in "scripts", "tests", "tools", "docs") {
        foreach ($file in Get-ChildItem -LiteralPath "$repo\$dir" -Recurse -File -Include "*.ps1", "*.psm1", "*.xaml", "*.md", "*.yml") {
            $rel = $file.FullName.Substring($repo.Length + 1)
            if (@($script:skip | Where-Object { $rel.StartsWith($_.Path) }).Count -gt 0) { continue }
            $script:targets.Add(@{ Rel = $rel; Docs = ($dir -eq "docs"); Lines = @(Get-Content -LiteralPath $file.FullName -Encoding UTF8) })
        }
    }
}

Describe "本文インデックスのコードの名前に、古い名前（pack の系）が戻らない" -Tag Meta {
    It "<Kind> <Pattern> を使っていない（新しい名前は <New>）" -ForEach $oldNames {
        $hits = New-Object System.Collections.Generic.List[string]
        foreach ($target in $script:targets) {
            # 文書は、関数名とファイル名だけを見る（画面の言い回しや計測の項目の説明を拾わないため）
            if ($target.Docs -and !$Docs) { continue }
            for ($i = 0; $i -lt $target.Lines.Count; $i++) {
                if ($target.Lines[$i] -cmatch $Pattern) { $hits.Add("$($target.Rel):$($i + 1)") }
            }
        }
        if ($hits.Count) {
            throw "古い名前が残っています。新しい名前は ${New} です（docs/design/index.md の用語の表）: $(($hits | Select-Object -First 10) -join ' / ')"
        }
    }
}
