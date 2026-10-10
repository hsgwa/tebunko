# インデックスの追加・編集・削除（tebunko\ui\index\index_edit.ps1 と index_list.ps1 の updateIndexSourceFile）のテスト。
# 画面と裏の列は偽物にする。ワークスペースが共有フォルダにある場面は、実在しない UNC のパスで表す。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"

    $script:calls = New-Object System.Collections.Generic.List[string]
    $script:statuses = New-Object System.Collections.Generic.List[string]
    $script:jobs = New-Object System.Collections.Generic.List[object]

    function setStatus { param([string]$text) $script:statuses.Add($text) }
    function startJob {
        # 裏の列には出さず、依頼を取っておく（テストが onDone を呼んで、終わったことにする）
        param ([scriptblock]$scriptBlock, [object[]]$arguments, [scriptblock]$onDone, [string]$queue = "default")
        $script:jobs.Add(@{ ScriptBlock = $scriptBlock; Arguments = $arguments; OnDone = $onDone; Queue = $queue })
    }
    function updateIndexingButton { $script:calls.Add("updateIndexingButton:$($script:indexBusy)") }
    function saveTargets { $script:calls.Add("saveTargets") }
    function refreshIndexViews { $script:calls.Add("refreshIndexViews") }
    function loadTargets { $script:calls.Add("loadTargets") }
    function updateFolderItemStatus { param($item) $script:calls.Add("updateFolderItemStatus") }

    . "${scriptsDir}\tebunko\ui\index_view.ps1"
    . "${scriptsDir}\tebunko\ui\index\index_edit.ps1"

    # index_list.ps1 は画面の型を使うため、読み込まずに updateIndexSourceFile だけを取り出す
    $ast = [System.Management.Automation.Language.Parser]::ParseFile("${scriptsDir}\tebunko\ui\index\index_list.ps1", [ref]$null, [ref]$null)
    $node = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq "updateIndexSourceFile" }, $true)
    . ([scriptblock]::Create($node.Extent.Text))

    function runJob {
        # 裏の仕事の中身を、この場の Mock が効くように作り直して（startJob と同じく文字列から）流す
        param ($job)
        $block = [scriptblock]::Create($job.ScriptBlock.ToString())
        $jobArguments = $job.Arguments
        & $block @jobArguments
    }

    function newItem {
        param ([string]$name, [string]$path)
        $item = [pscustomobject]@{ Name = $name; Path = $path; StatusChecked = $true }
        $item | Add-Member ScriptMethod SetName { param($n) $this.Name = $n }
        $item | Add-Member ScriptMethod SetPath { param($p) $this.Path = $p }
        return $item
    }
}

Describe "applyIndexEdit" -Tag Unit {
    BeforeEach {
        $script:calls.Clear()
        $script:statuses.Clear()
        $script:jobs.Clear()
        $script:indexBusy = $false
        $script:workspace = newTestWorkspace @{ IndexDir = "\\fileserver\共有\ws\work\index"; StatusFile = "\\fileserver\共有\ws\work\ingest_status.tsv" }
        $script:item = newItem "営業" "\\fileserver\共有\営業"
        $script:other = newItem "技術" "\\fileserver\共有\技術"
        $script:targetItems = New-Object System.Collections.Generic.List[object]
        $script:targetItems.Add($script:item)
        $script:targetItems.Add($script:other)
    }

    It "名前を変えるときは、画面のスレッドで改名せず、裏の列（network）に出して、indexBusy で他の操作を止める" {
        Mock renameIndex { throw "画面のスレッドで改名した" }

        applyIndexEdit $item @{ Path = $item.Path; Name = "営業部" }

        Should -Invoke renameIndex -Times 0
        $script:jobs.Count | Should -Be 1
        $script:jobs[0].Queue | Should -Be "network"
        $script:indexBusy | Should -Be $true
        (getIndexJobBlocker $false $script:indexBusy $false) | Should -Not -Be ""
        $script:statuses[0] | Should -Be (getIndexStoreJobStatus "rename" "営業" "営業部")
    }

    It "名前を変える間は行を変えず、検索の画面も読み直さない。結果が届いてから反映する" {
        applyIndexEdit $item @{ Path = $item.Path; Name = "営業部" }

        $item.Name | Should -Be "営業"
        $script:calls | Should -Not -Contain "refreshIndexViews"
        $script:calls | Should -Not -Contain "saveTargets"

        & $script:jobs[0].OnDone @() ""

        $item.Name | Should -Be "営業部"
        $script:indexBusy | Should -Be $false
        $script:calls | Should -Contain "saveTargets"
        $script:calls | Should -Contain "refreshIndexViews"
        $script:statuses[$script:statuses.Count - 1] | Should -Be "インデックスを変更しました（名前 [営業] → [営業部]）"
    }

    It "裏の仕事は、改名のあと、新しい名前と場所で元のフォルダの記録を書き直す" {
        Mock renameIndex { $script:calls.Add("renameIndex") }
        Mock writeSourceFolderFile { $script:calls.Add("writeSourceFolderFile:" + (($folders | ForEach-Object { "$($_.Name)=$($_.Path)" }) -join ",")) }

        applyIndexEdit $item @{ Path = "\\fileserver\共有\新営業"; Name = "営業部" }
        $script:calls.Clear()
        runJob $script:jobs[0]

        $script:calls[0] | Should -Be "renameIndex"
        $script:calls[1] | Should -Be "writeSourceFolderFile:営業部=\\fileserver\共有\新営業,技術=\\fileserver\共有\技術"
        Should -Invoke renameIndex -Times 1 -Exactly -ParameterFilter { $oldName -eq "営業" -and $newName -eq "営業部" -and $dir -eq "\\fileserver\共有\ws\work\index" }
    }

    It "名前の変更に失敗したら、理由をステータスに出し、一覧を保存済みの内容に戻す（行は変えない）" {
        applyIndexEdit $item @{ Path = $item.Path; Name = "営業部" }

        & $script:jobs[0].OnDone @() "届きません"

        $item.Name | Should -Be "営業"
        $script:indexBusy | Should -Be $false
        $script:statuses[$script:statuses.Count - 1] | Should -Be "インデックス [営業] の名前の変更に失敗しました：届きません"
        $script:calls | Should -Contain "loadTargets"
        $script:calls | Should -Contain "refreshIndexViews"
        $script:calls | Should -Not -Contain "saveTargets"
    }

    It "裏の仕事は、元のフォルダの記録だけが書けなくても失敗にせず、書けなかった理由を別に返す" {
        Mock renameIndex { }
        Mock writeSourceFolderFile { throw "届きません" }

        applyIndexEdit $item @{ Path = $item.Path; Name = "営業部" }
        $result = @(runJob $script:jobs[0])

        Should -Invoke renameIndex -Times 1 -Exactly
        $result[0].SourceFolderError | Should -Be "届きません"
    }

    It "記録だけが書けなかったときも、改名は反映して保存し、書けなかったことだけを知らせる（改名の失敗にはしない）" {
        applyIndexEdit $item @{ Path = $item.Path; Name = "営業部" }

        & $script:jobs[0].OnDone @(@{ SourceFolderError = "届きません" }) ""

        $item.Name | Should -Be "営業部"
        $script:indexBusy | Should -Be $false
        $script:calls | Should -Contain "saveTargets"
        $script:calls | Should -Not -Contain "loadTargets"
        $script:statuses[$script:statuses.Count - 1] | Should -Be (getSourceFolderFileFailedStatus "届きません")
        ($script:statuses -join "`n") | Should -Not -Match "名前の変更に失敗しました"
    }

    It "名前を変えている途中で閉じようとして待っていたなら、反映が終わってから閉じる" {
        $script:closed = 0
        $script:window = [pscustomobject]@{}
        $script:window | Add-Member ScriptMethod Close { $script:closed++ }
        $script:closeAfterIndexJob = $true
        applyIndexEdit $item @{ Path = $item.Path; Name = "営業部" }
        $script:calls.Clear()

        & $script:jobs[0].OnDone @() ""

        $script:calls | Should -Contain "saveTargets"
        $script:closed | Should -Be 1
        $script:closeAfterIndexJob | Should -Be $false
    }

    It "待っていなければ閉じない" {
        $script:closed = 0
        $script:window = [pscustomobject]@{}
        $script:window | Add-Member ScriptMethod Close { $script:closed++ }
        $script:closeAfterIndexJob = $false
        applyIndexEdit $item @{ Path = $item.Path; Name = "営業部" }

        & $script:jobs[0].OnDone @() ""

        $script:closed | Should -Be 0
    }

    It "ローカルのワークスペースの改名も、裏の仕事（default の列）で行う" {
        $script:workspace = newTestWorkspace @{ IndexDir = "C:\Users\test\tebunko_ws\work\index" }

        applyIndexEdit $item @{ Path = $item.Path; Name = "営業部" }

        $script:jobs.Count | Should -Be 1
        $script:jobs[0].Queue | Should -Be "default"
    }

    It "場所だけを変えるときは、裏の改名はせず、すぐ反映して元のフォルダの記録を書き直す" {
        applyIndexEdit $item @{ Path = "\\fileserver\共有\新営業"; Name = "営業" }

        $item.Path | Should -Be "\\fileserver\共有\新営業"
        $script:calls | Should -Contain "saveTargets"
        $script:calls | Should -Contain "refreshIndexViews"
        # 記録の書き直しだけが裏に出る（改名の仕事ではない）
        $script:jobs.Count | Should -Be 1
        $script:indexBusy | Should -Be $false
    }

    It "変わったところが無ければ、何もしない" {
        applyIndexEdit $item @{ Path = $item.Path; Name = $item.Name }

        $script:jobs.Count | Should -Be 0
        $script:calls.Count | Should -Be 0
    }
}

Describe "startIndexStoreJob（削除）" -Tag Unit {
    BeforeEach {
        $script:calls.Clear()
        $script:statuses.Clear()
        $script:jobs.Clear()
        $script:indexBusy = $false
        $script:workspace = newTestWorkspace @{ IndexDir = "\\fileserver\共有\ws\work\index"; StatusFile = "\\fileserver\共有\ws\work\ingest_status.tsv" }
    }

    It "裏の列に出し、終わったら indexBusy を戻して後始末を呼ぶ" {
        startIndexStoreJob "delete" "営業" "" @() { setStatus "済み" }

        $script:jobs[0].Queue | Should -Be "network"
        $script:indexBusy | Should -Be $true
        $script:statuses[0] | Should -Be (getIndexStoreJobStatus "delete" "営業" "")

        & $script:jobs[0].OnDone @() ""

        $script:indexBusy | Should -Be $false
        $script:statuses[$script:statuses.Count - 1] | Should -Be "済み"
        $script:calls | Should -Contain "refreshIndexViews"
    }

    It "失敗したら、理由をステータスに出す" {
        startIndexStoreJob "delete" "営業" "" @() { setStatus "済み" }

        & $script:jobs[0].OnDone @() "届きません"

        $script:statuses[$script:statuses.Count - 1] | Should -Be "インデックス [営業] の削除に失敗しました：届きません"
        $script:calls | Should -Not -Contain "loadTargets"
    }

    It "裏の仕事は removeIndex を呼ぶ" {
        Mock removeIndex { }
        startIndexStoreJob "delete" "営業" "" @() { }

        runJob $script:jobs[0]

        Should -Invoke removeIndex -Times 1 -Exactly -ParameterFilter { $name -eq "営業" }
    }
}

Describe "deleteIndexes（複数削除）" -Tag Unit {
    BeforeEach {
        $script:calls.Clear()
        $script:statuses.Clear()
        $script:jobs.Clear()
        $script:indexBusy = $false
        $script:closeAfterIndexJob = $false
        $script:targetItems = New-Object System.Collections.Generic.List[object]
        function testIndexOperable { param($operation) $true }
        function showConfirm { param($title, $heading, $hint, $choices) "delete" }
        function showBulkIndexResult { param($operation, $results) }
        function updateIndexListView { }
        function updateIndexSourceFile { }
    }

    It "<expected> の列に出す" -TestCases @(
        @{ dir = "\\fileserver\共有\ws\work\index"; expected = "network" }
        @{ dir = "C:\Users\test\tebunko_ws\work\index"; expected = "default" }
    ) {
        param ($dir, $expected)
        $script:workspace = newTestWorkspace @{ IndexDir = $dir }

        deleteIndexes @("営業", "技術")

        $script:jobs.Count | Should -Be 1
        $script:jobs[0].Queue | Should -Be $expected
        $script:indexBusy | Should -Be $true
    }

    It "削除の途中で閉じようとして待っていたなら、終わってから閉じる" {
        $script:workspace = newTestWorkspace @{ IndexDir = "C:\Users\test\tebunko_ws\work\index" }
        $script:closed = 0
        $script:window = [pscustomobject]@{}
        $script:window | Add-Member ScriptMethod Close { $script:closed++ }
        deleteIndexes @("営業", "技術")
        $script:closeAfterIndexJob = $true

        & $script:jobs[0].OnDone @() ""

        $script:closed | Should -Be 1
        $script:indexBusy | Should -Be $false
    }
}

Describe "updateIndexSourceFile" -Tag Unit {
    BeforeEach {
        $script:statuses.Clear()
        $script:jobs.Clear()
        $script:workspace = newTestWorkspace @{ IndexDir = "\\fileserver\共有\ws\work\index" }
        $script:targetItems = New-Object System.Collections.Generic.List[object]
        $script:targetItems.Add((newItem "営業" "\\fileserver\共有\営業"))
    }

    It "場所に触れず、書き直しを裏の列（network）に出す" {
        Mock writeSourceFolderFile { throw "画面のスレッドで書いた" }

        updateIndexSourceFile

        Should -Invoke writeSourceFolderFile -Times 0
        $script:jobs.Count | Should -Be 1
        $script:jobs[0].Queue | Should -Be "network"
    }

    It "裏の仕事は、今の一覧と IndexDir を引数で受けて書く" {
        Mock writeSourceFolderFile { }
        updateIndexSourceFile

        runJob $script:jobs[0]

        Should -Invoke writeSourceFolderFile -Times 1 -Exactly -ParameterFilter {
            $dir -eq "\\fileserver\共有\ws\work\index" -and @($folders).Count -eq 1 -and $folders[0].Name -eq "営業"
        }
    }

    It "書けなかったら、理由をステータスに出す" {
        updateIndexSourceFile

        & $script:jobs[0].OnDone @() "届きません"

        $script:statuses | Should -Be @((getSourceFolderFileFailedStatus "届きません"))
    }
}
