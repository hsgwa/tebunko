# ［インデックス管理］のエクスポート・インポートの判断（tebunko\ui\index\index_archive_view.ps1）のテスト。
BeforeAll {
    . "$PSScriptRoot\..\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\ui\index\index_list_view.ps1"
    . "${scriptsDir}\tebunko\ui\index\index_edit_view.ps1"
    . "${scriptsDir}\tebunko\ui\index\index_archive_view.ps1"

    function newItem {
        param ([string]$name, [string]$path)
        return [pscustomobject]@{ Name = $name; Path = $path }
    }
}

Describe "getImportResultStatus" -Tag Unit {
    It "件数を伝える" {
        $status = getImportResultStatus @{ Name = "営業"; Files = 12; Warnings = @() }
        $status | Should -Be "インデックス [営業] をインポートしました（12 ファイル）。"
    }

    It "Warnings があれば添える" {
        $status = getImportResultStatus @{ Name = "営業"; Files = 12; Warnings = @("元のフォルダ「C:\共有」は、インデックス [既存] としても登録されています。") }
        $status | Should -Match "既存"
    }
}

Describe "getImportSuggestedName / testImportNameCollision" -Tag Unit {
    # continueImportIndex は getUsedIndexNames の結果を変数に受けて渡す。0・1・2 件以上のどれでも、重なりを見落とさない
    BeforeAll {
        function newUsed([string[]]$names) {
            $items = @($names | ForEach-Object { newItem $_ "C:\data\$_" })
            $used = getUsedIndexNames $items
            return , $used
        }
    }

    It "<label>: 提案する名前は <expected>、確認が要るかは <collision>" -TestCases @(
        @{ label = "0 件"; names = @(); name = "営業"; expected = "営業"; collision = $false }
        @{ label = "1 件で重ならない"; names = @("経理"); name = "営業"; expected = "営業"; collision = $false }
        @{ label = "1 件で重なる"; names = @("営業"); name = "営業"; expected = "営業(2)"; collision = $true }
        @{ label = "2 件以上で重なる"; names = @("経理", "営業", "総務"); name = "営業"; expected = "営業(2)"; collision = $true }
        @{ label = "2 件以上で重ならない"; names = @("経理", "総務"); name = "営業"; expected = "営業"; collision = $false }
        @{ label = "(2) も使われている"; names = @("営業", "営業(2)"); name = "営業"; expected = "営業(3)"; collision = $true }
        @{ label = "大文字・小文字は区別しない"; names = @("Sales", "経理"); name = "sales"; expected = "sales(2)"; collision = $true }
    ) {
        param ($label, $names, $name, $expected, $collision)
        $used = newUsed $names
        getImportSuggestedName $name $used | Should -Be $expected
        testImportNameCollision $name $used | Should -Be $collision
    }

    It "集合が 1 要素の配列に入っていても（@(...) で受けた場合）、同じ答えを返す" {
        $used = @(newUsed @("経理", "営業"))
        testImportNameCollision "営業" $used | Should -BeTrue
        getImportSuggestedName "営業" $used | Should -Be "営業(2)"
    }

    It "設定から外したインデックスが残っている名前（フォルダが決まっていない）は、使用中として別の名前を提案し、選び方を出す" {
        $items = @((newItem "経理" "C:\data\経理"))
        $leftovers = @([pscustomobject]@{ Name = "営業"; Path = "C:\data\営業" })
        $used = getUsedIndexNames $items $null $leftovers
        getImportSuggestedName "営業" $used | Should -Be "営業(2)"
        testImportNameCollision "営業" $used | Should -BeTrue
        $choice = getImportCollisionChoiceView "営業(2)" "営業" $used ${importCollisionRename}
        $choice.Visible | Should -BeTrue
        $choice.Candidate | Should -Be "営業(2)"
        $choice.TargetName | Should -Be "営業"
    }
}


Describe "testIndexExportInput" -Tag Unit {
    It "<label>" -TestCases @(
        @{ label = "フォルダが空なら、指定するよう伝える"; folder = ""; expected = "書き出し先のフォルダを指定してください。" }
        @{ label = "空白だけでも同じ"; folder = "  "; expected = "書き出し先のフォルダを指定してください。" }
        @{ label = "フォルダがあれば問題なし"; folder = "C:\共有\書き出し"; expected = "" }
    ) {
        param ($label, $folder, $expected)
        testIndexExportInput $folder | Should -Be $expected
    }
}

Describe "getIndexExportHeader（エクスポートのダイアログの枠）" -Tag Unit {
    BeforeAll {
        $script:now = [datetime]"2026-09-27 10:00:00"
    }

    It "<label>" -TestCases @(
        @{ label = "1 件: 名前・ファイル数・zip の名前"; items = @(@{ Name = "営業部"; FileCountText = "2,530" }); title = "営業部"; size = "2,530 ファイル"; file = "営業部_インデックス_20260927.zip として保存します" }
        @{ label = "2 件: 件数と合計のファイル数、名前の形"; items = @(@{ Name = "営業"; FileCountText = "1,000" }, @{ Name = "総務"; FileCountText = "20" }); title = "2 件のインデックス"; size = "1,020 ファイル"; file = "インデックスごとに <名前>_インデックス_20260927.zip として保存します" }
        @{ label = "ファイル数が分からない（未作成）: 大きさの行は空"; items = @(@{ Name = "営業"; FileCountText = "－" }); title = "営業"; size = ""; file = "営業_インデックス_20260927.zip として保存します" }
    ) {
        param ($label, $items, $title, $size, $file)
        $header = getIndexExportHeader $items $script:now
        $header.Title | Should -Be $title
        $header.Size | Should -Be $size
        $header.FileText | Should -Be $file
        $header.Caution | Should -Be "このファイルには、元のファイルの本文と元のフォルダの場所が含まれます。元のファイルと同じように取り扱ってください。"
    }
}

Describe "getIndexExportHeader（大きさを数え終えたあと）" -Tag Unit {
    It "<label>" -TestCases @(
        @{ label = "1 件: ファイル数のあとに大きさ"; items = @(@{ Name = "営業部"; FileCountText = "2,530" }); bytes = 98566144; size = "2,530 ファイル・94.0 MB" }
        @{ label = "2 件: 合計のファイル数と合計の大きさ"; items = @(@{ Name = "営業"; FileCountText = "1,000" }, @{ Name = "総務"; FileCountText = "20" }); bytes = 52428800; size = "1,020 ファイル・50.0 MB" }
        @{ label = "0 バイトでも 0.1 MB と出す"; items = @(@{ Name = "営業"; FileCountText = "1" }); bytes = 0; size = "1 ファイル・0.1 MB" }
        @{ label = "ファイル数が分からないときは大きさも出さない"; items = @(@{ Name = "営業"; FileCountText = "－" }); bytes = 1048576; size = "" }
        @{ label = "大きさを渡さなければ今までどおり"; items = @(@{ Name = "営業"; FileCountText = "5" }); bytes = -1; size = "5 ファイル" }
    ) {
        param ($label, $items, $bytes, $size)
        (getIndexExportHeader $items ([datetime]"2026-09-27 10:00:00") $bytes).Size | Should -Be $size
    }

    It "数えている間に対象が無くなったときの文言" {
        getIndexExportMissingText | Should -Be "対象のインデックスが見つかりません"
    }
}

Describe "getIndexExportNotice" -Tag Unit {
    It "別の PC・ワークスペースで使えることを伝える" {
        getIndexExportNotice | Should -Match "インポートして使えます"
    }
}

Describe "testIndexImportInput" -Tag Unit {
    It "フォルダが空なら、指定するよう伝える" {
        testIndexImportInput "" "営業" | Should -Be "元のフォルダを指定してください。"
    }

    It "インデックス名として使えなければ理由を返す" {
        testIndexImportInput "C:\共有\営業部" "" | Should -Match "入力してください"
    }

    It "既にある名前でも断らない（上書きの確認に回すため）" {
        testIndexImportInput "C:\共有\営業部" "既存" | Should -Be ""
    }
}

Describe "getIndexImportNotice" -Tag Unit {
    It "名前と元のフォルダを変えられることを伝える" {
        getIndexImportNotice | Should -Match "名前と、元のフォルダの場所を変えられます"
    }
}

Describe "getIndexImportHeader（インポートのダイアログの枠）" -Tag Unit {
    It "<label>" -TestCases @(
        @{ label = "日時・ファイル数・大きさ・版"; info = @{ IndexName = "営業部"; Files = 2530; Bytes = [long](56.6 * 1MB); AppVersion = "0.3.0"; ExportedAt = "2026-09-27T10:00:00+09:00" }
            exported = "エクスポートした日時 2026/09/27 10:00"; size = "2,530 ファイル・56.6 MB・v0.3.0" }
        @{ label = "版に v が付いていても重ねない・日時が無ければ空"; info = @{ IndexName = "営業"; Files = 3; Bytes = 10KB; AppVersion = "v0.4.0"; ExportedAt = "" }
            exported = ""; size = "3 ファイル・0.1 MB・v0.4.0" }
        @{ label = "版が無ければ版を省く"; info = @{ IndexName = "営業"; Files = 1; Bytes = 2MB; AppVersion = ""; ExportedAt = "壊れた値" }
            exported = ""; size = "1 ファイル・2.0 MB" }
    ) {
        param ($label, $info, $exported, $size)
        $header = getIndexImportHeader $info
        $header.Name | Should -Be $info.IndexName
        $header.Size | Should -Be $size
        if ($info.ExportedAt -like "2026-*") {
            # 時差で日付が変わる環境でも崩れないよう、時刻は環境の時差に直した値と比べる
            $header.Exported | Should -Be ("エクスポートした日時 " + ([datetimeoffset]$info.ExportedAt).LocalDateTime.ToString("yyyy/MM/dd HH:mm"))
        } else {
            $header.Exported | Should -Be $exported
        }
    }
}

Describe "getImportSourceCheckText（元のフォルダの下の 1 行）" -Tag Unit {
    It "<label>" -TestCases @(
        @{ label = "空なら何も出さない"; folder = ""; exists = $true; text = ""; kind = "None" }
        @{ label = "あれば「見つかりました」"; folder = "C:\共有\営業部"; exists = $true; text = "見つかりました"; kind = "Ok" }
        @{ label = "無ければ、検索はできるが更新できないと伝える"; folder = "C:\共有\営業部"; exists = $false
            text = "元のフォルダがこの PC に見つかりません。検索はできますが、元のフォルダを設定するまで更新できません。"; kind = "Warn" }
        @{ label = "有無を調べていなければ何も出さない"; folder = "C:\共有\営業部"; exists = $null; text = ""; kind = "None" }
        @{ label = "ほかのインデックスと同じ"; folder = "C:\共有\顧客"; exists = $true; text = "「顧客」の元のフォルダと重なっています"; kind = "Ng" }
        @{ label = "ほかのインデックスの中"; folder = "C:\共有\顧客\東"; exists = $true; text = "「顧客」の元のフォルダと重なっています"; kind = "Ng" }
        @{ label = "ほかのインデックスを含む"; folder = "C:\共有"; exists = $null; text = "「顧客」の元のフォルダと重なっています"; kind = "Ng" }
    ) {
        param ($label, $folder, $exists, $text, $kind)
        $result = getImportSourceCheckText $folder $exists @(@{ Name = "顧客"; Path = "C:\共有\顧客" })
        $result.Text | Should -Be $text
        $result.Kind | Should -Be $kind
    }
}

Describe "getIndexArchiveCancelledStatus" -Tag Unit {
    It "<operation> を取りやめた文" -TestCases @(
        @{ operation = "エクスポート"; expected = "エクスポートを取りやめました" }
        @{ operation = "インポート"; expected = "インポートを取りやめました" }
    ) {
        param ($operation, $expected)
        getIndexArchiveCancelledStatus $operation | Should -Be $expected
    }
}

Describe "getImportCollisionChoiceView" -Tag Unit {
    # 今の一覧は 営業・総務。zip の名前は、重なる「営業」か、重ならない「経理」
    It "<label>" -TestCases @(
        @{ label = "初めの状態（zip の名前が重なり、名前の欄は別の名前の候補）は出す"; name = "営業(2)"; zip = "営業"; mode = "Rename"; visible = $true; effective = "Rename"; danger = $false }
        @{ label = "名前の欄を重なる名前に戻したら出す（大文字・小文字は区別しない）"; name = " 営業 "; zip = "営業"; mode = "Overwrite"; visible = $true; effective = "Overwrite"; danger = $true }
        @{ label = "重なる名前を手で入れたら、zip の名前が重ならなくても出す"; name = "総務"; zip = "経理"; mode = "Rename"; visible = $true; effective = "Rename"; danger = $false }
        @{ label = "名前の欄を手で一覧に無い名前にしたら隠す"; name = "営業の別件"; zip = "営業"; mode = "Overwrite"; visible = $false; effective = "Rename"; danger = $false }
        @{ label = "zip の名前が重ならず、名前の欄がそのままなら出さない"; name = "経理"; zip = "経理"; mode = "Rename"; visible = $false; effective = "Rename"; danger = $false }
        @{ label = "名前の欄が空なら出さない"; name = " "; zip = "経理"; mode = "Rename"; visible = $false; effective = "Rename"; danger = $false }
        @{ label = "zip が一覧にあり、名前の欄が別の一覧の名前のとき、置き換える相手は名前の欄の名前"; name = "総務"; zip = "営業"; mode = "Overwrite"; visible = $true; effective = "Overwrite"; danger = $true; target = "総務"; overwriteName = "総務" }
        @{ label = "zip が一覧に無く、名前の欄が一覧の名前のとき、置き換える相手は名前の欄の名前"; name = "総務"; zip = "経理"; mode = "Overwrite"; visible = $true; effective = "Overwrite"; danger = $true; target = "総務"; overwriteName = "総務" }
        @{ label = "置き換えるを選んだあと、名前の欄を候補に戻したら、置き換えにならず赤くもしない"; name = "営業(2)"; zip = "営業"; mode = "Overwrite"; visible = $true; effective = "Rename"; danger = $false; target = "営業" }
        @{ label = "置き換えるを選んだあと、名前の欄を一覧に無い名前にしたら、置き換えにならず赤くもしない"; name = "営業の別件"; zip = "営業"; mode = "Overwrite"; visible = $false; effective = "Rename"; danger = $false; target = "営業" }
    ) {
        param ($label, $name, $zip, $mode, $visible, $effective, $danger, $target, $overwriteName)
        $view = getImportCollisionChoiceView $name $zip @("営業", "総務") $mode
        $view.Visible | Should -Be $visible
        $view.Mode | Should -Be $effective
        $view.Danger | Should -Be $danger
        if ($target) { $view.TargetName | Should -Be $target }
        if ($overwriteName) { $view.OverwriteName | Should -Be $overwriteName }
    }

    It "選んだときに名前の欄へ入れる名前は、別の名前にするなら候補、置き換えるなら置き換える相手の名前" {
        $view = getImportCollisionChoiceView "営業(2)" "営業" @("営業", "総務")
        $view.Candidate | Should -Be "営業(2)"
        $view.OverwriteName | Should -Be "営業"
        $view.RenameText | Should -Be "別の名前にする"
        $view.OverwriteText | Should -Be "置き換える"
    }

    It "見出しに重なっている名前を、置き換えるときは今のインデックスが置き換わることを添える" {
        (getImportCollisionChoiceView "営業(2)" "営業" @("営業", "総務")).Heading | Should -Be "同じ名前のインデックス [営業] があります"
        (getImportCollisionChoiceView "営業(2)" "営業" @("営業", "総務")).Notice | Should -Be ""
        $overwrite = getImportCollisionChoiceView "総務" "営業" @("営業", "総務") "Overwrite"
        $overwrite.Heading | Should -Be "同じ名前のインデックス [総務] があります"
        $overwrite.Notice | Should -Be "今のインデックス [総務] は、インポートするもので置き換わります。"
        (getImportCollisionChoiceView "新規" "新規" @("営業")).Heading | Should -Be ""
    }

    # 「旧」は今の一覧に無く、設定から外して作業フォルダに残っているだけの名前（使えない名前の集合には入る）
    It "<label>" -TestCases @(
        @{ label = "外れた名前を zip が指す初めの状態は、置き換えを出さず、別の名前（旧(2)）だけにする"; name = "旧(2)"; zip = "旧"; mode = "Rename"; overwriteVisible = $false; effective = "Rename"; candidate = "旧(2)" }
        @{ label = "外れた名前を手で入れ、置き換えるを選んでいても、置き換えにならず赤くもしない"; name = "旧"; zip = "営業"; mode = "Overwrite"; overwriteVisible = $false; effective = "Rename"; candidate = "営業(2)" }
        @{ label = "一覧にある名前は、これまでどおり置き換えられる"; name = "総務"; zip = "旧"; mode = "Overwrite"; overwriteVisible = $true; effective = "Overwrite"; candidate = "旧(2)" }
    ) {
        param ($label, $name, $zip, $mode, $overwriteVisible, $effective, $candidate)
        $view = getImportCollisionChoiceView $name $zip @("営業", "総務", "旧") $mode @("営業", "総務")
        $view.Visible | Should -Be $true
        $view.OverwriteVisible | Should -Be $overwriteVisible
        $view.Mode | Should -Be $effective
        $view.Danger | Should -Be ($effective -eq "Overwrite")
        $view.Candidate | Should -Be $candidate
        if (!$overwriteVisible) {
            $view.Heading | Should -Be "作業フォルダに同じ名前のインデックス [旧] が残っています（置き換えはできません）".Replace("[旧]", "[$($view.TargetName)]")
            $view.Heading | Should -Not -Match "今のインデックス"
            $view.Notice | Should -Be ""
        }
    }

    It "一覧が空（listedNames が @()）なら、使えない名前はすべて外れて残っている名前として、置き換えを出さない" {
        $view = getImportCollisionChoiceView "旧" "営業" @("営業", "旧") "Overwrite" @()
        $view.OverwriteVisible | Should -Be $false
        $view.Mode | Should -Be "Rename"
        $view.Danger | Should -Be $false
    }

    It "一覧にある名前のときの見出しは「同じ名前のインデックス [総務] があります」" {
        $view = getImportCollisionChoiceView "総務" "旧" @("営業", "総務", "旧") "Rename" @("営業", "総務")
        $view.Heading | Should -Be "同じ名前のインデックス [総務] があります"
        $view.OverwriteVisible | Should -Be $true
    }

    It "一覧の名前を渡さなければ、使えない名前はすべて一覧にあるものとして扱う（これまでの動き）" {
        $view = getImportCollisionChoiceView "旧" "営業" @("営業", "総務", "旧") "Overwrite"
        $view.OverwriteVisible | Should -Be $true
        $view.Mode | Should -Be "Overwrite"
    }
}

Describe "testIndexImportInput（元のフォルダの重なり）" -Tag Unit {
    BeforeAll {
        $items = @((newItem "売上" "C:\data\売上"), (newItem "見積" "C:\data\見積"))
    }

    It "<label>: 別のインデックスと重なれば断る（追加・編集と同じ決まり）" -TestCases @(
        @{ label = "同じフォルダ"; folder = "C:\data\売上"; expected = "インデックス \[売上\] が既にあります" }
        @{ label = "中のフォルダ"; folder = "C:\data\売上\2024"; expected = "の中のフォルダです" }
        @{ label = "含むフォルダ"; folder = "C:\data"; expected = "があります" }
    ) {
        param ($label, $folder, $expected)
        testIndexImportInput $folder "新しい名前" $items | Should -Match $expected
    }

    It "同じ名前の行の元のフォルダは、上書きで置き換わるため断らない" {
        testIndexImportInput "C:\data\売上" "売上" $items | Should -Be ""
    }
}

Describe "getImportProgressView" -Tag Unit {
    It "<label>" -TestCases @(
        @{ label = "全体が分からない間は割合の分からない表示"; done = 0; total = 0; text = "インポートしています…"; detail = ""; indeterminate = $true; value = 0.0 }
        @{ label = "途中は割合と大きさ"; done = 1MB; total = 4MB; text = "インポートしています（25%）"; detail = "1.0 MB / 4.0 MB"; indeterminate = $false; value = 0.25 }
        @{ label = "展開が終わった後は割合の分からない表示"; done = 4MB; total = 4MB; text = "インポートを仕上げています…"; detail = ""; indeterminate = $true; value = 1.0 }
    ) {
        param ($label, $done, $total, $text, $detail, $indeterminate, $value)
        $view = getImportProgressView $done $total
        $view.Text | Should -Be $text
        $view.Detail | Should -Be $detail
        $view.Indeterminate | Should -Be $indeterminate
        $view.Value | Should -Be $value
    }
}

Describe "getImportIndexingConfirm" -Tag Unit {
    It "元のフォルダがあれば、importIndex が返した名前で尋ね、2-gram の索引のことを添える" {
        $confirm = getImportIndexingConfirm @{ Name = "営業(2)"; Enabled = $true }
        $confirm.Title | Should -Be "営業(2) をインポートしました"
        $confirm.Heading | Should -Be "高速化のためのインデックスを作成しますか？"
        $confirm.CreateText | Should -Be "作成する"
        $confirm.CancelText | Should -Be "あとで"
        $confirm.Hint | Should -Be "作成すると検索が速くなります。作成しなくても検索はできます（結果は同じ）。"
    }

    It "元のフォルダがこの PC に無ければ尋ねない" {
        getImportIndexingConfirm @{ Name = "営業"; Enabled = $false } | Should -BeNullOrEmpty
    }
}

Describe "getImportSkippedIndexingStatus" -Tag Unit {
    It "<label>" -TestCases @(
        @{ label = "Warnings が無い"; warnings = @(); expected = "インデックス [営業] をインポートしました（3 ファイル）。 次の更新までは、検索が遅くなることがあります。" }
        @{ label = "Warnings があれば間に入る"; warnings = @("注意です。"); expected = "インデックス [営業] をインポートしました（3 ファイル）。 注意です。 次の更新までは、検索が遅くなることがあります。" }
    ) {
        param ($label, $warnings, $expected)
        getImportSkippedIndexingStatus @{ Name = "営業"; Files = 3; Warnings = $warnings } | Should -Be $expected
    }
}

Describe "インポート後の画面に出す文言" -Tag Unit {
    It "内部の言葉（2-gram・索引）と、前の言い方（更新しますか・更新する）を画面に出さない" {
        $result = @{ Name = "営業"; Files = 1; Warnings = @(); Enabled = $true }
        $confirm = getImportIndexingConfirm $result
        $texts = @(
            (getImportProgressView 1MB 4MB).Text, (getImportProgressView 1MB 4MB).Detail, (getImportProgressView 4MB 4MB).Text,
            $confirm.Title, $confirm.Heading, $confirm.Hint, $confirm.CreateText, $confirm.CancelText, (getImportSkippedIndexingStatus $result)
        )
        foreach ($text in $texts) {
            $text | Should -Not -Match "2-gram|索引|更新しますか|更新する"
        }
    }
}
