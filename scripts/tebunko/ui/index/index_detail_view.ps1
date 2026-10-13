# ［インデックス管理］タブの判断: 詳細欄。
# 画面に触らないため、そのままテストできる（tests\tebunko\ui\index\index_detail_view.Tests.ps1）。

function getIndexDetailItem {
    # 詳細欄に出す行（詳細の名前・フォルダパスの欄と［変更…］、右クリックの［削除］［エクスポート…］の対象も同じ行。［アクション ▼］は別で、［更新］［エクスポート…］［削除…］はチェックの行、［インポート…］はチェックが要らない。可否は getIndexActionsEnabled が決める）。押した行があればそれ、
    # 無くてチェックが 1 件だけならその行、それ以外は $null。詳細を出す所と可否・対象を決める所で、この 1 つを使う
    param (
        [object]$selectedItem,   # 一覧で押した行（無ければ $null）
        [object[]]$checkedItems  # getIndexCheckedItems の結果
    )

    if ($null -ne $selectedItem) { return $selectedItem }
    $checked = @($checkedItems | Where-Object { $null -ne $_ })
    if ($checked.Count -eq 1) { return $checked[0] }
    return $null
}

function getIndexDetailMultiCount {
    # 詳細欄を「N 件を選択中」にするときの N（そうしないときは 0）。2 件以上チェックしているときだけ、その件数を返す。
    # 0 件・1 件のときは 0（詳細欄は押した行のまま）
    param (
        [int]$checked
    )

    if ($checked -ge 2) { return $checked }
    return 0
}

function getIndexDetailRowPlan {
    # 詳細の行の高さをどうするか。1 件も無いとき（empty）は詳細と境目を畳み（Hide）、1 件以上になったら戻す（Show）。
    # 空かどうかが変わっていないとき（別の行を選んだだけ）は何もしない（Keep）。境目をドラッグして変えた高さを、選び直しで戻さないため。
    # wasEmpty は前回の空かどうか（まだ一度も決めていなければ $null）。
    param (
        $wasEmpty,
        [bool]$empty
    )

    if ($null -ne $wasEmpty -and [bool]$wasEmpty -eq $empty) { return "Keep" }
    if ($empty) { return "Hide" }
    return "Show"
}

function getIndexDetailEditDecision {
    # 詳細の「基本設定」の名前・フォルダパスの欄を確定するときの判断。@{ Action; Message; Name; Path }
    #   Action: None（変わっていない）/ Invalid（直してほしい。Message に理由）/ Apply（Name・Path で反映する）
    # 名前は前後の空白を除き、フォルダパスは normalizeFolderPath で整えてから比べる（大文字・小文字だけの違いは名前の変更として扱う）
    param (
        [string]$name,
        [string]$path,
        $item,            # 直している行（Name・Path を持つ）
        $items            # 今の一覧
    )

    $newName = $name.Trim()
    $newPath = normalizeFolderPath $path
    if ($newName -ceq $item.Name -and $newPath -eq $item.Path) {
        return @{ Action = "None"; Message = ""; Name = $item.Name; Path = $item.Path }
    }
    $message = testIndexEditInput $path $newName $items $item
    if ($message -ne "") {
        return @{ Action = "Invalid"; Message = $message; Name = $item.Name; Path = $item.Path }
    }
    return @{ Action = "Apply"; Message = ""; Name = $newName; Path = $newPath }
}

function getIndexDetailView {
    # インデックスの詳細のパネルに出す内容を返す。1 つ選んでいるときだけ中身を出す
    # （何も選んでいない・複数のときは見出しだけ。一覧の下の合計は別の部品）。
    #   items         : 選んでいる行。@{ Name; Path; Enabled; FolderStatus; FolderMissing（元のフォルダが無いと分かったか）; FolderChecked（有無を調べ終えたか。$false のときだけ「確認中」。省くと調べ終えたものとする）; IndexText; IndexLevel; IndexSub; FileCountText;
    #                   LastIngestedText; FastText; FastLevel; FastToolTip; FastSteps; FastCheckedText }
    #   fastEntry     : getSystemIndexProgress の ByIndex のそのインデックスの値（@{ Folders; Waiting }）。無ければ $null
    #   multiCount    : 2 以上なら、詳細は出さず「N 件を選択中」と案内だけにする（getIndexDetailMultiCount）
    #   running       : 更新の進み具合（@{ Ratio; Processed; Total; Failed; Eta; Current }）。更新中の行の「インデックス」の箱に出す。無ければ $null
    # 戻り値: @{ Title; Selected; Name; Path; FolderStatus; FolderBrush（フォルダの状態の文の色のキー。ある = Ok／無い = Danger.Text／確認中・何も出さない = Ink.Muted）;
    #            FolderIcon（文の左に出すアイコン。ある = Check／無い = Close／確認中・何も出さない = 空。theme.xaml の Icon.<名前>）; Badge（@{ Text; Level }）; Updated; Count;
    #            Rows（右の「インデックス情報」の @{ Label; Value } の配列）; RowsKey;
    #            Fast（@{ Shown（反映の進みの棒を出すか）; Value（0～1）; Text; State; Level; Reason; Steps（対処の手順の箇条）; Checked; Note }）;
    #            Multi（複数を選んでいるか）; Hint（Multi のときの案内）;
    #            Run（@{ Shown（更新中の棒・件数・ファイル名を出すか）; Value（0～1）; CountText; FileText }）;
    #            Error・ErrorHint（元のフォルダが無いときの理由と直し方。無ければ空。バッジは「エラー」の Ng になる） }
    #   RowsKey は左右の表の中身から作る文字列。同じなら画面は行を置き直さない（高速検索の進みだけが変わったとき）
    param (
        [object[]]$items,
        $fastEntry = $null,
        [int]$multiCount = 0,
        $running = $null
    )

    $noRun = @{ Shown = $false; Value = 0.0; CountText = ""; FileText = "" }
    $emptyFast = @{ Shown = $false; Value = 0.0; Text = ""; State = ""; Level = "None"; Reason = ""; Steps = @(); Checked = ""; Note = "" }
    $none = @{
        Title = "インデックスの状態"; Selected = $false; Name = ""; Path = ""; FolderStatus = ""; FolderMissing = $false; FolderBrush = "Ink.Muted"; FolderIcon = ""
        Badge = @{ Text = ""; Level = "None" }; Updated = ""; Count = ""; Rows = @(); RowsKey = ""; Fast = $emptyFast
        Multi = $false; Hint = ""; Run = $noRun; Error = ""; ErrorHint = ""
    }
    if ($multiCount -ge 2) {
        $multi = $none.Clone()
        $multi.Title = "{0:#,0} 件を選択中" -f $multiCount
        $multi.Multi = $true
        $multi.Hint = "1 件だけ選ぶと、ここに詳細を表示します。"
        return $multi
    }
    $selected = @($items | Where-Object { $null -ne $_ })
    if ($selected.Count -ne 1) {
        return $none
    }
    $item = $selected[0]

    $hasCount = $item.FileCountText -and $item.FileCountText -ne "－"
    $count = if ($hasCount) { "$($item.FileCountText) ファイル" } else { "まだ更新していません" }
    $updated = if ($item.LastIngestedText) { "最終更新 $($item.LastIngestedText)" } else { "" }

    $rows = New-Object 'System.Collections.Generic.List[object]'
    [void]$rows.Add(@{ Label = "対象ファイル数"; Value = $(if ($hasCount) { "$($item.FileCountText) ファイル" } else { "－" }) })
    [void]$rows.Add(@{ Label = "最終更新"; Value = $(if ($item.LastIngestedText) { [string]$item.LastIngestedText } else { "－" }) })

    $fast = @{ Shown = $false; Value = 0.0; Text = ""; State = [string]$item.FastText; Level = [string]$item.FastLevel; Reason = ""; Steps = @(); Checked = [string]$item.FastCheckedText; Note = "" }
    if ($item.FastLevel -eq "Warn" -and $item.FastToolTip) {
        $fast.Reason = ([string]$item.FastToolTip -split "`n")[0]
        $fast.Steps = @($item.FastSteps)
    }
    if ($null -ne $fastEntry -and $fastEntry.Folders -gt 0) {
        $done = [Math]::Max(0, [Math]::Min([int]$fastEntry.Folders, [int]$fastEntry.Folders - [int]$fastEntry.Waiting))
        $fast.Shown = $true
        $fast.Value = $done / [double]$fastEntry.Folders
        $fast.Text = "反映済み $done / $($fastEntry.Folders) フォルダ"
        if ($done -lt [int]$fastEntry.Folders) {
            $fast.Note = "反映待ちのフォルダは通常の検索で調べます。検索結果は変わりませんが、時間がかかります。"
        }
    }
    $run = $noRun
    if ($item.IndexLevel -eq "Run") {
        $run = @{ Shown = $true; Value = 0.0; CountText = ""; FileText = "" }
        if ($null -ne $running) {
            $parts = New-Object 'System.Collections.Generic.List[string]'
            if ([int]$running.Total -gt 0) {
                $counts = "{0:#,0} / {1:#,0} 件" -f [int]$running.Processed, [int]$running.Total
                if ([int]$running.Failed -gt 0) {
                    $counts += "（失敗 {0:#,0} 件）" -f [int]$running.Failed
                }
                [void]$parts.Add($counts)
            }
            if ($running.Eta) {
                [void]$parts.Add([string]$running.Eta)
            }
            $run.Value = [Math]::Max(0.0, [Math]::Min(1.0, [double]$running.Ratio))
            $run.CountText = $parts -join "・"
            $run.FileText = if ($running.Current) { "読み取り中のファイル：$($running.Current)" } else { "" }
        }
    }
    # 元のフォルダが見つからないとき（更新中は更新の表示を優先する）は、バッジを「エラー」にして、理由と直し方を出す
    $badge = @{ Text = [string]$item.IndexText; Level = [string]$item.IndexLevel; Sub = [string]$item.IndexSub }
    $errorText = ""
    $errorHint = ""
    if ($item.FolderMissing -and $item.IndexLevel -ne "Run") {
        $badge = @{ Text = "エラー"; Level = "Ng"; Sub = "" }
        $errorText = "フォルダが見つかりません。"
        $errorHint = "フォルダが移動・削除されたか、ネットワークにつながっていない可能性があります。`nフォルダパスの［変更…］でフォルダを選び直してください。"
    }
    $folderBrush = if ($item.FolderMissing) { "Danger.Text" } elseif ($item.FolderChecked -eq $false) { "Ink.Muted" } else { "Ok" }
    $folderIcon = if ($item.FolderMissing) { "Close" } elseif ($item.FolderChecked -eq $false -or !$item.FolderStatus) { "" } else { "Check" }
    $fixed = @(
        $item.Name, $item.Path, $item.FolderStatus, $folderBrush, $folderIcon, $badge.Text, $badge.Level, $badge.Sub, $updated, $count
        $fast.State, $fast.Level, $fast.Reason, (@($fast.Steps) -join "|"), $fast.Checked
    ) -join "`t"
    $rowsKey = $fixed + "`n" + (@($rows | ForEach-Object { "$($_.Label)`t$($_.Value)" }) -join "`n")
    return @{
        Title = "$($item.Name) - 詳細"; Selected = $true; Name = [string]$item.Name; Path = [string]$item.Path
        FolderStatus = [string]$item.FolderStatus; FolderMissing = [bool]$item.FolderMissing; FolderBrush = $folderBrush; FolderIcon = $folderIcon
        Badge = $badge
        Updated = $updated; Count = $count; Rows = $rows.ToArray(); RowsKey = $rowsKey; Fast = $fast
        Multi = $false; Hint = ""; Run = $run; Error = $errorText; ErrorHint = $errorHint
    }
}
