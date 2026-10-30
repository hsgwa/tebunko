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

function getIndexFileStateLabels {
    # 取り込んだファイルの一覧の「状態」の文言（Kind → 文言）。Kind は getIndexFileEntries と同じ（Done / Failed / Pending）
    return @{ Done = "取り込み済"; Failed = "失敗"; Pending = "未取り込み"; NoReason = "（原因は記録されていません）" }
}

function selectIndexFileRows {
    # 一覧に出す行を、絞り込みで選ぶ。filter が Failed なら失敗の行（Kind が Failed）だけ、それ以外は全部。行は元の順のまま
    param (
        [object[]]$rows,
        [string]$filter = "All"
    )

    $all = @($rows | Where-Object { $null -ne $_ })
    # 呼ぶ側は @() で受ける（1 件でも空でも配列にするため）
    if ($filter -ne "Failed") { return $all }
    return @($all | Where-Object { $_.Kind -eq "Failed" })
}

function getIndexFailedEntryView {
    # 詳細の「更新に失敗したファイル N 件［一覧で見る］」の 1 行。失敗が 1 件以上あるときだけ出す
    param (
        [int]$failedCount
    )

    if ($failedCount -le 0) { return @{ Shown = $false; Text = "" } }
    return @{ Shown = $true; Text = ("更新に失敗したファイル {0:#,0} 件" -f $failedCount) }
}

function getIndexFilesEmptyNotice {
    # 絞り込みの結果が 0 件のときの案内。失敗のみで 0 件なら「失敗したファイルはありません」、それ以外は空（見出しの要約を出す）
    param (
        [string]$filter,
        [int]$shownCount
    )

    if ($filter -eq "Failed" -and $shownCount -le 0) { return "失敗したファイルはありません" }
    return ""
}

function getIndexFilesHeading {
    # 取り込んだファイルの一覧の見出し（total は一覧の行の数。済・失敗・未取り込みの全部）。total が 0 のときは「まだ取り込んだファイルはありません」
    param (
        [int]$total,
        [int]$failed = 0,
        [int]$pending = 0
    )

    if ($total -le 0) { return "まだ取り込んだファイルはありません" }
    $text = "ファイル {0:#,0} 件" -f $total
    $notes = New-Object 'System.Collections.Generic.List[string]'
    if ($failed -gt 0) { $notes.Add(("失敗 {0:#,0} 件" -f $failed)) }
    if ($pending -gt 0) { $notes.Add(("未取り込み {0:#,0} 件" -f $pending)) }
    if ($notes.Count -gt 0) { $text += "（うち" + ($notes -join "・") + "）" }
    return $text
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
        $items,           # 今の一覧
        $leftovers = $null # 設定から外れたまま残してあるインデックス（getWorkspaceIndexFolders の結果。省くと調べない）
    )

    $newName = $name.Trim()
    $newPath = normalizeFolderPath $path
    if ($newName -ceq $item.Name -and $newPath -eq $item.Path) {
        return @{ Action = "None"; Message = ""; Name = $item.Name; Path = $item.Path }
    }
    $message = testIndexEditInput $path $newName $items $item $leftovers
    if ($message -ne "") {
        return @{ Action = "Invalid"; Message = $message; Name = $item.Name; Path = $item.Path }
    }
    return @{ Action = "Apply"; Message = ""; Name = $newName; Path = $newPath }
}

function getIndexDetailInfoRows {
    # 詳細の右の「インデックス情報」の @{ Label; Value } の並び。getIndexDetailView が使う（hasCount: 取り込んだ件数が分かるか）。
    # 一覧（［ファイルの一覧］）を開かなくても、失敗・未取り込みの件数が分かるようにする（取り込んだ分があるときだけ）
    param (
        $item,
        [bool]$hasCount
    )

    $rows = New-Object 'System.Collections.Generic.List[object]'
    [void]$rows.Add(@{ Label = "対象ファイル数"; Value = $(if ($hasCount) { "$($item.FileCountText) ファイル" } else { "－" }) })
    [void]$rows.Add(@{ Label = "最終更新"; Value = $(if ($item.LastIngestedText) { [string]$item.LastIngestedText } else { "－" }) })
    if ($hasCount) {
        [void]$rows.Add(@{ Label = "うち失敗"; Value = ("{0:#,0} 件" -f [int]$item.FailedCount) })
        [void]$rows.Add(@{ Label = "うち未取り込み"; Value = ("{0:#,0} 件" -f [int]$item.PendingCount) })
    }
    return , $rows.ToArray()
}

function getIndexDetailView {
    # インデックスの詳細のパネルに出す内容を返す。1 つ選んでいるときだけ中身を出す
    # （何も選んでいない・複数のときは見出しだけ。一覧の下の合計は別の部品）。
    #   items         : 選んでいる行。@{ Name; Path; Enabled; FolderStatus; FolderMissing（元のフォルダが無いと分かったか）; FolderChecked（有無を調べ終えたか。$false のときだけ「確認中」。省くと調べ終えたものとする）; IndexText; IndexLevel; IndexSub; FileCountText;
    #                   LastIngestedText; FailedCount・PendingCount（取り込み一覧の失敗・未取り込みの件数。省くと 0）}
    #   multiCount    : 2 以上なら、詳細は出さず「N 件を選択中」と案内だけにする（getIndexDetailMultiCount）
    #   running       : 更新の進み具合（@{ Ratio; Processed; Total; Failed; Eta; Current }）。更新中の行の「インデックス」の箱に出す。無ければ $null
    # 戻り値: @{ Title; Selected; Name; Path; Subfolders（サブフォルダを含めるか。設定の subfolders）; FolderStatus; FolderBrush（フォルダの状態の文の色のキー。ある = Ok／無い = Danger.Text／確認中・何も出さない = Ink.Muted）;
    #            FolderIcon（文の左に出すアイコン。ある = Check／無い = Close／確認中・何も出さない = 空。theme.xaml の Icon.<名前>）; Badge（@{ Text; Level }）; Updated; Count;
    #            Rows（右の「インデックス情報」の @{ Label; Value } の配列）; RowsKey;
    #            Multi（複数を選んでいるか）; Hint（Multi のときの案内）;
    #            Run（@{ Shown（更新中の棒・件数・ファイル名を出すか）; Value（0～1）; CountText; FileText }）;
    #            Error・ErrorHint（元のフォルダが無いときの理由と直し方。無ければ空。バッジは「エラー」の Ng になる） }
    param (
        [object[]]$items,
        [int]$multiCount = 0,
        $running = $null
    )

    $noRun = @{ Shown = $false; Value = 0.0; CountText = ""; FileText = "" }
    $none = @{
        Title = "インデックスの状態"; Selected = $false; Name = ""; Path = ""; Subfolders = $true; FolderStatus = ""; FolderMissing = $false; FolderBrush = "Ink.Muted"; FolderIcon = ""
        Badge = @{ Text = ""; Level = "None" }; Updated = ""; Count = ""; Rows = @(); RowsKey = ""
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

    $rows = getIndexDetailInfoRows $item $hasCount

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
    ) -join "`t"
    $rowsKey = $fixed + "`n" + (@($rows | ForEach-Object { "$($_.Label)`t$($_.Value)" }) -join "`n")
    return @{
        Title = "$($item.Name) - 詳細"; Selected = $true; Name = [string]$item.Name; Path = [string]$item.Path; Subfolders = ($item.Subfolders -ne $false)
        FolderStatus = [string]$item.FolderStatus; FolderMissing = [bool]$item.FolderMissing; FolderBrush = $folderBrush; FolderIcon = $folderIcon
        Badge = $badge
        Updated = $updated; Count = $count; Rows = @($rows); RowsKey = $rowsKey
        Multi = $false; Hint = ""; Run = $run; Error = $errorText; ErrorHint = $errorHint
    }
}
