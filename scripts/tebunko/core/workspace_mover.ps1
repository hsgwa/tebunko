# ワークスペースの切り替え（中身を移す・消して最初からやり直す・選んだ先のインデックスを使う・searchExcludes の付け替え・保存・失敗したら戻す）。
# 画面（ui\settings\settings.ps1 の applyWorkspace・startWorkspaceMove）は、確かめて利用者に選ばせたあと ChangeWorkspace を呼び、結果を出して画面を切り替えるだけにする。
# ListEntries などの確かめは画面のスレッドで、ChangeWorkspace は裏のスレッド（startWorkspaceMove の startJob）で使う。切り替えの前に、画面は本文インデックスのファイルを読んでいる検索を止める（移せないため）。
#
# クラスの本体では $script: 変数・ほかのクラスの型・呼び出し元の変数を使わない（docs/design/structure/classes.md）。
# 設定ファイルの場所はコンストラクタで受け取る。ワークスペースの中の場所は Workspace（workspace.ps1）から組み立てる。
# searchExcludes の付け替えは、IndexCatalog と同じ部品（search_excludes.ps1 の moveSearchExcludes）を通す。
# メソッドの中のローカル変数は、プロパティと大文字・小文字だけの違いにしない
class WorkspaceMover {
    hidden [string]$SettingsPath

    WorkspaceMover([string]$settingsPath) {
        $this.SettingsPath = $settingsPath
    }

    [string[]] ListEntries([string]$dir) {
        # ワークスペース dir にある、tebunko のファイル・フォルダ（Workspace.Entries のうち、あるもの）のフルパスを返す
        return @((newWorkspace $dir).Entries() | Where-Object {
            $long = toLongPath $_
            [System.IO.Directory]::Exists($long) -or [System.IO.File]::Exists($long)
        })
    }

    [string[]] Conflicts([string]$from, [string]$to) {
        # ワークスペース from の中身を to へ移すとき、to に同じ名前が既にあるもの（移せないもの）の名前を返す
        return @($this.ListEntries($from) | ForEach-Object { [System.IO.Path]::GetFileName($_) } | Where-Object {
            $long = toLongPath (Join-Path $to $_)
            [System.IO.Directory]::Exists($long) -or [System.IO.File]::Exists($long)
        })
    }

    [hashtable] ChangeWorkspace([string]$mode, [string]$from, [string]$to) {
        # ワークスペースを from から to に変える（設定に保存する）。mode:
        #   "use"   … to にあるインデックスを使う。from の中身は移さず、インデックスの一覧を to のものにする
        #   "move"  … from の中身を to へ移す
        #   "reset" … to の tebunko のファイル・フォルダを消してから、from の中身を移す
        # 戻り値: @{ Ok; Stage（失敗した段階。"remove"・"move"）; Message（失敗の理由）; Targets（"use" で使うインデックスの数）; Count（"move"・"reset" で移した数）; Leftovers（別のドライブへ写したあと元を消せずに残った項目の場所）}。
        # 失敗したときは設定を変えない（移せなかったら、移した分は from に戻してある）
        if ($mode -eq "use") {
            # 共有されたワークスペースなどを使う。今の中身は移さず、インデックスの一覧をこのワークスペースのものにする
            $targets = $this.UseTargets($to)
            writeWorkspaceFolder $to $this.SettingsPath
            return @{ Ok = $true; Stage = ""; Message = ""; Targets = $targets; Count = 0 }
        }
        if ($mode -eq "reset") {
            try {
                [void]$this.RemoveEntries($to)
            } catch {
                return @{ Ok = $false; Stage = "remove"; Message = $_.Exception.Message; Targets = 0; Count = 0 }
            }
        }
        try {
            $result = $this.MoveEntries($from, $to)
        } catch {
            return @{ Ok = $false; Stage = "move"; Message = $_.Exception.Message; Targets = 0; Count = 0; Leftovers = @() }
        }
        # 検索対象ツリーでチェックを外したフォルダも、移した先のインデックスに付け替える
        [void](moveSearchExcludes $from $to $this.SettingsPath)
        writeWorkspaceFolder $to $this.SettingsPath
        return @{ Ok = $true; Stage = ""; Message = ""; Targets = 0; Count = $result.Count; Leftovers = @($result.Leftovers) }
    }

    [hashtable] MoveEntries([string]$from, [string]$to) {
        return $this.MoveEntries($from, $to, 5, 200)
    }

    [hashtable] MoveEntries([string]$from, [string]$to, [int]$tries, [int]$waitMilliseconds) {
        # ワークスペース from の中身（tebunko のファイル・フォルダ）を to へ移す。ほかのファイルは移さない。
        # 戻り値: @{ Count（移した数）; Leftovers（別のドライブへ写したあと、元を消せずに残った元の場所の一覧）}
        # フォルダは、ほかの処理が一時的に掴んでいても、待つ時間を倍にしながら tries 回まで試して移す（moveDirectoryRetry）。
        # to に同じ名前があれば、何も移さずに例外にする。途中で移せなかったら、移した分を from へ戻してから例外にする
        # （どちらかのワークスペースに中身がそろった状態にし、半分ずつに分かれたままにしない）。
        # 戻せなかった項目があれば、その名前と、残っている場所（to）を例外の文に書く
        $conflicts = @($this.Conflicts($from, $to))
        if ($conflicts.Count -gt 0) {
            throw "「${to}」には、すでに $($conflicts -join '、') があります。空のフォルダを選んでください。"
        }
        $entries = @($this.ListEntries($from))
        [System.IO.Directory]::CreateDirectory((toLongPath $to)) | Out-Null
        $moved = New-Object System.Collections.Generic.List[string]
        $leftovers = New-Object System.Collections.Generic.List[string]
        try {
            foreach ($source in $entries) {
                $left = $this.MoveEntry($source, (Join-Path $to ([System.IO.Path]::GetFileName($source))), $tries, $waitMilliseconds)
                $moved.Add($source)
                if ($left) {
                    $leftovers.Add($left)
                }
            }
        } catch {
            $reason = $_.Exception.Message
            $unrestored = New-Object System.Collections.Generic.List[string]
            for ($i = $moved.Count - 1; $i -ge 0; $i--) {
                try {
                    [void]$this.MoveEntry((Join-Path $to ([System.IO.Path]::GetFileName($moved[$i]))), $moved[$i], $tries, $waitMilliseconds)
                } catch {
                    # 戻せなかったものは、移した先にある。握りつぶさずに控え、例外の文で伝える
                    $unrestored.Add([System.IO.Path]::GetFileName($moved[$i]))
                }
            }
            $where = if ($unrestored.Count -gt 0) {
                "元へ戻せなかった項目があります（$($unrestored -join '、')）。それは「${to}」にあります。ほかの中身は「${from}」にあります。"
            } else {
                "中身は「${from}」に残しています。"
            }
            throw "ワークスペースの中身を「${to}」へ移せませんでした（${reason}）。ファイルを開いているアプリを閉じてから、もう一度変えてください。${where}"
        }
        return @{ Count = $moved.Count; Leftovers = @($leftovers) }
    }

    [int] RemoveEntries([string]$dir) {
        # ワークスペース dir の tebunko のファイル・フォルダ（Workspace.Entries）を削除し、削除した数を返す。ほかのファイルは消さない。
        # ワークスペースに選んだフォルダのインデックスを使わず、消して最初からやり直すときに使う
        $entries = @($this.ListEntries($dir))
        foreach ($entry in $entries) {
            $long = toLongPath $entry
            # フォルダかファイルかを分けるために Exists が要る。有無が分からないものは、後ろの File.Delete が例外にする
            if ([System.IO.Directory]::Exists($long)) {
                removeDirectoryRetry $entry
            } else {
                [System.IO.File]::Delete($long)
            }
        }
        return $entries.Count
    }

    [int] UseTargets([string]$dir) {
        # ワークスペース dir の取り込み一覧にあるクロール対象フォルダを、インデックスの一覧（設定の targetFolders）にし、その数を返す。
        # ほかの人が作ったワークスペースを使うとき、一覧をそのワークスペースに合わせる（合わせないままインデックス作成をすると、
        # 一覧に無いインデックスは、インデックス更新の確認に「削除予定」として出る。findDroppedIndexes）。取り込み一覧が無ければ一覧は変えない
        # 前の版のワークスペースは取り込み一覧がワークスペース直下（OldStatusFile）にあるため、そちらも読む。
        # 移行の途中（まだ作り直していないインデックスがある）は両方にあるので、新しい方を先に、無い名前を前の版から足す
        $target = newWorkspace $dir.TrimEnd("\")
        $map = [ordered]@{}
        foreach ($statusPath in @($target.StatusFile, $target.OldStatusFile)) {
            $part = getIndexNameMap $statusPath
            foreach ($name in $part.Keys) {
                if (!$map.Contains($name)) {
                    $map[$name] = $part[$name]
                }
            }
        }
        if ($map.Count -eq 0) {
            return 0
        }
        writeTargetFolders @($map.Keys | ForEach-Object { [pscustomobject]@{ Name = $_; Path = $map[$_]; Enabled = $true } }) $this.SettingsPath
        return $map.Count
    }

    hidden [string] MoveEntry([string]$source, [string]$dest, [int]$tries, [int]$waitMilliseconds) {
        # ファイル・フォルダを 1 つ移す。同じドライブならそのまま移し（フォルダは moveDirectoryRetry で試し直す）、
        # 別のドライブのフォルダは写してから元を消す（Directory.Move は別のドライブへ移せないため）。
        # 写している途中で失敗したら、写した分を消して例外にする。
        # 写し終えたあとに元を消せなければ、移し先を正として例外にせず、残った元の場所を返す（消せたら空の文字）
        $longSource = toLongPath $source
        $longDest = toLongPath $dest
        if (![System.IO.Directory]::Exists($longSource)) {
            # File.Move は別のドライブへも移せる
            [System.IO.File]::Move($longSource, $longDest)
            return ""
        }
        if (testSameDriveRoot $source $dest) {
            moveDirectoryRetry $source $dest $tries $waitMilliseconds
            return ""
        }
        try {
            $this.CopyTree($longSource, $longDest)
        } catch {
            $original = $_
            try {
                removeDirectoryRetry $dest
            } catch {
                # 片付けに失敗しても、元の失敗の理由を隠さない
            }
            throw $original
        }
        try {
            removeDirectoryRetry $source
        } catch {
            return $source
        }
        return ""
    }

    hidden [void] CopyTree([string]$source, [string]$dest) {
        # フォルダを中身ごと写す（\\?\ 付きのパスを受け取る）
        [System.IO.Directory]::CreateDirectory($dest) | Out-Null
        foreach ($file in [System.IO.Directory]::GetFiles($source)) {
            [System.IO.File]::Copy($file, [System.IO.Path]::Combine($dest, [System.IO.Path]::GetFileName($file)))
        }
        foreach ($sub in [System.IO.Directory]::GetDirectories($source)) {
            $this.CopyTree($sub, [System.IO.Path]::Combine($dest, [System.IO.Path]::GetFileName($sub)))
        }
    }
}
