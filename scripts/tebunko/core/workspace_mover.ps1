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
        # 戻り値: @{ Ok; Stage（失敗した段階。"remove"・"move"）; Message（失敗の理由）; Targets（"use" で使うインデックスの数）; Count（"move"・"reset" で移した数）}。
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
            $count = $this.MoveEntries($from, $to)
        } catch {
            return @{ Ok = $false; Stage = "move"; Message = $_.Exception.Message; Targets = 0; Count = 0 }
        }
        # 検索対象ツリーでチェックを外したフォルダも、移した先のインデックスに付け替える
        [void](moveSearchExcludes $from $to $this.SettingsPath)
        writeWorkspaceFolder $to $this.SettingsPath
        return @{ Ok = $true; Stage = ""; Message = ""; Targets = 0; Count = $count }
    }

    [int] MoveEntries([string]$from, [string]$to) {
        # ワークスペース from の中身（tebunko のファイル・フォルダ）を to へ移し、移した数を返す。ほかのファイルは移さない。
        # to に同じ名前があれば、何も移さずに例外にする。途中で移せなかったら、移した分を from へ戻してから例外にする
        # （どちらかのワークスペースに中身がそろった状態にし、半分ずつに分かれたままにしない）
        $conflicts = @($this.Conflicts($from, $to))
        if ($conflicts.Count -gt 0) {
            throw "「${to}」には、すでに $($conflicts -join '、') があります。空のフォルダを選んでください。"
        }
        $entries = @($this.ListEntries($from))
        [System.IO.Directory]::CreateDirectory((toLongPath $to)) | Out-Null
        $moved = New-Object System.Collections.Generic.List[string]
        try {
            foreach ($source in $entries) {
                $this.MoveEntry($source, (Join-Path $to ([System.IO.Path]::GetFileName($source))))
                $moved.Add($source)
            }
        } catch {
            $reason = $_.Exception.Message
            for ($i = $moved.Count - 1; $i -ge 0; $i--) {
                try {
                    $this.MoveEntry((Join-Path $to ([System.IO.Path]::GetFileName($moved[$i]))), $moved[$i])
                } catch {
                    # 戻せなかったものは、移した先に残る（例外のメッセージで両方の場所を伝える）
                }
            }
            throw "ワークスペースの中身を「${to}」へ移せませんでした（${reason}）。ファイルを開いているアプリを閉じてから、もう一度変えてください。中身は「${from}」に残しています。"
        }
        return $moved.Count
    }

    [int] RemoveEntries([string]$dir) {
        # ワークスペース dir の tebunko のファイル・フォルダ（Workspace.Entries）を削除し、削除した数を返す。ほかのファイルは消さない。
        # ワークスペースに選んだフォルダのインデックスを使わず、消して最初からやり直すときに使う
        $entries = @($this.ListEntries($dir))
        foreach ($entry in $entries) {
            $long = toLongPath $entry
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
        $map = getIndexNameMap (newWorkspace $dir.TrimEnd("\")).StatusFile
        if ($map.Count -eq 0) {
            return 0
        }
        writeTargetFolders @($map.Keys | ForEach-Object { [pscustomobject]@{ Name = $_; Path = $map[$_]; Enabled = $true } }) $this.SettingsPath
        return $map.Count
    }

    hidden [void] MoveEntry([string]$source, [string]$dest) {
        # ファイル・フォルダを 1 つ移す。同じドライブならそのまま移し、別のドライブのフォルダは写してから元を消す
        # （Directory.Move は別のドライブへ移せないため）。写している途中で失敗したら、写した分を消して例外にする
        $longSource = toLongPath $source
        $longDest = toLongPath $dest
        if (![System.IO.Directory]::Exists($longSource)) {
            # File.Move は別のドライブへも移せる
            [System.IO.File]::Move($longSource, $longDest)
            return
        }
        if ([System.IO.Path]::GetPathRoot($source) -eq [System.IO.Path]::GetPathRoot($dest)) {
            [System.IO.Directory]::Move($longSource, $longDest)
            return
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
        removeDirectoryRetry $source
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
