# インデックスの改名・削除・名前の割り当て・一覧と元のフォルダの記録の保存。
# インデックス 1 つの情報は複数の保存先に分かれて置かれている: 設定の targetFolders、取り込み一覧、インデックスのフォルダ（store\content\<名前>）、目録（store\format.tsv）の行と
# 元のフォルダの記録（source_folder.txt）、設定の searchExcludes。
# このクラスは、改名・削除でそれらを書き換える手順と順番、途中で失敗したときの扱い（止める・残す）だけを持つ。一覧（targetFolders）と元のフォルダの記録の書き直しは、画面が SaveTargets・WriteSourceFile（改名では RenameIndexWithSource）で呼ぶ。保存先の形式は知らず、
# 保存先ごとの読み書きは既存の関数（searchExcludes は core\search_excludes.ps1）と、作る側から渡された StatusLedger に任せる。
# 画面は呼んで結果を出すだけにする（画面の手順をここに引き寄せる）。
#
# クラスの本体では $script: 変数・ほかのクラスの型・呼び出し元の変数を使わない（docs/design/structure/classes.md）。
# 場所は、コンストラクタで受け取る workspace（Workspace のインスタンス）と設定ファイルの場所で決める。
# 取り込み一覧は ledger（StatusLedger のインスタンス。作る側が作って渡す）を通して書く。
# 別のスレッドへはインスタンスを渡さず、ワークスペースの場所（文字列）を渡して newIndexCatalog で作り直す。
# メソッドの中のローカル変数は、プロパティと大文字・小文字だけの違いにしない
class IndexCatalog {
    hidden $Workspace
    hidden [string]$SettingsPath
    hidden $Ledger

    IndexCatalog($workspace, [string]$settingsPath, $ledger) {
        $this.Workspace = $workspace
        $this.SettingsPath = $settingsPath
        $this.Ledger = $ledger
    }

    [void] AssertReachable() {
        # インデックスの削除・名前の変更の裏の仕事の先頭で呼ぶ。ワークスペースに届かないときは例外にする
        # （届かないのに「何もせず成功」にして、設定だけ変えてしまわないため）
        assertWorkspaceReachable $this.Workspace.IndexDir
    }

    [void] RenameIndex([string]$oldName, [string]$newName) {
        # インデックス名を変える（画面の詳細の名前の欄）。インデックスのフォルダ（store\content\<名前>）を改名し、
        # 取り込み一覧の記録も書き換えるため、名前を変えてもインデックスは作り直さない。
        # インデックス作成中は呼ばない（画面はインデックス作成中この操作を無効にする）。
        # フォルダの改名か取り込み一覧の書き換えに失敗したら例外にする（gram 索引と searchExcludes の整理は、失敗しても止めない）
        $dir = $this.Workspace.IndexDir
        # 大文字・小文字だけの変更も改名するため、同じ名前かは大文字・小文字を区別して比べる（-ceq）
        if ($oldName -eq "" -or $newName -eq "" -or $oldName -ceq $newName) {
            return
        }

        # 移す元・先の名前は、置き場（IndexDir）の外へ出ないものだけ使う。何も動かさずに例外にする。
        # 新しい名前は、インデックス名の決まり（testIndexName。予約名も拒む）に合うものだけ使う。
        # 元の名前は、既にある名前なので、予約名（aux.backup など）は通す（testStoredIndexName）。
        # 最後が . のような、決まりに合わない前の版の名前は、設定を読むときに空になって届かず、取り込み一覧の読み込み（LegacyFolders）と、
        # 設定から外れたインデックスの削除の流れが受け持つ（最後が . のフォルダは、Test-Path とフォルダ操作で別のフォルダを見るため、ここでは扱わない）
        $from = Join-Path $dir $oldName
        $to   = Join-Path $dir $newName
        if ((testStoredIndexName $oldName) -ne "" -or !(testPathUnder $from $dir)) {
            throw "元のインデックス名が使えないため、名前を変えられません: ${oldName}"
        }
        if ((testIndexName $newName @()) -ne "" -or !(testPathUnder $to $dir)) {
            throw "新しいインデックス名が使えないため、名前を変えられません: ${newName}"
        }
        # 元のフォルダは Directory.Exists で見る（aux.backup のような予約名のフォルダは、Test-Path では無いことになるため）。
        # ただし testPathUnder は GetFullPath に頼るため、最後の区切りをデバイス名として解く古い版の Windows では、
        # 上の確かめを通らず、何も動かさずに例外になる
        if ([System.IO.Directory]::Exists((toLongPath $from))) {
            if ([string]::Equals($oldName, $newName, [System.StringComparison]::OrdinalIgnoreCase)) {
                # 大文字・小文字だけを変える場合は、そのままでは改名できないため一時名を経由する
                $tmp = Join-Path $dir "${oldName}_rename_$([System.Diagnostics.Process]::GetCurrentProcess().Id)"
                # 移すのは、置き場（IndexDir）の下で途中にリンクが無いときだけ（moveDirectoryUnder）。偽なら何も動かさずに例外にする
                if (!(moveDirectoryUnder $from $tmp $dir) -or !(moveDirectoryUnder $tmp $to $dir)) {
                    throw "置き場の外か、途中にリンクがあるため、名前を変えられません: ${oldName}"
                }
            } else {
                if ([System.IO.Directory]::Exists((toLongPath $to)) -or [System.IO.File]::Exists((toLongPath $to))) {
                    throw "「${newName}」のフォルダが既にあるため、名前を変えられません: ${to}"
                }
                if (!(moveDirectoryUnder $from $to $dir)) {
                    throw "置き場の外か、途中にリンクがあるため、名前を変えられません: ${oldName}"
                }
            }
        }
        $this.Ledger.RenameIndexName($oldName, $newName)
        # 古い置き場（作り直す前の形式）の同じ名前のフォルダと、目録の行も改名に合わせる。目録を書き換えられない（新しい版）ときは行を残す
        renameOldIndexFolder $oldName $newName $this.Workspace
        [void](renameIndexFormatRow $oldName $newName $this.Workspace.FormatFile)
        # 検索用の索引（gram 索引）は古い名前の分を消す。次のインデックス作成で作り直す
        removeGramIndexOfWorkspace $oldName $dir
        # 外したフォルダの記録（searchExcludes）は旧名・新名の下のものを消す（付け替えず、検索対象に戻す）
        removeSearchExcludesUnder $from $this.SettingsPath
        removeSearchExcludesUnder $to $this.SettingsPath
    }

    [hashtable] RenameIndexWithSource([string]$oldName, [string]$newName, $folders) {
        # 名前を変え（RenameIndex）、続けて元のフォルダの記録（source_folder.txt）を書き直す。
        # 書き直しに失敗しても、改名は済んでいるので改名の失敗にはしない。書けなかった理由を SourceFolderError で返す（成功なら空）。
        # 改名の失敗は、RenameIndex の例外のまま呼び出し元へ出る
        $this.RenameIndex($oldName, $newName)
        $sourceError = ""
        try {
            $this.WriteSourceFile($folders)
        } catch {
            $sourceError = $_.Exception.Message
        }
        return @{ SourceFolderError = $sourceError }
    }

    [void] RemoveIndex([string]$name) {
        # インデックスを削除する（画面の［削除］）。インデックスのフォルダ（store\content\<名前>）と、取り込み一覧の記録を削除する。
        # インデックス作成中は呼ばない（画面はインデックス作成中この操作を無効にする）。
        # フォルダを消せなければ例外にして、取り込み一覧の記録は残す
        $dir = $this.Workspace.IndexDir
        if ($name -eq "") {
            return
        }
        $target = Join-Path $dir $name
        if ((testStoredIndexName $name) -ne "" -or !(testPathUnder $target $dir)) {
            # 名前は取り込み一覧（共有のワークスペースで書き換えられうる）から来ることもある。置き場の外は消さず、
            # 例外にして取り込み一覧の記録を残す（フォルダを消せなかったときと同じ）
            throw "インデックス名が使えないため、削除できません: ${name}"
        } else {
            # 中に長いパス（260文字超）のTSVがあっても削除でき、中のリンク（ジャンクション・シンボリックリンク）の先は消さない
            # （removeDirectoryRetry。フォルダが無ければ何もしない。aux.backup のような予約名のフォルダも、Test-Path ではなくここで見つけて消す。
            # ただし testPathUnder は GetFullPath に頼るため、最後の区切りをデバイス名として解く古い版の Windows では、
            # 上の確かめを通らず、何も消さずに例外になる）。
            # 消せなければ例外にして、取り込み一覧の記録を残す（フォルダが残ったまま記録だけ消えないように）
            if (!(removeDirectoryUnder $target $dir)) {
                throw "置き場の外か、途中にリンクがあるため、削除できませんでした: ${name}"
            }
        }
        # 古い置き場（作り直す前の形式）の同じ名前のフォルダと、目録の行も消す
        $oldTarget = Join-Path $this.Workspace.OldIndexDir $name
        if (testPathUnder $oldTarget $this.Workspace.OldIndexDir) {
            if (!(removeDirectoryUnder $oldTarget $this.Workspace.OldIndexDir)) {
                throw "置き場の外か、途中にリンクがあるため、削除できませんでした: ${name}"
            }
        }
        [void](removeIndexFormatRow $name $this.Workspace.FormatFile)
        $this.Ledger.RemoveIndexName($name)
        removeGramIndexOfWorkspace $name $dir
        # 外したフォルダの記録（searchExcludes）も消す
        removeSearchExcludesUnder $target $this.SettingsPath
    }

    [object[]] RemoveIndexes([string[]]$names) {
        # 選んだインデックスをまとめて削除する（画面の［削除］）。1 つずつ RemoveIndex を呼び、途中で 1 つ失敗しても残りを続ける。
        # 戻り値は名前ごとの結果の配列（@{ Name; Ok; Reason }）。Reason は失敗したときの理由（成功なら ""）。
        # 画面のスレッドでは呼ばず、別スレッドの仕事の中で呼ぶ
        $results = New-Object System.Collections.Generic.List[object]
        foreach ($name in @($names)) {
            if ([string]::IsNullOrEmpty($name)) {
                $results.Add([pscustomobject]@{ Name = [string]$name; Ok = $false; Reason = "インデックス名が空です" })
                continue
            }
            try {
                $this.RemoveIndex($name)
                $results.Add([pscustomobject]@{ Name = $name; Ok = $true; Reason = "" })
            } catch {
                $results.Add([pscustomobject]@{ Name = $name; Ok = $false; Reason = $_.Exception.Message })
            }
        }
        return $results.ToArray()
    }

    [object[]] AssignAndSaveNames($folders) {
        # 名前の決まっていないインデックス（設定ファイルを直接書き換えた場合など）に名前を割り当てて、設定に保存する。保存した一覧を返す。
        # 取り込み一覧の読み込みは排他の外で済ませ、設定の読み直しと書き込みだけを saveAssignedIndexNames が排他の中で行う
        $status = $this.Ledger.ReadStatus()
        $assigned = @(assignIndexNames $folders $status.Folders)
        return @(saveAssignedIndexNames $assigned $this.SettingsPath)
    }

    [object[]] StoreAssignedNames($assigned) {
        # 割り当てた名前（assignIndexNames の結果）を設定に保存し、保存した一覧（@{ Name; Path; Enabled }）を返す。
        # インデクサが、司令のスレッドで作ったこのクラスの経由で名前を保存する
        return @(saveAssignedIndexNames $assigned $this.SettingsPath)
    }

    [void] SaveTargets($folders) {
        # インデックスの一覧（@{ Name; Path; Enabled } の配列）を設定に書く
        writeTargetFolders $folders $this.SettingsPath
    }

    [void] WriteSourceFile($folders) {
        # インデックスのフォルダの source_folder.txt を、一覧（@{ Name; Path } の配列）に合わせて書き直す（インデックスが無いフォルダには書かない）
        writeSourceFolderFile $folders $this.Workspace.IndexDir
    }
}

function newIndexCatalog {
    # ワークスペースの場所（Dir。文字列）から IndexCatalog を作る。取り込み一覧を書く StatusLedger もここで作って渡す。
    # 画面のスレッドでも、別のスレッドの仕事の中でも使う（スレッドをまたいでインスタンスを渡さず、場所を渡して作り直すため）
    param (
        [string]$workspaceDir,
        [string]$settingsPath = ${settingsFile}
    )

    $target = [Workspace]::new($workspaceDir)
    return [IndexCatalog]::new($target, $settingsPath, [StatusLedger]::new($target))
}
