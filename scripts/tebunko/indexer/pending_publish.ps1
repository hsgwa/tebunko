# 取り込んだ・元のファイルが無くなったフォルダを、本文インデックスに書き出してよいかを決める（PendingPublish）。
# 書き出し（publishIndexFolders）は司令（invokeIndexerBody）が呼ぶ。ここは数えて返すだけ（docs/design/structure/classes.md）。

class PendingPublish {
    # フォルダ（フルパス）→ そのフォルダで無くなった元のファイル名の集まり。まだ書き出していないもの
    hidden [System.Collections.Generic.Dictionary[string,object]]$Folders

    PendingPublish() {
        $this.Folders = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    }

    [void] MarkFolder([string]$folder) {
        # 前回のインデックス作成で本文インデックスに入れていないフォルダを、書き出し待ちにする（無くなったファイルの記録は無し）
        if (!$this.Folders.ContainsKey($folder)) {
            $this.Folders[$folder] = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        }
    }

    [void] Add([string]$relPath, [bool]$removed) {
        # 取り込んだ・無くなった元のファイルのフォルダを、書き出し待ちにする
        $folder = [System.IO.Path]::GetDirectoryName((getBookDir $relPath))
        if (!$this.Folders.ContainsKey($folder)) {
            $this.Folders[$folder] = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        }
        if ($removed) {
            [void]$this.Folders[$folder].Add([System.IO.Path]::GetFileName($relPath))
        }
    }

    [hashtable] TakeFlushable([string[]]$keepFolders) {
        # 書き出してよいフォルダ（keepFolders に無いもの）を取り出し、書き出し待ちから外して返す。
        # keepFolders は、まだ取り込みが続くため今回は書き出さないフォルダ（取り込み中のファイルがある・まだ渡していないファイルがある）
        $keep = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($folder in @($keepFolders)) {
            if ($folder) {
                [void]$keep.Add($folder)
            }
        }
        $flush = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($folder in @($this.Folders.Keys)) {
            if (!$keep.Contains($folder)) {
                $flush[$folder] = $this.Folders[$folder]
                [void]$this.Folders.Remove($folder)
            }
        }
        return $flush
    }
}
