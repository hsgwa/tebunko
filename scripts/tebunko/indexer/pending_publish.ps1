# 取り込んだ・元のファイルが無くなったフォルダを、本文インデックスに書き出してよいかを決める（PendingPublish）。
# フォルダごとの取り込み中の数（Busy）とまだ渡していない数（Pending）を数え、どちらも 0 になったフォルダを
# 書き出してよいフォルダとして返す（TakeFlushable）。取り込みの前に無くなったフォルダ（Skip）・後回し（Complete。
# 成功・失敗と同じ扱い）も、この数に反映する。書き出し（publishIndexFolders）は司令（invokeIndexerBody）が呼ぶ。

class PendingPublish {
    # フォルダ（フルパス）→ そのフォルダで無くなった元のファイル名の集まり。まだ書き出していないもの
    hidden [System.Collections.Generic.Dictionary[string,object]]$Folders
    # フォルダごとの、取り込み中の数（Busy）・まだ渡していない数（Pending）。どちらも 0 で書き出してよい
    hidden [System.Collections.Generic.Dictionary[string,int]]$Busy
    hidden [System.Collections.Generic.Dictionary[string,int]]$Pending

    PendingPublish() {
        $this.Folders = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
        $this.Busy = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([System.StringComparer]::OrdinalIgnoreCase)
        $this.Pending = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([System.StringComparer]::OrdinalIgnoreCase)
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

    [void] AddPending([string]$folder) {
        # 取り込み対象を数えるとき、そのファイルの分だけ「まだ渡していない数」を1つ足す
        $this.Pending[$folder] = [int]$this.Pending[$folder] + 1
    }

    [void] Dispatch([string]$folder) {
        # ファイルを1つ、取り込みのスレッド（または司令のスレッド）に渡す（渡していない→取り込み中）
        $this.Pending[$folder] = [int]$this.Pending[$folder] - 1
        $this.Busy[$folder] = [int]$this.Busy[$folder] + 1
    }

    [void] Skip([string]$folder) {
        # ファイルを1つ、渡さずに済ませる（取り込みの直前に元のファイルが無くなった等。取り込み中にはしない）
        $this.Pending[$folder] = [int]$this.Pending[$folder] - 1
    }

    [void] Complete([string]$folder) {
        # 取り込み中のファイルが1つ終わる（成功・失敗・後回しのどれでも）
        $this.Busy[$folder] = [int]$this.Busy[$folder] - 1
    }

    [System.Collections.Generic.Dictionary[string,object]] TakeFlushable() {
        # 取り込み中・まだ渡していないファイルがどちらも無いフォルダを取り出し、書き出し待ちから外して返す
        $keep = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($folder in @($this.Busy.Keys | Where-Object { $this.Busy[$_] -gt 0 })) {
            [void]$keep.Add($folder)
        }
        foreach ($folder in @($this.Pending.Keys | Where-Object { $this.Pending[$_] -gt 0 })) {
            [void]$keep.Add($folder)
        }
        return $this.take($keep)
    }

    [System.Collections.Generic.Dictionary[string,object]] TakeAll() {
        # 取り込み中・まだ渡していない数を見ずに、書き出し待ちのフォルダをすべて取り出す
        # （インデックス作成の終わり。中止・続けられないエラーでも、取り込んだ分は残さず書き出す）
        return $this.take((New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)))
    }

    hidden [System.Collections.Generic.Dictionary[string,object]] take($keep) {
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
