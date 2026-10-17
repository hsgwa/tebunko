# 取り込みのジャーナル（store\journal\ingest.tsv）の読み書き（状態層。IngestJournal）。
# 書くのは司令（invokeIndexerBody）だけで、1 行ずつ追記して Flush する（取り込みのワーカーは書かない。書き手が 1 つなので排他は要らない。
# インデックス作成は鍵で 1 本しか動かない）。行の形と読み方は判断層（indexer_decide.ps1 の convertToIngestJournalLine・convertFromIngestJournal）。
# 保存先は、コンストラクタで受け取った場所から決める（メソッドの中で ${workspace} を読まない。docs/design/structure/classes.md）

class IngestJournal {
    [string]$Path
    hidden [System.IO.StreamWriter]$Writer

    IngestJournal([string]$path) {
        $this.Path = $path
    }

    [string[]] ReadPrevious() {
        # 前回のジャーナルの行を返す。無ければ空。最後の行が改行で終わっていなければ（書いている途中で強制終了した）、その行だけ捨てる
        if (!(Test-Path -LiteralPath $this.Path)) {
            return [string[]]@()
        }
        $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
        $stream = New-Object System.IO.FileStream($this.Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
        $reader = New-Object System.IO.StreamReader($stream, (New-Object System.Text.UTF8Encoding($true)))
        $text = ""
        try {
            $text = $reader.ReadToEnd()
        } finally {
            $reader.Dispose()
        }
        $end = $text.LastIndexOf("`n")
        if ($end -lt 0) {
            return [string[]]@()
        }
        $lines = New-Object System.Collections.Generic.List[string]
        foreach ($line in $text.Substring(0, $end).Split("`n")) {
            $line = $line.TrimEnd("`r")
            if ($line -ne "") {
                $lines.Add($line)
            }
        }
        return $lines.ToArray()
    }

    [void] AddRecord([string]$line) {
        # 1 記録を末尾に追記して Flush する（強制終了しても、書いた行は残る）。ファイルは書き始めで開き、Close まで開いたままにする
        if ($null -eq $this.Writer) {
            [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($this.Path)) | Out-Null
            $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
            $stream = New-Object System.IO.FileStream($this.Path, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, $share)
            $this.Writer = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($true)))
        }
        $this.Writer.Write($line + "`r`n")
        $this.Writer.Flush()
    }

    [void] Reset([string[]]$lines) {
        # ジャーナルを、渡した行だけにする（一時ファイルに書いて入れ替える。途中で止まっても前のジャーナルか新しいジャーナルのどちらかが残る）
        $this.Close()
        writeTextLinesAtomic $this.Path $lines
    }

    [void] Close() {
        if ($null -ne $this.Writer) {
            try { $this.Writer.Dispose() } catch {}
            $this.Writer = $null
        }
    }
}

function getLegacyIngestingPath {
    # 前の版が使っていた、取り込み中のファイルの記録（ワークスペース直下の ingesting.txt）の場所
    param (
        [string]$workspaceDir = $workspace.Dir
    )

    return Join-Path $workspaceDir "ingesting.txt"
}

function readLegacyIngesting {
    # 前の版の取り込み中のファイルの記録を読み、@{ RelPath = 相対パス; Count = 続けて取り込みを始めて終わらなかった回数 } の配列を返す
    # （始めの 1 回だけ読んで、ジャーナルの持ち越しに移す。移したら removeLegacyIngesting で消す）。1 行に 1 ファイル（"回数<TAB>相対パス"）。
    # 記録が無ければ空。壊れた行と、相対パスが使えない行（testIndexRelPath）は読み飛ばす
    param (
        [string]$path = (getLegacyIngestingPath)
    )

    $result = New-Object System.Collections.Generic.List[hashtable]
    $checkedNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::Ordinal)
    foreach ($line in @(readListFile $path)) {
        $fields = $line.Split("`t")
        $count = 0
        if ($fields.Count -ne 2 -or -not [int]::TryParse($fields[0], [ref]$count) -or $count -lt 1 -or $fields[1] -eq "") {
            continue
        }
        if (testIndexRelPath $fields[1] $false $checkedNames) {
            continue
        }
        $result.Add(@{ RelPath = $fields[1]; Count = $count })
    }
    return $result.ToArray()
}

function removeLegacyIngesting {
    param (
        [string]$path = (getLegacyIngestingPath)
    )

    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
    }
}
