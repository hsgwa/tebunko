# インデックス作成の段階ごとの時間と、取り込みの種類（拡張子）ごとの時間を足し合わせる（IngestTiming）。
# ファイルにも画面にも触らない集計。司令のスレッドで作り、段の区切りで Lap・AddStage を呼ぶ。
# 取り込みの 1 件の時間（抽出・コピー・置き場への移し）はワーカーが測り、結果に載せて司令が AddFile で足す。
# 段階: Startup（起動）・Crawl（クロール）・Ingest（取り込み。中のまとめは IngestFlush に別に出す）・FinishFlush（仕上げのまとめ）・Gram（gram 索引）・
#   Other（全体 - 上の合計）。StartupFlush（起動の中のまとめ）・IngestFlush（取り込みの中のまとめ）は内訳で、段の合計には足さない。
#   Wait（画面の返事を待った時間）は全体に含めない。時刻は Stopwatch だけを使う

class IngestTiming {
    # 段階（名前 → ミリ秒）
    hidden [hashtable]$Stages
    # 種類（拡張子。小文字）→ @{ Ok; Failed; Reroute; Postponed; ExtractMs; CopyMs; PublishMs; OkMs }
    hidden [hashtable]$Kinds
    hidden [System.Diagnostics.Stopwatch]$Watch
    hidden [double]$LapStartMs = 0
    hidden [double]$TotalMs = -1

    IngestTiming() {
        $this.Stages = @{}
        $this.Kinds = @{}
        $this.Watch = [System.Diagnostics.Stopwatch]::new()
    }

    [void] Start() {
        # 全体の計りを始め、区切りの印を今に置く
        $this.Watch.Restart()
        $this.LapStartMs = 0
    }

    [void] Mark() {
        # 区切りの印を今に置く（ここまでの時間はどの段にも足さない）
        $this.LapStartMs = $this.Watch.Elapsed.TotalMilliseconds
    }

    [void] Lap([string]$stage) {
        # 前の印から今までの時間を stage に足し、印を今に置く
        $now = $this.Watch.Elapsed.TotalMilliseconds
        $this.AddStage($stage, $now - $this.LapStartMs)
        $this.LapStartMs = $now
    }

    [void] AddStage([string]$stage, [double]$ms) {
        # 同じ段階が何度かあれば足す
        $current = 0.0
        if ($this.Stages.ContainsKey($stage)) {
            $current = [double]$this.Stages[$stage]
        }
        $this.Stages[$stage] = $current + [Math]::Max(0.0, $ms)
    }

    [void] Finish() {
        # 全体の時間を決める（画面の返事を待った時間は含めない）
        $this.Finish($this.Watch.Elapsed.TotalMilliseconds)
    }

    [void] Finish([double]$ms) {
        $wait = 0.0
        if ($this.Stages.ContainsKey("Wait")) {
            $wait = [double]$this.Stages["Wait"]
        }
        $this.TotalMs = [Math]::Max(0.0, $ms - $wait)
    }

    static [string] GetOutcome($result) {
        # 取り込みの結果（invokeIngestTask の戻り値）を、Ok・Reroute・Postponed（後回し。Postponed・Stopped）・Failed（TimedOut を含む残り）に分ける
        if ($result.Ok) { return "Ok" }
        if ($result.Reroute) { return "Reroute" }
        if ($result.Postponed -or $result.Stopped) { return "Postponed" }
        return "Failed"
    }

    static [string] GetKind([string]$relPath) {
        # 種類 = 拡張子（小文字）。無ければ "(なし)"
        $ext = [System.IO.Path]::GetExtension($relPath).ToLowerInvariant()
        if ($ext -eq "") { return "(なし)" }
        return $ext
    }

    hidden [hashtable] KindEntry([string]$kind) {
        if (!$this.Kinds.ContainsKey($kind)) {
            $this.Kinds[$kind] = @{ Ok = 0; Failed = 0; Reroute = 0; Postponed = 0; ExtractMs = 0.0; CopyMs = 0.0; PublishMs = 0.0; OkMs = 0.0 }
        }
        return $this.Kinds[$kind]
    }

    [void] AddFile([string]$kind, [string]$outcome, [double]$copyMs, [double]$extractMs, [double]$publishMs) {
        # 1 件（回し直した分は、回し直す前の分と後の分を別に）の時間を、その種類に足す
        $entry = $this.KindEntry($kind)
        $entry[$outcome] = [int]$entry[$outcome] + 1
        $entry.CopyMs += $copyMs
        $entry.ExtractMs += $extractMs
        $entry.PublishMs += $publishMs
        if ($outcome -eq "Ok") {
            $entry.OkMs += ($copyMs + $extractMs + $publishMs)
        }
    }

    [void] AddPostponed([string]$kind) {
        # 結果を受け取らずに後回しにした分（諦めたワーカー・渡さずに残した分）。件数だけ足し、時間は足さない
        $entry = $this.KindEntry($kind)
        $entry.Postponed = [int]$entry.Postponed + 1
    }

    hidden [double] Sum([string[]]$names) {
        $sum = 0.0
        foreach ($name in $names) {
            if ($this.Stages.ContainsKey($name)) {
                $sum += [double]$this.Stages[$name]
            }
        }
        return $sum
    }

    hidden [double] GetTotalMs() {
        if ($this.TotalMs -ge 0) {
            return $this.TotalMs
        }
        return $this.Sum(@("Startup", "Crawl", "Ingest", "FinishFlush", "Gram"))
    }

    [hashtable] ToHashtable() {
        # 受け渡しの口（StageTimings）に入れる形。段は秒、種類ごとは秒と件数、AvgMs は成功の 1 件の平均（ミリ秒。成功が 0 件なら $null）
        $total = $this.GetTotalMs()
        $parts = $this.Sum(@("Startup", "Crawl", "Ingest", "FinishFlush", "Gram"))
        $stageTable = [ordered]@{}
        foreach ($name in @("Startup", "Crawl", "Ingest", "FinishFlush", "Gram")) {
            $stageTable[$name] = [Math]::Round($this.Sum(@($name)) / 1000, 3)
        }
        $stageTable["StartupFlush"] = [Math]::Round($this.Sum(@("StartupFlush")) / 1000, 3)
        $stageTable["IngestFlush"] = [Math]::Round($this.Sum(@("IngestFlush")) / 1000, 3)
        $stageTable["Other"] = [Math]::Round([Math]::Max(0.0, $total - $parts) / 1000, 3)
        $stageTable["Total"] = [Math]::Round($total / 1000, 3)
        $stageTable["Wait"] = [Math]::Round($this.Sum(@("Wait")) / 1000, 3)
        $kindTable = [ordered]@{}
        foreach ($kind in @($this.Kinds.Keys | Sort-Object)) {
            $entry = $this.Kinds[$kind]
            $avg = $null
            if ([int]$entry.Ok -gt 0) {
                $avg = [Math]::Round([double]$entry.OkMs / [int]$entry.Ok, 1)
            }
            $kindTable[$kind] = [ordered]@{
                Count = [int]$entry.Ok + [int]$entry.Failed; Ok = [int]$entry.Ok; Failed = [int]$entry.Failed
                Reroute = [int]$entry.Reroute; Postponed = [int]$entry.Postponed
                ExtractSeconds = [Math]::Round([double]$entry.ExtractMs / 1000, 3); CopySeconds = [Math]::Round([double]$entry.CopyMs / 1000, 3)
                PublishSeconds = [Math]::Round([double]$entry.PublishMs / 1000, 3); AvgMs = $avg
            }
        }
        return @{ Stages = $stageTable; Kinds = $kindTable }
    }

    hidden [string] Sec([double]$seconds) {
        # 秒の表記（小数 1 桁。ロケールに依らない）。クラス自身の型を書かない（同じセッションで部品を読み直すと、型が別になって呼べなくなるため）
        return $seconds.ToString("0.0", [System.Globalization.CultureInfo]::InvariantCulture)
    }

    [string[]] FormatLogLines() {
        # インデックス作成ログに書く行（段階の 1 行と、種類ごとの行。種類は抽出の時間の長い順）
        $table = $this.ToHashtable()
        $s = $table.Stages
        $startup = $this.Sec($s.Startup); $startupFlush = $this.Sec($s.StartupFlush)
        $crawl = $this.Sec($s.Crawl); $ingest = $this.Sec($s.Ingest)
        $ingestFlush = $this.Sec($s.IngestFlush); $finishFlush = $this.Sec($s.FinishFlush)
        $gram = $this.Sec($s.Gram); $other = $this.Sec($s.Other); $total = $this.Sec($s.Total)
        $line = "段階ごとの時間（秒）: 起動 $startup（うちまとめ $startupFlush） / クロール $crawl / 取り込み $ingest（うちまとめ $ingestFlush） / 仕上げのまとめ $finishFlush / gram 索引 $gram / その他 $other / 全体 $total"
        if ([double]$s.Wait -gt 0) {
            $line += " / 画面の確認待ち $($this.Sec($s.Wait))（全体に含めない）"
        }
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add($line)
        $ordered = @($table.Kinds.Keys | Sort-Object @{ Expression = { -[double]$table.Kinds[$_].ExtractSeconds } }, @{ Expression = { $_ } })
        foreach ($kind in $ordered) {
            $k = $table.Kinds[$kind]
            $avg = "-"
            if ($null -ne $k.AvgMs) {
                $avg = [Math]::Round([double]$k.AvgMs, 0).ToString([System.Globalization.CultureInfo]::InvariantCulture) + " ms"
            }
            $extract = $this.Sec($k.ExtractSeconds); $copy = $this.Sec($k.CopySeconds); $publish = $this.Sec($k.PublishSeconds)
            $lines.Add("  $kind $($k.Count) 件: 抽出 $extract / コピー $copy / 置き場への移し $publish / 1 件 $avg（失敗 $($k.Failed)・回し直し $($k.Reroute)・後回し $($k.Postponed)）")
        }
        return $lines.ToArray()
    }
}
