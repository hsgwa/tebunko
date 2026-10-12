# Shift_JIS（CP932）で表せない文字を見つける（scripts\ と、作った単一 .ps1 の検査が共有する）。
# best-fit（近い字への置き換え）を使わないので、`〜`（U+301C）のように別の字に化ける字も見つかる。
function findNonCp932Chars {
    # 行ごとに、CP932 で表せない文字を "<行番号>: <文字> (U+XXXX)" の一覧で返す。無ければ空
    param ([string[]]$lines)

    $strict = [Text.Encoding]::GetEncoding(932, [Text.EncoderExceptionFallback]::new(), [Text.DecoderExceptionFallback]::new())
    $found = [System.Collections.Generic.List[string]]::new()
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -notmatch '[^\x00-\x7F]') { continue }
        foreach ($char in $lines[$i].ToCharArray()) {
            if ([int]$char -lt 0x80) { continue }
            try { [void]$strict.GetBytes([string]$char) }
            catch { $found.Add(("{0}: {1} (U+{2:X4})" -f ($i + 1), $char, [int]$char)) }
        }
    }
    return , $found.ToArray()
}
