# ボタンの中身（字面・アイコン）の位置を、描いた絵の画素から測る（tests\gui の見た目のテストと、tests\shared の部品のテストが使う）。
# レイアウトの数（ActualWidth など）ではなく、実際に描かれた「インク」（面の色と違う画素）の範囲を測る。
# 日本語のフォントは字面が行の枠の上か下に寄るので、枠の数だけを見ても、見た目のずれは分からない。

# Root を Width x Height で並べて描き、Target の範囲を枠（Frame。省略すると Target 自身）として、
# 枠の中にあるインクのまとまり（アイコン・文字。5 画素以上の横の空きで分ける）ごとに、
# 枠の中心からのずれ（Dx・Dy。右・下が正。画素）を返す。Groups の各要素は @{ Dx; Dy; Width; Height }
function measureInk {
    param (
        [System.Windows.FrameworkElement]$Root,
        [System.Windows.FrameworkElement]$Target,
        [System.Windows.FrameworkElement]$Frame = $null,
        [int]$Width = 900,
        [int]$Height = 600
    )

    if ($null -eq $Frame) { $Frame = $Target }
    # 本物の窓（tebunko.xaml）と同じ描き方にする（画素にそろえる・Display の文字の描き方）
    $Root.UseLayoutRounding = $true
    $Root.SnapsToDevicePixels = $true
    [System.Windows.Media.TextOptions]::SetTextFormattingMode($Root, [System.Windows.Media.TextFormattingMode]::Display)
    [System.Windows.Media.TextOptions]::SetTextRenderingMode($Root, [System.Windows.Media.TextRenderingMode]::ClearType)
    $Root.Measure((New-Object System.Windows.Size($Width, $Height)))
    $Root.Arrange((New-Object System.Windows.Rect(0, 0, $Width, $Height)))
    $Root.UpdateLayout()

    $bitmap = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($Width, $Height, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
    $bitmap.Render($Root)
    $stride = $Width * 4
    $bytes = New-Object byte[] ($stride * $Height)
    $bitmap.CopyPixels($bytes, $stride, 0)

    $bounds = $Frame.TransformToAncestor($Root).TransformBounds((New-Object System.Windows.Rect(0, 0, $Frame.ActualWidth, $Frame.ActualHeight)))
    $x0 = [int][Math]::Round($bounds.X)
    $y0 = [int][Math]::Round($bounds.Y)
    $bw = [int][Math]::Round($bounds.Width)
    $bh = [int][Math]::Round($bounds.Height)

    # 面の色: 枠の内側の 4 か所（左・右・上・下の中ほど）の最頻値
    $points = @(
        @(($x0 + 3), ($y0 + [int]($bh / 2))), @(($x0 + $bw - 4), ($y0 + [int]($bh / 2))),
        @(($x0 + [int]($bw / 2)), ($y0 + 3)), @(($x0 + [int]($bw / 2)), ($y0 + $bh - 4)))
    $colors = foreach ($p in $points) {
        $o = $p[1] * $stride + $p[0] * 4
        "{0},{1},{2}" -f $bytes[$o + 2], $bytes[$o + 1], $bytes[$o]
    }
    $face = @(($colors | Group-Object | Sort-Object Count -Descending | Select-Object -First 1).Name -split "," | ForEach-Object { [int]$_ })

    $columnTop = New-Object 'int[]' $bw
    $columnBottom = New-Object 'int[]' $bw
    for ($c = 0; $c -lt $bw; $c++) { $columnTop[$c] = -1; $columnBottom[$c] = -1 }
    for ($y = $y0 + 3; $y -le $y0 + $bh - 4; $y++) {
        for ($x = $x0 + 3; $x -le $x0 + $bw - 4; $x++) {
            $o = $y * $stride + $x * 4
            $diff = [Math]::Max([Math]::Max([Math]::Abs([int]$bytes[$o + 2] - $face[0]), [Math]::Abs([int]$bytes[$o + 1] - $face[1])), [Math]::Abs([int]$bytes[$o] - $face[2]))
            if ($diff -ge 40) {
                $c = $x - $x0
                if ($columnTop[$c] -lt 0 -or $y -lt $columnTop[$c]) { $columnTop[$c] = $y }
                if ($y -gt $columnBottom[$c]) { $columnBottom[$c] = $y }
            }
        }
    }

    $groups = New-Object System.Collections.ArrayList
    $start = -1; $last = -1; $gap = 0
    $close = {
        param ($from, $to)
        $top = 99999; $bottom = -1
        for ($c = $from; $c -le $to; $c++) {
            if ($columnTop[$c] -ge 0) {
                if ($columnTop[$c] -lt $top) { $top = $columnTop[$c] }
                if ($columnBottom[$c] -gt $bottom) { $bottom = $columnBottom[$c] }
            }
        }
        if ($to - $from + 1 -lt 3) { return }   # 1〜2 画素幅の縦線（枠・区切り線）は中身ではない
        [void]$groups.Add(@{
            Dx = (($from + $to + 1) / 2) - ($bw / 2)
            Dy = (($top + $bottom + 1) / 2) - ($y0 + $bh / 2)
            Width = $to - $from + 1
            Height = $bottom - $top + 1
        })
    }
    for ($c = 0; $c -lt $bw; $c++) {
        if ($columnTop[$c] -ge 0) {
            if ($start -lt 0) { $start = $c }
            $last = $c; $gap = 0
        } elseif ($start -ge 0) {
            $gap++
            if ($gap -ge 5) { & $close $start $last; $start = -1 }
        }
    }
    if ($start -ge 0) { & $close $start $last }

    # 全体（まとまりをすべて含む範囲）の中心のずれ
    $dx = $null; $dy = $null
    if ($groups.Count -gt 0) {
        $left = ($groups | ForEach-Object { $_.Dx - $_.Width / 2 } | Measure-Object -Minimum).Minimum
        $right = ($groups | ForEach-Object { $_.Dx + $_.Width / 2 } | Measure-Object -Maximum).Maximum
        $top = ($groups | ForEach-Object { $_.Dy - $_.Height / 2 } | Measure-Object -Minimum).Minimum
        $bottom = ($groups | ForEach-Object { $_.Dy + $_.Height / 2 } | Measure-Object -Maximum).Maximum
        $dx = ($left + $right) / 2
        $dy = ($top + $bottom) / 2
    }
    return @{ Groups = @($groups.ToArray()); Dx = $dx; Dy = $dy; Frame = @{ X = $x0; Y = $y0; Width = $bw; Height = $bh } }
}
