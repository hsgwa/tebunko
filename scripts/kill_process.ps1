# Excelプロセス終了
#
# 変換処理の中断などで残ったExcelプロセスを強制終了する。

$processes = @(Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue)

if ($processes.Count -eq 0) {
    Write-Host "Excelプロセスは実行されていません。" -ForegroundColor Green
    pause
    exit
}

# ウィンドウを持たないExcelは、変換処理などでバックグラウンド起動されたもの
$hidden = @($processes | Where-Object { $_.MainWindowHandle -eq [IntPtr]::Zero })
$visible = @($processes | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero })

Write-Host "Excelプロセスが起動中です。" -ForegroundColor Yellow
Write-Host "  バックグラウンド（画面に表示されていない）: $($hidden.Count) 個"
Write-Host "  画面に表示中                              : $($visible.Count) 個"
Write-Host ""
Write-Host "※ 変換処理の実行中にバックグラウンドのExcelを終了すると、変換中のファイルは失敗扱いになります。"
Write-Host "※ 画面に表示中のExcelを終了すると、保存していない内容は失われます。"
Write-Host ""
Write-Host "  1: バックグラウンドのExcelのみ終了"
Write-Host "  2: すべてのExcelを保存せずに終了"
$response = Read-Host "番号を入力してください（それ以外: 中止）"

switch ($response) {
    "1" { $targets = $hidden }
    "2" { $targets = $processes }
    default { $targets = $null }
}

if ($null -eq $targets) {
    Write-Host "処理を中止しました。"
} elseif ($targets.Count -eq 0) {
    Write-Host "終了対象のExcelプロセスはありません。" -ForegroundColor Green
} else {
    $targets | Stop-Process -Force
    Write-Host "Excelプロセスを $($targets.Count) 個終了させました。" -ForegroundColor Green
}

pause
