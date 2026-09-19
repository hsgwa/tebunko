$excepProcesses = Get-Process -Name "excel" -ErrorAction SilentlyContinue

if ($excepProcesses) {
    Write-Host "Excelプロセスが起動中です。" -ForegroundColor Yellow
    $response = Read-Host "すべてのExcelプロセスを保存せずに終了しますか？(y/N)"

    if ($response -eq "y") {
        $excelApp = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        if ($excelApp) {
            Stop-Process -Name "excel" -Force
            Write-Host "すべてのExcelプロセスを終了させました。" -ForegroundColor Green
        }
    } else {
        Write-Host "処理を中止しました。"
    }
} else {
    Write-Host "Excelプロセスは実行されていません。" -ForegroundColor Green
}

pause
