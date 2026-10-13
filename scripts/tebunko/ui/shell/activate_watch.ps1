# startGui の一部: 多重起動した 2 つ目の合図で、この画面を前面に出す見張り（activateTimer）。startGui の中で読み込む。

    # 多重起動したとき（2つ目のプロセスが $activateEvent を合図）に、この画面を前面へ出す。
    # 画面のスレッドで一定間隔にイベントを確認する（P/Invoke を使わず、WPF の Activate で前面化する）。
    # ※以前は C# の SingleInstance（AttachThreadInput 等の P/Invoke）で行っていたが、
    #   実行時コンパイル（csc.exe）を無くすため、DispatcherTimer＋Window.Activate に置き換えた。
    $activateTimer = New-Object System.Windows.Threading.DispatcherTimer
    $activateTimer.Interval = [TimeSpan]::FromMilliseconds(300)
    $activateTimer.Add_Tick({
        if ($activateEvent.WaitOne(0)) {
            if ($window.WindowState -eq [System.Windows.WindowState]::Minimized) {
                $window.WindowState = [System.Windows.WindowState]::Normal
            }
            [void]$window.Activate()
            # ほかのプロセスが前面のときは Activate が無視されることがあるため、最前面を一瞬立ててから戻す
            $window.Topmost = $true
            $window.Topmost = $false
        }
    })
    $activateTimer.Start()
