# 画面のスレッドの分割（shared\ui\shell.ps1 の startUiWork ほか）を、shell.ps1 全体を読み込まずに取り出す。
#   . "$PSScriptRoot\..\..\..\helpers\ui_work.ps1"   （dot-source は、取り出した関数を使う場の scope で行う）
# newTimer は偽物にする（タイマーは動かさない。Tick は $script:fakeTimers に取っておき、テストが手で呼べる）。
# 予算の時計（getUiWorkMilliseconds）は本物のまま。件数の少ないテストは、その場で最後まで回る。
# 結果を見る前に、まだ回っているかもしれない仕事は completeUiWork で終わらせる。

$script:fakeTimers = New-Object System.Collections.Generic.List[object]
function newTimer {
    param ([int]$milliseconds, [scriptblock]$onTick)
    $timer = [pscustomobject]@{ Tag = $null; Interval = $milliseconds; OnTick = $onTick; Running = $false }
    $timer | Add-Member ScriptMethod Start { $this.Running = $true }
    $timer | Add-Member ScriptMethod Stop { $this.Running = $false }
    $script:fakeTimers.Add($timer)
    return $timer
}

$uiWorkAst = [System.Management.Automation.Language.Parser]::ParseFile("${scriptsDir}\shared\ui\shell.ps1", [ref]$null, [ref]$null)
foreach ($uiWorkFunction in @("getUiWorkMilliseconds", "stopUiWork", "stopAllUiWork", "testUiWorkRunning", "failUiWork", "stepUiWork", "startUiWork", "completeUiWork")) {
    $uiWorkName = $uiWorkFunction
    $uiWorkNode = $uiWorkAst.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $uiWorkName }, $true)
    . ([scriptblock]::Create($uiWorkNode.Extent.Text))
}
$script:uiWorks = @{}
