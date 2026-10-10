# Office の見張りをクラス（OfficeWatchdog）にしたあと、古い関数・変数の名前が戻らないことを確かめるテスト。
# 古い名前で呼ぶと、呼び出しが try/catch に飲まれて静かに何も起きなくなるため、名前で止める。
Describe "見張りの古い名前" -Tag Meta {
    It "scripts と tests に、古い名前が残っていない" {
        $repo = (Resolve-Path "$PSScriptRoot\..\..").Path
        $self = Join-Path $PSScriptRoot "old_watchdog_names.Tests.ps1"
        $pattern = 'startWatchdog|stopWatchdog|updateWatchedPids|script:watchdog'
        $hits = foreach ($dir in "scripts", "tests") {
            Get-ChildItem -Path (Join-Path $repo $dir) -Recurse -File -Include *.ps1, *.psm1 |
                Where-Object { $_.FullName -ne $self } |
                Select-String -CaseSensitive -Pattern $pattern |
                ForEach-Object { "$($_.Path):$($_.LineNumber)" }
        }
        @($hits) | Should -BeNullOrEmpty
    }
}
