# Office の見張りをクラス（OfficeWatchdog）にしたあと、古い関数・変数の名前が戻らないことを確かめるテスト。
# 古い名前で呼ぶと、呼び出しが try/catch に飲まれて静かに何も起きなくなるため、名前で止める。
# 見る先は scripts と tests の .ps1、docs の .md。
Describe "見張りの古い名前" -Tag Meta {
    It "scripts・tests・docs に、古い名前が残っていない" {
        $repo = (Resolve-Path "$PSScriptRoot\..\..").Path
        $self = Join-Path $PSScriptRoot "old_watchdog_names.Tests.ps1"
        $pattern = 'startWatchdog|stopWatchdog|updateWatchedPids|script:watchdog'
        $hits = foreach ($target in @(@{ Dir = "scripts"; Include = "*.ps1", "*.psm1" }, @{ Dir = "tests"; Include = "*.ps1", "*.psm1" }, @{ Dir = "docs"; Include = "*.md" })) {
            Get-ChildItem -Path (Join-Path $repo $target.Dir) -Recurse -File -Include $target.Include |
                Where-Object { $_.FullName -ne $self } |
                Select-String -CaseSensitive -Pattern $pattern |
                ForEach-Object { "$($_.Path):$($_.LineNumber)" }
        }
        $because = "古い名前は OfficeWatchdog のメソッド（`$script:officeWatchdog の Start・Stop・UpdateWatchedPids）に置き換えた"
        @($hits) | Should -BeNullOrEmpty -Because $because
    }
}
