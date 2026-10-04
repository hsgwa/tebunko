# 部品（lib・indexerLib）の読み込み先（tebunko\core\parts.ps1）のテスト。
# 設計は docs/design/structure/single-script.md・docs/design/structure/classes.md「クラスのつなぎ方」
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "getPartLoad（zip 版。`${bundledParts} が無い）" -Tag Unit {
    It "<Name> は scripts 配下の実際のファイルを dot-source する Prelude を返す" -ForEach @(
        @{ Name = "lib"; Files = @("tebunko\lib.ps1") }
        @{ Name = "indexerLib"; Files = @("tebunko\indexer\indexer_lib.ps1", "tebunko\indexer\indexer_main.ps1") }
    ) {
        $load = getPartLoad $Name
        $load.State | Should -BeOfType [System.Management.Automation.Runspaces.InitialSessionState]
        foreach ($file in $Files) {
            $resolved = (Resolve-Path "$scriptsDir\$file").Path
            $load.Prelude | Should -Match ([regex]::Escape($resolved))
        }
    }

    It "indexer は ValidateSet に無い（単一 .ps1 版で画面のクラスを二重にコンパイルしないため廃止した）" {
        { getPartLoad indexer } | Should -Throw
    }
}

Describe "getPartLoad（単一 .ps1 版。`${bundledParts} がある）" -Tag Unit {
    BeforeEach {
        $global:bundledParts = @{
            lib        = 'function dummyLibMarker { "lib" }'
            indexerLib = 'function dummyIndexerMarker { "indexerLib" }'
        }
        $global:bundledScriptPath = "C:\fake\single.ps1"
        $global:bundledVersion = @{ Tag = "v9.9.9"; Sha = "deadbeef" }
    }
    AfterEach {
        Remove-Variable -Name bundledParts, bundledScriptPath, bundledVersion -Scope Global -ErrorAction SilentlyContinue
    }

    It "<Name> の部品があれば、importTebunkoPart を本体に持つ関数を State に登録し、Prelude はそれを呼ぶだけの 1 行" -ForEach @(
        @{ Name = "lib" }
        @{ Name = "indexerLib" }
    ) {
        $load = getPartLoad $Name
        $load.Prelude | Should -Be ". importTebunkoPart"
        @($load.State.Commands | Where-Object { $_.Name -eq "importTebunkoPart" }).Count | Should -Be 1
    }

    It "bundledScriptPath・bundledVersion を、呼び出したスレッドの値のまま State に渡す" {
        $load = getPartLoad lib
        $rs = [runspacefactory]::CreateRunspace($load.State)
        $rs.Open()
        try {
            $ps = [powershell]::Create()
            $ps.Runspace = $rs
            $result = $ps.AddScript('$bundledScriptPath; $bundledVersion.Tag').Invoke()
            $result[0] | Should -Be "C:\fake\single.ps1"
            $result[1] | Should -Be "v9.9.9"
        } finally {
            $rs.Dispose()
        }
    }

    It "bundledParts も State に渡す（取り込みのスレッドの中でもう一度 getPartLoad を呼んでも、パスの読み込みに落ちない）" {
        $load = getPartLoad lib
        $rs = [runspacefactory]::CreateRunspace($load.State)
        $rs.Open()
        try {
            $ps = [powershell]::Create()
            $ps.Runspace = $rs
            $result = $ps.AddScript('$bundledParts.Keys | Sort-Object').Invoke()
            @($result) | Should -Be @("indexerLib", "lib")
        } finally {
            $rs.Dispose()
        }
    }

    It "指定した名前の部品が入っていなければ例外を投げる" {
        $global:bundledParts = @{}
        { getPartLoad lib } | Should -Throw "*lib*"
    }
}

Describe "getPartLoad（同じランスペースを使い回しても、部品のクラスの型が変わらない）" -Tag Io {
    It "2 回のジョブをまたいでも、部品で定義したクラスは同じ型のまま" {
        $global:bundledParts = @{
            lib = @'
class ProbeTypeA {
    [string] Hello() { return "hi" }
}
function newProbeTypeA { return [ProbeTypeA]::new() }
'@
        }
        $load = getPartLoad lib
        $rs = [runspacefactory]::CreateRunspace($load.State)
        $rs.Open()
        try {
            $ps1 = [powershell]::Create()
            $ps1.Runspace = $rs
            $joined1 = joinWorkerScript $load.Prelude { newProbeTypeA }
            $type1 = $ps1.AddScript($joined1).Invoke()[0].GetType()
            $ps1.Dispose()

            $ps2 = [powershell]::Create()
            $ps2.Runspace = $rs
            $joined2 = joinWorkerScript $load.Prelude { newProbeTypeA }
            $type2 = $ps2.AddScript($joined2).Invoke()[0].GetType()
            $ps2.Dispose()

            $type1 | Should -Be $type2
        } finally {
            $rs.Dispose()
            Remove-Variable -Name bundledParts -Scope Global -ErrorAction SilentlyContinue
        }
    }
}

Describe "getPartLoad（別のランスペースの部品を取り違えない）" -Tag Io {
    It "同じ名前のクラスを別々のランスペースでそれぞれ使っても、実行したランスペースを取り違えない" {
        $global:bundledParts = @{
            lib = @'
class ProbeTypeB {
    [guid] WhichRunspace() { return [runspace]::DefaultRunspace.InstanceId }
}
function newProbeTypeB { return [ProbeTypeB]::new() }
'@
        }
        $load1 = getPartLoad lib
        $rs1 = [runspacefactory]::CreateRunspace($load1.State)
        $rs1.Open()
        $load2 = getPartLoad lib
        $rs2 = [runspacefactory]::CreateRunspace($load2.State)
        $rs2.Open()
        try {
            $ps1 = [powershell]::Create()
            $ps1.Runspace = $rs1
            $id1 = $ps1.AddScript((joinWorkerScript $load1.Prelude { (newProbeTypeB).WhichRunspace() })).Invoke()[0]
            $ps1.Dispose()

            $ps2 = [powershell]::Create()
            $ps2.Runspace = $rs2
            $id2 = $ps2.AddScript((joinWorkerScript $load2.Prelude { (newProbeTypeB).WhichRunspace() })).Invoke()[0]
            $ps2.Dispose()

            $id1 | Should -Be $rs1.InstanceId
            $id2 | Should -Be $rs2.InstanceId
            $id1 | Should -Not -Be $id2
        } finally {
            $rs1.Dispose()
            $rs2.Dispose()
            Remove-Variable -Name bundledParts -Scope Global -ErrorAction SilentlyContinue
        }
    }
}
