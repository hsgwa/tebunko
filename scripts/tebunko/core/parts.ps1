# 部品（lib・indexerLib）の読み込み先をまとめる。
# zip 版は scripts 配下の実際のファイルを dot-source し、単一 .ps1 版（展開せずに動く試験版）は
# 埋め込んだ部品の文字列（${bundledParts}）を関数（importTebunkoPart）としてランスペースに登録する。
# 設計は docs/design/structure/single-script.md・docs/design/structure/classes.md「クラスのつなぎ方」。
#
# 画面だけのクラス（SearchTarget・IndexNode など、ui\ 配下）は ${bundledParts} に含まれない。背景の
# スレッド・別のランスペースがこの部品だけを読み込んでも、画面のクラスを二重にコンパイルしない
# （PowerShell 5.1 は class を構文解析の時点でそのランスペースの型として登録するため、同じ名前の class を
# 別のランスペースでもう一度解析すると、型が別物になり画面のスレッドとキャストできなくなる）。
#
# 単一 .ps1 版かどうかは、先頭で定義する変数 ${bundledParts} の有無だけで決める
# （結合の道具 tools/new_single_script.ps1 が作る）。
# 別スレッド・別ランスペースで部品を読み込むところは、パスの文字列やソースを直接組み立てず、必ずここを通す。

function getPartLoad {
    # 戻り値は @{ State = <InitialSessionState>; Prelude = <部品を読み込む 1 行の文字列> }。
    # 呼び出し側は、この State でランスペース（または WorkerPool）を作り、Prelude を本体の前につないで
    # （shared\core\worker_pool.ps1 の joinWorkerScript）動かす。
    #
    # zip 版: State はただの既定値、Prelude はそのファイルを dot-source する 1 行。
    # 単一 .ps1 版: State に、埋め込んだ部品の文字列を本体に持つ関数 importTebunkoPart を登録し、
    #   Prelude はその関数を呼ぶだけの 1 行（". importTebunkoPart"）。関数の本体は、その State で
    #   ランスペースを作ったときに 1 回だけ構文解析される
    param (
        [ValidateSet("lib", "indexerLib")]
        [string]$name
    )

    $state = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
    $bundled = Get-Variable -Name bundledParts -ErrorAction SilentlyContinue
    if ($null -ne $bundled -and $bundled.Value) {
        $body = $bundled.Value[$name]
        if ($null -eq $body) {
            throw "単一 .ps1 版に ${name} の部品が入っていません。"
        }
        $entry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry("importTebunkoPart", $body)
        $state.Commands.Add($entry)
        # 部品の中（version.ps1 の readVersionFile など）が ${bundledScriptPath}・${bundledVersion} を見るため、
        # 呼び出したスレッドの値をそのまま渡す
        foreach ($varName in "bundledScriptPath", "bundledVersion") {
            $v = Get-Variable -Name $varName -ErrorAction SilentlyContinue
            if ($null -ne $v) {
                $entry2 = New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry($varName, $v.Value, "")
                $state.Variables.Add($entry2)
            }
        }
        return @{ State = $state; Prelude = ". importTebunkoPart" }
    }

    # indexerLib は、取り込みのスレッド・インデクサの司令のスレッドの両方から使うため、
    # 関数（indexer_lib.ps1）と、呼び出す口（invokeIndexerMain・indexer_main.ps1）をまとめて読み込む
    $paths = switch ($name) {
        "lib" { @("$PSScriptRoot\..\lib.ps1") }
        "indexerLib" { @("$PSScriptRoot\..\indexer\indexer_lib.ps1", "$PSScriptRoot\..\indexer\indexer_main.ps1") }
    }
    $prelude = ($paths | ForEach-Object {
        $resolved = (Resolve-Path $_).Path.Replace("'", "''")
        ". '$resolved'"
    }) -join "; "
    return @{ State = $state; Prelude = $prelude }
}
