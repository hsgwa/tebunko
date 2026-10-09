# indexer.ps1（起動口）を実際に動かすテストの共通の手伝い。
# tests\tebunko\indexer\indexer.Tests.ps1・tests\tebunko\indexer\index_compat.Tests.ps1 の両方が使う。
#   . "$PSScriptRoot\..\..\helpers\load.ps1"   を先にしてから dot-source する（${scriptsDir} を使うため）
#   . "$PSScriptRoot\..\..\helpers\indexer.ps1"

$indexerPath  = "${scriptsDir}\tebunko\indexer.ps1"
$runPath      = "${scriptsDir}\tebunko\indexer\indexer_run.ps1"
$settingsScriptPath = "${scriptsDir}\tebunko\core\settings.ps1"
$pathsPath    = "${scriptsDir}\tebunko\core\paths.ps1"

function findLine {
    # ファイルの中で pattern に一致する最初の行の番号を返す（テストが行番号を直接書かないようにする）
    param ([string]$path, [string]$pattern)
    $lines = [System.IO.File]::ReadAllLines($path)
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match $pattern) { return $i + 1 }
    }
    throw "${path} に ${pattern} がありません"
}

$script:rootCount = 0
function newRoot {
    # テストごとに別のツールの置き場所（setting.config・work\ を置くフォルダ）を作る
    $script:rootCount++
    $root = Join-Path $TestDrive "tool$($script:rootCount)"
    [System.IO.Directory]::CreateDirectory("$root\work") | Out-Null
    return $root
}

function writeTestSettings {
    # テスト用の置き場所に setting.config を書く。folders は @{ name; path; enabled } の配列
    param ([string]$root, [object[]]$folders)
    $settings = newSettings
    $settings.targetFolders = @($folders)
    # 既定のワークスペース（%USERPROFILE%\Documents\tebunko_ws）には利用者のインデックスがあるため、テスト用の work を指す
    $settings.workspaceFolder = "$root\work"
    writeSettings $settings "$root\setting.config"
}

function assertIndexerWorkspaceIsolated {
    # indexer.ps1 を動かす前に、取り込みの出力先（setting.config の workspaceFolder。空なら既定のワークスペース）が、
    # テスト用の置き場所（root）の中か、差し替えた既定のワークスペースの中であることを確かめる。外なら例外にする
    # （利用者の本物の既定のワークスペースに取り込みの跡を付けないため）。Real・DefaultDir は、本物の代わりを渡して確かめるときに使う
    param (
        [string]$root,
        [string]$real = (getRealDefaultWorkspace),
        [string]$defaultDir = (getDefaultWorkDir)
    )

    $folder = ([string](readSettings "$root\setting.config").workspaceFolder).Trim()
    $dir = if ($folder -eq "") { $defaultDir } else {
        [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($root, [System.Environment]::ExpandEnvironmentVariables($folder)))
    }
    assertNotRealWorkspace $dir $real
    if (!(testPathInside $dir $root) -and !(testPathInside $dir $defaultDir)) {
        throw "取り込みの出力先がテスト用の置き場所の外です: $dir"
    }
}

# 最後に動かしたインデックス作成の受け渡しの口（進み具合・エラーを確かめる）
$script:lastChannel = $null

function runIndexer {
    # テスト用の置き場所（root）で indexer.ps1 を動かし、終了コードを返す。
    #   options: @{ RetryFailed; ConfirmTargets; Workers（既定 0 = 取り込みのスレッドを使わない） }
    #   breaks : 途中で動かす処理 @{ Script; Pattern; Action }（Pattern に一致する行に来るたびに Action を動かす）
    param (
        [string]$root,
        [hashtable]$options = @{},
        [object[]]$breaks = @()
    )

    assertIndexerWorkspaceIsolated $root
    $workers = if ($options.ContainsKey("Workers")) { $options.Workers } else { 0 }
    $channel = newIndexerChannel ([bool]$options.RetryFailed) ([bool]$options.ConfirmTargets) $workers
    if ($options.ContainsKey("OnlyNames")) {
        $channel.OnlyNames = @($options.OnlyNames)
    }
    $script:lastChannel = $channel
    $global:indexerTestRoot = $root
    $global:capturedTmpDir = $null
    $global:capturedLegacyTmpParent = $null
    $points = New-Object System.Collections.Generic.List[object]
    try {
        # ${settingsFile} を決める行（settings.ps1）で、その前に ${rootDir} を差し替える（テスト用のフォルダには書き込めるため、setting.config・work もそこになる）。
        # Action は止まった場所の子のスコープで動く
        $points.Add((Set-PSBreakpoint -Script $settingsScriptPath -Line (findLine $settingsScriptPath '^\$\{settingsFile\}\s*=') -Action {
            Set-Variable -Name rootDir -Value $global:indexerTestRoot -Scope 1
        }))
        # ${legacyTmpParent}（前の版の片付けだけに使う場所 %TEMP%\tebunko）の既定値を決めた直後の行で差し替え、
        # テストが利用者の本物の %TEMP%\tebunko に触れないようにする
        $points.Add((Set-PSBreakpoint -Script $pathsPath -Line (findLine $pathsPath '^\$excelMaxPath = 218') -Action {
            Set-Variable -Name legacyTmpParent -Value (Join-Path $global:indexerTestRoot "legacy_tmp") -Scope 1
        }))
        # 取り込みの作業フォルダ（$tmpDir）が決まった直後の行で、その値を控える
        # （終わったあとはワークスペースの場所・前の版の片付けだけに使う場所のどちらも後片付けで消えるため、途中でしか確かめられない）。
        # ${legacyTmpParent} の差し替え（上のブレークポイント）が実際に initTmpDir まで効いていることも、
        # 同じ場所で控えて確かめる（差し替えの行自体が paths.ps1 の無関係な行に依存しているため、
        # ここで使われた値を見ないと、行順が変わって差し替えが上書きされても気付けない）
        $points.Add((Set-PSBreakpoint -Script $runPath -Line ((findLine $runPath '\$script:tmpDirReason = \$selected\.Reason') + 1) -Action {
            Set-Variable -Name capturedTmpDir -Value (Get-Variable -Name tmpDir -ValueOnly) -Scope Global
            Set-Variable -Name capturedLegacyTmpParent -Value (Get-Variable -Name legacyTmpParent -ValueOnly) -Scope Global
        }))
        foreach ($break in $breaks) {
            $points.Add((Set-PSBreakpoint -Script $break.Script -Line (findLine $break.Script $break.Pattern) -Action $break.Action))
        }
        & $indexerPath -Channel $channel *> $null
        return $channel.ExitCode
    } finally {
        foreach ($point in $points) { Remove-PSBreakpoint -Breakpoint $point }
        Remove-Variable -Name indexerTestRoot -Scope Global -ErrorAction SilentlyContinue
    }
    # $global:capturedTmpDir は呼び出し側が結果を確かめられるよう、次の runIndexer まで残す
}

function readTestStatus {
    param ([string]$root)
    return (readStatusFile "$root\work\ingest_status.tsv")
}

function readTestSystemState {
    param ([string]$root)
    return (readSystemIndexState "$root\work\system_index_state.tsv")
}

function readTestError {
    return $script:lastChannel.Error
}

function readTestProgress {
    return (readIndexingProgress $script:lastChannel)
}
