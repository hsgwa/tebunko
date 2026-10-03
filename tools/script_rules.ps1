# scripts\ の決まりを確かめるところ（tests/meta/layers.Tests.ps1・tools/check_release_package.ps1・
# tests/meta/safety.Tests.ps1・tools/new_single_script.ps1 の 4 つから読む）で共有する部品。
#
# - 読み込み口からのたどり方（dot-source の相手を見つける）
# - 禁止の語の一覧（危険な処理・ライブラリ。0 件であるべきもの）
#
# ここは tools\ のテスト以外のどこからも読み込まれる共有の置き場なので、スクリプト自身は scripts\ の層の決まり
# （判断層・状態層・画面層）の対象にしない（tools\ の道具・検査だけが読む）。

# ファイルが dot-source している先を返す（存在するかどうかは呼び出し側が確かめる）。
# ". "$PSScriptRoot\..." の形と、ui\ 配下のファイルが使う ". "$TebunkoDir\..." の形
#（起動口 gui.ps1 から渡される tebunko\ 直下）を見る。$tebunkoDir を省くと TebunkoDir の形は無視する
function getDotSourceTargets {
    param (
        [string]$path,
        [string]$tebunkoDir = ""
    )

    $dir = Split-Path $path -Parent
    $text = [System.IO.File]::ReadAllText($path)
    $result = @()
    foreach ($match in [regex]::Matches($text, '(?m)^\s*\.\s+"\$(PSScriptRoot|TebunkoDir)\\([^"]+)"')) {
        $base = if ($match.Groups[1].Value -eq "TebunkoDir") { $tebunkoDir } else { $dir }
        if (-not $base) { continue }
        $result += [System.IO.Path]::GetFullPath((Join-Path $base $match.Groups[2].Value))
    }
    return $result
}

# 起点のファイルから dot-source でたどれる、存在するファイルを全部返す（幅優先）。
# 存在しない先は黙って無視する（読み込み漏れの検査は「全部のファイルが起点からたどれるか」を見るので、
# 無い先を含めても含めなくても結果は変わらない。無い先を検査したいとき（配布物の検査）は getDotSourceTargets を直接使う）
function getReachableFiles {
    param (
        [string[]]$entries,
        [string]$tebunkoDir = ""
    )

    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    $queue = New-Object System.Collections.Queue
    foreach ($entry in $entries) {
        $full = (Resolve-Path -LiteralPath $entry).Path
        if ($seen.Add($full)) { $queue.Enqueue($full) }
    }
    while ($queue.Count -gt 0) {
        foreach ($next in (getDotSourceTargets $queue.Dequeue() $tebunkoDir)) {
            if ((Test-Path -LiteralPath $next) -and $seen.Add($next)) { $queue.Enqueue($next) }
        }
    }
    return $seen
}

# 禁止の語の一覧（docs/safety/checks.md「検査項目と結果」）。0 件であるべきものだけを集める
# （「1 か所だけ許す」といった個別の例外付きの検査は、tests/meta/safety.Tests.ps1 にそのまま残す）。
# キーは説明用の名前、値は検査する正規表現
$script:bannedCodePatterns = [ordered]@{
    InvokeExpression     = 'Invoke-Expression|[^-\w]iex[ (]|ScriptBlock\]::Create'
    EncodedCommand       = 'FromBase64String|EncodedCommand'
    Network              = 'Invoke-WebRequest|Invoke-RestMethod|WebClient|HttpClient|Net\.Sockets|Start-BitsTransfer|DownloadFile|DownloadString|System\.Net\.'
    PInvoke              = 'DllImport|GetDelegateForFunctionPointer'
    Registry             = 'HKLM|HKCU|HKEY_|Set-ItemProperty|New-ItemProperty|Registry::'
    PrivilegeService     = 'Set-Acl|icacls|schtasks|New-Service|Start-Service|sc\.exe|-Verb\s+RunAs'
    ExecutionPolicySet   = 'Set-ExecutionPolicy'
    ExecutionPolicyBypass = 'ExecutionPolicy\s+Bypass'
    Credential           = 'Get-Credential|ConvertTo-SecureString|PSCredential'
    WeakHash             = 'MD5|SHA1|RIPEMD|SHA(256|384|512)Managed|HashAlgorithm\]::Create'
    Remote               = 'Invoke-Command|New-PSSession|Enter-PSSession|WinRM'
}

# 禁止の語のどれかに当たった行を "名前: 行" の一覧（文字列）で返す。1 件も無ければ空文字列
function findBannedCode {
    param (
        [string]$text
    )

    $hits = New-Object System.Collections.Generic.List[string]
    foreach ($name in $script:bannedCodePatterns.Keys) {
        foreach ($match in [regex]::Matches($text, $script:bannedCodePatterns[$name])) {
            $hits.Add("${name}: $($match.Value)")
        }
    }
    return ($hits -join ", ")
}
