# tools\capture_screens.ps1 が使う、判断だけの関数（$ui・WPF・実際の画面に触らない）。
# 単体テストは tests\tools\capture_screens.Tests.ps1。tests\meta\screens.Tests.ps1 も、写真を撮る動きに触らず
# ${captureIds} だけを読むため、ここを読み込む（capture_screens.ps1 は画面を動かすため読み込まない）。

# 状態の一覧（docs\design\gui\screens\index.md「撮る状態の一覧」と同じ順）。ID は写真のファイル名になる。
# search-tab/result-menu・search-tab/preview-menu（結果・プレビューの右クリックのメニュー）は、UI オートメーションの
# Invoke でも、ネイティブの WM_CONTEXTMENU（キー操作と同じ扱い）でも開けなかった（フォーカスで出るツールヒントを
# 拾うだけだった）ため、「撮らないもの」に回し、この一覧には入れない（docs\design\gui\screens\index.md「撮らないもの」）
${captureIds} = @(
    "window/about", "window/close-confirm", "window/settings-broken",
    "window/leftover", "window/leftover-open", "window/leftover-many", "window/leftover-search", "window/leftover-killed", "window/leftover-partial",
    "index-tab/empty", "index-tab/normal", "index-tab/interrupted", "index-tab/add", "index-tab/add-error",
    "index-tab/edit", "index-tab/delete-confirm", "index-tab/checked", "index-tab/start-confirm",
    "index-tab/running", "index-tab/stop-confirm", "index-tab/done", "index-tab/failed",
    "search-tab/no-index", "search-tab/initial", "search-tab/results", "search-tab/no-results", "search-tab/limit",
    "search-tab/regex-error", "search-tab/tree-none", "search-tab/collapsed", "search-tab/filtered",
    "search-tab/missing-source", "search-tab/min-width",
    "settings-tab/normal", "settings-tab/running-warning", "settings-tab/empty-confirm", "settings-tab/nonempty-confirm",
    "settings-tab/index-confirm", "settings-tab/invalid-warning"
)

function resolveCaptureIds {
    # -Only で選んだ状態の ID の一覧を、一覧にある表記（大文字小文字は区別しない）に直して返す。
    # 指定が無ければ Ids をそのまま返す。知らない ID があれば例外にする
    param (
        [string[]]$Only,
        [string[]]$Ids
    )

    if (!$Only -or $Only.Count -eq 0) {
        return $Ids
    }
    $result = New-Object System.Collections.ArrayList
    foreach ($id in $Only) {
        $found = $Ids | Where-Object { $_ -eq $id }
        if (!$found) {
            throw "知らない状態の ID です: $id"
        }
        [void]$result.Add($found)
    }
    return @($result.ToArray())
}

function getCaptureImagePath {
    # ID（"<画面>/<状態>"）から、写真の置き場所（<OutDir>\<画面>\<状態>.png）を作る
    param (
        [string]$Id,
        [string]$OutDir
    )

    $parts = $Id -split "/", 2
    if ($parts.Count -ne 2 -or !$parts[0] -or !$parts[1]) {
        throw "ID の形が正しくありません（<画面>/<状態> の形にしてください）: $Id"
    }
    return (Join-Path $OutDir (Join-Path $parts[0] ($parts[1] + ".png")))
}

function getCaptureEnvironmentProblems {
    # 撮る前にそろえる条件に合っているかを確かめ、合わない理由の一覧を返す（すべて合えば空の配列）。
    # 実際の値（レジストリ・画面の数・Office のプロセス）を読むのは呼び出す側で行い、ここには読んだ値だけを渡す
    param (
        [int]$AppliedDpi,
        [bool]$LightTheme,
        [int]$ScreenCount,
        [int]$OfficeProcessCount,
        [int]$ExpectedDpi = 96
    )

    $problems = New-Object System.Collections.ArrayList
    if ($AppliedDpi -ne $ExpectedDpi) {
        [void]$problems.Add("表示の倍率が 100%（AppliedDPI が $ExpectedDpi）ではありません（今の値: $AppliedDpi）。表示の倍率を 100% にしてください。")
    }
    if (!$LightTheme) {
        [void]$problems.Add("Windows のテーマがライトモードではありません。設定でライトモードにしてください。")
    }
    if ($ScreenCount -ne 1) {
        [void]$problems.Add("画面が 1 つではありません（今の数: $ScreenCount）。画面を 1 つにしてください。")
    }
    if ($OfficeProcessCount -gt 0) {
        [void]$problems.Add("Excel・Word・PowerPoint が動いています（$OfficeProcessCount 個）。終了してから流してください。")
    }
    return @($problems.ToArray())
}

function testCaptureSensitiveText {
    # 部品の文字（Name・入力欄の Value）に、利用者名・コンピューター名・利用者のフォルダのパスが含まれるか
    param (
        [string]$Text,
        [string]$UserName,
        [string]$ComputerName,
        [string]$UserProfile
    )

    if ([string]::IsNullOrEmpty($Text)) {
        return $false
    }
    foreach ($needle in @($UserName, $ComputerName, $UserProfile)) {
        if ($needle -and $Text.IndexOf($needle, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            return $true
        }
    }
    return $false
}

function getCaptureRedactedText {
    # 利用者のフォルダのパス・利用者名・コンピューター名を、塗った上に描く架空の文字に置き換える。
    # 長い文字から先に置き換える（例えばコンピューター名が「<利用者名>-PC」のように利用者名を含むとき、
    # 利用者名を先に置き換えると、コンピューター名のほうが一部書き換わって一致しなくなるため）
    param (
        [string]$Text,
        [string]$UserProfile,
        [string]$UserName,
        [string]$ComputerName,
        [string]$UserProfileReplacement = "C:\Users\test",
        [string]$UserNameReplacement = "test",
        [string]$ComputerNameReplacement = "TEST-PC"
    )

    if ([string]::IsNullOrEmpty($Text)) {
        return $Text
    }
    $result = $Text
    $pairs = @(
        @{ Needle = $UserProfile; Replacement = $UserProfileReplacement },
        @{ Needle = $UserName; Replacement = $UserNameReplacement },
        @{ Needle = $ComputerName; Replacement = $ComputerNameReplacement }
    ) | Where-Object { ![string]::IsNullOrEmpty($_.Needle) } | Sort-Object { $_.Needle.Length } -Descending
    foreach ($pair in $pairs) {
        $result = [regex]::Replace($result, [regex]::Escape($pair.Needle), { param ($m) $pair.Replacement }, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    }
    return $result
}

function testCaptureImageSize {
    # 撮った画像の大きさが目安（既定 200 KB）以下か
    param (
        [long]$Bytes,
        [long]$LimitBytes = 200KB
    )

    return $Bytes -le $LimitBytes
}
