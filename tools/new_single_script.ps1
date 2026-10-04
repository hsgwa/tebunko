# 展開せずに動く単一 .ps1 版（試験版）を組み立てる。scripts\ 配下のソース（zip 版と共通）を 1 本の .ps1 に
# 畳み込み、work\release\ に書き出す（作業ツリーにはコミットしない。tests/meta/structure.Tests.ps1 の M5）。
# 設計は docs/design/structure/single-script.md。
#
#   .\tools\new_single_script.ps1 -Version v1.0.0
#
# 版の文字列（${bundledVersion}）は tools\new_version_text.ps1 と同じしくみ（-Version と git rev-parse HEAD）で決め、
# VERSION.txt を介さずコードに埋め込む（scripts\shared\core\version.ps1 の readVersionFile が ${bundledVersion} を見る）。
#
# 畳み込みの考え方（どのファイルも同じしくみで扱う）:
#   読み込み口の ". "$PSScriptRoot\..." "." "$TebunkoDir\..." " の行を、その相手のファイルの中身（同じしくみで
#   畳み込んだもの）に差し替える。相手が初めて出てきたときだけ中身を入れ、2 回目以降は空行にする
#   （HashSet で済んだファイルを覚えておく）。関数の中で読み込んでいる行も、その場での文字列の差し替えなので
#   そのまま働く。起動口（gui.ps1・indexer.ps1）だけは param ブロックが二重にならないよう、AST で本体だけを取り出す。
param (
    [Parameter(Mandatory = $true)]
    [string]$Version,
    [string]$OutFile
)

$ErrorActionPreference = "Stop"

if ($Version -notmatch '^[A-Za-z0-9._-]+$') {
    throw "バージョンに使えない文字が含まれています: $Version"
}

$rootDir = Split-Path $PSScriptRoot -Parent
. "$PSScriptRoot\script_rules.ps1"

if (!$OutFile) {
    $OutFile = Join-Path $rootDir "work\release\tebunko-$Version.ps1"
}

# ---- 版の文字列（VERSION.txt と同じ決め方。SHA が取れなければ止める） ----
$versionBytes = & "$PSScriptRoot\new_version_text.ps1" -Version $Version
$versionText = [System.Text.Encoding]::UTF8.GetString($versionBytes).TrimStart([char]0xFEFF)
$versionLines = $versionText -split "`r`n" | Where-Object { $_ -ne "" }
$sha = $versionLines[1]

# ---- 畳み込み ----

$scriptsDir = Join-Path $rootDir "scripts"
$tebunkoDir = Join-Path $scriptsDir "tebunko"
$script:visited = New-Object 'System.Collections.Generic.HashSet[string]'

# dot-source の読み込み行（PSScriptRoot・TebunkoDir の両方の形）を、相手の中身（再帰的に畳み込んだもの）に差し替える
function substituteLoaderLines {
    param (
        [string]$text,
        [string]$dir
    )

    $pattern = '^([ \t]*)\.\s+"\$(PSScriptRoot|TebunkoDir)\\([^"]+)"[ \t]*$'
    $lines = $text -split "`r`n"
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($line in $lines) {
        $match = [regex]::Match($line, $pattern)
        if (!$match.Success) {
            [void]$out.Add($line)
            continue
        }
        $base = if ($match.Groups[2].Value -eq "TebunkoDir") { $tebunkoDir } else { $dir }
        $target = [System.IO.Path]::GetFullPath((Join-Path $base $match.Groups[3].Value))
        [void]$out.Add((inlineFile $target))
    }
    return ($out -join "`r`n")
}

# 1 つのファイルを、中身（dot-source の相手を差し替え済み）に展開する。2 回目以降は空文字列
function inlineFile {
    param (
        [string]$path
    )

    if (!$script:visited.Add($path)) {
        return ""
    }
    $text = [System.IO.File]::ReadAllText($path)
    $dir = Split-Path $path -Parent
    return substituteLoaderLines $text $dir
}

# 起動口（param ブロックを持つことがあるファイル）の、param を除いた本体だけを取り出して畳み込む。
# 結合した単一 .ps1 は、統一した 1 つの param ブロック（ヘッダー）しか持てないため（二重になると構文エラーになる）。
# EndBlock.Extent.Text は param ブロックも含んでしまうため使わず、ParamBlock の終わりから後ろを使う
function inlineEntryBody {
    param (
        [string]$path
    )

    [void]$script:visited.Add($path)
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
    if ($errors.Count -gt 0) {
        throw "$path の読み取りに失敗しました: $($errors[0].Message)"
    }
    $text = [System.IO.File]::ReadAllText($path)
    $body = if ($ast.ParamBlock) { $text.Substring($ast.ParamBlock.Extent.EndOffset) } else { $text }
    $dir = Split-Path $path -Parent
    return substituteLoaderLines $body $dir
}

$guiPath = Join-Path $tebunkoDir "gui.ps1"
$indexerPath = Join-Path $tebunkoDir "indexer.ps1"

$parts = New-Object System.Collections.Generic.List[string]

[void]$parts.Add('# 自動で作る単一 .ps1 版（試験版）。tools\new_single_script.ps1 が組み立てる。手で編集しない。')
[void]$parts.Add('# 元のソースは scripts\ 配下（zip 版と共通）。設計は docs/design/structure/single-script.md')
[void]$parts.Add('param (')
[void]$parts.Add('    [string]$Part = "gui",')
[void]$parts.Add('    [switch]$RetryFailed,')
[void]$parts.Add('    $Channel = $null')
[void]$parts.Add(')')
[void]$parts.Add('')
[void]$parts.Add('${bundledScriptPath} = $PSCommandPath')
[void]$parts.Add("`${bundledVersion} = @{ Tag = '$Version'; Sha = '$sha' }")
[void]$parts.Add('${bundledXaml} = @{}')

# ---- xaml（画面定義）をここに埋め込む。キーは実行時に評価する式（インストール先ごとに $PSScriptRoot が違うため） ----
function addXamlEntries {
    param (
        [string]$dir,
        [string]$keyPrefix
    )

    foreach ($file in (Get-ChildItem -LiteralPath $dir -Filter *.xaml | Sort-Object Name)) {
        $content = [System.IO.File]::ReadAllText($file.FullName)
        if ($content -match "(?m)^'@") {
            throw "$($file.FullName) に、単一引用符のヒアストリングを閉じてしまう行があります（行の先頭が '@）。"
        }
        $keyExpr = '[System.IO.Path]::GetFullPath("$PSScriptRoot\' + $keyPrefix + $file.Name + '")'
        [void]$parts.Add("`${bundledXaml}[$keyExpr] = @'")
        [void]$parts.Add($content.TrimEnd("`r", "`n"))
        [void]$parts.Add("'@")
    }
}

addXamlEntries (Join-Path $tebunkoDir "xaml") "xaml\"
addXamlEntries (Join-Path $scriptsDir "shared\xaml") "..\shared\xaml\"

[void]$parts.Add('')

# ---- 部品（${bundledParts}）をここに埋め込む。lib・indexerLib の、画面（ui\ 配下）に触れない範囲だけ。 ----
# 背景のスレッド・別のランスペースは、ここに埋め込んだ文字列を関数（importTebunkoPart）として読み込み、
# 画面だけのクラス（SearchTarget・IndexNode など）を二重にコンパイルしない（scripts\tebunko\core\parts.ps1）。
# indexerLib は、取り込みのスレッド・インデクサの司令のスレッドの両方が使うため、indexer_lib.ps1（関数）と
# indexer_main.ps1（呼び出す口。invokeIndexerMain）の両方を 1 つの部品として畳み込む。
# 畳み込みは、本体（$script:visited）とは別の・部品ごとに新しい済みファイルの集合で行う
# （部品どうし・部品と本体で、同じファイルを二重に数えないようにするため）
$script:partVisited = @{}
function inlinePart {
    param ([string[]]$entryPaths)
    $script:visited = New-Object 'System.Collections.Generic.HashSet[string]'
    $text = (($entryPaths | ForEach-Object { inlineFile $_ })) -join "`r`n"
    return $text
}

$bundledPartEntries = @(
    @{ Name = "lib"; Paths = @((Join-Path $tebunkoDir "lib.ps1")) }
    @{ Name = "indexerLib"; Paths = @((Join-Path $tebunkoDir "indexer\indexer_lib.ps1"), (Join-Path $tebunkoDir "indexer\indexer_main.ps1")) }
)

[void]$parts.Add('${bundledParts} = @{}')
foreach ($partEntry in $bundledPartEntries) {
    $partText = inlinePart $partEntry.Paths
    $script:partVisited[$partEntry.Name] = [System.Collections.Generic.HashSet[string]]$script:visited
    if ($partText -match "(?m)^'@") {
        throw "単一 .ps1 の部品（$($partEntry.Name)）に、単一引用符のヒアストリングを閉じてしまう行があります（行の先頭が '@）。"
    }
    [void]$parts.Add("`${bundledParts}['$($partEntry.Name)'] = @'")
    [void]$parts.Add($partText.TrimEnd("`r", "`n"))
    [void]$parts.Add("'@")
}
[void]$parts.Add('')

# 本体（画面・起動口）の畳み込みは、部品の済み集合とは別の、まっさらな集合から始める
$script:visited = New-Object 'System.Collections.Generic.HashSet[string]'

# ---- 読み込み口の並び（driver の手順。設計どおり）。 ----
# -Part lib / indexerLib は、ほかのスレッドから dot-source で呼ばれ、必要な関数だけ読み込んで戻る（画面を出さない）。
# -Part indexer は、画面を出さずにインデックス作成だけ行って終わる。どちらでもなければ（既定）画面を起動する
[void]$parts.Add((inlineFile (Join-Path $tebunkoDir "ui\startup_error.ps1")))
[void]$parts.Add((inlineFile (Join-Path $tebunkoDir "lib.ps1")))
[void]$parts.Add('if ($Part -eq "lib") { return }')
[void]$parts.Add((inlineFile (Join-Path $tebunkoDir "indexer\indexer_lib.ps1")))
[void]$parts.Add('if ($Part -eq "indexerLib") { return }')
[void]$parts.Add((inlineFile (Join-Path $tebunkoDir "indexer\indexer_main.ps1")))
[void]$parts.Add('if ($Part -eq "indexer") {')
[void]$parts.Add((inlineEntryBody $indexerPath))
[void]$parts.Add('}')
[void]$parts.Add((inlineEntryBody $guiPath))

$combined = ($parts -join "`r`n") + "`r`n"

# ---- 道具の検査（自分が作った結合の結果を確かめる。1 つでも失敗すれば止める） ----

# 1) 構文として読める（0 個のエラー）
$checkTokens = $null
$checkErrors = $null
$combinedAst = [System.Management.Automation.Language.Parser]::ParseInput($combined, [ref]$checkTokens, [ref]$checkErrors)
if ($checkErrors.Count -gt 0) {
    throw "結合した .ps1 の構文エラー: $($checkErrors[0].Message)（$($checkErrors[0].Extent.StartLineNumber) 行目）"
}

# 2) 読み込み口からたどれるファイルが、もれなく・重複なく 1 回ずつ入っている
$reachable = getReachableFiles -entries @($guiPath, $indexerPath) -tebunkoDir $tebunkoDir
$missing = @($reachable | Where-Object { -not $script:visited.Contains($_) })
if ($missing.Count -gt 0) {
    throw "結合から漏れたファイルがあります: $($missing -join ', ')"
}
$extra = @($script:visited | Where-Object { -not $reachable.Contains($_) -and $_ -ne $guiPath -and $_ -ne $indexerPath })
if ($extra.Count -gt 0) {
    throw "結合に、読み込み口からたどれない余計なファイルが入っています: $($extra -join ', ')"
}

# 3) 読み込み行（$PSScriptRoot・$TebunkoDir の dot-source）が結果に残っていない
if ($combined -match '(?m)^[ \t]*\.\s+"\$(PSScriptRoot|TebunkoDir)\\[^"]+"[ \t]*$') {
    throw "結合した .ps1 に、差し替え忘れの読み込み行が残っています。"
}

# 4) 起動口（gui.ps1・indexer.ps1）の param の名前が、結合した頭の param にすべてある
function getEntryParamNames {
    param ([string]$path)
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
    if (!$ast.ParamBlock) { return @() }
    return @($ast.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
}
$headerNames = @($combinedAst.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
$entryNames = @(getEntryParamNames $guiPath) + @(getEntryParamNames $indexerPath)
$missingParams = @($entryNames | Where-Object { $headerNames -notcontains $_ } | Sort-Object -Unique)
if ($missingParams.Count -gt 0) {
    throw "起動口の param が、結合の頭の param にありません: $($missingParams -join ', ')"
}

# 5) 禁止の語が無い
$banned = findBannedCode $combined
if ($banned) {
    throw "結合した .ps1 に禁止の語が見つかりました: $banned"
}

# 6) 版の文字列が引数と同じ
if ($combined -notmatch [regex]::Escape("Tag = '$Version'; Sha = '$sha'")) {
    throw "版の文字列が埋め込めていません。"
}

# 7) 部品（${bundledParts}）が、入り口（lib.ps1・indexer_lib.ps1 + indexer_main.ps1）からたどれるファイルと
#    過不足なく一致し、画面（ui\ 配下）のファイルを含まない
foreach ($partEntry in $bundledPartEntries) {
    $partVisitedSet = $script:partVisited[$partEntry.Name]
    $partReachable = getReachableFiles -entries $partEntry.Paths
    $missingPart = @($partReachable | Where-Object { -not $partVisitedSet.Contains($_) })
    if ($missingPart.Count -gt 0) {
        throw "部品（$($partEntry.Name)）に漏れたファイルがあります: $($missingPart -join ', ')"
    }
    $extraPart = @($partVisitedSet | Where-Object { -not $partReachable.Contains($_) })
    if ($extraPart.Count -gt 0) {
        throw "部品（$($partEntry.Name)）に、入り口からたどれない余計なファイルが入っています: $($extraPart -join ', ')"
    }
    $uiFiles = @($partVisitedSet | Where-Object { $_ -match '\\ui\\' })
    if ($uiFiles.Count -gt 0) {
        throw "部品（$($partEntry.Name)）に画面（ui\）のファイルが含まれています: $($uiFiles -join ', ')"
    }
}

# ---- 書き出し（BOM 付き UTF-8・CRLF。work\release\ は .gitignore でコミットしない） ----

$outDir = Split-Path $OutFile -Parent
if (!(Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
}
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($OutFile, $combined, $utf8Bom)

Write-Host "作成しました: $OutFile"
return $OutFile
