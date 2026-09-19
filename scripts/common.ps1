# 各スクリプト共通のパス定義と関数。スクリプト・テストから dot-source して使う。

# フォルダ構成
${rootDir}   = Split-Path $PSScriptRoot -Parent
${workDir}   = "${rootDir}\work"
${indexDir}  = "${workDir}\index"
# Excelは [ ] を含むパスに保存できないため TEMP を使う。
# 変換を同時に複数実行しても互いのTSVを削除・移動しないよう、プロセスごとに分ける
${tmpDir}    = Join-Path ([System.IO.Path]::GetTempPath()) "win_grep\${PID}"
${outputDir} = "${rootDir}\output"

${utf8Bom} = New-Object System.Text.UTF8Encoding($true)

# インデックスのTSVで、セル内改行の代わりに使う文字（U+2028 LINE SEPARATOR）。
# TSVの1行 = Excelの1行を保つため、セル内改行はこの文字に置き換えて保存し、検索結果の出力時に改行へ戻す
${cellNewLine} = [string][char]0x2028

# 設定ファイル（画面が読み書きする。変換処理は変換対象フォルダを読む）。内容は JSON
${settingsFile} = "${rootDir}\setting.config"
# 以前の設定ファイル（設定ファイルと同じフォルダの config\*.txt）。設定ファイルが無いときだけ読み込んで移す
${legacyConfigDirName}        = "config"
${legacyTargetFolderFileName} = "変換対象フォルダパス.txt"
${legacyIndexPathFileName}    = "検索対象インデックスパス.txt"
${legacySourceReplaceFileName} = "元のフォルダの置き換え.txt"
${legacySearchOptionFileName} = "検索オプション.txt"

# インデックスのフォルダに置く、インデックス名と変換対象フォルダの対応（変換処理が作成する）。
# インデックスのフォルダごと別の場所・PCへコピーしても、検索結果から元のファイルの場所が分かるようにする。
# 拡張子を .tsv にすると検索対象になるため .txt にする
${sourceFolderFileName} = "元のフォルダ.txt"

# 変換一覧・出力（自動生成）
${statusFile} = "${workDir}\変換一覧.tsv"
${convertingFile} = "${workDir}\変換中.txt"  # 変換中のファイル。強制終了で残っていれば、そのファイルの変換中に止まった
${resultFile} = "${outputDir}\検索結果.txt"

# 画面（config_gui.ps1）と変換処理（office_to_tsv.ps1）の受け渡し
${stopRequestFile}  = "${workDir}\変換中止要求"    # 画面が作成すると、変換処理はファイルの切れ目で中止する
${convertErrorFile} = "${workDir}\変換エラー.txt"  # 変換処理を続けられないエラーのメッセージ（正常終了時は削除）
${convertLogFile}   = "${workDir}\変換ログ.txt"    # 変換処理の表示内容の記録（実行ごとに上書き）

# 変換一覧の列と状態
${statusColumns}   = @("相対パス", "更新日時", "サイズ", "状態", "TSV数", "変換日時", "エラー")
${statusFolderKey} = "変換対象フォルダ"
${stateNew}    = "未変換"
${stateDone}   = "済"
${stateFailed} = "失敗"

# 検索結果から元のファイルを開くときの開き方（設定 openMode の値）
${openModeNormal}   = "normal"    # そのまま開く（編集する）
${openModeReadOnly} = "readOnly"  # 読み取り専用で開く（誤って上書きしない）
${openModeNew}      = "new"       # 新規（元のファイルを基にした無題の文書）で開く。元のファイルを占有しない
${openModes}        = @(${openModeNormal}, ${openModeReadOnly}, ${openModeNew})

function newSettings {
    # 設定の既定値。設定ファイル（JSON）のキーと同じ
    return [ordered]@{
        targetFolders      = @()      # 変換対象フォルダ: @{ name（インデックス名）; path（今フォルダが置かれている場所）; enabled }（記載順。enabled が false は登録のみで変換しない）
        indexSources       = @()      # 変換しないインデックスの元のフォルダ: @{ name; path }（別のPC・場所で作ったインデックスを検索するとき）
        indexFolders       = @()      # 検索対象インデックスのフォルダ（空なら work\index）
        searchExcludes     = @()      # 画面の検索対象ツリーでチェックを外したフォルダ: @{ path（フルパス）; subfolders（false はフォルダ直下のファイルだけ） }
        useRegex           = $false   # 検索ワードを正規表現として扱う
        caseSensitive      = $false   # 英字の大文字と小文字を区別する
        fileFilter         = ""       # 対象ファイル（元のファイル名のワイルドカード。; 区切り、! で始まるものは除外。空ならすべて）
        openMode           = ${openModeNormal}  # 検索結果の元のファイルの開き方: 通常（編集する）/ 読み取り専用 / 新規（元のファイルを基にした無題の文書。占有しない）
    }
}

function readSettings {
    # 設定を newSettings と同じ形で返す。記載の無い項目は既定値。
    # 設定ファイルが無ければ、同じフォルダの config\ にある以前の設定ファイル（*.txt）から移して保存する（それも無ければ既定値）
    param (
        [string]$path = ${settingsFile}
    )

    $settings = newSettings
    if (!(Test-Path -LiteralPath $path)) {
        if (readLegacySettings $settings (Join-Path ([System.IO.Path]::GetDirectoryName($path)) ${legacyConfigDirName})) {
            writeSettings $settings $path
        }
        return $settings
    }

    $json = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
    if ($json.Trim() -eq "") {
        return $settings
    }
    try {
        $data = ConvertFrom-Json $json
    } catch {
        throw "$([System.IO.Path]::GetFileName($path)) を読み込めません。（$($_.Exception.Message)）"
    }
    foreach ($key in @($settings.Keys)) {
        $property = $data.PSObject.Properties[$key]
        if ($null -eq $property -or $null -eq $property.Value) {
            continue
        }
        if ($settings[$key] -is [array]) {
            $settings[$key] = @($property.Value | Where-Object { $null -ne $_ })
        } elseif ($settings[$key] -is [string]) {
            $settings[$key] = [string]$property.Value
        } else {
            $settings[$key] = [bool]$property.Value
        }
    }
    return $settings
}

function writeSettings {
    param (
        $settings,
        [string]$path = ${settingsFile}
    )

    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
    [System.IO.File]::WriteAllText($path, (ConvertTo-Json -InputObject $settings -Depth 5), (New-Object System.Text.UTF8Encoding($false)))
}

function updateSettings {
    # 設定ファイルを読み直して key の値だけを変えて保存する（ほかの項目は、ほかの画面・処理が保存した内容を保つ）
    param (
        [string]$key,
        $value,
        [string]$path = ${settingsFile}
    )

    $settings = readSettings $path
    $settings[$key] = $value
    writeSettings $settings $path
}

function readLegacySettings {
    # 以前の設定ファイル（1行1件のテキスト）を settings に読み込む。1つも無ければ $false
    param (
        $settings,
        [string]$dir
    )

    $found = $false
    $file = Join-Path $dir ${legacyTargetFolderFileName}
    if (Test-Path -LiteralPath $file) {
        $found = $true
        # 行頭が # の行はチェックなし
        # インデックス名は以前の設定ファイルには無いため空にする（変換時に割り当てる。assignIndexNames）
        $settings.targetFolders = @(readListFile $file | ForEach-Object { $_.Trim() } | ForEach-Object {
            [pscustomobject]@{ name = ""; path = (normalizeFolderPath $_.TrimStart("#")); enabled = -not $_.StartsWith("#") }
        } | Where-Object { $_.path -ne "" })
    }
    $file = Join-Path $dir ${legacyIndexPathFileName}
    if (Test-Path -LiteralPath $file) {
        $found = $true
        $settings.indexFolders = @(readListFile $file | ForEach-Object { $_.Trim() })
    }
    $file = Join-Path $dir ${legacySearchOptionFileName}
    if (Test-Path -LiteralPath $file) {
        $found = $true
        # "正規表現=オン" の行
        $settings.useRegex = @(readListFile $file | Where-Object { $_ -match "^\s*正規表現\s*=\s*オン\s*$" }).Count -gt 0
    }
    return $found
}

function readListFile {
    # 1行1件のファイルを読み込む（空行を除く）。ファイルが無ければ空配列
    param (
        [string]$path
    )

    if (!(Test-Path -LiteralPath $path)) {
        return @()
    }
    return @(Get-Content -LiteralPath $path -Encoding UTF8 | Where-Object { $_.Trim() -ne "" })
}

function writeListFile {
    param (
        [string]$path,
        [string[]]$lines
    )

    if ($null -eq $lines) {
        $lines = [string[]]@()
    }
    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
    [System.IO.File]::WriteAllLines($path, $lines, ${utf8Bom})
}

function formatFileTime {
    # 変換一覧に記録する日時の書式（秒まで）。更新の有無はこの文字列で比べる
    param (
        [datetime]$time
    )

    return $time.ToString("yyyy/MM/dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
}

function newStatusRow {
    # 変換一覧の1行（1ファイル）を作る
    param (
        [string]$relPath,
        [string]$updated = "",
        [string]$size = "",
        [string]$state = ${stateNew},
        [string]$tsvCount = "",
        [string]$converted = "",
        [string]$errorMessage = ""
    )

    return [pscustomobject]@{
        相対パス = $relPath
        更新日時 = $updated
        サイズ   = $size
        状態     = $state
        TSV数    = $tsvCount
        変換日時 = $converted
        エラー   = $errorMessage
    }
}

function toStatusLine {
    # 変換一覧の1行を文字列にする。エラーメッセージ等のタブ・改行は、1行1件を保つためスペースにする
    param (
        $row
    )

    return (@(${statusColumns} | ForEach-Object { [string]$row.$_ -replace "[\t\r\n]+", " " }) -join "`t")
}

function describeConvertError {
    # 変換で発生した例外から、変換一覧のエラー列・画面に表示する失敗の理由を返す。
    # Officeアプリのメッセージは分かりにくい（パスワード付きでも「入力したパスワードが間違っています」等）ため、
    # よくある原因は「原因（詳細: 元のメッセージ）」の形に言い換える
    param (
        [System.Exception]$exception
    )

    # メソッド呼び出しの例外（「"7" 個の引数を指定して "Open" を呼び出し中に例外が発生しました」等）は、中の例外が本来の理由
    $base = $exception.GetBaseException()
    $message = ($base.Message -replace "\s+", " ").Trim()
    $code = "{0:X8}" -f $base.HResult
    if ($message -eq "") {
        $message = "エラーコード 0x${code}"
    }

    # スクリプト自身が throw したメッセージは、利用者向けに書いてあるためそのまま使う
    if ($base.GetType() -eq [System.Management.Automation.RuntimeException]) {
        return $message
    }

    if ($message -match "パスワード|password") {
        # 元のメッセージ（パスワードが間違っています等）は、パスワードを入力していない利用者には誤解を招くため付けない
        return "読み取りパスワードが設定されているため開けません（パスワード付きのファイルは変換できません）"
    }

    $cause = $null
    if ($base -is [System.IO.FileNotFoundException] -or $base -is [System.IO.DirectoryNotFoundException]) {
        $cause = "ファイルが見つかりません（変換中に移動・削除・名前変更された可能性があります）"
    } elseif ($base -is [System.UnauthorizedAccessException] -or $code -eq "80070005") {
        $cause = "ファイルを読むアクセス権がありません"
    } elseif ($base -is [System.IO.PathTooLongException]) {
        $cause = "パスが長すぎるため読めません"
    } elseif ($base -is [System.OutOfMemoryException]) {
        # 巨大なシート（テキストにして約 1GB 超）は、整形（prettyTsv）で一度に読み込めずメモリ不足になる
        $cause = "シート・文書が大きすぎて変換できません（メモリが不足しました）"
    } elseif (@("80070020", "80070021") -contains $code) {
        # 共有違反・ロック違反
        $cause = "ほかのアプリ・利用者がファイルを使用中のため読めません（ファイルを閉じてから再変換してください）"
    } elseif (@("80040154", "80080005", "800401F3") -contains $code) {
        # クラス未登録・サーバーの起動失敗・ProgID 不正
        $cause = "Officeアプリ（Excel・Word・PowerPoint）を起動できませんでした（インストール・ライセンス認証の状態を確認してください）"
    } elseif (@("800706BA", "800706BE", "80010105", "80010108") -contains $code) {
        # RPC サーバーを利用できない・呼び出し失敗・サーバーで例外・切断
        $cause = "Officeアプリが異常終了したか、内部でエラーが発生しました（ファイルが壊れている、または大きすぎる可能性があります）"
    } elseif (@("80010001", "8001010A") -contains $code) {
        # 呼び出しの拒否・処理中のため後で再試行
        $cause = "Officeアプリが応答しませんでした（Officeアプリでダイアログを表示中などの可能性があります）"
    } elseif ($base -is [System.IO.InvalidDataException] -or $base -is [System.Xml.XmlException] -or
              $message -match "ファイル形式|破損|壊れ|corrupt|file format") {
        $cause = "ファイルが壊れているか、拡張子と中身の形式が一致していません"
    }

    if ($cause) {
        return "${cause}（詳細: ${message}）"
    }
    return $message
}

function readStatusFile {
    # 変換一覧を読み込み、@{ Folders; Rows } を返す。ファイルが無ければ空。
    #   Folders: 変換対象フォルダ @{ Path; Name（インデックス名） } の配列。以前の形式（インデックス名なし）は Name が空
    #   Rows   : 相対パス（"インデックス名\フォルダからの相対パス"。大文字・小文字を区別しない）→ 行
    # 変換中は1件ごとに行を追記するため、同じ相対パスの行は後の行を優先する。列数の合わない行（書き込み途中で中断した行など）は無視する
    param (
        [string]$path = ${statusFile}
    )

    $rows = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    $folders = New-Object System.Collections.Generic.List[object]
    $result = @{ Folders = $folders; Rows = $rows }
    if (!(Test-Path -LiteralPath $path)) {
        return $result
    }

    # 変換中に画面などから読んでも、変換側の追記・置き換えを妨げないよう共有を許して開く
    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $stream = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
    $reader = New-Object System.IO.StreamReader($stream, ${utf8Bom})
    try {
        $lines = $reader.ReadToEnd() -split "\r?\n"
    } finally {
        $reader.Dispose()
    }

    $header = ${statusColumns} -join "`t"
    foreach ($line in $lines) {
        $fields = $line.Split("`t")
        if ($fields[0] -eq ${statusFolderKey} -and ($fields.Count -eq 2 -or $fields.Count -eq 3)) {
            $folders.Add([pscustomobject]@{ Path = $fields[1]; Name = $(if ($fields.Count -eq 3) { $fields[2] } else { "" }) })
            continue
        }
        if ($line -eq $header -or $fields.Count -ne ${statusColumns}.Count -or $fields[0] -eq "") {
            continue
        }
        $rows[$fields[0]] = newStatusRow @fields
    }
    return $result
}

function writeStatusFile {
    # 変換一覧を書き出す（1ファイル1行）。途中で中断しても壊れないよう、一時ファイルに書いてから置き換える
    #   folders: 変換対象フォルダ @{ Path; Name } の配列。"変換対象フォルダ<TAB>パス<TAB>インデックス名" の行にする
    param (
        [object[]]$folders,
        [object[]]$rows,
        [string]$path = ${statusFile}
    )

    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($folder in @($folders | Where-Object { $_ })) {
        $lines.Add("${statusFolderKey}`t$($folder.Path)`t$($folder.Name)")
    }
    $lines.Add((${statusColumns} -join "`t"))
    foreach ($row in @($rows | Where-Object { $null -ne $_ })) {
        $lines.Add((toStatusLine $row))
    }

    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path)) | Out-Null
    $tmpPath = "${path}.tmp"
    [System.IO.File]::WriteAllLines($tmpPath, $lines, ${utf8Bom})
    if (Test-Path -LiteralPath $path) {
        # $null は空文字列として渡されて例外になるため、[NullString]::Value（バックアップを作らない）を渡す
        [System.IO.File]::Replace($tmpPath, $path, [NullString]::Value)
    } else {
        [System.IO.File]::Move($tmpPath, $path)
    }
}

function addStatusRow {
    # 変換一覧の末尾に1行追記する（readStatusFile では後の行が優先される）
    param (
        $row,
        [string]$path = ${statusFile}
    )

    [System.IO.File]::AppendAllText($path, "$(toStatusLine $row)`r`n", ${utf8Bom})
}

function readConvertingFile {
    # 変換中のファイルの記録を読み、@{ RelPath = 相対パス; Count = 続けて変換を始めて終わらなかった回数 } を返す。
    # 記録が無い・壊れている場合は $null
    param (
        [string]$path = ${convertingFile}
    )

    $lines = @(readListFile $path)
    if ($lines.Count -eq 0) {
        return $null
    }
    $fields = $lines[0].Split("`t")
    $count = 0
    if ($fields.Count -ne 2 -or -not [int]::TryParse($fields[0], [ref]$count) -or $count -lt 1 -or $fields[1] -eq "") {
        return $null
    }
    return @{ RelPath = $fields[1]; Count = $count }
}

function writeConvertingFile {
    # 変換を始めるファイルを "回数<TAB>相対パス" で記録する。変換が終われば removeConvertingFile で消す
    param (
        [string]$relPath,
        [int]$count,
        [string]$path = ${convertingFile}
    )

    writeListFile $path @("${count}`t${relPath}")
}

function removeConvertingFile {
    param (
        [string]$path = ${convertingFile}
    )

    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
    }
}

function normalizeFolderPath {
    # フォルダパスを1つの書き方にそろえる（書き方の違いで同じフォルダを別のフォルダとみなさないため）。
    #   ・前後の空白・" を取り除く
    #   ・環境変数（%USERPROFILE% など）を展開する
    #   ・/ を \ にそろえ、長いパス用の \\?\ ・ \\?\UNC\ を外す（ファイル操作に渡す直前に toLongPath で付け直す）
    #   ・重なった \ ・ . ・ .. を解決する
    #   ・相対パスは win_grep のフォルダ（$rootDir）からとみなして絶対パスにする
    #   ・末尾の \ を取り除く（ドライブ直下は "D:\" のまま。UNC の共有直下は "\\server\share"）
    # パスとして解釈できない場合（* ? を含む・共有名の無い \\server など）は、書かれたとおりに扱う
    param (
        [string]$path
    )

    $path = $path.Trim().Trim('"').Trim()
    if ($path -eq "") {
        return ""
    }
    if ($path.IndexOf("%") -ge 0) {
        $path = [System.Environment]::ExpandEnvironmentVariables($path).Trim()
    }
    $path = (fromLongPath ($path.Replace("/", "\"))).TrimEnd("\")
    if ($path -eq "") {
        return ""
    }
    # ドライブ名だけ（"D:"）はドライブ直下とする。
    # GetFullPath はそのドライブの「現在のフォルダ」を返すことがあるため、先に決める
    if ($path -match "^[A-Za-z]:$") {
        return "${path}\"
    }
    try {
        if (![System.IO.Path]::IsPathRooted($path)) {
            # 相対パスは win_grep のフォルダからとみなす
            $path = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine(${rootDir}, $path)).TrimEnd("\")
        } elseif ($path -match "^[A-Za-z]:\\" -or $path.StartsWith("\\")) {
            $path = [System.IO.Path]::GetFullPath($path).TrimEnd("\")
        }
        # \ ひとつで始まるパス（"\server\share"）は、書き間違えた UNC パスのことが多いため、
        # 今のドライブのパスに直さず、書かれたとおりに扱う
    } catch {
        # パスとして解釈できない場合は、書かれたとおりに扱う（存在しないフォルダとして扱われる）
    }
    if ($path -match "^[A-Za-z]:$") {
        return "${path}\"
    }
    return $path
}

function getPathUnderFolder {
    # path が folder 自身か folder の下なら、folder からの相対パス（folder 自身は ""）を返す。
    # folder の下でなければ $null（大文字・小文字は区別しない）。
    # パスの文字数で切り出すと、書き方が少し違うだけ（末尾の \ ・\\?\ 付きなど）で取り違えるため、この関数を通す
    param (
        [string]$path,
        [string]$folder
    )

    $folder = $folder.TrimEnd("\")
    $path = $path.TrimEnd("\")
    if ($folder -eq "" -or $path -eq "") {
        return $null
    }
    if ($path.Equals($folder, [System.StringComparison]::OrdinalIgnoreCase)) {
        return ""
    }
    if ($path.StartsWith("${folder}\", [System.StringComparison]::OrdinalIgnoreCase)) {
        return $path.Substring($folder.Length + 1)
    }
    return $null
}

# ドライブ文字の割り当て先を調べる Win32 API（QueryDosDevice）。
# net use・WMI と違い、subst で割り当てたドライブも分かり、切断中のネットワークドライブでも割り当てが分かる
${dosDeviceSource} = @'
using System;
using System.Text;
using System.Runtime.InteropServices;

namespace WinGrep
{
    public static class DosDevice
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        static extern uint QueryDosDeviceW(string lpDeviceName, StringBuilder lpTargetPath, uint ucchMax);

        // ドライブ文字（"Z:"）の割り当て先（デバイス名）を返す。割り当てが無ければ null
        public static string Query(string drive)
        {
            StringBuilder buffer = new StringBuilder(1024);
            if (QueryDosDeviceW(drive, buffer, (uint)buffer.Capacity) == 0) return null;
            return buffer.ToString();
        }
    }
}
'@

function parseDosDeviceTarget {
    # QueryDosDevice の結果を通常のパスにする。ネットワークドライブ・subst 以外（実際のディスク）は ""
    #   subst               : "\??\C:\data"                                              → "C:\data"
    #   ネットワークドライブ: "\Device\LanmanRedirector\;Z:0000000000012345\server\share" → "\\server\share"
    #                         "\Device\Mup\;Z:0000000000012345\server\share"              → "\\server\share"
    param (
        [string]$target
    )

    if ($target.StartsWith("\??\")) {
        return $target.Substring(4)
    }
    $i = $target.IndexOf("\;")
    if ($i -lt 0) {
        return ""
    }
    # "\;" の後は "<ドライブ文字>:<16進数>\server\share"
    $rest = $target.Substring($i + 2)
    $j = $rest.IndexOf("\")
    if ($j -lt 0) {
        return ""
    }
    return "\\" + $rest.Substring($j + 1)
}

${driveTargets} = $null

function getDriveTargets {
    # ドライブ文字（"Z:"）→ 割り当て先（ネットワークドライブは "\\server\share"、subst は割り当て元のフォルダ）を返す。
    # 同じプロセスでは1回だけ調べる（変換・検索の途中で割り当てが変わることは想定しない）
    if ($null -ne ${script:driveTargets}) {
        return ${script:driveTargets}
    }
    $map = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
    try {
        if (!("WinGrep.DosDevice" -as [type])) {
            Add-Type -TypeDefinition ${dosDeviceSource}
        }
        foreach ($letter in [char[]]"ABCDEFGHIJKLMNOPQRSTUVWXYZ") {
            $target = [WinGrep.DosDevice]::Query("${letter}:")
            if (!$target) {
                continue
            }
            $folder = (parseDosDeviceTarget $target).TrimEnd("\")
            if ($folder -ne "") {
                $map["${letter}:"] = $folder
            }
        }
    } catch {
        # 調べられない環境では別名なしとする（パスを書かれたとおりに使う）
    }
    ${script:driveTargets} = $map
    return $map
}

function getFolderPathAliases {
    # 同じ場所を指す別の書き方を、path 自身を先頭にして返す（ドライブの割り当てをたどる）。
    # ファイルサーバーのフォルダは、ネットワークドライブ（Z:\…）と UNC パス（\\server\share\…）のどちらでも書けるため、
    # 設定に書いた書き方と、インデックスに記録した書き方が違っても同じフォルダと分かるようにする。
    #   Z: が \\server\share のネットワークドライブのとき
    #     "Z:\見積"             → @("Z:\見積", "\\server\share\見積")
    #     "\\server\share\見積" → @("\\server\share\見積", "Z:\見積")
    param (
        [string]$path,
        $drives = (getDriveTargets)  # ドライブ文字 → 割り当て先（テストで差し替える）
    )

    $result = New-Object System.Collections.Generic.List[string]
    $result.Add($path)
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    [void]$seen.Add($path.TrimEnd("\"))

    foreach ($entry in $drives.GetEnumerator()) {
        $alias = $null
        # ドライブ文字 → 割り当て先
        $rest = getPathUnderFolder $path $entry.Key
        if ($null -ne $rest) {
            $alias = if ($rest -eq "") { $entry.Value } else { joinSourcePath $entry.Value $rest }
        } else {
            # 割り当て先 → ドライブ文字
            $rest = getPathUnderFolder $path $entry.Value
            if ($null -ne $rest) {
                $alias = if ($rest -eq "") { "$($entry.Key)\" } else { joinSourcePath "$($entry.Key)\" $rest }
            }
        }
        if ($alias -and $seen.Add($alias.TrimEnd("\"))) {
            $result.Add($alias)
        }
    }
    return $result.ToArray()
}

function testSameFolder {
    # 2つのパスが同じフォルダを指すか（末尾の \ ・大文字と小文字の違いと、ネットワークドライブ・subst の割り当てをたどる）
    param (
        [string]$a,
        [string]$b,
        $drives = (getDriveTargets)  # ドライブ文字 → 割り当て先（テストで差し替える）
    )

    $target = $b.TrimEnd("\")
    foreach ($alias in @(getFolderPathAliases $a $drives)) {
        if ($alias.TrimEnd("\").Equals($target, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }
    return $false
}

function getTargetFolders {
    # 変換対象フォルダの一覧（記載順）を返す: @{ Name; Path; Enabled }。
    #   Name   : インデックス名（work\index 直下のフォルダ名）。インデックスの「名前」で、フォルダの置き場所（Path）とは分けて持つ。
    #            Path を書き換えても Name が同じなら同じインデックスとして扱う（変換し直さない）。空なら変換時に割り当てる（assignIndexNames）
    #   Path   : そのフォルダが今置かれている場所
    #   Enabled: false はチェックなし（登録のみで変換しない）
    # 同じフォルダ・同じ名前は最初のものだけ使う（名前の重複は、2 つ目以降を空にして割り当て直す）。
    # 書き方が違うだけで同じフォルダを指す場合（ネットワークドライブと UNC パスなど）も同じフォルダとみなす
    param (
        [string]$path = ${settingsFile}
    )

    $folders = New-Object System.Collections.Generic.List[object]
    $seenPath = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $seenName = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($item in @((readSettings $path).targetFolders)) {
        $folder = normalizeFolderPath ([string]$item.path)
        if ($folder -eq "" -or -not $seenPath.Add($folder)) {
            continue
        }
        if (@($folders | Where-Object { testSameFolder $_.Path $folder }).Count -gt 0) {
            continue
        }
        $name = toSafeFileName ([string]$item.name).Trim()
        if ($name -ne "" -and -not $seenName.Add($name)) {
            $name = ""
        }
        $folders.Add([pscustomobject]@{ Name = $name; Path = $folder; Enabled = ($item.enabled -ne $false) })
    }
    return $folders.ToArray()
}

function writeTargetFolders {
    # 変換対象フォルダの一覧（@{ Name; Path; Enabled } の配列）を保存する
    param (
        [object[]]$folders,
        [string]$path = ${settingsFile}
    )

    updateSettings "targetFolders" ([object[]]@($folders | Where-Object { $_ } | ForEach-Object {
        [pscustomobject]@{ name = [string]$_.Name; path = $_.Path; enabled = [bool]$_.Enabled }
    })) $path
}

function readIndexSources {
    # 変換しないインデックスの元のフォルダ（indexSources）を @{ Name; Path } の配列で返す。
    # 別の PC・場所で作ったインデックスを検索するとき、そのインデックス名の元のフォルダを覚えておくために使う
    param (
        [string]$path = ${settingsFile}
    )

    $items = New-Object System.Collections.Generic.List[object]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($item in @((readSettings $path).indexSources)) {
        $name = ([string]$item.name).Trim()
        $folder = normalizeFolderPath ([string]$item.path)
        if ($name -eq "" -or $folder -eq "" -or -not $seen.Add($name)) {
            continue
        }
        $items.Add([pscustomobject]@{ Name = $name; Path = $folder })
    }
    return $items.ToArray()
}

function writeIndexSources {
    param (
        [object[]]$sources,
        [string]$path = ${settingsFile}
    )

    updateSettings "indexSources" ([object[]]@($sources | Where-Object { $_ } | ForEach-Object {
        [pscustomobject]@{ name = $_.Name; path = $_.Path }
    })) $path
}

function setIndexSourceFolder {
    # インデックス名に対する元のフォルダ（今そのフォルダが置かれている場所）を設定に記録する。
    # 変換対象フォルダにある名前ならそのパスを書き換え、無ければ indexSources に記録する。
    # 「名前」と「置き場所」を分けて持つため、フォルダを移した場合もこの 1 か所を書き換えるだけで済む
    param (
        [string]$name,
        [string]$folder,
        [string]$path = ${settingsFile}
    )

    $name = ([string]$name).Trim()
    $folder = normalizeFolderPath $folder
    if ($name -eq "" -or $folder -eq "") {
        return
    }

    $targets = @(getTargetFolders $path)
    if (@($targets | Where-Object { $_.Name -eq $name }).Count -gt 0) {
        writeTargetFolders @($targets | ForEach-Object {
            if ($_.Name -eq $name) { [pscustomobject]@{ Name = $_.Name; Path = $folder; Enabled = $_.Enabled } } else { $_ }
        }) $path
        return
    }

    $sources = @(@(readIndexSources $path | Where-Object { $_.Name -ne $name }) + @([pscustomobject]@{ Name = $name; Path = $folder }))
    writeIndexSources $sources $path
}

function getFolderLeafName {
    # フォルダ名（ドライブ直下はドライブ名、UNC の共有直下は共有名）を返す
    param (
        [string]$folderPath
    )

    $path = $folderPath.TrimEnd("\")
    $leaf = [System.IO.Path]::GetFileName($path)
    if ($leaf -eq "") {
        $leaf = $path.TrimEnd(":")
    }
    return $leaf
}

function newIndexName {
    # 変換対象フォルダのインデックス名（work\index 直下のフォルダ名）を作る。
    # フォルダ名（ドライブ直下はドライブ名）を使い、usedNames と重複すれば「名前(2)」「名前(3)」…とする
    param (
        [string]$folderPath,
        $usedNames  # HashSet[string]（大文字・小文字を区別しない）
    )

    $base = toSafeFileName (getFolderLeafName $folderPath)
    if ($base -eq "") {
        $base = "フォルダ"
    }

    $name = $base
    for ($i = 2; $usedNames.Contains($name); $i++) {
        $name = "${base}(${i})"
    }
    return $name
}

function assignIndexNames {
    # 変換対象フォルダ（getTargetFolders）にインデックス名を割り当て、@{ Path; Enabled; Name } の配列を返す。
    # インデックス名は設定に持つ（getTargetFolders の Name）。フォルダの置き場所（Path）を書き換えても名前は変わらないため、
    # フォルダを移しても同じインデックスとして扱える（インデックスを作り直さない）。
    # 名前が無い場合（新しく追加したフォルダ・以前の版の設定）は、前回の変換一覧の同じパスの名前を使い、
    # それも無ければフォルダ名から重複しない名前を作る
    param (
        [object[]]$targetFolders,
        [object[]]$previousFolders  # readStatusFile の Folders（@{ Path; Name }）
    )

    $previousNames = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($folder in @($previousFolders | Where-Object { $_ -and $_.Name })) {
        $previousNames[$folder.Path] = $folder.Name
    }

    # 設定にある名前・前回の名前は使わない（別のフォルダに同じ名前を付けてインデックスを取り違えないため）
    $used = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($name in $previousNames.Values) {
        [void]$used.Add($name)
    }
    foreach ($folder in @($targetFolders | Where-Object { $_ -and $_.Name })) {
        [void]$used.Add($folder.Name)
    }

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($folder in @($targetFolders | Where-Object { $_ })) {
        $name = [string]$folder.Name
        if ($name -eq "") {
            if ($previousNames.ContainsKey($folder.Path)) {
                $name = $previousNames[$folder.Path]
            } else {
                $name = newIndexName $folder.Path $used
            }
            [void]$used.Add($name)
        }
        $result.Add([pscustomobject]@{ Path = $folder.Path; Enabled = $folder.Enabled; Name = $name })
    }
    return $result.ToArray()
}

function splitIndexRelPath {
    # work\index からの相対パスを、先頭のインデックス名と残り（変換対象フォルダからの相対パス）に分ける: @{ Name; Rest }
    # （String.Split([char], 2) は .NET Framework では Split(params char[]) になり、2 も区切り文字とみなされるため使わない）
    param (
        [string]$relPath
    )

    $i = $relPath.IndexOf("\")
    if ($i -lt 0) {
        return @{ Name = $relPath; Rest = "" }
    }
    return @{ Name = $relPath.Substring(0, $i); Rest = $relPath.Substring($i + 1) }
}

function getIndexNameMap {
    # 変換一覧に記録したインデックス名 → 変換対象フォルダのパス（大文字・小文字を区別しない）。
    # 変換対象フォルダの行は先頭にあるため、見出し行まで読んで打ち切る（変換一覧が大きくても時間がかからないように）
    param (
        [string]$path = ${statusFile}
    )

    $map = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
    if (!(Test-Path -LiteralPath $path)) {
        return , $map
    }
    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $stream = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
    $reader = New-Object System.IO.StreamReader($stream, ${utf8Bom})
    try {
        $header = ${statusColumns} -join "`t"
        while ($null -ne ($line = $reader.ReadLine())) {
            if ($line -eq $header) {
                break
            }
            $fields = $line.Split("`t")
            if ($fields[0] -eq ${statusFolderKey} -and $fields.Count -eq 3 -and $fields[2]) {
                $map[$fields[2]] = $fields[1]
            }
        }
    } finally {
        $reader.Dispose()
    }
    return , $map
}

function readIndexFolders {
    # 設定した検索対象インデックスのフォルダ一覧を返す（設定が無ければ空）
    param (
        [string]$path = ${settingsFile}
    )

    return @((readSettings $path).indexFolders | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ -ne "" })
}

function writeIndexFolders {
    # 検索対象インデックスのフォルダ一覧を保存する（空なら work\index を使う）
    param (
        [string[]]$folders,
        [string]$path = ${settingsFile}
    )

    updateSettings "indexFolders" ([object[]]@($folders | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ -ne "" })) $path
}

function getIndexFolders {
    # 検索対象インデックスのフォルダ一覧を返す。設定が無い・空なら work\index を使う。
    param (
        [string]$path = ${settingsFile}
    )

    $folders = @(readIndexFolders $path)
    if ($folders.Count -gt 0) {
        return $folders
    }
    return @(${indexDir})
}

function readSearchExcludes {
    # 画面の検索対象ツリーでチェックを外したフォルダを @{ Path; Subfolders } の配列で返す（設定が無ければ空 = すべて検索する）。
    #   Subfolders: $true はフォルダ以下すべて、$false はフォルダ直下のファイルだけを外す
    param (
        [string]$path = ${settingsFile}
    )

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($item in @((readSettings $path).searchExcludes)) {
        $folder = ([string]$item.path).Trim().TrimEnd("\")
        if ($folder -eq "") {
            continue
        }
        $subfolders = if ($null -eq $item.subfolders) { $true } else { [bool]$item.subfolders }
        $result.Add([pscustomobject]@{ Path = $folder; Subfolders = $subfolders })
    }
    return $result.ToArray()
}

function writeSearchExcludes {
    # 検索対象ツリーでチェックを外したフォルダ（@{ Path; Subfolders } の配列）を保存する
    param (
        [object[]]$excludes,
        [string]$path = ${settingsFile}
    )

    updateSettings "searchExcludes" ([object[]]@($excludes | Where-Object { $_ -and $_.Path } | ForEach-Object {
        [ordered]@{ path = [string]$_.Path; subfolders = [bool]$_.Subfolders }
    })) $path
}

function toSafeFileName {
    # ファイル名に使えない文字を全角に変換
    param (
        [string]$name
    )

    $name = $name.Replace(">", "＞")
    $name = $name.Replace("<", "＜")
    $name = $name.Replace("\", "￥")
    $name = $name.Replace("*", "＊")
    $name = $name.Replace('"', '”')
    $name = $name.Replace(":", "：")
    $name = $name.Replace("?", "？")
    $name = $name.Replace("|", "｜")
    $name = $name.Replace("/", "／")

    return $name
}

${maxFileNameLength} = 255  # Windows のファイル名（パスの区切りの間の1つ）の上限

function encodeIndexPlace {
    # インデックスのTSVのファイル名に入れる場所（シート名・ページ・スライド）を符号化する。
    # ファイル名に使えない文字・制御文字と、区切りの _・符号化に使う % を "%XX"（16進数）にする（decodeIndexPlace で元に戻す）。
    #   ・全角に置き換えると `a"b` と `a”b` が同じファイル名になり、後のシートで上書きされるため、元に戻せる形にする
    #   ・場所に _ が残らないため、ファイル名に .xlsx_ 等を含むブックでも、最後の _ でファイル名と場所に分けられる
    param (
        [string]$place
    )

    return [regex]::Replace($place, '[\x00-\x1F"%*/:<>?\\_|]', { param($m) '%{0:X2}' -f [int][char]$m.Value })
}

function decodeIndexPlace {
    # encodeIndexPlace で符号化した場所を元に戻す（encodeIndexPlace が作る "%XX" だけを戻す）
    param (
        [string]$place
    )

    return [regex]::Replace($place, '%(?:[01][0-9A-F]|2[25AF]|3[ACEF]|5[CF]|7C)', { param($m) [string][char][Convert]::ToInt32($m.Value.Substring(1), 16) })
}

function toIndexFileName {
    # インデックスのTSVのファイル名 "<場所>.tsv" を返す（場所は encodeIndexPlace で符号化する）。
    # 元のファイル名はフォルダ名（= 元のファイル名そのもの）にするため、ファイル名には入れない。
    # ファイル名の上限（255文字）は長いパスの対応（toLongPath）でも超えられないため、超える場合は分かるメッセージで例外にする
    param (
        [string]$place
    )

    $name = "{0}.tsv" -f (encodeIndexPlace $place)
    if ($name.Length -gt ${maxFileNameLength}) {
        throw "変換結果のファイル名が長すぎるため保存できません（$($name.Length) 文字。上限 ${maxFileNameLength} 文字）: ${name}"
    }
    return $name
}

function toLongPath {
    # パスの先頭に \\?\ を付け、260文字を超えるパスもファイル操作（System.IO・-LiteralPath）で扱えるようにする。
    # ネットワークのパス（\\server\share\…）は \\?\UNC\server\share\… にする。付いていればそのまま返す。
    # Join-Path は \\?\ 付きのパスを扱えないため、パスを組み立てた後、ファイル操作に渡す直前に使う
    param (
        [string]$path
    )

    if (!$path -or $path.StartsWith("\\?\")) {
        return $path
    }
    $path = $path.Replace("/", "\")
    if ($path.StartsWith("\\")) {
        return "\\?\UNC\" + $path.Substring(2)
    }
    return "\\?\" + $path
}

function copyFileShared {
    # 元のファイルを占有せずにコピーする（変換は、このコピーを開いて行う）。
    # File.Copy は元のファイルをほかのアプリの書き込みを拒否して開くため、コピーの間は利用者が上書き保存できず、
    # 利用者がファイルを開いて編集中だとコピーできないことがある。
    # ここでは読み取りだけで開き、ほかのアプリの読み書き・削除・名前変更を妨げない。
    # コピー先は通常の属性（読み取り専用などを付けない）で作り、既にあれば上書きする
    param (
        [string]$sourcePath,
        [string]$destPath
    )

    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $source = New-Object System.IO.FileStream((toLongPath $sourcePath), [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
    try {
        $dest = New-Object System.IO.FileStream((toLongPath $destPath), [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        try {
            $source.CopyTo($dest, 1MB)
        } finally {
            $dest.Dispose()
        }
    } finally {
        $source.Dispose()
    }
}

function fromLongPath {
    # toLongPath で付けた \\?\ を外し、通常のパスに戻す（Get-ChildItem の FullName 等から相対パスを求めるため）
    param (
        [string]$path
    )

    if ($path.StartsWith("\\?\UNC\")) {
        return "\\" + $path.Substring(8)
    }
    if ($path.StartsWith("\\?\")) {
        return $path.Substring(4)
    }
    return $path
}

function replaceCellNewLine {
    # ダブルクォートで囲まれた1セル内の改行（LF・CR・CRLF）を、セル内改行を表す文字に置き換える
    param (
        [string]$inputString
    )

    # "" はクォート内のダブルクォートのエスケープだが、"…" "…" の2つに分けてマッチしても結果は同じ
    return [regex]::Replace($inputString, '"[^"]*"', { param($m) $m.Value -replace "\r\n|\r|\n", ${cellNewLine} })
}

function formatTsv {
    # Excelが出力したTSVを、1行目 = Excelの1行目、1列目 = A列 となるように整形する。
    # Excelは使用範囲（UsedRange）の左上のセルから出力するため、その行・列番号を firstRow・firstColumn に渡す。
    # セル内改行は置き換え、行末の空セルと末尾の空行は取り除く（途中の空行は行番号を保つため残す）。
    # 空白以外の文字が無ければ空文字を返す
    param (
        [string]$content,
        [int]$firstRow = 1,
        [int]$firstColumn = 1
    )

    $content = replaceCellNewLine $content
    $columnPadding = "`t" * ($firstColumn - 1)

    $lines = New-Object System.Collections.Generic.List[string]
    for ($i = 1; $i -lt $firstRow; $i++) {
        $lines.Add("")
    }
    foreach ($line in ($content -split "\r?\n")) {
        $lines.Add(($columnPadding + $line).TrimEnd("`t"))
    }
    while ($lines.Count -gt 0 -and $lines[$lines.Count - 1].Trim() -eq "") {
        $lines.RemoveAt($lines.Count - 1)
    }

    return ($lines -join "`r`n")
}

function prettyTsv {
    # Excelが出力したTSV（UTF-16）を整形してUTF-8で保存する。内容が空なら保存せず $false を返す
    #   firstRow, firstColumn: Excelが出力した範囲の左上のセルの行・列番号
    param (
        [string]$inputFilePath,
        [string]$outputFilePath,
        [int]$firstRow = 1,
        [int]$firstColumn = 1
    )

    $content = formatTsv ([System.IO.File]::ReadAllText((toLongPath $inputFilePath))) $firstRow $firstColumn
    if ($content -eq "") {
        return $false
    }

    [System.IO.File]::WriteAllText((toLongPath $outputFilePath), "${content}`r`n", ${utf8Bom})
    return $true
}

# インデックスのTSVのファイル名を、元のファイル名（book）と場所（sheet。encodeIndexPlace で符号化したもの）に分ける正規表現
# （大文字・小文字を区別しない。検索処理の C# でも使う）。
#   ・場所は _ を符号化してある（toIndexFileName）ため、最後の _ で分ける（ファイル名に .xlsx_ 等を含むブックでも正しく分かれる）
#   ・以前の版のTSV（場所を全角に置き換え、_ はそのまま）で最後の _ の前が拡張子にならないものは、最初の「拡張子_」で分ける
${indexFileNamePattern} = "^(?:(?<book>.*\.(?:xls|doc|ppt)[a-z]?)_(?<sheet>[^_]*)|(?<book>.*?\.(?:xls|doc|ppt)[a-z]?)_(?<sheet>.*))\.tsv$"

function splitIndexFileName {
    # 以前の形式（フラット）の "ファイル名.拡張子_場所.tsv" を、ファイル名と場所に分解する（場所は decodeIndexPlace で元に戻す）。
    # 今の形式は「ファイル名のフォルダ＋場所.tsv」のため splitIndexTsvPath を使う
    #   例: "ブック名.xlsx_シート名.tsv" / "文書.docx_ページ001.tsv" / "資料.pptx_スライド003%5Fノート.tsv"
    param (
        [string]$fileName
    )

    if ($fileName -match ${indexFileNamePattern}) {
        return @{ book = $Matches.book; sheet = (decodeIndexPlace $Matches.sheet) }
    }

    return @{ book = $fileName; sheet = "" }
}

# インデックスの「元のファイル名のフォルダ」と分かる名前（Officeファイルの拡張子で終わる）
${indexBookDirPattern} = "\.(?:xls|doc|ppt)[a-z]?$"

function splitIndexTsvPath {
    # インデックスフォルダからのTSVの相対パスを @{ Book（元のファイル名）; Place（場所）; RelDir（元のファイルのあるフォルダ） } に分解する。
    # 今の形式（<相対フォルダ>\<ファイル名.xlsx>\<場所>.tsv）と、以前の形式（<相対フォルダ>\<ファイル名.xlsx>_<場所>.tsv）の両方を扱う。
    # 場所には _ を符号化して入れる（encodeIndexPlace）ため、ファイル名に _ があれば以前の形式と分かる
    param (
        [string]$relPath
    )

    $fileName = [System.IO.Path]::GetFileName($relPath)
    $dir = [System.IO.Path]::GetDirectoryName($relPath)
    if ($fileName.IndexOf("_") -lt 0 -and $dir -and ([System.IO.Path]::GetFileName($dir) -match ${indexBookDirPattern})) {
        return @{
            Book   = [System.IO.Path]::GetFileName($dir)
            Place  = (decodeIndexPlace ([System.IO.Path]::GetFileNameWithoutExtension($fileName)))
            RelDir = [System.IO.Path]::GetDirectoryName($dir)
        }
    }

    $name = splitIndexFileName $fileName
    return @{ Book = $name.book; Place = $name.sheet; RelDir = $dir }
}

function toResultLine {
    # 検索結果1件を "ファイル名<TAB>場所<TAB>行番号<TAB>該当行" に整形する。
    # Excelに貼り付けたとき、該当行の各セルが元の列の順（4列目 = A列）に並ぶようにする
    param (
        [string]$book,
        [string]$location,
        [int]$lineNumber,
        [string]$line
    )

    if ($book -match "\.xls[a-z]?$") {
        # セル内改行を戻す（改行を含むセルは " で囲まれているため、貼り付けても1セルのまま）
        $line = $line.Replace(${cellNewLine}, "`n")
    } else {
        # Word・PowerPointの行（段落、または表の1行をタブ区切りにしたもの）は " で囲まれていない。
        # Excelは " で始まるセルを囲みの " と解釈して後続がずれるため、そのセルだけ " で囲む
        $line = @($line.Split("`t") | ForEach-Object {
            if ($_.StartsWith('"')) { '"' + $_.Replace('"', '""') + '"' } else { $_ }
        }) -join "`t"
    }
    # 場所（シート名）にタブ・改行が入っていると、列・行が分かれてしまうためスペースにする
    # （Excelのシート名はタブ・改行を含められる）
    $place = [regex]::Replace($location, '[\x00-\x1F]', " ")
    return "${book}`t${place}`t${lineNumber}`t${line}"
}

function countTsvFields {
    # TSVの1行をExcelに貼り付けたときのセル数を返す。
    # Excelと同じく、" で始まるセルは閉じる " までを1セルとする（中のタブ・改行は区切りとしない）
    param (
        [string]$line
    )

    # 先頭にタブを足し、どのセルも「タブ＋中身」で数える（^ を使うと、先頭の空のセルの後のタブを読み飛ばして1セル少なくなる）
    return [regex]::Matches("`t${line}", '\t(?:"(?:[^"]|"")*"[^\t]*|[^\t]*)').Count
}

function toColumnName {
    # 列番号を列名に変換する（1 → A、27 → AA）
    param (
        [int]$number
    )

    $name = ""
    while ($number -gt 0) {
        $number--
        $name = [string][char](65 + ($number % 26)) + $name
        $number = [math]::Floor($number / 26)
    }
    return $name
}

function toResultHeader {
    # 検索結果の見出し行 "ファイル名<TAB>場所<TAB>行<TAB>A<TAB>B…" を返す（列名は columnCount 列分）
    param (
        [int]$columnCount
    )

    $names = New-Object System.Collections.Generic.List[string]
    $names.AddRange([string[]]@("ファイル名", "場所", "行"))
    for ($i = 1; $i -le $columnCount; $i++) {
        $names.Add((toColumnName $i))
    }
    return ($names -join "`t")
}

# ----------------------------------------------------------------------------
# 検索・Officeプロセス・画面（config_gui.ps1）で共有する処理
# ----------------------------------------------------------------------------

# 画面の履歴（画面が作成する）
${historyFile}      = "${workDir}\検索履歴.txt"
${historyMax}       = 20

# 強制終了の対象: プロセス名 → 表示名
${officeProcessNames} = [ordered]@{ EXCEL = "Excel"; WINWORD = "Word"; POWERPNT = "PowerPoint" }

function isValidRegex {
    # 正規表現として正しいか
    param (
        [string]$pattern
    )

    try {
        [void][regex]::new($pattern)
        return $true
    } catch {
        return $false
    }
}

# 1行の照合にかけてよい時間。正規表現によっては終わらなくなるため、超えたら検索を止める
${regexTimeout} = [timespan]::FromSeconds(5)

function newSearchRegex {
    # 検索条件から、検索と一致箇所の強調で使う正規表現を作る（画面の検索と結果の表示で共通）。
    #   simpleMatch  : $true ならワードを文字どおりに探す。$false なら正規表現として探し、正規表現として不正なら文字どおりに探す
    #   caseSensitive: 英字の大文字・小文字を区別する（既定は区別しない）
    # @{ Regex; SimpleMatch（実際に文字どおり探すか） } を返す
    param (
        [string]$word,
        [bool]$simpleMatch = $true,
        [bool]$caseSensitive = $false
    )

    if (!$simpleMatch -and !(isValidRegex $word)) {
        $simpleMatch = $true
    }
    $pattern = if ($simpleMatch) { [regex]::Escape($word) } else { $word }
    # CultureInvariant: 大文字・小文字を区別しないときの照合が、区別する場合と同程度に速くなる（日本語の照合結果は変わらない）
    $options = [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
    if (!$caseSensitive) {
        $options = $options -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    }
    return @{
        Regex       = New-Object System.Text.RegularExpressions.Regex($pattern, $options, ${regexTimeout})
        SimpleMatch = $simpleMatch
    }
}

function newFileFilter {
    # 対象ファイルの指定（例: "*.xlsx;見積*;!*old*"）を、元のファイル名に対する正規表現 @{ Include; Exclude } にする（無い側は $null）。
    #   ; で区切る（全角の ； も可）。! で始まるものは除外。* は任意の文字列、? は任意の1文字。大文字・小文字を区別しない
    #   * も ? も無いものは部分一致（「見積」は「*見積*」）
    param (
        [string]$filter
    )

    $include = New-Object System.Collections.Generic.List[string]
    $exclude = New-Object System.Collections.Generic.List[string]
    foreach ($item in ([string]$filter).Split([char[]]";；")) {
        $item = $item.Trim()
        $list = $include
        if ($item.StartsWith("!") -or $item.StartsWith("！")) {
            $list = $exclude
            $item = $item.Substring(1).Trim()
        }
        if ($item -eq "") {
            continue
        }
        if ($item.IndexOfAny([char[]]"*?") -lt 0) {
            $item = "*${item}*"
        }
        $list.Add("^" + [regex]::Escape($item).Replace("\*", ".*").Replace("\?", ".") + "$")
    }

    $options = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
    $result = @{ Include = $null; Exclude = $null }
    if ($include.Count -gt 0) {
        $result.Include = New-Object System.Text.RegularExpressions.Regex(($include -join "|"), $options)
    }
    if ($exclude.Count -gt 0) {
        $result.Exclude = New-Object System.Text.RegularExpressions.Regex(($exclude -join "|"), $options)
    }
    return $result
}

# 検索処理の本体（TSVを読んで照合し、結果を作る）。Select-String と PowerShell で1件ずつ結果を作ると、
# 件数が多いとき（数万件）に数十秒かかるため C# にする。初めて検索するときにコンパイルする（initTsvSearcher）
${tsvSearcherSource} = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace WinGrep
{
    // 検索結果1件（TSVの1行）
    public class TsvSearchHit
    {
        public string Root { get; set; }
        public string RelPath { get; set; }
        public string RelDir { get; set; }
        public string FileName { get; set; }
        public string Book { get; set; }
        public string Location { get; set; }
        public int LineNumber { get; set; }
        public string Line { get; set; }
    }

    // 検索対象のTSV1件
    public class TsvSearchFile
    {
        public string Path;      // 読み込むパス（\\?\ 付き）
        public string Root;
        public string RelPath;
        public string RelDir;
        public string FileName;
        public string Book;
        public string Location;
    }

    public static class TsvSearcher
    {
        // TSVのファイル名の場所を元に戻す（decodeIndexPlace と同じ。encodeIndexPlace が作る "%XX" だけを戻す）
        static readonly Regex placeEscape = new Regex("%(?:[01][0-9A-F]|2[25AF]|3[ACEF]|5[CF]|7C)");

        public static string DecodePlace(string place)
        {
            if (place.IndexOf('%') < 0) return place;
            return placeEscape.Replace(place, m => ((char)Convert.ToInt32(m.Value.Substring(1), 16)).ToString());
        }

        // 検索対象のTSVを作る。元のファイル名（Book）が include に一致しない・exclude に一致するものは除く（null は条件なし）。
        // 今の形式は <相対フォルダ>\<ファイル名.xlsx>\<場所>.tsv、以前の形式は <相対フォルダ>\<ファイル名.xlsx>_<場所>.tsv
        // （場所は _ を符号化してあるため、ファイル名に _ があれば以前の形式。splitIndexTsvPath と同じ規則）
        public static List<TsvSearchFile> NewFiles(string[] paths, string[] roots, string[] relPaths, Regex fileNamePattern, Regex bookDirPattern, Regex include, Regex exclude)
        {
            List<TsvSearchFile> files = new List<TsvSearchFile>();
            for (int i = 0; i < paths.Length; i++)
            {
                TsvSearchFile f = new TsvSearchFile();
                f.Path = paths[i];
                f.Root = roots[i];
                f.RelPath = relPaths[i];
                f.RelDir = System.IO.Path.GetDirectoryName(relPaths[i]) ?? "";
                f.FileName = System.IO.Path.GetFileName(relPaths[i]);
                string bookDir = System.IO.Path.GetFileName(f.RelDir);
                if (f.FileName.IndexOf('_') < 0 && bookDir.Length > 0 && bookDirPattern.IsMatch(bookDir))
                {
                    f.Book = bookDir;
                    f.Location = DecodePlace(System.IO.Path.GetFileNameWithoutExtension(f.FileName));
                    f.RelDir = System.IO.Path.GetDirectoryName(f.RelDir) ?? "";
                }
                else
                {
                    Match m = fileNamePattern.Match(f.FileName);
                    f.Book = m.Success ? m.Groups["book"].Value : f.FileName;
                    f.Location = m.Success ? DecodePlace(m.Groups["sheet"].Value) : "";
                }
                if (include != null && !include.IsMatch(f.Book)) continue;
                if (exclude != null && exclude.IsMatch(f.Book)) continue;
                files.Add(f);
            }
            return files;
        }

        // files の start から count 件を読み、regex に一致する行を返す（1行に複数一致しても1件）。
        // max 以上なら、max を超えた（max+1 件目が見つかった）時点で止める（負なら上限なし）。
        // 読めないTSV（変換中に削除された等）は飛ばす
        public static List<TsvSearchHit> Search(List<TsvSearchFile> files, int start, int count, Regex regex, int max)
        {
            List<TsvSearchHit> hits = new List<TsvSearchHit>();
            int end = Math.Min(files.Count, start + count);
            for (int i = start; i < end; i++)
            {
                TsvSearchFile f = files[i];
                try
                {
                    using (FileStream fs = new FileStream(f.Path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                    using (StreamReader reader = new StreamReader(fs, Encoding.UTF8, true))
                    {
                        string line;
                        int number = 0;
                        while ((line = reader.ReadLine()) != null)
                        {
                            number++;
                            if (!regex.IsMatch(line)) continue;
                            TsvSearchHit h = new TsvSearchHit();
                            h.Root = f.Root;
                            h.RelPath = f.RelPath;
                            h.RelDir = f.RelDir;
                            h.FileName = f.FileName;
                            h.Book = f.Book;
                            h.Location = f.Location;
                            h.LineNumber = number;
                            h.Line = line;
                            hits.Add(h);
                            if (max >= 0 && hits.Count > max) return hits;
                        }
                    }
                }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
            return hits;
        }
    }
}
'@

function initTsvSearcher {
    # 検索処理の C# をコンパイルする（同じプロセスでは1回だけ。画面の検索スレッドとも共有される）
    if (!("WinGrep.TsvSearcher" -as [type])) {
        Add-Type -TypeDefinition ${tsvSearcherSource}
    }
}

function getIndexTsvFiles {
    # 検索対象インデックスのTSVを列挙し、@{ Folders; Files } を返す。
    #   folders: 検索対象インデックスのフォルダ（文字列。フォルダ以下すべて）、または
    #            @{ Root（インデックスのフォルダ）; RelPath（その中のフォルダ。空は Root 自身）; Recurse（$false は直下のファイルだけ） }
    #            （画面の検索対象ツリーで一部のフォルダだけを選んだとき。結果の相対パスは Root から求める）
    #   Folders: フォルダごとの @{ Path（指定どおり。RelPath があれば Root\RelPath）; Root（フルパス）; Exists; Count }
    #   Files  : TSVのフルパス → @{ Root; RelPath（インデックスフォルダからの相対パス） }。パス順。入れ子のフォルダでも重複しない
    param (
        [object[]]$folders = @(getIndexFolders)
    )

    $folderInfo = New-Object System.Collections.Generic.List[object]
    $files = @{}
    foreach ($target in $folders) {
        if ($target -is [string]) {
            $target = @{ Root = $target; RelPath = ""; Recurse = $true }
        }
        $relPath = ([string]$target.RelPath).Trim("\")
        $dir = if ($relPath) { "$(([string]$target.Root).TrimEnd('\'))\${relPath}" } else { [string]$target.Root }
        $exists = Test-Path -LiteralPath $target.Root -PathType Container
        if ($exists) {
            $root = (Resolve-Path -LiteralPath $target.Root).ProviderPath.TrimEnd("\")
            $fullDir = if ($relPath) { "${root}\${relPath}" } else { $root }
            $exists = [System.IO.Directory]::Exists((toLongPath $fullDir))
        }
        if (!$exists) {
            $folderInfo.Add(@{ Path = $dir; Root = ""; Exists = $false; Count = 0 })
            continue
        }

        # 長いパス（フォルダが約248文字超）の中も列挙できるよう \\?\ 付きで列挙し、キーは \\?\ の無い通常のパスにする
        $found = @(Get-ChildItem -LiteralPath (toLongPath $fullDir) -Filter "*.tsv" -File -Recurse:([bool]$target.Recurse))
        if (!$target.Recurse) {
            # 「フォルダ直下のファイル」には、元のファイル名のフォルダ（<ファイル名.xlsx>\<場所>.tsv）の中のTSVも含める
            $found += @(Get-ChildItem -LiteralPath (toLongPath $fullDir) -Directory |
                Where-Object { $_.Name -match ${indexBookDirPattern} } |
                ForEach-Object { Get-ChildItem -LiteralPath $_.FullName -Filter "*.tsv" -File })
        }
        $folderInfo.Add(@{ Path = $dir; Root = $root; Exists = $true; Count = $found.Count })
        foreach ($file in $found) {
            $fullName = fromLongPath $file.FullName
            $relative = getPathUnderFolder $fullName $root
            if ($null -eq $relative) {
                $relative = $file.Name  # 通常は起こらない（$root の下を列挙している）
            }
            $files[$fullName] = @{ Root = $root; RelPath = $relative }
        }
    }

    $sorted = [ordered]@{}
    foreach ($path in @($files.Keys | Sort-Object)) {
        $sorted[$path] = $files[$path]
    }
    return @{ Folders = $folderInfo.ToArray(); Files = $sorted }
}

function testIndexExists {
    # 検索対象インデックスにTSVが1件でもあるか（最初の1件が見つかった時点で打ち切る）
    param (
        [string[]]$folders = @(getIndexFolders)
    )

    foreach ($dir in $folders) {
        if (!(Test-Path -LiteralPath $dir -PathType Container)) {
            continue
        }
        try {
            $enumerator = [System.IO.Directory]::EnumerateFiles((toLongPath (Resolve-Path -LiteralPath $dir).ProviderPath), "*.tsv", [System.IO.SearchOption]::AllDirectories).GetEnumerator()
            if ($enumerator.MoveNext()) {
                return $true
            }
        } catch {
            # アクセスできないフォルダがある場合は、件数を数える方で判定する
            if (@(Get-ChildItem -LiteralPath $dir -Filter "*.tsv" -File -Recurse -ErrorAction SilentlyContinue).Count -gt 0) {
                return $true
            }
        }
    }
    return $false
}

function getIndexSummary {
    # 検索対象インデックスのTSVの件数と最新の更新日時を返す: @{ Count; LastWrite（無ければ $null）; Missing（存在しないフォルダ） }
    param (
        [string[]]$folders = @(getIndexFolders)
    )

    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $summary = @{ Count = 0; LastWrite = $null; Missing = @() }
    foreach ($dir in $folders) {
        if (!(Test-Path -LiteralPath $dir -PathType Container)) {
            $summary.Missing += $dir
            continue
        }
        foreach ($file in @(Get-ChildItem -LiteralPath (toLongPath (Resolve-Path -LiteralPath $dir).ProviderPath) -Filter "*.tsv" -File -Recurse -ErrorAction SilentlyContinue)) {
            if (!$seen.Add($file.FullName)) {
                continue
            }
            $summary.Count++
            if ($null -eq $summary.LastWrite -or $file.LastWriteTime -gt $summary.LastWrite) {
                $summary.LastWrite = $file.LastWriteTime
            }
        }
    }
    return $summary
}

function searchIndex {
    # インデックスのTSVをワードで検索し、ヒットした行を返す（画面の検索処理）。
    #   tsvFiles     : getIndexTsvFiles の Files
    #   simpleMatch  : $true なら文字どおりに検索する。$false なら正規表現として検索し、正規表現として不正なら文字どおりに検索する
    #   limit        : 件数の上限（0 は上限なし）。超えたら打ち切る
    #   onProgress   : chunkSize 件のTSVを検索するたびに呼ぶ { param($done, $total, $newHits) }
    #   shouldStop   : $true を返すと中止する
    #   caseSensitive: 英字の大文字・小文字を区別する（newSearchRegex）
    #   fileFilter   : 対象ファイル（newFileFilter）。元のファイル名が一致しないTSVは検索しない
    # @{ Hits; SimpleMatch（実際に文字どおり検索したか）; Total（対象ファイルで絞った後のTSVの数）; Truncated; Cancelled } を返す。
    # Hits の各要素は WinGrep.TsvSearchHit（Root; RelPath; RelDir; FileName; Book; Location; LineNumber; Line）
    param (
        [string]$word,
        $tsvFiles,
        [bool]$simpleMatch = $false,
        [int]$limit = 0,
        [int]$chunkSize = 100,
        [scriptblock]$onProgress = $null,
        [scriptblock]$shouldStop = $null,
        [bool]$caseSensitive = $false,
        [string]$fileFilter = ""
    )

    initTsvSearcher
    $search = newSearchRegex $word $simpleMatch $caseSensitive
    $filter = newFileFilter $fileFilter

    $count = $tsvFiles.Count
    $paths = New-Object string[] $count
    $roots = New-Object string[] $count
    $relPaths = New-Object string[] $count
    $i = 0
    foreach ($path in $tsvFiles.Keys) {
        $info = $tsvFiles[$path]
        # 260文字を超えるパスのTSVも読めるよう \\?\ 付きで読む
        $paths[$i] = toLongPath $path
        $roots[$i] = $info.Root
        $relPaths[$i] = $info.RelPath
        $i++
    }
    $fileNameRegex = New-Object System.Text.RegularExpressions.Regex(${indexFileNamePattern}, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $bookDirRegex = New-Object System.Text.RegularExpressions.Regex(${indexBookDirPattern}, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $files = [WinGrep.TsvSearcher]::NewFiles($paths, $roots, $relPaths, $fileNameRegex, $bookDirRegex, $filter.Include, $filter.Exclude)

    $hits = New-Object System.Collections.Generic.List[WinGrep.TsvSearchHit]
    $result = @{ Hits = $hits; SimpleMatch = $search.SimpleMatch; Total = $files.Count; Truncated = $false; Cancelled = $false }

    for ($i = 0; $i -lt $files.Count; $i += $chunkSize) {
        if ($shouldStop -and (& $shouldStop)) {
            $result.Cancelled = $true
            break
        }

        # 上限があれば、残りの件数を超えた時点で止める（残りちょうどで終われば打ち切りにしない）
        $max = if ($limit -gt 0) { $limit - $hits.Count } else { -1 }
        try {
            $newHits = [WinGrep.TsvSearcher]::Search($files, $i, $chunkSize, $search.Regex, $max)
        } catch [System.Text.RegularExpressions.RegexMatchTimeoutException] {
            throw "正規表現の照合に時間がかかりすぎるため、検索を中止しました。正規表現を見直してください。"
        } catch {
            if ($_.Exception.InnerException -is [System.Text.RegularExpressions.RegexMatchTimeoutException]) {
                throw "正規表現の照合に時間がかかりすぎるため、検索を中止しました。正規表現を見直してください。"
            }
            throw
        }
        if ($max -ge 0 -and $newHits.Count -gt $max) {
            $newHits.RemoveRange($max, $newHits.Count - $max)
            $result.Truncated = $true
        }
        $hits.AddRange($newHits)

        if ($onProgress) {
            & $onProgress ([math]::Min($i + $chunkSize, $files.Count)) $files.Count $newHits
        }
        if ($result.Truncated) {
            break
        }
    }
    return $result
}

function toSearchResultLines {
    # 検索結果を検索結果ファイルの形式にし、@{ Header（見出し行）; Lines } を返す（画面のファイル出力・コピーで共通）。
    # Excelに貼り付けたときに元のセル位置が分かるよう、行番号（TSVの行番号 = Excelの行番号）を付け、
    # 見出しには該当行の最大セル数分の列名（A, B, C…）を付ける
    param (
        [object[]]$hits
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $columnCount = 0
    foreach ($hit in $hits) {
        # ファイル名の前に、変換対象フォルダからの相対フォルダを付ける
        $line = toResultLine $hit.Book $hit.Location $hit.LineNumber $hit.Line
        if ($hit.RelDir) {
            $line = "$($hit.RelDir)\${line}"
        }
        $lines.Add($line)
        $columnCount = [math]::Max($columnCount, (countTsvFields $line) - 3)  # ファイル名・場所・行番号の3列を除く
    }
    return @{ Header = (toResultHeader $columnCount); Lines = $lines }
}

function writeSearchResult {
    # 1ワード分の検索結果を検索結果ファイルに書き出す
    param (
        [System.IO.TextWriter]$writer,
        [string]$word,
        [object[]]$hits
    )

    $result = toSearchResultLines $hits
    $writer.WriteLine("【検索文字列　${word}】 $($result.Lines.Count) 件")
    $writer.WriteLine($result.Header)
    foreach ($line in $result.Lines) {
        $writer.WriteLine($line)
    }
    $writer.WriteLine("")
}

function readTsvContext {
    # インデックスのTSVの lineNumber 行目と、その前後 before 行・after 行を @{ LineNumber; Line } の配列で返す（画面の選択行のプレビュー）。
    # 行の数え方は Select-String と同じ（Excel のTSVでは行番号 = シートの行番号）。必要な行まで読んだら止める。
    # ファイルが無い・読めない場合は空
    param (
        [string]$path,
        [int]$lineNumber,
        [int]$before = 3,
        [int]$after = 3
    )

    $rows = New-Object System.Collections.Generic.List[object]
    $first = [math]::Max(1, $lineNumber - $before)
    $last = $lineNumber + $after
    $reader = $null
    try {
        $reader = New-Object System.IO.StreamReader((toLongPath $path), [System.Text.Encoding]::UTF8)
        $number = 0
        while ($number -lt $last) {
            $line = $reader.ReadLine()
            if ($null -eq $line) {
                break
            }
            $number++
            if ($number -ge $first) {
                $rows.Add([pscustomobject]@{ LineNumber = $number; Line = $line })
            }
        }
    } catch {
        $rows.Clear()
    } finally {
        if ($reader) {
            $reader.Dispose()
        }
    }
    # 呼び出し側で @() にして使う（, を付けて返すと @() で1要素に包まれる）
    return $rows.ToArray()
}

function splitTsvCells {
    # TSVの1行をセルに分ける。" で始まるセルは閉じる " までを1セルとし、囲みの " を外して "" を " に戻す（countTsvFields と同じ区切り方）
    param (
        [string]$line
    )

    $cells = New-Object System.Collections.Generic.List[string]
    foreach ($match in [regex]::Matches("`t${line}", '\t(?:"(?:[^"]|"")*"[^\t]*|[^\t]*)')) {
        $cell = $match.Value.Substring(1)
        if ($cell -match '^"(?<inner>(?:[^"]|"")*)"(?<rest>.*)$') {
            $cell = $Matches.inner.Replace('""', '"') + $Matches.rest
        }
        $cells.Add($cell)
    }
    return , $cells.ToArray()
}

function writeSourceFolderFile {
    # インデックスのフォルダに、インデックス名と変換対象フォルダの対応（元のフォルダ.txt）を書き出す。
    # 1行目は説明、2行目以降は "インデックス名<TAB>変換対象フォルダ"
    param (
        [object[]]$folders,  # assignIndexNames の結果（@{ Path; Name }）
        [string]$dir = ${indexDir}
    )

    $lines = @("# 検索結果から元のファイルを開くときに使う、インデックス名と変換対象フォルダの対応です（変換のたびに作り直します）")
    $lines += @($folders | Where-Object { $_ -and $_.Name } | ForEach-Object { "$($_.Name)`t$($_.Path)" })
    writeListFile (Join-Path $dir ${sourceFolderFileName}) $lines
}

function readSourceFolderFile {
    # インデックスのフォルダの 元のフォルダ.txt を読み、インデックス名 → 変換対象フォルダ を返す（大文字・小文字を区別しない）。
    # ファイルが無ければ空
    param (
        [string]$dir
    )

    $map = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($line in @(readListFile (Join-Path $dir ${sourceFolderFileName}))) {
        $fields = $line.Split("`t")
        if ($fields.Count -eq 2 -and $fields[0] -ne "" -and $fields[1] -ne "") {
            $map[$fields[0]] = $fields[1]
        }
    }
    return , $map
}

function getSourceFolderMap {
    # インデックスのフォルダ（dir）の インデックス名 → 元のフォルダ（今そのフォルダが置かれている場所）を返す。
    # 次の順に読み、後のもので上書きする（後のものが優先）:
    #   1. dir の 元のフォルダ.txt … インデックスを作ったときの場所。インデックスをコピーしても付いてくる
    #   2. 既定のインデックス（work\index）なら変換一覧の記録
    #   3. 設定のインデックス名に対する場所（targetFolders・indexSources）… 利用者が指定した「今の場所」のため最も優先する
    param (
        [string]$dir,
        [string]$statusPath = ${statusFile},
        [string]$settingsPath = ${settingsFile}
    )

    $map = readSourceFolderFile $dir
    if ((Test-Path -LiteralPath ${indexDir} -PathType Container) -and
        (testSameFolder $dir (Resolve-Path -LiteralPath ${indexDir}).ProviderPath)) {
        foreach ($entry in (getIndexNameMap $statusPath).GetEnumerator()) {
            $map[$entry.Key] = $entry.Value
        }
    }
    foreach ($folder in @(getTargetFolders $settingsPath | Where-Object { $_.Name })) {
        $map[$folder.Name] = $folder.Path
    }
    foreach ($source in @(readIndexSources $settingsPath)) {
        $map[$source.Name] = $source.Path
    }
    return , $map
}

function joinSourcePath {
    # フォルダ・相対フォルダ（空でも可）・ファイル名をつなぐ（ドライブ直下 "D:\" でも \ が重ならないようにする）
    param (
        [string]$folder,
        [string]$rest,
        [string]$name = ""
    )

    $path = $folder.TrimEnd("\")
    foreach ($part in @($rest, $name)) {
        if ($part) {
            $path += "\$part"
        }
    }
    return $path
}

function getSourceLocation {
    # 検索結果の元のファイルの場所を @{ Name（インデックス名）; Folder（元のフォルダ）; Rest（その下の相対フォルダ）; Known } で返す。
    # インデックスのフォルダ（Root）の下は "インデックス名\相対フォルダ" のため、インデックス名から元のフォルダを引く（getSourceFolderMap）。
    # 検索対象にインデックス名のフォルダ（…\index\<インデックス名>）を直接指定した場合は、親フォルダの記録を使う。
    # 元のフォルダが分からなければ Known = $false・Folder = "" とする（Name は返すため、フォルダを選んでもらえば設定に記録できる）
    #   maps: フォルダ → getSourceFolderMap の結果 のキャッシュ（読んだ結果を追加する）
    param (
        $hit,
        [hashtable]$maps = @{}
    )

    $root = ([string]$hit.Root).TrimEnd("\")
    $relDir = [string]$hit.RelDir
    $candidates = New-Object System.Collections.Generic.List[object]
    if ($relDir) {
        $parts = splitIndexRelPath $relDir
        $candidates.Add(@{ Dir = $root; Name = $parts.Name; Rest = $parts.Rest })
    }
    $parent = Split-Path $root -Parent
    if ($parent) {
        $candidates.Add(@{ Dir = $parent; Name = (Split-Path $root -Leaf); Rest = $relDir })
    }

    foreach ($candidate in $candidates) {
        if (!$maps.ContainsKey($candidate.Dir)) {
            $maps[$candidate.Dir] = getSourceFolderMap $candidate.Dir
        }
        $map = $maps[$candidate.Dir]
        if ($map.ContainsKey($candidate.Name)) {
            return @{ Name = $candidate.Name; Folder = $map[$candidate.Name]; Rest = $candidate.Rest; Known = $true }
        }
    }

    # 分からない場合も、インデックス名と、その下の相対フォルダは分かる（検索対象にインデックス名のフォルダを直接指定した場合は Rest がすべて）
    if ($candidates.Count -gt 0) {
        return @{ Name = $candidates[0].Name; Folder = ""; Rest = $candidates[0].Rest; Known = $false }
    }
    return @{ Name = (Split-Path $root -Leaf); Folder = ""; Rest = ""; Known = $false }
}

function findMovedSource {
    # 元のファイルが見つからないとき、選んでもらったフォルダ（picked）の中から探す。
    # picked は元のフォルダ（インデックスのルート）に当たるフォルダでも、その下のどのフォルダ（ファイルのあるフォルダなど）に当たるフォルダでもよい。
    # 上の階層に当たるとみなす方から順に試し、見つかれば @{ Path（見つかったファイル）; Root（元のフォルダに当たるフォルダ。遡れなければ ""） }、
    # 見つからなければ $null を返す。Root を設定に記録すれば（setIndexSourceFolder）、同じインデックスのほかのファイルも開ける
    param (
        [string]$picked,
        [string]$rest,
        [string]$book
    )

    $picked = normalizeFolderPath $picked
    $segments = [string[]]@($rest.Split([char[]]@("\"), [System.StringSplitOptions]::RemoveEmptyEntries))
    for ($skip = 0; $skip -le $segments.Count; $skip++) {
        # 先頭の skip 個のフォルダは picked より上、残りは picked の下にあるとみなす
        $below = [string]::Join("\", $segments, $skip, $segments.Count - $skip)
        $candidate = joinSourcePath $picked $below $book
        if (Test-Path -LiteralPath (toLongPath $candidate) -PathType Leaf) {
            # 選んだフォルダは「元のフォルダ＋先頭 skip 個のフォルダ」に当たるため、skip 個上が元のフォルダに当たる
            $root = $picked
            for ($i = 0; $i -lt $skip -and $root; $i++) {
                $root = Split-Path $root -Parent
            }
            return @{ Path = $candidate; Root = [string]$root }
        }
    }
    return $null
}

function resolveSourcePath {
    # 検索結果の元のファイルのパスを返す（ファイルがあるかは確かめない）。元のフォルダが分からなければ $null
    #   maps: getSourceLocation のキャッシュ
    param (
        $hit,
        [hashtable]$maps = @{}
    )

    $location = getSourceLocation $hit $maps
    if (!$location.Known) {
        return $null
    }
    return (joinSourcePath $location.Folder $location.Rest $hit.Book)
}

function readSearchHistory {
    # 検索履歴（新しい順）
    param (
        [string]$path = ${historyFile}
    )

    return @(readListFile $path)
}

function addSearchHistory {
    # 検索履歴の先頭にワードを加えて保存し、履歴を返す。同じワードは先頭に移す
    param (
        [string]$word,
        [string]$path = ${historyFile},
        [int]$max = ${historyMax}
    )

    $word = $word.Trim()
    if ($word -eq "") {
        return @(readSearchHistory $path)
    }
    $items = @(@($word) + @(readSearchHistory $path | Where-Object { $_ -cne $word }) | Select-Object -First $max)
    writeListFile $path $items
    return $items
}

${searchOptionKeys} = [ordered]@{ UseRegex = "useRegex"; CaseSensitive = "caseSensitive"; FileFilter = "fileFilter" }

function readSearchOption {
    # 画面の検索オプションを @{ UseRegex; CaseSensitive; FileFilter } で返す。
    # 設定が無ければ、文字どおり・大文字と小文字を区別しない・対象ファイルはすべて
    param (
        [string]$path = ${settingsFile}
    )

    $settings = readSettings $path
    $option = @{}
    foreach ($name in ${searchOptionKeys}.Keys) {
        $option[$name] = $settings[${searchOptionKeys}[$name]]
    }
    return $option
}

function writeSearchOption {
    # 画面の検索オプションを保存する。option（readSearchOption と同じ形）にある項目だけを変える
    param (
        [hashtable]$option,
        [string]$path = ${settingsFile}
    )

    $settings = readSettings $path
    foreach ($name in ${searchOptionKeys}.Keys) {
        if ($option.ContainsKey($name)) {
            $settings[${searchOptionKeys}[$name]] = $option[$name]
        }
    }
    writeSettings $settings $path
}

function readOpenMode {
    # 元のファイルの開き方（${openModes} のいずれか）を返す。設定が無い・知らない値なら「通常」
    param (
        [string]$path = ${settingsFile}
    )

    $mode = (readSettings $path).openMode
    if (${openModes} -contains $mode) {
        return $mode
    }
    return ${openModeNormal}
}

function writeOpenMode {
    param (
        [string]$mode,
        [string]$path = ${settingsFile}
    )

    updateSettings "openMode" $mode $path
}

function getConversionState {
    # 変換一覧から変換の状態を返す:
    #   @{ Exists; Folders（変換対象フォルダ @{ Path; Name } の配列）; Total; Pending（未変換）; Failed; Done; ConvertedSince（since 以降に変換した件数）; Updated（変換一覧の更新日時）;
    #      FailedRows（失敗したファイルの行。変換日時の新しい順） }
    param (
        [datetime]$since = [datetime]::MaxValue,
        [string]$path = ${statusFile}
    )

    $state = @{ Exists = $false; Folders = @(); Total = 0; Pending = 0; Failed = 0; Done = 0; ConvertedSince = 0; Updated = $null; FailedRows = @() }
    if (!(Test-Path -LiteralPath $path)) {
        return $state
    }

    $state.Exists = $true
    $state.Updated = (Get-Item -LiteralPath $path).LastWriteTime
    $status = readStatusFile $path
    $state.Folders = $status.Folders.ToArray()
    $state.Total = $status.Rows.Count
    $sinceSecond = if ($since -eq [datetime]::MaxValue) { $since } else { $since.AddTicks(-($since.Ticks % [timespan]::TicksPerSecond)) }
    $failedRows = New-Object System.Collections.Generic.List[object]
    foreach ($row in $status.Rows.Values) {
        if ($row.状態 -eq ${stateNew}) {
            $state.Pending++
        } elseif ($row.状態 -eq ${stateFailed}) {
            $state.Failed++
            $failedRows.Add($row)
        } elseif ($row.状態 -eq ${stateDone}) {
            $state.Done++
        }

        $converted = [datetime]::MinValue
        if ($row.変換日時 -and [datetime]::TryParseExact($row.変換日時, "yyyy/MM/dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$converted) -and $converted -ge $sinceSecond) {
            $state.ConvertedSince++
        }
    }
    # 変換日時は "yyyy/MM/dd HH:mm:ss" のため、文字列の順で新しい順に並ぶ
    $state.FailedRows = @($failedRows | Sort-Object -Property @{ Expression = { [string]$_.変換日時 }; Descending = $true }, @{ Expression = { $_.相対パス } })
    return $state
}

function getOfficeProcesses {
    # 実行中の Excel・Word・PowerPoint を返す。
    # ウィンドウを持たない（MainWindowHandle が 0）プロセスは、変換処理などでバックグラウンド起動されたものとする
    $result = New-Object System.Collections.Generic.List[object]
    foreach ($process in @(Get-Process -Name @(${officeProcessNames}.Keys) -ErrorAction SilentlyContinue)) {
        $startTime = $null
        try {
            $startTime = $process.StartTime
        } catch {
            # 権限の無いプロセスは起動時刻を取得できない
        }
        $result.Add([pscustomobject]@{
            Id          = $process.Id
            ProcessName = $process.ProcessName
            AppName     = ${officeProcessNames}[$process.ProcessName]
            Background  = ($process.MainWindowHandle -eq [IntPtr]::Zero)
            StartTime   = $startTime
            MemoryMB    = [math]::Round($process.WorkingSet64 / 1MB)
            Title       = $process.MainWindowTitle
        })
    }
    return $result.ToArray()
}

function stopOfficeProcesses {
    # 指定したプロセスを保存せずに終了し、@{ Id; Stopped; Message } の配列を返す
    param (
        [int[]]$ids
    )

    $results = New-Object System.Collections.Generic.List[object]
    foreach ($id in $ids) {
        try {
            Stop-Process -Id $id -Force -ErrorAction Stop
            $results.Add(@{ Id = $id; Stopped = $true; Message = "" })
        } catch {
            $results.Add(@{ Id = $id; Stopped = $false; Message = $_.Exception.Message })
        }
    }
    return $results.ToArray()
}