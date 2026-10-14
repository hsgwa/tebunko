# TSV のインデックスへの取り込みと、作業フォルダ・外したフォルダのインデックスの後始末（tebunko\indexer\index_migrate.ps1）のテスト。
# 作業フォルダ（$tmpDir）・出力用のフォルダ（$publishDir）・インデックスのフォルダ（$indexDir）は indexer.ps1 が決めるため、
# テストごとに TestDrive の下に差し替える。
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
    . "${scriptsDir}\tebunko\indexer\index_migrate.ps1"

    # 存在しないプロセスID（Windows のプロセスIDは 4 の倍数のため、奇数は使われない）
    $deadPid = 999999999
}

Describe "publishTsv" -Tag Io {
    It "作業フォルダのTSVを、そのファイルのインデックスのフォルダへ移す" {
        $tmpDir = "$TestDrive\publish\tmp"
        $publishDir = "$TestDrive\publish\出力"
        $workspace = newTestWorkspace @{ PublishDir = $publishDir }
        $bookDir = "$TestDrive\publish\index\営業\見積.xlsx"
        newTsv "$tmpDir\Sheet1.tsv" @("a`tb")
        newTsv "$tmpDir\Sheet2.tsv" @("c")
        newTsv "$bookDir\古いシート.tsv" @("old")

        publishTsv $bookDir

        @(Get-ChildItem -LiteralPath $bookDir | ForEach-Object { $_.Name } | Sort-Object) | Should -Be @("Sheet1.tsv", "Sheet2.tsv")
        @(Get-ChildItem -LiteralPath $tmpDir).Count | Should -Be 0
        Test-Path -LiteralPath "$publishDir\見積.xlsx" | Should -Be $false
    }

    It "ファイル名に [ ] があっても取り込み、TSV 以外のファイルは取り込まない" {
        $tmpDir = "$TestDrive\publish2\tmp"
        $publishDir = "$TestDrive\publish2\出力"
        $workspace = newTestWorkspace @{ PublishDir = $publishDir }
        $bookDir = "$TestDrive\publish2\index\営業\[確定]見積.xlsx"
        newTsv "$tmpDir\[表]売上.tsv" @("a")
        newTsv "$tmpDir\作業.txt" @("x")

        publishTsv $bookDir

        @(Get-ChildItem -LiteralPath $bookDir | ForEach-Object { $_.Name }) | Should -Be @("[表]売上.tsv")
        Test-Path -LiteralPath "$tmpDir\作業.txt" | Should -Be $true
    }

    It "取り込んだ TSV が 0 件なら、前回のインデックスを空にする（文字の無いファイルになった）" {
        $tmpDir = "$TestDrive\publish3\tmp"
        $publishDir = "$TestDrive\publish3\出力"
        $workspace = newTestWorkspace @{ PublishDir = $publishDir }
        $bookDir = "$TestDrive\publish3\index\営業\空.xlsx"
        [System.IO.Directory]::CreateDirectory($tmpDir) | Out-Null
        newTsv "$bookDir\Sheet1.tsv" @("old")

        publishTsv $bookDir

        Test-Path -LiteralPath $bookDir -PathType Container | Should -Be $true
        @(Get-ChildItem -LiteralPath $bookDir).Count | Should -Be 0
    }

    It "260 文字を超えるパスのインデックスにも取り込む" {
        $tmpDir = "$TestDrive\publish4\tmp"
        $publishDir = "$TestDrive\publish4\出力"
        $workspace = newTestWorkspace @{ PublishDir = $publishDir }
        $bookDir = "$TestDrive\publish4\index\" + ("深いフォルダ" * 20) + "\" + ("もっと深いフォルダ" * 15) + "\見積.xlsx"
        newTsv "$tmpDir\Sheet1.tsv" @("a")
        try {
            $bookDir.Length | Should -BeGreaterThan 260
            publishTsv $bookDir
            [System.IO.File]::Exists((toLongPath "$bookDir\Sheet1.tsv")) | Should -Be $true
        } finally {
            # TestDrive の後片付けは長いパスを消せないため、ここで消す
            removeDirectoryRetry "$TestDrive\publish4"
        }
    }
}

Describe "clearTmpDir" -Tag Io {
    It "作業フォルダ直下のファイルだけを消す" {
        $tmpDir = "$TestDrive\clear"
        newTsv "$tmpDir\a.tsv" @("a")
        newTsv "$tmpDir\sub\b.tsv" @("b")

        clearTmpDir

        Test-Path -LiteralPath "$tmpDir\a.tsv" | Should -Be $false
        Test-Path -LiteralPath "$tmpDir\sub\b.tsv" | Should -Be $true
    }

    It "<kind>のファイルも消える（作業フォルダ直下の 1 つだけ。中のフォルダは残る）" -TestCases @(
        @{ kind = "隠し"; attribute = [System.IO.FileAttributes]::Hidden }
        @{ kind = "読み取り専用"; attribute = [System.IO.FileAttributes]::ReadOnly }
    ) {
        $tmpDir = "$TestDrive\clear_attr_$kind"
        newTsv "$tmpDir\a.tsv" @("a")
        newTsv "$tmpDir\sub\b.tsv" @("b")
        [System.IO.File]::SetAttributes("$tmpDir\a.tsv", $attribute)

        clearTmpDir

        [System.IO.File]::Exists("$tmpDir\a.tsv") | Should -Be $false
        [System.IO.File]::Exists("$tmpDir\sub\b.tsv") | Should -Be $true
    }

    It "作業フォルダそのものがジャンクションなら、先のファイルを消さず、ログに出す" {
        $outside = "$TestDrive\clear_outside"
        newTsv "$outside\外.tsv" @("外")
        if (!(newTestJunction "$TestDrive\clear_link" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        $tmpDir = "$TestDrive\clear_link"
        $log = New-Object System.IO.StringWriter
        $script:indexerLog = $log

        clearTmpDir

        Test-Path -LiteralPath "$outside\外.tsv" | Should -Be $true
        $log.ToString() | Should -Match "リンク"
    }

    It "作業フォルダ直下のファイルのシンボリックリンクは、リンクだけを消し、先は残す" {
        $tmpDir = "$TestDrive\clear_symlink"
        newTsv "$TestDrive\clear_symlink_outside\外.txt" @("外")
        [System.IO.Directory]::CreateDirectory($tmpDir) | Out-Null
        if (!(newTestSymlink "$tmpDir\リンク.txt" "$TestDrive\clear_symlink_outside\外.txt" -File)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        clearTmpDir

        Test-Path -LiteralPath "$tmpDir\リンク.txt" | Should -Be $false
        Test-Path -LiteralPath "$TestDrive\clear_symlink_outside\外.txt" | Should -Be $true
    }
}

Describe "removeEmptyDir" -Tag Io {
    It "削除の権限が無いフォルダ（共有フォルダで他の利用者のものなど）でも、例外にならずそのまま残す" {
        $parent = "$TestDrive\denied"
        $dir = "$parent\child"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
        foreach ($target in @($parent, $dir)) {
            $acl = Get-Acl -LiteralPath $target
            $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($identity, "FullControl", "Deny")))
            Set-Acl -LiteralPath $target -AclObject $acl
        }
        try {
            { removeEmptyDir $dir } | Should -Not -Throw
        } finally {
            # 後片付け（TestDrive の削除）が行えるよう、権限を元に戻す
            icacls $parent /reset /t /c 2>&1 | Out-Null
        }
        Test-Path -LiteralPath $dir | Should -Be $true
    }

    It "空のフォルダは消し、ジャンクションは（先が空でも）外さずに残す" {
        [System.IO.Directory]::CreateDirectory("$TestDrive\emptydir_plain") | Out-Null
        removeEmptyDir "$TestDrive\emptydir_plain"
        Test-Path -LiteralPath "$TestDrive\emptydir_plain" | Should -Be $false

        if (!(newTestJunction "$TestDrive\emptydir_link" "$TestDrive\emptydir_target")) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        removeEmptyDir "$TestDrive\emptydir_link"
        testPathIsLink "$TestDrive\emptydir_link" | Should -Be $true
    }

    It "置き場（tmp・出力用のフォルダの親）がジャンクションでも、片付けと終了のあともジャンクションが残り、先は消えない" {
        $base = "$TestDrive\rootlink"
        [System.IO.Directory]::CreateDirectory($base) | Out-Null
        foreach ($name in @("tmp", "publish")) {
            if (!(newTestJunction "$base\$name" "$TestDrive\rootlink_target\$name")) {
                Set-ItResult -Skipped -Because $script:linkError
                return
            }
        }
        $workspace = newTestWorkspace @{ PublishDir = "$base\publish\$PID"; TmpRoot = "$base\tmp" }
        ${legacyTmpParent} = "$TestDrive\rootlink_legacy"

        $result = initTmpDir
        $result.Reason | Should -Be ""
        $result.Dir | Should -Be (getWorkspaceTmpDir $workspace)
        $tmpDir = $result.Dir
        [System.IO.Directory]::CreateDirectory($workspace.PublishDir) | Out-Null
        removeStaleTmpDirs
        removeTmpDir

        testPathIsLink "$base\tmp" | Should -Be $true
        testPathIsLink "$base\publish" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\rootlink_target\tmp" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\rootlink_target\publish" | Should -Be $true
        Test-Path -LiteralPath $result.Dir | Should -Be $false
    }
}

Describe "removeTmpDir" -Tag Io {
    It "作業フォルダと出力用のフォルダを削除する" {
        $tmpDir = "$TestDrive\remove\tmp"
        $publishDir = "$TestDrive\remove\出力"
        $workspace = newTestWorkspace @{ PublishDir = $publishDir }
        newTsv "$tmpDir\a.tsv" @("a")
        newTsv "$publishDir\b.xlsx\b.tsv" @("b")

        removeTmpDir

        Test-Path -LiteralPath $tmpDir | Should -Be $false
        Test-Path -LiteralPath $publishDir | Should -Be $false
    }

    It "削除できなくても止まらず、次のフォルダも削除する" {
        $tmpDir = "$TestDrive\remove_fail\tmp"
        $publishDir = "$TestDrive\remove_fail\出力"
        $workspace = newTestWorkspace @{ PublishDir = $publishDir }
        newTsv "$tmpDir\使用中.tsv" @("a")
        newTsv "$publishDir\b.xlsx\b.tsv" @("b")
        # 1 つ目（作業フォルダ）の中のファイルをほかから開いておき、削除に失敗させる
        $stream = [System.IO.File]::Open("$tmpDir\使用中.tsv", "Open", "Read", "None")
        try {
            { removeTmpDir } | Should -Not -Throw
            Test-Path -LiteralPath "$tmpDir\使用中.tsv" | Should -Be $true
        } finally {
            $stream.Dispose()
        }
        Test-Path -LiteralPath $publishDir | Should -Be $false
    }

    It "空になった tmp\PC の鍵・tmp も消す" {
        $key = getMachineKey
        $tmpDir = "$TestDrive\remove_parent\tmp\$key\$PID"
        $publishDir = "$TestDrive\remove_parent\出力\$PID"
        $workspace = newTestWorkspace @{ PublishDir = $publishDir; TmpRoot = "$TestDrive\remove_parent\tmp" }
        newTsv "$tmpDir\a.tsv" @("a")
        newTsv "$publishDir\b.xlsx\b.tsv" @("b")

        removeTmpDir

        Test-Path -LiteralPath "$TestDrive\remove_parent\tmp\$key" | Should -Be $false
        Test-Path -LiteralPath "$TestDrive\remove_parent\tmp" | Should -Be $false
    }

    It "tmp\PC の鍵 にほかのフォルダが残っていれば、tmp\PC の鍵・tmp は消えない" {
        $key = getMachineKey
        $tmpDir = "$TestDrive\remove_parent_kept\tmp\$key\$PID"
        $workspace = newTestWorkspace @{ PublishDir = "$TestDrive\remove_parent_kept\出力\$PID"; TmpRoot = "$TestDrive\remove_parent_kept\tmp" }
        newTsv "$tmpDir\a.tsv" @("a")
        [System.IO.Directory]::CreateDirectory("$TestDrive\remove_parent_kept\tmp\$key\ほかのプロセス") | Out-Null

        removeTmpDir

        Test-Path -LiteralPath $tmpDir | Should -Be $false
        Test-Path -LiteralPath "$TestDrive\remove_parent_kept\tmp\$key" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\remove_parent_kept\tmp" | Should -Be $true
    }

    It "`$tmpDir が空（決める前・置けなかった）でも例外にならず、出力用のフォルダは消す" {
        $tmpDir = ""
        $publishDir = "$TestDrive\remove_untouched\出力\$PID"
        $workspace = newTestWorkspace @{ PublishDir = $publishDir; TmpRoot = "$TestDrive\remove_untouched\tmp" }
        newTsv "$publishDir\a.xlsx\a.tsv" @("a")

        { removeTmpDir } | Should -Not -Throw

        Test-Path -LiteralPath $publishDir | Should -Be $false
    }
}

Describe "removeStaleProcessDirs" -Tag Io {
    It "終了したプロセスのフォルダだけを削除し、実行中のプロセス・自分・プロセスID以外の名前のフォルダは残す" {
        $parent = "$TestDrive\stale"
        foreach ($name in @([string]$deadPid, "4", [string]$PID, "abc", "1234567890")) {
            [System.IO.Directory]::CreateDirectory("$parent\$name") | Out-Null
        }
        newTsv "$parent\$deadPid\a.tsv" @("a")

        removeStaleProcessDirs $parent

        Test-Path -LiteralPath "$parent\$deadPid" | Should -Be $false
        Test-Path -LiteralPath "$parent\4" | Should -Be $true  # System プロセス（実行中）
        Test-Path -LiteralPath "$parent\$PID" | Should -Be $true
        Test-Path -LiteralPath "$parent\abc" | Should -Be $true
        Test-Path -LiteralPath "$parent\1234567890" | Should -Be $true  # 10 桁はプロセスIDとみなさない
    }

    It "削除できないフォルダは、原因をログに出して次回に回す" {
        $parent = "$TestDrive\stale_locked"
        newTsv "$parent\$deadPid\a.tsv" @("a")
        Mock writeIndexerLog {}
        $held = [System.IO.File]::Open("$parent\$deadPid\a.tsv", [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
        try {
            { removeStaleProcessDirs $parent } | Should -Not -Throw
            Test-Path -LiteralPath "$parent\$deadPid" | Should -Be $true
        } finally {
            $held.Dispose()
        }
        Should -Invoke writeIndexerLog -Times 1 -Exactly -ParameterFilter { $text -like "*終了したプロセスの作業フォルダを削除できませんでした: *$deadPid（*" }
    }

    It "中にリンクかどうか確かめられないフォルダがあれば、何も消さず、その理由をログに出す" {
        $parent = "$TestDrive\stale_unknown"
        newTsv "$parent\$deadPid\中\a.tsv" @("a")
        newTsv "$parent\$deadPid\b.tsv" @("b")
        Mock writeIndexerLog {}
        Mock getLinkKind { return "Unknown" } -ParameterFilter { $item.Name -eq "中" }

        { removeStaleProcessDirs $parent } | Should -Not -Throw

        Test-Path -LiteralPath "$parent\$deadPid\中\a.tsv" | Should -Be $true
        Test-Path -LiteralPath "$parent\$deadPid\b.tsv" | Should -Be $true
        Should -Invoke writeIndexerLog -Times 1 -Exactly -ParameterFilter { $text -like "*確かめられないフォルダのため、消しませんでした*" }
    }

    It "終了したプロセスのフォルダの中にあるジャンクションは、先を消さずにジャンクションだけを外す" {
        $parent = "$TestDrive\stale_junction"
        $outside = "$TestDrive\stale_junction_outside"
        newTsv "$outside\外.tsv" @("外")
        newTsv "$parent\$deadPid\a.tsv" @("a")
        if (!(newTestJunction "$parent\$deadPid\リンク" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        removeStaleProcessDirs $parent

        Test-Path -LiteralPath "$parent\$deadPid" | Should -Be $false
        Test-Path -LiteralPath "$outside\外.tsv" | Should -Be $true
    }

    It "終了したプロセスの名前そのものがジャンクションなら、ジャンクションだけを外し、先は消さない" {
        $parent = "$TestDrive\stale_junction_name"
        $outside = "$TestDrive\stale_junction_name_outside"
        newTsv "$outside\外.tsv" @("外")
        [System.IO.Directory]::CreateDirectory($parent) | Out-Null
        if (!(newTestJunction "$parent\$deadPid" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        removeStaleProcessDirs $parent

        Test-Path -LiteralPath "$parent\$deadPid" | Should -Be $false
        Test-Path -LiteralPath "$outside\外.tsv" | Should -Be $true
    }
}

Describe "removeStaleTmpDirs" -Tag Io {
    It "作業フォルダ（ワークスペースの tmp）・出力用のフォルダ・前の版までの %TEMP% の、それぞれの親フォルダを片付ける" {
        $publishDir = "$TestDrive\stale_both\出力\$PID"
        $workspace = newTestWorkspace @{ PublishDir = $publishDir; TmpRoot = "$TestDrive\stale_both\tmp" }
        ${legacyTmpParent} = "$TestDrive\stale_both\legacy"
        $key = getMachineKey
        [System.IO.Directory]::CreateDirectory("$TestDrive\stale_both\tmp\$key\$deadPid") | Out-Null
        [System.IO.Directory]::CreateDirectory("$TestDrive\stale_both\出力\$deadPid") | Out-Null
        [System.IO.Directory]::CreateDirectory("$TestDrive\stale_both\legacy\$deadPid") | Out-Null

        removeStaleTmpDirs

        Test-Path -LiteralPath "$TestDrive\stale_both\tmp\$key\$deadPid" | Should -Be $false
        Test-Path -LiteralPath "$TestDrive\stale_both\出力\$deadPid" | Should -Be $false
        Test-Path -LiteralPath "$TestDrive\stale_both\legacy\$deadPid" | Should -Be $false
    }

    It "新規で開くためのコピー（open_copy）は、強制終了の片付けで消さない" {
        $workspace = newTestWorkspace @{ PublishDir = "$TestDrive\stale_open_copy\出力\$PID"; TmpRoot = "$TestDrive\stale_open_copy\tmp" }
        ${legacyTmpParent} = "$TestDrive\stale_open_copy\legacy"
        $copyDir = getOpenCopyDir $workspace
        [System.IO.Directory]::CreateDirectory("$copyDir\abc123") | Out-Null
        [System.IO.File]::WriteAllText("$copyDir\abc123\a.txt", "x")
        [System.IO.Directory]::CreateDirectory("$($workspace.TmpRoot)\$(getMachineKey)\$deadPid") | Out-Null

        removeStaleTmpDirs

        Test-Path -LiteralPath "$copyDir\abc123\a.txt" | Should -Be $true
        Test-Path -LiteralPath "$($workspace.TmpRoot)\$(getMachineKey)\$deadPid" | Should -Be $false
    }

    It "tmp\PC の鍵 がジャンクションなら、先にある終了済みのプロセスのフォルダも消さず、ログに出す" {
        $workspace = newTestWorkspace @{ PublishDir = "$TestDrive\stale_keylink\出力\$PID"; TmpRoot = "$TestDrive\stale_keylink\tmp" }
        ${legacyTmpParent} = "$TestDrive\stale_keylink\legacy"
        $outside = "$TestDrive\stale_keylink_outside"
        [System.IO.Directory]::CreateDirectory("$outside\$deadPid") | Out-Null
        [System.IO.Directory]::CreateDirectory("$TestDrive\stale_keylink\tmp") | Out-Null
        if (!(newTestJunction "$TestDrive\stale_keylink\tmp\$(getMachineKey)" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        $log = New-Object System.IO.StringWriter
        $script:indexerLog = $log

        removeStaleTmpDirs

        Test-Path -LiteralPath "$outside\$deadPid" | Should -Be $true
        $log.ToString() | Should -Match "リンク"
    }

    It "`$tmpDir が空（決める前）でも例外にならない" {
        $tmpDir = ""
        $workspace = newTestWorkspace @{ PublishDir = "$TestDrive\stale_untouched\出力\$PID"; TmpRoot = "$TestDrive\stale_untouched\tmp" }
        ${legacyTmpParent} = "$TestDrive\stale_untouched\legacy"

        { removeStaleTmpDirs } | Should -Not -Throw
    }

    It "空になった親（tmp\PC の鍵・tmp・出力用のフォルダ・前の版までの `$TEMP`\tebunko）も消す" {
        $publishDir = "$TestDrive\stale_parent\出力\$PID"
        $workspace = newTestWorkspace @{ PublishDir = $publishDir; TmpRoot = "$TestDrive\stale_parent\tmp" }
        ${legacyTmpParent} = "$TestDrive\stale_parent\legacy"
        $key = getMachineKey
        [System.IO.Directory]::CreateDirectory("$TestDrive\stale_parent\tmp\$key\$deadPid") | Out-Null
        [System.IO.Directory]::CreateDirectory("$TestDrive\stale_parent\出力\$deadPid") | Out-Null
        [System.IO.Directory]::CreateDirectory("$TestDrive\stale_parent\legacy\$deadPid") | Out-Null

        removeStaleTmpDirs

        Test-Path -LiteralPath "$TestDrive\stale_parent\tmp\$key" | Should -Be $false
        Test-Path -LiteralPath "$TestDrive\stale_parent\tmp" | Should -Be $false
        Test-Path -LiteralPath "$TestDrive\stale_parent\出力" | Should -Be $false
        Test-Path -LiteralPath "$TestDrive\stale_parent\legacy" | Should -Be $false
    }

    It "ほかのプロセス（動いている PID）のフォルダが残っていれば、tmp\PC の鍵・tmp は残る" {
        $workspace = newTestWorkspace @{ PublishDir = "$TestDrive\stale_kept\出力\$PID"; TmpRoot = "$TestDrive\stale_kept\tmp" }
        ${legacyTmpParent} = "$TestDrive\stale_kept\legacy"
        $key = getMachineKey
        [System.IO.Directory]::CreateDirectory("$TestDrive\stale_kept\tmp\$key\$deadPid") | Out-Null
        [System.IO.Directory]::CreateDirectory("$TestDrive\stale_kept\tmp\$key\$PID") | Out-Null

        removeStaleTmpDirs

        Test-Path -LiteralPath "$TestDrive\stale_kept\tmp\$key\$deadPid" | Should -Be $false
        Test-Path -LiteralPath "$TestDrive\stale_kept\tmp\$key\$PID" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\stale_kept\tmp\$key" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\stale_kept\tmp" | Should -Be $true
    }
}

Describe "initTmpDir" -Tag Io {
    It "ワークスペースの tmp の下にフォルダを作り、tmp・tmp\鍵・<PID> に NotContentIndexed を付ける" {
        $workspace = newTestWorkspace @{ TmpRoot = "$TestDrive\init_normal\tmp" }
        ${legacyTmpParent} = "$TestDrive\init_normal\legacy"

        $result = initTmpDir

        $result.Dir | Should -Be (getWorkspaceTmpDir $workspace)
        $result.Reason | Should -Be ""
        Test-Path -LiteralPath $result.Dir -PathType Container | Should -Be $true
        foreach ($dir in @($workspace.TmpRoot, (Split-Path $result.Dir -Parent), $result.Dir)) {
            ([System.IO.File]::GetAttributes($dir) -band [System.IO.FileAttributes]::NotContentIndexed) | Should -Not -Be 0
        }
    }

    It "ワークスペースのパスに [ ] があれば、作業フォルダを作らず、理由とスキップすることをログに書く" {
        $workspace = newTestWorkspace @{ TmpRoot = "$TestDrive\init_brackets\[共有]\tmp" }
        ${legacyTmpParent} = "$TestDrive\init_brackets\legacy"
        $log = New-Object System.IO.StringWriter
        $script:indexerLog = $log

        $result = initTmpDir

        $result.Dir | Should -BeNullOrEmpty
        $result.Reason | Should -Be "Brackets"
        Test-Path -LiteralPath "$TestDrive\init_brackets\[共有]\tmp" | Should -Be $false
        Test-Path -LiteralPath ${legacyTmpParent} | Should -Be $false
        $log.ToString() | Should -Match "\[ \]"
        $log.ToString() | Should -Match "スキップ"
    }

    It "強制終了などで残った、ほかのプロセスの作業フォルダを、場所を決める前に片付ける" {
        $workspace = newTestWorkspace @{ PublishDir = "$TestDrive\init_stale\出力\$PID"; TmpRoot = "$TestDrive\init_stale\tmp" }
        ${legacyTmpParent} = "$TestDrive\init_stale\legacy"
        $key = getMachineKey
        [System.IO.Directory]::CreateDirectory("$TestDrive\init_stale\tmp\$key\$deadPid") | Out-Null

        [void](initTmpDir)

        Test-Path -LiteralPath "$TestDrive\init_stale\tmp\$key\$deadPid" | Should -Be $false
    }

    It "tmp\PC の鍵 がジャンクションなら、作業フォルダを置けない扱い（Link）にし、先には何も作らず・消さない" {
        $workspace = newTestWorkspace @{ PublishDir = "$TestDrive\init_keylink\出力\$PID"; TmpRoot = "$TestDrive\init_keylink\tmp" }
        ${legacyTmpParent} = "$TestDrive\init_keylink\legacy"
        $outside = "$TestDrive\init_keylink_outside"
        newTsv "$outside\外.tsv" @("外")
        [System.IO.Directory]::CreateDirectory("$outside\$deadPid") | Out-Null
        [System.IO.Directory]::CreateDirectory("$TestDrive\init_keylink\tmp") | Out-Null
        if (!(newTestJunction "$TestDrive\init_keylink\tmp\$(getMachineKey)" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        $log = New-Object System.IO.StringWriter
        $script:indexerLog = $log

        $result = initTmpDir

        $result.Dir | Should -BeNullOrEmpty
        $result.Reason | Should -Be "Link"
        $log.ToString() | Should -Match "リンク"
        $log.ToString() | Should -Match "スキップ"
        Test-Path -LiteralPath "$outside\外.tsv" | Should -Be $true
        Test-Path -LiteralPath "$outside\$deadPid" | Should -Be $true
        Test-Path -LiteralPath "$outside\$PID" | Should -Be $false
    }

    It "先に置かれた <PID> のジャンクションは、先を消さずにジャンクションだけを外し、普通のフォルダとして作り直す" {
        $workspace = newTestWorkspace @{ PublishDir = "$TestDrive\init_pidlink\出力\$PID"; TmpRoot = "$TestDrive\init_pidlink\tmp" }
        ${legacyTmpParent} = "$TestDrive\init_pidlink\legacy"
        $outside = "$TestDrive\init_pidlink_outside"
        newTsv "$outside\外.tsv" @("外")
        $keyDir = "$TestDrive\init_pidlink\tmp\$(getMachineKey)"
        [System.IO.Directory]::CreateDirectory($keyDir) | Out-Null
        if (!(newTestJunction "$keyDir\$PID" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        $result = initTmpDir

        $result.Dir | Should -Be "$keyDir\$PID"
        testPathIsLink $result.Dir | Should -Be $false
        Test-Path -LiteralPath $result.Dir -PathType Container | Should -Be $true
        Test-Path -LiteralPath "$outside\外.tsv" | Should -Be $true
        @(Get-ChildItem -LiteralPath $outside -Force).Count | Should -Be 1
    }

    It "tmp そのものがジャンクションでも、これまでどおり作業フォルダを置ける" {
        $workspace = newTestWorkspace @{ PublishDir = "$TestDrive\init_rootlink\出力\$PID"; TmpRoot = "$TestDrive\init_rootlink\tmp" }
        ${legacyTmpParent} = "$TestDrive\init_rootlink\legacy"
        [System.IO.Directory]::CreateDirectory("$TestDrive\init_rootlink") | Out-Null
        if (!(newTestJunction "$TestDrive\init_rootlink\tmp" "$TestDrive\init_rootlink_target")) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        $result = initTmpDir

        $result.Dir | Should -Be (getWorkspaceTmpDir $workspace)
        $result.Reason | Should -Be ""
        Test-Path -LiteralPath $result.Dir -PathType Container | Should -Be $true
    }

    It "作業フォルダ（<PID>）を作ったあとリンクかどうか確かめられなければ、理由 Link で置けない扱いにし、ログに出す" {
        $workspace = newTestWorkspace @{ PublishDir = "$TestDrive\init_unplaced\出力\$PID"; TmpRoot = "$TestDrive\init_unplaced\tmp" }
        ${legacyTmpParent} = "$TestDrive\init_unplaced\legacy"
        Mock writeIndexerLog {}
        Mock testPathIsLink { return $true } -ParameterFilter { $path -like "*\$PID" }

        $result = initTmpDir

        $result.Dir | Should -Be ""
        $result.Reason | Should -Be "Link"
        Should -Invoke writeIndexerLog -Times 1 -Exactly -ParameterFilter { $text -like "*取り込みをすべてスキップします。" }
    }
}

Describe "newDirectoryReplacingExisting" -Tag Io {
    It "無ければ作り、あれば中身ごと作り直して `$true を返す" {
        newDirectoryReplacingExisting "$TestDrive\replace_new\d" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\replace_new\d" -PathType Container | Should -Be $true

        writeListFile "$TestDrive\replace_old\d\古.tsv" @("古")
        newDirectoryReplacingExisting "$TestDrive\replace_old\d" | Should -Be $true
        @(Get-ChildItem -LiteralPath "$TestDrive\replace_old\d" -Force).Count | Should -Be 0
    }

    It "リンクがあれば、先は消さず、リンクを外して普通のフォルダを作る" {
        writeListFile "$TestDrive\replace_link_target\外.tsv" @("外")
        [System.IO.Directory]::CreateDirectory("$TestDrive\replace_link") | Out-Null
        if (!(newTestJunction "$TestDrive\replace_link\d" "$TestDrive\replace_link_target")) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        newDirectoryReplacingExisting "$TestDrive\replace_link\d" | Should -Be $true

        testPathIsLink "$TestDrive\replace_link\d" | Should -Be $false
        readListFile "$TestDrive\replace_link_target\外.tsv" | Should -Be @("外")
    }

    It "作り直したあともリンク（または確かめられないもの）が残っていれば、`$false を返す" {
        Mock testPathIsLink { return $true } -ParameterFilter { $path -like "*\replace_unknown\d" }

        newDirectoryReplacingExisting "$TestDrive\replace_unknown\d" | Should -Be $false
    }
}

Describe "newWorkerTmpDir" -Tag Io {
    It "親フォルダに NotContentIndexed が付いていれば、作業フォルダ（w<番号>）にも同じ属性を付ける" {
        $parent = "$TestDrive\worker_nci\1234"
        [System.IO.Directory]::CreateDirectory($parent) | Out-Null
        [System.IO.File]::SetAttributes($parent, [System.IO.File]::GetAttributes($parent) -bor [System.IO.FileAttributes]::NotContentIndexed)

        $dir = newWorkerTmpDir $parent 2

        $dir | Should -Be "$parent\w2"
        Test-Path -LiteralPath $dir -PathType Container | Should -Be $true
        ([System.IO.File]::GetAttributes($dir) -band [System.IO.FileAttributes]::NotContentIndexed) | Should -Not -Be 0
    }

    It "親フォルダに NotContentIndexed が付いていなければ、作業フォルダにも付けない" {
        $parent = "$TestDrive\worker_plain\5678"
        [System.IO.Directory]::CreateDirectory($parent) | Out-Null

        $dir = newWorkerTmpDir $parent 1

        $dir | Should -Be "$parent\w1"
        ([System.IO.File]::GetAttributes($dir) -band [System.IO.FileAttributes]::NotContentIndexed) | Should -Be 0
    }

    It "親フォルダが空（置けなかった）なら、何も作らず空文字を返す" {
        { newWorkerTmpDir "" 1 } | Should -Not -Throw
        newWorkerTmpDir "" 1 | Should -BeNullOrEmpty
    }

    It "先に置かれた w<番号> のジャンクションは、先を消さずにジャンクションだけを外し、普通のフォルダとして作り直す" {
        $parent = "$TestDrive\worker_link\1234"
        $outside = "$TestDrive\worker_link_outside"
        newTsv "$outside\外.tsv" @("外")
        [System.IO.Directory]::CreateDirectory($parent) | Out-Null
        if (!(newTestJunction "$parent\w1" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        $dir = newWorkerTmpDir $parent 1

        $dir | Should -Be "$parent\w1"
        testPathIsLink $dir | Should -Be $false
        Test-Path -LiteralPath "$outside\外.tsv" | Should -Be $true
        @(Get-ChildItem -LiteralPath $outside -Force).Count | Should -Be 1
    }

    It "親フォルダがジャンクションなら、何も作らず空文字を返し、ログに出す" {
        $outside = "$TestDrive\worker_parentlink_outside"
        [System.IO.Directory]::CreateDirectory("$TestDrive\worker_parentlink") | Out-Null
        if (!(newTestJunction "$TestDrive\worker_parentlink\1234" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }
        $log = New-Object System.IO.StringWriter
        $script:indexerLog = $log

        newWorkerTmpDir "$TestDrive\worker_parentlink\1234" 1 | Should -BeNullOrEmpty

        @(Get-ChildItem -LiteralPath $outside -Force).Count | Should -Be 0
        $log.ToString() | Should -Match "リンク"
    }

    It "作り直した w<番号> がリンク（または確かめられないもの）のままなら、空文字を返す" {
        $parent = "$TestDrive\worker_unknown\1234"
        [System.IO.Directory]::CreateDirectory($parent) | Out-Null
        Mock testPathIsLink { return $true } -ParameterFilter { $path -like "*\w7" }

        newWorkerTmpDir $parent 7 | Should -BeNullOrEmpty
    }
}

Describe "findDroppedIndexes" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "今回の設定に無い名前だけを返す"; current = @("営業"); previous = @("営業", "技術", "経理"); expected = @("技術", "経理") }
        @{ name = "大文字・小文字だけ違う名前は、同じインデックスとして返さない"; current = @("sales"); previous = @("Sales"); expected = @() }
        @{ name = "前回が null なら空"; current = @("営業"); previous = $null; expected = @() }
    ) {
        param ($name, $current, $previous, $expected)
        $folders = @($current | ForEach-Object { [pscustomobject]@{ Path = "C:\data\$_"; Name = $_ } })
        $prev = if ($null -eq $previous) { $null } else { @($previous | ForEach-Object { [pscustomobject]@{ Path = "C:\data\$_"; Name = $_ } }) }
        $result = @(findDroppedIndexes $folders $prev)
        @($result | ForEach-Object { $_.Name }) -join "," | Should -Be ($expected -join ",")
    }
}

Describe "getDroppedStatusRows" -Tag Unit {
    It "外れたインデックスの前回の行だけを返す（名前の先頭が同じだけの別のインデックスは返さない）" {
        $rows = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
        $rows["技術\a.docx"] = [pscustomobject]@{ 相対パス = "技術\a.docx" }
        $rows["技術2\b.docx"] = [pscustomobject]@{ 相対パス = "技術2\b.docx" }
        $rows["営業\c.docx"] = [pscustomobject]@{ 相対パス = "営業\c.docx" }
        $result = @(getDroppedStatusRows @([pscustomobject]@{ Path = "C:\data\技術"; Name = "技術" }) $rows)
        @($result | ForEach-Object { $_.相対パス }) -join "," | Should -Be "技術\a.docx"
    }

    It "外れたものが無い・前回の行が無いなら空" {
        @(getDroppedStatusRows @() @{}).Count | Should -Be 0
        @(getDroppedStatusRows @([pscustomobject]@{ Path = "C:\x"; Name = "x" }) $null).Count | Should -Be 0
    }
}

Describe "writeStatusKeepingDropped" -Tag Unit {
    BeforeAll {
        function newFakeLedger {
            $ledger = [pscustomobject]@{ Folders = $null; Rows = $null }
            $ledger | Add-Member -MemberType ScriptMethod -Name WriteStatus -Value {
                param ($folders, $rows)
                $this.Folders = @($folders)
                $this.Rows = @($rows)
            }
            return $ledger
        }
    }

    It "外れたインデックス <kept> 件のとき、フォルダ <folderCount>・行 <rowCount> を渡す" -TestCases @(
        @{ kept = 0; folderCount = 1; rowCount = 1 }
        @{ kept = 1; folderCount = 2; rowCount = 2 }
    ) {
        param ($kept, $folderCount, $rowCount)
        $ledger = newFakeLedger
        $previous = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
        $previous["技術\a.docx"] = [pscustomobject]@{ 相対パス = "技術\a.docx" }
        $dropped = @()
        if ($kept -gt 0) { $dropped = @([pscustomobject]@{ Path = "C:\data\技術"; Name = "技術" }) }
        $folder = [pscustomobject]@{ Path = "C:\data\営業"; Name = "営業" }
        $row = [pscustomobject]@{ 相対パス = "営業\c.docx" }
        writeStatusKeepingDropped $ledger @($folder) @($row) $dropped $previous
        $ledger.Folders.Count | Should -Be $folderCount
        $ledger.Rows.Count | Should -Be $rowCount
        $ledger.Folders[0].Name | Should -Be "営業"
        if ($kept -gt 0) {
            $ledger.Folders[1].Name | Should -Be "技術"
            $ledger.Rows[1].相対パス | Should -Be "技術\a.docx"
        }
    }
}

Describe "writeStatusKeepingDropped（今の決まりに合わない前の版の名前）" -Tag Io {
    It "取りやめた回に書き戻した <name> のフォルダの行は、読み直すと LegacyFolders に入り、各ファイルの行は残らない" -TestCases @(
        @{ name = "営業." }
        @{ name = "営業 " }
    ) {
        param ($name)
        $path = "$TestDrive\keep\ingest_status.tsv"
        [System.IO.Directory]::CreateDirectory("$TestDrive\keep") | Out-Null
        $ledger = [pscustomobject]@{ StatusPath = $path }
        $ledger | Add-Member -MemberType ScriptMethod -Name WriteStatus -Value {
            param ($folders, $rows)
            writeStatusFile $folders $rows $this.StatusPath
        }
        $folder = [pscustomobject]@{ Path = "C:\data\営業"; Name = "営業" }
        $oldFolder = [pscustomobject]@{ Path = "C:\data\古い"; Name = $name }
        $row = [pscustomobject]@{ 相対パス = "営業\c.docx"; 更新日時 = "2025/01/10 12:34:56"; サイズ = "1"; 状態 = $stateDone; 抽出版 = "1"; エラー = "" }
        $oldRow = [pscustomobject]@{ 相対パス = "$name\a.docx"; 更新日時 = "2025/01/10 12:34:56"; サイズ = "1"; 状態 = $stateDone; 抽出版 = "1"; エラー = "" }
        # 前の版が書いた取り込み一覧（前の版の名前のフォルダの行と、その各ファイルの行）を読む
        writeStatusFile @($folder, $oldFolder) @($row, $oldRow) $path
        $before = readStatusFile $path
        @($before.LegacyFolders | ForEach-Object { $_.Name }) | Should -Be @($name)
        $before.Rows.ContainsKey("$name\a.docx") | Should -Be $false

        writeStatusKeepingDropped $ledger @($folder) @($row) $before.LegacyFolders.ToArray() $before.Rows

        $status = readStatusFile $path
        @($status.Folders | ForEach-Object { $_.Name }) | Should -Be @("営業")
        @($status.LegacyFolders | ForEach-Object { $_.Name }) | Should -Be @($name)
        @($status.Rows.Keys) | Should -Be @("営業\c.docx")
    }

    It "予約名（<name>）のフォルダの行は、前の版の名前にせず、ふつうのフォルダと各ファイルの行として読む" -TestCases @(
        @{ name = "aux.backup" }
        @{ name = "CON" }
    ) {
        param ($name)
        $path = "$TestDrive\keep_reserved\ingest_status.tsv"
        [System.IO.Directory]::CreateDirectory("$TestDrive\keep_reserved") | Out-Null
        $oldFolder = [pscustomobject]@{ Path = "C:\data\古い"; Name = $name }
        $oldRow = [pscustomobject]@{ 相対パス = "$name\a.docx"; 更新日時 = "2025/01/10 12:34:56"; サイズ = "1"; 状態 = $stateDone; 抽出版 = "1"; エラー = "" }
        writeStatusFile @($oldFolder) @($oldRow) $path

        $status = readStatusFile $path

        @($status.Folders | ForEach-Object { $_.Name }) | Should -Be @($name)
        $status.LegacyFolders.Count | Should -Be 0
        $status.Rows.ContainsKey("$name\a.docx") | Should -Be $true
    }
}

Describe "removeDroppedFolders" -Tag Io {
    It "クロール対象から削除されたフォルダのインデックスだけを削除する" {
        $indexDir = "$TestDrive\dropped\index"
        $workspace = newTestWorkspace @{ IndexDir = $indexDir }
        newTsv "$indexDir\営業\見積.xlsx\Sheet1.tsv" @("a")
        newTsv "$indexDir\技術\仕様.docx\ページ001.tsv" @("b")
        $folders = @([pscustomobject]@{ Path = "C:\data\営業"; Name = "営業" })
        $previous = @(
            [pscustomobject]@{ Path = "C:\data\営業"; Name = "営業" },
            [pscustomobject]@{ Path = "C:\data\技術"; Name = "技術" },
            [pscustomobject]@{ Path = "C:\data\消えた"; Name = "消えた" }  # インデックスのフォルダが無い
        )

        removeDroppedFolders @(findDroppedIndexes $folders $previous)

        Test-Path -LiteralPath "$indexDir\営業\見積.xlsx\Sheet1.tsv" | Should -Be $true
        Test-Path -LiteralPath "$indexDir\技術" | Should -Be $false
    }

    It "インデックス名に [ ] があり、中に 260 文字を超えるパスがあっても削除する" {
        $indexDir = "$TestDrive\dropped2\index"
        $workspace = newTestWorkspace @{ IndexDir = $indexDir }
        $deep = "$indexDir\[旧]営業\" + ("深いフォルダ" * 20) + "\" + ("もっと深いフォルダ" * 15)
        newTsv (toLongPath "$deep\見積.xlsx\Sheet1.tsv") @("a")
        newTsv "$indexDir\[旧]営業2\見積.xlsx\Sheet1.tsv" @("b")  # 名前の先頭が同じだけの別のインデックスは残す

        removeDroppedFolders @(findDroppedIndexes @([pscustomobject]@{ Path = "C:\data\営業2"; Name = "[旧]営業2" }) @(
            [pscustomobject]@{ Path = "C:\data\旧営業"; Name = "[旧]営業" },
            [pscustomobject]@{ Path = "C:\data\営業2"; Name = "[旧]営業2" }
        ))

        [System.IO.Directory]::Exists((toLongPath "$indexDir\[旧]営業")) | Should -Be $false
        Test-Path -LiteralPath "$indexDir\[旧]営業2\見積.xlsx\Sheet1.tsv" | Should -Be $true
    }

    It "インデックスの中にあるジャンクションの先は消さず、ジャンクションだけを外す" {
        $indexDir = "$TestDrive\dropped_junction\index"
        $workspace = newTestWorkspace @{ IndexDir = $indexDir }
        $outside = "$TestDrive\dropped_junction_outside"
        newTsv "$outside\外.tsv" @("外")
        newTsv "$indexDir\技術\仕様.docx\ページ001.tsv" @("b")
        if (!(newTestJunction "$indexDir\技術\リンク" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        removeDroppedFolders @(findDroppedIndexes @() @([pscustomobject]@{ Path = "C:\data\技術"; Name = "技術" }))

        Test-Path -LiteralPath "$indexDir\技術" | Should -Be $false
        Test-Path -LiteralPath "$outside\外.tsv" | Should -Be $true
    }

    It "インデックスのフォルダそのものがジャンクションなら、ジャンクションだけを外し、先は消さない" {
        $indexDir = "$TestDrive\dropped_rootlink\index"
        $workspace = newTestWorkspace @{ IndexDir = $indexDir }
        $outside = "$TestDrive\dropped_rootlink_outside"
        newTsv "$outside\外.tsv" @("外")
        [System.IO.Directory]::CreateDirectory($indexDir) | Out-Null
        if (!(newTestJunction "$indexDir\技術" $outside)) {
            Set-ItResult -Skipped -Because $script:linkError
            return
        }

        removeDroppedFolders @(findDroppedIndexes @() @([pscustomobject]@{ Path = "C:\data\技術"; Name = "技術" }))

        Test-Path -LiteralPath "$indexDir\技術" | Should -Be $false
        Test-Path -LiteralPath "$outside\外.tsv" | Should -Be $true
    }

    It "予約名に拡張子の付いた名前（<name>）のインデックスも消し、消えてから「削除しました」を記録する" -TestCases @(
        @{ name = "aux.backup" }
        @{ name = "CON.txt" }
    ) {
        $indexDir = "$TestDrive\dropped_reserved_$name\index"
        $workspace = newTestWorkspace @{ IndexDir = $indexDir }
        $long = toLongPath "$indexDir\$name\見積.xlsx"
        [System.IO.Directory]::CreateDirectory($long) | Out-Null
        [System.IO.File]::WriteAllText("$long\Sheet1.tsv", "a")
        Mock writeIndexerLog {}

        removeDroppedFolders @(findDroppedIndexes @() @([pscustomobject]@{ Path = "C:\data\予約"; Name = $name }))

        [System.IO.Directory]::Exists((toLongPath "$indexDir\$name")) | Should -Be $false
        Should -Invoke writeIndexerLog -Times 1 -Exactly -Scope It -ParameterFilter { $text -like "*削除しました*" }
    }

    It "フォルダが消えなければ例外にし、「削除しました」を記録しない" {
        Mock Start-Sleep {}
        Mock writeIndexerLog {}
        $indexDir = "$TestDrive\dropped_locked\index"
        $workspace = newTestWorkspace @{ IndexDir = $indexDir }
        newTsv "$indexDir\技術\仕様.docx\ページ001.tsv" @("b")
        $stream = [System.IO.File]::Open("$indexDir\技術\仕様.docx\ページ001.tsv", "Open", "Read", "None")
        try {
            { removeDroppedFolders @(findDroppedIndexes @() @([pscustomobject]@{ Path = "C:\data\技術"; Name = "技術" })) } | Should -Throw
        } finally {
            $stream.Dispose()
        }
        Should -Invoke writeIndexerLog -Times 0 -Exactly -Scope It -ParameterFilter { $text -like "*削除しました*" }
    }

    It "前回のクロール対象フォルダが無ければ何も削除しない" {
        $indexDir = "$TestDrive\dropped3\index"
        $workspace = newTestWorkspace @{ IndexDir = $indexDir }
        newTsv "$indexDir\営業\見積.xlsx\Sheet1.tsv" @("a")

        removeDroppedFolders @(findDroppedIndexes @() $null)

        Test-Path -LiteralPath "$indexDir\営業\見積.xlsx\Sheet1.tsv" | Should -Be $true
    }

    It "インデックス名の大文字・小文字だけが違うフォルダは、同じフォルダとして残す" {
        $indexDir = "$TestDrive\dropped4\index"
        $workspace = newTestWorkspace @{ IndexDir = $indexDir }
        newTsv "$indexDir\Sales\見積.xlsx\Sheet1.tsv" @("a")

        removeDroppedFolders @(findDroppedIndexes @([pscustomobject]@{ Path = "C:\data\sales"; Name = "sales" }) @([pscustomobject]@{ Path = "C:\data\Sales"; Name = "Sales" }))

        Test-Path -LiteralPath "$indexDir\Sales\見積.xlsx\Sheet1.tsv" | Should -Be $true
    }

    It "置き場の外を指す名前は削除せず、例外にもしないで、次の名前を続ける" -TestCases @(
        @{ badName = "..\外" }
        @{ badName = "..\..\外" }
        @{ badName = "営業\..\..\外" }
        @{ badName = ".." }
        @{ badName = "." }
        @{ badName = "" }
    ) {
        $indexDir = "$TestDrive\dropped5\index"
        $workspace = newTestWorkspace @{ IndexDir = $indexDir }
        newTsv "$TestDrive\dropped5\外\a.tsv" @("a")
        newTsv "$TestDrive\外\a.tsv" @("a")
        newTsv "$indexDir\営業\見積.xlsx\Sheet1.tsv" @("b")
        newTsv "$indexDir\技術\仕様.docx\ページ001.tsv" @("c")
        Mock writeIndexerLog {}

        { removeDroppedFolders @(
            [pscustomobject]@{ Path = "C:\data\壊れた"; Name = $badName },
            [pscustomobject]@{ Path = "C:\data\技術"; Name = "技術" }
        ) } | Should -Not -Throw

        Test-Path -LiteralPath "$TestDrive\dropped5\外\a.tsv" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\外\a.tsv" | Should -Be $true
        Test-Path -LiteralPath "$indexDir\営業\見積.xlsx\Sheet1.tsv" | Should -Be $true
        Test-Path -LiteralPath "$indexDir\技術" | Should -Be $false
        Should -Invoke writeIndexerLog -Times 1 -Exactly -Scope It -ParameterFilter { $text -like "*削除しませんでした*" -and $color -eq "Yellow" }
    }

    It "前の版で作られた名前（最後の . ・前後の空白）のフォルダも、実在のフォルダを列挙して削除し、似た名前の別のインデックスは残す" -TestCases @(
        @{ legacyName = "営業." }
        @{ legacyName = "営業 " }
        @{ legacyName = " 営業" }
    ) {
        param ($legacyName)
        $indexDir = "$TestDrive\dropped6\index"
        $workspace = newTestWorkspace @{ IndexDir = $indexDir }
        newTsv (toLongPath "$indexDir\$legacyName\見積.xlsx\Sheet1.tsv") @("a")
        newTsv "$indexDir\営業\見積.xlsx\Sheet1.tsv" @("b")   # 名前の最後の . や空白を取ると同じになる別のインデックス
        newTsv "$TestDrive\dropped6\外\a.tsv" @("c")

        removeDroppedFolders @([pscustomobject]@{ Path = "C:\data\旧営業"; Name = $legacyName })

        [System.IO.Directory]::Exists((toLongPath "$indexDir\$legacyName")) | Should -Be $false
        Test-Path -LiteralPath "$indexDir\営業\見積.xlsx\Sheet1.tsv" | Should -Be $true
        Test-Path -LiteralPath "$TestDrive\dropped6\外\a.tsv" | Should -Be $true
    }
}
