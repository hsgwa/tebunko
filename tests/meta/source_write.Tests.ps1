# ファイルに書く・属性や日時を変える API の許可リスト。
# 取り込み対象の元のファイルを書き換えない決まり（docs/safety/file-access.md「取り込み対象のファイルは書き換えない」）を、
# 変数の名前でなく書く API そのものを見て守る。scripts の下で、ファイルに書く・属性や日時を変える・ファイルを作る呼び出しは、
# 理由つきの一覧（${writeAllowed}）にある関数の中だけにする。一覧に無い関数に足すと落ちるので、書き先が
# ワークスペースの置き場（インデックス・作業フォルダ・設定・ログ）であることを確かめて、Reason を書いて足す。
# 見張るもの（構文木で見つける。フォルダを消す・移す・作るは remove_entry.Tests.ps1 の担当で、ここでは見ない）:
#   - コマンド: Set-Content・Add-Content・Out-File・Copy-Item・Move-Item・Rename-Item・New-Item・Set-ItemProperty・Clear-Content とその別名、
#     New-Object の StreamWriter・FileStream・BinaryWriter、リダイレクト（> >>。$null への捨て先を除く）
#   - メソッド: SetAttributes・Set*Time・WriteAll*・AppendAll*・AppendText・CreateText・Create・OpenWrite・Copy・Move・Replace・
#     ExtractToFile・Save・SaveAs・SaveAs2・SaveCopyAs・SetAccessControl、Open（FileAccess.Read と FileMode.Open を両方書いた形を除く）、
#     StreamWriter・FileStream・BinaryWriter の生成（[型]::new も）
#   - 型: FileMode（Open 以外）・FileAccess（Read 以外）を書いた所。型を変数に入れて渡す書き方（[System.IO.FileMode] そのもの・キャスト）も拾う
# 守れない範囲（構文木で見つけられない書き方）:
#   - コマンド名・メソッド名を変数や文字列の組み立てで作る呼び出し（& $name ... や $obj.$member(...)）
#   - 外のプログラムに書かせる書き方（cmd /c copy・robocopy・Excel の COM を除く）
#   - 一覧にある関数の中で、書き先を取り込み対象のフォルダにしてしまうこと。書き先の確かめは各関数のテストが守る
#     （元は copyFileShared が読み取りだけで開いて作業フォルダの写しを作る。Office は写しを開く）
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"

    ${writeCommands} = @{
        "set-content" = "Set-Content"; "sc" = "Set-Content"
        "add-content" = "Add-Content"; "ac" = "Add-Content"
        "out-file" = "Out-File"
        "copy-item" = "Copy-Item"; "copy" = "Copy-Item"; "cp" = "Copy-Item"; "cpi" = "Copy-Item"
        "move-item" = "Move-Item"; "move" = "Move-Item"; "mv" = "Move-Item"; "mi" = "Move-Item"
        "rename-item" = "Rename-Item"; "ren" = "Rename-Item"; "rni" = "Rename-Item"
        "new-item" = "New-Item"; "ni" = "New-Item"
        "set-itemproperty" = "Set-ItemProperty"; "sp" = "Set-ItemProperty"
        "clear-content" = "Clear-Content"; "clc" = "Clear-Content"
    }
    ${writeMembers} = @("SetAttributes", "SetLastWriteTime", "SetLastWriteTimeUtc", "SetCreationTime", "SetCreationTimeUtc",
        "SetLastAccessTime", "SetLastAccessTimeUtc", "SetAccessControl", "WriteAllText", "WriteAllLines", "WriteAllBytes",
        "AppendAllText", "AppendAllLines", "AppendText", "CreateText", "Create", "CreateNew", "OpenWrite", "Copy", "CopyTo", "Move", "MoveTo",
        "Replace", "ExtractToFile", "Save", "SaveAs", "SaveAs2", "SaveCopyAs")
    ${writerTypes} = @("StreamWriter", "FileStream", "BinaryWriter")
    ${writeProperties} = @("Attributes", "IsReadOnly", "LastWriteTime", "LastWriteTimeUtc", "CreationTime", "CreationTimeUtc", "LastAccessTime", "LastAccessTimeUtc")

    function getLastName {
        param ([string]$name)
        return $name.Split(".")[-1]
    }

    function testTypeNamed {
        # 型の式（[型]）が、名前（System.IO. の有無は問わない）のどれかか
        param ($ast, [string[]]$names)
        return ($ast -is [System.Management.Automation.Language.TypeExpressionAst]) -and ($names -contains (getLastName $ast.TypeName.Name))
    }

    function testReadOnlyOpen {
        # Open・FileStream の生成の引数が、FileMode.Open と FileAccess.Read の両方を書いた形か（読み取りだけで開く）
        param ($arguments)
        $hasMode = $false
        $hasRead = $false
        foreach ($argument in @($arguments)) {
            if ($argument -is [System.Management.Automation.Language.MemberExpressionAst] -and $argument.Member -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                if ((testTypeNamed $argument.Expression @("FileMode")) -and $argument.Member.Value -eq "Open") { $hasMode = $true }
                if ((testTypeNamed $argument.Expression @("FileAccess")) -and $argument.Member.Value -eq "Read") { $hasRead = $true }
            }
        }
        return $hasMode -and $hasRead
    }

    function testWriteMemberGate {
        # 名前が紛らわしいメンバー（Copy・Move・Replace・Create・CreateNew。string.Replace・Array.Copy・[powershell]::Create() など）は、
        # ファイルを扱う型の静的な呼び出し（型が変数で読めない場合も含む）だけを見る。ほかのメンバーは型を問わず見る
        param ($node, [string]$member)
        if (@("Copy", "Move", "Replace", "Create", "CreateNew") -notcontains $member) {
            return $true
        }
        if (!$node.Static) {
            return $false
        }
        if ($node.Expression -isnot [System.Management.Automation.Language.TypeExpressionAst]) {
            return $true
        }
        return (@("File", "FileInfo", "ZipFile", "ZipFileExtensions", "FileSystem") -contains (getLastName $node.Expression.TypeName.Name))
    }

    function findOwner {
        # ノードを含む、いちばん内側の関数（クラスのメソッドを含む）の名前。関数の外は "(スクリプト直下)"
        param ($node)
        for ($parent = $node.Parent; $null -ne $parent; $parent = $parent.Parent) {
            if ($parent -is [System.Management.Automation.Language.FunctionDefinitionAst] -or $parent -is [System.Management.Automation.Language.FunctionMemberAst]) {
                return $parent.Name
            }
        }
        return "(スクリプト直下)"
    }

    function findWriteCalls {
        # ast の中の、ファイルに書く・属性や日時を変える・ファイルを作る呼び出しを、@{ Owner; Call } で返す
        param ($ast)
        $nodes = $ast.FindAll({
                param ($n)
                $n -is [System.Management.Automation.Language.CommandAst] -or
                $n -is [System.Management.Automation.Language.MemberExpressionAst] -or
                $n -is [System.Management.Automation.Language.FileRedirectionAst] -or
                $n -is [System.Management.Automation.Language.AssignmentStatementAst] -or
                $n -is [System.Management.Automation.Language.TypeConstraintAst] -or
                $n -is [System.Management.Automation.Language.TypeExpressionAst]
            }, $true)
        foreach ($node in $nodes) {
            $call = @()
            if ($node -is [System.Management.Automation.Language.AssignmentStatementAst]) {
                # $info.Attributes = ...・$info.LastWriteTime = ...（FileInfo・DirectoryInfo のプロパティへの代入）
                $left = $node.Left
                if ($left -is [System.Management.Automation.Language.MemberExpressionAst] -and $left.Member -is [System.Management.Automation.Language.StringConstantExpressionAst] -and
                    ${writeProperties} -contains $left.Member.Value) {
                    $call = @("$($left.Member.Value) =")
                }
            } elseif ($node -is [System.Management.Automation.Language.FileRedirectionAst]) {
                if ($node.Location.Extent.Text -ne '$null') { $call = @("リダイレクト") }
            } elseif ($node -is [System.Management.Automation.Language.CommandAst]) {
                $name = (getLastName "$($node.GetCommandName())".Split("\")[-1]).ToLowerInvariant()
                if ($name -eq "new-object") {
                    foreach ($element in $node.CommandElements) {
                        if ($element -is [System.Management.Automation.Language.StringConstantExpressionAst] -and ${writerTypes} -contains (getLastName $element.Value)) {
                            # FileMode.Open と FileAccess.Read を両方書いた FileStream は、読むだけ
                            $modeNodes = @($node.FindAll({ param ($n) $n -is [System.Management.Automation.Language.MemberExpressionAst] }, $true))
                            if (!((getLastName $element.Value) -eq "FileStream" -and (testReadOnlyOpen $modeNodes))) {
                                $call = @("New-Object $(getLastName $element.Value)")
                            }
                        }
                    }
                } elseif (${writeCommands}.ContainsKey($name)) {
                    $call = @(${writeCommands}[$name])
                }
            } elseif ($node -is [System.Management.Automation.Language.MemberExpressionAst]) {
                $member = if ($node.Member -is [System.Management.Automation.Language.StringConstantExpressionAst]) { $node.Member.Value } else { "" }
                $isInvoke = $node -is [System.Management.Automation.Language.InvokeMemberExpressionAst]
                $arguments = if ($isInvoke) { @($node.Arguments) } else { @() }
                if (testTypeNamed $node.Expression @("FileMode")) {
                    if ($member -ne "Open") { $call = @("FileMode.$member") }
                } elseif (testTypeNamed $node.Expression @("FileAccess")) {
                    if ($member -ne "Read") { $call = @("FileAccess.$member") }
                } elseif ($isInvoke -and $member -eq "new" -and $node.Static -and $node.Expression -is [System.Management.Automation.Language.TypeExpressionAst] -and
                    ${writerTypes} -contains (getLastName $node.Expression.TypeName.Name)) {
                    $typeName = getLastName $node.Expression.TypeName.Name
                    if (!($typeName -eq "FileStream" -and (testReadOnlyOpen $arguments))) { $call = @("$typeName::new") }
                } elseif ($isInvoke -and $member -eq "Open") {
                    # 引数なしの .Open()（runspace・COM のアプリなど）は、ファイルを開くものではない
                    if ($arguments.Count -gt 0 -and !(testReadOnlyOpen $arguments)) { $call = @("Open") }
                } elseif ($isInvoke -and ${writeMembers} -contains $member) {
                    if (testWriteMemberGate $node $member) { $call = @($member) }
                }
            } elseif ($node -is [System.Management.Automation.Language.TypeConstraintAst]) {
                # [System.IO.FileMode]"Create"・[System.IO.FileMode]$mode（キャスト・型の制約）
                $typeName = getLastName $node.TypeName.Name
                if (@("FileMode", "FileAccess") -contains $typeName) { $call = @("$typeName（型の参照）") }
            } elseif ($node -is [System.Management.Automation.Language.TypeExpressionAst]) {
                # [FileMode]::Create のようにメンバーを書いた所は、上の MemberExpressionAst が拾う。
                # 型だけを書いた所（[System.IO.FileMode] を変数に入れる・キャスト・Parse の引数）は、変数越しに渡すかもしれないので拾う
                $typeName = getLastName $node.TypeName.Name
                if (@("FileMode", "FileAccess") -contains $typeName -and
                    !($node.Parent -is [System.Management.Automation.Language.MemberExpressionAst] -and $node.Parent.Expression -eq $node)) {
                    $call = @("$typeName（型の参照）")
                }
            }
            foreach ($one in $call) {
                @{ Owner = (findOwner $node); Call = $one }
            }
        }
    }

    function parseCode {
        param ([string]$code)
        $errors = $null
        $tokens = $null
        return [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)
    }

    function getScriptFiles {
        $files = @(Get-ChildItem -LiteralPath ${scriptsDir} -Filter "*.ps1" -Recurse)
        $files.Count | Should -BeGreaterThan 0 -Because "scripts の下の .ps1 が見つからないと、何も調べずに通ってしまう"
        return $files
    }

    function getRelativeName {
        param ($file)
        return $file.FullName.Substring(${scriptsDir}.Length).TrimStart("\")
    }

    function getAllWriteCalls {
        # scripts の下の全部の書く呼び出しを "ファイル|関数|呼び出し" の配列で返す
        $found = New-Object System.Collections.Generic.List[string]
        foreach ($file in getScriptFiles) {
            $errors = $null
            $tokens = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
            $relative = getRelativeName $file
            foreach ($call in @(findWriteCalls $ast)) {
                $found.Add("$relative|$($call.Owner)|$($call.Call)")
            }
        }
        return @($found | Sort-Object -Unique)
    }

    # 書く API を使ってよい関数。足すときは Reason に、書き先が取り込み対象でなく、ワークスペースの置き場か利用者が選んだ出力先であることを書く。
    # File = scripts からの相対パス、Owner = 関数（クラスのメソッドを含む）の名前、Call = 呼び出し
    ${writeAllowed} = @(
        @{ File = "shared\core\close_trace.ps1"; Owner = "writeCloseTrace"; Call = "AppendAllText"; Reason = "閉じる処理の診断記録に追記する（ログの置き場）" }
        @{ File = "shared\core\data_dir.ps1"; Owner = "testWritableFolder"; Call = "FileAccess.Write"; Reason = "設定の置き場に書けるかを、使い捨てのファイル（その場で消す）で確かめる。置き場は呼ぶ側が決めたツールのフォルダ" }
        @{ File = "shared\core\data_dir.ps1"; Owner = "testWritableFolder"; Call = "FileMode.CreateNew"; Reason = "設定の置き場に書けるかを、使い捨てのファイル（その場で消す）で確かめる。置き場は呼ぶ側が決めたツールのフォルダ" }
        @{ File = "shared\core\data_dir.ps1"; Owner = "testWritableFolder"; Call = "New-Object FileStream"; Reason = "設定の置き場に書けるかを、使い捨てのファイル（その場で消す）で確かめる。置き場は呼ぶ側が決めたツールのフォルダ" }
        @{ File = "shared\core\diagnostics.ps1"; Owner = "rotateLogFile"; Call = "Move"; Reason = "診断ログの世代を入れ替える（ログの置き場の中）" }
        @{ File = "shared\core\diagnostics.ps1"; Owner = "rotateLogFile"; Call = "Replace"; Reason = "診断ログの世代を入れ替える（ログの置き場の中）" }
        @{ File = "shared\core\fs.ps1"; Owner = "copyFileShared"; Call = "CopyTo"; Reason = "元のファイルは読み取りだけで開き、書くのは作業フォルダの写し（書き先は呼ぶ側が作業フォルダから組み立てる）" }
        @{ File = "shared\core\fs.ps1"; Owner = "copyFileShared"; Call = "FileAccess.Write"; Reason = "元のファイルは読み取りだけで開き、書くのは作業フォルダの写し（書き先は呼ぶ側が作業フォルダから組み立てる）" }
        @{ File = "shared\core\fs.ps1"; Owner = "copyFileShared"; Call = "FileMode.Create"; Reason = "元のファイルは読み取りだけで開き、書くのは作業フォルダの写し（書き先は呼ぶ側が作業フォルダから組み立てる）" }
        @{ File = "shared\core\fs.ps1"; Owner = "copyFileShared"; Call = "New-Object FileStream"; Reason = "元のファイルは読み取りだけで開き、書くのは作業フォルダの写し（書き先は呼ぶ側が作業フォルダから組み立てる）" }
        @{ File = "shared\core\fs.ps1"; Owner = "openLockFile"; Call = "FileAccess.ReadWrite"; Reason = "ワークスペースのロックファイルを作って開く（元のファイルではない）" }
        @{ File = "shared\core\fs.ps1"; Owner = "openLockFile"; Call = "FileMode.OpenOrCreate"; Reason = "ワークスペースのロックファイルを作って開く（元のファイルではない）" }
        @{ File = "shared\core\fs.ps1"; Owner = "openLockFile"; Call = "FileStream::new"; Reason = "ワークスペースのロックファイルを作って開く（元のファイルではない）" }
        @{ File = "shared\core\fs.ps1"; Owner = "removeEntryOnly"; Call = "SetAttributes"; Reason = "消す前に読み取り専用の属性を外す。呼ぶのは、置き場の下かを確かめる入口を通った削除だけ（remove_entry.Tests.ps1 が守る）" }
        @{ File = "shared\core\fs.ps1"; Owner = "writeListFile"; Call = "WriteAllLines"; Reason = "一覧・設定のファイルを書く。書き先は呼ぶ側がワークスペースか設定の置き場から決める" }
        @{ File = "shared\core\fs.ps1"; Owner = "writeTextLinesAtomic"; Call = "Move"; Reason = "一時ファイルに書いて置き換える。書き先は呼ぶ側がワークスペースか設定の置き場から決める" }
        @{ File = "shared\core\fs.ps1"; Owner = "writeTextLinesAtomic"; Call = "Replace"; Reason = "一時ファイルに書いて置き換える。書き先は呼ぶ側がワークスペースか設定の置き場から決める" }
        @{ File = "shared\core\fs.ps1"; Owner = "writeTextLinesAtomic"; Call = "WriteAllLines"; Reason = "一時ファイルに書いて置き換える。書き先は呼ぶ側がワークスペースか設定の置き場から決める" }
        @{ File = "shared\core\tsv_file.ps1"; Owner = "prettyTsv"; Call = "WriteAllText"; Reason = "Excel が出した作業フォルダの TSV（UTF-16）を整形して UTF-8 で書き直す" }
        @{ File = "shared\office\office_process.ps1"; Owner = "addOfficeRecord"; Call = "Move"; Reason = "起こした Office のプロセスの記録を書く（ワークスペースの外の一時の場所）" }
        @{ File = "shared\office\office_process.ps1"; Owner = "addOfficeRecord"; Call = "WriteAllText"; Reason = "起こした Office のプロセスの記録を書く（ワークスペースの外の一時の場所）" }
        @{ File = "shared\ui\app_host.ps1"; Owner = "loadXaml"; Call = "Save"; Reason = "XAML をメモリのストリームへ書く（ファイルではない）" }
        @{ File = "shared\ui\app_host.ps1"; Owner = "writeErrorLog"; Call = "AppendAllText"; Reason = "画面のエラーログに追記する（ログの置き場）" }
        @{ File = "tebunko\core\open_copy.ps1"; Owner = "clearOldOpenCopies"; Call = "SetAttributes"; Reason = "古い写し（開くための写しの置き場）を消す前に読み取り専用を外す。置き場の下かを確かめてから" }
        @{ File = "tebunko\core\settings.ps1"; Owner = "repairBrokenSettings"; Call = "Move"; Reason = "壊れた設定ファイル（setting.config）を別の名前へ退避する" }
        @{ File = "tebunko\core\workspace_mover.ps1"; Owner = "CopyTree"; Call = "Copy"; Reason = "ワークスペースを、利用者が選んだ移し先へ複写する（元のファイルではなくワークスペースの中身）" }
        @{ File = "tebunko\core\workspace_mover.ps1"; Owner = "MoveEntry"; Call = "Move"; Reason = "ワークスペースを、利用者が選んだ移し先へ移す（元のファイルではなくワークスペースの中身）" }
        @{ File = "tebunko\index\content_index_store.ps1"; Owner = "writeContentIndexFile"; Call = "Move"; Reason = "本文インデックスのファイルを書く（インデックスの置き場）" }
        @{ File = "tebunko\index\content_index_store.ps1"; Owner = "writeContentIndexFile"; Call = "Replace"; Reason = "本文インデックスのファイルを書く（インデックスの置き場）" }
        @{ File = "tebunko\index\content_index_store.ps1"; Owner = "writeContentIndexFile"; Call = "WriteAllText"; Reason = "本文インデックスのファイルを書く（インデックスの置き場）" }
        @{ File = "tebunko\index\gram_index_store.ps1"; Owner = "getFileGramValues"; Call = "CopyTo"; Reason = "メモリの集合を配列へ写す（ファイルではない）" }
        @{ File = "tebunko\index\gram_index_store.ps1"; Owner = "writeGramSegment"; Call = "FileAccess.Write"; Reason = "gram 索引の部分を書く（インデックスの置き場）" }
        @{ File = "tebunko\index\gram_index_store.ps1"; Owner = "writeGramSegment"; Call = "FileMode.Create"; Reason = "gram 索引の部分を書く（インデックスの置き場）" }
        @{ File = "tebunko\index\gram_index_store.ps1"; Owner = "writeGramSegment"; Call = "FileStream::new"; Reason = "gram 索引の部分を書く（インデックスの置き場）" }
        @{ File = "tebunko\index\gram_index_store.ps1"; Owner = "writeGramSegment"; Call = "Move"; Reason = "gram 索引の部分を書く（インデックスの置き場）" }
        @{ File = "tebunko\index\gram_index_store.ps1"; Owner = "writeGramToc"; Call = "Move"; Reason = "gram 索引の目次を書く（インデックスの置き場）" }
        @{ File = "tebunko\index\gram_index_store.ps1"; Owner = "writeGramToc"; Call = "Replace"; Reason = "gram 索引の目次を書く（インデックスの置き場）" }
        @{ File = "tebunko\index\gram_index_store.ps1"; Owner = "writeGramToc"; Call = "WriteAllText"; Reason = "gram 索引の目次を書く（インデックスの置き場）" }
        @{ File = "tebunko\index\index_archive.ps1"; Owner = "exportIndexCore"; Call = "Move"; Reason = "エクスポートの zip を、利用者が選んだ出力先へ書く（一時ファイルに書いてから移す）" }
        @{ File = "tebunko\index\index_archive.ps1"; Owner = "exportIndexCore"; Call = "Open"; Reason = "エクスポートの zip を、利用者が選んだ出力先へ書く（一時ファイルに書いてから移す）" }
        @{ File = "tebunko\index\index_archive.ps1"; Owner = "extractArchiveEntryToFile"; Call = "FileAccess.Write"; Reason = "インポートする zip の部品を、作業フォルダへ展開する" }
        @{ File = "tebunko\index\index_archive.ps1"; Owner = "extractArchiveEntryToFile"; Call = "FileMode.CreateNew"; Reason = "インポートする zip の部品を、作業フォルダへ展開する" }
        @{ File = "tebunko\index\index_archive.ps1"; Owner = "extractArchiveEntryToFile"; Call = "New-Object FileStream"; Reason = "インポートする zip の部品を、作業フォルダへ展開する" }
        @{ File = "tebunko\index\index_archive.ps1"; Owner = "openIndexArchive"; Call = "Open"; Reason = "インポートする zip を読み取りで開く" }
        @{ File = "tebunko\index\index_archive.ps1"; Owner = "writeArchiveFile"; Call = "CopyTo"; Reason = "エクスポートの zip の部品を書く（作業フォルダの一時ファイル）" }
        @{ File = "tebunko\index\index_format_store.ps1"; Owner = "updateIndexFormat"; Call = "FileAccess.ReadWrite"; Reason = "インデックスの目録ファイルを書き換える（ワークスペース）" }
        @{ File = "tebunko\index\index_format_store.ps1"; Owner = "updateIndexFormat"; Call = "FileMode.OpenOrCreate"; Reason = "インデックスの目録ファイルを書き換える（ワークスペース）" }
        @{ File = "tebunko\index\index_format_store.ps1"; Owner = "updateIndexFormat"; Call = "New-Object FileStream"; Reason = "インデックスの目録ファイルを書き換える（ワークスペース）" }
        @{ File = "tebunko\index\index_store.ps1"; Owner = "publishIndexFiles"; Call = "Move"; Reason = "作ったインデックスを置き場へ移す。移す前に両側が置き場の下かを確かめる" }
        @{ File = "tebunko\indexer\extract_office.ps1"; Owner = "extractDocument"; Call = "Move"; Reason = "作業フォルダの写しを、旧形式の名前へ移す（元のファイルではない）" }
        @{ File = "tebunko\indexer\extract_office.ps1"; Owner = "extractWithPowerPoint"; Call = "Open"; Reason = "作業フォルダの写しを開き、作業フォルダへ保存する（元は開かない）" }
        @{ File = "tebunko\indexer\extract_office.ps1"; Owner = "extractWithPowerPoint"; Call = "SaveAs"; Reason = "作業フォルダの写しを開き、作業フォルダへ保存する（元は開かない）" }
        @{ File = "tebunko\indexer\extract_office.ps1"; Owner = "extractWithWord"; Call = "Open"; Reason = "作業フォルダの写しを開き、作業フォルダへ保存する（元は開かない）" }
        @{ File = "tebunko\indexer\extract_office.ps1"; Owner = "extractWithWord"; Call = "SaveAs2"; Reason = "作業フォルダの写しを開き、作業フォルダへ保存する（元は開かない）" }
        @{ File = "tebunko\indexer\extract_office.ps1"; Owner = "extractWorkbook"; Call = "Open"; Reason = "作業フォルダの写しを開き、作業フォルダへ保存する（元は開かない）" }
        @{ File = "tebunko\indexer\extract_office.ps1"; Owner = "extractWorkbook"; Call = "SaveAs"; Reason = "作業フォルダの写しを開き、作業フォルダへ保存する（元は開かない）" }
        @{ File = "tebunko\indexer\extract_office.ps1"; Owner = "writeUnits"; Call = "WriteAllLines"; Reason = "取り出した本文を作業フォルダへ書く" }
        @{ File = "tebunko\indexer\extract_text.ps1"; Owner = "extractTextFile"; Call = "WriteAllLines"; Reason = "取り出した本文を作業フォルダへ書く（元は copyFileShared の写しから読む）" }
        @{ File = "tebunko\indexer\indexer_run.ps1"; Owner = "invokeIndexer"; Call = "New-Object StreamWriter"; Reason = "インデックス作成のログを書く（ワークスペース）" }
        @{ File = "tebunko\indexer\indexer_state.ps1"; Owner = "addStatusRow"; Call = "AppendAllText"; Reason = "取り込み一覧（ingest_status.tsv）に追記する（ワークスペース）" }
        @{ File = "tebunko\indexer\ingest_journal.ps1"; Owner = "AddRecord"; Call = "FileAccess.Write"; Reason = "取り込みのジャーナルに追記する（ワークスペース）" }
        @{ File = "tebunko\indexer\ingest_journal.ps1"; Owner = "AddRecord"; Call = "FileMode.Append"; Reason = "取り込みのジャーナルに追記する（ワークスペース）" }
        @{ File = "tebunko\indexer\ingest_journal.ps1"; Owner = "AddRecord"; Call = "New-Object FileStream"; Reason = "取り込みのジャーナルに追記する（ワークスペース）" }
        @{ File = "tebunko\indexer\ingest_journal.ps1"; Owner = "AddRecord"; Call = "New-Object StreamWriter"; Reason = "取り込みのジャーナルに追記する（ワークスペース）" }
        @{ File = "tebunko\search\search_export.ps1"; Owner = "invokeSearchExport"; Call = "Move"; Reason = "検索結果の出力ファイルを、利用者が選んだ出力先へ移す" }
        @{ File = "tebunko\search\search_export.ps1"; Owner = "invokeSearchExport"; Call = "Replace"; Reason = "検索結果の出力ファイルを、利用者が選んだ出力先へ移す" }
        @{ File = "tebunko\search\search_export.ps1"; Owner = "testSearchExportTarget"; Call = "FileAccess.ReadWrite"; Reason = "出力先に書けるかを、開くだけで確かめる（書き込まない。出力先は利用者が選んだ場所）" }
        @{ File = "tebunko\search\search_export.ps1"; Owner = "testSearchExportTarget"; Call = "Open"; Reason = "出力先に書けるかを、開くだけで確かめる（書き込まない。出力先は利用者が選んだ場所）" }
        @{ File = "tebunko\search\search_export.ps1"; Owner = "writeSearchExportBody"; Call = "FileAccess.Write"; Reason = "検索結果の本体を作業フォルダの一時ファイルへ書く" }
        @{ File = "tebunko\search\search_export.ps1"; Owner = "writeSearchExportBody"; Call = "FileMode.Create"; Reason = "検索結果の本体を作業フォルダの一時ファイルへ書く" }
        @{ File = "tebunko\search\search_export.ps1"; Owner = "writeSearchExportBody"; Call = "FileStream::new"; Reason = "検索結果の本体を作業フォルダの一時ファイルへ書く" }
        @{ File = "tebunko\search\search_export.ps1"; Owner = "writeSearchExportBody"; Call = "StreamWriter::new"; Reason = "検索結果の本体を作業フォルダの一時ファイルへ書く" }
        @{ File = "tebunko\search\search_export.ps1"; Owner = "writeSearchExportFile"; Call = "CopyTo"; Reason = "検索結果を出力ファイルへ書く（利用者が選んだ出力先）" }
        @{ File = "tebunko\search\search_export.ps1"; Owner = "writeSearchExportFile"; Call = "FileAccess.Write"; Reason = "検索結果を出力ファイルへ書く（利用者が選んだ出力先）" }
        @{ File = "tebunko\search\search_export.ps1"; Owner = "writeSearchExportFile"; Call = "FileMode.Create"; Reason = "検索結果を出力ファイルへ書く（利用者が選んだ出力先）" }
        @{ File = "tebunko\search\search_export.ps1"; Owner = "writeSearchExportFile"; Call = "FileStream::new"; Reason = "検索結果を出力ファイルへ書く（利用者が選んだ出力先）" }
        @{ File = "tebunko\search\search_export.ps1"; Owner = "writeSearchExportFile"; Call = "StreamWriter::new"; Reason = "検索結果を出力ファイルへ書く（利用者が選んだ出力先）" }
        @{ File = "tebunko\search\search_overflow.ps1"; Owner = "newSearchOverflowWriter"; Call = "FileAccess.Write"; Reason = "検索結果の溢れを作業フォルダへ書く（根は TmpRoot の search_overflow）" }
        @{ File = "tebunko\search\search_overflow.ps1"; Owner = "newSearchOverflowWriter"; Call = "FileMode.Create"; Reason = "検索結果の溢れを作業フォルダへ書く（根は TmpRoot の search_overflow）" }
        @{ File = "tebunko\search\search_overflow.ps1"; Owner = "newSearchOverflowWriter"; Call = "FileStream::new"; Reason = "検索結果の溢れを作業フォルダへ書く（根は TmpRoot の search_overflow）" }
        @{ File = "tebunko\search\search_overflow.ps1"; Owner = "newSearchOverflowWriter"; Call = "StreamWriter::new"; Reason = "検索結果の溢れを作業フォルダへ書く（根は TmpRoot の search_overflow）" }
        @{ File = "tebunko\ui\open_source.ps1"; Owner = "openInExcel"; Call = "Open"; Reason = "利用者が選んだ開き方で元を開く（tebunko は書かない。「通常」で開いたあとの保存は利用者の操作）" }
        @{ File = "tebunko\ui\settings\settings.ps1"; Owner = "updateSearchLimitView"; Call = "IsReadOnly ="; Reason = "画面の入力欄の IsReadOnly（ファイルではない）" }
        @{ File = "tebunko\ui\startup_error.ps1"; Owner = "writeStartupErrorFile"; Call = "Add-Content"; Reason = "起動できなかったときのエラーの記録を、ツールのフォルダへ書く" }
        @{ File = "tebunko\ui\startup_error.ps1"; Owner = "writeStartupErrorFile"; Call = "Move-Item"; Reason = "起動できなかったときのエラーの記録を、ツールのフォルダへ書く" }
    )
}

Describe "ファイルに書く API の許可リスト" -Tag Meta {
    It "書く・属性や日時を変える・ファイルを作る呼び出しは、理由つきの一覧にある関数の中だけにある" {
        $unknown = @(getAllWriteCalls | Where-Object {
                $key = $_
                @(${writeAllowed} | Where-Object { "$($_.File)|$($_.Owner)|$($_.Call)" -eq $key }).Count -eq 0
            })
        $unknown | Should -BeNullOrEmpty -Because "元のファイルに書かないため、書く API は一覧の関数だけに置く。書き先がワークスペースの置き場だと確かめて、Reason を書いて足す"
    }

    It "一覧の呼び出しは、いまも存在する（消した・名前を変えた呼び出しを一覧に残さない）" {
        $current = getAllWriteCalls
        $stale = @(${writeAllowed} | Where-Object { $current -notcontains "$($_.File)|$($_.Owner)|$($_.Call)" } |
                ForEach-Object { "$($_.File) › $($_.Owner) › $($_.Call)" })
        $stale | Should -BeNullOrEmpty
    }

    It "一覧の行には、理由が書いてある" {
        @(${writeAllowed} | Where-Object { [string]::IsNullOrWhiteSpace($_.Reason) } | ForEach-Object { "$($_.File) › $($_.Owner) › $($_.Call)" }) | Should -BeNullOrEmpty
    }

    It "一覧に同じ行が 2 つない" {
        $keys = @(${writeAllowed} | ForEach-Object { "$($_.File)|$($_.Owner)|$($_.Call)" })
        @($keys | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name }) | Should -BeNullOrEmpty
    }
}

Describe "書く API の見つけ方（自己テスト）" -Tag Meta {
    It "書く呼び出しを見つける: <code>" -TestCases @(
        @{ code = '[System.IO.File]::SetAttributes($p, "Normal")'; call = "SetAttributes" }
        @{ code = '[IO.File]::SetLastWriteTime($p, $d)'; call = "SetLastWriteTime" }
        @{ code = '$info.SetCreationTimeUtc($d)'; call = "SetCreationTimeUtc" }
        @{ code = '$info.Attributes = 0'; call = "Attributes =" }
        @{ code = '$info.IsReadOnly = $false'; call = "IsReadOnly =" }
        @{ code = '$info.LastWriteTimeUtc = $d'; call = "LastWriteTimeUtc =" }
        @{ code = '[System.IO.File]::WriteAllText($p, "x")'; call = "WriteAllText" }
        @{ code = '[System.IO.File]::AppendAllText($p, "x")'; call = "AppendAllText" }
        @{ code = '[System.IO.File]::Copy($a, $b, $true)'; call = "Copy" }
        @{ code = '[System.IO.File]::Move($a, $b)'; call = "Move" }
        @{ code = '[System.IO.File]::Create($p)'; call = "Create" }
        @{ code = '[System.IO.File]::Replace($a, $b, $c)'; call = "Replace" }
        @{ code = '$type::Copy($a, $b)'; call = "Copy" }
        @{ code = '$wb.SaveAs($p)'; call = "SaveAs" }
        @{ code = '$doc.Save($p)'; call = "Save" }
        @{ code = 'Set-Content -LiteralPath $p -Value "x"'; call = "Set-Content" }
        @{ code = 'sc $p "x"'; call = "Set-Content" }
        @{ code = 'Microsoft.PowerShell.Management\Set-Content $p "x"'; call = "Set-Content" }
        @{ code = '"x" | Out-File $p'; call = "Out-File" }
        @{ code = 'Copy-Item $a $b'; call = "Copy-Item" }
        @{ code = 'cp $a $b'; call = "Copy-Item" }
        @{ code = 'Set-ItemProperty -LiteralPath $p -Name IsReadOnly -Value $false'; call = "Set-ItemProperty" }
        @{ code = 'New-Item -ItemType File -Path $p'; call = "New-Item" }
        @{ code = '"x" > $p'; call = "リダイレクト" }
        @{ code = '"x" >> $p'; call = "リダイレクト" }
        @{ code = 'New-Object System.IO.StreamWriter($p)'; call = "New-Object StreamWriter" }
        @{ code = 'New-Object -TypeName System.IO.FileStream -ArgumentList $p, "Create"'; call = "New-Object FileStream" }
        @{ code = 'New-Object System.IO.FileStream($p, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite)'; call = "New-Object FileStream" }
        @{ code = '[System.IO.StreamWriter]::new($p)'; call = "StreamWriter::new" }
        @{ code = '[System.IO.FileStream]::new($p, [System.IO.FileMode]::Open)'; call = "FileStream::new" }
        @{ code = '[System.IO.File]::Open($p, [System.IO.FileMode]::OpenOrCreate)'; call = "Open" }
        @{ code = '[System.IO.File]::Open($p, $mode, $access)'; call = "Open" }
        @{ code = '[System.IO.File]::Open($p, [System.IO.FileMode]::Open)'; call = "Open" }
        @{ code = '$m = [System.IO.FileMode]::Create'; call = "FileMode.Create" }
        @{ code = '$m = [System.IO.FileMode]::Append'; call = "FileMode.Append" }
        @{ code = '$a = [System.IO.FileAccess]::Write'; call = "FileAccess.Write" }
        @{ code = '$a = [System.IO.FileAccess]::ReadWrite'; call = "FileAccess.ReadWrite" }
        @{ code = '$t = [System.IO.FileMode]; [Enum]::Parse($t, "Create")'; call = "FileMode（型の参照）" }
        @{ code = '$mode = [System.IO.FileMode]"Create"'; call = "FileMode（型の参照）" }
        @{ code = '[Enum]::Parse([System.IO.FileAccess], "Write")'; call = "FileAccess（型の参照）" }
    ) {
        @(findWriteCalls (parseCode $code) | ForEach-Object { $_.Call }) | Should -Contain $call
    }

    It "読むだけの呼び出しは見つけない: <code>" -TestCases @(
        @{ code = '[System.IO.File]::ReadAllText($p)' }
        @{ code = '[System.IO.File]::OpenRead($p)' }
        @{ code = '[System.IO.File]::Open($p, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)' }
        @{ code = '[System.IO.File]::Open($p, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)' }
        @{ code = '[System.IO.FileStream]::new($p, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)' }
        @{ code = 'New-Object System.IO.FileStream($p, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)' }
        @{ code = 'Get-Content -LiteralPath $p'; }
        @{ code = 'Get-ChildItem $p | ForEach-Object { $_.Name }' }
        @{ code = '[System.IO.File]::GetAttributes($p)'; }
        @{ code = '[System.IO.File]::GetLastWriteTimeUtc($p)' }
        @{ code = '$info.LastWriteTimeUtc' }
        @{ code = '$x = $info.Attributes -bor 1' }
        @{ code = '$info.Length = 3; $t.Name = "x"' }
        @{ code = '$null = Get-Item $p *> $null' }
        @{ code = 'Get-Item $p 2>$null' }
        @{ code = '[System.IO.Directory]::CreateDirectory($d)' }
        @{ code = '$text.Replace("a", "b")' }
        @{ code = '[Array]::Copy($a, $b, 3)' }
        @{ code = '$ps = [powershell]::Create()' }
        @{ code = '$runspace.Open()' }
        @{ code = 'Remove-Item $p -Force' }
    ) {
        @(findWriteCalls (parseCode $code)).Count | Should -Be 0
    }

    It "呼び出しは、いちばん内側の関数の名前で数える（関数の外はスクリプト直下）" {
        $ast = parseCode 'function outer { function inner { [System.IO.File]::WriteAllText($a, "x") }; Set-Content $b "y" }; Out-File $c'
        $owners = @(findWriteCalls $ast | ForEach-Object { "$($_.Owner)|$($_.Call)" })
        $owners | Should -Contain "inner|WriteAllText"
        $owners | Should -Contain "outer|Set-Content"
        $owners | Should -Contain "(スクリプト直下)|Out-File"
        $owners.Count | Should -Be 3
    }
}
