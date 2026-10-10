# インデックスのフォルダ（IndexDir・content_index）を消す関数を、許した関数の一覧に限る。
# 設定から外れたインデックスを確認なしに消していた不具合（利用者のデータが黙って消える）の再発を防ぐ。
# 消す関数を足すときは、利用者の確認を取ってから消すことを確かめたうえで、下の一覧に理由つきで足す。
#
# 判定の単位は関数（クラスのメソッドを含む）。本体に IndexDir・content_index・getWorkspaceEntries のどれかと、Remove-Item・removeDirectoryRetry・
# Directory.Delete・.Delete( のどれかがある関数を対象にする（変数を経由した削除、$dir = Join-Path $ws.IndexDir ... のあとの削除も拾うため）。
# 拾えないもの: 上の語を本体に書かずに、引数で渡されたパスを消す関数（moveWorkspaceEntry は引数のフォルダを移し終えたあとに消す。ワークスペースを移す moveWorkspace の中だけで呼ばれ、
# 拾えていない）。そういう関数を足すときは、呼ぶ側の関数の名前をこの先頭の説明に書く。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"

    # 許した関数: File（ファイル名）・Function（関数・メソッドの名前）・Reason（利用者の確認を取る、または利用者のデータではない理由）
    ${indexDeleteAllowed} = @(
        @{ File = "index_store.ps1"; Function = "removeIndex"; Reason = "画面の［削除］の流れ。確認ダイアログで利用者が選んだインデックスだけを消す" }
        @{ File = "index_migrate.ps1"; Function = "removeDroppedFolders"; Reason = "設定から外れたインデックス。インデックス更新の確認ダイアログで利用者が［更新を開始］を押した回でだけ呼ぶ（findDroppedIndexes は見つけるだけ）" }
        @{ File = "index_archive.ps1"; Function = "exportIndexCore"; Reason = "書き出しの途中のファイル（一時ファイル）だけを消す。インデックスのフォルダは消さない" }
        @{ File = "index_archive.ps1"; Function = "expandImportArchive"; Reason = "インポートの作業フォルダ（workDir）だけを消す。インデックスのフォルダは消さない" }
        @{ File = "index_archive.ps1"; Function = "swapInImportedIndexDir"; Reason = "消さずに移すだけ（前のフォルダは退避する）。コメントに removeDirectoryRetry の語があって拾われる" }
        @{ File = "index_archive.ps1"; Function = "restoreSwappedIndexDir"; Reason = "インポートの途中で失敗したとき、入れたばかりのフォルダを消して、退避した前のフォルダを戻す（取り消し）" }
        @{ File = "index_archive.ps1"; Function = "importIndexCore"; Reason = "作業フォルダを消す。上書きは、画面の確認（上書きする／別の名前）で利用者が選んだ回だけ" }
        @{ File = "index_store.ps1"; Function = "publishIndexFiles"; Reason = "取り込みで作り直す元のファイル 1 つ分のフォルダを、更新の確認を経た取り込みの結果に置き換える" }
        @{ File = "system_index.ps1"; Function = "removeSystemIndexOf"; Reason = "システムインデックス（インデックスから作り直せる派生物）だけを消す" }
        @{ File = "workspace.ps1"; Function = "clearLegacySystemIndex"; Reason = "前の版のシステムインデックス（作り直せる派生物）だけを消す" }
        @{ File = "workspace.ps1"; Function = "removeWorkspaceEntries"; Reason = "ワークスペースを変えるときの確認で、利用者が［最初からやり直す］を選んだ回だけ" }
        @{ File = "pack_store.ps1"; Function = "convertIndexFolderToPack"; Reason = "集約ファイルに入れ終えた、元のファイルごとの TSV のフォルダだけを消す（同じ内容が集約ファイルに残る）" }
    )

    function findIndexDeleteFunctions {
        # scripts の下で、インデックスのフォルダを消す関数を @{ File; Function } の配列で返す
        $found = New-Object System.Collections.Generic.List[object]
        foreach ($file in @(Get-ChildItem -LiteralPath ${scriptsDir} -Filter "*.ps1" -Recurse)) {
            $errors = $null
            $tokens = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
            $functions = $ast.FindAll({
                    param ($n)
                    $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -or $n -is [System.Management.Automation.Language.FunctionMemberAst]
                }, $true)
            foreach ($function in $functions) {
                $body = $function.Extent.Text
                if ($body -match "IndexDir|content_index|getWorkspaceEntries" -and $body -match "Remove-Item|removeDirectoryRetry|Directory\]::Delete|\.Delete\(") {
                    $found.Add(@{ File = $file.Name; Function = $function.Name })
                }
            }
        }
        return $found.ToArray()
    }
}

Describe "インデックスのフォルダを消す関数" -Tag Meta {
    It "許した一覧に無い関数が、インデックスのフォルダを消していない" {
        $unknown = @(findIndexDeleteFunctions | Where-Object {
                $item = $_
                @(${indexDeleteAllowed} | Where-Object { $_.File -eq $item.File -and $_.Function -eq $item.Function }).Count -eq 0
            } | ForEach-Object { "$($_.File) › $($_.Function)" })
        $unknown | Should -BeNullOrEmpty -Because "インデックスを消す関数は、利用者の確認を取ってから消すこと。足すなら tests/meta/index_delete.Tests.ps1 の一覧に理由つきで足す"
    }

    It "許した一覧の関数は、いまも存在する（消した・名前を変えた関数を一覧に残さない）" {
        $current = @(findIndexDeleteFunctions)
        $stale = @(${indexDeleteAllowed} | Where-Object {
                $item = $_
                @($current | Where-Object { $_.File -eq $item.File -and $_.Function -eq $item.Function }).Count -eq 0
            } | ForEach-Object { "$($_.File) › $($_.Function)" })
        $stale | Should -BeNullOrEmpty
    }

    It "removeDroppedFolders は、インデクサの中で確認のあとにだけ呼ばれる" {
        $text = [System.IO.File]::ReadAllText("${scriptsDir}\tebunko\indexer\indexer_run.ps1")
        $calls = [regex]::Matches($text, "removeDroppedFolders")
        $calls.Count | Should -Be 1
        $text.IndexOf("removeDroppedFolders") | Should -BeGreaterThan $text.IndexOf("WaitForApproval")
    }
}
