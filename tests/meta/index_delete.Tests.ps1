# インデックスのフォルダ（IndexDir・content_index）を消す関数を、許した関数の一覧に限る。
# 設定から外れたインデックスを確認なしに消していた不具合（利用者のデータが黙って消える）の再発を防ぐ。
# 消す関数を足すときは、利用者の確認を取ってから消すことを確かめたうえで、下の一覧に理由つきで足す。
#
# 判定の単位は関数（クラスのメソッドを含む）。本体に IndexDir または content_index と、Remove-Item または
# Directory.Delete の両方がある関数を対象にする（変数を経由した削除、$dir = Join-Path $ws.IndexDir ... のあとの Remove-Item も拾うため）。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"

    # 許した関数: File（ファイル名）・Function（関数・メソッドの名前）・Reason（利用者の確認を取る、または利用者のデータではない理由）
    ${indexDeleteAllowed} = @(
        @{ File = "index_store.ps1"; Function = "removeIndex"; Reason = "画面の［削除］の流れ。確認ダイアログで利用者が選んだインデックスだけを消す" }
        @{ File = "index_migrate.ps1"; Function = "removeDroppedFolders"; Reason = "設定から外れたインデックス。インデックス更新の確認ダイアログで利用者が［更新を開始］を押した回でだけ呼ぶ（findDroppedIndexes は見つけるだけ）" }
        @{ File = "index_archive.ps1"; Function = "exportIndexCore"; Reason = "書き出しの途中のファイル（一時ファイル）だけを消す。インデックスのフォルダは消さない" }
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
                if ($body -match "IndexDir|content_index" -and $body -match "Remove-Item|Directory\]::Delete|\.Delete\(") {
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
