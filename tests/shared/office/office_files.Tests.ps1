# Office ファイルの判定（shared\office\office_files.ps1）のテスト
. "$PSScriptRoot\..\..\helpers\load.ps1"

Describe "testOfficeFile" -Tag Unit {
    It "Excel・Word・PowerPoint のファイルを見分ける（大文字・小文字は区別しない）" {
        testOfficeFile "見積.xlsx" | Should Be $true
        testOfficeFile "報告.DOCM" | Should Be $true
        testOfficeFile "資料.ppt" | Should Be $true
    }

    It "Office 以外のファイル・拡張子の無い名前は false" {
        testOfficeFile "メモ.txt" | Should Be $false
        testOfficeFile "データ" | Should Be $false
        testOfficeFile "" | Should Be $false
    }
}
