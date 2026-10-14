# Office ファイルの拡張子と種類の表（shared\office\office_files.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"
}

Describe "officeFileKinds・officeExtensions" -Tag Unit {
    It "取り込みの対象の拡張子は、表の拡張子と同じ並びである" {
        @(${officeExtensions}) | Should -Be @(${officeFileKinds}.Keys)
    }

    It "表の種類は Excel・Word・PowerPoint のどれかである" {
        foreach ($kind in @(${officeFileKinds}.Values)) {
            @("Excel", "Word", "PowerPoint") -contains $kind | Should -Be $true
        }
    }
}

Describe "getOfficeFileKind" -Tag Unit {
    # 表（officeFileKinds）にある拡張子だけが Office。大文字小文字は区別しない。表に無い拡張子（.xlsxyz・.xlsa など）は空
    It "<name> の種類は「<kind>」" -ForEach @(
        @{ name = "見積.xlsx"; kind = "Excel" }
        @{ name = "見積.xlsm"; kind = "Excel" }
        @{ name = "見積.xls"; kind = "Excel" }
        @{ name = "見積.xlsb"; kind = "Excel" }
        @{ name = "報告.docx"; kind = "Word" }
        @{ name = "報告.docm"; kind = "Word" }
        @{ name = "報告.doc"; kind = "Word" }
        @{ name = "発表.pptx"; kind = "PowerPoint" }
        @{ name = "発表.pptm"; kind = "PowerPoint" }
        @{ name = "発表.ppt"; kind = "PowerPoint" }
        @{ name = "見積.XLSX"; kind = "Excel" }
        @{ name = "報告.Doc"; kind = "Word" }
        @{ name = "営業部\2024\見積.xlsx"; kind = "Excel" }
        @{ name = ".xlsx"; kind = "Excel" }
        @{ name = "a.xlsxyz"; kind = "" }
        @{ name = "a.xlsa"; kind = "" }
        @{ name = "a.docz"; kind = "" }
        @{ name = "a.pptz"; kind = "" }
        @{ name = "a.xlsx.bak"; kind = "" }
        @{ name = "a.xlsx."; kind = "" }
        @{ name = "a.xl"; kind = "" }
        @{ name = "ひな形.xltx"; kind = "" }
        @{ name = "ひな形.dotx"; kind = "" }
        @{ name = "xls"; kind = "" }
        @{ name = "営業部.xlsx\メモ"; kind = "" }
        @{ name = "a.txt"; kind = "" }
        @{ name = ""; kind = "" }
        @{ name = $null; kind = "" }
    ) {
        getOfficeFileKind $name | Should -BeExactly $kind
    }

    It "名前に使えない文字があっても例外にしない" {
        { getOfficeFileKind "a<b>|c.xlsx" } | Should -Not -Throw
        getOfficeFileKind "a<b>|c.xlsx" | Should -BeExactly "Excel"
    }
}
